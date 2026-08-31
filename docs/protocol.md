# Conversion Server Protocol

Wire contract between `hwp2pdf.backends.remote_http.RemoteHttpBackend` (client)
and `hwp2pdf.server.http_server` (server). Both sides import the paths, headers
and event kinds from `src/hwp2pdf/server/protocol.py`, so this document
describes that module rather than duplicating it.

Setup instructions live in [remote-server.md](remote-server.md).

## Versioning

`API_VERSION` is currently **1**. The client refuses to talk to a server that
reports a different value. Bump it whenever an existing field changes meaning;
adding a new optional field does not require a bump.

**This is not the build version.** The client compares `API_VERSION` only, never
`hwp2pdf`'s `yyyy.MM.dd.N`, so a server and a client from different releases work
together as long as the protocol did not change -- which is the usual case. Only
a release that bumps `API_VERSION` forces the two sides to be updated together.

## Model

A **job** is one batch session. An **item** is one source file converted to one
output format, matching the inner loop of `jobs.run_batch` exactly. A job holds
an append-only event log; clients read it with a cursor.

```
POST   /v1/jobs                          -> job_id
PUT    /v1/jobs/{id}/inputs/{item}       (raw bytes; idempotent)
POST   /v1/jobs/{id}/items/{item}/run    -> 202
GET    /v1/jobs/{id}/events?cursor=N&wait=25
GET    /v1/jobs/{id}/outputs/{item}      -> raw bytes
POST   /v1/jobs/{id}/cancel
DELETE /v1/jobs/{id}
```

## Endpoints

| Method | Path | Auth | Purpose |
|---|---|---|---|
| GET | `/v1/health` | none | `{app, version, api, auth_required}`. Unauthenticated on purpose so a client can tell *unreachable* from *bad token* |
| GET | `/v1/capabilities` | bearer | `{os, hwp_installed, hwp_detail, hwp_running, formats, extensions, shares, max_upload_bytes, queue_depth}` |
| POST | `/v1/jobs` | bearer | `{lang, safe_temp}` -> `{job_id, queue_depth}`; `201` |
| PUT | `/v1/jobs/{id}/inputs/{item}` | bearer | `application/octet-stream`, `Content-Length` required; `204`. Written to a `.part` file and renamed, so a retry is safe |
| POST | `/v1/jobs/{id}/items/{item}/run` | bearer | `{name, output_format, force_one_page, share?, rel?, out_rel?}`; `202` |
| GET | `/v1/jobs/{id}/events` | bearer | `?cursor=N&wait=S` -> `{events, cursor, cancelled, queue_depth}` |
| GET | `/v1/jobs/{id}/outputs/{item}` | bearer | Converted bytes |
| POST | `/v1/jobs/{id}/cancel` | bearer | Marks the job cancelled and wakes pollers |
| DELETE | `/v1/jobs/{id}` | bearer | Closes the engine session and removes the work directory |

## Events

Cursor-based long polling, not SSE: a reconnect after a network drop is just the
same GET with the last cursor, no stream parser is needed on the client, and a
handler thread is never pinned indefinitely.

```json
{"seq": 1, "kind": "session", "state": "started"}
{"seq": 2, "kind": "log", "text": "HWPFrame.HwpObject started.", "level": "info"}
{"seq": 3, "kind": "item", "item": "00001-PDF", "status": "ok",
 "actual": "PDF", "message": "", "notices": []}
```

`kind` is one of `session`, `log`, `item`. `status` is `ok`, `failed` or
`blocked`. `level` is `info`, `warning` or `error`. Log events are re-emitted on
the client prefixed with `서버:` / `server:` so users can tell where a message
came from.

`GET .../events` returns as soon as there is anything after `cursor`, or after
`wait` seconds with an empty list. `wait` is clamped to 25 s.

## Localization

The client sends `lang` when it creates the job; the server renders every
message with the same `hwp2pdf.i18n.TEXT` table. There is one source of truth
for Korean and English strings, and failure text arrives ready to display.

## Auth

`Authorization: Bearer <token>`, compared with `hmac.compare_digest`. A bad or
missing token gets `401` with `WWW-Authenticate: Bearer`. `/v1/health` is the
only unauthenticated route. The server refuses to start on a non-loopback
address without a token.

## Status codes

| Code | Meaning |
|---|---|
| `400` | Malformed JSON, unknown output format, missing upload, or a share path that escapes its root |
| `401` | Missing or wrong token |
| `404` | Unknown job, item, route, or no output produced |
| `411` | `PUT` without `Content-Length` |
| `413` | Upload exceeds `--max-upload-bytes` |
| `429` | Conversion queue is full (`--max-queue`) |

## Concurrency

The Hancom COM engine is a per-process, per-thread singleton, so exactly one
worker thread owns it and jobs are converted strictly in order. Uploads,
downloads and event polling run on handler threads. `queue_depth` is reported by
`/v1/capabilities` and on every event page so a client can show how much is
waiting ahead of it.

## Share transport

When a job item carries `share` + `rel` + `out_rel`, the server resolves them
against a configured `--share-root` and converts the file in place, writing the
output next to it. `resolved.is_relative_to(root)` is enforced on both paths;
anything else is a `400`. No bytes cross the wire in either direction.
