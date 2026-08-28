"""Wire contract shared by the server and the remote client.

Keep both sides importing from here so a change cannot drift on one side only.
Bump ``API_VERSION`` whenever an existing field changes meaning; the client
refuses to talk to a server with a different major contract.
"""

API_VERSION = 1
DEFAULT_PORT = 8765

AUTH_HEADER = "Authorization"
AUTH_SCHEME = "Bearer"
SHA256_HEADER = "X-Hwp2pdf-Sha256"

PATH_HEALTH = "/v1/health"
PATH_CAPABILITIES = "/v1/capabilities"
PATH_JOBS = "/v1/jobs"


def job_path(job_id: str) -> str:
    return f"{PATH_JOBS}/{job_id}"


def input_path(job_id: str, item_id: str) -> str:
    return f"{PATH_JOBS}/{job_id}/inputs/{item_id}"


def run_path(job_id: str, item_id: str) -> str:
    return f"{PATH_JOBS}/{job_id}/items/{item_id}/run"


def output_path(job_id: str, item_id: str) -> str:
    return f"{PATH_JOBS}/{job_id}/outputs/{item_id}"


def events_path(job_id: str, cursor: int, wait: int) -> str:
    return f"{PATH_JOBS}/{job_id}/events?cursor={cursor}&wait={wait}"


def cancel_path(job_id: str) -> str:
    return f"{PATH_JOBS}/{job_id}/cancel"


# Event kinds pushed onto a job's append-only event log.
EVENT_LOG = "log"          # {"kind","seq","text","level"}
EVENT_ITEM = "item"        # {"kind","seq","item","status","actual","message","notices"}
EVENT_SESSION = "session"  # {"kind","seq","state"}

ITEM_OK = "ok"
ITEM_FAILED = "failed"
ITEM_BLOCKED = "blocked"

TRANSPORT_UPLOAD = "upload"
TRANSPORT_SHARE = "share"

# Server defaults; each is overridable from the command line.
DEFAULT_MAX_UPLOAD_BYTES = 512 * 1024 * 1024
DEFAULT_MAX_QUEUE = 8
DEFAULT_JOB_TTL_SECONDS = 3600
DEFAULT_EVENT_WAIT_SECONDS = 25
