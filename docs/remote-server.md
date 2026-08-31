# 변환 서버 설정 / Conversion Server Setup

macOS(및 Linux)에는 한컴오피스 자동화 엔진이 없습니다. `hwp2pdf`의 mac 앱은
**한컴오피스가 설치된 Windows 컴퓨터에서 도는 변환 서버**에 붙어서 동작합니다.
파일 탐색·스킵/덮어쓰기 판정·CSV 로그·진행률·중지는 전부 mac 쪽에 남고,
서버는 "파일 하나를 한 포맷으로 저장"만 담당합니다.

```
[mac] hwp2pdf.app ── HTTP ──► [Windows] hwp2pdf-cli serve ──► 한컴오피스 한글 (COM)
      결과 파일은 mac이 지정한 위치에 그대로 저장됩니다.
```

---

## 1. 서버 준비 (Windows)

### 1.1 요구 사항

- Windows 10/11
- 한컴오피스 한글 설치 (COM 자동화 `HWPFrame.HwpObject` 사용 가능)
- `hwp2pdf` 설치본 또는 배포 zip의 `hwp2pdf-cli.exe`

### 1.2 토큰 생성과 첫 실행

```powershell
hwp2pdf-cli.exe serve --init
```

`--init`은 토큰이 없으면 만들어 `%LOCALAPPDATA%\hwp2pdf\server_token.txt`에 저장하고
화면에 한 번 출력합니다. 이 토큰을 mac 클라이언트에 입력하면 됩니다.
`--token`이나 환경변수 `HWP2PDF_TOKEN`으로 직접 지정할 수도 있습니다.

시작 배너에서 확인할 것:

```
hwp2pdf conversion server v2026.08.28.1 (API 1)
  listening   http://100.124.117.75:8765
  auth        token AbCdEf...
  hangul      yes (SOFTWARE\HNC\HwpRun)
  shares      (none)
  max upload  512 MB
```

`hangul`이 `NO`면 한컴오피스나 pywin32 쪽 문제입니다. 먼저 그것부터 해결하세요.

### 1.3 ⚠️ 서비스로 등록하지 마세요

한글 COM 자동화는 **로그인된 대화형 데스크톱 세션**이 필요합니다.
Windows 서비스(`sc create`, NSSM 등)로 등록하면 Session 0에서 실행되어
데스크톱이 없고, 좀비 `Hwp.exe`가 남습니다 (`docs/known-issues.md` §2).

자동 시작이 필요하면 **작업 스케줄러에 "사용자가 로그온할 때" 트리거**로 등록하고
**"사용자가 로그온한 경우에만 실행"** 을 선택하세요. 도우미 스크립트가 있습니다:

```powershell
.\scripts\install_serve_task.ps1 -Bind tailscale
```

기본으로 쓰이는 `hwp2pdf-serve.exe`는 **창 없는(windowless) 빌드**라 데스크톱에
콘솔 창이 뜨지 않습니다 — 실수로 닫아서 서버가 죽는 일이 없습니다. 대신 출력은
`%LOCALAPPDATA%\hwp2pdf\server.log`에 쌓입니다(2MB마다 `.log.1`로 롤링).

```powershell
Get-Content "$env:LOCALAPPDATA\hwp2pdf\server.log" -Tail 20
```

`hwp2pdf-cli.exe serve`(콘솔 빌드)로 띄우면 검은 창이 뜨고, **그 창을 닫으면 서버가
죽습니다**(`LastTaskResult`가 `0xC000013A` = `STATUS_CONTROL_C_EXIT`). 스크립트가
등록하는 작업에는 재시작 설정이 들어 있어 1분 안에 되살아나지만, 그 사이 변환은 끊깁니다.

등록된 작업에는 트리거가 셋 있습니다 — **로그온 시**, **시스템 시작 시**, 그리고
**10분마다 재확인**. 마지막 것이 자가 치유 역할을 해서, 어떤 이유로든 서버가 죽어 있으면
10분 안에 다시 뜹니다(이미 떠 있으면 아무 일도 하지 않습니다).

수동으로 즉시 띄우려면:

```powershell
Start-ScheduledTask -TaskName "hwp2pdf serve"
```

⚠️ **한글 COM은 로그인된 데스크톱 세션이 필요합니다.** 업데이트로 재부팅된 뒤 아무도
로그인하지 않으면 서버는 뜰 수 없습니다(절전 중에도 마찬가지). 다음 순서로 대비하세요:

1. **"업데이트 후 자동 로그인"을 켜 둡니다** — 설정 > 계정 > 로그인 옵션.
   Windows Hello PIN과 함께 써도 정상 동작합니다. 잠긴 화면으로 복구되며, 잠긴 세션도
   대화형 데스크톱이라 한글 자동화가 됩니다.
   다만 Windows가 **재시작을 연달아 두 번** 거는 경우(업데이트 오케스트레이터 → 서비싱)
   이 기능이 두 번째 재시작을 넘기지 못해 잠금 화면에 머물 수 있습니다.
2. 그 상황에 대비해 **Tailscale 무인 모드**를 켜 둡니다(위 참고). 변환은 못 해도
   원격 접근은 살아 있어 로그인만 하면 바로 복구됩니다.
3. 작업 스케줄러의 10분 주기 트리거가 로그인 직후 서버를 자동으로 띄웁니다.

> 고전 방식 자동 로그인(`AutoAdminLogon`)은 권하지 않습니다. Windows Hello를 쓰는
> 기기에서는 TPM 기반 암호 없는 로그인을 끄고 암호를 저장해야 해서, 얻는 것보다
> 잃는 보안이 큽니다.

---

## 2. 연결 방법 3가지

주소만 다를 뿐 나머지는 동일합니다. **`--bind`는 "어느 네트워크에 노출할지"를 정합니다.**

### 2.1 Tailscale (권장)

테일넷 안에서만 열리고, WireGuard로 암호화되며, 방화벽 규칙이 필요 없습니다.

**서버(Windows):**

```powershell
hwp2pdf-cli.exe serve --bind tailscale --init
```

`--bind tailscale`은 `tailscale ip -4`로 이 컴퓨터의 100.x 주소를 찾아 **거기에만** 바인드합니다.
LAN이나 공인망에는 아예 열리지 않으므로 방화벽 예외를 추가할 필요가 없습니다.

**⚠️ 서버에서 Tailscale 무인 모드를 켜 두세요.** 이걸 안 하면 재부팅 후 아무도
로그인하지 않았을 때 **Tailscale 노드까지 오프라인이 되어 원격 접근 자체가 끊깁니다**
(`tailscaled` 서비스는 떠 있어도 사용자 프로필이 없어 노드가 내려갑니다).
켜 두면 세션이 없어도 연결이 유지되므로, 최소한 SSH로 들어가 상황을 볼 수 있습니다.

```powershell
tailscale set --unattended=true
tailscale debug prefs | Select-String ForceDaemon   # "ForceDaemon": true 확인
```

**클라이언트(mac):** 앱의 `변환 서버` 칸에 MagicDNS 이름 또는 100.x 주소를 입력합니다.

```
주소   http://<호스트이름>.<테일넷>.ts.net:8765
토큰   <서버가 출력한 토큰>
```

예: 호스트 `namun-ji`, 테일넷 `tail0fdba8.ts.net` → `http://namun-ji.tail0fdba8.ts.net:8765`
(테일넷 이름은 `tailscale status --json`의 `MagicDNSSuffix`에서 확인)

MagicDNS가 꺼져 있으면 `tailscale status`에 보이는 100.x 주소를 그대로 쓰면 됩니다.

**TLS까지 원하면** (평문 HTTP가 싫을 때) Tailscale이 인증서를 대신 처리해 줍니다:

```powershell
hwp2pdf-cli.exe serve --bind 127.0.0.1
tailscale serve --bg 8765
```

그러면 클라이언트 주소가 `https://<호스트>.<테일넷>.ts.net`이 됩니다.
Tailscale 자체가 이미 WireGuard로 암호화하므로 필수는 아닙니다.

### 2.2 같은 LAN

```powershell
hwp2pdf-cli.exe serve --bind 0.0.0.0 --init
```

Windows 방화벽에서 인바운드 포트를 열어야 합니다. **개인 프로필로만** 여세요:

```powershell
.\scripts\allow_firewall.ps1          # 아래 규칙을 그대로 실행합니다
# New-NetFirewallRule -DisplayName "hwp2pdf serve" -Direction Inbound `
#   -Protocol TCP -LocalPort 8765 -Profile Private -Action Allow
```

클라이언트 주소는 `http://<Windows LAN IP>:8765`.

> ⚠️ LAN에서는 통신이 **평문 HTTP**입니다. 신뢰할 수 있는 사내망에서만 쓰거나
> `--tls-cert` / `--tls-key`로 TLS를 켜세요. 공인 인터넷에 절대 노출하지 마세요.

### 2.3 mac 안의 Windows VM (Parallels / VMware / UTM)

VM 안에서:

```powershell
hwp2pdf-cli.exe serve --bind 0.0.0.0 --init
```

- **공유 네트워크(NAT)**: VM의 IP(`ipconfig`)를 그대로 사용 →
  `http://10.211.55.x:8765` (Parallels 기본 대역)
- **포트 포워딩**: Parallels/VMware의 네트워크 설정에서 호스트 8765 → 게스트 8765를
  전달하면 mac에서 `http://127.0.0.1:8765`로 붙을 수 있습니다.

VM은 mac과 파일시스템을 공유하기 쉬우므로 **공유 폴더 모드**(§3)와 특히 잘 맞습니다.

---

## 3. 전송 방식: 업로드 vs 공유 폴더

| 모드 | 동작 | 언제 |
|---|---|---|
| **업로드**(기본) | 원본을 서버로 올리고 결과를 내려받음 | 어디서나 동작. 별도 설정 없음 |
| **공유 폴더** | 경로만 전달, 서버가 원본 자리에서 변환 | mac과 Windows가 같은 저장소를 볼 때. 큰 파일에 훨씬 빠름 |
| **자동** | 원본이 매핑된 마운트 아래면 공유 폴더, 아니면 업로드 | 기본값 |

**서버 쪽** — 노출할 폴더를 이름과 함께 등록합니다 (여러 번 지정 가능):

```powershell
hwp2pdf-cli.exe serve --bind tailscale --share-root work=D:\shared --share-root pub=\\nas\pub
```

**클라이언트 쪽** — mac의 마운트 지점을 같은 이름에 매핑합니다.
`~/Library/Application Support/hwp2pdf/settings.json`:

```json
{
  "server": {
    "url": "http://namun-ji.tail0fdba8.ts.net:8765",
    "token": "...",
    "transport": "auto",
    "shares": [
      { "name": "work", "local_mount": "/Volumes/shared" }
    ]
  }
}
```

변환 대상이 `/Volumes/shared` 아래면 업로드 없이 `{"share":"work","rel":"..."}`만 전달되고,
결과도 공유 폴더에 직접 쓰이므로 다운로드도 생략됩니다.
그 밖의 경로는 자동으로 업로드 모드로 떨어집니다.

서버는 `share_root` 밖으로 나가는 경로(`../` 등)를 **400으로 거부**합니다.

---

## 4. 클라이언트 설정 (mac)

### GUI

앱의 `변환 서버` 패널에 주소·토큰·전송 방식을 넣고 **연결 테스트**를 누릅니다.
성공하면 서버 버전·한글 설치 여부·대기열 길이가 표시됩니다.
설정은 `~/Library/Application Support/hwp2pdf/settings.json`에 저장됩니다(권한 `0600`).

### CLI

```bash
hwp2pdf ~/문서/보고서폴더 --pdf --docx \
    --server http://namun-ji.tail0fdba8.ts.net:8765 \
    --token  <토큰>
```

환경변수도 쓸 수 있습니다 (저장된 설정보다 우선):

```bash
export HWP2PDF_SERVER_URL=http://namun-ji.tail0fdba8.ts.net:8765
export HWP2PDF_TOKEN=...
hwp2pdf ~/문서/보고서폴더 --pdf
```

### 점검

```bash
./scripts/check_macos.sh                       # 파이썬·Tk·의존성·서버 도달 확인
python scripts/smoke_remote.py <url> <token> <sample.hwp>   # 실제 변환 왕복 검증
```

---

## 5. 서버 옵션

| 옵션 | 기본값 | 설명 |
|---|---|---|
| `--bind` | `127.0.0.1` | `tailscale` / 주소 / `0.0.0.0` |
| `--port` | `8765` | |
| `--token` | (파일) | 없으면 `%LOCALAPPDATA%\hwp2pdf\server_token.txt` |
| `--init` | | 토큰이 없으면 생성해 저장하고 출력 |
| `--no-auth` | | 루프백 바인드일 때만 허용 |
| `--share-root NAME=PATH` | | 공유 폴더 등록 (반복 가능) |
| `--max-upload-bytes` | 512 MB | 초과 시 `413` |
| `--max-queue` | 8 | 초과 시 `429` |
| `--job-ttl` | 3600초 | 유휴 작업 정리 주기 |
| `--tls-cert` / `--tls-key` | | 평문 HTTP 대신 TLS |
| `--quiet` | | 요청 로그 끄기 |
| `--log-file` | (창 없을 때 `server.log`) | 출력을 파일로 남김 |

루프백이 아닌 주소에 바인드하면서 토큰이 없으면 **서버가 기동을 거부합니다.**

---

## 6. 문제 해결

| 증상 | 원인과 조치 |
|---|---|
| `변환 서버에 연결하지 못했습니다` | 서버가 안 떠 있거나 주소/포트가 틀림. 방화벽. Tailscale이면 `tailscale status`로 피어가 online인지 확인 |
| `변환 서버 인증에 실패했습니다` (401) | 토큰 불일치. 서버 배너의 토큰 앞 6자와 클라이언트 설정을 대조 |
| `변환 서버와 버전이 맞지 않습니다` | 양쪽 `hwp2pdf` 버전이 다름. 같은 릴리스로 맞추세요 |
| `변환 서버에 한컴오피스 한글이 없습니다` | 서버 배너의 `hangul` 줄 확인. 한컴오피스/pywin32 설치 필요 |
| `변환 서버가 처리 중인 작업이 많습니다` (429) | 큐가 찼습니다. 한컴 엔진은 직렬 처리라 동시 실행이 불가능합니다. 잠시 후 재시도하거나 `--max-queue`를 올리세요 |
| `파일이 서버의 업로드 상한을 초과했습니다` (413) | 공유 폴더 모드를 쓰거나 `--max-upload-bytes`를 올리세요 |
| `공유 폴더에서 변환 결과를 찾지 못했습니다` | mac 마운트와 서버 `--share-root`가 같은 저장소를 가리키는지, 쓰기 권한이 있는지 확인 |
| 서버에서 한컴 대화상자가 떠서 멈춤 | 서버를 대화형 세션에서 실행 중인지 확인. 서비스로 돌리면 발생합니다 |
| 잘 되다가 갑자기 연결 실패 | `server.log` 마지막 줄과 `Get-ScheduledTaskInfo`의 `LastTaskResult`를 보세요. `3221225786`(`0xC000013A`)이면 콘솔 빌드의 창이 닫힌 경우입니다(창 없는 `hwp2pdf-serve.exe`를 쓰면 안 생깁니다). `Start-ScheduledTask`로 재시작 |
| 부팅/절전 복귀 직후 연결 실패 | `--bind tailscale`은 Tailscale이 올라올 때까지 최대 3분 기다립니다(`--tailscale-wait`). 그래도 안 되면 `server.log`를 보세요. 10분 주기 트리거가 다시 시도합니다 |
| 업데이트 재부팅 후 계속 안 됨 | 아무도 로그인하지 않은 상태일 수 있습니다. 한글 COM은 데스크톱 세션이 필수라 자동 로그인이 필요합니다 |
| 서버가 뭘 하는지 안 보임 | 창 없는 빌드라 정상입니다. `%LOCALAPPDATA%\hwp2pdf\server.log`를 보거나 맥 앱의 `연결 테스트`를 누르세요 |
| 배치 도중 네트워크 끊김 | 폴링은 커서 기반이라 자동 재개됩니다. 최종 실패한 파일만 CSV에 `FAILED`로 남고 나머지는 계속 진행합니다 |

---

## 7. 보안

- **공인 인터넷에 노출하지 마세요.** Tailscale, VPN, 또는 신뢰된 LAN에서만 사용합니다.
- 토큰은 로그에 남기지 않습니다. 클라이언트 설정 파일은 `0600`으로 저장됩니다.
- 토큰을 명령줄(`--token`)에 넣으면 프로세스 목록에 보일 수 있습니다.
  가능하면 환경변수나 저장된 설정을 쓰세요.
- Tailscale ACL로 접근을 더 좁힐 수 있습니다:

  ```json
  {
    "acls": [
      { "action": "accept", "src": ["user@example.com"], "dst": ["namun-ji:8765"] }
    ]
  }
  ```

- 공유 폴더 모드는 서버가 등록한 `--share-root` 안쪽만 접근합니다. 경로 탈출은 거부됩니다.

---

## English summary

macOS has no Hancom automation engine, so the mac app talks to a **conversion
server running on a Windows machine with Hancom Office installed**. File
discovery, skip/overwrite rules, the CSV log, progress and stop all stay on the
Mac; the server only converts one file to one format at a time.

**Server (Windows).** Run `hwp2pdf-cli.exe serve --init` once to mint a token
(stored in `%LOCALAPPDATA%\hwp2pdf\server_token.txt`). Then pick how it is
exposed:

- **Tailscale (recommended)** — `serve --bind tailscale` binds only this
  machine's 100.x address, so nothing is exposed on the LAN and no firewall rule
  is needed. Point the client at `http://<host>.<tailnet>.ts.net:8765`. For TLS,
  bind loopback and run `tailscale serve --bg 8765`.
- **LAN** — `serve --bind 0.0.0.0` plus a Private-profile firewall rule
  (`scripts/allow_firewall.ps1`). Traffic is plain HTTP; use `--tls-cert` /
  `--tls-key` or keep it on a trusted network.
- **Local Windows VM** — bind `0.0.0.0` inside the VM and use its NAT address,
  or forward host `127.0.0.1:8765` to the guest.

**Never register the server as a Windows Service.** Hangul automation needs an
interactive desktop session; Session 0 leaves zombie `Hwp.exe` processes. Use a
Scheduled Task with "run only when user is logged on"
(`scripts/install_serve_task.ps1`).

**Transports.** Uploads work everywhere. If the Mac and the Windows host see the
same storage, register it with `--share-root NAME=PATH` on the server and map the
mount in `shares` on the client; sources are then passed by path and outputs are
written in place, skipping both transfers. Paths that escape a share root are
rejected with 400.

**Client (macOS).** Fill in the `Conversion server` panel and press *Test
connection*, or use the CLI:

```bash
hwp2pdf ~/Documents/reports --pdf --docx --server http://host:8765 --token <token>
# or export HWP2PDF_SERVER_URL / HWP2PDF_TOKEN
```

Verify with `./scripts/check_macos.sh` and
`python scripts/smoke_remote.py <url> <token> <sample.hwp>`.

The wire contract is documented in [protocol.md](protocol.md).
