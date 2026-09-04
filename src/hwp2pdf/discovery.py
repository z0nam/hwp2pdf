"""Turning what a user knows into a conversion server address.

Someone setting this up for the first time knows, at best, the name of the
Windows box next to them -- and often not even that. Three paths lead from
there to a working address, in the order they are likely to succeed:

* an **invite string** the server prints, pasted into the address field. It
  carries the token too, and survives subnets, VPNs and port forwards, so it is
  the one path that works when nothing else does.
* a **bare host name or IP**, expanded here into a full URL. Windows registers
  its host name over mDNS, so on macOS ``namun-ji`` usually also answers as
  ``namun-ji.local``; both are tried.
* the **Find button**, which probes the machines this one can already see --
  Tailscale peers, then hosts in the local ARP table. When that finds nothing,
  a second, explicit press sweeps the attached /24 networks; it is never the
  first thing tried, because 254 connections per network look like a port scan
  to a corporate security appliance.

No discovery protocol of its own is needed: ``/v1/health`` answers without a
token and names the application, so any reachable server -- including one built
before this module existed -- identifies itself when asked.

Stdlib only, and every lookup here is best-effort: a missing ``tailscale``
binary, an unparseable ``arp`` table or an unreachable host yields an empty
result rather than an error.
"""

import base64
import concurrent.futures
import http.client
import ipaddress
import json
import re
import socket
import subprocess
import urllib.error
import urllib.parse
import urllib.request

from hwp2pdf import certs
from hwp2pdf.constants import APP_NAME
from hwp2pdf.server import protocol

#: Prefix of a pasted invite. Looks like a URL so chat clients keep it in one
#: piece, but it is ours: the payload is base64url, not a host name.
INVITE_SCHEME = "hwp2pdf://"

#: A probe is a single request to a machine that is either there or not, so it
#: fails fast; discovery fans many of them out at once.
PROBE_TIMEOUT = 1.5
PROBE_WORKERS = 24

#: Sweeping trades patience for breadth: hundreds of addresses, nearly all of
#: them dead, so each gets less time and more of them run at once.
SWEEP_TIMEOUT = 0.6
SWEEP_WORKERS = 64
#: A machine on more networks than this is a router or a lab, not the setup
#: this feature is for; sweeping all of them would take longer than it is worth.
SWEEP_MAX_NETWORKS = 3

#: What an office or home LAN actually looks like. Spelled out rather than
#: leaning on ``is_private``, which also accepts the documentation ranges and
#: the shared-address space Tailscale hands out -- neither is a subnet anyone
#: wants swept.
LAN_NETWORKS = (
    ipaddress.ip_network("10.0.0.0/8"),
    ipaddress.ip_network("172.16.0.0/12"),
    ipaddress.ip_network("192.168.0.0/16"),
)

TAILSCALE_BINARIES = (
    "tailscale",
    r"C:\Program Files\Tailscale\tailscale.exe",
    r"C:\Program Files (x86)\Tailscale\tailscale.exe",
    "/usr/local/bin/tailscale",
    "/opt/homebrew/bin/tailscale",
)

VIA_TAILSCALE = "tailscale"
VIA_LAN = "lan"

_IPV4 = re.compile(r"\b(\d{1,3}(?:\.\d{1,3}){3})\b")
#: A host name urllib will accept. Anything else -- a space, a control
#: character -- is a typo, and is rejected here rather than deep in http.client.
_HOSTNAME = re.compile(r"^[a-z0-9_](?:[a-z0-9_\-.]*[a-z0-9_])?$", re.IGNORECASE)
#: An address as the platform's own network tool prints it: "inet 1.2.3.4"
#: (ifconfig, ip addr), "inet addr:1.2.3.4" (old Linux), or the "IPv4
#: Address . . . : 1.2.3.4" line ipconfig prints in any language.
_INET = re.compile(
    r"(?:\binet\s+(?:addr:)?|IPv4[^:\n]*:\s*)(\d{1,3}(?:\.\d{1,3}){3})"
)


# -- addresses -----------------------------------------------------------

def normalize_server_url(text: str) -> str:
    """Expand what a user typed into a full base URL, or "" if unusable.

    A bare host or IP gains ``http://`` and the default port. A URL typed with
    a scheme is taken at its word -- no port is injected, because someone who
    wrote ``https://convert.example.com`` means port 443, not 17650.
    """
    raw = (text or "").strip()
    if not raw:
        return ""

    typed_scheme = "://" in raw
    try:
        parsed = urllib.parse.urlsplit(raw if typed_scheme else f"http://{raw}")
        host = parsed.hostname
        port = parsed.port
    except ValueError:
        # A bare IPv6 address without brackets lands here; so does "host:garbage".
        return ""
    if parsed.scheme not in ("http", "https") or not host:
        return ""

    if ":" in host:
        if not _is_ip(host):
            return ""
        netloc = f"[{host}]"
    elif not _HOSTNAME.match(host):
        return ""
    else:
        netloc = host
    if port is not None:
        netloc = f"{netloc}:{port}"
    elif not typed_scheme:
        netloc = f"{netloc}:{protocol.DEFAULT_PORT}"

    return urllib.parse.urlunsplit((parsed.scheme, netloc, parsed.path.rstrip("/"), "", ""))


def url_candidates(text: str) -> list:
    """Addresses worth trying for what the user typed, best first.

    A single-label name gets an mDNS variant: Windows publishes its host name
    as ``<name>.local``, which is how a Mac resolves it without a DNS server on
    the network.
    """
    primary = normalize_server_url(text)
    if not primary:
        return []

    parsed = urllib.parse.urlsplit(primary)
    host = parsed.hostname or ""
    if "." in host or ":" in host or _is_ip(host):
        return [primary]

    mdns = urllib.parse.urlunsplit((
        parsed.scheme,
        f"{host}.local:{parsed.port}" if parsed.port else f"{host}.local",
        parsed.path,
        "",
        "",
    ))
    return [primary, mdns]


def _is_ip(host: str) -> bool:
    try:
        ipaddress.ip_address(host)
    except ValueError:
        return False
    return True


# -- invites -------------------------------------------------------------

def make_invite(url: str, token: str = "") -> str:
    """One paste-able string carrying both halves of a connection."""
    payload = {"u": normalize_server_url(url) or (url or "").strip(), "t": token or ""}
    blob = base64.urlsafe_b64encode(
        json.dumps(payload, separators=(",", ":")).encode("utf-8")
    ).decode("ascii")
    # Padding is restored on parse; "=" is what chat clients most often eat.
    return INVITE_SCHEME + blob.rstrip("=")


def parse_invite(text: str):
    """Decode an invite into ``{"url", "token"}``, or None if it is not one.

    Returns None for anything that is not an invite *and* for a corrupt one --
    both mean "this is not an address I can use", and the caller reports the
    same thing either way.
    """
    raw = (text or "").strip()
    if not raw.lower().startswith(INVITE_SCHEME):
        return None

    blob = raw[len(INVITE_SCHEME):].strip().strip("/")
    padded = blob + "=" * (-len(blob) % 4)
    try:
        payload = json.loads(base64.urlsafe_b64decode(padded.encode("ascii")).decode("utf-8"))
    except (ValueError, UnicodeDecodeError):
        return None
    if not isinstance(payload, dict):
        return None

    url = normalize_server_url(payload.get("u") or "")
    if not url:
        return None
    token = payload.get("t")
    return {"url": url, "token": token if isinstance(token, str) else ""}


# -- probing -------------------------------------------------------------

def probe(url: str, timeout: float = PROBE_TIMEOUT):
    """Ask one address whether a conversion server lives there.

    ``/v1/health`` needs no token by design, so this identifies password-
    protected servers too -- the token is only needed to convert.
    """
    base = normalize_server_url(url)
    if not base:
        return None
    try:
        with certs.urlopen(base + protocol.PATH_HEALTH, timeout=timeout) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except (urllib.error.URLError, http.client.HTTPException, OSError, ValueError):
        return None
    if not isinstance(payload, dict) or payload.get("app") != APP_NAME:
        return None

    return {
        "url": base,
        "version": str(payload.get("version") or ""),
        "api": payload.get("api"),
        "auth_required": bool(payload.get("auth_required")),
        "compatible": payload.get("api") == protocol.API_VERSION,
    }


# -- candidate sources ---------------------------------------------------

def _run(command, timeout=10):
    """Run a helper binary, returning its stdout or "" if it is not usable."""
    try:
        result = subprocess.run(
            command, capture_output=True, text=True, timeout=timeout,
            errors="replace",
            # Windows would otherwise flash a console window for each probe.
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except (OSError, subprocess.SubprocessError):
        return ""
    return result.stdout if result.returncode == 0 else ""


def tailscale_peers() -> list:
    """Online Tailscale peers as ``{"name", "address"}``, newest CLI or not."""
    for binary in TAILSCALE_BINARIES:
        raw = _run([binary, "status", "--json"])
        if not raw:
            continue
        try:
            status = json.loads(raw)
        except ValueError:
            continue

        peers = []
        for peer in (status.get("Peer") or {}).values():
            if not isinstance(peer, dict) or not peer.get("Online"):
                continue
            address = next(
                (ip for ip in (peer.get("TailscaleIPs") or []) if _is_ipv4(ip)), ""
            )
            if address:
                peers.append({
                    "name": (peer.get("HostName") or "").split(".")[0],
                    "address": address,
                })
        return peers
    return []


def arp_neighbours() -> list:
    """IPv4 addresses this machine has recently talked to on the local network.

    Cheap and quiet -- it reads a table the OS already keeps rather than
    touching every address in the subnet -- but it only knows hosts there has
    been traffic with, so a server that has been idle may be missing.
    """
    # "-n" matters: plain "arp -a" reverse-resolves every entry, which on a Mac
    # full of AWDL peers takes longer than the whole discovery budget. Windows
    # arp has no -n and is numeric anyway, so it falls through to "arp -a".
    raw = _run(["arp", "-an"]) or _run(["arp", "-a"])
    seen = []
    for address in _IPV4.findall(raw):
        if address in seen or not _is_ipv4(address):
            continue
        parsed = ipaddress.ip_address(address)
        # Link-local hits are AirDrop/AWDL peers, never a server someone set up.
        if (parsed.is_multicast or parsed.is_unspecified
                or parsed.is_loopback or parsed.is_link_local):
            continue
        if address.endswith(".255"):
            continue
        seen.append(address)
    return seen


def _is_ipv4(value: str) -> bool:
    try:
        return isinstance(ipaddress.ip_address(value), ipaddress.IPv4Address)
    except ValueError:
        return False


def local_networks() -> list:
    """The private /24 networks this machine is attached to.

    The default route is not enough: a Mac on both Wi-Fi and Ethernet reaches
    the conversion server over whichever one the server happens to share, which
    may not be the one carrying the default route. Nor is the host name, which
    on a tailnet resolves to the Tailscale address alone. So the platform's own
    tool is asked, and every private address it names contributes its /24.

    Only the network matters, never the mask, so a stray broadcast address in
    the output is harmless -- it names the same /24 -- and a dotted netmask is
    dropped for not being private.
    """
    raw = "\n".join(filter(None, (
        _run(["ip", "-4", "-o", "addr"]),   # modern Linux
        _run(["ifconfig"]),                 # macOS, BSD, older Linux
        _run(["ipconfig"]),                 # Windows
    )))

    networks = []
    for address in _INET.findall(raw):
        try:
            parsed = ipaddress.ip_address(address)
        except ValueError:
            continue
        if not any(parsed in network for network in LAN_NETWORKS):
            continue
        network = ipaddress.ip_network(f"{address}/24", strict=False)
        if network not in networks:
            networks.append(network)
    return networks[:SWEEP_MAX_NETWORKS]


def subnet_hosts() -> list:
    """Every address on the attached /24 networks, this machine excluded."""
    mine = {address for address in _INET.findall(
        "\n".join(filter(None, (
            _run(["ip", "-4", "-o", "addr"]), _run(["ifconfig"]), _run(["ipconfig"]),
        )))
    )}
    hosts = []
    for network in local_networks():
        hosts.extend(str(host) for host in network.hosts() if str(host) not in mine)
    return hosts


def _reverse_name(address: str) -> str:
    try:
        return socket.gethostbyaddr(address)[0].split(".")[0]
    except OSError:
        return ""


# -- discovery -----------------------------------------------------------

def candidates(wide: bool = False) -> dict:
    """Addresses worth probing, mapped to how they were found.

    ``wide`` adds every address on the attached /24 networks. It is the
    escalation the user asks for after a quiet search, never the default.
    """
    found = {}
    for peer in tailscale_peers():
        url = f"http://{peer['address']}:{protocol.DEFAULT_PORT}"
        found.setdefault(url, {"name": peer["name"], "via": VIA_TAILSCALE})
    for address in arp_neighbours():
        url = f"http://{address}:{protocol.DEFAULT_PORT}"
        found.setdefault(url, {"name": "", "via": VIA_LAN})
    if wide:
        for address in subnet_hosts():
            url = f"http://{address}:{protocol.DEFAULT_PORT}"
            found.setdefault(url, {"name": "", "via": VIA_LAN})
    return found


def discover(timeout: float = PROBE_TIMEOUT, workers: int = PROBE_WORKERS,
             wide: bool = False) -> list:
    """Every conversion server this machine can currently reach.

    Compatible servers sort first, then Tailscale peers, then by name: the
    entry a user most likely wants is the one at the top.
    """
    targets = candidates(wide)
    if not targets:
        return []

    servers = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=min(workers, len(targets))) as pool:
        probes = {pool.submit(probe, url, timeout): url for url in targets}
        for future in concurrent.futures.as_completed(probes):
            health = future.result()
            if not health:
                continue
            source = targets[probes[future]]
            host = urllib.parse.urlsplit(health["url"]).hostname or ""
            servers.append({
                **health,
                "via": source["via"],
                "name": source["name"] or _reverse_name(host) or host,
            })

    servers.sort(key=lambda s: (not s["compatible"], s["via"] != VIA_TAILSCALE, s["name"]))
    return servers
