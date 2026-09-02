"""Getting from what a user knows to an address the client can open.

The probing tests run against the real server so they cannot drift from the
``/v1/health`` contract discovery depends on; the source tests replace the
helper binaries, since a runner has neither Tailscale nor a populated ARP table.
"""

import json
import threading

import pytest

from fakes import FakeBackend

from hwp2pdf import discovery
from hwp2pdf.server import protocol
from hwp2pdf.server.http_server import create_server

DEFAULT = protocol.DEFAULT_PORT


@pytest.fixture
def health_server():
    """A real conversion server on a loopback port, with no Hangul behind it."""
    httpd = create_server(
        "127.0.0.1", 0,
        backend_factory=FakeBackend,
        hwp_probe=lambda: {"installed": False, "detail": "test", "running": []},
        token="",
        quiet=True,
    )
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()
    try:
        host, port = httpd.server_address[:2]
        yield f"http://{host}:{port}"
    finally:
        httpd.shutdown()
        httpd.server_close()
        thread.join(timeout=5)


# -- normalization -------------------------------------------------------

@pytest.mark.parametrize("typed, expected", [
    ("namun-ji", f"http://namun-ji:{DEFAULT}"),
    ("NAMUN-JI", f"http://namun-ji:{DEFAULT}"),
    ("  namun-ji  ", f"http://namun-ji:{DEFAULT}"),
    ("namun-ji:9000", "http://namun-ji:9000"),
    ("192.168.0.5", f"http://192.168.0.5:{DEFAULT}"),
    ("192.168.0.5:80", "http://192.168.0.5:80"),
    ("[fd7a::1]:17650", "http://[fd7a::1]:17650"),
    ("http://host:8765/", "http://host:8765"),
])
def test_what_a_user_types_becomes_a_usable_url(typed, expected):
    assert discovery.normalize_server_url(typed) == expected


def test_a_typed_scheme_is_taken_at_its_word():
    # Someone who wrote https means 443 behind a proxy, not the default port.
    assert discovery.normalize_server_url("https://convert.example.com") == (
        "https://convert.example.com"
    )


@pytest.mark.parametrize("typed", [
    "", "   ", "ftp://host", "namun-ji:abc",
    "fd7a::1",  # bare IPv6 is ambiguous with host:port; brackets are required
])
def test_unusable_addresses_are_rejected_rather_than_guessed(typed):
    assert discovery.normalize_server_url(typed) == ""


def test_a_single_label_name_also_gets_an_mdns_spelling():
    # Windows publishes its host name over mDNS, which is how a Mac finds it.
    assert discovery.url_candidates("namun-ji") == [
        f"http://namun-ji:{DEFAULT}",
        f"http://namun-ji.local:{DEFAULT}",
    ]


@pytest.mark.parametrize("typed", ["192.168.0.5", "host.example.com", "[fd7a::1]:1"])
def test_addresses_that_already_resolve_get_no_mdns_variant(typed):
    assert len(discovery.url_candidates(typed)) == 1


def test_nothing_typed_yields_no_candidates():
    assert discovery.url_candidates("") == []


# -- invites -------------------------------------------------------------

def test_an_invite_carries_both_halves_of_a_connection():
    invite = discovery.make_invite("namun-ji", "s3cr3t")
    assert invite.startswith(discovery.INVITE_SCHEME)
    assert discovery.parse_invite(invite) == {
        "url": f"http://namun-ji:{DEFAULT}",
        "token": "s3cr3t",
    }


def test_an_invite_without_a_token_round_trips():
    invite = discovery.make_invite(f"http://100.64.0.1:{DEFAULT}")
    assert discovery.parse_invite(invite) == {
        "url": f"http://100.64.0.1:{DEFAULT}",
        "token": "",
    }


def test_an_invite_survives_lost_padding():
    # "=" is what chat clients and copy-paste most often drop.
    invite = discovery.make_invite("namun-ji", "s3cr3t")
    assert discovery.parse_invite(invite + "==") == discovery.parse_invite(invite)


@pytest.mark.parametrize("text", [
    "",
    "http://namun-ji:17650",          # a plain address is not an invite
    "hwp2pdf://",
    "hwp2pdf://!!!not-base64!!!",
    "hwp2pdf://" + "eyJ4IjoxfQ",      # valid base64, but no address in it
])
def test_anything_that_is_not_a_usable_invite_returns_none(text):
    assert discovery.parse_invite(text) is None


# -- probing -------------------------------------------------------------

def test_probe_identifies_a_real_server(health_server):
    found = discovery.probe(health_server)
    assert found["url"] == health_server
    assert found["api"] == protocol.API_VERSION
    assert found["compatible"] is True
    assert found["auth_required"] is False


def test_probe_reaches_a_server_named_without_a_scheme_or_port(health_server):
    # The GUI hands over whatever the user typed, so probe must normalize too.
    bare = health_server.removeprefix("http://")
    assert discovery.probe(bare)["url"] == health_server


def test_a_token_protected_server_still_identifies_itself():
    # Health is deliberately unauthenticated: discovery must see locked servers.
    httpd = create_server(
        "127.0.0.1", 0,
        backend_factory=FakeBackend,
        hwp_probe=lambda: {"installed": False, "detail": "test", "running": []},
        token="secret",
        quiet=True,
    )
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()
    try:
        host, port = httpd.server_address[:2]
        found = discovery.probe(f"http://{host}:{port}")
        assert found["auth_required"] is True
    finally:
        httpd.shutdown()
        httpd.server_close()
        thread.join(timeout=5)


@pytest.mark.parametrize("url", ["", "not a url", "http://127.0.0.1:1"])
def test_probing_something_that_is_not_a_server_returns_none(url):
    assert discovery.probe(url, timeout=0.5) is None


def test_probe_ignores_a_stranger_answering_on_the_port(monkeypatch):
    class Response:
        def read(self):
            return json.dumps({"app": "something else", "api": 1}).encode()

        def __enter__(self):
            return self

        def __exit__(self, *_exc):
            return False

    monkeypatch.setattr(discovery.urllib.request, "urlopen", lambda *a, **k: Response())
    assert discovery.probe("http://host:17650") is None


# -- candidate sources ---------------------------------------------------

TAILSCALE_STATUS = json.dumps({
    "Peer": {
        "nodekey:a": {"HostName": "namun-ji", "Online": True,
                      "TailscaleIPs": ["100.124.117.75", "fd7a::1"]},
        "nodekey:b": {"HostName": "offline-box", "Online": False,
                      "TailscaleIPs": ["100.64.0.9"]},
        "nodekey:c": {"HostName": "v6-only", "Online": True,
                      "TailscaleIPs": ["fd7a::2"]},
    }
})


def test_only_online_peers_with_an_ipv4_are_offered(monkeypatch):
    monkeypatch.setattr(discovery, "_run", lambda *a, **k: TAILSCALE_STATUS)
    assert discovery.tailscale_peers() == [
        {"name": "namun-ji", "address": "100.124.117.75"}
    ]


@pytest.mark.parametrize("output", ["", "not json at all"])
def test_a_missing_or_broken_tailscale_is_not_an_error(monkeypatch, output):
    monkeypatch.setattr(discovery, "_run", lambda *a, **k: output)
    assert discovery.tailscale_peers() == []


ARP_OUTPUT = """\
? (192.168.8.1) at 0:11:22:33:44:55 on en0 ifscope [ethernet]
? (192.168.8.20) at 8c:85:90:aa:bb:cc on en0 ifscope [ethernet]
? (192.168.8.20) at 8c:85:90:aa:bb:cc on en1 ifscope [ethernet]
? (169.254.7.7) at f4:ce:23:c7:2c:52 on en0 [ethernet]
? (224.0.0.251) at 1:0:5e:0:0:fb on en0 ifscope permanent [ethernet]
? (192.168.8.255) at ff:ff:ff:ff:ff:ff on en0 ifscope [ethernet]
"""


def test_arp_yields_real_neighbours_only(monkeypatch):
    monkeypatch.setattr(discovery, "_run", lambda *a, **k: ARP_OUTPUT)
    # Link-local is AirDrop, multicast and broadcast are not hosts, and the
    # same address on two interfaces is still one machine.
    assert discovery.arp_neighbours() == ["192.168.8.1", "192.168.8.20"]


def test_an_unreadable_arp_table_is_not_an_error(monkeypatch):
    monkeypatch.setattr(discovery, "_run", lambda *a, **k: "")
    assert discovery.arp_neighbours() == []


# -- sweeping --------------------------------------------------------------

IFCONFIG = """\
en1: flags=8863 mtu 1500
\tinet 192.168.8.2 netmask 0xffffff00 broadcast 192.168.8.255
en0: flags=8863 mtu 1500
\tinet 192.168.161.11 netmask 0xffffff00 broadcast 192.168.161.255
bridge0: flags=8863 mtu 1500
\tinet 169.254.10.82 netmask 0xffff0000
utun5: flags=8051 mtu 1280
\tinet 100.123.230.95 --> 100.123.230.95 netmask 0xffffffff
lo0: flags=8049 mtu 16384
\tinet 127.0.0.1 netmask 0xff000000
"""


def test_every_attached_network_is_swept_not_just_the_default_route(monkeypatch):
    # A Mac on Wi-Fi and Ethernet reaches the server over whichever one it
    # shares, which need not be the interface carrying the default route.
    monkeypatch.setattr(discovery, "_run", lambda *a, **k: IFCONFIG)
    assert [str(net) for net in discovery.local_networks()] == [
        "192.168.8.0/24", "192.168.161.0/24",
    ]


def test_tailscale_loopback_and_link_local_are_never_swept(monkeypatch):
    monkeypatch.setattr(discovery, "_run", lambda *a, **k: IFCONFIG)
    swept = {str(net) for net in discovery.local_networks()}
    assert not any(net.startswith(("100.", "169.254.", "127.")) for net in swept)


@pytest.mark.parametrize("address", [
    "203.0.113.7",    # documentation range -- ipaddress calls this "private"
    "8.8.8.8",        # genuinely public
    "100.123.230.95",  # Tailscale's shared-address space
    "169.254.10.82",  # link-local
    "127.0.0.1",
])
def test_only_real_lan_ranges_are_swept(monkeypatch, address):
    monkeypatch.setattr(discovery, "_run", lambda *a, **k: f"\tinet {address} netmask 0xffffff00")
    assert discovery.local_networks() == []


def test_only_a_few_networks_are_swept(monkeypatch):
    many = "\n".join(f"\tinet 10.{n}.0.5 netmask 0xffffff00" for n in range(10))
    monkeypatch.setattr(discovery, "_run", lambda *a, **k: many)
    assert len(discovery.local_networks()) == discovery.SWEEP_MAX_NETWORKS


def test_the_sweep_covers_both_networks_and_skips_this_machine(monkeypatch):
    monkeypatch.setattr(discovery, "_run", lambda *a, **k: IFCONFIG)
    hosts = discovery.subnet_hosts()
    assert len(hosts) == 2 * 254 - 2          # both /24s, less our two addresses
    assert "192.168.8.1" in hosts and "192.168.161.254" in hosts
    assert "192.168.8.2" not in hosts and "192.168.161.11" not in hosts


def test_a_machine_on_no_private_network_has_nothing_to_sweep(monkeypatch):
    monkeypatch.setattr(discovery, "_run", lambda *a, **k: "")
    assert discovery.subnet_hosts() == []


def test_sweeping_is_never_what_the_first_search_does(monkeypatch):
    monkeypatch.setattr(discovery, "tailscale_peers", list)
    monkeypatch.setattr(discovery, "arp_neighbours", lambda: ["192.168.8.9"])
    monkeypatch.setattr(discovery, "subnet_hosts", lambda: ["192.168.8.77"])

    assert len(discovery.candidates()) == 1
    assert len(discovery.candidates(wide=True)) == 2


# -- discovery -----------------------------------------------------------

def test_discovery_finds_a_running_server_among_dead_candidates(health_server, monkeypatch):
    host, _, port = health_server.removeprefix("http://").rpartition(":")
    monkeypatch.setattr(discovery, "tailscale_peers", list)
    monkeypatch.setattr(discovery, "arp_neighbours", lambda: ["203.0.113.9", host])
    # The port the fixture got is arbitrary, so candidates() is bypassed here.
    monkeypatch.setattr(discovery, "candidates", lambda wide=False: {
        f"http://203.0.113.9:{port}": {"name": "", "via": discovery.VIA_LAN},
        health_server: {"name": "stub", "via": discovery.VIA_LAN},
    })

    found = discovery.discover(timeout=0.5)
    assert [server["url"] for server in found] == [health_server]
    assert found[0]["name"] == "stub"


def test_discovery_with_nothing_to_probe_is_empty(monkeypatch):
    monkeypatch.setattr(discovery, "candidates", lambda wide=False: {})
    assert discovery.discover() == []


def test_compatible_tailscale_peers_sort_first(monkeypatch):
    rows = {
        "http://a:1": {"name": "lan-old", "via": discovery.VIA_LAN},
        "http://b:1": {"name": "tail", "via": discovery.VIA_TAILSCALE},
        "http://c:1": {"name": "lan-new", "via": discovery.VIA_LAN},
    }
    health = {
        "http://a:1": {"compatible": False},
        "http://b:1": {"compatible": True},
        "http://c:1": {"compatible": True},
    }
    monkeypatch.setattr(discovery, "candidates", lambda wide=False: rows)
    monkeypatch.setattr(discovery, "probe", lambda url, timeout=0: {
        "url": url, "version": "1", "api": 1, "auth_required": False,
        **health[url],
    })
    assert [s["name"] for s in discovery.discover()] == ["tail", "lan-new", "lan-old"]
