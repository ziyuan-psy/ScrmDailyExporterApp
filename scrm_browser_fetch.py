#!/usr/bin/env python3
"""Same-origin SCRM fetch through the logged-in Chrome page.

This is a fallback for SCRM endpoints that reject direct Python HTTPS requests
but accept the same request from the authenticated browser origin.
"""

from __future__ import annotations

import base64
import hashlib
import json
import os
import socket
import struct
import time
from typing import Any, Dict, List, Optional
from urllib.parse import urlparse
from urllib.request import urlopen


SCRM_HOST = "scrm.cotticoffee.cc"


class BrowserFetchError(RuntimeError):
    pass


class WebSocketConnection:
    def __init__(self, url: str, timeout: float = 10.0) -> None:
        parsed = urlparse(url)
        if parsed.scheme != "ws":
            raise BrowserFetchError(f"Only ws:// CDP URLs are supported: {url}")
        self.host = parsed.hostname or "127.0.0.1"
        self.port = parsed.port or 80
        self.path = parsed.path
        if parsed.query:
            self.path = f"{self.path}?{parsed.query}"
        self.timeout = timeout
        self.sock: Optional[socket.socket] = None

    def __enter__(self) -> "WebSocketConnection":
        self.sock = socket.create_connection((self.host, self.port), timeout=self.timeout)
        self.sock.settimeout(self.timeout)
        key = base64.b64encode(os.urandom(16)).decode("ascii")
        request = (
            f"GET {self.path} HTTP/1.1\r\n"
            f"Host: {self.host}:{self.port}\r\n"
            "Upgrade: websocket\r\n"
            "Connection: Upgrade\r\n"
            f"Sec-WebSocket-Key: {key}\r\n"
            "Sec-WebSocket-Version: 13\r\n"
            "\r\n"
        )
        self.sock.sendall(request.encode("ascii"))
        response = self._recv_until(b"\r\n\r\n")
        if b" 101 " not in response.split(b"\r\n", 1)[0]:
            raise BrowserFetchError("Chrome CDP websocket handshake failed")
        accept = base64.b64encode(
            hashlib.sha1((key + "258EAFA5-E914-47DA-95CA-C5AB0DC85B11").encode("ascii")).digest()
        ).decode("ascii")
        if accept.encode("ascii") not in response:
            raise BrowserFetchError("Chrome CDP websocket accept header mismatch")
        return self

    def __exit__(self, *_: Any) -> None:
        if self.sock:
            try:
                self.sock.close()
            finally:
                self.sock = None

    def _recv_exact(self, size: int) -> bytes:
        if not self.sock:
            raise BrowserFetchError("Websocket is not connected")
        chunks: List[bytes] = []
        remaining = size
        while remaining:
            chunk = self.sock.recv(remaining)
            if not chunk:
                raise BrowserFetchError("Websocket closed unexpectedly")
            chunks.append(chunk)
            remaining -= len(chunk)
        return b"".join(chunks)

    def _recv_until(self, marker: bytes) -> bytes:
        if not self.sock:
            raise BrowserFetchError("Websocket is not connected")
        data = bytearray()
        while marker not in data:
            chunk = self.sock.recv(4096)
            if not chunk:
                raise BrowserFetchError("Connection closed before websocket handshake finished")
            data.extend(chunk)
        return bytes(data)

    def send_text(self, text: str) -> None:
        if not self.sock:
            raise BrowserFetchError("Websocket is not connected")
        payload = text.encode("utf-8")
        if len(payload) < 126:
            header = bytes([0x81, 0x80 | len(payload)])
        elif len(payload) < 65536:
            header = bytes([0x81, 0x80 | 126]) + struct.pack("!H", len(payload))
        else:
            header = bytes([0x81, 0x80 | 127]) + struct.pack("!Q", len(payload))
        mask = os.urandom(4)
        masked = bytes(byte ^ mask[index % 4] for index, byte in enumerate(payload))
        self.sock.sendall(header + mask + masked)

    def send_pong(self, payload: bytes) -> None:
        if not self.sock:
            return
        if len(payload) > 125:
            payload = payload[:125]
        mask = os.urandom(4)
        masked = bytes(byte ^ mask[index % 4] for index, byte in enumerate(payload))
        self.sock.sendall(bytes([0x8A, 0x80 | len(payload)]) + mask + masked)

    def recv_text(self) -> str:
        fragments: List[bytes] = []
        while True:
            first, second = self._recv_exact(2)
            fin = bool(first & 0x80)
            opcode = first & 0x0F
            masked = bool(second & 0x80)
            length = second & 0x7F
            if length == 126:
                length = struct.unpack("!H", self._recv_exact(2))[0]
            elif length == 127:
                length = struct.unpack("!Q", self._recv_exact(8))[0]
            mask_key = self._recv_exact(4) if masked else b""
            payload = self._recv_exact(length) if length else b""
            if masked:
                payload = bytes(byte ^ mask_key[index % 4] for index, byte in enumerate(payload))

            if opcode == 0x8:
                raise BrowserFetchError("Chrome CDP websocket closed")
            if opcode == 0x9:
                self.send_pong(payload)
                continue
            if opcode == 0xA:
                continue
            if opcode in (0x1, 0x0):
                fragments.append(payload)
                if fin:
                    return b"".join(fragments).decode("utf-8")


class CdpClient:
    def __init__(self, websocket_url: str) -> None:
        self.websocket_url = websocket_url
        self.next_id = 0

    def __enter__(self) -> "CdpClient":
        self.ws = WebSocketConnection(self.websocket_url).__enter__()
        return self

    def __exit__(self, *args: Any) -> None:
        self.ws.__exit__(*args)

    def call(self, method: str, params: Optional[Dict[str, Any]] = None, timeout: float = 10.0) -> Dict[str, Any]:
        self.next_id += 1
        message_id = self.next_id
        self.ws.send_text(json.dumps({"id": message_id, "method": method, "params": params or {}}))
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            data = json.loads(self.ws.recv_text())
            if data.get("id") != message_id:
                continue
            if "error" in data:
                raise BrowserFetchError(f"CDP {method} failed: {data['error']}")
            return data.get("result") or {}
        raise BrowserFetchError(f"CDP {method} timed out")


def find_scrm_tab(port: int) -> Dict[str, Any]:
    tabs = json.loads(urlopen(f"http://127.0.0.1:{port}/json/list", timeout=3).read().decode("utf-8"))
    for tab in tabs:
        if tab.get("type") == "page" and SCRM_HOST in (tab.get("url") or ""):
            websocket_url = tab.get("webSocketDebuggerUrl")
            if websocket_url:
                return tab
    raise BrowserFetchError(f"No SCRM page tab found on Chrome debug port {port}")


def browser_safe_headers(headers: Dict[str, str]) -> Dict[str, str]:
    forbidden = {
        "cookie",
        "host",
        "origin",
        "referer",
        "user-agent",
        "content-length",
        "accept-encoding",
        "connection",
    }
    return {key: value for key, value in headers.items() if key.lower() not in forbidden}


def post_json(
    port: int,
    path: str,
    payload: Dict[str, Any],
    headers: Dict[str, str],
    referer: str,
    timeout: float,
) -> str:
    tab = find_scrm_tab(port)
    expression = f"""
(async () => {{
  const path = {json.dumps(path)};
  const payload = {json.dumps(payload, ensure_ascii=False)};
  const headers = {json.dumps(browser_safe_headers(headers), ensure_ascii=False)};
  const currentToken = sessionStorage.getItem('current-token') || '';
  if (currentToken) headers.Authorization = currentToken.startsWith('Bearer ') ? currentToken : 'Bearer ' + currentToken;
  const res = await fetch(path, {{
    method: 'POST',
    headers,
    referrer: {json.dumps(referer)},
    body: JSON.stringify(payload)
  }});
  const text = await res.text();
  return {{status: res.status, text}};
}})()
"""
    with CdpClient(tab["webSocketDebuggerUrl"]) as client:
        result = client.call(
            "Runtime.evaluate",
            {"expression": expression, "returnByValue": True, "awaitPromise": True},
            timeout=max(10.0, timeout),
        )
    value = ((result.get("result") or {}).get("value")) or {}
    status = int(value.get("status") or 0)
    text = str(value.get("text") or "")
    if status >= 400:
        raise BrowserFetchError(f"Browser fetch HTTP {status} for {path}: {text[:300]}")
    return text
