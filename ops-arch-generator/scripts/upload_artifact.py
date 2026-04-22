#!/usr/bin/env python3
"""Upload generated architecture artifacts to a cloud target."""

from __future__ import annotations

import argparse
import base64
import json
import mimetypes
import os
import shutil
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path


def _env(name: str, fallback: str | None = None) -> str | None:
    value = os.getenv(name)
    if value is None or value == "":
        return fallback
    return value


def load_upload_config(args: argparse.Namespace | None = None) -> dict:
    args = args or argparse.Namespace()
    return {
        "mode": getattr(args, "upload_mode", None) or getattr(args, "mode", None) or _env("OPS_ARCH_UPLOAD_MODE", "none"),
        "target": getattr(args, "upload_target", None) or getattr(args, "target", None) or _env("OPS_ARCH_UPLOAD_TARGET"),
        "url": getattr(args, "upload_url", None) or getattr(args, "url", None) or _env("OPS_ARCH_UPLOAD_URL"),
        "username": getattr(args, "upload_username", None) or getattr(args, "username", None) or _env("OPS_ARCH_UPLOAD_USERNAME"),
        "password": getattr(args, "upload_password", None) or getattr(args, "password", None) or _env("OPS_ARCH_UPLOAD_PASSWORD"),
        "token": getattr(args, "upload_token", None) or getattr(args, "token", None) or _env("OPS_ARCH_UPLOAD_TOKEN"),
        "timeout": int(getattr(args, "upload_timeout", None) or getattr(args, "timeout", None) or _env("OPS_ARCH_UPLOAD_TIMEOUT", "60")),
    }


def _content_type(path: Path) -> str:
    guessed, _ = mimetypes.guess_type(path.name)
    return guessed or "application/octet-stream"


def _copy_upload(file_path: Path, target: str) -> dict:
    if not target:
        raise ValueError("copy mode requires a target path")

    raw_target = Path(target).expanduser()
    if raw_target.exists() and raw_target.is_dir():
        raw_target.mkdir(parents=True, exist_ok=True)
        destination = raw_target / file_path.name
    elif raw_target.suffix:
        raw_target.parent.mkdir(parents=True, exist_ok=True)
        destination = raw_target
    else:
        raw_target.mkdir(parents=True, exist_ok=True)
        destination = raw_target / file_path.name

    shutil.copy2(file_path, destination)
    return {
        "mode": "copy",
        "location": str(destination.resolve()),
        "status": "uploaded",
    }


def _resolve_remote_url(file_path: Path, url: str | None, target: str | None) -> str:
    remote_url = url or target
    if not remote_url:
        raise ValueError("remote upload requires a URL or target")
    if remote_url.endswith("/"):
        return remote_url + urllib.parse.quote(file_path.name)
    return remote_url


def _put_upload(
    file_path: Path,
    mode: str,
    url: str | None,
    target: str | None,
    username: str | None,
    password: str | None,
    token: str | None,
    timeout: int,
) -> dict:
    remote_url = _resolve_remote_url(file_path, url, target)
    data = file_path.read_bytes()
    request = urllib.request.Request(remote_url, data=data, method="PUT")
    request.add_header("Content-Type", _content_type(file_path))

    if username or password:
        auth = f"{username or ''}:{password or ''}".encode("utf-8")
        request.add_header("Authorization", "Basic " + base64.b64encode(auth).decode("ascii"))
    if token:
        request.add_header("Authorization", f"Bearer {token}")

    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            status = getattr(response, "status", response.getcode())
    except urllib.error.HTTPError as exc:
        raise RuntimeError(f"{mode} upload failed with HTTP {exc.code}: {exc.reason}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"{mode} upload failed: {exc.reason}") from exc

    return {
        "mode": mode,
        "location": remote_url,
        "status": "uploaded",
        "http_status": status,
    }


def upload_file(
    file_path: str | Path,
    mode: str = "none",
    target: str | None = None,
    url: str | None = None,
    username: str | None = None,
    password: str | None = None,
    token: str | None = None,
    timeout: int = 60,
) -> dict:
    path = Path(file_path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(path)

    selected_mode = (mode or "none").strip().lower()
    if selected_mode in {"", "none"}:
        return {"mode": "none", "location": str(path), "status": "local-only"}
    if selected_mode == "copy":
        return _copy_upload(path, target or "")
    if selected_mode in {"webdav", "http-put"}:
        return _put_upload(path, selected_mode, url, target, username, password, token, timeout)
    raise ValueError(f"Unsupported upload mode: {selected_mode}")


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Upload a generated architecture artifact.")
    parser.add_argument("--file", required=True, help="Artifact file path")
    parser.add_argument("--mode", help="Upload mode: none, copy, webdav, http-put")
    parser.add_argument("--target", help="Local target path or remote destination")
    parser.add_argument("--url", help="Remote upload URL")
    parser.add_argument("--username", help="Basic auth username")
    parser.add_argument("--password", help="Basic auth password")
    parser.add_argument("--token", help="Bearer token")
    parser.add_argument("--timeout", type=int, default=60, help="HTTP timeout in seconds")
    return parser


def main() -> int:
    parser = build_arg_parser()
    args = parser.parse_args()
    config = load_upload_config(args)
    result = upload_file(
        file_path=args.file,
        mode=config["mode"],
        target=config["target"],
        url=config["url"],
        username=config["username"],
        password=config["password"],
        token=config["token"],
        timeout=config["timeout"],
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
