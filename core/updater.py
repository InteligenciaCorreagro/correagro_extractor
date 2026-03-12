import json
import os
import re
import subprocess
import tempfile
import urllib.error
import urllib.request
from dataclasses import dataclass

from PySide6.QtCore import QObject, Signal


DEFAULT_REPO = "InteligenciaCorreagro/correagro_extractor"
DEFAULT_ASSET_NAME = "Correagro-Extractor-Setup.exe"


def _repo_from_any(value: str) -> str:
    if not value:
        return ""
    s = value.strip()
    m = re.search(r"github\.com[:/]+([^/]+)/([^/]+?)(?:\.git)?$", s, re.IGNORECASE)
    if m:
        return f"{m.group(1)}/{m.group(2)}"
    if "/" in s and " " not in s:
        return s
    return ""


def get_update_repo() -> str:
    env = os.environ.get("CORREAGRO_UPDATE_REPO", "").strip()
    repo = _repo_from_any(env)
    if repo:
        return repo
    try:
        here = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        p = os.path.join(here, "update_repo.txt")
        if os.path.exists(p):
            with open(p, "r", encoding="utf-8") as f:
                repo = _repo_from_any(f.read())
                if repo:
                    return repo
    except Exception:
        pass
    return DEFAULT_REPO


@dataclass
class ReleaseInfo:
    tag: str
    url: str
    asset_name: str


def _http_json(url: str) -> dict:
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": "correagro-extractor-updater",
            "Accept": "application/vnd.github+json",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            data = resp.read().decode("utf-8", errors="replace")
        return json.loads(data)
    except urllib.error.HTTPError as e:
        if e.code == 404:
            raise RuntimeError("Repo privado o release inexistente (HTTP 404).") from e
        if e.code == 403:
            raise RuntimeError("Acceso denegado o rate-limit de GitHub (HTTP 403).") from e
        raise RuntimeError(f"Error HTTP al consultar GitHub (HTTP {e.code}).") from e
    except urllib.error.URLError as e:
        raise RuntimeError("No hay conexión o GitHub está bloqueado.") from e


def get_latest_release(repo: str, preferred_asset_name: str = DEFAULT_ASSET_NAME) -> ReleaseInfo | None:
    if not repo or "/" not in repo:
        return None
    api = f"https://api.github.com/repos/{repo}/releases/latest"
    payload = _http_json(api)
    tag = str(payload.get("tag_name") or "").strip()
    html_url = str(payload.get("html_url") or "").strip()
    assets = payload.get("assets") or []
    best = None
    for a in assets:
        name = str(a.get("name") or "")
        url = str(a.get("browser_download_url") or "")
        if not name or not url:
            continue
        if name == preferred_asset_name:
            return ReleaseInfo(tag=tag, url=url, asset_name=name)
        if best is None and name.lower().endswith(".exe"):
            best = ReleaseInfo(tag=tag, url=url, asset_name=name)
    if best is not None:
        return best
    if html_url:
        return ReleaseInfo(tag=tag, url=html_url, asset_name="")
    return None


def _normalize_tag(s: str) -> str:
    if not s:
        return ""
    s = s.strip()
    return s[1:] if s.lower().startswith("v") else s


def is_update_available(current_version: str, latest_tag: str) -> bool:
    cur = _normalize_tag(current_version)
    lat = _normalize_tag(latest_tag)
    return bool(cur and lat and cur != lat)


def download_file(url: str, filename: str) -> str:
    tmp_dir = tempfile.mkdtemp(prefix="correagro_update_")
    out = os.path.join(tmp_dir, filename)
    req = urllib.request.Request(url, headers={"User-Agent": "correagro-extractor-updater"})
    with urllib.request.urlopen(req, timeout=60) as resp, open(out, "wb") as f:
        f.write(resp.read())
    return out


def run_installer(installer_path: str) -> None:
    subprocess.Popen([installer_path, "/VERYSILENT", "/NORESTART", "/SUPPRESSMSGBOXES"], close_fds=True)


class UpdateWorker(QObject):
    finished = Signal(object, object)

    def __init__(self, current_version: str, parent=None):
        super().__init__(parent)
        self.current_version = current_version

    def run(self):
        try:
            repo = get_update_repo()
            rel = get_latest_release(repo)
            if rel is None or not rel.tag:
                self.finished.emit(None, None)
                return
            if not is_update_available(self.current_version, rel.tag):
                self.finished.emit(None, None)
                return
            self.finished.emit(rel, None)
        except Exception as e:
            self.finished.emit(None, str(e))
