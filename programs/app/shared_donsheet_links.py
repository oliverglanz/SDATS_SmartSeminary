from __future__ import annotations

import json
from pathlib import Path
from urllib.parse import urlparse


MAX_SHARED_DONSHEET_LINKS = 15
_ALLOWED_HOST_SUFFIXES = (".sharepoint.com", ".onedrive.live.com", ".1drv.com")
_ALLOWED_HOSTS = {"1drv.ms"}


class SharedDonSheetLinkError(ValueError):
    """Raised when saved DonSheet link data is invalid."""


def _validate_link(label: str, url: str) -> dict[str, str]:
    label = label.strip()
    url = url.strip()
    parsed = urlparse(url)
    hostname = (parsed.hostname or "").lower().rstrip(".")
    if not label:
        raise SharedDonSheetLinkError("Enter a name for the shared DonSheet link.")
    if parsed.scheme != "https" or not hostname:
        raise SharedDonSheetLinkError("Enter a valid HTTPS SharePoint or OneDrive link.")
    if hostname not in _ALLOWED_HOSTS and not hostname.endswith(_ALLOWED_HOST_SUFFIXES):
        raise SharedDonSheetLinkError("Only SharePoint or OneDrive links can be saved.")
    return {"label": label, "url": url}


def load_shared_donsheet_links(storage_path: Path) -> list[dict[str, str]]:
    if not storage_path.is_file():
        return []
    try:
        data = json.loads(storage_path.read_text(encoding="utf-8"))
        links = data.get("links", [])
        if not isinstance(links, list):
            raise ValueError
        return [_validate_link(str(item["label"]), str(item["url"])) for item in links]
    except (json.JSONDecodeError, KeyError, TypeError, ValueError) as error:
        raise SharedDonSheetLinkError("The saved shared-DonSheet link list is invalid.") from error


def save_shared_donsheet_links(storage_path: Path, links: list[dict[str, str]]) -> None:
    if len(links) > MAX_SHARED_DONSHEET_LINKS:
        raise SharedDonSheetLinkError(f"Only {MAX_SHARED_DONSHEET_LINKS} shared DonSheet links may be saved.")
    cleaned = [_validate_link(str(item.get("label", "")), str(item.get("url", ""))) for item in links]
    labels = [item["label"].casefold() for item in cleaned]
    if len(labels) != len(set(labels)):
        raise SharedDonSheetLinkError("Each shared DonSheet link needs a unique name.")

    storage_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = storage_path.with_suffix(storage_path.suffix + ".tmp")
    temporary_path.write_text(json.dumps({"links": cleaned}, indent=2) + "\n", encoding="utf-8")
    temporary_path.replace(storage_path)
