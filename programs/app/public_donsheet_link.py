from __future__ import annotations

from io import BytesIO
from pathlib import Path
import re
from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse
from urllib.request import Request, urlopen
import zipfile


MAX_DOWNLOAD_BYTES = 50 * 1024 * 1024
ALLOWED_HOST_SUFFIXES = (".sharepoint.com", ".onedrive.live.com", ".1drv.com")
ALLOWED_HOSTS = {"1drv.ms"}


class PublicDonSheetLinkError(ValueError):
    """Raised when an anonymous Microsoft 365 link cannot supply an XLSX file."""


def _is_allowed_microsoft_host(hostname: str) -> bool:
    hostname = hostname.lower().rstrip(".")
    return hostname in ALLOWED_HOSTS or hostname.endswith(ALLOWED_HOST_SUFFIXES)


def _download_candidates(url: str) -> list[str]:
    parsed = urlparse(url)
    query = parse_qsl(parsed.query, keep_blank_values=True)
    without_download = [(key, value) for key, value in query if key.lower() != "download"]
    download_url = urlunparse(parsed._replace(query=urlencode([*without_download, ("download", "1")])))
    return [url] if download_url == url else [url, download_url]


def _filename(response, fallback_url: str) -> str:
    disposition = response.headers.get("Content-Disposition", "")
    match = re.search(r"filename\*?=(?:UTF-8''|\")?([^\";]+)", disposition, flags=re.IGNORECASE)
    if match:
        candidate = Path(match.group(1).strip()).name
        if candidate.lower().endswith(".xlsx"):
            return candidate
    query_name = dict(parse_qsl(urlparse(fallback_url).query)).get("file", "")
    candidate = Path(query_name).name
    return candidate if candidate.lower().endswith(".xlsx") else "DonSheet_from_public_link.xlsx"


def _is_xlsx(payload: bytes) -> bool:
    if not zipfile.is_zipfile(BytesIO(payload)):
        return False
    with zipfile.ZipFile(BytesIO(payload)) as archive:
        names = set(archive.namelist())
    return "[Content_Types].xml" in names and "xl/workbook.xml" in names


def download_public_donsheet(url: str) -> tuple[str, bytes]:
    """Download an anonymously accessible Microsoft 365 XLSX link without credentials."""
    url = url.strip()
    parsed = urlparse(url)
    if parsed.scheme != "https" or not parsed.hostname or parsed.username or parsed.password or parsed.port:
        raise PublicDonSheetLinkError("Enter a valid HTTPS SharePoint or OneDrive public link.")
    if not _is_allowed_microsoft_host(parsed.hostname):
        raise PublicDonSheetLinkError("Only public SharePoint or OneDrive links are accepted.")

    last_content_type = ""
    for candidate in _download_candidates(url):
        request = Request(candidate, headers={"User-Agent": "SmartSeminary/1.0"})
        try:
            with urlopen(request, timeout=20) as response:  # nosec B310: host is allow-listed above
                final_url = getattr(response, "geturl", lambda: candidate)()
                final_host = urlparse(final_url).hostname or ""
                if not _is_allowed_microsoft_host(final_host):
                    raise PublicDonSheetLinkError("The public link redirected outside Microsoft 365 and was blocked.")
                length = int(response.headers.get("Content-Length", "0") or 0)
                if length > MAX_DOWNLOAD_BYTES:
                    raise PublicDonSheetLinkError("The linked workbook is larger than the 50 MB import limit.")
                payload = response.read(MAX_DOWNLOAD_BYTES + 1)
                if len(payload) > MAX_DOWNLOAD_BYTES:
                    raise PublicDonSheetLinkError("The linked workbook is larger than the 50 MB import limit.")
                last_content_type = response.headers.get("Content-Type", "")
                if _is_xlsx(payload):
                    return _filename(response, candidate), payload
        except PublicDonSheetLinkError:
            raise
        except Exception as error:
            raise PublicDonSheetLinkError(f"Could not retrieve the public Microsoft 365 link: {error}") from error

    content_hint = f" The link returned {last_content_type}." if last_content_type else ""
    raise PublicDonSheetLinkError(
        "The link did not return an Excel workbook. Confirm that it is an 'Anyone with the link' "
        "link with download allowed, rather than a link that requires Microsoft 365 sign-in."
        + content_hint
    )
