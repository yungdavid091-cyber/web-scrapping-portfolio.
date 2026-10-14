"""
ScrapeX Pro - Configurable Web Scraper
For educational use on websites you are allowed to scrape.

Install in Pydroid 3:
    pip install requests beautifulsoup4 lxml pandas openpyxl

Features:
- CSS selectors for record containers and fields
- Single-page, page-number, and next-link pagination
- Retries, timeouts, and request delays
- Duplicate detection
- Same-host checks and page/item limits
- CSV, JSON, and Excel output
- Optional pandas data cleaning
- Repairs common currency encoding issues
- Optional normalized price columns
- UTF-8-compatible exports
- Automatic unique filenames
- Detailed cleaning and scraping report

Example test site:
    https://books.toscrape.com/
"""

import csv
import json
import os
import re
import time

from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from urllib.parse import urljoin, urlparse, urldefrag

import requests
import pandas as pd

from bs4 import BeautifulSoup
from requests.exceptions import RequestException


# ============================================================
# SETTINGS
# ============================================================

REQUEST_TIMEOUT = 20
MAX_RETRIES = 3
RETRY_BACKOFF_SECONDS = 2
REQUEST_DELAY_SECONDS = 1.0

MAX_ITEMS_HARD_LIMIT = 1000
MAX_PAGES_HARD_LIMIT = 100

USER_AGENT = "ScrapeX-Pro-EducationalScraper/2.0"

TEMPORARY_HTTP_CODES = {429, 500, 502, 503, 504}


# ============================================================
# TEXT CLEANING
# ============================================================

def repair_text_encoding(value):
    """Repair common currency and punctuation encoding problems."""

    if value is None:
        return ""

    value = str(value)

    replacements = {
        "Â£": "£",
        "Â€": "€",
        "Â¥": "¥",
        "Â¢": "¢",
        "â‚¬": "€",
        "â€™": "’",
        "â€œ": "“",
        "â€": "”",
        "â€“": "–",
        "â€”": "—",
    }

    for wrong, correct in replacements.items():
        value = value.replace(wrong, correct)

    return value


def clean_text(value):
    """Repair encoding and normalize whitespace."""

    return " ".join(
        repair_text_encoding(value).split()
    )


def normalize_price(value):
    """
    Convert common price strings to two decimal places.

    Examples:
        £51.77    -> 51.77
        £50.10    -> 50.10
        $1,200.50 -> 1200.50
        £850      -> 850.00
        19,99     -> 19.99

    This function does not convert currencies.
    """

    value = clean_text(value)

    if not value:
        return ""

    candidate = re.sub(r"[^\d,.\-]", "", value)

    if not candidate or not re.search(r"\d", candidate):
        return ""

    if "," in candidate and "." in candidate:

        # The last separator is treated as the decimal separator.
        if candidate.rfind(".") > candidate.rfind(","):
            candidate = candidate.replace(",", "")
        else:
            candidate = candidate.replace(".", "")
            candidate = candidate.replace(",", ".")

    elif "," in candidate:

        parts = candidate.split(",")

        if len(parts) == 2 and len(parts[1]) in (1, 2):
            candidate = parts[0] + "." + parts[1]
        else:
            candidate = candidate.replace(",", "")

    try:
        number = Decimal(candidate)

        if not number.is_finite():
            return ""

        number = number.quantize(
            Decimal("0.01"),
            rounding=ROUND_HALF_UP
        )

        return f"{number:.2f}"

    except (InvalidOperation, ValueError):
        return ""


# ============================================================
# GENERAL HELPERS
# ============================================================

def log(message, level="INFO"):
    print(f"[{level}] {message}")


def ask_yes_no(question, default="y"):
    while True:

        suffix = "Y/n" if default.lower() == "y" else "y/N"

        answer = input(
            f"{question} ({suffix}): "
        ).strip().lower()

        if not answer:
            answer = default.lower()

        if answer in ("y", "yes"):
            return True

        if answer in ("n", "no"):
            return False

        print("Please enter y or n.")


def ask_number(prompt, minimum, maximum, default=None):

    while True:

        suffix = f" [{default}]" if default is not None else ""

        raw = input(
            f"{prompt}{suffix}: "
        ).strip()

        if not raw and default is not None:
            return default

        try:
            value = int(raw)

            if minimum <= value <= maximum:
                return value

        except ValueError:
            pass

        print(
            f"Please enter a whole number "
            f"from {minimum} to {maximum}."
        )


def validate_http_url(url):
    """Validate a complete HTTP or HTTPS URL."""

    url = url.strip()

    try:
        parsed = urlparse(url)

        if parsed.scheme.lower() not in ("http", "https"):
            raise ValueError(
                "Enter a URL beginning with http:// or https://."
            )

        if not parsed.hostname:
            raise ValueError("The URL must contain a hostname.")

        if parsed.username or parsed.password:
            raise ValueError(
                "URLs containing usernames or passwords are not accepted."
            )

        # Accessing port can raise ValueError for malformed ports.
        _ = parsed.port

        normalized, _fragment = urldefrag(url)

        return normalized

    except ValueError as exc:
        raise ValueError(f"Invalid URL: {exc}") from exc


def same_host(url_a, url_b):
    """Compare hostnames case-insensitively."""

    host_a = (urlparse(url_a).hostname or "").lower()
    host_b = (urlparse(url_b).hostname or "").lower()

    return bool(host_a and host_b and host_a == host_b)


def safe_filename(name, default="scraped_data.csv"):
    """Keep output filenames in the current working directory."""

    name = name.strip() or default
    name = os.path.basename(name)

    if not name or name in (".", ".."):
        return default

    return name


def unique_filename(filename):
    """Return a filename that does not already exist."""

    filename = safe_filename(filename)

    stem, extension = os.path.splitext(filename)

    candidate = filename
    counter = 1

    while os.path.exists(candidate):

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        if counter == 1:
            candidate = f"{stem}_{stamp}{extension}"
        else:
            candidate = f"{stem}_{stamp}_{counter}{extension}"

        counter += 1

    return candidate


# ============================================================
# HTML PARSER
# ============================================================

def parse_html(html):
    """Use lxml when available; otherwise use html.parser."""

    try:
        return BeautifulSoup(html, "lxml")

    except Exception as exc:
        log(
            f"lxml parser unavailable; using built-in parser: {exc}",
            "WARN"
        )

        return BeautifulSoup(html, "html.parser")


# ============================================================
# HTTP REQUESTS
# ============================================================

def fetch_html(session, url):
    """Fetch a page with retries and a timeout."""

    for attempt in range(1, MAX_RETRIES + 1):

        try:

            response = session.get(
                url,
                timeout=REQUEST_TIMEOUT
            )

            if response.status_code in TEMPORARY_HTTP_CODES:

                if attempt < MAX_RETRIES:

                    wait = RETRY_BACKOFF_SECONDS * attempt

                    log(
                        f"HTTP {response.status_code}; "
                        f"retrying in {wait}s "
                        f"({attempt}/{MAX_RETRIES}).",
                        "WARN"
                    )

                    time.sleep(wait)
                    continue

            response.raise_for_status()

            # requests normally detects the response encoding.
            # response.text provides decoded HTML for BeautifulSoup.
            return response.text

        except RequestException as exc:

            if attempt >= MAX_RETRIES:

                log(
                    f"Could not fetch {url}: {exc}",
                    "ERROR"
                )

                return None

            wait = RETRY_BACKOFF_SECONDS * attempt

            log(
                f"Request failed; retrying in {wait}s "
                f"({attempt}/{MAX_RETRIES}): {exc}",
                "WARN"
            )

            time.sleep(wait)

    return None


# ============================================================
# DATA EXTRACTION
# ============================================================

def extract_value(container, selector, attribute=""):
    """Extract text or an HTML attribute from a record."""

    if selector == ".":
        found = container

    else:

        try:
            found = container.select_one(selector)

        except Exception as exc:
            raise ValueError(
                f"Invalid CSS selector {selector!r}: {exc}"
            ) from exc

    if found is None:
        return ""

    if attribute:

        value = found.get(attribute, "")

        if isinstance(value, list):
            return " ".join(str(part) for part in value)

        return str(value)

    return clean_text(
        found.get_text(" ", strip=True)
    )


def absolute_link(base_url, value):

    if not value:
        return ""

    return urljoin(base_url, value.strip())


def normalize_record(row, trim_whitespace=True):
    """Repair text encoding in each extracted field."""

    cleaned = {}

    for field, value in row.items():

        value = repair_text_encoding(value)

        if trim_whitespace:
            value = clean_text(value)

        cleaned[field] = value

    return cleaned


# ============================================================
# OPTIONAL PANDAS CLEANING
# ============================================================

def clean_dataframe(rows, fields):
    """
    Use pandas for consistent column ordering and optional
    duplicate handling. Original field values are preserved.
    """

    dataframe = pd.DataFrame(rows, columns=fields)

    # Convert missing pandas values to empty strings.
    dataframe = dataframe.fillna("")

    for column in dataframe.columns:
        dataframe[column] = dataframe[column].map(
            lambda value: repair_text_encoding(value)
            if isinstance(value, str)
            else value
        )

    return dataframe


def add_normalized_price_columns(dataframe):
    """Optionally add normalized columns for price/cost fields."""

    price_candidates = [
        column
        for column in dataframe.columns
        if "price" in str(column).casefold()
        or "cost" in str(column).casefold()
    ]

    if not price_candidates:
        log("No price or cost columns detected.")
        return dataframe

    print("\nPossible price columns:")

    for number, column in enumerate(price_candidates, start=1):
        print(f"{number}. {column}")

    if not ask_yes_no(
        "Add normalized price columns while keeping original prices?"
    ):
        return dataframe

    raw = input(
        "Choose column numbers separated by commas "
        "(Enter = all detected price columns): "
    ).strip()

    if not raw:
        selected = price_candidates

    else:

        try:
            indexes = [
                int(item.strip()) - 1
                for item in raw.split(",")
            ]

            if any(
                index < 0 or index >= len(price_candidates)
                for index in indexes
            ):
                raise ValueError

            selected = list(dict.fromkeys(
                price_candidates[index]
                for index in indexes
            ))

        except ValueError:
            log(
                "Invalid selection. Using all detected price columns.",
                "WARN"
            )
            selected = price_candidates

    for column in selected:

        base_name = f"{column}_clean"
        new_name = base_name
        suffix = 2

        while new_name in dataframe.columns:
            new_name = f"{base_name}_{suffix}"
            suffix += 1

        dataframe[new_name] = dataframe[column].map(
            normalize_price
        )

    return dataframe


# ============================================================
# PAGINATION AND SCRAPING ENGINE
# ============================================================

def scrape_site(
    start_url,
    record_selector,
    field_rules,
    pagination,
    max_items,
    max_pages,
    trim_whitespace=True
):

    start_url = validate_http_url(start_url)

    session = requests.Session()

    session.headers.update({
        "User-Agent": USER_AGENT,
        "Accept": "text/html,application/xhtml+xml",
    })

    fields = list(field_rules.keys())

    all_rows = []
    seen_items = set()
    visited_urls = set()

    current_url = start_url
    page_number = 1
    duplicates_removed = 0

    try:

        while (
            len(all_rows) < max_items
            and page_number <= max_pages
        ):

            if pagination["type"] == "page_number":

                current_url = pagination["template"].format(
                    page=page_number
                )

                current_url = validate_http_url(current_url)

            if not same_host(start_url, current_url):

                log(
                    "URL points to another hostname. Stopping.",
                    "WARN"
                )

                break

            current_url, _fragment = urldefrag(current_url)

            if current_url in visited_urls:

                log(
                    "A page URL repeated; stopping to prevent a loop.",
                    "WARN"
                )

                break

            visited_urls.add(current_url)

            log(
                f"Fetching page {page_number}: {current_url}"
            )

            html = fetch_html(session, current_url)

            if html is None:
                break

            soup = parse_html(html)

            try:
                records = soup.select(record_selector)

            except Exception as exc:
                raise ValueError(
                    f"Invalid record CSS selector: {exc}"
                ) from exc

            if not records:

                log(
                    "No matching record containers found; stopping.",
                    "WARN"
                )

                break

            new_rows = []

            for container in records:

                if len(all_rows) + len(new_rows) >= max_items:
                    break

                row = {}

                for field, (selector, attribute) in field_rules.items():

                    value = extract_value(
                        container,
                        selector,
                        attribute
                    )

                    if (
                        attribute.lower() in ("href", "src")
                        and value
                    ):
                        value = absolute_link(
                            current_url,
                            value
                        )

                    row[field] = value

                row = normalize_record(
                    row,
                    trim_whitespace=trim_whitespace
                )

                key = tuple(
                    row.get(field, "")
                    for field in fields
                )

                if key in seen_items:
                    duplicates_removed += 1
                    continue

                seen_items.add(key)
                new_rows.append(row)

            all_rows.extend(new_rows)

            log(
                f"Containers: {len(records)} | "
                f"New unique records: {len(new_rows)} | "
                f"Total: {len(all_rows)}"
            )

            if len(all_rows) >= max_items:

                log("Maximum item limit reached.")
                break

            if pagination["type"] == "single":
                break

            if pagination["type"] == "page_number":

                page_number += 1

            elif pagination["type"] == "next_link":

                try:
                    next_tag = soup.select_one(
                        pagination["selector"]
                    )

                except Exception as exc:
                    raise ValueError(
                        f"Invalid next-page CSS selector: {exc}"
                    ) from exc

                if next_tag is None:

                    log(
                        "Next-page link not found; finished."
                    )

                    break

                next_value = next_tag.get(
                    pagination["attribute"],
                    ""
                )

                if not next_value:

                    log(
                        "Next-page link has no URL attribute.",
                        "WARN"
                    )

                    break

                next_url = validate_http_url(
                    absolute_link(
                        current_url,
                        next_value
                    )
                )

                if not same_host(start_url, next_url):

                    log(
                        "Next-page link points to another hostname.",
                        "WARN"
                    )

                    break

                current_url = next_url
                page_number += 1

            time.sleep(REQUEST_DELAY_SECONDS)

        if (
            page_number > max_pages
            and len(all_rows) < max_items
        ):

            log(
                f"Maximum page limit ({max_pages}) reached.",
                "WARN"
            )

    finally:
        session.close()

    return all_rows, fields, duplicates_removed


# ============================================================
# SAVE OUTPUT
# ============================================================

def save_data(dataframe, filename):
    """Save CSV, JSON, or Excel without overwriting existing files."""

    filename = safe_filename(filename)

    extension = os.path.splitext(filename)[1].lower()

    if extension not in (".csv", ".json", ".xlsx"):

        raise ValueError(
            "Choose a .csv, .json, or .xlsx output filename."
        )

    script_path = os.path.realpath(__file__)
    output_path = os.path.realpath(filename)

    if output_path == script_path:

        raise ValueError(
            "Output filename cannot overwrite the scraper script."
        )

    filename = unique_filename(filename)

    extension = os.path.splitext(filename)[1].lower()

    if extension == ".csv":

        # UTF-8 with BOM helps spreadsheet applications recognize
        # currency symbols and other non-ASCII text.
        dataframe.to_csv(
            filename,
            index=False,
            encoding="utf-8-sig"
        )

    elif extension == ".json":

        dataframe.to_json(
            filename,
            orient="records",
            force_ascii=False,
            indent=2
        )

    elif extension == ".xlsx":

        dataframe.to_excel(
            filename,
            index=False,
            engine="openpyxl"
        )

    return os.path.abspath(filename)


# ============================================================
# REPORT
# ============================================================

def save_report(
    output_path,
    start_url,
    records_saved,
    duplicates_removed,
    pages_limit,
    items_limit
):

    report_name = os.path.splitext(output_path)[0] + "_report.txt"
"""
ScrapeX Pro - Configurable Web Scraper
For educational use on websites you are allowed to scrape.

Install in Pydroid 3:
    pip install requests beautifulsoup4 lxml pandas openpyxl

Features:
- CSS selectors for record containers and fields
- Single-page, page-number, and next-link pagination
- Retries, timeouts, and request delays
- Duplicate detection
- Same-host checks and page/item limits
- CSV, JSON, and Excel output
- Optional pandas data cleaning
- Repairs common currency encoding issues
- Optional normalized price columns
- UTF-8-compatible exports
- Automatic unique filenames
- Detailed cleaning and scraping report

Example test site:
    https://books.toscrape.com/
"""

import csv
import json
import os
import re
import time

from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from urllib.parse import urljoin, urlparse, urldefrag

import requests
import pandas as pd

from bs4 import BeautifulSoup
from requests.exceptions import RequestException


# ============================================================
# SETTINGS
# ============================================================

REQUEST_TIMEOUT = 20
MAX_RETRIES = 3
RETRY_BACKOFF_SECONDS = 2
REQUEST_DELAY_SECONDS = 1.0

MAX_ITEMS_HARD_LIMIT = 1000
MAX_PAGES_HARD_LIMIT = 100

USER_AGENT = "ScrapeX-Pro-EducationalScraper/2.0"

TEMPORARY_HTTP_CODES = {429, 500, 502, 503, 504}


# ============================================================
# TEXT CLEANING
# ============================================================

def repair_text_encoding(value):
    """Repair common currency and punctuation encoding problems."""

    if value is None:
        return ""

    value = str(value)

    replacements = {
        "Â£": "£",
        "Â€": "€",
        "Â¥": "¥",
        "Â¢": "¢",
        "â‚¬": "€",
        "â€™": "’",
        "â€œ": "“",
        "â€": "”",
        "â€“": "–",
        "â€”": "—",
    }

    for wrong, correct in replacements.items():
        value = value.replace(wrong, correct)

    return value


def clean_text(value):
    """Repair encoding and normalize whitespace."""

    return " ".join(
        repair_text_encoding(value).split()
    )


def normalize_price(value):
    """
    Convert common price strings to two decimal places.

    Examples:
        £51.77    -> 51.77
        £50.10    -> 50.10
        $1,200.50 -> 1200.50
        £850      -> 850.00
        19,99     -> 19.99

    This function does not convert currencies.
    """

    value = clean_text(value)

    if not value:
        return ""

    candidate = re.sub(r"[^\d,.\-]", "", value)

    if not candidate or not re.search(r"\d", candidate):
        return ""

    if "," in candidate and "." in candidate:

        # The last separator is treated as the decimal separator.
        if candidate.rfind(".") > candidate.rfind(","):
            candidate = candidate.replace(",", "")
        else:
            candidate = candidate.replace(".", "")
            candidate = candidate.replace(",", ".")

    elif "," in candidate:

        parts = candidate.split(",")

        if len(parts) == 2 and len(parts[1]) in (1, 2):
            candidate = parts[0] + "." + parts[1]
        else:
            candidate = candidate.replace(",", "")

    try:
        number = Decimal(candidate)

        if not number.is_finite():
            return ""

        number = number.quantize(
            Decimal("0.01"),
            rounding=ROUND_HALF_UP
        )

        return f"{number:.2f}"

    except (InvalidOperation, ValueError):
        return ""


# ============================================================
# GENERAL HELPERS
# ============================================================

def log(message, level="INFO"):
    print(f"[{level}] {message}")


def ask_yes_no(question, default="y"):
    while True:

        suffix = "Y/n" if default.lower() == "y" else "y/N"

        answer = input(
            f"{question} ({suffix}): "
        ).strip().lower()

        if not answer:
            answer = default.lower()

        if answer in ("y", "yes"):
            return True

        if answer in ("n", "no"):
            return False

        print("Please enter y or n.")


def ask_number(prompt, minimum, maximum, default=None):

    while True:

        suffix = f" [{default}]" if default is not None else ""

        raw = input(
            f"{prompt}{suffix}: "
        ).strip()

        if not raw and default is not None:
            return default

        try:
            value = int(raw)

            if minimum <= value <= maximum:
                return value

        except ValueError:
            pass

        print(
            f"Please enter a whole number "
            f"from {minimum} to {maximum}."
        )


def validate_http_url(url):
    """Validate a complete HTTP or HTTPS URL."""

    url = url.strip()

    try:
        parsed = urlparse(url)

        if parsed.scheme.lower() not in ("http", "https"):
            raise ValueError(
                "Enter a URL beginning with http:// or https://."
            )

        if not parsed.hostname:
            raise ValueError("The URL must contain a hostname.")

        if parsed.username or parsed.password:
            raise ValueError(
                "URLs containing usernames or passwords are not accepted."
            )

        # Accessing port can raise ValueError for malformed ports.
        _ = parsed.port

        normalized, _fragment = urldefrag(url)

        return normalized

    except ValueError as exc:
        raise ValueError(f"Invalid URL: {exc}") from exc


def same_host(url_a, url_b):
    """Compare hostnames case-insensitively."""

    host_a = (urlparse(url_a).hostname or "").lower()
    host_b = (urlparse(url_b).hostname or "").lower()

    return bool(host_a and host_b and host_a == host_b)


def safe_filename(name, default="scraped_data.csv"):
    """Keep output filenames in the current working directory."""

    name = name.strip() or default
    name = os.path.basename(name)

    if not name or name in (".", ".."):
        return default

    return name


def unique_filename(filename):
    """Return a filename that does not already exist."""

    filename = safe_filename(filename)

    stem, extension = os.path.splitext(filename)

    candidate = filename
    counter = 1

    while os.path.exists(candidate):

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        if counter == 1:
            candidate = f"{stem}_{stamp}{extension}"
        else:
            candidate = f"{stem}_{stamp}_{counter}{extension}"

        counter += 1

    return candidate


# ============================================================
# HTML PARSER
# ============================================================

def parse_html(html):
    """Use lxml when available; otherwise use html.parser."""

    try:
        return BeautifulSoup(html, "lxml")

    except Exception as exc:
        log(
            f"lxml parser unavailable; using built-in parser: {exc}",
            "WARN"
        )

        return BeautifulSoup(html, "html.parser")


# ============================================================
# HTTP REQUESTS
# ============================================================

def fetch_html(session, url):
    """Fetch a page with retries and a timeout."""

    for attempt in range(1, MAX_RETRIES + 1):

        try:

            response = session.get(
                url,
                timeout=REQUEST_TIMEOUT
            )

            if response.status_code in TEMPORARY_HTTP_CODES:

                if attempt < MAX_RETRIES:

                    wait = RETRY_BACKOFF_SECONDS * attempt

                    log(
                        f"HTTP {response.status_code}; "
                        f"retrying in {wait}s "
                        f"({attempt}/{MAX_RETRIES}).",
                        "WARN"
                    )

                    time.sleep(wait)
                    continue

            response.raise_for_status()

            # requests normally detects the response encoding.
            # response.text provides decoded HTML for BeautifulSoup.
            return response.text

        except RequestException as exc:

            if attempt >= MAX_RETRIES:

                log(
                    f"Could not fetch {url}: {exc}",
                    "ERROR"
                )

                return None

            wait = RETRY_BACKOFF_SECONDS * attempt

            log(
                f"Request failed; retrying in {wait}s "
                f"({attempt}/{MAX_RETRIES}): {exc}",
                "WARN"
            )

            time.sleep(wait)

    return None


# ============================================================
# DATA EXTRACTION
# ============================================================

def extract_value(container, selector, attribute=""):
    """Extract text or an HTML attribute from a record."""

    if selector == ".":
        found = container

    else:

        try:
            found = container.select_one(selector)

        except Exception as exc:
            raise ValueError(
                f"Invalid CSS selector {selector!r}: {exc}"
            ) from exc

    if found is None:
        return ""

    if attribute:

        value = found.get(attribute, "")

        if isinstance(value, list):
            return " ".join(str(part) for part in value)

        return str(value)

    return clean_text(
        found.get_text(" ", strip=True)
    )


def absolute_link(base_url, value):

    if not value:
        return ""

    return urljoin(base_url, value.strip())


def normalize_record(row, trim_whitespace=True):
    """Repair text encoding in each extracted field."""

    cleaned = {}

    for field, value in row.items():

        value = repair_text_encoding(value)

        if trim_whitespace:
            value = clean_text(value)

        cleaned[field] = value

    return cleaned


# ============================================================
# OPTIONAL PANDAS CLEANING
# ============================================================

def clean_dataframe(rows, fields):
    """
    Use pandas for consistent column ordering and optional
    duplicate handling. Original field values are preserved.
    """

    dataframe = pd.DataFrame(rows, columns=fields)

    # Convert missing pandas values to empty strings.
    dataframe = dataframe.fillna("")

    for column in dataframe.columns:
        dataframe[column] = dataframe[column].map(
            lambda value: repair_text_encoding(value)
            if isinstance(value, str)
            else value
        )

    return dataframe


def add_normalized_price_columns(dataframe):
    """Optionally add normalized columns for price/cost fields."""

    price_candidates = [
        column
        for column in dataframe.columns
        if "price" in str(column).casefold()
        or "cost" in str(column).casefold()
    ]

    if not price_candidates:
        log("No price or cost columns detected.")
        return dataframe

    print("\nPossible price columns:")

    for number, column in enumerate(price_candidates, start=1):
        print(f"{number}. {column}")

    if not ask_yes_no(
        "Add normalized price columns while keeping original prices?"
    ):
        return dataframe

    raw = input(
        "Choose column numbers separated by commas "
        "(Enter = all detected price columns): "
    ).strip()

    if not raw:
        selected = price_candidates

    else:

        try:
            indexes = [
                int(item.strip()) - 1
                for item in raw.split(",")
            ]

            if any(
                index < 0 or index >= len(price_candidates)
                for index in indexes
            ):
                raise ValueError

            selected = list(dict.fromkeys(
                price_candidates[index]
                for index in indexes
            ))

        except ValueError:
            log(
                "Invalid selection. Using all detected price columns.",
                "WARN"
            )
            selected = price_candidates

    for column in selected:

        base_name = f"{column}_clean"
        new_name = base_name
        suffix = 2

        while new_name in dataframe.columns:
            new_name = f"{base_name}_{suffix}"
            suffix += 1

        dataframe[new_name] = dataframe[column].map(
            normalize_price
        )

    return dataframe


# ============================================================
# PAGINATION AND SCRAPING ENGINE
# ============================================================

def scrape_site(
    start_url,
    record_selector,
    field_rules,
    pagination,
    max_items,
    max_pages,
    trim_whitespace=True
):

    start_url = validate_http_url(start_url)

    session = requests.Session()

    session.headers.update({
        "User-Agent": USER_AGENT,
        "Accept": "text/html,application/xhtml+xml",
    })

    fields = list(field_rules.keys())

    all_rows = []
    seen_items = set()
    visited_urls = set()

    current_url = start_url
    page_number = 1
    duplicates_removed = 0

    try:

        while (
            len(all_rows) < max_items
            and page_number <= max_pages
        ):

            if pagination["type"] == "page_number":

                current_url = pagination["template"].format(
                    page=page_number
                )

                current_url = validate_http_url(current_url)

            if not same_host(start_url, current_url):

                log(
                    "URL points to another hostname. Stopping.",
                    "WARN"
                )

                break

            current_url, _fragment = urldefrag(current_url)

            if current_url in visited_urls:

                log(
                    "A page URL repeated; stopping to prevent a loop.",
                    "WARN"
                )

                break

            visited_urls.add(current_url)

            log(
                f"Fetching page {page_number}: {current_url}"
            )

            html = fetch_html(session, current_url)

            if html is None:
                break

            soup = parse_html(html)

            try:
                records = soup.select(record_selector)

            except Exception as exc:
                raise ValueError(
                    f"Invalid record CSS selector: {exc}"
                ) from exc

            if not records:

                log(
                    "No matching record containers found; stopping.",
                    "WARN"
                )

                break

            new_rows = []

            for container in records:

                if len(all_rows) + len(new_rows) >= max_items:
                    break

                row = {}

                for field, (selector, attribute) in field_rules.items():

                    value = extract_value(
                        container,
                        selector,
                        attribute
                    )

                    if (
                        attribute.lower() in ("href", "src")
                        and value
                    ):
                        value = absolute_link(
                            current_url,
                            value
                        )

                    row[field] = value

                row = normalize_record(
                    row,
                    trim_whitespace=trim_whitespace
                )

                key = tuple(
                    row.get(field, "")
                    for field in fields
                )

                if key in seen_items:
                    duplicates_removed += 1
                    continue

                seen_items.add(key)
                new_rows.append(row)

            all_rows.extend(new_rows)

            log(
                f"Containers: {len(records)} | "
                f"New unique records: {len(new_rows)} | "
                f"Total: {len(all_rows)}"
            )

            if len(all_rows) >= max_items:

                log("Maximum item limit reached.")
                break

            if pagination["type"] == "single":
                break

            if pagination["type"] == "page_number":

                page_number += 1

            elif pagination["type"] == "next_link":

                try:
                    next_tag = soup.select_one(
                        pagination["selector"]
                    )

                except Exception as exc:
                    raise ValueError(
                        f"Invalid next-page CSS selector: {exc}"
                    ) from exc

                if next_tag is None:

                    log(
                        "Next-page link not found; finished."
                    )

                    break

                next_value = next_tag.get(
                    pagination["attribute"],
                    ""
                )

                if not next_value:

                    log(
                        "Next-page link has no URL attribute.",
                        "WARN"
                    )

                    break

                next_url = validate_http_url(
                    absolute_link(
                        current_url,
                        next_value
                    )
                )

                if not same_host(start_url, next_url):

                    log(
                        "Next-page link points to another hostname.",
                        "WARN"
                    )

                    break

                current_url = next_url
                page_number += 1

            time.sleep(REQUEST_DELAY_SECONDS)

        if (
            page_number > max_pages
            and len(all_rows) < max_items
        ):

            log(
                f"Maximum page limit ({max_pages}) reached.",
                "WARN"
            )

    finally:
        session.close()

    return all_rows, fields, duplicates_removed


# ============================================================
# SAVE OUTPUT
# ============================================================

def save_data(dataframe, filename):
    """Save CSV, JSON, or Excel without overwriting existing files."""

    filename = safe_filename(filename)

    extension = os.path.splitext(filename)[1].lower()

    if extension not in (".csv", ".json", ".xlsx"):

        raise ValueError(
            "Choose a .csv, .json, or .xlsx output filename."
        )

    script_path = os.path.realpath(__file__)
    output_path = os.path.realpath(filename)

    if output_path == script_path:

        raise ValueError(
            "Output filename cannot overwrite the scraper script."
        )

    filename = unique_filename(filename)

    extension = os.path.splitext(filename)[1].lower()

    if extension == ".csv":

        # UTF-8 with BOM helps spreadsheet applications recognize
        # currency symbols and other non-ASCII text.
        dataframe.to_csv(
            filename,
            index=False,
            encoding="utf-8-sig"
        )

    elif extension == ".json":

        dataframe.to_json(
            filename,
            orient="records",
            force_ascii=False,
            indent=2
        )

    elif extension == ".xlsx":

        dataframe.to_excel(
            filename,
            index=False,
            engine="openpyxl"
        )

    return os.path.abspath(filename)


# ============================================================
# REPORT
# ============================================================

def save_report(
    output_path,
    start_url,
    records_saved,
    duplicates_removed,
    pages_limit,
    items_limit
):

    report_name = os.path.splitext(output_path)[0] + "_report.txt"
# Avoid replacing an existing report.
    report_name = unique_filename(report_name)
  
    with open(
        report_name,
        "w",
        encoding="utf-8-sig"
    ) as report:

        report.write("SCRAPEX PRO SCRAPING REPORT\n")
        report.write("=" * 36 + "\n")
        report.write(
            f"Date: {datetime.now().isoformat(timespec='seconds')}\n"
        )
        report.write(f"Starting URL: {start_url}\n")
        report.write(f"Records saved: {records_saved}\n")
        report.write(f"Duplicates removed: {duplicates_removed}\n")
        report.write(f"Configured maximum pages: {pages_limit}\n")
        report.write(f"Configured maximum records: {items_limit}\n")
        report.write(f"Output file: {output_path}\n")
        report.write(
            "\nURL checks restrict navigation to the same hostname.\n"
        )
        report.write(
            "Price normalization is best-effort and does not "
            "convert currencies.\n"
        )

    return os.path.abspath(report_name)


# ============================================================
# MAIN PROGRAM
# ============================================================

def main():

    print("=" * 58)
    print("                 SCRAPEX PRO")
    print("          Reliable Educational Scraper")
    print("=" * 58)

    print(
        "Only scrape websites you are allowed to access.\n"
        "Respect site terms, robots.txt, and reasonable request rates.\n"
    )

    try:

        start_url = validate_http_url(
            input("Starting URL: ")
        )

    except ValueError as exc:

        log(str(exc), "ERROR")
        return

    max_items = ask_number(
        f"Maximum records (1-{MAX_ITEMS_HARD_LIMIT})",
        1,
        MAX_ITEMS_HARD_LIMIT,
        100
    )

    max_pages = ask_number(
        f"Maximum pages (1-{MAX_PAGES_HARD_LIMIT})",
        1,
        MAX_PAGES_HARD_LIMIT,
        20
    )

    record_selector = input(
        "\nCSS selector for ONE record/container "
        "(e.g. article.product_pod): "
    ).strip()

    if not record_selector:

        log("A record selector is required.", "ERROR")
        return

    raw_fields = input(
        "Field names separated by commas "
        "(e.g. title,price,link): "
    ).strip()

    fields = [
        item.strip()
        for item in raw_fields.split(",")
        if item.strip()
    ]

    if not fields or len(set(fields)) != len(fields):

        log(
            "Enter at least one unique field name.",
            "ERROR"
        )

        return

    field_rules = {}

    for field in fields:

        print(f"\nConfigure field: {field}")

        selector = input(
            "  CSS selector inside record "
            "(use . for the whole record): "
        ).strip()

        if not selector:

            log(
                "Each field needs a CSS selector.",
                "ERROR"
            )

            return

        attribute = input(
            "  HTML attribute to read "
            "(blank means text; e.g. href, title): "
        ).strip()

        field_rules[field] = (
            selector,
            attribute
        )

    trim_whitespace = ask_yes_no(
        "\nNormalize whitespace and repair text encoding?"
    )

    print("\nPagination type:")
    print("1. One page only")
    print("2. Page number in URL")
    print("3. Next-page link in the page")

    choice = input(
        "Choose 1, 2, or 3: "
    ).strip()

    if choice == "1":

        pagination = {"type": "single"}

    elif choice == "2":

        template = input(
            "URL template containing {page}, "
            "e.g. https://site.test/items?page={page}: "
        ).strip()

        if "{page}" not in template:

            log(
                "The URL template must contain {page}.",
                "ERROR"
            )

            return

        try:

            test_url = validate_http_url(
                template.replace("{page}", "1")
            )

        except ValueError as exc:

            log(str(exc), "ERROR")
            return

        if not same_host(start_url, test_url):

            log(
                "The page-number template must use the same hostname.",
                "ERROR"
            )

            return

        pagination = {
            "type": "page_number",
            "template": template
        }

    elif choice == "3":

        next_selector = input(
            "CSS selector for next-page link "
            "(e.g. li.next a): "
        ).strip()

        if not next_selector:

            log(
                "A next-page selector is required.",
                "ERROR"
            )

            return

        next_attribute = input(
            "Link URL attribute [href]: "
        ).strip() or "href"

        pagination = {
            "type": "next_link",
            "selector": next_selector,
            "attribute": next_attribute
        }

    else:

        log("Choose 1, 2, or 3.", "ERROR")
        return

    print("\nOutput formats:")
    print("1. CSV - recommended for Google Sheets")
    print("2. JSON")
    print("3. Excel workbook (.xlsx)")

    output_choice = input(
        "Choose 1, 2, or 3 [1]: "
    ).strip() or "1"

    if output_choice == "1":

        default_filename = "scraped_data.csv"

    elif output_choice == "2":

        default_filename = "scraped_data.json"

    elif output_choice == "3":

        default_filename = "scraped_data.xlsx"

    else:

        log("Choose 1, 2, or 3.", "ERROR")
        return

    output_name = input(
        f"Output filename [{default_filename}]: "
    ).strip() or default_filename

    output_name = safe_filename(output_name)

    expected_extension = {
        "1": ".csv",
        "2": ".json",
        "3": ".xlsx"
    }[output_choice]

    if not output_name.lower().endswith(expected_extension):

        output_name = os.path.splitext(output_name)[0]
        output_name += expected_extension

    try:

        rows, field_order, duplicates_removed = scrape_site(
            start_url=start_url,
            record_selector=record_selector,
            field_rules=field_rules,
            pagination=pagination,
            max_items=max_items,
            max_pages=max_pages,
            trim_whitespace=trim_whitespace
        )

        if not rows:

            log(
                "No records were collected. Check the URL and CSS selectors.",
                "WARN"
            )

            return

        dataframe = clean_dataframe(
            rows,
            field_order
        )

        dataframe = add_normalized_price_columns(
            dataframe
        )

        if ask_yes_no(
            "Remove duplicate rows from the final dataset?"
        ):

            before = len(dataframe)

            dataframe = dataframe.drop_duplicates(
                keep="first"
            ).reset_index(drop=True)

            extra_removed = before - len(dataframe)

            duplicates_removed += extra_removed

            log(
                f"Additional duplicate rows removed: {extra_removed}"
            )

        output_path = save_data(
            dataframe,
            output_name
        )

        report_path = save_report(
            output_path=output_path,
            start_url=start_url,
            records_saved=len(dataframe),
            duplicates_removed=duplicates_removed,
            pages_limit=max_pages,
            items_limit=max_items
        )

        print("\n" + "=" * 58)
        print("SCRAPING FINISHED")
        print("=" * 58)
        print("Records saved:", len(dataframe))
        print("Duplicates removed:", duplicates_removed)
        print("Output file:", output_path)
        print("Report file:", report_path)
        print("=" * 58)

        log(
            "Tip: For Google Sheets, try importing the CSV through "
            "the spreadsheet import menu if currency symbols display "
            "incorrectly."
        )

    except KeyboardInterrupt:

        log("Stopped by user.", "WARN")

    except (ValueError, OSError, csv.Error) as exc:

        log(str(exc), "ERROR")

    except Exception as exc:

        log(
            f"Unexpected error: {exc}",
            "ERROR"
        )


if __name__ == "__main__":
    main()