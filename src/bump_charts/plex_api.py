"""Plex API interactions for part keys and pricing."""

import base64
import json
import logging
import os
from datetime import datetime
from typing import Tuple, Optional, List, Dict

import requests
from requests.auth import HTTPBasicAuth

from .config import PLEX_PART_KEY_URL, PLEX_PRICE_URL, PART_KEY_CACHE_PATH, PCN_AUTH_MAP, DEFAULT_PCN

logger = logging.getLogger("compare_prices")


def get_auth_for_pcn(pcn: str) -> Tuple[str, str]:
    """Decode base64 auth for a given PCN."""
    b64 = PCN_AUTH_MAP.get(str(pcn), PCN_AUTH_MAP[DEFAULT_PCN])
    decoded = base64.b64decode(b64).decode("utf-8").strip()
    if ":" in decoded:
        return tuple(decoded.split(":", 1))
    raise ValueError(f"Invalid base64 for PCN {pcn}")


def load_part_key_cache() -> Dict:
    """Load cached part keys from file."""
    if os.path.exists(PART_KEY_CACHE_PATH):
        try:
            with open(PART_KEY_CACHE_PATH, "r") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_part_key_cache(cache: Dict) -> None:
    """Save part key cache to file."""
    try:
        with open(PART_KEY_CACHE_PATH, "w") as f:
            json.dump(cache, f)
    except Exception:
        pass


def retrieve_part_key(
    part_no: str,
    customer_code_raw: str,
    username: str,
    password: str,
    cache: Dict
) -> Tuple[Optional[str], str]:
    """
    Retrieve Part_Key from Plex API.
    
    Returns:
        Tuple of (part_key, status_message)
    """
    status = ""
    normalized_customer = (customer_code_raw or "").strip().lower()
    cache_key = _build_cache_key(part_no, normalized_customer)
    
    # Check cache
    part_key = cache.get(cache_key)
    if not part_key and not normalized_customer:
        part_key = cache.get(part_no)
    if part_key:
        return part_key, status
    
    try:
        # Try lookup with Part_No first
        rows, headers = _query_part_key_api(
            username, password, part_no, "Part_No"
        )
        part_key, status = _select_part_key_from_rows(
            rows, headers, normalized_customer, True
        )
        
        # If Part_No lookup failed, try Customer_Part_No as fallback
        if not part_key:
            logger.info(
                "[885] Part_Key not found via Part_No lookup for %s; "
                "retrying with Customer_Part_No.",
                part_no
            )
            rows, headers = _query_part_key_api(
                username, password, part_no, "Customer_Part_No"
            )
            part_key, status = _select_part_key_from_rows(
                rows, headers, normalized_customer, True
            )
        
        # Cache the result
        if part_key:
            cache[cache_key] = part_key
            if not normalized_customer:
                cache[part_no] = part_key
            save_part_key_cache(cache)
    
    except requests.Timeout:
        part_key = None
        status = "Timed Out"
        logger.error("[885] Timeout retrieving Part_Key for %s", part_no)
    except Exception:
        logger.exception("[885] Error retrieving Part_Key for %s", part_no)
        part_key = None
    
    return part_key, status


def query_price_api(
    part_key: str,
    date_begin_str: str,
    username: str,
    password: str,
    chart_price: float,
    customer_name: str,
    all_customer_names: set
) -> Tuple[float, str, str]:
    """
    Query Plex API for pricing information.
    
    Returns:
        Tuple of (api_price, status_message, po_no)
    """
    payload = {
        "inputs": {
            "Part_Key": part_key,
            "Shipper_Date_Begin": date_begin_str,
        }
    }
    
    api_price = chart_price
    status = ""
    po_no = ""
    today = datetime.today()
    
    logger.info(
        "[19770] Requesting price for part_key=%s payload=%s",
        part_key, payload
    )
    
    try:
        response = requests.post(
            PLEX_PRICE_URL,
            auth=HTTPBasicAuth(username, password),
            json=payload,
            timeout=24,
        )
        logger.info("[19770] Response status %s for part_key %s", response.status_code, part_key)
        
        if response.status_code == 200:
            data = response.json()
            table = data.get("tables", [])[0] if data.get("tables") else None
            rows = table.get("rows", []) if table else []
            headers = table.get("columns", []) if table else []
            
            logger.info("[19770] Returned %s rows for part_key=%s", len(rows), part_key)
            
            if not rows or not headers:
                status = "No PO's exist yet"
                return api_price, status, po_no
            
            api_price, status, po_no = _select_price_from_rows(
                rows, headers, customer_name, all_customer_names, 
                chart_price, date_begin_str, today
            )
        else:
            status = f"PO API error: {response.status_code} {response.text}"
            logger.error("[19770] Error response: %s", status)
    
    except requests.Timeout:
        status = "Timed Out"
        logger.error("[19770] Timeout retrieving price for part_key %s", part_key)
    except Exception as exc:
        status = f"PO API exception: {exc}"
        logger.exception("[19770] Exception retrieving price for part_key %s", part_key)
    
    return api_price, status, po_no


def _build_cache_key(part_no: str, customer_code: str) -> str:
    """Build a cache key for part key lookup."""
    if not customer_code:
        customer_code = "none"
    return f"{part_no}::{customer_code}"


def _query_part_key_api(
    username: str,
    password: str,
    part_no: str,
    lookup_field: str
) -> Tuple[List, List]:
    """Execute part key API query."""
    payload = {"inputs": {lookup_field: part_no, "Active_Only": 1}}
    
    logger.info("[885] Requesting Part_Key for %s with payload %s", part_no, payload)
    response = requests.post(
        PLEX_PART_KEY_URL,
        auth=HTTPBasicAuth(username, password),
        json=payload,
        timeout=24,
    )
    logger.info("[885] Response status %s for part %s", response.status_code, part_no)
    
    if response.status_code == 200:
        data = response.json()
        table = data.get("tables", [])[0] if data.get("tables") else None
        rows = table.get("rows", []) if table else []
        headers = table.get("columns", []) if table else []
        logger.info("[885] Returned %s candidate rows for %s", len(rows), part_no)
        return rows, headers
    
    logger.error(
        "[885] Non-200 response %s retrieving Part_Key for %s: %s",
        response.status_code, part_no, response.text
    )
    return [], []


def _select_part_key_from_rows(
    rows: List,
    headers: List,
    customer_code: str,
    enforce_customer: bool
) -> Tuple[Optional[str], str]:
    """Select best Part_Key from response rows."""
    if not rows or not headers:
        return None, ""
    
    try:
        part_key_idx = headers.index("Part_Key")
        status_idx = headers.index("Part_Status")
        revision_idx = headers.index("Revision")
        cust_code_idx = headers.index("Customer_Code")
    except ValueError:
        return None, ""
    
    status_order = ["Service", "Production", "Prototype", "Development", "Obsolete"]
    
    # Filter by customer if specified
    matching_rows = []
    if customer_code:
        for row in rows:
            row_code = str(row[cust_code_idx]).strip().lower()
            if row_code == customer_code:
                matching_rows.append(row)
    
    candidate_rows = matching_rows if enforce_customer else (matching_rows or rows)
    if not candidate_rows:
        return None, ""
    
    # Sort by status, then revision
    def sort_key(row):
        status_val = str(row[status_idx]) if row[status_idx] is not None else ""
        try:
            revision_val = float(row[revision_idx]) if row[revision_idx] not in [None, ""] else -float("inf")
        except Exception:
            revision_val = -float("inf")
        rank = status_order.index(status_val) if status_val in status_order else len(status_order)
        return (rank, -revision_val)
    
    best_row = sorted(candidate_rows, key=sort_key)[0]
    selected_key = best_row[part_key_idx]
    
    logger.info(
        "[885] Selected Part_Key %s (status=%s, revision=%s)",
        selected_key,
        best_row[status_idx] if status_idx is not None else "",
        best_row[revision_idx] if revision_idx is not None else ""
    )
    
    return selected_key, ""


def _select_price_from_rows(
    rows: List,
    headers: List,
    customer_name: str,
    all_customer_names: set,
    chart_price: float,
    date_begin_str: str,
    today: datetime
) -> Tuple[float, str, str]:
    """Select best price from API response rows."""
    try:
        cust_name_idx = headers.index("Customer_Name")
        price_idx = headers.index("Unit_Price")
        date_idx = headers.index("Require_Ship_Date")
        po_no_idx = headers.index("PO_No") if "PO_No" in headers else None
    except ValueError:
        return chart_price, "PO API missing required columns", ""
    
    # Parse start date threshold
    start_threshold = None
    if date_begin_str:
        try:
            start_threshold = datetime.fromisoformat(date_begin_str.replace("Z", "+00:00")).replace(tzinfo=None)
        except Exception:
            pass
    
    # Filter rows by customer if applicable
    matched_customer = None
    if customer_name:
        matched_customer = _fuzzy_match_customer(customer_name, all_customer_names)
    
    filtered_rows = []
    if matched_customer:
        for row in rows:
            api_cust = str(row[cust_name_idx]).strip().lower()
            if matched_customer in api_cust or api_cust in matched_customer:
                if abs(len(api_cust) - len(matched_customer)) <= 2:
                    filtered_rows.append(row)
    else:
        filtered_rows = rows
    
    if not filtered_rows:
        if matched_customer:
            return chart_price, f"No PO's under {customer_name} found", ""
        else:
            return chart_price, "No PO's exist yet", ""
    
    # Select best row from candidates
    selected_row, selected_date, used_old_price = _pick_best_price_row(
        filtered_rows, price_idx, date_idx, po_no_idx,
        chart_price, start_threshold, today
    )
    
    if selected_row is None:
        return chart_price, "No suitable PO found", ""
    
    api_price = selected_row[price_idx]
    po_no = selected_row[po_no_idx] if po_no_idx is not None else ""
    status = "Success"
    
    if used_old_price and selected_date:
        status = f"Old Price from {selected_date:%Y-%m-%d}"
    
    logger.info(
        "[19770] Selected PO %s with price %s (status=%s)",
        po_no, api_price, status
    )
    
    return api_price, status, po_no


def _fuzzy_match_customer(target: str, candidates: set) -> Optional[str]:
    """Fuzzy match customer name."""
    target = target.lower().strip()
    for cand in candidates:
        cand_l = cand.lower().strip()
        if target in cand_l or cand_l in target:
            if abs(len(target) - len(cand_l)) <= 2:
                return cand
    return None


def _pick_best_price_row(
    rows: List,
    price_idx: int,
    date_idx: int,
    po_no_idx: Optional[int],
    chart_price: float,
    start_threshold: Optional[datetime],
    today: datetime
) -> Tuple[Optional[List], Optional[datetime], bool]:
    """Select the best price row based on date and price proximity.

    - Prefer rows with Require_Ship_Date on/after start_threshold (Excel date).
    - If none qualify, fall back to the oldest available row and flag status accordingly.
    """
    if not rows:
        return None, None, False
    
    import pandas as pd
    
    usable_recent = []
    older_rows = []
    for row in rows:
        try:
            ts = pd.to_datetime(row[date_idx], errors="coerce")
            if pd.isna(ts):
                row_date = None
            else:
                if isinstance(ts, pd.Timestamp) and ts.tzinfo is not None:
                    ts = ts.tz_convert(None)
                row_date = ts.to_pydatetime().replace(tzinfo=None) if hasattr(ts, "to_pydatetime") else ts
        except Exception:
            row_date = None
        
        # Skip future dates
        if row_date and row_date > today:
            continue
        
        if start_threshold and row_date is not None and row_date < start_threshold:
            older_rows.append((row, row_date))
        else:
            usable_recent.append((row, row_date))
    
    # If we have recent/valid rows (>= threshold or unknown date), use them
    if usable_recent:
        best_date = max(item[1] for item in usable_recent if item[1] is not None) if any(item[1] for item in usable_recent) else None
        candidates = [item[0] for item in usable_recent if item[1] == best_date] if best_date else [item[0] for item in usable_recent]
        
        if len(candidates) > 1 and chart_price is not None:
            def price_gap(row):
                try:
                    return abs(float(row[price_idx]) - chart_price)
                except Exception:
                    return float("inf")
            selected = min(candidates, key=price_gap)
        else:
            selected = candidates[0]
        return selected, best_date, False
    
    # No recent rows; fall back to oldest available (if any)
    if older_rows:
        oldest_date = min(item[1] for item in older_rows if item[1] is not None) if any(item[1] for item in older_rows) else None
        candidates = [item[0] for item in older_rows if item[1] == oldest_date] if oldest_date else [item[0] for item in older_rows]
        
        if len(candidates) > 1 and chart_price is not None:
            def price_gap(row):
                try:
                    return abs(float(row[price_idx]) - chart_price)
                except Exception:
                    return float("inf")
            selected = min(candidates, key=price_gap)
        else:
            selected = candidates[0]
        return selected, oldest_date, True
    
    return None, None, False
