"""Excel reader for bump chart parts."""

import pandas as pd
from datetime import datetime
from typing import List, Optional
from .models import ExcelPart
from .config import PRICE_COLUMN_PRIORITIES, CUSTOM_PRICE_COLUMN_MAP


def get_excel_part_info(excel_path: str) -> List[ExcelPart]:
    """
    Extract the most recent valid part number and price from each row in Excel file.
    
    For each sheet, finds the main table (with Part Number and Part Description),
    then locates price blocks with dates. Returns the most recent dated price <= today.
    
    Args:
        excel_path: Path to the Excel file
        
    Returns:
        List of ExcelPart objects
    """
    today = datetime.today().date()
    xls = pd.ExcelFile(excel_path)
    all_parts = []
    
    for sheet_name in xls.sheet_names:
        print(f"\n=== Processing Sheet: {sheet_name} ===")
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
        
        # Find the main table header row (must have 'Part Number' and 'Part Description')
        start_table_row = _find_main_table_header(df)
        
        if start_table_row is None:
            print(f"  No main table found in sheet {sheet_name}")
            continue
            
        print(f"  Main table starts at row {start_table_row}")
        
        # Get column indices from main table header
        header = df.iloc[start_table_row].astype(str).str.strip().tolist()
        col_indices = {col.lower(): idx for idx, col in enumerate(header)}
        
        # Find all price blocks
        price_blocks = _find_price_blocks(df, start_table_row)
        
        print(f"  Found {len(price_blocks)} price blocks in sheet {sheet_name}")
        for idx, block in enumerate(price_blocks):
            price_col_display = block["default_price_col"] if block["default_price_col"] is not None else "None"
            print(
                f"    Block {idx}: date={block['date']}, part_col={block['part_col']}, price_col={price_col_display}"
            )
        
        # Extract parts from main table rows
        parts_in_sheet = _extract_parts_from_rows(
            df, start_table_row, col_indices, price_blocks, today
        )
        all_parts.extend(parts_in_sheet)
    
    print(f"\nTotal parts extracted: {len(all_parts)}")
    return all_parts


def _find_main_table_header(df: pd.DataFrame) -> int:
    """Find the row containing main table headers (Part Number and Part Description)."""
    for i, row in df.iterrows():
        header = row.astype(str).str.strip().tolist()
        header_lower = [c.lower() for c in header]
        if "part number" in header_lower and "part description" in header_lower:
            return i
    return None


def _find_price_blocks(df: pd.DataFrame, start_table_row: int) -> List[dict]:
    """Find all price blocks (Part Number columns with associated price columns and dates)."""
    price_blocks = []
    
    for i, row in df.iterrows():
        header_values = row.astype(str).str.strip().tolist()
        header_lower = [col.lower() for col in header_values]
        for idx, col in enumerate(header_lower):
            if col == "part number":
                default_price_col = _find_price_column(header_values, idx)
                date_val = _find_date_above(df, i, idx)
                price_blocks.append({
                    "header_row": i,
                    "header_values": header_values,
                    "part_col": idx,
                    "default_price_col": default_price_col,
                    "date": date_val
                })
    
    return price_blocks


def _find_price_column(header_values: list, part_col_idx: int) -> Optional[int]:
    """Find default price column (exact match DDP/FCA, then contains match)."""
    header_lower = [col.lower() for col in header_values]
    start_idx = part_col_idx + 1
    
    # Try exact match first
    for priority in PRICE_COLUMN_PRIORITIES:
        for j in range(start_idx, len(header_lower)):
            if header_lower[j] == priority:
                return j
    
    # Fallback to contains match
    for priority in PRICE_COLUMN_PRIORITIES:
        for j in range(start_idx, len(header_lower)):
            if priority in header_lower[j]:
                return j
    
    return None


def _map_customer_to_header(customer_name: str) -> Optional[str]:
    """Map a Plex Customer Code value to the target price column header text."""
    customer_lower = customer_name.lower()
    for key, mapped_header in CUSTOM_PRICE_COLUMN_MAP.items():
        if key.lower() in customer_lower:
            return mapped_header
    return None


def _select_price_column(
    header_values: list,
    part_col_idx: int,
    customer_name: Optional[str]
) -> Optional[int]:
    """Choose price column for a row (multi-customer or single).

    - If comma-separated customer: use mapped header (contains match).
    - If no comma (single customer/no customer): try exact DDP/FCA, then contains.
    """
    header_lower = [col.lower() for col in header_values]
    start_idx = part_col_idx + 1
    
    # If customer name provided, try the mapped header (contains match)
    if customer_name:
        mapped_header = _map_customer_to_header(customer_name)
        if mapped_header:
            target = mapped_header.lower()
            for j in range(start_idx, len(header_lower)):
                if target in header_lower[j]:
                    return j
            # Mapped header not found; fall through to default
    
    # Default logic: exact match first, then contains
    # Try exact match
    for priority in PRICE_COLUMN_PRIORITIES:
        for j in range(start_idx, len(header_lower)):
            if header_lower[j] == priority:
                return j
    
    # Fallback to contains match
    for priority in PRICE_COLUMN_PRIORITIES:
        for j in range(start_idx, len(header_lower)):
            if priority in header_lower[j]:
                return j
    
    return None


def _find_date_above(df: pd.DataFrame, header_row: int, col_idx: int) -> pd.Timestamp:
    """Find a valid date in cells above the given position."""
    for drow in range(header_row - 1, -1, -1):
        cell = df.iloc[drow, col_idx]
        try:
            date_val = pd.to_datetime(cell, errors="coerce").date()
            if pd.notna(date_val):
                return date_val
        except Exception:
            continue
    return None


def _extract_parts_from_rows(
    df: pd.DataFrame,
    start_table_row: int,
    col_indices: dict,
    price_blocks: List[dict],
    today
) -> List[ExcelPart]:
    """Extract parts from main table rows by finding most recent valid price."""
    parts = []
    
    # Get column indices for metadata
    program_idx = col_indices.get("program")
    pcn_idx = col_indices.get("fisher pcn")
    customer_idx = col_indices.get("plex customer code")
    oem_plant_idx = col_indices.get("oem plant")
    desc_idx = col_indices.get("part description")
    
    # Process each data row
    for k in range(start_table_row + 1, len(df)):
        row = df.iloc[k]
        
        # Get metadata from main table
        program = row[program_idx] if program_idx is not None else ""
        pcn = row[pcn_idx] if pcn_idx is not None and str(row[pcn_idx]).strip() else "95506"
        customer_raw = row[customer_idx] if customer_idx is not None else ""
        oem_plant = row[oem_plant_idx] if oem_plant_idx is not None else ""
        description = row[desc_idx] if desc_idx is not None else ""
        
        # Determine if this row needs custom price column selection (comma-separated customers)
        customer_text = str(customer_raw).strip()
        has_custom_prices = "," in customer_text
        customer_names = [c.strip() for c in customer_text.split(",") if c.strip()] if has_custom_prices else [customer_text]
        
        for customer_name in customer_names:
            # Find most recent valid price for this row/customer
            result = _find_most_recent_price(
                row,
                k,
                price_blocks,
                today,
                customer_name if has_custom_prices else None
            )
            
            if result is not None:
                part_number, price, date = result
                parts.append(ExcelPart(
                    part_number=part_number,
                    price=price,
                    program=str(program).strip(),
                    pcn=str(pcn).strip() if str(pcn).strip() else "95506",
                    customer=str(customer_name).strip(),
                    oem_plant=str(oem_plant).strip(),
                    description=str(description).strip(),
                    date=date
                ))
    
    return parts


def _find_most_recent_price(
    row,
    row_idx: int,
    price_blocks: List[dict],
    today,
    customer_name: Optional[str]
):
    """Find the most recent valid price (date <= today) for a row."""
    latest_date = None
    latest_part_number = None
    latest_price = None
    
    for block in price_blocks:
        date_val = block["date"]
        if date_val is None or pd.isna(date_val) or date_val > today:
            continue
        
        if row_idx <= block["header_row"]:
            continue
        
        price_col = _select_price_column(
            block["header_values"],
            block["part_col"],
            customer_name
        )
        if price_col is None:
            continue

        part_number = row[block["part_col"]]
        price = row[price_col]
        
        try:
            part_number_int = int(str(part_number).strip())
            price_val = float(str(price).replace("$", "").replace(",", "").strip())
        except Exception:
            continue
        
        # Use the most recent valid date
        if (latest_date is None) or (date_val > latest_date):
            latest_date = date_val
            latest_part_number = part_number_int
            latest_price = price_val
    
    if latest_part_number is not None and latest_price is not None:
        return latest_part_number, latest_price, latest_date
    
    return None
