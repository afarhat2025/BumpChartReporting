"""Main script for comparing bump chart prices with Plex API."""

import logging
import os
import sys
from datetime import datetime

# Add src to path for imports
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from bump_charts.config import (
    INPUT_FILES,
    OUTPUT_PREFIX,
    NETWORK_SHARE_PATH,
    LOG_FILEPATH,
    RESULTS_DIR,
    DEFAULT_PCN,
)
from bump_charts.readers import get_excel_part_info
from bump_charts.customers import load_customer_metadata
from bump_charts.plex_api import (
    get_auth_for_pcn, load_part_key_cache, retrieve_part_key, query_price_api
)
from bump_charts.reports import write_results, send_success_email_with_attachments
from bump_charts.models import PriceResult
from bump_charts.utils import compute_effective_month_start, format_price

# Configure logging
logging.basicConfig(
    filename=LOG_FILEPATH,
    filemode="w",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("compare_prices")


def main():
    """Main entry point for price comparison."""
    # Load dependencies
    _, customer_code_to_name, customer_names = load_customer_metadata()
    part_key_cache = load_part_key_cache()
    
    attachment_paths = []
    run_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_output_dir = os.path.join(RESULTS_DIR, run_timestamp)
    os.makedirs(run_output_dir, exist_ok=True)
    
    # Process each input file
    for filename in INPUT_FILES:
        logger.info("=" * 80)
        logger.info("Processing file: %s", filename)
        logger.info("=" * 80)
        
        # Construct file paths
        input_filepath = os.path.join(NETWORK_SHARE_PATH, filename)
        base_name = os.path.splitext(os.path.basename(filename))[0]
        output_filename = OUTPUT_PREFIX + base_name + ".xlsx"
        output_filepath = os.path.join(run_output_dir, output_filename)
        
        # Get parts from Excel
        excel_parts = get_excel_part_info(input_filepath)
        logger.info("Found %s Excel parts in %s", len(excel_parts), filename)
        
        # Process parts
        price_lookup_cache = {}
        results = []
        
        _print_header()
        
        for part in excel_parts:
            result = _process_part(
                part,
                customer_code_to_name,
                customer_names,
                part_key_cache,
                price_lookup_cache
            )
            if result:
                if result.status == "No part found (no Part_Key)":
                    # Do not include in outbound report; log only
                    logger.info(
                        "Skipping part %s due to missing Part_Key (status logged only)",
                        result.part_number
                    )
                    continue
                results.append(result)
        
        # Write results for this file
        write_results(results, output_filepath, send_email=False)
        attachment_paths.append(output_filepath)
        logger.info("Completed processing %s. Processed %s parts.", filename, len(results))
    
    # Send one email with all attachments
    if attachment_paths:
        send_success_email_with_attachments(attachment_paths)
    
    logger.info("All files processed successfully.")


def _print_header():
    """Print the report header."""
    header_fmt = "{:<12} {:<12} {:<12} {:<12} {:<10}   {:<50} {:<12} {:<10} {:<25} {:<12}"
    header_line = header_fmt.format(
        "Part Number", "Program", "Chart Price", "Plex Price", "Delta",
        "Status", "OEM Plant", "PCN", "Customer", "Description", "Date"
    )
    separator_line = "-" * 215
    
    logger.info(header_line)
    logger.info(separator_line)
    print(header_line)
    print(separator_line)


def _process_part(part, customer_code_to_name, customer_names, part_key_cache, price_lookup_cache):
    """Process a single part and return a PriceResult."""
    part_no = str(part.part_number)
    excel_price = part.price
    pcn = str(part.pcn or DEFAULT_PCN)
    
    try:
        username, password = get_auth_for_pcn(pcn)
    except Exception as e:
        logger.error("Failed to get auth for PCN %s: %s", pcn, e)
        return None
    
    # Compute effective date
    _, shipper_begin_str = compute_effective_month_start(part.date)
    
    # Get customer code and name
    customer_code_raw = str(part.customer or "").strip().lower()
    customer_name = customer_code_to_name.get(customer_code_raw, "")
    
    # Retrieve part key
    part_key, lookup_status = retrieve_part_key(
        part_no, customer_code_raw, username, password, part_key_cache
    )
    
    # Query price from API
    api_price = excel_price
    status = lookup_status
    po_no = ""
    
    if not part_key:
        status = lookup_status or "No part found (no Part_Key)"
    else:
        cache_key = (part_key, customer_code_raw, pcn, shipper_begin_str)
        if cache_key in price_lookup_cache:
            api_price, status, po_no = price_lookup_cache[cache_key]
        else:
            api_price, status, po_no = query_price_api(
                part_key,
                shipper_begin_str,
                username,
                password,
                excel_price,
                customer_name,
                customer_names
            )
            price_lookup_cache[cache_key] = (api_price, status, po_no)
    
    # Calculate delta
    try:
        delta = float(api_price) - float(excel_price)
        delta_rounded = round(delta, 5)
        if delta_rounded == -0.0:
            delta_rounded = 0.0
    except Exception:
        delta_rounded = 0.0
    
    # Log and print the result
    row_fmt = "{:<12} {:<12} {:<12} {:<12} {:<10}   {:<50} {:<12} {:<10} {:<25} {:<12}"
    row_line = row_fmt.format(
        part_no,
        str(part.program),
        format_price(excel_price),
        format_price(api_price),
        f"{delta_rounded:.5f}",
        status,
        str(part.oem_plant),
        pcn,
        customer_name,
        str(part.description),
        str(part.date),
    )
    logger.info(row_line)
    print(row_line)
    
    # Create result object
    return PriceResult(
        part_number=part_no,
        chart_price=excel_price,
        plex_price=api_price,
        delta=delta_rounded,
        description=part.description,
        pcn=pcn,
        program=part.program,
        oem_plant=part.oem_plant,
        customer=customer_name,
        po_no=po_no,
        effective_date=part.date,
        status=status,
        part_key=part_key,
    )


if __name__ == "__main__":
    main()
