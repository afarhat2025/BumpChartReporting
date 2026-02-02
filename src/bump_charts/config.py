"""Configuration and constants for the bump chart comparison tool."""

import os
from dotenv import load_dotenv

load_dotenv()

# Local paths (relative to project root)
SCRIPT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
PROJECT_ROOT = SCRIPT_DIR
DATA_DIR = os.path.join(PROJECT_ROOT, "data")
RESULTS_DIR = os.path.join(PROJECT_ROOT, "Results")

# Paths
NETWORK_SHARE_PATH = "/mnt/network_share"
CUSTOMER_LIST_PATH = os.path.join(DATA_DIR, "CustomerList.csv")
PART_KEY_CACHE_PATH = os.path.join(DATA_DIR, "part_key_cache.json")
LOG_FILEPATH = os.path.join(PROJECT_ROOT, "compare_prices.log")

# Input files to process - Add/remove files as needed
INPUT_FILES = [
    "3. GM Bump Charts/2R Exec - BT1YL_T1SUV/Bump Chart - BT1YL - T1SUV 2nd Row Executive - Current Q1 2026.xlsx",
    "3. GM Bump Charts/C121 Relaunch/C121 Relaunch Bump Chart - ECR 7065022 - 1.26.26.xlsm",
    "3. GM Bump Charts/EVO Front Structure/Fisher - EVO FRT STRUCTURE BUMP CHART - All Programs - Q1 2026 Material Index Adjustment.xlsx",
    "3. GM Bump Charts/Synergy/Synergy_2025 BT Uplift Generic Component Bump Chart_Updated-Plex-Codes.xlsx",
    "14. STLA Bump Charts/STLA/JL Bump Chart Q1 2026.xlsx"
]

# Output file naming
OUTPUT_PREFIX = "Delta-Report_"

# APIs
PLEX_PART_KEY_URL = "https://cloud.plex.com/api/datasources/885/execute"
PLEX_PRICE_URL = "https://cloud.plex.com/api/datasources/19770/execute"

# Defaults
DEFAULT_PCN = "SCS"

# PCN Authentication (base64 encoded username:password)
PCN_AUTH_MAP = {
    "BRO": os.getenv("BRO_AUTH"),
    "EVV": os.getenv("EVV_AUTH"),
    "TROY": os.getenv("TROY_AUTH"),
    "SCS": os.getenv("SCS_AUTH"),
    "MX": os.getenv("MX_AUTH"),
}

# Email
SMTP_SERVER = "10.10.1.151"
SMTP_PORT = 25
SMTP_USER = "no-reply@fisherco.com"
SUCCESS_RECIPIENTS = ["salesbumpchartanalysis@fisherco.com"]
ERROR_RECIPIENT = "alex.farhat@fisherco.com"
EMAIL_SUBJECT = "GM Bump Chart Report"

# Result column order
RESULT_COLUMNS = [
    "Part Number",
    "Chart Price",
    "Plex Price",
    "Delta",
    "Description",
    "PCN",
    "Program",
    "OEM Plant",
    "Customer",
    "PO_No",
    "Effective Date",
    "Status",
    "Part_Key",
]

# Price block search priorities (DDP first, then FCA)
PRICE_COLUMN_PRIORITIES = [
    "ddp price",  # First priority
    "fca price",  # Fallback
]

# Custom price column mapping for multi-customer rows
# Keys match Plex Customer Code text (case-insensitive contains),
# values match price column headers to search for (also contains).
CUSTOM_PRICE_COLUMN_MAP = {
    "lear": "factory zero",
    "adient": "lansing",
    "magna": "springhill",
}
