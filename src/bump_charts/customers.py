"""Customer metadata management."""

import pandas as pd
from typing import Dict, Set, Tuple

from .config import CUSTOMER_LIST_PATH


def load_customer_metadata() -> Tuple[pd.DataFrame, Dict[str, str], Set[str]]:
    """
    Load customer information from CSV.
    
    Returns:
        Tuple of (dataframe, code_to_name_dict, name_set)
    """
    customer_df = pd.read_csv(CUSTOMER_LIST_PATH, dtype=str)
    code_to_name = dict(
        zip(
            customer_df["Customer_Code"].str.strip().str.lower(),
            customer_df["Customer_Name"].str.strip(),
        )
    )
    name_set = set(customer_df["Customer_Name"].str.strip().str.lower())
    return customer_df, code_to_name, name_set
