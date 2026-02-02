"""Data models and utility functions for bump chart processing."""

from dataclasses import dataclass
from datetime import datetime
from typing import Optional


@dataclass
class ExcelPart:
    """Represents a part extracted from Excel."""
    part_number: int
    price: float
    program: str
    pcn: str
    customer: str
    oem_plant: str
    description: str
    date: Optional[datetime]


@dataclass
class PriceResult:
    """Result of price comparison for a part."""
    part_number: str
    chart_price: float
    plex_price: float
    delta: float
    description: str
    pcn: str
    program: str
    oem_plant: str
    customer: str
    po_no: str
    effective_date: Optional[datetime]
    status: str
    part_key: Optional[str]

    def to_dict(self):
        """Convert to dictionary for DataFrame."""
        return {
            "Part Number": self.part_number,
            "Chart Price": self.chart_price,
            "Plex Price": self.plex_price,
            "Delta": self.delta,
            "Description": self.description,
            "PCN": self.pcn,
            "Program": self.program,
            "OEM Plant": self.oem_plant,
            "Customer": self.customer,
            "PO_No": self.po_no,
            "Effective Date": self.effective_date,
            "Status": self.status,
            "Part_Key": self.part_key,
        }
