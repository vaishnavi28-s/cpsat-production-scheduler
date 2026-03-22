from dataclasses import dataclass, field


@dataclass
class PrintJob:
    """Represents a single print job fetched from Snowflake."""

    JOB: str
    PRESS_LOCATION: str
    SEND_TO_LOCATION: str
    PRODUCTTYPE: str
    PAPER: str
    FINISHTYPE: str
    FINISHINGOP: str
    DELIVERYDATE: str
    INKSS1: str
    INKSS2: str
    QUANTITYORDERED: int = 0
    PAGES: int = 0

    def __post_init__(self):
        self.QUANTITYORDERED = int(self.QUANTITYORDERED or 0)
        self.PAGES = int(self.PAGES or 0)

    def group_key(self) -> tuple:
        """Group jobs by product type, press location, destination and paper."""
        return (
            self.PRODUCTTYPE,
            self.PRESS_LOCATION,
            self.SEND_TO_LOCATION,
            self.PAPER,
        )

    def is_cover(self) -> bool:
        return "cover" in (self.PRODUCTTYPE or "").strip().lower()

    def is_jacket(self) -> bool:
        return "jacket" in (self.PRODUCTTYPE or "").strip().lower()
