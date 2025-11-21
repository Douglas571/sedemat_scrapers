from pydantic import BaseModel
from typing import Optional

class Payment(BaseModel):
  date: str
  reference: str
  description: str
  amount: float

  bank: Optional[str] = None
  account_number: Optional[str] = None
  matched_settlement_code: Optional[str] = None
  settlement_date: Optional[str] = None
  settlement_description: Optional[str] = None
  