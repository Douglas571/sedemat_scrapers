from datetime import date
from pydantic import BaseModel, field_validator, validator
from typing import Optional

class Transaction(BaseModel):
  date: date
  reference: str
  description: str
  amount: float

  bank: Optional[str] = None
  account_number: Optional[str] = None
  matched_settlement_code: Optional[str] = None
  settlement_date: Optional[date] = None
  settlement_description: Optional[str] = None

  @field_validator('bank', mode='before')
  def validate_bank(cls, v):
    allowed_banks = ['Banco de Venezuela', 'BIOPAGO', 'BDT']
    if v not in allowed_banks:
      raise ValueError(f'bank must be one of {allowed_banks}')
    return v

  @field_validator('account_number', mode='before')
  def validate_account_number(cls, v):
    allowed_account_numbers = ['1892', '9290']
    if v not in allowed_account_numbers:
      raise ValueError(f'account_number must be one of {allowed_account_numbers}')
    return v
