from pydantic import BaseModel, field_validator, field_serializer
from typing import List, Optional, Tuple, Union
from datetime import date

from .transaction import Transaction

class Settlement(BaseModel):
  legal_name: str
  rif_cedula: str
  code: str
  description: str
  account_number: str
  bank: str
  amount: float
  reference: List[str]
  reference_raw: str

  paid_at: List[date] = []
  paid_at_raw: str

  settled_at: date

  is_exonerated: bool = False
  is_voided: bool = False

  payments: List[Transaction] = []
  not_found_payments: List[str] = []

  amount_difference: float = 0.0

  is_verified: bool = True

  @field_serializer('reference', mode='plain')
  def serialize_settled_at(self, value: List[str]) -> str:
    return '-'.join(value)
  
  @field_serializer('paid_at', mode='plain')
  def serialize_paid_at(self, value: List[date]) -> str:
    if len(value) == 0:
      return None
    elif len(value) == 1:
      return value[0].strftime('%d/%m/%Y')
    else:
      return '-'.join([v.strftime('%d/%m/%Y') for v in value])
    
  @field_serializer('not_found_payments', mode='plain')
  def serialize_not_found_payments(self, value: List[str]) -> str:
    return '-'.join(value)

  # @field_validator('bank', mode='before')
  # def validate_bank(cls, v):
  #   allowed_banks = ['Banco de Venezuela', 'BIOPAGO', 'BDT']
  #   if v not in allowed_banks:
  #     raise ValueError(f'bank must be one of {allowed_banks}')
  #   return v

  # @field_validator('account_number', mode='before')
  # def validate_account_number(cls, v):
  #   allowed_account_numbers = ['1892', '9290']
  #   if v not in allowed_account_numbers:
  #     raise ValueError(f'account_number must be one of {allowed_account_numbers}')
  #   return v
