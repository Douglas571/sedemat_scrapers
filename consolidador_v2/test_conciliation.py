from main import asigne_payments_to_settlements, parse_date_intervals, dates_are_close_by, is_date_in_range

import pytest
from datetime import date, datetime

def test_parse_date_intervals():
    # Test valid range
    assert parse_date_intervals("01/01/2024 - 05/01/2024") == [date(2024, 1, 1), date(2024, 1, 5)]
    # Test 2-digit years
    assert parse_date_intervals("01-01-25") == [date(2025, 1, 1), date(2025, 1, 1)]
    # Test multiple dates (finds min/max)
    assert parse_date_intervals("10/02/2024, 01/01/2024, 20/02/2024") == [date(2024, 1, 1), date(2024, 2, 20)]
    # Test invalid date
    assert parse_date_intervals("99/99/2024") is None

def test_dates_are_close_by():
    d1 = date(2024, 1, 1)
    d2 = date(2024, 1, 3)
    assert dates_are_close_by(d1, d2, 2) is True
    assert dates_are_close_by(d1, d2, 1) is False

def test_is_date_in_range():
    target = date(2024, 1, 2)
    time_range = [date(2024, 1, 1), date(2024, 1, 3)]
    
    assert is_date_in_range(target, time_range) is True
    # Test Biopago offset logic (-1/+1 day)
    biopago_target = date(2023, 12, 31)
    assert is_date_in_range(biopago_target, time_range, isBiopago=True) is True

def test_matching_logic_bdt_reference():
    # Mock a settlement
    settlements = [{
        'num_comprobante': '101',
        'legal_name': 'Test Corp',
        'rif_cedula': 'J-123',
        'pago_por': 'Service',
        'fecha': date(2024, 1, 1),
        'fecha_pago': date(2024, 1, 1),
        'banco': 'BDT',
        'referencia': '123456',
        'monto': 100.0,
        'payments': [{'reference': '123456', 'amount': 100.0, 'not_found': True}],
        'is_verified': True,
        'is_exonerated': False
    }]
    
    # Mock a BDT bank payment
    payments = [{
        'date': date(2024, 1, 1),
        'reference': '333333',
        'description': 'Something 123456', # BDT matches on description suffix
        'amount': 100.0,
        'bank': 'BDT',
        'account_number': '9290',
        'matched_settlement': 'None'
    }]
    
    asigne_payments_to_settlements(payments, settlements)
    
    # Assertions
    assert settlements[0]['is_verified'] is True
    assert settlements[0]['payments'][0]['not_found'] is False
    assert payments[0]['matched_settlement'] == '101'

def test_matching_logic_amount_difference():
    settlements = [{
        'num_comprobante': '102',
        'monto': 500.0,
        'fecha_pago': date(2024, 1, 1),
        'banco': 'Other',
        'referencia': '777777',
        'payments': [{'reference': '777777', 'amount': 500.0, 'not_found': True}],
        'is_exonerated': False,
        'legal_name': 'X', 'rif_cedula': 'Y', 'pago_por': 'Z', 'fecha': date(2024, 1, 1)
    }]
    
    # Payment exists but amount is different
    payments = [{
        'date': date(2024, 1, 1),
        'reference': '777777',
        'description': 'Normal payment',
        'amount': 450.0, # Difference of 50
        'bank': 'Other',
        'matched_settlement': 'None'
    }]
    
    asigne_payments_to_settlements(payments, settlements)
    
    assert settlements[0]['is_verified'] is False
    assert settlements[0]['amount_difference'] == -50.0