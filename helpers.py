from datetime import datetime, timedelta
from typing import Tuple, List

def calculate_stay_duration(check_in: str, check_out: str) -> int:
    """Calculate number of nights between check-in and check-out dates.
    
    Args:
        check_in (str): Check-in date (YYYY-MM-DD)
        check_out (str): Check-out date (YYYY-MM-DD)
        
    Returns:
        int: Number of nights
    """
    check_in_date = datetime.strptime(check_in, "%Y-%m-%d")
    check_out_date = datetime.strptime(check_out, "%Y-%m-%d")
    return (check_out_date - check_in_date).days

def calculate_total_amount(nights: int, nightly_rate: float) -> float:
    """Calculate total amount for stay.
    
    Args:
        nights (int): Number of nights
        nightly_rate (float): Price per night
        
    Returns:
        float: Total amount
    """
    return nights * nightly_rate

def get_date_range(period: str) -> Tuple[datetime, datetime]:
    """Get start and end dates for a given period.
    
    Args:
        period (str): One of 'week', 'month', 'year'
        
    Returns:
        Tuple[datetime, datetime]: Start and end dates
    """
    today = datetime.now()
    
    if period == 'week':
        start_date = today - timedelta(days=today.weekday())
        end_date = start_date + timedelta(days=6)
    elif period == 'month':
        start_date = today.replace(day=1)
        if today.month == 12:
            end_date = today.replace(year=today.year + 1, month=1, day=1) - timedelta(days=1)
        else:
            end_date = today.replace(month=today.month + 1, day=1) - timedelta(days=1)
    elif period == 'year':
        start_date = today.replace(month=1, day=1)
        end_date = today.replace(month=12, day=31)
    else:
        raise ValueError("Period must be one of: 'week', 'month', 'year'")
    
    return start_date, end_date

def format_currency(amount: float) -> str:
    """Format amount as currency string.
    
    Args:
        amount (float): Amount to format
        
    Returns:
        str: Formatted currency string
    """
    return f"₺{amount:,.2f}"

def validate_dates(check_in: str, check_out: str) -> bool:
    """Validate check-in and check-out dates.
    
    Args:
        check_in (str): Check-in date (YYYY-MM-DD)
        check_out (str): Check-out date (YYYY-MM-DD)
        
    Returns:
        bool: True if dates are valid, False otherwise
    """
    try:
        check_in_date = datetime.strptime(check_in, "%Y-%m-%d")
        check_out_date = datetime.strptime(check_out, "%Y-%m-%d")
        
        if check_in_date >= check_out_date:
            return False
            
        if check_in_date < datetime.now().replace(hour=0, minute=0, second=0, microsecond=0):
            return False
            
        return True
    except ValueError:
        return False 