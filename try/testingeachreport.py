def get_days_in_month(month):
    """Returns the number of days in a given month."""
    if month in [1, 3, 5, 7, 8, 10, 12]:  # Months with 31 days
        return 31
    elif month in [4, 6, 9, 11]:  # Months with 30 days
        return 30
    elif month == 2:  # February
        return 29  # Assuming leap year; adjust as needed for non-leap years
    else:
        raise ValueError("Invalid month number. Must be between 1 and 12.")
    
x = get_days_in_month(5)
print(x)