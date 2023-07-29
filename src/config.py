""" Configuration for the Patching functions."""


class Config:
    """ Common Variables"""
    FILE = 'site.xlsx'
    DEBUG = False
    CHANGE_TICKET = 'CHG0046373'
    WEEK_NUMBER = 3
    TIMEZONE = 'Eastern'
    DAYS = 13
    ZONE = {'Eastern': 'EDT',
            'Central': 'CDT',
            'Mountain': 'MDT',
            'Pacific': 'PDT'}