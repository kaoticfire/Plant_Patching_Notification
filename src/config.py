""" Configuration for the Patching functions."""


class Config:
    """ Common Variables"""
    FILE = 'site.xlsx'
    DEBUG = False
    CHANGE_TICKET = 'CHG0046418'
    WEEK_NUMBER = 3
    TIMEZONE = 'Eastern'
    DAYS = 13
    SPOKE_TO = ''
    ZONE = {'Eastern': 'EDT',
            'Central': 'CDT',
            'Mountain': 'MDT',
            'Pacific': 'PDT'}