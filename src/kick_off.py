""" Email when about to start with plant patching. """
from datetime import datetime as dt
from dateutil.relativedelta import relativedelta
from config import Config as cfg
from filelog import file_log
from mail import __mail_notification

if dt.today() <= dt(int(dt.today().year), int(dt.today().month), cfg.DAYS):
    month = dt.date(dt.today() - relativedelta(months=1))
else:
    month = dt.today()

if cfg.DEBUG:
    to_who = 'virgil.hoover@cavalry.solutions'
else:
    to_who = 'overwatch@cavalry.solutions'

msg_subj = f'{cfg.CHANGE_TICKET} - Plant Prod Monthly Patching - {month}'
msg_body= f'<html><body><p>Overwatch,<br />Patching for the ' \
                   f'plants will be starting in 2 hours.<br /><br />' \
                   f'Virgil Hoover<br />SUPPORT@CAVALRY.SOLUTIONS<br />' \
                   f'NETWORK OPERATIONS CENTER ///<br />' \
                   f'<a href="tel:+1 (720) 279-2233">+1 (720) 279-2233' \
                   f'</a> - 24X7 OVERWATCH</p></body></html>'
__mail_notification(to_who, msg_subj, msg_body)
file_log('Start email sent')
