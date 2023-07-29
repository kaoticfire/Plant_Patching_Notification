""" Email when about to start with plant patching. """
from openpyxl import load_workbook
from datetime import datetime as dt
from dateutil.relativedelta import relativedelta
from filelog import file_log
from time import sleep
from config import Config as cfg
from mail import __mail_notification
from typing import IO


def _notify(change_ticket: str, date_stamp: str,
                        week_num: int, sites: list, start_time: int) -> None:
    if cfg.DEBUG:
        to = 'virgil.hoover@cavalry.solutions'
    else:
       to = 'overwatch@cavalry.solutions'
    if week_num:
        subj = f'{change_ticket} - Plant Prod Monthly Patching - {date_stamp}' \
               f' - Week {week_num}'
    else:
        subj = f'{change_ticket} - Plant Prod Monthly Patching - {date_stamp}'
    content = f'<html><body><p>Overwatch,<br />Patching for the ' \
              f'below plants will start at {start_time}:00 CDT ' \
              f'(16:00 {cfg.ZONE[cfg.TIMEZONE]}). Please ' \
              f'reach out and make sure they are aware of the <strong>' \
              f'automated process</strong>. This will ' \
              f'include the file server, the DAHS servers, the ' \
              f'EDNA server, and any CDMS workstations. Please ' \
              f'reply to this email and ' \
              f'<strong style="text-transform:uppercase">note in ' \
              f'the change</strong> who you spoke with<br />Thank ' \
              f'you,<br /><br />{"<br />".join(map(str, sites))}' \
              f'<br /><br />Virgil Hoover<br />' \
              f'SUPPORT@CAVALRY.SOLUTIONS<br />' \
              f'NETWORK OPERATIONS CENTER ///<br />' \
              f'<a href="tel:+1 (720) 279-2233">+1 (720) 279-2233' \
              f'</a> - 24X7 OVERWATCH</p></body></html>'
    __mail_notification(to, subj, content)


def get_data(week: int, _file: str, ticket: str, current: str, sheet: str) \
        -> None:
    """ Get data from the Excel worksheet and email the correct people

    week: The week number regarding the patching
    _file: The file in which to pull data from
    ticket: The change ticket associated with the patching
    current: the month and year for the change in MMM YYYY format
    sheet: The sheet to pull data from in the file
    """

    current_time = dt.now()
    current_hour = dt.strftime(current_time, '%H')
    hour = ''
    workbook = IO()
    try:
        # Times are in Mountain Standard Time
        if current_hour == '13':  # Eastern
            hour = 15  # sent at 13:00, start at 14:00
        elif current_hour == '14':  # Central
            hour = 16  # sent at 14:00 start at 15:00
        elif current_hour == '15':  # Mountain
            hour = 17  # sent at 15:00 start at 16:00
        elif current_hour == '16':  # Pacific
            hour = 18  # sent at 16:00 start at 17:00
        else:
            assert ValueError, f'Unknown sheet: The sheet ({sheet}) does not ' \
                               f'exist or is not found.'
            quit()

        # Get a list of lists for sending emails
        workbook = load_workbook(_file)
        worksheet = workbook[sheet]
        records = []
        for i in range(2, worksheet.max_row + 1):
            record = worksheet.cell(row=i, column=1)
            if not record:
                continue
            records.append(record.value)
        _notify(ticket, current, week, records, hour)
        file_log(f'Email sent for the {cfg.TIMEZONE} timezone plants. | '
                f'Change {cfg.CHANGE_TICKET} | Week {cfg.WEEK_NUMBER}')

    except FileNotFoundError:
        print('No file found to process.')
        sleep(5)
        quit()

    except PermissionError:
        print('This file is locked by another process.')
        sleep(5)
        quit()

    except KeyError:
        print('The fle has been corrupted')
        sleep(5)
        quit()

    except IOError:
        print('There was an error trying to access the file.')
        sleep(5)
        quit()

    finally:
        workbook.close()





if __name__ == '__main__':
    # If current date is within first 7 days of the month, subtract one month
    if dt.today() <= dt(int(dt.today().year), int(dt.today().month), cfg.DAYS):
        month = dt.date(dt.today() - relativedelta(months=1))
    else:
        month = dt.today()

    get_data(cfg.WEEK_NUMBER,
             cfg.FILE,
             cfg.CHANGE_TICKET,
             month.strftime('%b %Y'),
             cfg.TIMEZONE
             )
    print(f'\n{cfg.TIMEZONE} email sent')
