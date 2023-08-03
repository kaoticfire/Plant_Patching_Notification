""" Email when all clear with plant patching. """
from openpyxl import load_workbook
from datetime import datetime as dt
from dateutil.relativedelta import relativedelta
from time import sleep
from config import Config as cfg
from mail import __mail_notification
from filelog import file_log


def __all_clear(change_ticket: str, date_stamp: str,
                week_num: int, site: list) -> None:
    if cfg.DEBUG:
        to = 'virgil.hoover@cavalry.solutions'
    else:
        to = 'HistorianSupport@calpine.com; ' \
             'brad.gibson@calpine.com; ' \
             'David.Symons@calpine.com; ' \
             'Laura.Morrical@calpine.com; ' \
             'overwatch@cavalry.solutions'

    if week_num:
        subj = f'{change_ticket} - Plant Prod Monthly Patching - ' \
                          f'{date_stamp} - Week {week_num}'
    else:
        subj = f'{change_ticket} - Plant Prod Monthly Patching - ' \
                         f'{date_stamp}'
    content= f'<html><body><p>Overwatch,<br />All patching for the ' \
             f'plant sites has been completed. Please resume normal' \
             f' monitoring for all involved sites.<br /><br />' \
             f'{"<br />".join(map(str, site))}<br /><br />' \
             f'Virgil Hoover<br />SUPPORT@CAVALRY.SOLUTIONS<br />' \
             f'NETWORK OPERATIONS CENTER ///<br /><a href="tel:+1 ' \
             f'(720) 279-2233">+1 (720) 279-2233</a> - 24X7 ' \
             f'OVERWATCH</p></body></html>'
    __mail_notification(to, subj, content)


def get_data(week: int, file: str, ticket: str, current: str) -> None:
    """ Get data from the Excel worksheet and email the correct people

    week: The week number regarding the patching
    file: The file in which to pull data from
    ticket: The change ticket associated with the patching
    current: the month and year for the change in MMM YYYY format
    """

    workbook = ''
    sites = []
    try:
        # Get a list of lists for sending emails
        workbook = load_workbook(file)
        sheet_names = workbook.sheetnames
        records = []
        listed_records = ''
        for name in sheet_names:
            worksheet = workbook[name]
            for i in range(2, worksheet.max_row + 1):
                for j in range(1, worksheet.max_column + 1):
                    record = worksheet.cell(row=i, column=j)
                    records.append(record.value)

            temp = filter(lambda item: item is not None, records)
            listed_records = list(temp)
        for _ in range(0, len(listed_records), 3):
            file_log(f'Completion email sent for the '
                     f'{listed_records[_:_ + 3][0]}. | Change '
                     f'{cfg.CHANGE_TICKET} | Week {cfg.WEEK_NUMBER}')
            sites.append(listed_records[_:_ + 3][0])
        sleep(15)
        __all_clear(ticket, current, week, sites)

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
        print('The was an error trying to access the file or a portion of it.')
        sleep(5)
        quit()

    finally:
        workbook.close()



if __name__ == '__main__':
    # If current date is within first 13 days of the month, subtract one month
    if dt.today() <= dt(int(dt.today().year), int(dt.today().month), cfg.DAYS):
        month = dt.date(dt.today() - relativedelta(months=1))
    else:
        month = dt.today()

    get_data(cfg.WEEK_NUMBER,
             cfg.FILE,
             cfg.CHANGE_TICKET,
             month.strftime('%b %Y')
             )
