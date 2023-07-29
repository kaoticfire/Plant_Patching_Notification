""" The main notification function """
from win32com.client import Dispatch
from time import sleep


def __mail_notification(recipients: str, subject: str, body: str) -> None:
    """ Send email with Attachment

    recipients: who is receiving the email
    subject: the reason for the email
    body: what you want the email to see
    """

    outlook = Dispatch('outlook.application')
    message = outlook.CreateItem(0)
    message.to = recipients
    message.Subject = subject
    message.HTMLBody = body
    try:
        message.Send()

    except IOError:
        print('Error: Problem sending email')
        sleep(5)
        quit()