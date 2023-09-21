""" Default logging for the project."""
from logging import getLogger
from logging.config import dictConfig
from yaml import safe_load
from os import getenv


def file_log(message: str) -> None:
    """ Keep record of outbound emails. """
    with open('config.yaml', 'r') as file:
        dictConfig(safe_load(file.read()))
    logger = getLogger(getenv('UserName'))
    logger.info(message)


if __name__ == '__main__':
    msg = input('Enter your message: ')
    file_log(msg)
