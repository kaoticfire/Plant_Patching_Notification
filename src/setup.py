""" A script to set the Logger initially. """
from time import sleep
from os import getenv


def find_n_replace(_file: str) -> list:
    """ A script to Update the On-Call Listings.

    :param _file: the file to search and to write back to

    :returns: a list containing the exit code and a message if needed
    """
    try:
        # Read the file
        with open(_file, 'r') as fr:
            temp_data = fr.read()

        # Search for text and replace it with some new text
        if not DEBUG:
            temp_data = temp_data.replace('<user>', getenv('Username'))

    except FileNotFoundError:
        msg = '- The file was not found, please locate the file and try again.'
        return [1, msg]

    except IOError:
        msg = '- An issue with accessing the file, please try again later.'
        return [2, msg]

    # Overwrite the file
    if not DEBUG:
        with open(_file, 'w') as fw:
            fw.write(temp_data)

    return [None, None]


if __name__ == '__main__':
    DEBUG = False
    folder = 'C:/Users/getenv('Username')/OneDrive/'
    master_file = 'config.yaml'
    exit_code, message = find_n_replace(master_file)
    if exit_code:
        print(exit_code, message)
        sleep(5)
    else:
        exit(0)
