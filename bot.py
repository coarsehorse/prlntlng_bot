import requests
import json
import time
import datetime
import os
import subprocess

CREDENTIALS_FILE = 'credentials.json'
TELEGRAM_API_URL = 'https://api.telegram.org/'
TELEGRAM_FILE_API_URL = 'https://api.telegram.org/file/'
LAST_PROC_MESS_FILE = 'last_processed_message.txt'
LOG_FILE = 'log.txt'
ALLOWED_MIME_TYPES = [
    'application/pdf',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'application/msword',
    'text/plain']
FILES_DIR = 'Files'


def read_credentials():
    """
    Read access token from CREDENTIALS_FILE.

    :returns: credentials json'
    """
    with open(CREDENTIALS_FILE, encoding='utf-8', mode='r') as f_content:
        return json.load(f_content)


def get_updates(access_token, offset=0):
    """
    Get new messages for bot from Telegram API using get_updates method.

    :param access_token: string that contains access token in format
        'bot_num:token'
    :param offset: the last last update id + 1
    :returns: decoded response json.
    """
    resp = requests.get(TELEGRAM_API_URL + 'bot' +
                        access_token + '/getUpdates?offset=' + str(offset))
    return resp


def send_message(access_token, chat_id, text):
    """
    Send message to specified chat.

    :param access_token: string with bot access token
    :param chat_id: integer with chat id
    :param text: string with message to be sent
    """
    requests.get(TELEGRAM_API_URL + 'bot' + access_token +
                 '/sendMessage?' + 'chat_id=' + str(chat_id) +
                 '&text=' + str(text))


def download_file(access_token, file_id):
    """
    Download file by file_id and save to FILES_DIR.
    FILES_DIR will be automatically created if not exists.

    :param access_token: string with bot access token
    :param file_id: integer with id of file to be downloaded
    :returns: string with downloaded file path or None on error.
    """
    resp = requests.get(TELEGRAM_API_URL + 'bot' + access_token +
                        '/getFile?file_id=' + str(file_id))
    if resp.ok:
        file_path = resp.json()['result']['file_path']
        download_link = TELEGRAM_FILE_API_URL + 'bot' + access_token \
            + '/' + file_path
        # Create folder if not exists
        if not os.path.exists(FILES_DIR):
            os.makedirs(FILES_DIR)
        # Open connection for file downloading
        resp = requests.get(download_link, allow_redirects=True)
        if resp.ok:
            # Get file name from url
            file_name = file_path
            if '/' in file_path:
                file_name = file_path.split('/')[-1]
            # Avoid duplicates
            ms_now = int(round(time.time() * 1000))
            if '.' in file_name:
                before_ext = '.'.join(file_name.split('.')[0:-1])
                ext = file_name.split('.')[-1]
                file_name = before_ext + '_' + str(ms_now) + '.' + ext
            else:
                file_name = file_name + '_' + ms_now
            # Get future file path
            file_path = os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                FILES_DIR, file_name)
            # Download file
            with open(file_path, mode='wb') as file:
                file.write(resp.content)
            return file_path
        else:
            return None
    else:
        return None

# Start


credentials = read_credentials()
access_token = credentials['access_token']
default_printer = credentials['default_printer']
last_processed_message = 0
content = ''
if not os.path.exists(LAST_PROC_MESS_FILE):
    with open(LAST_PROC_MESS_FILE, encoding='utf-8', mode='w'):
        pass  # just create a file
else:
    with open(LAST_PROC_MESS_FILE, encoding='utf-8', mode='r') as f_content:
        content = f_content.read().strip()
if content:  # if not empty
    last_processed_message = int(content)

print('lpm: ' + str(last_processed_message), flush=True)

# Long polling for updates
print("Bot started. Listening new messages...", flush=True)
while True:
    resp = get_updates(access_token, last_processed_message + 1)
    if resp.ok:
        for upd in resp.json(encoding='utf-8')['result']:
            # Read data from message
            update_id = upd['update_id']
            user_id = upd['message']['from']['id']
            user_name = upd['message']['from']['first_name']
            user_nickname = '@' + upd['message']['from']['username']
            message_text = ''
            if 'text' in upd['message']:
                message_text = upd['message']['text']
            mime_type = ''
            file_id = ''
            if 'document' in upd['message']:
                if 'mime_type' in upd['message']['document']:
                    mime_type = upd['message']['document']['mime_type']
                if 'file_id' in upd['message']['document']:
                    file_id = upd['message']['document']['file_id']
            chat_id = upd['message']['chat']['id']
            # Log message
            log_msg = str(datetime.datetime.now()) + "\t" \
                + "User '" + user_name + "'(" + user_nickname + ", " \
                + "id=" + str(user_id) + ")" + "\t" \
                + "message: '" + message_text + "'"
            with open(LOG_FILE, encoding='utf-8', mode='a') as log_f:
                log_f.write(log_msg + "\n")
            # Process request
            if message_text:  # if message exists
                send_message(
                    access_token, chat_id,
                    "Привет!\n" +
                    "Этот бот может напечатать документ на принтаре.\n" +
                    "Для того чтобы начать печать, отправь " +
                    "документ сюда.")
            elif mime_type in ALLOWED_MIME_TYPES:
                # process
                print("'" + str(mime_type) + "'", flush=True)
                file_path = download_file(access_token, file_id)
                if file_path:
                    print(file_path, flush=True)
                    # Print document
                    args = '"C:\\\\Program Files\\\\gs\\\\bin\\\\gswin32c" ' \
                           '-sDEVICE=mswinpr2 ' \
                           '-dBATCH ' \
                           '-dNOPAUSE ' \
                           '-dFitPage ' \
                           '-sOutputFile="%printer%' + default_printer + '" '
                    ghostscript = args \
                        + os.path.join(os.getcwd(), 'myFile.pdf') \
                            .replace('\\', '\\\\')
                    subprocess.call(ghostscript, shell=True)
                    print('Printed successfully', flush=True)
                else:
                    log_msg = str(datetime.datetime.now()) + "\t" \
                        + 'Error during downloading file_id: ' \
                        + str(file_id) \
                        + ' Sent by user ' + str(user_nickname)
                    send_message(access_token, chat_id, log_msg)
                    with open(LOG_FILE, encoding='utf-8', mode='a') as log_f:
                            log_f.write(log_msg + "\n")
            else:
                send_message(
                    access_token, chat_id,
                    "Не поддерживаемый тип документа!\n" +
                    "Я умею работать только с" +
                    " .doc .docx .pdf .txt")
            # Update last processed message id
            last_processed_message = update_id
        # Save last processed message id
        with open(LAST_PROC_MESS_FILE, encoding='utf-8', mode='w') as lpm_f:
            lpm_f.write(str(last_processed_message))
    else:
        log_msg = str(datetime.datetime.now()) + "\t" \
            + 'Error, status code: ' + str(resp.status_code) \
            + ', responce text: ' + str(resp.text)
        with open(LOG_FILE, encoding='utf-8', mode='a') as log_f:
                log_f.write(log_msg + "\n")
