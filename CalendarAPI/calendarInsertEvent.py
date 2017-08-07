# -*- coding: utf-8 -*-
'''
Created on 2017/07/30
'''

from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
import pandas as pd
import ConfigParser
import logging


try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-insert-event.json
SCOPES = 'https://www.googleapis.com/auth/calendar'
APPLICATION_NAME = 'Google Calendar API Python Insert Events'


def setup_logger():
    """
    Setup Logger.
    """
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)

    log_file_path = os.path.join(os.path.abspath('..'), 'log')
    log_file_path = os.path.join(log_file_path, 'calendarInsertEvent.log')
    fh = logging.FileHandler(log_file_path)
    logger.addHandler(fh)

    sh = logging.StreamHandler()
    logger.addHandler(sh)

    formatter = logging.Formatter(
        '%(asctime)s:%(lineno)d:%(levelname)s:%(message)s')
    fh.setFormatter(formatter)
    sh.setFormatter(formatter)


logger = setup_logger()


def setup_config():
    """
    Setup Config.
    """
    # 設定ファイル読み込み
    conf_dir_path = os.path.join(os.path.abspath('..'), 'conf')
    conf_file_path = os.path.join(conf_dir_path, 'settings.ini')
    conf_file = ConfigParser.SafeConfigParser()
    conf_file.read(conf_file_path)

    # configの値はglobal変数に
    global CLIENT_SECRET_FILE_NAME
    global TMP_CREDENTIAL_FILE_NAME
    global USER_FILE_NAME
    global RESOURCE_FILE_NAME
    global EVENT_FILE_NAME

    # 設定ファイルから各ファイル名を取得
    CLIENT_SECRET_FILE_NAME = unicode(conf_file.get(
        'settings', 'CLIENT_SECRET_FILE_NAME'), 'UTF-8')
    USER_FILE_NAME = unicode(conf_file.get(
        'settings', 'USER_FILE_NAME'), 'UTF-8')
    EVENT_FILE_NAME = unicode(conf_file.get(
        'settings', 'EVENT_FILE_NAME'), 'UTF-8')
    RESOURCE_FILE_NAME = unicode(conf_file.get(
        'settings', 'RESOURCE_FILE_NAME'), 'UTF-8')
    TMP_CREDENTIAL_FILE_NAME = unicode(conf_file.get(
        'settings', 'TMP_CREDENTIAL_FILE_NAME'), 'UTF-8')


setup_config()


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, TMP_CREDENTIAL_FILE_NAME)

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:

        credential_path = os.path.join(os.path.abspath('..'), 'credentials')
        secret_file = os.path.join(credential_path, CLIENT_SECRET_FILE_NAME)

        flow = client.flow_from_clientsecrets(secret_file, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            logger.error('ImportError argparse')
#             credentials = tools.run(flow, store)
        logger.info('Storing credentials to ' + credential_path)
    return credentials


def get_calendar_id(resource_df, values):
    """
    Get Calendar ID
        values['施設']とresource_df[name]が同じ行の
        emailをcalendar_idとして返却
    """
    # values['施設']と同じresource_df['姓名']のメールアドレスを取得
    calendar_id = resource_df.ix[resource_df['name']
                                 == values['施設']]['email'].values[0:1][0]

    return calendar_id


def create_api_body(user_df, values):
    """
    create API body
    """
    # values['登録者']と同じuser_df['姓名']のメールアドレスを取得
    email = user_df.ix[user_df['姓名'] ==
                       values['登録者']]['メールアドレス'].values[0:1][0]

    # 開始時間
    start_time = values['開始日時'].replace('/', '-')
    start_time = start_time.replace('|', 'T')
    start_time = start_time + '+09:00'

    # 終了時間
    end_time = values['終了日時'].replace('/', '-')
    end_time = end_time.replace('|', 'T')
    end_time = end_time + '+09:00'

    body = {
        "summary": "",
        "start": {
            "dateTime": start_time,
            "timeZone": "Asia/Tokyo",
        },
        "end": {
            "dateTime": end_time,
            "timeZone": "Asia/Tokyo",
        },
        "attendees": [
            {"email": email},
        ],
    }

    return body


def main():
    """
    Creates a Google Calendar API service object and
    create events on the user's calendar.
    """
    logger.info('Start Create the Calendar events')

    # 認証情報取得しAPIクライアント生成
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    # CSVファイルディレクトリ
    csv_files_path = os.path.join(os.path.abspath('..'), 'csvfiles')

    # 施設CSV読み込み
    resource_csv_file = os.path.join(csv_files_path, RESOURCE_FILE_NAME)
    resource_df = pd.read_csv(resource_csv_file)

    # 登録者CSV読み込み
    user_csv_file = os.path.join(csv_files_path, USER_FILE_NAME)
    user_df = pd.read_csv(user_csv_file)
    user_df['姓名'] = user_df[['姓', '名']].apply(
        lambda x: '{}　{}'.format(x[0], x[1]), axis=1)

    # 予定CSV読み込み
    event_csv_file = os.path.join(csv_files_path, EVENT_FILE_NAME)
    event_df = pd.read_csv(event_csv_file)

    for i, values in event_df.iterrows():
        # 予定CSV件数繰り返し
        try:
            logger.info(
                "EVENT_CSV : LineNo." + i + values['施設'] + values['開始日時'] +
                values['終了日時'] + values['登録者'])

            # APIのbody作成
            body = create_api_body(user_df, values)

            # 予定登録対象施設決定
            calendar_id = get_calendar_id(resource_df, values)

            # API呼び出し
            event = service.events().insert(calendarId=calendar_id, body=body).\
                execute()

            # API呼び出し結果確認
            if not event.get('htmlLink'):
                logger.error('Failed create event.')
            else:
                logger.info("Event created : %s" % event.get('htmlLink'))

        except Exception as e:
            logger.exception(e)


if __name__ == '__main__':
    main()
