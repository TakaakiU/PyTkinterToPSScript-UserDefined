import datetime


#   datetime設定
TIMEDELTA_SET = datetime.timedelta(hours=9)
TIMEZONE_JST = datetime.timezone(TIMEDELTA_SET, 'JST')


class ctrlString():
    def now_filename():
        datetime_now = datetime.datetime.now(TIMEZONE_JST)
        return datetime_now.strftime('%Y%m%d-%H%M%S')

    def now_label():
        datetime_now = datetime.datetime.now(TIMEZONE_JST)
        return datetime_now.strftime('[%Y年%m月%d日 %H:%M:%S]')
