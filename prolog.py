import datetime
import os
import time
from rich import print as rprint
from rich.panel import Panel


def salary_log():
    salary_content = '小主,今天发工资，真开心'
    current_day = int(time.strftime('%d'))
    today = datetime.datetime.today()
    year, month, day = today.year, today.month, today.day
    if current_day != 20:
        current_date = datetime.date(year, month, day)
        next_month = month
        next_year = year
        if current_day > 20:
            next_month = month + 1
            if month == 12:
                next_month = 1
                next_year = year + 1
        next_salary_date = datetime.date(next_year, next_month, 20)
        differ_day = str(next_salary_date.__sub__(current_date).days)
        salary_content = '小主,距离下次工资发放还有' + differ_day + '天'
    return salary_content


def every_day_log():
    return '小主,好久不见,甚是想念'


def current_time_log():
    current_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    return '当前时间为:' + current_time


def notice_day_log():
    today = datetime.datetime.today()
    year, month, day = today.year, today.month, today.day
    weekday = datetime.date(year, month, day).strftime("%A")
    notice_day = ""
    match weekday:
        case 'Monday':
            notice_day = '小主,今天又是元气满满的一天，加油~~~'
        case 'Tuesday':
            notice_day = '小主,今天周二了啦,距离周末还剩三天,加油~~~'
        case 'Wednesday':
            notice_day = '小主,今天周三了啦,距离周末还二天,加油~~~'
        case 'Thursday':
            notice_day = '小主,今天周四了啦,距离周末还一天,加油~~~'
        case 'Friday':
            notice_day = '小主,这周过得好快呀，今天就到周五了,即将迎来两天的小假期哟~~~好激动，嘻嘻:smiley:'
        case 'Saturday':
            notice_day = '小主,今天是周末的第一天,要好好休息,去玩吧~~~'
        case 'Sunday':
            notice_day = '小主,今天休息啦~~~让我也休息休息嘛'
    return notice_day


def format_prolog_content(*content_list):
    content = '\n'
    for content_index, content_item in enumerate(content_list):
        if content_index == len(content_list)-1:
            content += content_item + '\n'
        else:
            content += content_item + '\n\n'
    return content


# 打印控制台
def print_log():
    os.system('cls')
    content = format_prolog_content(current_time_log(), every_day_log(), salary_log(), notice_day_log())
    title = '欢迎使用Monster小工具'
    border_style = 'magenta'
    rprint(Panel.fit("[bold red]"+content, border_style=border_style, title=title))
    print()
