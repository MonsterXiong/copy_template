
from rich.console import Console
console = Console()


def error_log(content):
    print()
    Console(style='red').print(content)
    print()


def info_log(content):
    print()
    Console(style='magenta').print(content)
    print()


def input_log(question):
    print()
    answer = None
    is_empty = True
    while is_empty is True:
        answer = Console(style='blue').input(question)
        if len(answer) == 0:
            error_log('请输入内容')
        else:
            is_empty = False
    print()
    return answer