import os
import time
import prolog
import tools
from utils.log.index import error_log, input_log, info_log


def get_file(path_dir):
    # 给定一个目录和，拿到目录下的所有文件
    # path_dir = r'C:\\Users\\Monster\\Desktop\\9月资金调拨单'
    external_list = ['.xlsx', '.xlsm']
    file_list = tools.filter_by_condition(tools.get_file_list(path_dir), external_list)
    return file_list


def read_content(file_list):
    # 读取文件
    exclude_list = ['组织编码归属']
    content = []
    for file in file_list:
        tools.read_sheet(file, exclude_list, content)
    return content


def main():
    while True:
        path_dir = input_log('请输入文件夹路径：')
        start_time = time.time()
        try:
            file_list = get_file(tools.transform_path(path_dir))
        except FileNotFoundError:
            error_log("哎呀,小主是不是犯糊涂了呀,没有这个文件夹呀，快去确认一下下")
        else:
            if len(file_list) <= 0:
                error_log('小主,对不起,我没能找到您要的文件')
                continue
            info_log('--------小主,我已经拿到文件列表了--------')
            content = read_content(file_list)
            info_log('--------小主,Monster准备写入数据了哦--------')

            filename = '汇总.xlsx'

            current_filename = os.path.join(os.getcwd(), filename)
            if os.path.exists(current_filename):
                error_log('小主,文件已存在')
                is_cover = input_log('小主,要不要覆盖? [Y]覆盖：')
                if is_cover.upper() != 'Y':
                    while True:
                        new_filename = input_log('请输入新的文件名字(不需要包含扩展名)：')
                        filename = new_filename + '.xlsx'
                        current_filename = os.path.join(os.getcwd(), filename)
                        if os.path.exists(current_filename):
                            error_log('小主,文件已存在,请重新输入')
                            continue
                        else:
                            break
            tools.write_sheet(filename, content)
            # # todo:开始拆分
            # 准备拆分调播单
            end_time = time.time()
            spend_time = '耗时:' + str(end_time - start_time) + '秒'
            info_log(spend_time)
            info_log('---------小主,我的工作已完成,快夸夸我真棒---------')
            is_continue = input_log('是否继续？  [N]关闭：')
            if is_continue.upper() == 'N':
                break
            else:
                prolog.print_log()


if __name__ == '__main__':
    prolog.print_log()
    main()



