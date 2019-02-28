# -*-  encoding=utf-8   -*-

# 需求：
#       已有部门列表的excel文件和个人信息的excel文件，个人文件放在data文件夹
#       需要将data下的个人文件通过部门列表文件分类保存到相应的部门目录
# 处理：
#       由于部门列表只有姓名和部门，可能存在名字重复的情况，流程如下
#       1、先将全部个税信息文件拷贝到未匹配到部门文件夹
#       2、部门列表到个人文件进行逐行匹配，有以下几种情况：
#           匹配到，且文件存在，则移动文件到部门文件夹
#           匹配到，但文件不存在，则提示姓名重复
#           匹配不到，则写入未有个税文件的文档
#           剩下的是部门列表中没有的人，保存在'未匹配到部门的'文件夹

import os
import shutil
import xlrd


def main():
    my_path = os.getcwd()
    old_data_path = my_path + '\\' + 'data'
    # 读取old_data_path目录下的所有文件名并存在file_list里面
    file_list = os.listdir(old_data_path)
    xls_file = '部门列表.xls'
    xlsx_file = '部门列表.xlsx'
    department_file = ''
    data_path = my_path + '\\' + '未匹配到部门的'
    tag = 0

    # 判断未匹配到部门的文件夹是否存在，不存在则创建
    if not os.path.exists(data_path):
        os.makedirs(data_path)
    # 复制全部文件到未匹配到部门文件夹下
    for file_name in file_list:
        o_file = old_data_path + '\\' + file_name
        n_file = data_path + '\\' + file_name
        shutil.copyfile(o_file, n_file)

    # 判断部门列表文件类型，没有文件则退出程序
    if os.path.isfile('部门列表.xls'):
        department_file = xls_file
    elif os.path.isfile('部门列表.xlsx'):
        department_file = xlsx_file
    else:
        print('未找到部门列表文件!')
        os.system('pause')
        exit()

    # 文件列表增加换行符
    with open('list.txt', 'w', encoding='utf-8') as xls_list:
        for xls in file_list:
            xls_list.write(xls + '\n')
        xls_list.close()

    with xlrd.open_workbook(filename=department_file) as department_file:
        # 读取xls文件的第一张表
        sheet1 = department_file.sheet_by_index(0)
        # 读取第一和第二列
        # cols_1 = sheet1.col_values(0)
        # cols_2 = sheet1.col_values(1)
        # 读取部门列表，开始匹配
        for rows_number in range(0, sheet1.nrows):
            line_data = sheet1.row_values(rows_number)
            name = line_data[0]
            department_dir_file = line_data[1]
            department_dir = my_path + '\\分部门已匹配的\\' + department_dir_file
            # 判断部门文件夹是否存在
            if not os.path.exists(department_dir):
                os.makedirs(department_dir)
            else:
                pass
            # 匹配个税文件列表list.txt文件
            with open('list.txt', encoding='utf-8') as list_file:
                for line_list in list_file:
                    old_file = data_path + '\\' + line_list[:-1]
                    new_file = department_dir + '\\' + line_list[:-1]
                    if name in line_list:
                        if os.path.isfile(old_file):
                            # 如果名字在个税文件目录中且源文件夹中有个税文件。
                            shutil.move(old_file, new_file)
                            print('已将' + name + '归类到:' + department_dir_file)
                            tag = 1
                            break
                        else:
                            # 部门列表中有名字重复的会将个税文件保存在部门列表文件中第一个出现的部门，
                            # 由于第一次匹配已经将文件移动到其他文件夹，所以同名的名字在第一次以后会
                            # 在‘未匹配’文件夹中找不到，此时程序将姓名保存到‘名字有重复.txt’中。
                            print('注意名字重复：' + name + '的重复数据')
                            with open('名字有重复的.txt', 'a', encoding='utf-8') as one_more_list:
                                one_more_list.write(name + '\n')
                                one_more_list.close()
                            tag = 1
                    else:
                        pass
                # 如果遍历了个税文件列表没有找到部门列表中的人，则将姓名保存到'未找到个人文件的.txt'
                # 文件中。
                if tag == 0:
                    with open('未找到个人文件的.txt', 'a', encoding='utf-8') as xls_list:
                        print('没有' + name + '的个税文件。')
                        word = name + ',' + department_dir_file
                        xls_list.write(word + '\n')
                        xls_list.close()
                else:
                    tag = 0
                list_file.close()
    # 根据部门列表文件夹匹配完没有移动到所属部门的个税文件就是未匹配到部门的
    os.remove(my_path + '\\list.txt')


if __name__ == '__main__':
    main()

