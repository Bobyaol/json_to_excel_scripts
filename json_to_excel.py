# 这个脚本作用是将.json文件转换为excel文件
# 因为该json文件, 存在多级嵌套dict情况, 因此需要将json文件解析到excel的多个sheet中
# 一个一级sheet: release信息
# 三个二级sheet: repos, release_trend, mark_line_data信息
# 两个三级sheet: mark_line_data中的commits信息, repos中的datas信息

import json
import xlsxwriter
import sys
import re

file_path = sys.argv[1]
key_1_1   = 'release'
key_2_0   = 'total'
key_2_1   = 'repos'
key_2_2   = 'release_trend'
key_2_3   = 'mark_line_data'
key_3_1   = 'commits'
key_3_2   = 'datas'



def write_data_for_key_and_worksheet(dict_for_key, worksheet):
    # 表头
    row = 0
    col = 0
    for index, i in enumerate(dict_for_key.items()):
        if index == 0:
            single_row = i[1].keys()
            single_row = list(single_row)
            length = len(single_row)
            for i_2 in range(length):
                worksheet.write (row, col+i_2, str(single_row[i_2]))
            row += 1
    
    # 数据
    row = 1
    col = 0
    for index, i in enumerate(dict_for_key.items()):
        single_row = i[1].values()
        single_row = list(single_row)
        length = len(single_row)
        for i_2 in range(length):
            worksheet.write (row, col+i_2, str(single_row[i_2]))
        row += 1



with open(file_path, encoding='utf-8') as f:
    total_content = (json.load(f))[0]
    print("\n")
    keys = list(total_content.keys())
    print(f"keys = {keys}")
    print(f"开始处理{file_path}")
    filename = f'{file_path}.xlsx'
    workbook = xlsxwriter.Workbook(filename)



#1_1#############################################################################################################

    print(f"开始处理{key_1_1}")
    sheet_name_1_1 = key_1_1
    worksheet_1_1 = workbook.add_worksheet(sheet_name_1_1)
    worksheet_1_1.set_column_pixels(0, 20, 200)

    total_content_1_1 = total_content.copy()
    total_content_1_1.pop(key_2_0)
    total_content_1_1.pop(key_2_1)
    total_content_1_1.pop(key_2_2)
    total_content_1_1.pop(key_2_3)

    # print(f"total_content_1_1 = {total_content_1_1}")


    dict_for_key_1_1 = {}
    dict_for_key_1_1[0] = total_content_1_1

    
    write_data_for_key_and_worksheet(dict_for_key_1_1, worksheet_1_1)


#2_0##############################################################################################################

    print(f"开始处理{key_2_0}")
    sheet_name_2_0 = key_2_0
    worksheet_2_0 = workbook.add_worksheet(sheet_name_2_0)
    worksheet_2_0.set_column_pixels(0, 20, 200)


    second_half_dict_2_0 = total_content[key_2_0]

    total_content_2_0 = total_content.copy()
    total_content_2_0.pop(key_2_0)
    total_content_2_0.pop(key_2_1)
    total_content_2_0.pop(key_2_2)
    total_content_2_0.pop(key_2_3)
    first_half_dict_2_0 = total_content_2_0


    dict_for_key_2_0 = {}
    index_2_0 = 0
    for i_2_0 in [second_half_dict_2_0]:
        dict_for_key_2_0[index_2_0] = {**first_half_dict_2_0, **{"*":"*"}, **i_2_0}
        index_2_0 += 1
    
    # print(f"dict_for_key_2_0 = {dict_for_key_2_0}")

    write_data_for_key_and_worksheet(dict_for_key_2_0, worksheet_2_0)


#2_1##############################################################################################################

    print(f"开始处理{key_2_1}")
    sheet_name_2_1 = key_2_1
    worksheet_2_1 = workbook.add_worksheet(sheet_name_2_1)
    worksheet_2_1.set_column_pixels(0, 20, 200)


    second_half_dict_2_1 = total_content[key_2_1]

    total_content_2_1 = total_content.copy()
    total_content_2_1.pop(key_2_0)
    total_content_2_1.pop(key_2_1)
    total_content_2_1.pop(key_2_2)
    total_content_2_1.pop(key_2_3)
    first_half_dict_2_1 = total_content_2_1


    dict_for_key_2_1 = {}
    index_2_1 = 0
    for i_2_1 in second_half_dict_2_1:
        dict_for_key_2_1[index_2_1] = {**first_half_dict_2_1, **{"*":"*"}, **i_2_1}
        index_2_1 += 1
    
    # print(f"dict_for_key_2_1 = {dict_for_key_2_1}")

    write_data_for_key_and_worksheet(dict_for_key_2_1, worksheet_2_1)


#2_2##############################################################################################################

    print(f"开始处理{key_2_2}")
    sheet_name_2_2 = key_2_2
    worksheet_2_2 = workbook.add_worksheet(sheet_name_2_2)
    worksheet_2_2.set_column_pixels(0, 20, 200)


    second_half_dict_2_2 = total_content[key_2_2]

    total_content_2_2 = total_content.copy()
    total_content_2_2.pop(key_2_0)
    total_content_2_2.pop(key_2_1)
    total_content_2_2.pop(key_2_2)
    total_content_2_2.pop(key_2_3)
    first_half_dict_2_2 = total_content_2_2


    dict_for_key_2_2 = {}
    index_2_2 = 0
    for i_2_2 in second_half_dict_2_2:
        dict_for_key_2_2[index_2_2] = {**first_half_dict_2_2, **{"*":"*"}, **i_2_2}
        index_2_2 += 1
    
    # print(f"dict_for_key_2_2 = {dict_for_key_2_2}")

    write_data_for_key_and_worksheet(dict_for_key_2_2, worksheet_2_2)



#2_3##############################################################################################################

    print(f"开始处理{key_2_3}")
    sheet_name_2_3 = key_2_3
    worksheet_2_3 = workbook.add_worksheet(sheet_name_2_3)
    worksheet_2_3.set_column_pixels(0, 20, 200)


    second_half_dict_2_3 = total_content[key_2_3]

    total_content_2_3 = total_content.copy()
    total_content_2_3.pop(key_2_0)
    total_content_2_3.pop(key_2_1)
    total_content_2_3.pop(key_2_2)
    total_content_2_3.pop(key_2_3)
    first_half_dict_2_3 = total_content_2_3


    dict_for_key_2_3 = {}
    index_2_3 = 0
    for i_2_3 in second_half_dict_2_3:
        dict_for_key_2_3[index_2_3] = {**first_half_dict_2_3, **{"*":"*"}, **i_2_3}
        index_2_3 += 1
    
    # print(f"dict_for_key_2_3 = {dict_for_key_2_3}")

    write_data_for_key_and_worksheet(dict_for_key_2_3, worksheet_2_3)





#3_1##############################################################################################################

    print(f"开始处理{key_3_1}")
    sheet_name_3_1 = key_3_1
    worksheet_3_1 = workbook.add_worksheet(sheet_name_3_1)
    worksheet_3_1.set_column_pixels(0, 20, 200)



    second_half_dict_3_1 = total_content[key_2_2]

    total_content_3_1 = total_content.copy()
    total_content_3_1.pop(key_2_0)
    total_content_3_1.pop(key_2_1)
    total_content_3_1.pop(key_2_2)
    total_content_3_1.pop(key_2_3)
    first_half_dict_3_1 = total_content_3_1


    dict_for_key_3_1 = {}
    index_3_1 = 0
    for i in total_content[key_2_2]:
        if key_3_1 in i.keys():
            for commit in i[key_3_1]:
                # print(f"commit = {commit}")
                if 'title' in commit.keys():
                    title = commit['title']
                    title_key = re.findall(r'\[(.*?)\]',title)
                    if len(title_key) != 0:
                        title_key = title_key[0]
                dict_for_key_3_1[index_3_1] = {**first_half_dict_3_1, **{"*":"*"}, **i, **{"$":"$"}, **commit, **{"title_key": title_key}}
                index_3_1 += 1



    write_data_for_key_and_worksheet(dict_for_key_3_1, worksheet_3_1)


#3_2##############################################################################################################

    print(f"开始处理{key_3_2}")
    sheet_name_3_2 = key_3_2
    worksheet_3_2 = workbook.add_worksheet(sheet_name_3_2)
    worksheet_3_2.set_column_pixels(0, 20, 200)



    second_half_dict_3_2 = total_content[key_2_1]

    total_content_3_2 = total_content.copy()
    total_content_3_2.pop(key_2_0)
    total_content_3_2.pop(key_2_1)
    total_content_3_2.pop(key_2_2)
    total_content_3_2.pop(key_2_3)
    first_half_dict_3_2 = total_content_3_2


    dict_for_key_3_2 = {}
    index_3_2 = 0
    for i in total_content[key_2_1]:
        if key_3_2 in i.keys():
            for data in i[key_3_2]:
                data['data-date'] = data['date']
                data['data-project_id'] = data['project_id']
                data['data-commit_num'] = data['commit_num']
                data['data-dev_equivalent'] = data['dev_equivalent']
                data['data-loc'] = data['loc']
                data['data-developers'] = data['developers']

                data.pop('date')
                data.pop('project_id')
                data.pop('commit_num')
                data.pop('dev_equivalent')
                data.pop('loc')
                data.pop('developers')

                # print(f"commit = {commit}")
                # print(f"i = {i}")
                # print(f"data = {data}")
                dict_for_key_3_2[index_3_2] = {**first_half_dict_3_2, **{"*":"*"}, **i, **{"$":"$"}, **data}
                index_3_2 += 1




    write_data_for_key_and_worksheet(dict_for_key_3_2, worksheet_3_2)





    workbook.close()



