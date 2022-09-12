# 这个脚本作用是将.json文件转换为excel文件
# 因为该json文件, 存在多级嵌套dict情况, 因此需要将json文件解析到excel的多个sheet中
# 一个一级sheet: primary_email信息
# 三个二级sheet: total, metrics, baselines信息
# 本脚本文件是对更新格式之后的contibutors.json的处理

import json
import xlsxwriter
import sys
import re

file_path = sys.argv[1]
key_1_1   = 'primary_email'
key_2_0   = 'total'
key_2_1   = 'metrics'
key_2_2   = 'baselines'




def write_data_for_key_and_worksheet(dict_for_key, worksheet):
    # 表头
    row = 0
    col = 0
    for index, i in enumerate(dict_for_key.items()):
        if index == 0:
            single_row = i[1].keys()
            # print(f"i[1] = {i[1]}")
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
    total_content = json.load(f) # total_content is a list
    total_content_list = total_content.copy()
    total_content = {}
    # print("\n")
    index_0_0 = 0
    for i in total_content_list:
        total_content[index_0_0] = i
        index_0_0 += 1
        # print(f"len(total_content) = {len(total_content)}")
    keys = list(total_content.keys())
    # print(f"keys = {keys}")
    print(f"开始处理{file_path}")
    filename = f'{file_path}.xlsx'
    workbook = xlsxwriter.Workbook(filename)



#1_1#############################################################################################################

    print(f"开始处理{key_1_1}")
    sheet_name_1_1 = key_1_1
    worksheet_1_1 = workbook.add_worksheet(sheet_name_1_1)
    worksheet_1_1.set_column_pixels(0, 20, 200)

    total_content_1_1 = {}
    for index_1_1, i_1_1 in enumerate(total_content.items()):
        total_content_1_1[index_1_1] = {}
        total_content_1_1[index_1_1]['primary_email'] = i_1_1[1]['primary_email']
        total_content_1_1[index_1_1]['linked_emails'] = i_1_1[1]['linked_emails']
        total_content_1_1[index_1_1]['user_id'] = i_1_1[1]['user_id']
        total_content_1_1[index_1_1]['user_name'] = i_1_1[1]['user_name']
    dict_for_key_1_1 = total_content_1_1

    write_data_for_key_and_worksheet(dict_for_key_1_1, worksheet_1_1)


# #2_0##############################################################################################################

    print(f"开始处理{key_2_0}")
    sheet_name_2_0 = key_2_0
    worksheet_2_0 = workbook.add_worksheet(sheet_name_2_0)
    worksheet_2_0.set_column_pixels(0, 20, 200)


    first_half_dict_2_0 = {}
    for index_2_0, i_2_0 in enumerate(total_content.items()):
        first_half_dict_2_0[i_2_0[1]['primary_email']] = {}
        first_half_dict_2_0[i_2_0[1]['primary_email']]['primary_email'] = i_2_0[1]['primary_email']
        first_half_dict_2_0[i_2_0[1]['primary_email']]['linked_emails'] = i_2_0[1]['linked_emails']
        first_half_dict_2_0[i_2_0[1]['primary_email']]['user_id'] = i_2_0[1]['user_id']
        first_half_dict_2_0[i_2_0[1]['primary_email']]['user_name'] = i_2_0[1]['user_name']


    # print(f"first_half_dict_2_0 = {first_half_dict_2_0}")


    second_half_dict_2_0 = {}
    for index_2_0_, i_2_0_ in enumerate(total_content.items()):
        second_half_dict_2_0[index_2_0_] = i_2_0_[1][key_2_0]
        second_half_dict_2_0[index_2_0_]['primary_email'] = i_2_0_[1]['primary_email']

        # print(f"second_half_dict_2_0 = {second_half_dict_2_0}")


    dict_for_key_2_0 = {}
    index_2_0 = 0
    for i_2_0 in second_half_dict_2_0.items():
        # print(f"i_2_0 = {i_2_0}")
        # print(f"first_half_dict_2_0[i_2_0[1] = {first_half_dict_2_0[i_2_0[1]]}")
        first_dict = first_half_dict_2_0[i_2_0[1]['primary_email']]
        # print(f"first_dict = {first_dict}")
        dict_for_key_2_0[index_2_0] = {**first_dict, **{"*":"*"}, **i_2_0[1]}
        index_2_0 += 1


    write_data_for_key_and_worksheet(dict_for_key_2_0, worksheet_2_0)


# #2_1##############################################################################################################

    print(f"开始处理{key_2_1}")
    sheet_name_2_1 = key_2_1
    worksheet_2_1 = workbook.add_worksheet(sheet_name_2_1)
    worksheet_2_1.set_column_pixels(0, 20, 200)


    first_half_dict_2_1 = {}
    for index_2_1, i_2_1 in enumerate(total_content.items()):
        first_half_dict_2_1[i_2_1[1]['primary_email']] = {}
        first_half_dict_2_1[i_2_1[1]['primary_email']]['primary_email'] = i_2_1[1]['primary_email']
        first_half_dict_2_1[i_2_1[1]['primary_email']]['linked_emails'] = i_2_1[1]['linked_emails']
        first_half_dict_2_1[i_2_1[1]['primary_email']]['user_id'] = i_2_1[1]['user_id']
        first_half_dict_2_1[i_2_1[1]['primary_email']]['user_name'] = i_2_1[1]['user_name']

    # print(f"first_half_dict_2_1 = {first_half_dict_2_1}")


    second_half_dict_2_1 = {}
    metric_index = 0
    for index_2_1_, i_2_1_ in enumerate(total_content.items()):

        # print("\n")
        # print(f"i_2_1_ = {i_2_1_}")
        metrics = i_2_1_[1]['metrics']
        primary_email = i_2_1_[1]['primary_email']
        # linked_emails = i_2_1_[1]['linked_emails']
        # user_id = i_2_1_[1]['user_id']

        # print("\n")
        # print(f"metrics = {metrics}")

        for metric in metrics:
            second_half_dict_2_1[metric_index] = {}
            second_half_dict_2_1[metric_index]['primary_email'] = primary_email
            # second_half_dict_2_1[metric_index]['linked_emails'] = linked_emails
            # second_half_dict_2_1[metric_index]['user_id'] = user_id
            second_half_dict_2_1[metric_index]['date'] = metric['date']
            if 'commit_num' in metric.keys():
                second_half_dict_2_1[metric_index]['commit_num'] = metric['commit_num']
            else:
                second_half_dict_2_1[metric_index]['commit_num'] = 0

            if 'commit_num' in metric.keys():
                second_half_dict_2_1[metric_index]['dev_equivalent'] = metric['dev_equivalent']
            else:
                second_half_dict_2_1[metric_index]['dev_equivalent'] = 0

            if 'loc' in metric.keys():
                second_half_dict_2_1[metric_index]['loc'] = metric['loc']
            else:
                second_half_dict_2_1[metric_index]['loc'] = 0

            if 'commit_num_trend' in metric.keys():
                second_half_dict_2_1[metric_index]['commit_num_trend'] = metric['commit_num_trend']
            else:
                second_half_dict_2_1[metric_index]['commit_num_trend'] = 0

            if 'dev_equivalent_trend' in metric.keys():
                second_half_dict_2_1[metric_index]['dev_equivalent_trend'] = metric['dev_equivalent_trend']
            else:
                second_half_dict_2_1[metric_index]['dev_equivalent_trend'] = 0

            if 'loc_trend' in metric.keys():
                second_half_dict_2_1[metric_index]['loc_trend'] = metric['loc_trend']
            else:
                second_half_dict_2_1[metric_index]['loc_trend'] = 0
            metric_index += 1
            # print(f"metric_index  = {metric_index}")

    # print("\n")
    # print(f"second_half_dict_2_1 = {second_half_dict_2_1}")



    dict_for_key_2_1 = {}
    index_2_1 = 0
    for i_2_1 in second_half_dict_2_1.items():
        # print(f"i_2_1 = {i_2_1}")
        first_dict = first_half_dict_2_1[i_2_1[1]['primary_email']]
        # print(f"first_dict = {first_dict}")
        dict_for_key_2_1[index_2_1] = {**first_dict, **{"*":"*"}, **i_2_1[1]}
        index_2_1 += 1


    write_data_for_key_and_worksheet(dict_for_key_2_1, worksheet_2_1)



# #2_2##############################################################################################################

    print(f"开始处理{key_2_2}")
    sheet_name_2_2 = key_2_2
    worksheet_2_2 = workbook.add_worksheet(sheet_name_2_2)
    worksheet_2_2.set_column_pixels(0, 20, 200)


    first_half_dict_2_2 = {}
    for index_2_2, i_2_2 in enumerate(total_content.items()):
        first_half_dict_2_2[i_2_2[1]['primary_email']] = {}
        first_half_dict_2_2[i_2_2[1]['primary_email']]['primary_email'] = i_2_2[1]['primary_email']
        first_half_dict_2_2[i_2_2[1]['primary_email']]['linked_emails'] = i_2_2[1]['linked_emails']
        first_half_dict_2_2[i_2_2[1]['primary_email']]['user_id'] = i_2_2[1]['user_id']
        first_half_dict_2_2[i_2_2[1]['primary_email']]['user_name'] = i_2_2[1]['user_name']


    # print(f"first_half_dict_2_2 = {first_half_dict_2_2}")


    second_half_dict_2_2 = {}
    for index_2_2_, i_2_2_ in enumerate(total_content.items()):
        second_half_dict_2_2[index_2_2_] = i_2_2_[1][key_2_2]
        second_half_dict_2_2[index_2_2_]['primary_email'] = i_2_2_[1]['primary_email']

        # print(f"second_half_dict_2_2 = {second_half_dict_2_2}")


    dict_for_key_2_2 = {}
    index_2_2 = 0
    for i_2_2 in second_half_dict_2_2.items():
        # print(f"i_2_2 = {i_2_2}")
        first_dict = first_half_dict_2_2[i_2_2[1]['primary_email']]
        # print(f"first_dict = {first_dict}")
        dict_for_key_2_2[index_2_2] = {**first_dict, **{"*":"*"}, **i_2_2[1]}
        index_2_2 += 1


    write_data_for_key_and_worksheet(dict_for_key_2_2, worksheet_2_2)



# #end##############################################################################################################



    workbook.close()



