import os
import shutil
import docx
import random
import datetime

def get_random_question(exam, result, test_point, key):
    test_paper = []
    result_paper = []
    for tPkey in test_point.keys():
        random_list = [int(i) for i in range(len(exam[tPkey]))]
        bei_xuan_ti = len(exam[tPkey]) - 1
        for i in range(int(test_point[tPkey])):
            number = random.randint(0, bei_xuan_ti)
            test_paper.append(exam[tPkey][random_list[number]])
            result_paper.append(result[tPkey][random_list[number]])
            random_list.remove(random_list[number])
            bei_xuan_ti -= 1
    return test_paper, result_paper, key

def read_table_from_docx(file_name):
    doc_list = docx.Document(file_name)
    table = doc_list.tables
    return table[0]

def parse_table(table):
    dictionary_list = {}
    dan_xuan = {}
    dan_xuan_da_an = {}
    duo_xuan = {}
    duo_xuan_da_an = {}
    key_list = {}
    
    for n in range(1, len(table.rows)):
        rows_value = table.rows[n]
        rows_cell = [cell.text for cell in rows_value.cells]
        if rows_cell[2] in dictionary_list.keys():
            if rows_cell[3] not in dictionary_list[rows_cell[2]]:
                dictionary_list[rows_cell[2]].append(rows_cell[3])
        else:
            dictionary_list[rows_cell[2]] = []
            
        if rows_cell[3] not in key_list.keys():
            dan_xuan[rows_cell[3]] = []
            dan_xuan_da_an[rows_cell[3]] = []
            duo_xuan[rows_cell[3]] = []
            duo_xuan_da_an[rows_cell[3]] = []
            key_list[rows_cell[3]] = rows_cell[-1]
            dan_xuan[rows_cell[3]].append(rows_cell[6])
            dan_xuan_da_an[rows_cell[3]].append(rows_cell[7])
            duo_xuan[rows_cell[3]].append(rows_cell[10])
            duo_xuan_da_an[rows_cell[3]].append(rows_cell[11])
        else:
            dan_xuan[rows_cell[3]].append(rows_cell[6])
            dan_xuan_da_an[rows_cell[3]].append(rows_cell[7])
            duo_xuan[rows_cell[3]].append(rows_cell[10])
            duo_xuan_da_an[rows_cell[3]].append(rows_cell[11])

    return dan_xuan, dan_xuan_da_an, duo_xuan, duo_xuan_da_an, key_list

def generate_exam_paper(tests, result, dan_xuan, dan_xuan_da_an, duo_xuan, duo_xuan_da_an, random_question_dic):
    for i in key_list.keys():
        rand_number = random.randint(1, 2)
        if rand_number == 1 and must_be_duoxuan == 2:
            is_danxuan += int(key_list[i])
            if is_danxuan <= max_d_danxuan:
                random_question_dic[rand_number][i] = int(key_list[i])
            else:
                must_be_duoxuan = 0
                more_from_danxuan = is_danxuan - max_d_danxuan
                rand_number = 2
                if must_be_danxuan == 0 and is_duoxuan < max_duoxuan:
                    is_duoxuan += int(key_list[i])
                    random_question_dic[rand_number][i] = more_from_danxuan
        else:
            if must_be_danxuan == 0:
                rand_number = 2
                is_duoxuan += int(key_list[i])
                if is_duoxuan <= max_duoxuan:
                    random_question_dic[rand_number][i] = int(key_list[i])
                else:
                    must_be_danxuan = 2
                    more_from_duoxuan = is_duoxuan - max_duoxuan
                    rand_number = 1
                    if must_be_duoxuan == 2 and is_danxuan < max_d_danxuan:
                        is_danxuan += int(key_list[i])
                        random_question_dic[rand_number][i] = more_from_duoxuan

def write_to_file(file_name, title, content):
    with open(file_name, "w") as file:
        file.write(f"{' '*18}{title}\n")
        number = 0
        for i in content.keys():
            for x in content[i]:
                if int(i) == 1 and number == 0:
                    file.writelines("i am the king"+"\n")
                else:
                    if int(i) == 2 and number == len(content[1]):
                        file.writelines("i am the QQQ"+"\n")
                number += 1
                file.writelines(f"{number}. {x}\n")

def main():
    file_name = '1.docx'
    table = read_table_from_docx(file_name)
    dan_xuan, dan_xuan_da_an, duo_xuan, duo_xuan_da_an, key_list = parse_table(table)

    tests = {1: [], 2: []}
    result = {1: [], 2: []}
    is_danxuan = 0
    max_d_danxuan = 6
    must_be_danxuan = 0
    is_duoxuan = 0
    max_duoxuan = 4
    must_be_duoxuan = 2
    random_question_dic = {1: {}, 2: {}}
    
    generate_exam_paper(tests, result, dan_xuan, dan_xuan_da_an, duo_xuan, duo_xuan_da_an, random_question_dic)

    now = datetime.datetime.now()
    other_style_time = now.strftime("%Y年%m月")

    write_to_file("考试.txt", f"{other_style_time}考试", tests)
    write_to_file("答案.txt", f"{other_style_time}答案", result)

if __name__ == "__main__":
    main()
