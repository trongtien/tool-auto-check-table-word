import docx
from numpy import random



'''
    index   0
    a       1
    b       2
    c       3
    d       4
    e       5
    f       6
    g       7
    h       8

'''

'''
1d      -> (1, 4)   : max -> 4 
2e      -> (2, 5)   : max -> 5
3c      -> (3, 3)   : max -> 3
4b      -> (4, 2)   : max -> 2
5a      -> (5, 1)   : max -> 2
6a      -> (6, 1)   : max -> 2
7e      -> (7, 5)   : max -> 5
8c      -> (8, 3)   : max -> 4
9d      -> (9, 4)   : max -> 4
10h     -> (10, 8)  : max -> 8
11e     -> (11, 5)  : max -> 5
12e     -> (12, 5)  : max -> 5
13d     -> (13, 4)  : max -> 4
14c     -> (12, 3)  : max -> 3
15a     -> (15, 1)  : max -> 2
16c     -> (16, 3)  : max -> 3
17e     -> (17, 5)  : max -> 5
18c     -> (18, 3)  : max -> 3
19d     -> (19, 4)  : max -> 4
20b     -> (20, 2)  : max -> 4
21c     -> (21, 3)  : max -> 3
22d     -> (22, 4)  : max -> 4
'''

enum_answer = {
    1: dict(selected_answer=4, max=4, cell=1),
    2: dict(selected_answer=5, max=5, cell=2),
    3: dict(selected_answer=3, max=3, cell=3),
    4: dict(selected_answer=2, max=2, cell=4),
    5: dict(selected_answer=1, max=2, cell=5),
    6: dict(selected_answer=1, max=2, cell=6),
    7: dict(selected_answer=5, max=5, cell=7),
    8: dict(selected_answer=3, max=4, cell=8),
    9: dict(selected_answer=4, max=4, cell=9),
    10: dict(selected_answer=8, max=8, cell=10),
    11: dict(selected_answer=5, max=5, cell=1),
    12: dict(selected_answer=5, max=5, cell=2),
    13: dict(selected_answer=4, max=4, cell=3),
    14: dict(selected_answer=3, max=3, cell=4),
    15: dict(selected_answer=1, max=2, cell=5),
    16: dict(selected_answer=3, max=3, cell=6),
    17: dict(selected_answer=5, max=5, cell=7),
    18: dict(selected_answer=3, max=3, cell=8),
    19: dict(selected_answer=4, max=4, cell=9),
    20: dict(selected_answer=2, max=4, cell=10),
    21: dict(selected_answer=3, max=3, cell=1),
    22: dict(selected_answer=4, max=4, cell=2)
}

path = 'C:/Users/base/Documents/word/BT-N4.docx'
doc = docx.Document(path)

def automationRandomSelectedAnswer(num_max, selected_answer):
    # 1-> num_max
    num_list_random = list(range(1, num_max + 1))
    num_random = random.choice(num_list_random)

    if num_random == selected_answer:
        return automationRandomSelectedAnswer(num_max, selected_answer)

    return num_random


def random_question_correl_answer(num_cell_root):
    question = []
    # 12 -> 18
    num_correl_answer_list_random = list(range(12, 19))
    num_correl_answer = random.choice(num_correl_answer_list_random)
    num_cell = num_cell_root


    for _ in range(num_correl_answer):
        num_cell_random = random.choice(num_cell)
        num_cell.remove(num_cell_random)

        question.append(num_cell_random)

    return question 

def automationTick(index: str, path_doc_save_file: str):
    print('==============================>')
    print("start mapping cell")

    file_name_suffix = "BT-N4-"+index+".docx"
    value_default_cell='x'

    # Max row table tick
    question = random_question_correl_answer(list(range(1, 23)))
    print('+++++++++++++++++++++++++++++')
    print('Question tick', len(question))
    print('+++++++++++++++++++++++++++++')

    # Get table row
    table = doc.tables[0]

    for numCell in list(range(1, 23)):
        
        is_exist_question = numCell in question

        selected_answer = enum_answer.get(numCell)['selected_answer']
        num_cell_set_value = enum_answer.get(numCell)['cell']


        max_answer = enum_answer.get(numCell)['max']
        selected_answer_random = automationRandomSelectedAnswer(max_answer, selected_answer)

        '''
            [11, 20] -> +11
            [21, 30] -> +21
        '''
        if numCell in list(range(11, 21)):
            selected_answer = selected_answer + 11
            selected_answer_random = selected_answer_random + 11

        if numCell in list(range(21, 31)):
            selected_answer = selected_answer + 21
            selected_answer_random = selected_answer_random + 21

        '''
            Add value table cell
        '''
        print(numCell,'is_exist_question', is_exist_question, 'num_cell_set_value', num_cell_set_value, 'selected_answer', selected_answer)
        if is_exist_question :
            table.cell(num_cell_set_value, selected_answer).text = value_default_cell
            print('(', num_cell_set_value,',',selected_answer, ')')
        else:
            table.cell(num_cell_set_value, selected_answer_random).text = value_default_cell
            print('(', num_cell_set_value,',',selected_answer_random, ')')
       

    doc.save(path_doc_save_file + file_name_suffix)
    print('+++++++++++++++++++++++++++++')
    print("Save file success")
    print('==============================>')
    

def main():
    num_word_generator = 1
    path_save_file = 'C:/Users/base/Documents/word/convert/'
    for index in range(num_word_generator):
        automationTick(str(index+1), path_save_file)


main()