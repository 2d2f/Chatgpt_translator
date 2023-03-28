import openai
import streamlit as st
from openpyxl import load_workbook
from ast import literal_eval
from konlpy.tag import Okt

openai.api_key = "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"

def slice_dict(d, max_length):
    """
    Slices a dictionary into several shorter dictionaries with the total length of the values less than max_length.
    """
    result = []
    current_dict = {}
    current_length = 0
    for key, value in d.items():
        value_length = len(str(value))
        if current_length + value_length > max_length:
            result.append(current_dict)
            current_dict = {}
            current_length = 0
        current_dict[key] = value
        current_length += value_length
    result.append(current_dict)
    return result



#################### 엑셀 불러온 후 모든 글자 긁어오기 #####################
wb = load_workbook(r"C:\Users\bgo006\Desktop\CorDA\project\chatgpt\translator\test_eng.xlsx",data_only=True)
ws_list = wb.sheetnames
trans_dict = {}
# print(ws_list)
for order, wsname in enumerate(ws_list):
    ws = wb[wsname]
    max_row = ws.max_row
    max_col = ws.max_column
    print(max_row, max_col)
    for row in range(1,max_row+1):
        for col in range(1, max_col+1):
            if ws.cell(row, col).value == None:
                continue
            
            key = wsname+"-:-"+str(row)+"-:-"+str(col)
            val = ws.cell(row,col).value

            trans_dict[key] = val

# print(trans_dict)

###################### 1,500자 내로 자르기 ###################
sliced_dicts = slice_dict(trans_dict,1500)
answer_dicts = {}

for sliced_dict in sliced_dicts:
    messages = []
    print(len(str(sliced_dict)))
    
    messages.append({"role": "system", "content": 'Dictionary is one of the type of variables in python that contains keys and values. I want to translate the values of dictionary.'})
    # messages.append({"role": "system", "content": 'Please translate sentenses and words from English to Korean. What you should translate are values in below dictionary and output type is also dictionary which has same keys with input dictionary'})
    messages.append({"role": "system", "content": 'Please translate the values in below dictionary from English to Korean. What you should translate are all the sentenses and words and output type is also dictionary which has same keys with input dictionary'})
    messages.append({"role": "system", "content": str(sliced_dict)})

    completions = openai.ChatCompletion.create(
        # model="gpt-4",
        model="gpt-3.5-turbo",
        messages=messages
    )
    answer = completions.choices[0]['message']['content']
    answer_dict = literal_eval(answer)
    print(answer_dict)
    answer_dicts.update(answer_dict)

for key_answer in answer_dicts:
    val_answer = answer_dicts[key_answer]
    print(val_answer)
    key_answer_list = key_answer.split("-:-")
    wsname_answer = key_answer_list[0]
    print(wsname_answer)
    row_answer = int(key_answer_list[1])
    print(row_answer)
    col_answer = int(key_answer_list[2])
    print(col_answer)
    wb[wsname_answer].cell(row_answer,col_answer).value = val_answer
    print(wb[wsname_answer].cell(row_answer,col_answer).value)

wb.save(r"C:\Users\bgo006\Desktop\CorDA\project\chatgpt\translator\test_eng_output.xlsx")

wb.close()