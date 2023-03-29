import openai
import streamlit as st
from openpyxl import load_workbook
from ast import literal_eval
from konlpy.tag import Okt
import sys
import requests
import time
import re

def import_excel(file_path):

    if ".xlsm" in file:
        wb = load_workbook(file_path,data_only=True,keep_vba=True)
    elif ".xlsx" in file:
        wb = load_workbook(file_path,data_only=True)
    else :
        print("파일 형식 오류")
        sys.exit()

    return wb

def make_dict(ws_list):
    trans_dict = {}
    for order, wsname in enumerate(ws_list):
        ws = wb[wsname]
        max_row = ws.max_row
        max_col = ws.max_column
        print(max_row, max_col)
        for row in range(1,max_row+1):
            for col in range(1, max_col+1):
                target = ws.cell(row, col).value
                if target == None:
                    continue
                elif is_korean_sentence(str(target)):
                    continue

                key = str(order)+"-"+str(row)+"-"+str(col)
                val = target

                trans_dict[key] = val

    return trans_dict


def slice_dict(dict, max_length):
    """
    Slices a dictionary into several shorter dictionaries with the total length of the values less than max_length.
    """
    result = []
    current_dict = {}
    current_length = 0
    for key, value in dict.items():
        key_length = len(str(key))
        value_length = len(str(value))
        if current_length +key_length + value_length > max_length:
            result.append(current_dict)
            current_dict = {}
            current_length = 0
        current_dict[key] = value
        current_length += key_length
        current_length += value_length
    result.append(current_dict)
    return result


def is_korean_sentence(sentence):
    # Define a regular expression that only matches Korean characters
    korean_pattern = re.compile(r'^[ㄱ-ㅎ가-힣]+$')
    
    # Check if the sentence matches the Korean pattern
    if korean_pattern.match(sentence):
        return True
    else:
        return False




############### main ################


openai.api_key = openai.api_key = st.secrets["OPENAI_KEY"]

st.set_page_config(layout="wide")

st.title('Assurance DA')
st.header('File Translator')
st.write('번역할 파일을 선택하세요.')
st.write('Developed by Assurance DA (beomsun.go@pwc.com)')

file = st.file_uploader(
    "파일을 선택하세요(xlsx, xlsm만 가능)",
    type=['xlsx', 'xlsm']
)

if file is not None:
    st.subheader(file)
    # file_path = r"C:\Users\bgo006\Desktop\CorDA\project\chatgpt\translator\sample_eng_2.xlsm"
    lang = "English"

    #################### 엑셀 불러온 후 모든 글자 긁어오기 #####################
    wb = import_excel(file_path=file)

    ws_list = wb.sheetnames


    #################### dictionary 생성 ###################
    trans_dict = make_dict(ws_list = ws_list)
    st.subheader("Excel data was loaded.")

    ###################### 1,500자 내로 자르기 ###################
    sliced_dicts = slice_dict(trans_dict,2500) # 한자는 1,300자로 하는게 안전한듯 # 영어는 2500자?
    st.subheader("Input dictionary was created.")

    answer_dicts = {}
    st.subheader("번역시작")
    for trytime, sliced_dict in enumerate(sliced_dicts):
        messages = []
        st.subheader(f"input : {trytime}")
        st.subheader(len(str(sliced_dict)))
#         st.subheader(str(sliced_dict))
        messages.append({"role": "system", "content": 'Dictionary is one of the type of variables in python that contains keys and values.'})
        # messages.append({"role": "system", "content": 'Please translate sentenses and words from English to Korean. What you should translate are values in below dictionary and output type is also dictionary which has same keys with input dictionary'})
        messages.append({"role": "system", "content": f'Please translate all the {lang} sentenses and words in the dictionary below into Korean. What you should translate are all the sentenses and words and output type is also dictionary which has same keys with input dictionary'})
        messages.append({"role": "system", "content": str(sliced_dict)})
        try:
            print("try : 1")
            completions = openai.ChatCompletion.create(
                # model="gpt-4",
                model="gpt-3.5-turbo",
                messages=messages,
                timeout=60
            )
            print("used token :"+str(completions.usage['total_tokens']))
            answer = completions.choices[0]['message']['content']
#             print(answer)
            answer_dict = literal_eval(answer)
            # print(answer_dict)
            answer_dicts.update(answer_dict)
        except requests.exceptions.Timeout:
            time.sleep(5)
            print("try : 2 - timeout")
            completions = openai.ChatCompletion.create(
                # model="gpt-4",
                model="gpt-3.5-turbo",
                messages=messages,
                timeout=60
            )
            print("used token :"+str(completions.usage['total_tokens']))
            answer = completions.choices[0]['message']['content']
#             print(answer)
            answer_dict = literal_eval(answer)
            # print(answer_dict)
            answer_dicts.update(answer_dict)
        except SyntaxError:
            time.sleep(5)
            print("try : 2 - syntax")
            completions = openai.ChatCompletion.create(
                # model="gpt-4",
                model="gpt-3.5-turbo",
                messages=messages,
                timeout=60
            )
            print("used token :"+str(completions.usage['total_tokens']))
            answer = completions.choices[0]['message']['content']
#             print(answer)
            answer_dict = literal_eval(answer)
            # print(answer_dict)
            answer_dicts.update(answer_dict)
        except :
            continue

    for key_answer in answer_dicts:
        val_answer = answer_dicts[key_answer]
        key_answer_list = key_answer.split("-")
        wsname_answer = ws_list[int(key_answer_list[0])]
        row_answer = int(key_answer_list[1])
        col_answer = int(key_answer_list[2])
        wb[wsname_answer].cell(row_answer,col_answer).value = val_answer
#         print(val_answer, wsname_answer, row_answer, col_answer, wb[wsname_answer].cell(row_answer,col_answer).value)
    st.subheader("번역완료")
    #### output 생성 ####

    # output_path = file_path[:-5]+"_output."+file_path[-4:]
    output_file = f"{file.name.split('.')[0]}_output.xlsx"
    wb.save(output_file)
    st.success(f"Modified data saved to {output_file}.")
    wb.close()
