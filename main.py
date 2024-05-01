# gspread 사용법(python gspread - google spreadsheet)
# url: https://greeksharifa.github.io/references/2023/04/10/gspread-usage/

import gspread
import re
from google.oauth2.service_account import Credentials

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

credentials = Credentials.from_service_account_file(
    'C:/Users/leekw/Desktop/gspread-python/service_account.json', # change back slash to '/'
    scopes=scopes
)

gc = gspread.authorize(credentials)

#gc = gspread.service_account()

#sh = gc.open("age")

sh = gc.open_by_url('https://docs.google.com/spreadsheets/d/1uxfzbLTm9lSKRM1dDCFxkjJj-Lr4knUGedVC3jZ1Up0/edit?usp=sharing')

print(sh.sheet1.get("A2"))

worksheet = sh.get_worksheet(0)

worksheet = sh.worksheet("Sheet1")

worksheet_list = sh.worksheets()  # 전체 worksheet 목록

# 특정 1개의 셀 값을 얻는 방법은 여러 가지가 있다.

# 1행 3열의 셀 값 얻기
val1 = worksheet.cell(2, 3).value
val2  = worksheet.acell('C1').value
val3 = worksheet.get('C1') # 사실 get은 범위 연산도 가능하다. 출력되는 결과를 보면 2차원 리스트로 조금 다르게 생겼다.

print(val1, val2, val3)
# 결과:
# 'MR-full-mAP'
# 'MR-full-mAP'
# [['MR-full-mAP']]


# cell formula를 사용하는 방법도 있다.
cell = worksheet.acell('C1', value_render_option='FORMULA').value  # 또는
cell = worksheet.cell(1, 3, value_render_option='FORMULA').value # 1행 3열 -> C1

# 특정 행이나 열 전체의 셀 값을 가져올 수 있다.

values_list = worksheet.row_values(1)
print(values_list)
values_list = worksheet.col_values(1)
print(values_list)

#이외에도 범위를 지정해서 셀 값을 가져올 수 있다.

##get_all_values()는 worksheet 전체의 값을 가져온다.
##get()는 지정한 범위의 셀 값을 전부 가져온다.
##batch_get()는 1번의 api call로 여러 범위의 셀 값을 가져올 수 있다.
##update()는 값을 가져오는 것이 아니라 원하는 값으로 셀을 수정할 수 있다.
##이제 batch_update()는 무슨 함수인지 알 것이다.

# worksheet의 모든 값을 list의 list 또는 dict의 list 형태로 가져올 수 있다.

list_of_lists = worksheet.get_all_values()
list_of_dicts = worksheet.get_all_records()

print("list_of_lists:", list_of_lists)
print("list_of_dicts:", list_of_dicts)

#셀 값 찾기
#기본적으로, 셀은 다음과 같은 값들을 attribute로 갖는다.
"""
value = cell().value
row_number = cell().row
column_number = cell().col
"""

#특정 문자열을 가지는 셀을 하나 또는 전부 찾는 방법은 아래와 같다.
cell = worksheet.find("age")
print("'age' is at Row%s, Col%s" % (cell.row, cell.col))

# 전부 찾기
cell_list = worksheet.findall("age")
#정규표현식을 사용할 수도 있다.

amount_re = re.compile(r'(Big|Enormous) dough')
cell = worksheet.find(amount_re)

# 전부 찾기
cell_list = worksheet.findall(amount_re)
print(cell_list)

#셀 값 업데이트하기
#먼저, 선택한 범위의 셀 내용을 지우는 방법은 다음과 같다.

# 리스트 안에 요소로 셀 범위나 이름을 부여한 named range를 넣을 수 있다.
#worksheet.batch_clear(["A1:B1", "C2:E2", "named_range"])

# 전체를 지울 수도 있다.
#worksheet.clear()

#특정 1개 또는 특정 범위의 셀을 지정한 값으로 업데이트할 수 있다.

worksheet.update('B1', 'Gorio')

# 1행 2열, 즉 B1
worksheet.update_cell(1, 2, 'Gorio')

# 범위 업데이트
#worksheet.update('A1:B2', [[1, 2], [3, 4]])

#서식 지정
#셀 값을 단순히 채우는 것 말고도 서식을 지정할 수도 있다.

# 볼드체로 지정
worksheet.format('A1:B1', {'textFormat': {'bold': True}})

#여러가지 설정을 같이 할 수도 있다. dict에 원하는 값을 지정하여 업데이트하면 된다.
worksheet.format("A2:B2", {
    "backgroundColor": {
      "red": 0.0,
      "green": 0.0,
      "blue": 0.0
    },
    "horizontalAlignment": "CENTER",
    "textFormat": {
      "foregroundColor": {
        "red": 1.0,
        "green": 1.0,
        "blue": 1.0
      },
      "fontSize": 12,
      "bold": True
    }
})

#또한, 셀 서식뿐만 아니라 워크시트의 일부 서식을 지정할 수도 있다.

#예를 들어 행 또는 열 고정은 다음과 같이 값을 얻거나 지정할 수 있다.
row_count = get_frozen_row_count(worksheet)
col_count = get_frozen_column_count(worksheet)
result1 = set_frozen(worksheet, rows=1)
result2 = set_frozen(worksheet, cols=1)
result3 = set_frozen(worksheet, rows=1, cols=0)

#셀의 높이나 너비를 지정하거나 데이터 유효성 검사, 조건부 서식 등을 지정할 수도 있다.

#설치 방법 및 사용법은 다음을 참고하자.

# https://gspread-formatting.readthedocs.io/en/latest/

# numpy, pandas와 같이 사용하기
# 워크시트 전체를 numpy array로 만들 수 있다.

import numpy as np
array = np.array(worksheet.get_all_values())

# 시트에 있는 header를 떼고 싶으면 그냥 [1:]부터 시작하면 된다.
array = np.array(worksheet.get_all_values()[1:])

# 물론 numpy array를 시트에 올릴 수도 있다.

import numpy as np

array = np.array([[1, 2, 3], [4, 5, 6]])
worksheet.update('A2', array.tolist())

#워크시트 전체를 불러와 pandas dataframe으로 만들고 싶으면 다음과 같이 쓰면 된다.
import pandas as pd
dataframe = pd.DataFrame(worksheet.get_all_records())

# dataframe의 header와 value를 전부 worksheet에 쓰는 코드 예시는 다음과 같다.
import pandas as pd
worksheet.update([dataframe.columns.values.tolist()] + dataframe.values.tolist())

#더 많은 기능은 다음 github을 참고하자.

#https://github.com/aiguofer/gspread-pandas
#https://github.com/robin900/gspread-dataframe