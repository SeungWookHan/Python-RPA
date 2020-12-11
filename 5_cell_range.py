from openpyxl.utils.cell import coordinate_from_string
from openpyxl import Workbook
from random import *

wb = Workbook()
ws = wb.active

# 1줄씩 데이터 넣기
ws.append(["번호", "영어", "수학"])  # A B C
for i in range(1, 11):  # 10개 데이터 넣기
    ws.append([i, randint(0, 100), randint(0, 100)])

col_B = ws["B"]  # 영어 column만 가져오기
# print(col_B)
# for cell in col_B:
#     print(cell.value)

# col_range = ws["B:C"]  # 영어, 수학 column 함께 가져오기
# for cols in col_range:
#     for cell in cols:
#         print(cell.value)

row_title = ws[1]  # 1번째 row만 가져오기
# for cell in row_title:
#     print(cell.value)


row_range = ws[2:6]  # 2 번째 줄에서 6(포함) 기존 슬라이싱과 다름
for rows in row_range:
    for cell in rows:
        # print(cell.value, end=" ")
        # print(cell.coordinate, end=" ")  # A10,
        xy = coordinate_from_string(cell.coordinate)
        # print(xy, end=" ")  # 튜플 형태로 받아옴
        print(xy[0], end="")  # A
        print(xy[1], end=" ")  # B
    print()

wb.save("sample.xlsx")
