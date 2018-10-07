import openpyxl

# 엑셀 파일 불러오기 sample.xlsx
wb = openpyxl.load_workbook('sample.xlsx')
ws = wb['Sheet1']
i = 1

for row in ws.rows:
    print(("B" + str(i) + " : " + str(row[1].value)))
    i += 1

print(str(ws.max_row - 1) + "개의 행이 있습니다.")
# 시작 번호
num = 1
# 마지막 번호
end_value = ws.max_row - 1
# A열 1행
nameA = input("A열 Name 입력 : ")
print('** Name: ' + nameA + '-')
start_numA = int(input("A열 시작번호 : "))
print('** Name: ' + nameA + '-' + str(start_numA) + ' Type: ')
input_typeA = input("A열 Type 입력 : ")
print('** Name: ' + nameA + '-' + str(start_numA) + ' Type: ' + str(input_typeA))

# A열 2행
input_second = input("두번째 줄에 들어갈 것 입력 : ")
print('*' + input_second)
input_chkType = input("두번째 줄에 type을 입력할래? (y/n) : ")
if input_chkType == 'y':
    data_type = ", type=" + input_typeA.upper()
    print('*' + input_second + data_type)
else:
    data_type = ""
    print('*' + input_second + data_type)

# A열 3행
input_third = input("A열 3행ex) cover : ")
print(input_third + "-")
input_thirdA = int(input("A열 3행 시작번호 : "))
print(input_third + "-" + str(start_numA) + ",")
input_chkThird = input("세번째 줄 입력할거 있는가  (없으면 빈칸 있으면 입력 ex)11,11,) : ")
print(input_third + "-" + str(start_numA) + "," + input_chkThird + str(ws['B' + str(num)].value))

for i in range(1, end_value * 3, 3):
    ws['A' + str(i)] = '** Name: ' + nameA + '-' + str(start_numA) + ' Type: ' + str(input_typeA)
    ws['A' + str(i + 1)] = '*' + input_second + data_type
    ws['A' + str(i + 2)] = input_third + "-" + str(start_numA) + "," + input_chkThird + str(ws['B' + str(num)].value)
    print(ws['A' + str(i)].value)
    print(ws['A' + str(i + 1)].value)
    print(ws['A' + str(i + 2)].value)
    start_numA += 1
    num += 1

# 엑셀 파일 저장
save_name = input("저장할 파일명 입력 : ")
wb.save(save_name + ".xlsx")
wb.close()
exit()
