# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import openpyxl
import random
xlsx_file = 'member.xlsx'
wb_obj = openpyxl.load_workbook(xlsx_file)
sheet = wb_obj.active

menber_list = []
for row in sheet:
    t_list = []
    for cell in row:
        t_list.append(cell.value)
    menber_list.append(t_list)

# # 1그룹 선정
# h1 = [] # 행정 6급
# h2 = [] #행정 6급 외
# for ml in menber_list:
#     if ml[1] =='행정' and ml[2] =='6급':
#         h1.append(ml)
#     else:
#         h2.append(ml)
#
#
# hrcNum = list(range(1,  len(h1)))
# ehrcNum = list(range(1, len(h2)))
# for i in range(0,20):
#     # 랜덤함수 추출
#     hrc = random.choice(hrcNum)
#     hrcNum.remove(hrc)
#     ehrc = random.choice(ehrcNum)
#     ehrcNum.remove(ehrc)
#
#
#     #1그룹 추출
#     print("1그룹, 행정6급," + str(h1[hrc][0]) +", " + str(h1[hrc][3])+", " + str(h1[hrc][4]))
#     menber_list.remove(h1[hrc])
#     print("1그룹, 행정6급," + str(h2[ehrc][0]) + ", " + str(h2[ehrc][3]) + ", " + str(h2[ehrc][4]))
#     menber_list.remove(h2[ehrc])

def printMember(gnumber, arg1, arg2,rnm):
    print()
    # 2그룹 선정
    h1 = []  # 행정 7급
    h2 = []  # 행정 7급 외
    for ml in menber_list:
        if (ml[1] == arg1 and ml[2] == arg2 ) or (ml[1] == arg1 and '' == arg2):
            h1.append(ml)
        else:
            h2.append(ml)

    hrcNum = list(range(1, len(h1)))
    ehrcNum = list(range(1, len(h2)))
    for i in range(0, rnm):
        # 랜덤함수 추출
        hrc = random.choice(hrcNum)
        hrcNum.remove(hrc)
        ehrc = random.choice(ehrcNum)
        ehrcNum.remove(ehrc)

        # 1그룹 추출
        print(gnumber + str(h1[hrc][0]) + ", " + str(h1[hrc][3]) + ", " + str(h1[hrc][4]))
        menber_list.remove(h1[hrc])
        print(gnumber + str(h2[ehrc][0]) + ", " + str(h2[ehrc][3]) + ", " + str(h2[ehrc][4]))
        menber_list.remove(h2[ehrc])
def printMemberEx1(gnumber, arg1, arg2,rnm,rnm2):
    print()
    # 2그룹 선정
    h1 = []  # 행정 7급
    h2 = []  # 행정 7급 외
    for ml in menber_list:
        if (ml[1] == arg1 and ml[2] == arg2 ) or (ml[1] == arg1 and '' == arg2):
            h1.append(ml)
        else:
            h2.append(ml)

    hrcNum = list(range(1, len(h1)))
    ehrcNum = list(range(1, len(h2)))
    for i in range(0, rnm):
        # 랜덤함수 추출
        hrc = random.choice(hrcNum)
        hrcNum.remove(hrc)
        print(gnumber + str(h1[hrc][0]) + ", " + str(h1[hrc][3]) + ", " + str(h1[hrc][4]))
        menber_list.remove(h1[hrc])

    for i in range(0, rnm2):
        ehrc = random.choice(ehrcNum)
        ehrcNum.remove(ehrc)
        print(gnumber + str(h2[ehrc][0]) + ", " + str(h2[ehrc][3]) + ", " + str(h2[ehrc][4]))
        menber_list.remove(h2[ehrc])

def printMemberEx2(gnumber, arg1, arg2,rnm,rnm2,rnm3):
    print()
    # 2그룹 선정
    h1 = []  # 행정 7급
    h2 = []  # 행정 7급 외
    h3 = []
    for ml in menber_list:
        if (ml[1] == arg1):
            h1.append(ml)
        elif  (ml[1] == arg2):
            h2.append(ml)
        else:
            h3.append(ml)

    hrcNum = list(range(1, len(h1)))
    ehrcNum = list(range(1, len(h2)))
    ohrcNum = list(range(1, len(h3)))
    for i in range(0, rnm):
        # 랜덤함수 추출
        hrc = random.choice(hrcNum)
        hrcNum.remove(hrc)
        print(gnumber + str(h1[hrc][0]) + ", " + str(h1[hrc][3]) + ", " + str(h1[hrc][4]))
        menber_list.remove(h1[hrc])
    for i in range(0, rnm2):
        ehrc = random.choice(ehrcNum)
        ehrcNum.remove(ehrc)
        print(gnumber + str(h2[ehrc][0]) + ", " + str(h2[ehrc][3]) + ", " + str(h2[ehrc][4]))
        menber_list.remove(h2[ehrc])
    for i in range(0, rnm3):
        ohrc = random.choice(ohrcNum)
        ohrcNum.remove(ohrc)
        print(gnumber + str(h3[ohrc][0]) + ", " + str(h3[ohrc][3]) + ", " + str(h3[ohrc][4]))
        menber_list.remove(h3[ohrc])
def printMemberEx3(gnumber, arg1, rnm1,rnm2):
    print()
    # 2그룹 선정
    h1 = []  # 행정 7급
    h2 = []  # 행정 7급 외
    h3 = []
    for ml in menber_list:
        if (ml[1] == arg1.split(',')[0]):
            h1.append(ml)
        elif  (ml[1] == arg1.split(',')[1] or ml[1] == arg1.split(',')[2]):
            h2.append(ml)


    hrcNum = list(range(1, len(h1)))
    ehrcNum = list(range(1, len(h2)))
    ohrcNum = list(range(1, len(h3)))
    for i in range(0, rnm1):
        # 랜덤함수 추출
        hrc = random.choice(hrcNum)
        hrcNum.remove(hrc)
        print(gnumber + str(h1[hrc][0]) + ", " + str(h1[hrc][3]) + ", " + str(h1[hrc][4]))
        menber_list.remove(h1[hrc])
    for i in range(0, rnm2):
        ehrc = random.choice(ehrcNum)
        ehrcNum.remove(ehrc)
        print(gnumber + str(h2[ehrc][0]) + ", " + str(h2[ehrc][3]) + ", " + str(h2[ehrc][4]))
        menber_list.remove(h2[ehrc])


try:
    printMemberEx2('9그룹, 농업.임업,','농업','임업',7,3,10)
except:
    print("9그룹, 농업.임업 오류 ")
try:
    printMemberEx3('10그룹, 해양수.위생.조리,', '해양수산,위생,조리', 18, 2)
except:
    print("10그룹, 해양수.위생.조리 오류 ")
try:
    printMember('1그룹, 행정6급,','행정','6급',20)
except:
    print("1그룹, 행정6급 오류 ")
try:
    printMember('2그룹, 행정7급,','행정','7급',15)
except:
    print("2그룹, 행정7급 오류 ")
try:
    printMember('3그룹, 행정8급,','행정','',10)
except:
    print("3그룹, 행정8급 오류 ")
try:
    printMember('4그룹, 행정9급,','행정','',10)
except:
    print("4그룹, 행정9급 오류 ")

try:
    printMember('5그룹, 사서,','사서','',10)
except:
    print("5그룹, 사서 오류 ")
try:
    printMember('6그룹, 전산,','전산','',10)
except:
    print("6그룹, 전산 오류 ")
try:
    printMember('7그룹, 공업,','공업','',15)
except:
    print("7그룹, 공업 오류 ")
try:
    printMemberEx1('8그룹, 시설,','시설','',10,20)
except:
    print("8그룹, 시설 오류 ")
try:
    printMember('11그룹, 기타직렬,','기타직렬','',10)
except:
    print("11그룹, 기타직렬 오류 ")

try:
    printMember('12그룹, 관리운영직,','관리운영직','',5)
except:
    print("12그룹, 관리운영직 오류 ")

try:
    printMember('13그룹, 대학회계 7급(사무원),','대학회계','',20)
except:
    print("13그룹, 대학회계 오류 ")
try:
    printMember('14그룹, 대학회계,','대학회계','',5)
except:
    print("14그룹, 대학회계 오류 ")
try:
    printMember('15그룹, 대학회계,','대학회계','',5)
except:
    print("15그룹, 대학회계 오류 ")


print("프로그램 출력종료...")
input()







