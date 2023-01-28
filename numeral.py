from collections import deque

import pandas as pnd
import xlsxwriter as xls

def numeral(num):
    num = int(num)
    if num == 0:
        return "ноль"
    
    count = -1
    num_spl = deque()
    pref = [["тысяча","тысячи","тысяч"],["миллион","миллиона","миллионов"],["миллиард","миллиарда","миллиардов"]]
    
    while num > 0:
        if num % 1000 != 0:
            if count >= 0:
                if len(str(num)) > 1:
                    if all([str(num)[-1] == "1",str(num)[-2] != "1"]):
                        num_spl.appendleft(pref[count][0])
                    elif all([any([str(num)[-1] == "2", str(num)[-1] == "3", str(num)[-1] == "4"]), str(num)[-2] != "1"]):
                        num_spl.appendleft(pref[count][1])
                    elif all([any([
                        str(num)[-1] == "5", str(num)[-1] == "6", str(num)[-1] == "7",
                        str(num)[-1] == "8", str(num)[-1] == "9", str(num)[-1] == "0"
                    ])]) or all([str(num)[-2] == "1"]):
                        num_spl.appendleft(pref[count][2])
                    else:
                        pass
                else:
                    if str(num)[-1] == "1":
                        num_spl.appendleft(pref[count][0])
                    elif any([str(num)[-1] == "2", str(num)[-1] == "3", str(num)[-1] == "4"]):
                        num_spl.appendleft(pref[count][1])
                    elif any([
                        str(num)[-1] == "5", str(num)[-1] == "6", str(num)[-1] == "7",
                        str(num)[-1] == "8", str(num)[-1] == "9", str(num)[-1] == "0"
                    ]):
                        num_spl.appendleft(pref[count][2])
                    else:
                        pass
            else:
                pass
            num_spl.appendleft(num % 1000)
            num = num // 1000
            count += 1
        else:
            num = num // 1000
            count += 1
    point = 0
    for i in range(len(num_spl)-1,-1,-1):
        if "тысяч" in str(num_spl[i]):
            point = 1
        if i % 2 == 0:
            num_spl[i],point = convert(num_spl[i],point)

    return " ".join(num_spl)

def convert(num, point):
    rez = []
    ones = {
        "3" : "три", "4" : "четыре", "5" : "пять", "6" : "шесть",
        "7" : "семь", "8" : "восемь", "9" : "девять", "0" : "",
    }
    if point == 1:
        ones["1"] = "одна"
        ones["2"] = "две"
    else:
        ones["1"] = "один"
        ones["2"] = "два"

    teen = {
        "10" : "десять", "11" : "одиннадцать", "12" : "двенадцать", "13" : "тринадцать",
        "14" : "четырнадцать", "15" : "пятнадцать", "16" : "шестнадцать",
        "17" : "семнадцать", "18" : "восемнадцать", "19" : "девятнадцать",
    }
    dec = {
        "2" : "двадцать", "3" : "тридцать", "4" : "сорок", "5" :"пятьдесят", "6" : "шестьдесят",
        "7" : "семьдесят", "8" : "восемьдесят", "9" : "девяносто", "0" : "",
    }
    hundr = {
        "1" : "сто", "2" : "двести", "3" : "триста", "4" : "четыреста", "5" : "пятьсот",
        "6" : "шестьсот", "7" : "семьсот", "8" : "восемьсот", "9" : "девятьсот",
    }

    num_s = str(num)
    if len(num_s) == 3 and num_s[1] == "1":
        rez.append(hundr[num_s[0]])
        rez.append(teen[num_s[1:]])
    elif len(num_s) == 2 and num_s[0] == "1":
        rez.append(teen[num_s[0:]])
    elif len(num_s) == 3 and num_s[1] != "1":
        rez.append(hundr[num_s[0]])
        rez.append(dec[num_s[1]])
        rez.append(ones[num_s[2]])
    elif len(num_s) == 2 and num_s[0] != "1":
        rez.append(dec[num_s[0]])
        rez.append(ones[num_s[1]])
    else:
        rez.append(ones[num_s[0]])
    rez = " ".join(rez).rstrip()
    return " ".join(rez.split()), 0


def read_and_write(path_in,path_out):
    data = pnd.read_excel(path_in)
    workbook = xls.Workbook(path_out)
    worksheet = workbook.add_worksheet("numerals")
    for ind,item in enumerate(data["ИТОГО"]):
        rez = numeral(item)
        worksheet.write(ind,0,rez)
    workbook.close()


if __name__ == "__main__":
    l = ["1495000","982421518","891275","23","4121","400","910000","902000",]
    for num in l:
        print(numeral(num))
    """ r_xls = "test.xlsx"
    w_xls = "train.xlsx"
    read_and_write(r_xls,w_xls) """
