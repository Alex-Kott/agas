# coding=utf-8
import re
import sys
sys.path.insert(0, './venv/lib/python2.7/site-packages/')

from openpyxl import load_workbook
from pathlib import Path
from bs4 import BeautifulSoup
from openpyxl.styles import PatternFill


def is_interesting(tr, azs_prefixes):
    for prefix in azs_prefixes:
        if tr.text.find(prefix) != -1:
            return True
    return False


def parse_report(file_name):
    azs_prefixes = ['ALK', 'CHO', 'EKA', 'HMO', 'IVO', 'KDK', 'KEO', 'KYK', 'MSK', 'NNO', 'NSO',
                    'OMO', 'SPB', 'SVO', 'TUO', 'YAR', 'IAR', 'KRR', 'MOW', 'MSK', 'NN', 'RYZ',
                    'CEK', 'MSK', 'NN', 'SPB', 'TUO', 'YAR']
    with open(str(file_name)) as file:
        soup = BeautifulSoup(file.read(), "lxml")

    tbody = soup.find("tbody")
    trs = tbody.find_all("tr")
    interesting_tr = []
    for tr in trs:
        if is_interesting(tr, azs_prefixes):
            interesting_tr.append(tr)

    data = {}
    for entry in interesting_tr:
        tds = entry.find_all("td")
        player_name = tds[1].text.strip('\n')
        try:
            azs_number = int(re.findall(r'(?<=\D{3}[_-])\d+', player_name)[0])
            data[azs_number] = int(tds[5].text)
        except:
            pass

    return data


def is_empty(s, n):
    print ws['{}{}'.format(s, n)]
    if ws['{}{}'.format(s, n)] == '':
        return True
    return True


def parse_arguments(arguments):
    if len(arguments) != 3:
        raise Exception("Must be 2 argument (.html-file and .xlsx-file)")

    args = [Path(arg) for arg in arguments]

    for arg in args:
        if arg.suffix == ".html":
            agas_file_name = arg
        elif arg.suffix == ".xlsx":
            order_file_name = arg
        elif arg.suffix == ".py":
            pass
        else:
            raise Exception("Unknown file")

    return agas_file_name, order_file_name


def print_row(row):
    for i in row:
        print i.coordinate
    line = ['({}) {}'.format(cell.value, cell.coordinate) for cell in row if cell.value is not None]
    print "    ".join(line)

if __name__ == "__main__":
    (agas_file_name, order_file_name) = parse_arguments(sys.argv)

    report = parse_report(agas_file_name)
    wb = load_workbook(str(order_file_name))
    ws = wb.active

    flag = False
    for row in ws.iter_rows():
        if row[1].value == u"№ АЗС:":
            print_row(row)
            print '\n\n'
            flag = True
            continue
        elif row[1].value == None:
            flag = False
        if flag == False:
            continue

        azs_code = row[1].value
        val = 0
        if type(row[12].value) is long:
            n = row[12].value
            try:
                row16 = report[azs_code] - n
                val = report[azs_code]
                row[15].value = val
                if row16 < 0:
                    row[16].fill = PatternFill(fill_type='solid',
                                               start_color='FF9900',
                                               end_color='FF9900')
            except KeyError as e:
                row[15].value = "Данных нет"
        elif type(row[12].value) is str:
            row[15].value = val

                





    wb.save('file.xlsx')
