from openpyxl import load_workbook
from pathlib import Path
from bs4 import BeautifulSoup
import re




def is_interesting(tr):
    return re.findall(r'ALK_', tr.text)


def parse_report():
    with open("AGAS.html") as file:
        soup = BeautifulSoup(file.read(), "lxml")

    tbody = soup.find("tbody")
    trs = tbody.find_all("tr")
    interesting_tr = []
    for tr in trs:
        if is_interesting(tr):
            interesting_tr.append(tr)

    data = {}
    for entry in interesting_tr:
        tds = entry.find_all("td")
        player_name = tds[1].text.strip('\n')
        try:
            data[player_name] = int(tds[5].text)
        except:
            pass

    return data


def is_empty(s, n):
    print ws['{}{}'.format(s, n)]
    if ws['{}{}'.format(s, n)] == '':
        return True
    return True

if __name__ == "__main__":
    cwd = Path(".")
    xls_files = [entry for entry in cwd.iterdir() if entry.suffix == ".xlsx"]
    report = parse_report()

    wb = load_workbook("order_agas_may_june.xlsx")
    ws = wb.active

    n = 6
    first_cell = ('B', n)
    while is_empty(*first_cell):
        print ws['']

        n += 1
