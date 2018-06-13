from openpyxl import load_workbook
from pathlib import Path
import pickle
from bs4 import BeautifulSoup
import requests as req

wb = load_workbook("order_agas_may_june.xlsx")


# print(wb.get_sheet_names())


def prettify(elem):
    """Return a pretty-printed XML string for the Element.
    """
    rough_string = tostring(elem, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="\t", encoding="utf8")


def get_pretty_names(args):
    rb = xlrd.open_workbook(args.good_name_file)
    sheet = rb.sheet_by_index(0)
    goods = []
    for rownum in range(1, sheet.nrows):
        row = sheet.row_values(rownum)
        goods.append((int(row[1]), row[0]))

    return dict(goods)


def get_raw_data(args):
    tree = parse(args.raw_file)
    root = tree.getroot()
    raw_data = []
    for item in root.findall(u'Цена_x0020_Прайса_x0020_Розничная'):
        good = {}
        good['price_date'] = item.find(u'Цена_x0020_Прайса_x0020_Розничная_x0020_Дата').text
        good['retail_price'] = item.find(u'Цена_x0020_Прайса_x0020_Розничная').text
        good['name'] = item.find(u'Номенклатура_x0020_Наименование').text
        good['gas_station'] = item.find(u'Объект_x0020_Управления_x0020_Родитель_x0020_ASPB').text
        good['code'] = item.find(u'Номенклатура_x0020_Эталон_x0020_Код').text
        raw_data.append(good)
    return raw_data


def generate_file(good_names, raw_data):
    menu = Element('menu')

    for i in raw_data:
        code = i['gas_station'].strip()
        int_code = re.findall(r'\d+', code)[0]

        if int_code != args.object_code:
            continue
        # item = SubElement(menu, 'Item')
        item_name = SubElement(menu, 'Item_Name_{}'.format(i['code']))
        item_name.text = good_names.get(int(i['code']), i['name'])
        item_price = SubElement(menu, 'Item_Price_{}'.format(i['code']))
        item_price.text = int(float(i['retail_price'].replace(',', '.')))

        with open(args.output_file, 'w') as f:
            f.write(prettify(menu))

        # orderer ask do not write xml-header (<?xml version="1.0" encoding="utf8"?>)
        lines = open(args.output_file).readlines()
        open(args.output_file, 'w').writelines(lines[1:])


def download_file(file_link):
    file_name = os.path.basename(file_link)
    try:
        f = urlopen(file_link)
        local_file_name = file_name
        local_file = open(local_file_name, "wb")
        local_file.write(f.read())
        local_file.close()
        f.close()
    finally:
        return file_name


def log(msg):
    now = datetime.datetime.now()
    report = '{} {}\n'.format(now, msg)
    with open('log.txt', 'a') as f:
        f.write(report)


def parse_report():
    with open("AGAS.html") as file:
        soup = BeautifulSoup(file.read(), "lxml")

    tbody = soup.find("tbody")
    trs = tbody.find_all("tr")
    for tr in trs:
        pass


if __name__ == "__main__":
    cwd = Path(".")
    # print(cwd.absolute())
    xls_files = [entry for entry in cwd.iterdir() if entry.suffix == ".xlsx"]
    parse_report()
