from bs4 import BeautifulSoup
from collections import namedtuple
import xlwt
import glob

Soup = namedtuple('Soup', 'name soup')
Table = namedtuple('Table', 'header rows')


def idea(filenames):
    soups = [Soup(filename, BeautifulSoup(open(filename))) for filename in filenames]
    soup_tables = {}

    def is_header(class_):
        if class_:
            return class_.startswith('datasheetTopLabel')
        else:
            return False

    def is_row(class_):
        if class_:
            return class_.startswith('rowEven') or class_.startswith('rowOdd')
        else:
            return False

    for soup in soups:
        tables = soup.soup.find_all('table')
        soup_tables[soup.name] = []
        for table in [t for t in tables if not t.find_all('table')]:
            header_cells = [row for row in table.find_all('td', {'class': is_header})]
            row_cells = []
            for row in table.find_all('tr', {'class': is_row}):
                row_cells.append([td for td in row.find_all('td')])
            tbl = Table(header_cells, row_cells)
            soup_tables[soup.name].append(tbl)

    font0 = xlwt.Font()
    font0.name = 'Times New Roman'
    font0.colour_index = 2
    font0.bold = False

    style0 = xlwt.XFStyle()
    style0.font = font0

    wb = xlwt.Workbook()
    ws = wb.add_sheet('This is awesome')

    row = 0
    for name, tables in soup_tables.items():
        for table in tables:
            if len(table[0]) > 0:
                for col, header_cell in enumerate([cell for cell in table.header
                                                   if cell.get_text(strip=True)]):
                    ws.write(row, col, header_cell.string)
                row += 1
                for roow in table.rows:
                    for col, cell in enumerate([r for r in roow if r.get_text(strip=True)]):
                        ws.write(row, col, cell.get_text(strip=True))
                    row += 1

    wb.save('example.xls')
 
 
if __name__ == '__main__':
    filenames = glob.glob('html/*.html')
    idea(filenames)