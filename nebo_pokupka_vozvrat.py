import xlrd, re, os, string, pprint
import openpyxl as op

def saveIntoWb(tables): #type=dict
    wb = op.Workbook()
    for i in tables.keys():
        ws = wb.create_sheet(0)
        ws.title = i
        data = tables[i]
        for nrow in range(0, len(data)):
            res_nrow = nrow + 1
            row_data = data[nrow]
            for ncol in range(0, len(row_data)):
                ws.cell(row=res_nrow, column=ncol + 1).value = row_data[ncol]

    wb.save('result.xlsx')


def colIndexes(caps=False):
    a = list(string.ascii_lowercase)
    if caps:
        b = [i.upper() for i in a]
        a = b
    expansion = []
    for i in a:
        for k in a:
            expansion.append(i + k)
    a += expansion
    d = {a[i]:i for i in range(len(a))}
    return d


def getTable(fname):
    sb = xlrd.open_workbook(fname, logfile=open(os.devnull, 'w'))
    ws = sb.sheet_by_index(0)
    rows = ws.nrows
    cols = ws.ncols
    result = []
    row = 0
    while row < rows:
        result.append([ws.cell_value(row, i) for i in range(0, cols)])
        row += 1
    return result


def extractTable(table, cols, cond_col, condition): #type(table)=list, type(cols)=str, type(condition)=str
    ci = colIndexes()
    result = []
    head = []
    i = table[0]
    for c in cols:
        head.append(i[ci[c]])
    result.append(head)
    
    for i in table:
        if i[ci[cond_col]] == condition:
            new_row = []
            for c in cols:
                new_row.append(i[ci[c]])

            new_row_ = []
            for i in new_row:
                try:
                    s = int(i)
                    new_row_.append(s)
                except:
                    new_row_.append(i)
            new_row = new_row_
            result.append(new_row)
        else:
            pass
    return result


if __name__ == '__main__':
    table_arr = getTable('source.xls')
    d = {}
    d['pokupka'] = extractTable(table_arr, 'aghlmr', 'l', u'Вручен')
    d['vozvrat'] = extractTable(table_arr, 'aghlmrpqt', 'p', u'Возврат')
    saveIntoWb(d)
    
    
