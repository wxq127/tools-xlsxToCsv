import xlrd
import csv
import codecs
import os

# fromDir = os.path.join(os.getcwd(), 'excel')
fromDir = os.getcwd()
toDir = os.path.join(fromDir, '../csv')


def startConvert():
    if not os.path.exists(toDir):
        os.mkdir(toDir)
    count = 0
    for filename in os.listdir(fromDir):
        try:
            if filename.endswith('.xlsx') and not filename.startswith('~$'):
                xlsx_to_csv(os.path.join(fromDir, filename), os.path.join(
                    toDir, filename.replace('.xlsx', '.csv')))
                count += 1
                print('转换成功：{}'.format(filename))
        except Exception as e:
            print('转换{} 出错'.format(filename))
            raise e
    print('成功转换了 %d 个文件' % count)


def xlsx_to_csv(formFile, toFile):
    workbook = xlrd.open_workbook(formFile)
    table = workbook.sheet_by_index(0)
    with codecs.open(toFile, 'w', encoding='gb18030') as f:
        write = csv.writer(f)
        for row_num in range(table.nrows):
            row_value = table.row_values(row_num)
            res_row = []
            for col in row_value:
                # 42是异常空值#N/A
                if type(col) == int and col == 42:
                    col = '#N/A'
                # excel中的整数和小数都会当做float处理 如果是整数 则转换为int就不会出现xx.0的情况
                elif type(col) == float and col % 1 == 0:
                    col = '{:.0f}'.format(col)
                res_row.append(col)
            write.writerow(res_row)


if __name__ == '__main__':
    try:
        startConvert()
    except Exception as e:
        print('转换任务出错，结束！ error={}'.format(e))
