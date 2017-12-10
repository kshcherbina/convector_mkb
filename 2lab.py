
import xlrd

def xls_to_txt():
    project_name = input('Введите название:')
    author_name = input('Введите имя автора:')
    f = open('result.mkb', 'w')
    f.write(project_name+'\n'+author_name+'\n')
        
    while True:
        try:
            f_name = input('Введите путь к файлу:')
            table = xlrd.open_workbook(f_name)
            break
        except IOError:
            print('No such file or directory')
    worksheet = table.sheet_by_index(0)

    for i in range(worksheet.nrows):
        row = worksheet.row_values(rowx = i, start_colx = 0, end_colx = worksheet.ncols)
        for j in row:
            if j!='':
                if i == 1:
                    f.write(str(j)+'\n')
                elif i>2 and i<worksheet.nrows-1:
        
                    f.write(str(j)+',')
        if i != 1:
            f.write('\n')
    f.close()

if __name__ == "__main__":
    xls_to_txt()