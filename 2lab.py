import xlrd

def xls_to_txt():
    project_name = input('Введите название:')
    author_name = input('Введите имя автора:')
    f = open('result.mkb', 'w')
    f.write(project_name+'\n'+author_name+'\n\n')
    
    while True:
        try:
            f_name = input('Введите путь к файлу:')
            table = xlrd.open_workbook(f_name)
            break
        except IOError:
            print('No such file or directory')
    worksheet = table.sheet_by_index(0)

    count = 1


    for i in range(worksheet.nrows):
        row = worksheet.row_values(rowx = i, start_colx = 0, end_colx = worksheet.ncols)
        row.insert(1,',0.05')

        if i>2 and i<worksheet.nrows-1:
            for j in range(0,len(row)-1,2):
                if j<2:
                    f.write(row[0])
                    f.write(row[1])
                else:
                    print(j)
                    f.write(','+str(count)+','+str(row[j])+','+str(row[j+1]))
                    count+=1
                
            count = 1
            f.write('\n')

        elif i == 1:
            for j in range(len(row)):
                if row[j] !='' and j!=1:
                    f.write(str(row[j])+'\n')
            f.write('\n')

    f.close()
    
if __name__ == "__main__":
    xls_to_txt()