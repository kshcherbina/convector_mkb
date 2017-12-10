import xlrd
from tkinter import *

def xls_to_txt(project_name,author_name,f_name):
    
    #project_name = input('Введите название:')
    #author_name = input('Введите имя автора:')
    f = open('result.mkb', 'w')
    f.write('База данных о '+project_name+'.'+'\n'+'Автор: '+author_name+'.'+'\n\n')
    
    
    try:
            
        table = xlrd.open_workbook(f_name) #открытие файла
        
    except IOError:
        print('No such file or directory')
        return 0
    worksheet = table.sheet_by_index(0) 

    count = 1


    for i in range(worksheet.nrows):
        row = worksheet.row_values(rowx = i, start_colx = 0, end_colx = worksheet.ncols)
        row.insert(1,',0.05')

        if i>2 and i<worksheet.nrows-1:
            for j in range(0,len(row)-1,2):
                #записываем название и коэффициенты
                if j<2:
                    f.write(row[0])
                    f.write(row[1])
                #записываем ответы
                else:
                    
                    f.write(','+str(count)+','+str(row[j])+','+str(row[j+1]))
                    count+=1
                
            count = 1
            f.write('\n')
        #Запись вопросов 
        elif i == 1:
            for j in range(len(row)):
                if row[j] !='' and j!=1:
                    f.write(str(row[j])+'\n')
            f.write('\n')

    f.close()
    print('Success')

def GUI():
    root = Tk()
    root.title(u'Convector in mkb')
    label = Label(root, text = 'База данных о',width=20,bd=3)
    ent = Entry(root,width=20,bd=3)
    label1 = Label(root, text = 'Автор',width=20,bd=3)
    ent1 = Entry(root,width=20,bd=3)
    label2 = Label(root, text = 'Путь к файлу',width=20,bd=3)
    ent2 = Entry(root,width=20,bd=3)

    label.pack()
    ent.pack()
    label1.pack()
    ent1.pack()
    label2.pack()
    ent2.pack()

    but = Button(root,
          text="Нажать", #надпись на кнопке
          width=30,height=5, #ширина и высота
          bg="white",fg="blue") #цвет фона и надписи

    but.bind('<Button-1>', lambda event: xls_to_txt(ent.get(),ent1.get(),ent2.get()))



    but.pack()
    
    root.mainloop()
    

if __name__ == "__main__":
    GUI()