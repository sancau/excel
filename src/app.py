
from tkinter import *
from tkinter import messagebox, filedialog

from script import process_files

root = Tk()
TITLE = 'ЦНИИ РТК'
src = StringVar()
# src.set('Не задана')
# DEBUG ONLY
src.set('c:/kub')
dest = StringVar()
# dest.set('Не задана')
# DEBUG ONLY
dest.set('c:/kub_output')

def choose_source():
    value = filedialog.askdirectory()
    if value:
        src.set(value)
    
def choose_destination():
    value = filedialog.askdirectory()
    if value:
        dest.set(value)
    
def process():
    try:
        src_path = src.get()
        dest_path = dest.get()
        if not src_path or src_path == 'Не задана':
            messagebox.showinfo(TITLE, 'Не задана папка источник')
            return
        if not dest_path or dest_path == 'Не задана':
            messagebox.showinfo(TITLE, 'Не задана папка вывода')
            return
        process_files(src_path, dest_path)
        messagebox.showinfo(TITLE, 'Файлы успешно обработаны')
    except Exception as ex:     
        messagebox.showinfo(TITLE, 'Ошибка при обработке файлов: %s' % ex)

Button(root, text='Выбрать папку источник', 
    command=choose_source).place(x=10, y=10)
Label(root, textvariable=src).place(x=10, y=40)

Button(root, text='Выбрать папку назначения', 
    command=choose_destination).place(x=10, y=70)
Label(root, textvariable=dest).place(x=10, y=100)

Button(root, text='Обработка', command=process).place(x=10, y=130)

root.maxsize(300,170)
root.minsize(300,170)
root.wm_title("ЦНИИ РТК")
root.mainloop()
