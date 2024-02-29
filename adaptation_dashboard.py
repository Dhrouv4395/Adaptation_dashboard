import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import *
import pandas as pd 
from csv import DictWriter
from ttkthemes import ThemedStyle
import os


def GetValue(event):
    date_entrybox.delete(0, END)
    name_entrybox.delete(0, END)
    date_range_entrybox.delete(0, END)
    hours_entrybox.delete(0, END)
    p_details_entrybox.delete(0, END)
    row_id = listBox.selection()[0]
    select = listBox.set(row_id)
    date_entrybox.insert(0,select['Date'])
    name_entrybox.insert(0,select['Name'])
    date_range_entrybox.insert(0,select['Date_range'])
    hours_entrybox.insert(0,select['Hours'])
    machine_choose.set(select['Machine'])
    iso_choose.set(select['Iso'])
    task_choose.set(select['Task'])
    p_details_entrybox.insert(0,select['Project'])
    network_id_entrybox.insert(0, select['Network'])
    article_no_entrybox.insert(0, select['Article'])
    sap_entrybox.insert(0, select['Sap'])
    activity_entrybox.insert(0, select['Activity'])


def clear_data():
   date_entrybox.delete(0, END)
   name_entrybox.delete(0, END)
   date_range_entrybox.delete(0, END)
   hours_entrybox.delete(0, END)
   p_details_entrybox.delete(0, END)
   machine_choose.set('')
   iso_choose.set('')
   task_choose.set('')
   network_id_entrybox.delete(0, END)
   article_no_entrybox.delete(0, END)
   sap_entrybox.delete(0, END)
   activity_entrybox.delete(0, END)
   date_entrybox.focus_set()

def add_csv():
   data = [listBox.item(item,'values') for item in listBox.get_children()]
   dff = pd.DataFrame(data)
   dff.drop_duplicates(inplace=True)
   with open(r'N:\CCI\Common Use\Adaptation\adaptation.csv','w', newline='') as file:
      dff.to_csv(file, index=False)
   clear_data()
   
def add_data():
   if date_entrybox.get() == '' and name_entrybox.get()=='':
      messagebox.showinfo("info", "Can't be empty.")
   else:
      data = [date_entrybox.get(),name_entrybox.get(),str(machine_choose.get()),str(iso_choose.get()),
              str(task_choose.get()),p_details_entrybox.get(),date_range_entrybox.get(),hours_entrybox.get(),
            network_id_entrybox.get(),article_no_entrybox.get(),sap_entrybox.get(),activity_entrybox.get()]
      listBox.insert('','end',values=data)
      add_csv()
      print("Inserted successfully...")
      messagebox.showinfo("information", "Inserted successfully...")

def update_data():
   if date_entrybox.get() == '' and name_entrybox.get()=='':
      messagebox.showinfo("info", "Can't be empty.")
   else:
      selected_item = listBox.selection()
      if selected_item:
         data = [date_entrybox.get(),name_entrybox.get(),str(machine_choose.get()),str(iso_choose.get()),
              str(task_choose.get()),p_details_entrybox.get(),date_range_entrybox.get(),hours_entrybox.get(),
            network_id_entrybox.get(),article_no_entrybox.get(),sap_entrybox.get(),activity_entrybox.get()]
         listBox.item(selected_item,values=data)
         add_csv()
         print("Updated successfully...")
         messagebox.showinfo("information", "Updated successfully...")

def delete_data():
   selected_item = listBox.selection()
   if selected_item:
      listBox.delete(selected_item)
      print("Deleted successfully...")
      messagebox.showinfo("information", "Deleted successfully...")
      add_csv()

def show():
   auth = ['Date','Iso','Name','Machine','Task']##0,1,2,3,4,5
   try:
      with open(r'N:\CCI\Common Use\Adaptation\adaptation.csv','r') as file:
         df = pd.read_csv(file)#,usecols=[0,1,2,3,4,5]
         # print(df)
         for index, row in df.iterrows():
            listBox.insert('','end',values=list(row),)
   except Exception as e:
      print(e)
            

root = Tk()
root.geometry("1150x500")

Label(root, text="G+D",fg="darkblue",font=(None, 30)).place(x=950, y=5)
Label(root, text="Date").place(x=10, y=10)
Label(root, text="Name").place(x=10, y=40)
Label(root, text="Date_range").place(x=10, y=70)
Label(root, text="Hours").place(x=10, y=100)

Label(root, text="Machine").place(x=220, y=10)
Label(root, text="ISO").place(x=220, y=40)
Label(root, text="Task").place(x=220, y=70)
Label(root, text="P_details").place(x=220, y=100)

Label(root, text="Network ID").place(x=450, y=10)
Label(root, text="Article No").place(x=450, y=40)
Label(root, text="SAP").place(x=450, y=70)
Label(root, text="Activity").place(x=450, y=100)

name = tk.StringVar()

machine_choose = tk.StringVar()
machinechoosen=ttk.Combobox(root, width=15, textvariable=machine_choose)
machinechoosen['values'] = ('BPS-C4','BPS-C2/C5','GSNA')
machinechoosen.place(x=280,y=10)
machinechoosen.current(0)

iso_choose = tk.StringVar()
isochoosen=ttk.Combobox(root, width=15, textvariable=iso_choose)
isochoosen['values'] = ('ALL','DZD','AOA','ARS','AMD','AWG','AZN','AUD','BSD','BHD','BDT','BBD','BYR','BYN','BMD','BTN','BOB','BAM',
                        'BWP','BRL','BND','BGN','KHR','XAF','CAD','CVE','XAF','CLP','CNY','COP','CDF','CRC','HRK','CUC', 'CUP','ANG','CZK',
                        'DKK','DJF','DOP','EGP','XAF','ERN','EEK','ETP','EUR','GD5','GD2000','CPS','GEL','GHS','GTQ','GNF','GYD','HNL',
                        'HKD','HUF','INR','IDR','IRR','IQD','ILS','JMD','JPY','JOD','KZT','KES','KWD','LAK','LBP','LSL','LYD','MOP','MKD',
                        'MGF','MWK','MYR','MVR','MRO','MUR','MXN','MDL','MNT','MAD','MZN','MMK','NAD','NPR','XPF','NZD','NIO','NGN','NOK',
                        'OMR','PKR','USD','PGK','PYG','PEN','PHP','PLN','QAR','RON','RUB','RWF','SAR','RSD','SLL','SGD','ZAR','KRW','LKR',
                        'SDG','SRD','SZL','SEK','CHF','SYP','TWD','TJS','TZS','THB','TTD','TND','TRY','UGX','UAH','AED','GBP, GBP-SCO, GBP-NI',
                        'UZS','XOF','VEF','VND','YER','ZMK','ZWD',)
isochoosen.place(x=280,y=40)
isochoosen.current(0)

task_choose = tk.StringVar()
taskchoosen=ttk.Combobox(root, width=15, textvariable=task_choose)
taskchoosen['values'] = ('CF','ClearCase Update','Finetune','Map_Config','New Emission','Tree Update','Review')
taskchoosen.place(x=280,y=70)
taskchoosen.current(0)

p_details_entrybox = tk.Entry(root)
p_details_entrybox.place(x=280,y=100)
date_entrybox = Entry(root)
date_entrybox.place(x=75, y=10)

name_entrybox = Entry(root)
name_entrybox.place(x=75, y=40)

date_range_entrybox = Entry(root)
date_range_entrybox.place(x=75, y=70)

hours_entrybox = Entry(root)
hours_entrybox.place(x=75, y=100)

network_id_entrybox = Entry(root)
network_id_entrybox.place(x=530, y=10)

article_no_entrybox = Entry(root)
article_no_entrybox.place(x=530, y=40)

sap_entrybox = Entry(root)
sap_entrybox.place(x=530, y=70)

activity_entrybox = Entry(root)
activity_entrybox.place(x=530, y=100)

Button(root, text="Add", command=add_data, height=2, width= 13).place(x=30, y=140)
Button(root, text="update", command=update_data, height=2, width= 13).place(x=160, y=140)
Button(root, text="Delete", command=delete_data, height=2, width= 13).place(x=290, y=140)

# style = ttk.Style()
# style.theme_use('xpnative')#clam,alt,default,classic#('winnative', 'clam', 'alt', 'default', 'classic', 'vista', 'xpnative')
# print(style.theme_names())

# style.configure("mystyle.Treeview", highlightthickness=0, bd=0, font=('Times', 11)) # Modify the font of the body
# style.configure("mystyle.Treeview.Heading", font=('Times', 12,'bold')) # Modify the font of the headings
# style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})]) # Remove the borders
def openNewWindow(var):
   if var == 'new':
      newWindow = Toplevel(root)
      newWindow.geometry("300x100")
      newWindow.title("new")
      Label(newWindow, 
         text ="Coming Soon!!!.").pack()
   elif var == 'open':
      newWindow = Toplevel(root)
      newWindow.geometry("300x100")
      newWindow.title("open")
      Label(newWindow, 
         text ="Coming Soon!!!.").pack()
   else:
      newWindow = Toplevel(root)
      newWindow.geometry("300x100")
      newWindow.title("About")
      Label(newWindow, 
         text ="Dashboard Demo Work!!!").pack()


style = ThemedStyle(root)
def apply_theme(var):
   style.set_theme(var)

# MenuBar-----------------------------------------------------------------------------------------------
menu = Menu(root)
root.config(menu=menu)
filemenu = Menu(menu, tearoff=0)
themes = Menu(menu,tearoff=0)
menu.add_cascade(label='File', menu=filemenu)
menu.add_cascade(label='Themes', menu=themes)
themes.add('command',label='arc',command=lambda: apply_theme('arc'))
themes.add('command',label='elegance',command=lambda: apply_theme('elegance'))
themes.add('command',label='blue',command=lambda: apply_theme('blue'))
themes.add('command',label='breeze',command=lambda: apply_theme('breeze'))
themes.add('command',label='equilux',command=lambda: apply_theme('equilux'))
themes.add('command',label='winxpblue',command=lambda: apply_theme('winxpblue'))
themes.add('command',label='plastik',command=lambda: apply_theme('plastik'))##
themes.add('command',label='scidmint',command=lambda: apply_theme('scidmint'))##
themes.add('command',label='keramik',command=lambda: apply_theme('keramik'))##
themes.add('command',label='clam',command=lambda: apply_theme('clam'))##
themes.add('command',label='black',command=lambda: apply_theme('black'))##
themes.add('command',label='default',command=lambda: apply_theme('default'))##
themes.add('command',label='scidpurple',command=lambda: apply_theme('scidpurple'))
themes.add('command',label='ubuntu',command=lambda: apply_theme('ubuntu'))##
themes.add('command',label='smog',command=lambda: apply_theme('smog'))##
themes.add('command',label='kroc',command=lambda: apply_theme('kroc'))##
themes.add('command',label='scidgrey',command=lambda: apply_theme('scidgrey'))
themes.add('command',label='itft1',command=lambda: apply_theme('itft1'))##
themes.add('command',label='winnative',command=lambda: apply_theme('winnative'))##
themes.add('command',label='radiance',command=lambda: apply_theme('radiance'))##
themes.add('command',label='scidblue',command=lambda: apply_theme('scidblue'))
themes.add('command',label='adapta',command=lambda: apply_theme('adapta'))
themes.add('command',label='classic',command=lambda: apply_theme('classic'))
themes.add('command',label='vista',command=lambda: apply_theme('vista'))
themes.add('command',label='scidpink',command=lambda: apply_theme('scidpink'))
themes.add('command',label='yaru',command=lambda: apply_theme('yaru'))
themes.add('command',label='clearlooks',command=lambda: apply_theme('clearlooks'))
themes.add('command',label='scidgreen',command=lambda: apply_theme('scidgreen'))
themes.add('command',label='aquativo',command=lambda: apply_theme('aquativo'))

filemenu.add('command', label='New', command=lambda: openNewWindow('new'))
filemenu.add('command', label='Open...', command=lambda: openNewWindow('open'))
filemenu.add_separator()
filemenu.add_command(label='Exit', command=root.quit)
helpmenu = Menu(menu, tearoff=0)
menu.add_cascade(label='Help', menu=helpmenu)
helpmenu.add('command', label='About', command=lambda: openNewWindow('about'))
#---------------------------------------------------------------------------------------------------------

cols = ('Date','Name','Machine','Iso','Task','Project','Date_range','Hours','Network','Article','Sap','Activity')
listBox = ttk.Treeview(root,style='mystyle.Treeview', columns=cols, show='headings')
listBox.bind('<Double-Button-1>',GetValue)
listBox.tag_configure('odd', background='#E8E8E8')

for col in cols:
    listBox.heading(col, text=col)
    listBox.grid(row=1, column=1, sticky=tk.NSEW,)
    listBox.column(col,width=90)
    listBox.place(x=10, y=200)

show()

root.mainloop()