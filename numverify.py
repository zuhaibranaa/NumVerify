# Import Libraries
from tkinter import END
import pandas as pd
try:
    import tkinter as tk
    import tkinter.ttk as ttk
    from tkinter import filedialog
except ImportError:
    import Tkinter as tk
    import ttk
    import tkFileDialog as filedialog
from tkfilebrowser import askopendirname, askopenfilenames, asksaveasfilename
import requests
# Crafting UI

# Crafting Window
window = tk.Tk()
window.title('Num Verify')
window.resizable(width=False, height=False)
window.configure(background = 'White')
window.geometry('300x530+450+90')
# Crafting Frames
frame_GREETINGS = tk.Frame(master=window, borderwidth=2)
frame_GET_API_KEY = tk.Frame(master=window,width=300, relief=tk.RAISED, borderwidth=4)
frame_GET_EXCEL_FILE = tk.Frame(master=window,width=300, relief=tk.RAISED, borderwidth=4)
frame_GET_COLUMN_NAME = tk.Frame(master=window,width=300, relief=tk.RAISED, borderwidth=4)
frame_GET_FINAL_FILE = tk.Frame(master=window,width=300, relief=tk.RAISED, borderwidth=4)
frame_PROCESS = tk.Frame()
# Crafting Components For frame_GREETINGS
welcome = tk.Label(master=frame_GREETINGS,text='Welcome To Numverify\nGUI Based Verification Interface',font=20, height=3, width=300,bg='blue',fg='yellow')
welcome.pack()
# Crafting Components For frame_GET_API_KEY
api_label = tk.Label(master=frame_GET_API_KEY, width=300, text='Enter Your Api Key :', bg='gray', fg='yellow', borderwidth=1)
api_label.pack()
ent_api = tk.Entry(master=frame_GET_API_KEY, width=300, relief=tk.SUNKEN ,bg='white')
ent_api.pack()
# Crafting Components For frame_GET_COLUMN_NAME
column_label = tk.Label(master=frame_GET_COLUMN_NAME, width=300, text='Enter Column Name :', bg='gray', fg='yellow', borderwidth=1)
column_label.pack()
ent_column = tk.Entry(master=frame_GET_COLUMN_NAME, width=300, relief=tk.SUNKEN ,bg='white')
ent_column.insert(END,'PhoneNumber')
ent_column.pack()
# Crafting Components For frame_GET_EXCEL_FILE
ent1 = tk.Entry(master=frame_GET_EXCEL_FILE, width=300, bg='white', relief=tk.SUNKEN)
# Browse Function
def browsefunc():
    filename =tk.filedialog.askopenfilename(filetypes=(("excell files","*.xlsx"),("All files","*.*")))
    ent1.insert(tk.END, filename)
# Browse Button
btn = tk.Button(master=frame_GET_EXCEL_FILE, text='Browse Excel File', width=300, bg='green', fg='white',command=browsefunc)
btn.pack()
ent1.pack()
# Crafting Components For frame_GET_FINAL_FILE
final_label = tk.Label(master=frame_GET_FINAL_FILE, width=300, text='Enter Output File Name :\n Output Name Will Be (input-output.xlsx) ', bg='gray', fg='yellow', borderwidth=1)
final_label.pack()
ent_final = tk.Entry(master=frame_GET_FINAL_FILE, width=300, bg='white', relief=tk.SUNKEN)
ent_final.pack()
# Crafting Components For frame_PROCESS
def process_data():
    fil = ent1.get()
    index= (fil.find('.'))
    output = fil[:index]
    access_key = ent_api.get()
    final_file_name = ent_final.get()
    print('Please Wait Data is Being processed')
    dataset = pd.read_excel(fil,engine='openpyxl')
    numbers = dataset[ent_column.get()].tolist()
    validity = [requests.get(f'http://apilayer.net/api/validate?access_key={access_key}&number={num}',headers={'Content-Type': 'application/json'}) for num in numbers]
    valid = []
    number = []
    location = []
    international_format = []
    local_format = []
    country_name = []
    carrier = []
    country_code = []
    line_type = []
    country_prefix = []
    for k in validity:
        for i in range(10):
            if i == 0:
                valid.append(k.json()['valid'])
            elif i == 1:
                number.append(k.json()['number'])
            elif i == 2:
                location.append(k.json()['location'])
            elif i == 3:
                international_format.append(k.json()['international_format'])
            elif i == 4:
                local_format.append(k.json()['local_format'])
            elif i == 5:
                country_name.append(k.json()['country_name'])
            elif i == 6:
                carrier.append(k.json()['carrier'])
            elif i == 7:
                country_code.append(k.json()['country_code'])
            elif i == 8:
                line_type.append(k.json()['line_type'])
            else:
                country_prefix.append(k.json()['country_prefix'])
    dataset['valid'] = valid
    dataset['local_format'] = local_format
    dataset['international_format'] = international_format
    dataset['country_prefix'] = country_prefix
    dataset['country_code'] = country_code
    dataset['country_name'] = country_name
    dataset['location'] = location
    dataset['carrier'] = valid
    dataset['line_type'] = line_type
    dataset.to_excel(f'{output}-{final_file_name}.xlsx')
    tk.messagebox.showinfo(title='Success',message=f'Data Stored as {output}-{final_file_name}.xlsx')
submit = tk.Button(master=frame_PROCESS, text='Process Data', width=300, bg='green', fg='white', command=process_data)
submit.pack()
# Packing Frames
frame_GREETINGS.pack()   
frame_GET_API_KEY.pack(padx=2, pady=5)
frame_GET_FINAL_FILE.pack(padx=2, pady=5)
frame_GET_COLUMN_NAME.pack(padx=2, pady=5)
frame_GET_EXCEL_FILE.pack(padx=2, pady=5)
frame_PROCESS.pack(padx=2, pady=(150,10))

# Crafting Functions
   
window.mainloop()
