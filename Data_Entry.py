import tkinter
import openpyxl
import datetime
import tkinter.messagebox




now = datetime.datetime.now()
now_hour = float(now.strftime("%H"))
now_minute = float(now.strftime("%M"))
timeEnd = round(now_hour+(now_minute/60),2)
date = datetime.date.today()




filepath = r"C:\Users\Dell\OneDrive\Desktop\Python\data_entry.xlsx"

def remove_text(arr):
    return[element for element in arr if isinstance(element,(int,float))]

def findBalance():
    workbook = openpyxl.load_workbook(filepath)
    sheet=workbook.active

    trans_col = []
    transactions =[]

    for cell in sheet['F']:
        trans_col.append(cell.value)

    transactions = remove_text(trans_col)
    balance = sum(transactions)

    return balance


balance = findBalance()
print(balance)


def enter_data():
    #print("hulallalalalallla")
    activity = activityEntry.get()
    cost = int(costEntry.get())
    timeStart = float(timeEntry.get())
    period = timeEnd-timeStart
    change = cost*period


    workbook = openpyxl.load_workbook(filepath)
    sheet=workbook.active
    sheet.append([activity,cost,timeStart,timeEnd,period,change,date])

    workbook.save(filepath)

    new_balance = findBalance()


    tkinter.messagebox.showinfo("Ujjwal Bank of Aligarh",  f"Your data is successfully saved!\n New Balance: {new_balance}")
    



window = tkinter.Tk()
window.title('Enter Your Data below__')
window.geometry("300x200")

frame = tkinter.Frame(window)
frame.pack()

activity = tkinter.Label(frame,text='Activity:')
activity.grid(row=0,column=0)
activityEntry = tkinter.Entry(frame)
activityEntry.grid(row=0,column=1)


cost = tkinter.Label(frame,text='Cost:')
cost.grid(row=1,column=0)
costEntry = tkinter.Entry(frame)
costEntry.grid(row=1,column=1)


timeStart = tkinter.Label(frame,text='Start Time:')
timeStart.grid(row=2,column=0)
timeEntry = tkinter.Entry(frame)
timeEntry.grid(row=2,column=1)

currentBalance = tkinter.Label(frame,text="Current Balance: ")
currentBalance.grid(row=3,column=0)
blance_val = tkinter.Label(frame,text=balance)
blance_val.grid(row=3,column=1)

for widget in frame.winfo_children():
    widget.grid_configure(pady=5) 


button = tkinter.Button(frame,text='Enter Data',command=enter_data)
button.grid(row=4,column=1)

window.mainloop()