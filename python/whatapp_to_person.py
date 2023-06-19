import pywhatkit
import keyboard
import xlwings as xw
 
ws = xw.Book("./excel/whatapp_to_person.xlsx").sheets['list']
arr = ws.range("A2:C10").value

def send(phone,name,message):
    try:{
    pywhatkit.sendwhatmsg_instantly(phone_no=phone,message=message + "to" + name,wait_time=10,tab_close=True)}
    finally:{
    }
for i in range(len(arr)):    
    send(arr[i][1],arr[i][0],arr[i][2])

