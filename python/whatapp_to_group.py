import pywhatkit
import keyboard
import xlwings as xw
 
ws = xw.Book("./excel/whatapp_to_group.xlsx").sheets['list']
arr = ws.range("A2:C10").value

def send(phone,name,mess):
    try:{
    pywhatkit.sendwhatmsg_to_group_instantly(group_id=phone, message=mess + "to" + name,wait_time=3,tab_close=True,close_time=3)
    }
    finally:{
    }
for i in range(len(arr)):    
    send(arr[i][1],arr[i][0],arr[1][2])
