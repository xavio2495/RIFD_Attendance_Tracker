import gspread
import serial
import winsound as w
from datetime import datetime
try:
    arduino = serial.Serial("COM4",timeout=0) #checks if connections are good
    print("Serial port connected & RC522 is working")
except:
    print('Please check the port or RC522')
    exit()
#value from arduino -- passed from rfid reader RC522
while True:
    inputx=""
    x=""
    count=0
    while count!=1:
            x=arduino.readline() #getting data from serial port COM5
            y=str(x)
            if len(y)==16:
                    inputx=str(y[2:-3]) #getting only input with values and skipping empty ones
                    print("True value from port:",inputx) 
                    count=1


    #w.PlaySound("digitone",w.SND_FILENAME)
    w.Beep(1000,500)
    print("Data read, Moving on to marking...")

    #Accessing the google sheet
    sa = gspread.service_account(filename="datajson3.json")
    sh = sa.open("Data")
    wks = sh.worksheet("Sheet")

    #Date from system and from Gsheet
    date_s = str(wks.acell('A1').value)
    date_rt = datetime.now().strftime("%d/%m/%Y")
    print("time check done.")
    # to find location of the input in sheet
    val=wks.get('A3:A11')
    ind=val.index([inputx]) + 3
    s= "C" +str(ind)
    t= "D" +str(ind)
    print(t)
    
    #On next day update date and reset attendance
    if date_rt != date_s:                          #-----executed on first input of the day
        wks.update('A1',[[date_rt]])
        for x in range(9):
            z='D' + str(x+3)
            y='C' + str(x+3)
            wks.update(y,"") 
            wks.update(z,"") 
        wks.update(s,[["Present"]]) #mark attendance of the person and log timestamp
        wks.update(t,[[str(datetime.now().strftime("%H:%M:%S"))]]) 
    else:
         #-----marking of attendance executed normally
        wks.update(t,[[str(datetime.now().strftime("%H:%M:%S"))]]) 
        wks.update(s,[[str("Present")]])
        
    print("User registered... \n\n waiting for next...")
