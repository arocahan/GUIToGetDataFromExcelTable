import openpyxl
from tkinter import *

root = Tk()
root.title('Rechnungen')
root.geometry('850x500')

#--------------------------------------------Variables-----------------------------------------------
apothekeArray = []

#--------------------------------------------excelFile (Table)---------------------------------------------
excelFile = openpyxl.load_workbook(r'C:\Users\Arael Roca Hanson\Desktop\Pharmy_plan.xlsx', data_only=True)
worksheet = excelFile['MJJA']

#--------------------------------------------GUI - list of Apotheke mit Aufträge im X Monat-----------------------------------------------

frame1 = LabelFrame(root, text = 'Date Bsp. 2021-07 ->', padx= 15, pady= 15)
frame1.grid(row = 1, column = 1)

frame = LabelFrame(root, text= "Apotheke eingeben", padx= 15, pady= 15)
frame.grid(column=2, row=1)

dateEntry = Entry(frame1, width= 40)
dateEntry.grid(row = 1, column = 1)

ApoEntry = Entry(frame, width= 40)
ApoEntry.grid(row = 1, column = 1)

my_text = Text(root, width=50, height=15, padx= 5, pady= 5)
my_text.grid(row = 2, column= 1)

my_text_Daten = Text(root, width=50, height=15, padx= 5, pady= 5)
my_text_Daten.grid(row = 2, column= 2)
 

def passDateValue():
    zeitDau = dateEntry.get()
    for i in range(1, 1000):
        s = worksheet['A' + str(i)].value
        ss = str(s)
    #print(ss)
        if zeitDau in ss:
            apotheke = worksheet['C' + str(i)].value
            if apotheke not in apothekeArray and apotheke != None:
                apothekeArray.append(apotheke)
    my_text.delete(1.0 ,END)
    for x in apothekeArray:
    
        my_text.insert(END, x + '\n')

def Stundenhonorar():
    zeitDau = dateEntry.get()
    apoName = ApoEntry.get()
    totalStunde = 0
    my_text_Daten.delete(1.0 ,END)
    for i in range(1, 1000):
        s = worksheet['A' + str(i)].value
        ss = str(s)
        #print(ss)
        if zeitDau in ss:
            apo = worksheet['C' + str(i)].value
            datum = worksheet['A' + str(i)].value
            von = worksheet['D' + str(i)].value
            bis = worksheet['E' + str(i)].value
            stunden = worksheet['F' + str(i)].value
            if apo == apoName:
                my_text_Daten.insert(END, str(datum)[0:10] + ': ' + str(von)[0:5] + ' - ' + str(bis)[0:5] + ' Uhr (' + str(stunden) + 'h)\n')
                totalStunde += stunden 
    my_text_Daten.insert(END, str(totalStunde) + '\n') 
    my_text_Daten.insert(END, str(totalStunde * 120) + '\n')    

def Fahrtzeit():
    zeitDau = dateEntry.get()
    apoName = ApoEntry.get()
    my_text_Daten.delete(1.0 ,END)
    totalFahrtzeit = 0
    for i in range(2, 1000):
        s = worksheet['A' + str(i)].value
        ss = str(s)
        if zeitDau in ss:
            apo = worksheet['C' + str(i)].value
            datum = worksheet['A' + str(i)].value
            stunden = worksheet['F' + str(i)].value
            Fahrtzeit = worksheet['G' + str(i)].value
            if Fahrtzeit != 0 and Fahrtzeit != None:
                FahrtzeitBez = float(Fahrtzeit) - 1.0

            if apo == apoName and FahrtzeitBez > 0:
                my_text_Daten.insert(END, str(datum)[0:10]+ ': ' + str(Fahrtzeit) + 'h, davon ' + str(FahrtzeitBez) + 'h bezahlt \n')
                totalFahrtzeit += FahrtzeitBez 
    my_text_Daten.insert(END, str(totalFahrtzeit) + '\n') 
    my_text_Daten.insert(END, str(totalFahrtzeit * 120) + '\n') 

def Km():
    zeitDau = dateEntry.get()
    apoName = ApoEntry.get()
    my_text_Daten.delete(1.0 ,END)
    totalKm = 0
    for i in range(1, 1000):
        s = worksheet['A' + str(i)].value
        ss = str(s)
    
        if zeitDau in ss:
            apo = worksheet['C' + str(i)].value
            datum = worksheet['A' + str(i)].value
            von = worksheet['D' + str(i)].value
            bis = worksheet['E' + str(i)].value
            stunden = worksheet['F' + str(i)].value
            km = worksheet['J' + str(i)].value
    
            if apo == apoName and (km != None and km != 0):
                my_text_Daten.insert(END, str(datum)[0:10]+ ': ' + str(km) + ' km \n')
                totalKm += km 
    my_text_Daten.insert(END, str(totalKm) + '\n') 
    my_text_Daten.insert(END, str(totalKm * 0.7) + '\n')  

def Ticket():
    zeitDau = dateEntry.get()
    apoName = ApoEntry.get()
    my_text_Daten.delete(1.0 ,END)
    totalTicket = 0
    for i in range(1, 1000):
        s = worksheet['A' + str(i)].value
        ss = str(s)
        if zeitDau in ss:
            apo = worksheet['C' + str(i)].value
            datum = worksheet['A' + str(i)].value
            ticket = worksheet['I' + str(i)].value
    
            if apo == apoName and (ticket != 0 and ticket != None):
                my_text_Daten.insert(END, str(datum)[0:10]+ ': ' + str(ticket) + ' CHF \n')
                totalTicket += ticket 
    my_text_Daten.insert(END, str(totalTicket) +'\n') 

def Parkgebuehr():
    zeitDau = dateEntry.get()
    apoName = ApoEntry.get()
    my_text_Daten.delete(1.0 ,END)
    totalParkgebuehr = 0
    for i in range(1, 1000):
        s = worksheet['A' + str(i)].value
        ss = str(s)
        if zeitDau in ss:
            apo = worksheet['C' + str(i)].value
            datum = worksheet['A' + str(i)].value
            parkticket = worksheet['L' + str(i)].value
            if apo == apoName and (parkticket != 0 and parkticket != None):
                my_text_Daten.insert(END, str(datum)[0:10]+ ': ' + str(parkticket) + ' CHF\n')
                totalParkgebuehr += parkticket 
    my_text_Daten.insert(END, str(totalParkgebuehr) + '\n') 

def Verpflegungskosten():
    zeitDau = dateEntry.get()
    apoName = ApoEntry.get()
    my_text_Daten.delete(1.0 ,END)
    totalUebernachtung = 0
    for i in range(1, 1000):
        s = worksheet['A' + str(i)].value
        ss = str(s)
        if zeitDau in ss:
            apo = worksheet['C' + str(i)].value
            datum = worksheet['A' + str(i)].value
            uberNach = worksheet['M' + str(i)].value
    
            if apo == apoName and (uberNach != 0 and uberNach != None):
                my_text_Daten.insert(END, str(datum)[0:10]+ ': ' + str(uberNach) + ' CHF \n') 
                totalUebernachtung += uberNach 
    my_text_Daten.insert(END, str(totalUebernachtung))


button = Button(frame1, text='Generate Apo', width=20, command = passDateValue).grid(row = 2, column=1)

Stundenhonorar = Button(frame, text='Stundenhonorar', width=20, command = Stundenhonorar)
Stundenhonorar.grid(row = 2, column = 1)

Fahrtzeit = Button(frame, text='Fahrtzeit', width=20, command = Fahrtzeit)
Fahrtzeit.grid(row = 3, column = 1)

Km = Button(frame, text='Km', width=20, command = Km)
Km.grid(row = 4, column = 1)

Ticket = Button(frame, text='Ticket', width=20, command = Ticket)
Ticket.grid(row = 5, column = 1)

Parkgebuehr = Button(frame, text='Parkgebuehr', width=20, command = Parkgebuehr)
Parkgebuehr.grid(row = 6, column = 1)

Verpflegungskosten = Button(frame, text='Verpflegungskosten', width=20, command = Verpflegungskosten)
Verpflegungskosten.grid(row = 7, column = 1)


root.mainloop()

