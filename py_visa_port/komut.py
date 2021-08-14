# -*- coding: utf-8 -*-
"""
Created on Sat Aug 14 09:21:02 2021

@author: User
"""

import pyvisa

dcload_id='USB0::0x0A69::0x087C::63206EL00431::INSTR'
ac_id='USB0::0x0A69::0x0883::96160900000094::INSTR'
batload_id='USB0::0x0A69::0x083E::636005'
batload_id=batload_id[0:28]        
            

rm = pyvisa.ResourceManager()
print("USB Cihazlar:\t",rm.list_resources()) # visa usb port okuma 
usb=list(rm.list_resources())# usb port listesi

liste=str()
deger=str()

for i in range(len(usb)):
    deger=usb[i]
    deger=deger[0:28]
    if deger==batload_id:
        print("deger bulundu")
        print(usb[i])
        batload_id=usb[i]
    else:
        i=+1

    
    

#################################################################################    
if (ac_id in usb):#burada ac güç kaynağı bul
    print("AC KAYNAK BAĞLANDI")
    ac = rm.open_resource(ac_id)#belirtilen usb portu açma 
else:
    print("AC KAYNAK BULUNAMADI")
#################################################################################            
    
if (dcload_id in usb):
    print("DC YÜK BAĞLANDI")#burada dc bul
    dcload=rm.open_resource(dcload_id)
else:
    print("DC YÜK BULUNAMADI")
        
#################################################################################  
        
if (batload_id in usb):
    print("BATARYA YÜK BAĞLANDI")#burada batarya yükü bul
    batload=rm.open_resource(batload_id)
else:
    print("BATARYA YÜK BULUNAMADI")

while deger != "0":
    
    deger=(input("input read: "))
    batload.write(deger)
    gelen=batload.query("CHAN?")
    print(gelen)

