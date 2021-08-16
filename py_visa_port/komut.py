# -*- coding: utf-8 -*-
"""
Created on Sat Aug 14 09:21:02 2021

@author: User
"""

import pyvisa
import time
import locale
import openpyxl
locale.setlocale(locale.LC_ALL, 'turkish')

dcload_id='USB0::0x0A69::0x087C::63206'
ac_id='USB0::0x0A69::0x0883::961609'
batload_id='USB0::0x0A69::0x083E::636005'


#AC Güç kaynağı=ac
#DC Çıkış yükü=dcload
#DC akü yükü=batload
#DC Güç kaynağı=dc



# exel dosyası aç
path="C:\\forms\\0K7_V1.xlsx"
wb_obj=openpyxl.load_workbook(path) # belirtilen yerde ki excell aç
data = wb_obj.active           # aktif et


#zaman oku ve kaydet
zaman=str(time.strftime('%c'))#sistem zman bilgisi okuma
zaman=zaman[0:10]
data["D10"]=zaman
 
      

rm = pyvisa.ResourceManager()
print("USB Cihazlar:\t",rm.list_resources()) # visa usb port okuma 
usb=list(rm.list_resources())# usb port listesi

deger=str()
konum=str()


for a in range(len(usb)):
    deger=usb[a]
    deger=deger[0:28]
    if deger==batload_id:
        print("Batarya yük bulundu")
        print(usb[a])
        batload_id=usb[a]
        a=0
    else:
        a=+1
        
for  b in range(len(usb)):
    deger=usb[b]
    deger=deger[:27]
    if deger==dcload_id:
        print("Yük bulundu")
        print(usb[b])
        dcload_id=usb[b]
        b=0
    else:
        b=+1
        
for c in range(len(usb)):
    deger=usb[c]
    deger=deger[:28]
    if deger==ac_id:
        print("AC Kaynak bulundu")
        print(usb[c])
        ac_id=usb[c]
        c=0
    else:
        c=+1


    

#################################################################################    
if (ac_id in usb):#burada ac güç kaynağı bul
    print("AC KAYNAK BAĞLANDI")
    ac = rm.open_resource(ac_id)#belirtilen usb portu açma 
    gate1=True
else:
    print("AC KAYNAK BULUNAMADI")
    gate1=False
#################################################################################            
    
if (dcload_id in usb):
    print("DC YÜK BAĞLANDI")#burada dc bul
    dcload=rm.open_resource(dcload_id)
    gate2=True
else:
    print("DC YÜK BULUNAMADI")
    gate2=False
        
#################################################################################  
        
if (batload_id in usb):
    print("BATARYA YÜK BAĞLANDI")#burada batarya yükü bul
    batload=rm.open_resource(batload_id)
    gate3=True
else:
    print("BATARYA YÜK BULUNAMADI")
    gate3=False


def giris_regule():
    try:
        ac.write("VOLT:AC 220")
        ac.write("FREQ 50")
        ac.write("OUTP ON")
        time.sleep(6)
        
        dcload.write("MODE CCH")
        dcload.write("CURR:STAT:L1 5")
        dcload.write("LOAD ON")
        time.sleep(2)

        ac.write("VOLT:AC 176")
        time.sleep(2)
        cikisvolt=dcload.query("MEAS:VOLT?")
        cikisvolt=cikisvolt[:4]
        data["F3"]=("{}V".format(cikisvolt))
        cikisvolt=float(cikisvolt)
        
        if (cikisvolt > 27.5 and cikisvolt < 28.5):
            data["G3"]="√"
            print("Giriş Volt:176V Çıkış Volt:",cikisvolt)
        else:
            data["G3"]="X"
            print("HATA    Giriş Volt:176V Çıkış Volt:",cikisvolt)
        
        
        
        
        
        ac.write("VOLT:AC 220")
        time.sleep(2)
        cikisvolt=dcload.query("MEAS:VOLT?")
        cikisvolt=cikisvolt[:4]
        data["F4"]=("{}V".format(cikisvolt))
        cikisvolt=float(cikisvolt)
        
        if (cikisvolt > 27.5 and cikisvolt < 28.5):
            data["G4"]="√"
            print("Giriş Volt:220 Çıkış volt:",cikisvolt)
        else:
            data["G4"]="X"
            print("HATA    Giriş Volt:220V Çıkış Volt:",cikisvolt)
        
        
        
        
        ac.write("VOLT:AC 265")
        time.sleep(2)
        cikisvolt=dcload.query("MEAS:VOLT?")
        cikisvolt=cikisvolt[:4]
        data["F5"]=("{}V".format(cikisvolt))
        cikisvolt=float(cikisvolt)
        
        if (cikisvolt > 27.5 and cikisvolt < 28.5):
            data["G5"]="√"
            print("Giriş Volt:265V Çıkış volt:",cikisvolt)
        else:
            data["G5"]="X"
            print("HATA    Giriş Volt:265V Çıkış Volt:",cikisvolt)
            
            
            
            
            
        ac.write("VOLT:AC 300")
        time.sleep(2)
        cikisvolt=dcload.query("MEAS:VOLT?")
        cikisvolt=cikisvolt[:4]
        data["F7"]=("{}V".format(cikisvolt))
        cikisvolt=float(cikisvolt)
        
        if (cikisvolt < 1.0):
            data["G7"]="√"
            print("Giriş Volt:300V Çıkış volt:",cikisvolt)
        else:
            data["G7"]="X"
            print("HATA    Giriş Volt:300V Çıkış Volt:",cikisvolt)
            
        
        ac.write("VOLT:AC 165")
        time.sleep(2)
        cikisvolt=dcload.query("MEAS:VOLT?")
        cikisvolt=cikisvolt[:4]
        data["F6"]=("{}V".format(cikisvolt))
        cikisvolt=float(cikisvolt)
        
        if (cikisvolt < 1.0):
            data["G6"]="√"
            print("Giriş Volt:165V Çıkış volt:",cikisvolt)
        else:
            data["G6"]="X"
            print("HATA    Giriş Volt:165V Çıkış Volt:",cikisvolt)
            
            
            
            
        ac.write("VOLT:AC 220")
        ac.write("FREQ 47")
        time.sleep(2)
        cikisvolt=dcload.query("MEAS:VOLT?")
        cikisvolt=cikisvolt[:4]
        data["F8"]=("{}V".format(cikisvolt))
        cikisvolt=float(cikisvolt)
        
        if (cikisvolt > 27.5 and cikisvolt < 28.5):
            data["G8"]="√"
            print("Giriş Frekans:47Hz Çıkış volt:",cikisvolt)
        else:
            data["G8"]="X"
            print("HATA    Giriş Frekans:47Hz Çıkış Volt:",cikisvolt)
        
            
        
        ac.write("VOLT:AC 220")
        ac.write("FREQ 63")
        time.sleep(2)
        cikisvolt=dcload.query("MEAS:VOLT?")
        cikisvolt=cikisvolt[:4]
        data["F9"]=("{}V".format(cikisvolt))
        cikisvolt=float(cikisvolt)
        
        if (cikisvolt > 27.5 and cikisvolt < 28.5):
            data["G9"]="√"
            print("Giriş Frekans:63Hz Çıkış volt:",cikisvolt)
        else:
            data["G9"]="X"
            print("HATA!    Giriş Frekans:63Hz Çıkış Volt:",cikisvolt)  
            
            
        print("Test Başarılı :)")
            
    except:
        print("Bağlantı hatası!")



def verim_olc():
    
    ac.write("VOLT:AC 220")
    ac.write("FREQ 50")
    time.sleep(2)
    verim=ac.query("MEAS:POW:AC:TOT?")
    cikisvolt=dcload.query("MEAS:VOLT?")
    verim=float(verim)
    cikisvolt=float(cikisvolt)
    
    try:
        verim=((verim/(cikisvolt*19.5))*100)
        if verim>=90.0:
             verim=str(verim)
             verim=verim[:4]
             data["E13"]=("%{}".format(verim))
             data["F13"]="√"
             print("Verim:",verim)
        else:
             verim=str(verim)
             verim=verim[:4]
             data["E13"]=("%{}".format(verim))
             print("Verim Düşük!")
             data["F13"]="X"
    except:
        print("Verim hesaplama hatası!")
     
    try:
        
        dcload.write("LOAD OFF")
        time.sleep(1)
        bosverim=ac.query("MEAS:POW:AC:TOT?")
        bosverim=float(bosverim)
    
        if bosverim <=30.0:
            bosverim=str(bosverim)
            bosverim=bosverim[:4]
            data["E14"]=("{}W".format(bosverim))
            data["F14"]="√"
            print("Giriş Aktif:",bosverim)
        else:
            bosverim=str(bosverim)
            bosverim=bosverim[:4]
            data["E14"]=("{}W".format(bosverim))
            data["F14"]="X"
            print("Giriş Aktif yüksek:")
                  
    except:
        print("Giriş aktif hesaplama hatası!")
        
def asiri_yuk():
    
    
    dcload.write("LOAD ON")
    time.sleep(1)
    cikisvolt=dcload.query("MEAS:VOLT?")
    cikisvolt=float(cikisvolt)
    
    if cikisvolt <= 28.5 and cikisvolt >=27.5:
        data["E18"]="√"
        print("Çıkış akım:19.5A")
    else:
        data["E18"]="X"
        print("Cıkış akım hata!")
       
    dcload.write("VOLT:L1 16")
    dcload.write("MODE CVH")
    time.sleep(2)
    cikisvolt=dcload.query("MEAS:VOLT?")
    cikisvolt=float(cikisvolt)
    if cikisvolt<1.0:
        data["E19"]="√"
        print("Hiccup Modu")
    else:
        data["E19"]="X"
        print("Hiccup Modu HATA!")
    
    
          
    
    
deger=(input("Seri NO: "))
konum="C:\\Users\\User\\Desktop\\pytest_master\\"
if (gate1 and gate2)==True:
    giris_regule()
    verim_olc()
    asiri_yuk()
    ac.write("OUTP OFF")
    dcload.write("LOAD OFF")
    wb_obj.save(filename="{konu}{dege}.xlsx".format(konu=konum, dege=deger))
else:
    print("bağlantı yok")


