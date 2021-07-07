import pyvisa
import time
import locale
import openpyxl
locale.setlocale(locale.LC_ALL, 'turkish')

#AC Güç kaynağı=ac
#DC Çıkış yükü=dcload
#DC akü yükü=batload
#DC Güç kaynağı=dc

cikisvolt=str()


path="C:\\Users\\berka\\Desktop\\deneme.xlsx"
wb_obj=openpyxl.load_workbook(path) # belirtilen yerde ki excell aç
data = wb_obj.active           # aktif et


rm = pyvisa.ResourceManager()
print("USB Cihazlar:\t",rm.list_resources()) # visa usb port okuma 
usb=list(rm.list_resources())# usb port listesi

zaman=str(time.strftime('%c'))#sistem zman bilgisi okuma

seriNo=str(input("Cihaz seri nunarasını okutunuz:"))
zaman=zaman[0:9]
data["D10"]=zaman
data["B10"]=seriNo




#################################################################################    
if ('ASRL1::INSTR' in usb):#burada ac güç kaynağı bul
    print("AC KAYNAK BAĞLANDI")
    ac = rm.open_resource("ASRL1::INSTR")#belirtilen usb portu açma 
else:
    print("AC KAYNAK BULUNAMADI")
#################################################################################            
    
if ('ASRL11::INSTR' in usb):
    print("DC YÜK BAĞLANDI")#burada dc bul
    dcload=rm.open_resource("ASRL11::INSTR")
else:
    print("BATARYA YÜK BULUNAMADI")
        
#################################################################################  
        
if ('ASRL6::INSTR' in usb):
    print("BATARYA YÜK BAĞLANDI")#burada batarya yükü bul
    batload=rm.open_resource("ASRL6::INSTR")
else:
    print("BATARYA YÜK BULUNAMADI")

#################################################################################
#Çıkış gerilim regülasyon testi

try:
    ac.write("OUTP ON")
    
    ac.write("VOLT AC 100;FREQ 50")
    time.sleep(0.5)
    cikisvolt=dcload.query("MEAS:VOLT?")
    data["E3"]=cikisvolt
    
    ac.write("VOLT AC 220")
    time.sleep(0.5)
    cikisvolt=dcload.query("MEAS:VOLT?")
    data["E5"]=cikisvolt
    
    ac.write("VOLT AC 240")
    time.sleep(0.5)
    cikisvolt=dcload.query("MEAS:VOLT?")
    data["E7"]=cikisvolt
    
    dcload.write("CURR:STAT:L1 39.5")
    dcload.write("LOAD ON")
    time.sleep(0.5)
    cikisvolt=dcload.query("MEAS:VOLT?")
    data["E8"]=cikisvolt
    
    ac.write("VOLT AC 220")
    time.sleep(0.5)
    cikisvolt=dcload.query("MEAS:VOLT?")
    data["E6"]=cikisvolt
    
    ac.write("VOLT AC 100")
    time.sleep(0.5)
    cikisvolt=dcload.query("MEAS:VOLT?")
    data["E4"]=cikisvolt
    
##############################################################################◘
#Frekans çıkış ölçüm

    ac.write("VOLT AC 220;FREQ 47")
    time.sleep(0.5)
    cikisvolt=dcload.query("MEAS:VOLT?")
    data["M3"]=cikisvolt
    
    
    ac.write("VOLT AC 220;FREQ 63")
    time.sleep(0.5)
    cikisvolt=dcload.query("MEAS:VOLT?")
    data["M4"]=cikisvolt
    
    wb_obj.save(filename="C:\\Users\\berka\\Desktop\\python_ders\\%s.xlsx"%(seriNo))
    
    
    
except:
    print("Test hata!")
    data["G10"]="TEST HATASI"
    wb_obj.save(filename="C:\\Users\\berka\\Desktop\\python_ders\\%s.xlsx"%(seriNo))
        

    






 