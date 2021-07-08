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

dcload_id='USB0::0x0A69::0x087C::63206EL00431::INSTR'
ac_id='USB0::0x0A69::0x0883::96160900000094::INSTR'
batload_id='USB0::0x0A69::0x083E::636005004631::INSTR'

path="C:\\Users\\User\\Desktop\\SETUP\\program\\deneme.xlsx"
wb_obj=openpyxl.load_workbook(path) # belirtilen yerde ki excell aç
data = wb_obj.active           # aktif et


rm = pyvisa.ResourceManager()
print("USB Cihazlar:\t",rm.list_resources()) # visa usb port okuma 
usb=list(rm.list_resources())# usb port listesi

zaman=str(time.strftime('%c'))#sistem zman bilgisi okuma
zaman=zaman[0:10]
data["D10"]=zaman




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
    batload.write("LOAD OFF")
    batload.write("CONF:PARA:MODE MASTER")
    batload.write("CONF:PARA:INIT ON")
    batload.write("MODE CRL")
    batload.write("CURR:STAT:L2 1")
    batload.write("LOAD ON")
else:
    print("BATARYA YÜK BULUNAMADI")

#################################################################################
seriNo=str(input("Cihaz seri nunarasını okutunuz:"))
data["B10"]=seriNo

#Çıkış gerilim regülasyon testi

try:
    ac.write("VOLT:AC 220")
    ac.write("FREQ 50")
    ac.write("OUTP ON")
    time.sleep(6)
    ac.write("VOLT:AC 175")
    ac.write("FREQ 50")
    time.sleep(2)
    cikisvolt=dcload.query("MEAS:VOLT?")
    data["E3"]=cikisvolt
    
    ac.write("VOLT:AC 220")
    time.sleep(2)
    cikisvolt=dcload.query("MEAS:VOLT?")
    data["E5"]=cikisvolt
    
    ac.write("VOLT:AC 240")
    time.sleep(2)
    cikisvolt=dcload.query("MEAS:VOLT?")
    data["E7"]=cikisvolt
    
    dcload.write("MODE CCH")
    dcload.write("CURR:STAT:L1 19.5")
    dcload.write("LOAD ON")
    time.sleep(2)
    cikisvolt=dcload.query("MEAS:VOLT?")
    data["E8"]=cikisvolt
    
    ac.write("VOLT:AC 220")
    time.sleep(2)
    cikisvolt=dcload.query("MEAS:VOLT?")
    data["E6"]=cikisvolt
    
    ac.write("VOLT:AC 175")
    time.sleep(2)
    cikisvolt=dcload.query("MEAS:VOLT?")
    data["E4"]=cikisvolt
    
##############################################################################◘
#Frekans çıkış ölçüm

    ac.write("VOLT:AC 220")
    ac.write("FREQ 47")
    time.sleep(2)
    cikisvolt=dcload.query("MEAS:VOLT?")
    data["M3"]=cikisvolt
    
    
    ac.write("FREQ 63")
    time.sleep(2)
    cikisvolt=dcload.query("MEAS:VOLT?")
    data["M4"]=cikisvolt
    ac.write("OUTP OFF")
    dcload.write("LOAD OFF")
    
    print("Test Başarılı :)")
    wb_obj.save(filename="C:\\Users\\User\\Desktop\\SETUP\\program\\%s.xlsx"%(seriNo))
    
    
    
except:
    print("Test hata!")
    data["G10"]="TEST HATASI"
    wb_obj.save(filename="C:\\Users\\User\\Desktop\\SETUP\\program\\%s.xlsx"%(seriNo))
        
