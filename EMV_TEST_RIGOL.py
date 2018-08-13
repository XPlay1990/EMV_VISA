import visa
import time
import winsound
import threading
import win32api
##################################################################
# be you can use the script please install the following packages

# pip install pypiwin32
# pip install pyvisa

#################################################################
#################################################################

    
def turnoff():
    print(inst.write(":OUTP OFF"))

def MyThread1():
    winsound.Beep(beepFrequency, duration)

def setup():
    print(inst.write("*IDN?"))
    print(inst.write(":SYST:PRES:TYPE FAC"))
    print(inst.write(":SYST:PRES"))
    print(inst.write(":LEV 0.018V")) #carrier amplitude [dBm]
                                     # Levels for the Power Amplifier Hubert A1200
                                     #  ":LEV 0.006V" ->1VRMS
                                     #  ":LEV 0.018V" ->3VRMS
                                     #  ":LEV 0.030V" ->5VRMS
                                     #  ":LEV 0.048V" ->8VRMS
                                     #  ":LEV 0.060V" ->10VRMS
    print(inst.write(":AM:DEPT 80")) #modulation depth
    print(inst.write(":AM:FREQ " + str(1000)))
    print(inst.write(":AM:STAT ON"))
    print(inst.write(":MOD:STAT ON"))
    print("Setup Complete!")


restTime = 1
dwellTime = 2

beepFrequency = 1200  # Set Frequency To 2500 Hertz
duration = (dwellTime*1000)  # Set Duration To 1000 ms == 1 second

actualFreq = 6000000
endFreq = 230000000
percentage = 0.01 #0.01

rm = visa.ResourceManager()
inst = rm.open_resource('USB0::0x1AB1::0x099C::DSG8A193900193::INSTR')

try:

    setup()

    while True:
         print(inst.write(":FREQ "+ str(actualFreq)))
         print(inst.write(":OUTP ON"))
         t1 = threading.Thread(target=MyThread1, args=[])
         t1.start()
         time.sleep(dwellTime)
         print(inst.write(":OUTP OFF"))
         
         actualFreq = actualFreq * (percentage + 1)
         print(actualFreq)
         
         time.sleep(restTime)
         if(actualFreq >= endFreq):
            break
        
except KeyboardInterrupt:
    print("W: interrupt received, stopping…")
finally:
# clean up
    turnoff()
    exit()