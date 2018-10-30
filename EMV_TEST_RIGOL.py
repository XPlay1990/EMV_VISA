import visa
import time
import winsound
import threading
import win32api
import win32com.client as wincl
import csv
import socket

##################################################################
# be you can use the script please install the following packages

# pip install pypiwin32
# pip install pyvisa

#################################################################
#################################################################
m_TCPIP_Connection = socket.socket(socket.AF_INET, socket.SOCK_STREAM)


def initTCPIPServer():
    # global m_TCPIP_Connection
    TCP_IP = '127.0.0.1'
    TCP_PORT = 1882
    BUFFER_SIZE = 1024
    m_TCPIP_Connection.connect((TCP_IP, TCP_PORT))
    print("Init TCP-IP:")


def sendDataTCPIPServer(actualFreq):
    # global m_TCPIP_Connection
    try:
        actualFreq = int(actualFreq)
        m_TCPIP_Connection.send(("frequency:" + str(actualFreq) + "\n").encode())
    except Exception as e:
        print(e)


def closeTCPIPServer():
    # global m_TCPIP_Connection
    m_TCPIP_Connection.close()


def turnoff():
    print(inst.write(":OUTP OFF"))
    sendDataTCPIPServer(0)


def MyThread1():
    winsound.Beep(beepFrequency, duration)


def setup():
    print(inst.write("*IDN?"))
    print(inst.write(":SYST:PRES:TYPE FAC"))
    print(inst.write(":SYST:PRES"))
    print(inst.write(":LEV 0.018V"))  # carrier amplitude [dBm]
    # Levels for the Power Amplifier Hubert A1200
    #  ":LEV 0.006V" ->1VRMS
    #  ":LEV 0.018V" ->3VRMS
    #  ":LEV 0.030V" ->5VRMS
    #  ":LEV 0.048V" ->8VRMS
    #  ":LEV 0.060V" ->10VRMS
    print(inst.write(":AM:DEPT 80"))  # modulation depth
    print(inst.write(":AM:FREQ " + str(1000)))
    print(inst.write(":AM:STAT ON"))
    print(inst.write(":MOD:STAT ON"))
    print("Setup Complete!")


def outputOn():
    print(inst.write(":FREQ " + str(actualFreq)))
    print(inst.write(":OUTP ON"))
    sendDataTCPIPServer(actualFreq)
    print(actualFreq)
    # timestamp = time.strftime("%b-%d-%Y_Time_%H-%M-%S", time.localtime())
    timestamp = time.strftime("%H-%M-%S", time.localtime())
    csvfile.write(str(timestamp) + ";" + str(float(actualFreq)).replace('.', ',') + "\n")
    t1 = threading.Thread(target=MyThread1, args=[])
    t1.start()


def outputOff():
    print(inst.write(":OUTP OFF"))
    sendDataTCPIPServer(0)


speak = wincl.Dispatch("SAPI.SpVoice")
speak.Speak("Micha, USB Stecker ab")
restTime = 1
dwellTime = 2

beepFrequency = 1200  # Set Frequency To 2500 Hertz
duration = (dwellTime * 1000)  # Set Duration To 1000 ms == 1 second

actualFreq = 150000  # 1800000 ## 150000
endFreq = 230000000
percentage = 0.01  # 0.01

rm = visa.ResourceManager()
inst = rm.open_resource('USB0::0x1AB1::0x099C::DSG8A193900193::INSTR')

try:

    setup()

    timestamp = time.strftime("%b-%d-%Y_Time_%H-%M-%S", time.localtime())
    filename = "RIGOL_Log_" + timestamp + ".csv"

    csvfile = open(filename, 'w', 1)
    ## fieldnames = ['Date', 'Freqeuncy_Value']
    ## writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    ## writer.writeheader()

    initTCPIPServer()

    while True:
        outputOn()
        time.sleep(dwellTime)
        outputOff()
        actualFreq = actualFreq * (percentage + 1)
        # actualFreq = int(actualFreq)
        time.sleep(restTime)
        if (actualFreq >= endFreq):
            break

except KeyboardInterrupt:
    print("W: interrupt received, stopping…")
finally:
    # clean up
    turnoff()
    csvfile.close()
    closeTCPIPServer()
    exit()









