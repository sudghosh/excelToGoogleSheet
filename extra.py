import os
import datetime

def writeFile(MESSAGES):
    FILE = 'LOGS.txt'
    
    with open(FILE, 'a+') as file:
        MESSAGES = datetime.datetime.now().strftime('%d-%m-%Y %H-%M-%S') + " " +  MESSAGES +   '\n'
        file.write(MESSAGES)
        file.close
