from datetime import datetime

RETIREDATE = '05-04-2022'

def currentDate():
        now = datetime.now()
        current = now.strftime("%m-%d-%Y")
        return current

##print(currentDate())
