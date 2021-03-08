def readLog():
	f = open("data_log.txt", "r")
	data = f.read()
	f.close()
	return data

def writeLog(data):
	f = open("data_log.txt", "w")
	f.write(data)
	f.close()
