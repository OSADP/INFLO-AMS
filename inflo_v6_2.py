import win32com.client as com
import os
import time
from random import randrange
import random


def InitializeINFLO(i):
	realTime = i / Sim_Resolution - INFLOFreq
	DCinfo = ReadDataCollections(realTime)
	CVData = FetchCVData()
	#print CVData	
	WriteData(DCinfo, CVData)
	time.sleep(INFLOSleep)
	SpeedData = GetSpeed()
	if len(SpeedData[1]) >= 80:
		SpeedDecMatrix = SpeedDecision(SpeedData)
		UpdateSD(SpeedDecMatrix)

def FetchCVData(fields = ('No','SimSec','VehType','Speed','Lane','Pos')):
	CVData = []
	linklist = [514,515,516,517,523,524,527,531,226,216,214,107,106,105,90,83]
	all_veh_attributes = Vissim.Net.Vehicles.GetMultipleAttributes(fields)
	#print all_veh_attributes
	for cnt in range(len(all_veh_attributes)):
		#link in linklist, vehicle type in CV, 
		if float(all_veh_attributes[cnt][4].split('-')[0]) in linklist and all_veh_attributes[cnt][2]=='1':
			vehList = [0,0,0,0,0,0]
			vehList[0] = all_veh_attributes[cnt][0] #ID
			vehList[1] = all_veh_attributes[cnt][3]	#speed
			vehList[2] = float(all_veh_attributes[cnt][4].split('-')[0])	#link
			vehList[3] = all_veh_attributes[cnt][5]/5280	#position
			if vehList[1] < 25:
				vehList[4] = "Yes"
			else:
				vehList[4]="No"
				#print vehList[2]
			vehList[5] = vehList[3] + LinkLength(vehList[2])
			CVData.append(vehList)
	#print CVData
	return CVData

def ReadDataCollections(realTime):
	"""
	This module returns the unique data collection measurements made at time t = realTime.
	Data consists of [realTime, realTime-20, myVolume, myOccupancy, mySpeed]

    """
	
	DCinfo = [[0 for x in range(5)] for x in range(19)] #blank array
	for j in range(1,20):
		DC = Vissim.Net.DataCollectionMeasurements.ItemByKey(j)
		mySpeed = DC.AttValue('Speed(Current,Avg,All)')
		myVolume = DC.AttValue('Vehs(Current,Total,All)')
		myOccupancy = DC.AttValue('Pers(Current,Total,All)')
		
		#assign VISSIM data to dynamic array
		DCinfo[j-1][0] = realTime
		DCinfo[j-1][1] = realTime + 20
		DCinfo[j-1][2] = myVolume
		DCinfo[j-1][3] = myOccupancy
		DCinfo[j-1][4] = mySpeed

	return DCinfo

def LinkLength(linkNum):
	"""
	This module returns the milemarker starting point for each LinkIDs

	"""
	if linkNum == 545:
		return 0
	elif linkNum == 514:
		return 0.91
	elif linkNum == 515:
		return 1.26
	elif linkNum == 516:
		return 1.54
	elif linkNum == 517:
		return 1.74
	elif linkNum == 523:
		return 2.24
	elif linkNum == 524:
	    return 2.46
	elif linkNum == 527:
		return 3.87
	elif linkNum == 531:
		return 4.19
	elif linkNum == 226:
		return 4.52
	elif linkNum == 216:
		return 5.09
	elif linkNum == 214:
		return 5.31
	elif linkNum == 107:
		return 5.54
	elif linkNum == 106:
		return 6.61
	elif linkNum == 105:
		return 6.84
	elif linkNum == 90:
		return 7.04
	elif linkNum == 83:
		return 7.21
	else:
		return 0  #link number not valued

def WriteData(DCinfo, newCVData):
	"""
	This module writes TSSData and CVData to the access tables to be read by the INFLO application.
	TSSInput includes: DSId, Volume, Occupancy and Average Speed
	CVData includes: VehicleID, Speed, MMLocation and QueuedState

    """
	#CurrDir = os.getcwd() #current directory
	syncfile = 'D:\\RKK\\AMSRuns\\Data.txt'
	file = open(syncfile, "w") #make text file
	conn = com.Dispatch(r'ADODB.Connection')
	db = 'D:\\RKK\\AMSRuns\\INFLODatabase.accdb'
	DSN = ('PROVIDER = Microsoft.ACE.OLEDB.12.0;Data Source = ' + db)
	conn.Open(DSN)
	rs = com.Dispatch(r'ADODB.Recordset')
	state1 = "DELETE * FROM TME_TSSData_Input"
	state2 = "DELETE * FROM TME_CVData_Input"
	rs = conn.Execute(state1) #clear TSSData sheet
	rs = conn.Execute(state2) #clear CVData sheet
	CVData = newCVData
	    
    #custom built method of reading array to TSSData Sheet DB via SQL
	for i in range(1,20):
		sql1 = "INSERT INTO TME_TSSData_Input (DSId, Volume, Occupancy, AvgSpeed) VALUES (" + str(i) + "," + str(DCinfo[i-1][2]) + "," + str(DCinfo[i-1][3]) + "," + str(DCinfo[i-1][4]) + ");"
		rs = conn.Execute(sql1)
   
	#custom built method of reading array to CVData Sheet DB via SQL
	for i in range(1,len(CVData)):
		sql2 = "INSERT INTO TME_CVData_Input (NomadicDeviceID, Speed, MMLocation, CVQueuedState) \
				VALUES (" + str(CVData[i-1][0]) + "," + str(CVData[i-1][1]) + "," + str(CVData[i-1][5]) \
				+ "," + str(CVData[i-1][4]) + ");"
		rs = conn.Execute(sql2)
		
	conn.close
    #write lines to new text file, which tells INFLO to work
	file.write("TSSData\n")
	file.write("CVData\n")
	file.close() #dont forget to close!
		
def GetSpeed():
	"""
	This module reads and returns the harmonized speed values from the Access Tables.

    """
	conn = com.Dispatch(r'ADODB.Connection')
	#CurrDir = os.getcwd() #current directory
	db = 'D:\\RKK\\AMSRuns\\INFLODatabase.accdb'
	DSN = ('PROVIDER = Microsoft.ACE.OLEDB.12.0;Data Source = ' + db + ';')
	conn.Open(DSN)
	rs = com.Dispatch(r'ADODB.Recordset')
	strsql = "SELECT TMEOutput_SPDHARMMessage_CV.* FROM TMEOutput_SPDHARMMessage_CV;"
	rs.Open(strsql, conn, 1, 3)
	SpeedData = rs.GetRows()
	sql = "DELETE * FROM TMEOutput_SPDHARMMessage_CV"
	conn.Execute(sql)
	conn.Close()
	return SpeedData

def MyRound(speed, base=5):
	"""
	This module rounds off the harmonized speed to nearest '5's.

    """
	return int(base * round(float(speed)/base))	

def SpeedDecision(SpeedData):
	"""
	This module converts the harmonized speeds to Speed Decisions.

    """
	SpeedDecMatrix = []
	for i in range (0,80):
		if SpeedData[2][i] == 0:
			SpeedDecMatrix.append(1)
		else:
			speed = SpeedData[2][i]
			if speed < 30:
				speed = 30
			newspeed = MyRound(speed)
			SpeedDecMatrix.append(newspeed)
	return SpeedDecMatrix

def UpdateSD(SpeedDecMatrix):
	SpeedDs = SpeedDecMatrix
	Vissim.Net.DesSpeedDecisions.ItemByKey(24).SetAttValue('DesSpeedDistr(1)' , SpeedDs[0])
	Vissim.Net.DesSpeedDecisions.ItemByKey(25).SetAttValue('DesSpeedDistr(1)' , SpeedDs[0])
	Vissim.Net.DesSpeedDecisions.ItemByKey(26).SetAttValue('DesSpeedDistr(1)' , SpeedDs[0])
	Vissim.Net.DesSpeedDecisions.ItemByKey(27).SetAttValue('DesSpeedDistr(1)' , SpeedDs[0])
	Vissim.Net.DesSpeedDecisions.ItemByKey(28).SetAttValue('DesSpeedDistr(1)' , SpeedDs[0])
	Vissim.Net.DesSpeedDecisions.ItemByKey(29).SetAttValue('DesSpeedDistr(1)' , SpeedDs[1])
	Vissim.Net.DesSpeedDecisions.ItemByKey(30).SetAttValue('DesSpeedDistr(1)' , SpeedDs[1])
	Vissim.Net.DesSpeedDecisions.ItemByKey(31).SetAttValue('DesSpeedDistr(1)' , SpeedDs[1])
	Vissim.Net.DesSpeedDecisions.ItemByKey(32).SetAttValue('DesSpeedDistr(1)' , SpeedDs[1])
	Vissim.Net.DesSpeedDecisions.ItemByKey(33).SetAttValue('DesSpeedDistr(1)' , SpeedDs[1])
	Vissim.Net.DesSpeedDecisions.ItemByKey(34).SetAttValue('DesSpeedDistr(1)' , SpeedDs[2])
	Vissim.Net.DesSpeedDecisions.ItemByKey(35).SetAttValue('DesSpeedDistr(1)' , SpeedDs[2])
	Vissim.Net.DesSpeedDecisions.ItemByKey(36).SetAttValue('DesSpeedDistr(1)' , SpeedDs[2])
	Vissim.Net.DesSpeedDecisions.ItemByKey(37).SetAttValue('DesSpeedDistr(1)' , SpeedDs[2])
	Vissim.Net.DesSpeedDecisions.ItemByKey(38).SetAttValue('DesSpeedDistr(1)' , SpeedDs[2])
	Vissim.Net.DesSpeedDecisions.ItemByKey(39).SetAttValue('DesSpeedDistr(1)' , SpeedDs[3])
	Vissim.Net.DesSpeedDecisions.ItemByKey(40).SetAttValue('DesSpeedDistr(1)' , SpeedDs[3])
	Vissim.Net.DesSpeedDecisions.ItemByKey(41).SetAttValue('DesSpeedDistr(1)' , SpeedDs[3])
	Vissim.Net.DesSpeedDecisions.ItemByKey(42).SetAttValue('DesSpeedDistr(1)' , SpeedDs[3])
	Vissim.Net.DesSpeedDecisions.ItemByKey(43).SetAttValue('DesSpeedDistr(1)' , SpeedDs[3])
	Vissim.Net.DesSpeedDecisions.ItemByKey(44).SetAttValue('DesSpeedDistr(1)' , SpeedDs[4])
	Vissim.Net.DesSpeedDecisions.ItemByKey(45).SetAttValue('DesSpeedDistr(1)' , SpeedDs[4])
	Vissim.Net.DesSpeedDecisions.ItemByKey(46).SetAttValue('DesSpeedDistr(1)' , SpeedDs[4])
	Vissim.Net.DesSpeedDecisions.ItemByKey(47).SetAttValue('DesSpeedDistr(1)' , SpeedDs[4])
	Vissim.Net.DesSpeedDecisions.ItemByKey(48).SetAttValue('DesSpeedDistr(1)' , SpeedDs[4])
	Vissim.Net.DesSpeedDecisions.ItemByKey(49).SetAttValue('DesSpeedDistr(1)' , SpeedDs[5])
	Vissim.Net.DesSpeedDecisions.ItemByKey(50).SetAttValue('DesSpeedDistr(1)' , SpeedDs[5])
	Vissim.Net.DesSpeedDecisions.ItemByKey(51).SetAttValue('DesSpeedDistr(1)' , SpeedDs[5])
	Vissim.Net.DesSpeedDecisions.ItemByKey(52).SetAttValue('DesSpeedDistr(1)' , SpeedDs[5])
	Vissim.Net.DesSpeedDecisions.ItemByKey(53).SetAttValue('DesSpeedDistr(1)' , SpeedDs[5])
	Vissim.Net.DesSpeedDecisions.ItemByKey(54).SetAttValue('DesSpeedDistr(1)' , SpeedDs[6])
	Vissim.Net.DesSpeedDecisions.ItemByKey(55).SetAttValue('DesSpeedDistr(1)' , SpeedDs[6])
	Vissim.Net.DesSpeedDecisions.ItemByKey(56).SetAttValue('DesSpeedDistr(1)' , SpeedDs[6])
	Vissim.Net.DesSpeedDecisions.ItemByKey(57).SetAttValue('DesSpeedDistr(1)' , SpeedDs[6])
	Vissim.Net.DesSpeedDecisions.ItemByKey(58).SetAttValue('DesSpeedDistr(1)' , SpeedDs[6])
	Vissim.Net.DesSpeedDecisions.ItemByKey(59).SetAttValue('DesSpeedDistr(1)' , SpeedDs[7])
	Vissim.Net.DesSpeedDecisions.ItemByKey(60).SetAttValue('DesSpeedDistr(1)' , SpeedDs[7])
	Vissim.Net.DesSpeedDecisions.ItemByKey(61).SetAttValue('DesSpeedDistr(1)' , SpeedDs[7])
	Vissim.Net.DesSpeedDecisions.ItemByKey(62).SetAttValue('DesSpeedDistr(1)' , SpeedDs[7])
	Vissim.Net.DesSpeedDecisions.ItemByKey(63).SetAttValue('DesSpeedDistr(1)' , SpeedDs[7])
	Vissim.Net.DesSpeedDecisions.ItemByKey(64).SetAttValue('DesSpeedDistr(1)' , SpeedDs[8])
	Vissim.Net.DesSpeedDecisions.ItemByKey(65).SetAttValue('DesSpeedDistr(1)' , SpeedDs[8])
	Vissim.Net.DesSpeedDecisions.ItemByKey(66).SetAttValue('DesSpeedDistr(1)' , SpeedDs[8])
	Vissim.Net.DesSpeedDecisions.ItemByKey(67).SetAttValue('DesSpeedDistr(1)' , SpeedDs[8])
	Vissim.Net.DesSpeedDecisions.ItemByKey(68).SetAttValue('DesSpeedDistr(1)' , SpeedDs[8])
	Vissim.Net.DesSpeedDecisions.ItemByKey(69).SetAttValue('DesSpeedDistr(1)' , SpeedDs[9])
	Vissim.Net.DesSpeedDecisions.ItemByKey(70).SetAttValue('DesSpeedDistr(1)' , SpeedDs[9])
	Vissim.Net.DesSpeedDecisions.ItemByKey(71).SetAttValue('DesSpeedDistr(1)' , SpeedDs[9])
	Vissim.Net.DesSpeedDecisions.ItemByKey(72).SetAttValue('DesSpeedDistr(1)' , SpeedDs[9])
	Vissim.Net.DesSpeedDecisions.ItemByKey(73).SetAttValue('DesSpeedDistr(1)' , SpeedDs[9])
	Vissim.Net.DesSpeedDecisions.ItemByKey(74).SetAttValue('DesSpeedDistr(1)' , SpeedDs[10])
	Vissim.Net.DesSpeedDecisions.ItemByKey(75).SetAttValue('DesSpeedDistr(1)' , SpeedDs[10])
	Vissim.Net.DesSpeedDecisions.ItemByKey(76).SetAttValue('DesSpeedDistr(1)' , SpeedDs[10])
	Vissim.Net.DesSpeedDecisions.ItemByKey(77).SetAttValue('DesSpeedDistr(1)' , SpeedDs[10])
	Vissim.Net.DesSpeedDecisions.ItemByKey(78).SetAttValue('DesSpeedDistr(1)' , SpeedDs[10])
	Vissim.Net.DesSpeedDecisions.ItemByKey(79).SetAttValue('DesSpeedDistr(1)' , SpeedDs[11])
	Vissim.Net.DesSpeedDecisions.ItemByKey(80).SetAttValue('DesSpeedDistr(1)' , SpeedDs[11])
	Vissim.Net.DesSpeedDecisions.ItemByKey(81).SetAttValue('DesSpeedDistr(1)' , SpeedDs[11])
	Vissim.Net.DesSpeedDecisions.ItemByKey(82).SetAttValue('DesSpeedDistr(1)' , SpeedDs[11])
	Vissim.Net.DesSpeedDecisions.ItemByKey(83).SetAttValue('DesSpeedDistr(1)' , SpeedDs[11])
	Vissim.Net.DesSpeedDecisions.ItemByKey(84).SetAttValue('DesSpeedDistr(1)' , SpeedDs[12])
	Vissim.Net.DesSpeedDecisions.ItemByKey(85).SetAttValue('DesSpeedDistr(1)' , SpeedDs[12])
	Vissim.Net.DesSpeedDecisions.ItemByKey(86).SetAttValue('DesSpeedDistr(1)' , SpeedDs[12])
	Vissim.Net.DesSpeedDecisions.ItemByKey(87).SetAttValue('DesSpeedDistr(1)' , SpeedDs[12])
	Vissim.Net.DesSpeedDecisions.ItemByKey(88).SetAttValue('DesSpeedDistr(1)' , SpeedDs[12])
	Vissim.Net.DesSpeedDecisions.ItemByKey(89).SetAttValue('DesSpeedDistr(1)' , SpeedDs[13])
	Vissim.Net.DesSpeedDecisions.ItemByKey(90).SetAttValue('DesSpeedDistr(1)' , SpeedDs[13])
	Vissim.Net.DesSpeedDecisions.ItemByKey(91).SetAttValue('DesSpeedDistr(1)' , SpeedDs[13])
	Vissim.Net.DesSpeedDecisions.ItemByKey(92).SetAttValue('DesSpeedDistr(1)' , SpeedDs[13])
	Vissim.Net.DesSpeedDecisions.ItemByKey(93).SetAttValue('DesSpeedDistr(1)' , SpeedDs[13])
	Vissim.Net.DesSpeedDecisions.ItemByKey(94).SetAttValue('DesSpeedDistr(1)' , SpeedDs[14])
	Vissim.Net.DesSpeedDecisions.ItemByKey(95).SetAttValue('DesSpeedDistr(1)' , SpeedDs[14])
	Vissim.Net.DesSpeedDecisions.ItemByKey(96).SetAttValue('DesSpeedDistr(1)' , SpeedDs[14])
	Vissim.Net.DesSpeedDecisions.ItemByKey(97).SetAttValue('DesSpeedDistr(1)' , SpeedDs[14])
	Vissim.Net.DesSpeedDecisions.ItemByKey(98).SetAttValue('DesSpeedDistr(1)' , SpeedDs[14])
	Vissim.Net.DesSpeedDecisions.ItemByKey(99).SetAttValue('DesSpeedDistr(1)' , SpeedDs[15])
	Vissim.Net.DesSpeedDecisions.ItemByKey(100).SetAttValue('DesSpeedDistr(1)' , SpeedDs[15])
	Vissim.Net.DesSpeedDecisions.ItemByKey(101).SetAttValue('DesSpeedDistr(1)' , SpeedDs[15])
	Vissim.Net.DesSpeedDecisions.ItemByKey(102).SetAttValue('DesSpeedDistr(1)' , SpeedDs[15])
	Vissim.Net.DesSpeedDecisions.ItemByKey(103).SetAttValue('DesSpeedDistr(1)' , SpeedDs[15])
	Vissim.Net.DesSpeedDecisions.ItemByKey(104).SetAttValue('DesSpeedDistr(1)' , SpeedDs[16])
	Vissim.Net.DesSpeedDecisions.ItemByKey(105).SetAttValue('DesSpeedDistr(1)' , SpeedDs[16])
	Vissim.Net.DesSpeedDecisions.ItemByKey(106).SetAttValue('DesSpeedDistr(1)' , SpeedDs[16])
	Vissim.Net.DesSpeedDecisions.ItemByKey(107).SetAttValue('DesSpeedDistr(1)' , SpeedDs[16])
	Vissim.Net.DesSpeedDecisions.ItemByKey(108).SetAttValue('DesSpeedDistr(1)' , SpeedDs[16])
	Vissim.Net.DesSpeedDecisions.ItemByKey(109).SetAttValue('DesSpeedDistr(1)' , SpeedDs[17])
	Vissim.Net.DesSpeedDecisions.ItemByKey(110).SetAttValue('DesSpeedDistr(1)' , SpeedDs[17])
	Vissim.Net.DesSpeedDecisions.ItemByKey(111).SetAttValue('DesSpeedDistr(1)' , SpeedDs[17])
	Vissim.Net.DesSpeedDecisions.ItemByKey(112).SetAttValue('DesSpeedDistr(1)' , SpeedDs[17])
	Vissim.Net.DesSpeedDecisions.ItemByKey(113).SetAttValue('DesSpeedDistr(1)' , SpeedDs[17])
	Vissim.Net.DesSpeedDecisions.ItemByKey(114).SetAttValue('DesSpeedDistr(1)' , SpeedDs[18])
	Vissim.Net.DesSpeedDecisions.ItemByKey(115).SetAttValue('DesSpeedDistr(1)' , SpeedDs[18])
	Vissim.Net.DesSpeedDecisions.ItemByKey(116).SetAttValue('DesSpeedDistr(1)' , SpeedDs[18])
	Vissim.Net.DesSpeedDecisions.ItemByKey(117).SetAttValue('DesSpeedDistr(1)' , SpeedDs[18])
	Vissim.Net.DesSpeedDecisions.ItemByKey(118).SetAttValue('DesSpeedDistr(1)' , SpeedDs[18])
	Vissim.Net.DesSpeedDecisions.ItemByKey(119).SetAttValue('DesSpeedDistr(1)' , SpeedDs[19])
	Vissim.Net.DesSpeedDecisions.ItemByKey(120).SetAttValue('DesSpeedDistr(1)' , SpeedDs[19])
	Vissim.Net.DesSpeedDecisions.ItemByKey(121).SetAttValue('DesSpeedDistr(1)' , SpeedDs[19])
	Vissim.Net.DesSpeedDecisions.ItemByKey(122).SetAttValue('DesSpeedDistr(1)' , SpeedDs[19])
	Vissim.Net.DesSpeedDecisions.ItemByKey(123).SetAttValue('DesSpeedDistr(1)' , SpeedDs[20])
	Vissim.Net.DesSpeedDecisions.ItemByKey(124).SetAttValue('DesSpeedDistr(1)' , SpeedDs[20])
	Vissim.Net.DesSpeedDecisions.ItemByKey(125).SetAttValue('DesSpeedDistr(1)' , SpeedDs[20])
	Vissim.Net.DesSpeedDecisions.ItemByKey(126).SetAttValue('DesSpeedDistr(1)' , SpeedDs[20])
	Vissim.Net.DesSpeedDecisions.ItemByKey(127).SetAttValue('DesSpeedDistr(1)' , SpeedDs[21])
	Vissim.Net.DesSpeedDecisions.ItemByKey(128).SetAttValue('DesSpeedDistr(1)' , SpeedDs[21])
	Vissim.Net.DesSpeedDecisions.ItemByKey(129).SetAttValue('DesSpeedDistr(1)' , SpeedDs[21])
	Vissim.Net.DesSpeedDecisions.ItemByKey(130).SetAttValue('DesSpeedDistr(1)' , SpeedDs[21])
	Vissim.Net.DesSpeedDecisions.ItemByKey(131).SetAttValue('DesSpeedDistr(1)' , SpeedDs[22])
	Vissim.Net.DesSpeedDecisions.ItemByKey(132).SetAttValue('DesSpeedDistr(1)' , SpeedDs[22])
	Vissim.Net.DesSpeedDecisions.ItemByKey(133).SetAttValue('DesSpeedDistr(1)' , SpeedDs[22])
	Vissim.Net.DesSpeedDecisions.ItemByKey(134).SetAttValue('DesSpeedDistr(1)' , SpeedDs[22])
	Vissim.Net.DesSpeedDecisions.ItemByKey(135).SetAttValue('DesSpeedDistr(1)' , SpeedDs[23])
	Vissim.Net.DesSpeedDecisions.ItemByKey(136).SetAttValue('DesSpeedDistr(1)' , SpeedDs[23])
	Vissim.Net.DesSpeedDecisions.ItemByKey(137).SetAttValue('DesSpeedDistr(1)' , SpeedDs[23])
	Vissim.Net.DesSpeedDecisions.ItemByKey(138).SetAttValue('DesSpeedDistr(1)' , SpeedDs[23])
	Vissim.Net.DesSpeedDecisions.ItemByKey(139).SetAttValue('DesSpeedDistr(1)' , SpeedDs[24])
	Vissim.Net.DesSpeedDecisions.ItemByKey(140).SetAttValue('DesSpeedDistr(1)' , SpeedDs[24])
	Vissim.Net.DesSpeedDecisions.ItemByKey(141).SetAttValue('DesSpeedDistr(1)' , SpeedDs[24])
	Vissim.Net.DesSpeedDecisions.ItemByKey(142).SetAttValue('DesSpeedDistr(1)' , SpeedDs[24])
	Vissim.Net.DesSpeedDecisions.ItemByKey(143).SetAttValue('DesSpeedDistr(1)' , SpeedDs[25])
	Vissim.Net.DesSpeedDecisions.ItemByKey(144).SetAttValue('DesSpeedDistr(1)' , SpeedDs[25])
	Vissim.Net.DesSpeedDecisions.ItemByKey(145).SetAttValue('DesSpeedDistr(1)' , SpeedDs[25])
	Vissim.Net.DesSpeedDecisions.ItemByKey(146).SetAttValue('DesSpeedDistr(1)' , SpeedDs[25])
	Vissim.Net.DesSpeedDecisions.ItemByKey(147).SetAttValue('DesSpeedDistr(1)' , SpeedDs[26])
	Vissim.Net.DesSpeedDecisions.ItemByKey(148).SetAttValue('DesSpeedDistr(1)' , SpeedDs[26])
	Vissim.Net.DesSpeedDecisions.ItemByKey(149).SetAttValue('DesSpeedDistr(1)' , SpeedDs[26])
	Vissim.Net.DesSpeedDecisions.ItemByKey(150).SetAttValue('DesSpeedDistr(1)' , SpeedDs[26])
	Vissim.Net.DesSpeedDecisions.ItemByKey(151).SetAttValue('DesSpeedDistr(1)' , SpeedDs[26])
	Vissim.Net.DesSpeedDecisions.ItemByKey(152).SetAttValue('DesSpeedDistr(1)' , SpeedDs[27])
	Vissim.Net.DesSpeedDecisions.ItemByKey(153).SetAttValue('DesSpeedDistr(1)' , SpeedDs[27])
	Vissim.Net.DesSpeedDecisions.ItemByKey(154).SetAttValue('DesSpeedDistr(1)' , SpeedDs[27])
	Vissim.Net.DesSpeedDecisions.ItemByKey(155).SetAttValue('DesSpeedDistr(1)' , SpeedDs[27])
	Vissim.Net.DesSpeedDecisions.ItemByKey(156).SetAttValue('DesSpeedDistr(1)' , SpeedDs[27])
	Vissim.Net.DesSpeedDecisions.ItemByKey(157).SetAttValue('DesSpeedDistr(1)' , SpeedDs[28])
	Vissim.Net.DesSpeedDecisions.ItemByKey(158).SetAttValue('DesSpeedDistr(1)' , SpeedDs[28])
	Vissim.Net.DesSpeedDecisions.ItemByKey(159).SetAttValue('DesSpeedDistr(1)' , SpeedDs[28])
	Vissim.Net.DesSpeedDecisions.ItemByKey(160).SetAttValue('DesSpeedDistr(1)' , SpeedDs[28])
	Vissim.Net.DesSpeedDecisions.ItemByKey(161).SetAttValue('DesSpeedDistr(1)' , SpeedDs[28])
	Vissim.Net.DesSpeedDecisions.ItemByKey(162).SetAttValue('DesSpeedDistr(1)' , SpeedDs[29])
	Vissim.Net.DesSpeedDecisions.ItemByKey(163).SetAttValue('DesSpeedDistr(1)' , SpeedDs[29])
	Vissim.Net.DesSpeedDecisions.ItemByKey(164).SetAttValue('DesSpeedDistr(1)' , SpeedDs[29])
	Vissim.Net.DesSpeedDecisions.ItemByKey(165).SetAttValue('DesSpeedDistr(1)' , SpeedDs[29])
	Vissim.Net.DesSpeedDecisions.ItemByKey(166).SetAttValue('DesSpeedDistr(1)' , SpeedDs[29])
	Vissim.Net.DesSpeedDecisions.ItemByKey(167).SetAttValue('DesSpeedDistr(1)' , SpeedDs[30])
	Vissim.Net.DesSpeedDecisions.ItemByKey(168).SetAttValue('DesSpeedDistr(1)' , SpeedDs[30])
	Vissim.Net.DesSpeedDecisions.ItemByKey(169).SetAttValue('DesSpeedDistr(1)' , SpeedDs[30])
	Vissim.Net.DesSpeedDecisions.ItemByKey(170).SetAttValue('DesSpeedDistr(1)' , SpeedDs[30])
	Vissim.Net.DesSpeedDecisions.ItemByKey(171).SetAttValue('DesSpeedDistr(1)' , SpeedDs[30])
	Vissim.Net.DesSpeedDecisions.ItemByKey(172).SetAttValue('DesSpeedDistr(1)' , SpeedDs[31])
	Vissim.Net.DesSpeedDecisions.ItemByKey(173).SetAttValue('DesSpeedDistr(1)' , SpeedDs[31])
	Vissim.Net.DesSpeedDecisions.ItemByKey(174).SetAttValue('DesSpeedDistr(1)' , SpeedDs[31])
	Vissim.Net.DesSpeedDecisions.ItemByKey(175).SetAttValue('DesSpeedDistr(1)' , SpeedDs[31])
	Vissim.Net.DesSpeedDecisions.ItemByKey(176).SetAttValue('DesSpeedDistr(1)' , SpeedDs[31])
	Vissim.Net.DesSpeedDecisions.ItemByKey(177).SetAttValue('DesSpeedDistr(1)' , SpeedDs[32])
	Vissim.Net.DesSpeedDecisions.ItemByKey(178).SetAttValue('DesSpeedDistr(1)' , SpeedDs[32])
	Vissim.Net.DesSpeedDecisions.ItemByKey(179).SetAttValue('DesSpeedDistr(1)' , SpeedDs[32])
	Vissim.Net.DesSpeedDecisions.ItemByKey(180).SetAttValue('DesSpeedDistr(1)' , SpeedDs[32])
	Vissim.Net.DesSpeedDecisions.ItemByKey(181).SetAttValue('DesSpeedDistr(1)' , SpeedDs[32])
	Vissim.Net.DesSpeedDecisions.ItemByKey(182).SetAttValue('DesSpeedDistr(1)' , SpeedDs[33])
	Vissim.Net.DesSpeedDecisions.ItemByKey(183).SetAttValue('DesSpeedDistr(1)' , SpeedDs[33])
	Vissim.Net.DesSpeedDecisions.ItemByKey(184).SetAttValue('DesSpeedDistr(1)' , SpeedDs[33])
	Vissim.Net.DesSpeedDecisions.ItemByKey(185).SetAttValue('DesSpeedDistr(1)' , SpeedDs[33])
	Vissim.Net.DesSpeedDecisions.ItemByKey(186).SetAttValue('DesSpeedDistr(1)' , SpeedDs[34])
	Vissim.Net.DesSpeedDecisions.ItemByKey(187).SetAttValue('DesSpeedDistr(1)' , SpeedDs[34])
	Vissim.Net.DesSpeedDecisions.ItemByKey(188).SetAttValue('DesSpeedDistr(1)' , SpeedDs[34])
	Vissim.Net.DesSpeedDecisions.ItemByKey(189).SetAttValue('DesSpeedDistr(1)' , SpeedDs[34])
	Vissim.Net.DesSpeedDecisions.ItemByKey(190).SetAttValue('DesSpeedDistr(1)' , SpeedDs[35])
	Vissim.Net.DesSpeedDecisions.ItemByKey(191).SetAttValue('DesSpeedDistr(1)' , SpeedDs[35])
	Vissim.Net.DesSpeedDecisions.ItemByKey(192).SetAttValue('DesSpeedDistr(1)' , SpeedDs[35])
	Vissim.Net.DesSpeedDecisions.ItemByKey(193).SetAttValue('DesSpeedDistr(1)' , SpeedDs[35])
	Vissim.Net.DesSpeedDecisions.ItemByKey(194).SetAttValue('DesSpeedDistr(1)' , SpeedDs[35])
	Vissim.Net.DesSpeedDecisions.ItemByKey(195).SetAttValue('DesSpeedDistr(1)' , SpeedDs[36])
	Vissim.Net.DesSpeedDecisions.ItemByKey(196).SetAttValue('DesSpeedDistr(1)' , SpeedDs[36])
	Vissim.Net.DesSpeedDecisions.ItemByKey(197).SetAttValue('DesSpeedDistr(1)' , SpeedDs[36])
	Vissim.Net.DesSpeedDecisions.ItemByKey(198).SetAttValue('DesSpeedDistr(1)' , SpeedDs[36])
	Vissim.Net.DesSpeedDecisions.ItemByKey(199).SetAttValue('DesSpeedDistr(1)' , SpeedDs[36])
	Vissim.Net.DesSpeedDecisions.ItemByKey(200).SetAttValue('DesSpeedDistr(1)' , SpeedDs[37])
	Vissim.Net.DesSpeedDecisions.ItemByKey(201).SetAttValue('DesSpeedDistr(1)' , SpeedDs[37])
	Vissim.Net.DesSpeedDecisions.ItemByKey(202).SetAttValue('DesSpeedDistr(1)' , SpeedDs[37])
	Vissim.Net.DesSpeedDecisions.ItemByKey(203).SetAttValue('DesSpeedDistr(1)' , SpeedDs[37])
	Vissim.Net.DesSpeedDecisions.ItemByKey(204).SetAttValue('DesSpeedDistr(1)' , SpeedDs[37])
	Vissim.Net.DesSpeedDecisions.ItemByKey(205).SetAttValue('DesSpeedDistr(1)' , SpeedDs[38])
	Vissim.Net.DesSpeedDecisions.ItemByKey(206).SetAttValue('DesSpeedDistr(1)' , SpeedDs[38])
	Vissim.Net.DesSpeedDecisions.ItemByKey(207).SetAttValue('DesSpeedDistr(1)' , SpeedDs[38])
	Vissim.Net.DesSpeedDecisions.ItemByKey(208).SetAttValue('DesSpeedDistr(1)' , SpeedDs[38])
	Vissim.Net.DesSpeedDecisions.ItemByKey(209).SetAttValue('DesSpeedDistr(1)' , SpeedDs[38])
	Vissim.Net.DesSpeedDecisions.ItemByKey(210).SetAttValue('DesSpeedDistr(1)' , SpeedDs[39])
	Vissim.Net.DesSpeedDecisions.ItemByKey(211).SetAttValue('DesSpeedDistr(1)' , SpeedDs[39])
	Vissim.Net.DesSpeedDecisions.ItemByKey(212).SetAttValue('DesSpeedDistr(1)' , SpeedDs[39])
	Vissim.Net.DesSpeedDecisions.ItemByKey(213).SetAttValue('DesSpeedDistr(1)' , SpeedDs[39])
	Vissim.Net.DesSpeedDecisions.ItemByKey(214).SetAttValue('DesSpeedDistr(1)' , SpeedDs[39])
	Vissim.Net.DesSpeedDecisions.ItemByKey(215).SetAttValue('DesSpeedDistr(1)' , SpeedDs[40])
	Vissim.Net.DesSpeedDecisions.ItemByKey(216).SetAttValue('DesSpeedDistr(1)' , SpeedDs[40])
	Vissim.Net.DesSpeedDecisions.ItemByKey(217).SetAttValue('DesSpeedDistr(1)' , SpeedDs[40])
	Vissim.Net.DesSpeedDecisions.ItemByKey(218).SetAttValue('DesSpeedDistr(1)' , SpeedDs[40])
	Vissim.Net.DesSpeedDecisions.ItemByKey(219).SetAttValue('DesSpeedDistr(1)' , SpeedDs[40])
	Vissim.Net.DesSpeedDecisions.ItemByKey(220).SetAttValue('DesSpeedDistr(1)' , SpeedDs[41])
	Vissim.Net.DesSpeedDecisions.ItemByKey(221).SetAttValue('DesSpeedDistr(1)' , SpeedDs[41])
	Vissim.Net.DesSpeedDecisions.ItemByKey(222).SetAttValue('DesSpeedDistr(1)' , SpeedDs[41])
	Vissim.Net.DesSpeedDecisions.ItemByKey(223).SetAttValue('DesSpeedDistr(1)' , SpeedDs[41])
	Vissim.Net.DesSpeedDecisions.ItemByKey(224).SetAttValue('DesSpeedDistr(1)' , SpeedDs[41])
	Vissim.Net.DesSpeedDecisions.ItemByKey(225).SetAttValue('DesSpeedDistr(1)' , SpeedDs[42])
	Vissim.Net.DesSpeedDecisions.ItemByKey(226).SetAttValue('DesSpeedDistr(1)' , SpeedDs[42])
	Vissim.Net.DesSpeedDecisions.ItemByKey(227).SetAttValue('DesSpeedDistr(1)' , SpeedDs[42])
	Vissim.Net.DesSpeedDecisions.ItemByKey(228).SetAttValue('DesSpeedDistr(1)' , SpeedDs[42])
	Vissim.Net.DesSpeedDecisions.ItemByKey(229).SetAttValue('DesSpeedDistr(1)' , SpeedDs[42])
	Vissim.Net.DesSpeedDecisions.ItemByKey(230).SetAttValue('DesSpeedDistr(1)' , SpeedDs[43])
	Vissim.Net.DesSpeedDecisions.ItemByKey(231).SetAttValue('DesSpeedDistr(1)' , SpeedDs[43])
	Vissim.Net.DesSpeedDecisions.ItemByKey(232).SetAttValue('DesSpeedDistr(1)' , SpeedDs[43])
	Vissim.Net.DesSpeedDecisions.ItemByKey(233).SetAttValue('DesSpeedDistr(1)' , SpeedDs[43])
	Vissim.Net.DesSpeedDecisions.ItemByKey(234).SetAttValue('DesSpeedDistr(1)' , SpeedDs[43])
	Vissim.Net.DesSpeedDecisions.ItemByKey(235).SetAttValue('DesSpeedDistr(1)' , SpeedDs[44])
	Vissim.Net.DesSpeedDecisions.ItemByKey(236).SetAttValue('DesSpeedDistr(1)' , SpeedDs[44])
	Vissim.Net.DesSpeedDecisions.ItemByKey(237).SetAttValue('DesSpeedDistr(1)' , SpeedDs[44])
	Vissim.Net.DesSpeedDecisions.ItemByKey(238).SetAttValue('DesSpeedDistr(1)' , SpeedDs[44])
	Vissim.Net.DesSpeedDecisions.ItemByKey(239).SetAttValue('DesSpeedDistr(1)' , SpeedDs[44])
	Vissim.Net.DesSpeedDecisions.ItemByKey(240).SetAttValue('DesSpeedDistr(1)' , SpeedDs[45])
	Vissim.Net.DesSpeedDecisions.ItemByKey(241).SetAttValue('DesSpeedDistr(1)' , SpeedDs[45])
	Vissim.Net.DesSpeedDecisions.ItemByKey(242).SetAttValue('DesSpeedDistr(1)' , SpeedDs[45])
	Vissim.Net.DesSpeedDecisions.ItemByKey(243).SetAttValue('DesSpeedDistr(1)' , SpeedDs[45])
	Vissim.Net.DesSpeedDecisions.ItemByKey(244).SetAttValue('DesSpeedDistr(1)' , SpeedDs[45])
	Vissim.Net.DesSpeedDecisions.ItemByKey(245).SetAttValue('DesSpeedDistr(1)' , SpeedDs[46])
	Vissim.Net.DesSpeedDecisions.ItemByKey(246).SetAttValue('DesSpeedDistr(1)' , SpeedDs[46])
	Vissim.Net.DesSpeedDecisions.ItemByKey(247).SetAttValue('DesSpeedDistr(1)' , SpeedDs[46])
	Vissim.Net.DesSpeedDecisions.ItemByKey(248).SetAttValue('DesSpeedDistr(1)' , SpeedDs[46])
	Vissim.Net.DesSpeedDecisions.ItemByKey(249).SetAttValue('DesSpeedDistr(1)' , SpeedDs[46])
	Vissim.Net.DesSpeedDecisions.ItemByKey(250).SetAttValue('DesSpeedDistr(1)' , SpeedDs[47])
	Vissim.Net.DesSpeedDecisions.ItemByKey(251).SetAttValue('DesSpeedDistr(1)' , SpeedDs[47])
	Vissim.Net.DesSpeedDecisions.ItemByKey(252).SetAttValue('DesSpeedDistr(1)' , SpeedDs[47])
	Vissim.Net.DesSpeedDecisions.ItemByKey(253).SetAttValue('DesSpeedDistr(1)' , SpeedDs[47])
	Vissim.Net.DesSpeedDecisions.ItemByKey(254).SetAttValue('DesSpeedDistr(1)' , SpeedDs[47])
	Vissim.Net.DesSpeedDecisions.ItemByKey(255).SetAttValue('DesSpeedDistr(1)' , SpeedDs[48])
	Vissim.Net.DesSpeedDecisions.ItemByKey(256).SetAttValue('DesSpeedDistr(1)' , SpeedDs[48])
	Vissim.Net.DesSpeedDecisions.ItemByKey(257).SetAttValue('DesSpeedDistr(1)' , SpeedDs[48])
	Vissim.Net.DesSpeedDecisions.ItemByKey(258).SetAttValue('DesSpeedDistr(1)' , SpeedDs[48])
	Vissim.Net.DesSpeedDecisions.ItemByKey(259).SetAttValue('DesSpeedDistr(1)' , SpeedDs[48])
	Vissim.Net.DesSpeedDecisions.ItemByKey(260).SetAttValue('DesSpeedDistr(1)' , SpeedDs[49])
	Vissim.Net.DesSpeedDecisions.ItemByKey(261).SetAttValue('DesSpeedDistr(1)' , SpeedDs[49])
	Vissim.Net.DesSpeedDecisions.ItemByKey(262).SetAttValue('DesSpeedDistr(1)' , SpeedDs[49])
	Vissim.Net.DesSpeedDecisions.ItemByKey(263).SetAttValue('DesSpeedDistr(1)' , SpeedDs[49])
	Vissim.Net.DesSpeedDecisions.ItemByKey(264).SetAttValue('DesSpeedDistr(1)' , SpeedDs[50])
	Vissim.Net.DesSpeedDecisions.ItemByKey(265).SetAttValue('DesSpeedDistr(1)' , SpeedDs[50])
	Vissim.Net.DesSpeedDecisions.ItemByKey(266).SetAttValue('DesSpeedDistr(1)' , SpeedDs[50])
	Vissim.Net.DesSpeedDecisions.ItemByKey(267).SetAttValue('DesSpeedDistr(1)' , SpeedDs[50])
	Vissim.Net.DesSpeedDecisions.ItemByKey(268).SetAttValue('DesSpeedDistr(1)' , SpeedDs[51])
	Vissim.Net.DesSpeedDecisions.ItemByKey(269).SetAttValue('DesSpeedDistr(1)' , SpeedDs[51])
	Vissim.Net.DesSpeedDecisions.ItemByKey(270).SetAttValue('DesSpeedDistr(1)' , SpeedDs[51])
	Vissim.Net.DesSpeedDecisions.ItemByKey(271).SetAttValue('DesSpeedDistr(1)' , SpeedDs[51])
	Vissim.Net.DesSpeedDecisions.ItemByKey(272).SetAttValue('DesSpeedDistr(1)' , SpeedDs[52])
	Vissim.Net.DesSpeedDecisions.ItemByKey(273).SetAttValue('DesSpeedDistr(1)' , SpeedDs[52])
	Vissim.Net.DesSpeedDecisions.ItemByKey(274).SetAttValue('DesSpeedDistr(1)' , SpeedDs[52])
	Vissim.Net.DesSpeedDecisions.ItemByKey(275).SetAttValue('DesSpeedDistr(1)' , SpeedDs[52])
	Vissim.Net.DesSpeedDecisions.ItemByKey(276).SetAttValue('DesSpeedDistr(1)' , SpeedDs[53])
	Vissim.Net.DesSpeedDecisions.ItemByKey(277).SetAttValue('DesSpeedDistr(1)' , SpeedDs[53])
	Vissim.Net.DesSpeedDecisions.ItemByKey(278).SetAttValue('DesSpeedDistr(1)' , SpeedDs[53])
	Vissim.Net.DesSpeedDecisions.ItemByKey(279).SetAttValue('DesSpeedDistr(1)' , SpeedDs[53])
	Vissim.Net.DesSpeedDecisions.ItemByKey(280).SetAttValue('DesSpeedDistr(1)' , SpeedDs[54])
	Vissim.Net.DesSpeedDecisions.ItemByKey(281).SetAttValue('DesSpeedDistr(1)' , SpeedDs[54])
	Vissim.Net.DesSpeedDecisions.ItemByKey(282).SetAttValue('DesSpeedDistr(1)' , SpeedDs[54])
	Vissim.Net.DesSpeedDecisions.ItemByKey(283).SetAttValue('DesSpeedDistr(1)' , SpeedDs[54])
	Vissim.Net.DesSpeedDecisions.ItemByKey(284).SetAttValue('DesSpeedDistr(1)' , SpeedDs[55])
	Vissim.Net.DesSpeedDecisions.ItemByKey(285).SetAttValue('DesSpeedDistr(1)' , SpeedDs[55])
	Vissim.Net.DesSpeedDecisions.ItemByKey(286).SetAttValue('DesSpeedDistr(1)' , SpeedDs[55])
	Vissim.Net.DesSpeedDecisions.ItemByKey(287).SetAttValue('DesSpeedDistr(1)' , SpeedDs[55])
	Vissim.Net.DesSpeedDecisions.ItemByKey(288).SetAttValue('DesSpeedDistr(1)' , SpeedDs[55])
	Vissim.Net.DesSpeedDecisions.ItemByKey(289).SetAttValue('DesSpeedDistr(1)' , SpeedDs[56])
	Vissim.Net.DesSpeedDecisions.ItemByKey(290).SetAttValue('DesSpeedDistr(1)' , SpeedDs[56])
	Vissim.Net.DesSpeedDecisions.ItemByKey(291).SetAttValue('DesSpeedDistr(1)' , SpeedDs[56])
	Vissim.Net.DesSpeedDecisions.ItemByKey(292).SetAttValue('DesSpeedDistr(1)' , SpeedDs[56])
	Vissim.Net.DesSpeedDecisions.ItemByKey(293).SetAttValue('DesSpeedDistr(1)' , SpeedDs[57])
	Vissim.Net.DesSpeedDecisions.ItemByKey(294).SetAttValue('DesSpeedDistr(1)' , SpeedDs[57])
	Vissim.Net.DesSpeedDecisions.ItemByKey(295).SetAttValue('DesSpeedDistr(1)' , SpeedDs[57])
	Vissim.Net.DesSpeedDecisions.ItemByKey(296).SetAttValue('DesSpeedDistr(1)' , SpeedDs[57])
	Vissim.Net.DesSpeedDecisions.ItemByKey(297).SetAttValue('DesSpeedDistr(1)' , SpeedDs[57])
	Vissim.Net.DesSpeedDecisions.ItemByKey(298).SetAttValue('DesSpeedDistr(1)' , SpeedDs[58])
	Vissim.Net.DesSpeedDecisions.ItemByKey(299).SetAttValue('DesSpeedDistr(1)' , SpeedDs[58])
	Vissim.Net.DesSpeedDecisions.ItemByKey(300).SetAttValue('DesSpeedDistr(1)' , SpeedDs[58])
	Vissim.Net.DesSpeedDecisions.ItemByKey(301).SetAttValue('DesSpeedDistr(1)' , SpeedDs[58])
	Vissim.Net.DesSpeedDecisions.ItemByKey(302).SetAttValue('DesSpeedDistr(1)' , SpeedDs[58])
	Vissim.Net.DesSpeedDecisions.ItemByKey(303).SetAttValue('DesSpeedDistr(1)' , SpeedDs[59])
	Vissim.Net.DesSpeedDecisions.ItemByKey(304).SetAttValue('DesSpeedDistr(1)' , SpeedDs[59])
	Vissim.Net.DesSpeedDecisions.ItemByKey(305).SetAttValue('DesSpeedDistr(1)' , SpeedDs[59])
	Vissim.Net.DesSpeedDecisions.ItemByKey(306).SetAttValue('DesSpeedDistr(1)' , SpeedDs[59])
	Vissim.Net.DesSpeedDecisions.ItemByKey(307).SetAttValue('DesSpeedDistr(1)' , SpeedDs[59])
	Vissim.Net.DesSpeedDecisions.ItemByKey(308).SetAttValue('DesSpeedDistr(1)' , SpeedDs[60])
	Vissim.Net.DesSpeedDecisions.ItemByKey(309).SetAttValue('DesSpeedDistr(1)' , SpeedDs[60])
	Vissim.Net.DesSpeedDecisions.ItemByKey(310).SetAttValue('DesSpeedDistr(1)' , SpeedDs[60])
	Vissim.Net.DesSpeedDecisions.ItemByKey(311).SetAttValue('DesSpeedDistr(1)' , SpeedDs[60])
	Vissim.Net.DesSpeedDecisions.ItemByKey(312).SetAttValue('DesSpeedDistr(1)' , SpeedDs[60])
	Vissim.Net.DesSpeedDecisions.ItemByKey(313).SetAttValue('DesSpeedDistr(1)' , SpeedDs[61])
	Vissim.Net.DesSpeedDecisions.ItemByKey(314).SetAttValue('DesSpeedDistr(1)' , SpeedDs[61])
	Vissim.Net.DesSpeedDecisions.ItemByKey(315).SetAttValue('DesSpeedDistr(1)' , SpeedDs[61])
	Vissim.Net.DesSpeedDecisions.ItemByKey(316).SetAttValue('DesSpeedDistr(1)' , SpeedDs[61])
	Vissim.Net.DesSpeedDecisions.ItemByKey(317).SetAttValue('DesSpeedDistr(1)' , SpeedDs[61])
	Vissim.Net.DesSpeedDecisions.ItemByKey(318).SetAttValue('DesSpeedDistr(1)' , SpeedDs[62])
	Vissim.Net.DesSpeedDecisions.ItemByKey(319).SetAttValue('DesSpeedDistr(1)' , SpeedDs[62])
	Vissim.Net.DesSpeedDecisions.ItemByKey(320).SetAttValue('DesSpeedDistr(1)' , SpeedDs[62])
	Vissim.Net.DesSpeedDecisions.ItemByKey(321).SetAttValue('DesSpeedDistr(1)' , SpeedDs[62])
	Vissim.Net.DesSpeedDecisions.ItemByKey(322).SetAttValue('DesSpeedDistr(1)' , SpeedDs[63])
	Vissim.Net.DesSpeedDecisions.ItemByKey(323).SetAttValue('DesSpeedDistr(1)' , SpeedDs[63])
	Vissim.Net.DesSpeedDecisions.ItemByKey(324).SetAttValue('DesSpeedDistr(1)' , SpeedDs[63])
	Vissim.Net.DesSpeedDecisions.ItemByKey(325).SetAttValue('DesSpeedDistr(1)' , SpeedDs[63])
	Vissim.Net.DesSpeedDecisions.ItemByKey(326).SetAttValue('DesSpeedDistr(1)' , SpeedDs[64])
	Vissim.Net.DesSpeedDecisions.ItemByKey(327).SetAttValue('DesSpeedDistr(1)' , SpeedDs[64])
	Vissim.Net.DesSpeedDecisions.ItemByKey(328).SetAttValue('DesSpeedDistr(1)' , SpeedDs[64])
	Vissim.Net.DesSpeedDecisions.ItemByKey(329).SetAttValue('DesSpeedDistr(1)' , SpeedDs[64])
	Vissim.Net.DesSpeedDecisions.ItemByKey(330).SetAttValue('DesSpeedDistr(1)' , SpeedDs[65])
	Vissim.Net.DesSpeedDecisions.ItemByKey(331).SetAttValue('DesSpeedDistr(1)' , SpeedDs[65])
	Vissim.Net.DesSpeedDecisions.ItemByKey(332).SetAttValue('DesSpeedDistr(1)' , SpeedDs[65])
	Vissim.Net.DesSpeedDecisions.ItemByKey(333).SetAttValue('DesSpeedDistr(1)' , SpeedDs[65])
	Vissim.Net.DesSpeedDecisions.ItemByKey(334).SetAttValue('DesSpeedDistr(1)' , SpeedDs[66])
	Vissim.Net.DesSpeedDecisions.ItemByKey(335).SetAttValue('DesSpeedDistr(1)' , SpeedDs[66])
	Vissim.Net.DesSpeedDecisions.ItemByKey(336).SetAttValue('DesSpeedDistr(1)' , SpeedDs[66])
	Vissim.Net.DesSpeedDecisions.ItemByKey(337).SetAttValue('DesSpeedDistr(1)' , SpeedDs[66])
	Vissim.Net.DesSpeedDecisions.ItemByKey(338).SetAttValue('DesSpeedDistr(1)' , SpeedDs[67])
	Vissim.Net.DesSpeedDecisions.ItemByKey(339).SetAttValue('DesSpeedDistr(1)' , SpeedDs[67])
	Vissim.Net.DesSpeedDecisions.ItemByKey(340).SetAttValue('DesSpeedDistr(1)' , SpeedDs[67])
	Vissim.Net.DesSpeedDecisions.ItemByKey(341).SetAttValue('DesSpeedDistr(1)' , SpeedDs[67])
	Vissim.Net.DesSpeedDecisions.ItemByKey(342).SetAttValue('DesSpeedDistr(1)' , SpeedDs[67])
	Vissim.Net.DesSpeedDecisions.ItemByKey(343).SetAttValue('DesSpeedDistr(1)' , SpeedDs[68])
	Vissim.Net.DesSpeedDecisions.ItemByKey(344).SetAttValue('DesSpeedDistr(1)' , SpeedDs[68])
	Vissim.Net.DesSpeedDecisions.ItemByKey(345).SetAttValue('DesSpeedDistr(1)' , SpeedDs[68])
	Vissim.Net.DesSpeedDecisions.ItemByKey(346).SetAttValue('DesSpeedDistr(1)' , SpeedDs[68])
	Vissim.Net.DesSpeedDecisions.ItemByKey(347).SetAttValue('DesSpeedDistr(1)' , SpeedDs[68])
	Vissim.Net.DesSpeedDecisions.ItemByKey(348).SetAttValue('DesSpeedDistr(1)' , SpeedDs[69])
	Vissim.Net.DesSpeedDecisions.ItemByKey(349).SetAttValue('DesSpeedDistr(1)' , SpeedDs[69])
	Vissim.Net.DesSpeedDecisions.ItemByKey(350).SetAttValue('DesSpeedDistr(1)' , SpeedDs[69])
	Vissim.Net.DesSpeedDecisions.ItemByKey(351).SetAttValue('DesSpeedDistr(1)' , SpeedDs[69])
	Vissim.Net.DesSpeedDecisions.ItemByKey(352).SetAttValue('DesSpeedDistr(1)' , SpeedDs[69])
	Vissim.Net.DesSpeedDecisions.ItemByKey(353).SetAttValue('DesSpeedDistr(1)' , SpeedDs[70])
	Vissim.Net.DesSpeedDecisions.ItemByKey(354).SetAttValue('DesSpeedDistr(1)' , SpeedDs[70])
	Vissim.Net.DesSpeedDecisions.ItemByKey(355).SetAttValue('DesSpeedDistr(1)' , SpeedDs[70])
	Vissim.Net.DesSpeedDecisions.ItemByKey(356).SetAttValue('DesSpeedDistr(1)' , SpeedDs[70])
	Vissim.Net.DesSpeedDecisions.ItemByKey(357).SetAttValue('DesSpeedDistr(1)' , SpeedDs[70])
	Vissim.Net.DesSpeedDecisions.ItemByKey(358).SetAttValue('DesSpeedDistr(1)' , SpeedDs[71])
	Vissim.Net.DesSpeedDecisions.ItemByKey(359).SetAttValue('DesSpeedDistr(1)' , SpeedDs[71])
	Vissim.Net.DesSpeedDecisions.ItemByKey(360).SetAttValue('DesSpeedDistr(1)' , SpeedDs[71])
	Vissim.Net.DesSpeedDecisions.ItemByKey(361).SetAttValue('DesSpeedDistr(1)' , SpeedDs[71])
	Vissim.Net.DesSpeedDecisions.ItemByKey(362).SetAttValue('DesSpeedDistr(1)' , SpeedDs[71])
	Vissim.Net.DesSpeedDecisions.ItemByKey(363).SetAttValue('DesSpeedDistr(1)' , SpeedDs[72])
	Vissim.Net.DesSpeedDecisions.ItemByKey(364).SetAttValue('DesSpeedDistr(1)' , SpeedDs[72])
	Vissim.Net.DesSpeedDecisions.ItemByKey(365).SetAttValue('DesSpeedDistr(1)' , SpeedDs[72])
	Vissim.Net.DesSpeedDecisions.ItemByKey(366).SetAttValue('DesSpeedDistr(1)' , SpeedDs[72])
	Vissim.Net.DesSpeedDecisions.ItemByKey(367).SetAttValue('DesSpeedDistr(1)' , SpeedDs[72])
	Vissim.Net.DesSpeedDecisions.ItemByKey(368).SetAttValue('DesSpeedDistr(1)' , SpeedDs[73])
	Vissim.Net.DesSpeedDecisions.ItemByKey(369).SetAttValue('DesSpeedDistr(1)' , SpeedDs[73])
	Vissim.Net.DesSpeedDecisions.ItemByKey(370).SetAttValue('DesSpeedDistr(1)' , SpeedDs[73])
	Vissim.Net.DesSpeedDecisions.ItemByKey(371).SetAttValue('DesSpeedDistr(1)' , SpeedDs[73])
	Vissim.Net.DesSpeedDecisions.ItemByKey(372).SetAttValue('DesSpeedDistr(1)' , SpeedDs[73])
	Vissim.Net.DesSpeedDecisions.ItemByKey(373).SetAttValue('DesSpeedDistr(1)' , SpeedDs[74])
	Vissim.Net.DesSpeedDecisions.ItemByKey(374).SetAttValue('DesSpeedDistr(1)' , SpeedDs[74])
	Vissim.Net.DesSpeedDecisions.ItemByKey(375).SetAttValue('DesSpeedDistr(1)' , SpeedDs[74])
	Vissim.Net.DesSpeedDecisions.ItemByKey(376).SetAttValue('DesSpeedDistr(1)' , SpeedDs[74])
	Vissim.Net.DesSpeedDecisions.ItemByKey(377).SetAttValue('DesSpeedDistr(1)' , SpeedDs[74])
	Vissim.Net.DesSpeedDecisions.ItemByKey(378).SetAttValue('DesSpeedDistr(1)' , SpeedDs[75])
	Vissim.Net.DesSpeedDecisions.ItemByKey(379).SetAttValue('DesSpeedDistr(1)' , SpeedDs[75])
	Vissim.Net.DesSpeedDecisions.ItemByKey(380).SetAttValue('DesSpeedDistr(1)' , SpeedDs[75])
	Vissim.Net.DesSpeedDecisions.ItemByKey(381).SetAttValue('DesSpeedDistr(1)' , SpeedDs[75])
	Vissim.Net.DesSpeedDecisions.ItemByKey(382).SetAttValue('DesSpeedDistr(1)' , SpeedDs[75])
	Vissim.Net.DesSpeedDecisions.ItemByKey(383).SetAttValue('DesSpeedDistr(1)' , SpeedDs[76])
	Vissim.Net.DesSpeedDecisions.ItemByKey(384).SetAttValue('DesSpeedDistr(1)' , SpeedDs[76])
	Vissim.Net.DesSpeedDecisions.ItemByKey(385).SetAttValue('DesSpeedDistr(1)' , SpeedDs[76])
	Vissim.Net.DesSpeedDecisions.ItemByKey(386).SetAttValue('DesSpeedDistr(1)' , SpeedDs[76])
	Vissim.Net.DesSpeedDecisions.ItemByKey(387).SetAttValue('DesSpeedDistr(1)' , SpeedDs[76])
	Vissim.Net.DesSpeedDecisions.ItemByKey(388).SetAttValue('DesSpeedDistr(1)' , SpeedDs[77])
	Vissim.Net.DesSpeedDecisions.ItemByKey(389).SetAttValue('DesSpeedDistr(1)' , SpeedDs[77])
	Vissim.Net.DesSpeedDecisions.ItemByKey(390).SetAttValue('DesSpeedDistr(1)' , SpeedDs[77])
	Vissim.Net.DesSpeedDecisions.ItemByKey(391).SetAttValue('DesSpeedDistr(1)' , SpeedDs[77])
	Vissim.Net.DesSpeedDecisions.ItemByKey(392).SetAttValue('DesSpeedDistr(1)' , SpeedDs[77])
	Vissim.Net.DesSpeedDecisions.ItemByKey(393).SetAttValue('DesSpeedDistr(1)' , SpeedDs[78])
	Vissim.Net.DesSpeedDecisions.ItemByKey(394).SetAttValue('DesSpeedDistr(1)' , SpeedDs[78])
	Vissim.Net.DesSpeedDecisions.ItemByKey(395).SetAttValue('DesSpeedDistr(1)' , SpeedDs[78])
	Vissim.Net.DesSpeedDecisions.ItemByKey(396).SetAttValue('DesSpeedDistr(1)' , SpeedDs[78])
	Vissim.Net.DesSpeedDecisions.ItemByKey(397).SetAttValue('DesSpeedDistr(1)' , SpeedDs[79])
	Vissim.Net.DesSpeedDecisions.ItemByKey(398).SetAttValue('DesSpeedDistr(1)' , SpeedDs[79])
	Vissim.Net.DesSpeedDecisions.ItemByKey(399).SetAttValue('DesSpeedDistr(1)' , SpeedDs[79])
	Vissim.Net.DesSpeedDecisions.ItemByKey(400).SetAttValue('DesSpeedDistr(1)' , SpeedDs[79])
			
def runSimulation():
	"""
	This is the simulation manager and starts and runs the simulation along with calling DMA Applications.

    """

	flag_read_additionally = False
	Vissim.LoadNet(network, flag_read_additionally)
	Vissim.LoadLayout(layout)
	Vissim.Simulation.SetAttValue('RandSeed', Random_Seed)
	
	tot_steps = Sim_Resolution * End_of_Simulation	
	for i in range(tot_steps):
		Vissim.Simulation.RunSingleStep()
		currtime = float(i)/Sim_Resolution
		
		if i == 1:
		#	print Vissim.Net.DesSpeedDecisions.ItemByKey(72).AttValue('DesSpeedDistr(1)')
			Vissim.Graphics.CurrentNetworkWindow.SetAttValue("QuickMode",1)

		if i == INFLOStart * Sim_Resolution and INFLO == True:
			print "\nINFLO Activated"
			
		tracker = currtime % INFLOFreq
		if INFLO == True:
			if tracker == 0 and currtime > INFLOStart and currtime < INFLOStop:
				InitializeINFLO(i)
				
		if i == INFLOStop * Sim_Resolution and INFLO == True:
			print "\nINFLO Dectivated"

def startup(seed, network, layout):
	"""
	This module sets up variables for individual simulations.

    """
	global Random_Seed, End_of_Simulation, Sim_Speed, Sim_Resolution
	global Pen_Rate, IZDelay, IZFreq, NumIncidents
	global INFLOStart, INFLOStop, INFLOFreq, INFLOSleep


	#Simulation Variables
	Random_Seed = seed
	End_of_Simulation = 21600
	Sim_Speed = 10
	Sim_Resolution = 3
	NumIncidents = 3
	
	INFLOStart = 3600
	INFLOStop = 16200
	INFLOFreq = 20
	INFLOSleep = 6
	
	runSimulation()


if __name__ == "__main__":	
	"""
	This main module starts the simulation manager and manages the database across multiple simulations.

    """
	global INC_ZONE, INFLO
	global network, layout

	"""
	Vissim = com.Dispatch("Vissim.Vissim-64.600")
	network = 'D:\\RKK\\AMSRuns\\AugustRuns\\Cluster1\\inflo_1_50MPv6.inpx'
	layout = 'D:\\RKK\\AMSRuns\\AugustRuns\\Cluster1\\inflo_1_50MPv6.layx'
	seed = 1
	INFLO = True
	startup(seed,network,layout)
	Vissim = None
	time.sleep(5)
	print "Simulation Completed"

	
	Vissim = com.Dispatch("Vissim.Vissim-64.600")
	network = 'D:\\RKK\\AMSRuns\\AugustRuns\\Cluster2\\inflo_2_50MP_v6.inpx'
	layout = 'D:\\RKK\\AMSRuns\\AugustRuns\\Cluster2\\inflo_2_50MP_v6.layx'
	seed = 1
	INFLO = True
	startup(seed,network,layout)
	Vissim = None
	time.sleep(5)
	print "Simulation Completed"
	


	#Cluster 3
	Vissim = com.Dispatch("Vissim.Vissim-64.600")
	network = 'D:\\RKK\\AMSRuns\\AugustRuns\\Cluster3\\inflo_3_50MPv6.inpx'
	layout = 'D:\\RKK\\AMSRuns\\AugustRuns\\Cluster3\\inflo_3_50MPv6.layx'
	seed = 1
	INFLO = True
	startup(seed,network,layout)
	Vissim = None
	time.sleep(5)
	print "Simulation Completed"
	"""

	Vissim = com.Dispatch("Vissim.Vissim-64.600")
	network = 'D:\\RKK\\AMSRuns\\AugustRuns\\Cluster4\\inflo_4_50MPv6.inpx'
	layout = 'D:\\RKK\\AMSRuns\\AugustRuns\\Cluster4\\inflo_4_50MPv6.layx'
	seed = 2
	INFLO = True
	startup(seed,network,layout)
	Vissim = None
	time.sleep(5)
	print "Simulation Completed"
	

	#Cluster 4
	Vissim = com.Dispatch("Vissim.Vissim-64.600")
	network = 'D:\\RKK\\AMSRuns\\AugustRuns\\Cluster4\\inflo_4_50MPv6.inpx'
	layout = 'D:\\RKK\\AMSRuns\\AugustRuns\\Cluster4\\inflo_4_50MPv6.layx'
	seed = 5
	INFLO = True
	startup(seed,network,layout)
	Vissim = None
	time.sleep(5)
	print "Simulation Completed"
	"""

	#Cluster 3
	Vissim = com.Dispatch("Vissim.Vissim-64.600")
	network = 'D:\\RKK\\AMSRuns\\AugustRuns\\Cluster3\\inflo_3_25MPv6.inpx'
	layout = 'D:\\RKK\\AMSRuns\\AugustRuns\\Cluster3\\inflo_3_25MPv6.layx'
	seed = 1
	INFLO = True
	startup(seed,network,layout)
	Vissim = None
	time.sleep(5)
	print "Simulation Completed"
	

	Vissim = com.Dispatch("Vissim.Vissim-64.600")
	network = 'D:\\RKK\\AMSRuns\\AugustRuns\\Cluster3\\inflo_3_10MPv6.inpx'
	layout = 'D:\\RKK\\AMSRuns\\AugustRuns\\Cluster3\\inflo_3_10MPv6.layx'
	seed = 1
	INFLO = True
	startup(seed,network,layout)
	Vissim = None
	time.sleep(5)
	print "Simulation Completed"
	"""

