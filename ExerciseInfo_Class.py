#! /Users/mikeyb/Applications/python3
#CycleInfo_Class

class ExerciseInfo:
	eDate = ''
	type = ''
	startTime = ''
	distTot = ''
	distUnit = ''
	
	durTot = ''
	hourTot = ''
	minTot = ''
	secTot = ''
	
	rating = ''
	
	calTot = ''
	avgHeartRate = ''
	userNotes = ''
	
	temperature = ''
	
	gear = ''
	
	source = ''
	
	# Filename the data came from or link to data
	originLoc = ''
	

	def __init__(self):
		self.epName = ''
	
# 	def __str__(self):
# 		print (epName)

	def cycleDateTime():
		return self.runDate + ' ' + self.runTime
