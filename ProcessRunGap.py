#! /Users/mikeyb/Applications/python3

import zipfile
import json
import os,glob
import re
import datetime
import math
from ExerciseInfo_Class import ExerciseInfo
import applescript
import configparser


def breakTimeFromSeconds(totTimeSec):
	hourTot = math.floor(totTimeSec/60/60)
	minTot = math.floor((totTimeSec/60/60 - hourTot) * 60)
	secTot = math.floor(((totTimeSec/60/60 - hourTot) * 60 - minTot) *60)
	return hourTot, minTot, secTot
def formatNumbersTime(h, m, s):
	durTotNumbers = str(h) + 'h ' + str(m) + 'm ' + str(s) + 's'
	return durTotNumbers
def formatSheetsTime(h, m, s):
	durTotSheets = str(h) + ':' + str(m) + ':' + str(s)
	return durTotSheets


#######################################################
# MAIN
#######################################################
def main():
	# Get config details
	config = configparser.ConfigParser()
	config.read("../configs/newExerciseConfig.txt")
	
	pathToAppleScript = config['applescript']['script_path']
	appleScriptName = config['applescript']['sheet_name']

	runGapConfigs = config['rungap']
	monitorDir = runGapConfigs['monitor_dir']
	tempDir = runGapConfigs['temp_dir']
# 	dateTimeSheetFormat = runGapConfigs['date_time_sheet_format']
	dateTimeSheetFormat = '%m/%d/%Y %H:%M:%S'
	extraDetailsDir = runGapConfigs['extra_details_dir']

	compressFileRegex = re.compile(r'(.zip|.gz)$')
	jsonFileRegex = re.compile(r'(metadata.json)$')
	jsonExtRegex = re.compile(r'(.json)$')

	# ) Read applescript file for reading and updating exercise spreadseeht
	scptFile = open(pathToAppleScript + 'AddExercise.txt')
	scptTxt = scptFile.read()
	scpt = applescript.AppleScript(scptTxt)
	scpt.call('initialize',appleScriptName)

	zipFiles = [] 
	# Uncompress files from monitor directory into temp directory
	for filename in os.listdir(monitorDir):
		# Checks if compressed file
		if compressFileRegex.search(filename):
			z = zipfile.ZipFile(monitorDir + filename,mode='r')
			z.extractall(path=tempDir)
			zipFiles.append(monitorDir + filename)

	exLst = []
	# Loop through files and load exercise data to a list
	for filename in os.listdir(tempDir):
		if jsonFileRegex.search(filename):
			ex = ExerciseInfo()
			print('\nProcess ' + filename)
			with open(tempDir + filename) as data_file:
				data = json.load(data_file)
				ex.source = data['source']
				ex.originLoc = filename
			
				# Get the start time from file in UTC
				d = datetime.datetime.strptime(data['startTime']['time'],'%Y-%m-%dT%H:%M:%SZ')
				# Convert start time to current time zone
				sTime = d.replace(tzinfo=datetime.timezone.utc).astimezone(tz=None)
				ex.startTime = sTime
	
				MILES_IN_KILOMETERS = 0.621371
				METERS_IN_KILOMETERS = 1000
	
				eDistanceTotMeters = data['distance']
				ex.distTot = eDistanceTotMeters / METERS_IN_KILOMETERS * MILES_IN_KILOMETERS

				ex.hourTot, ex.minTot, ex.secTot = breakTimeFromSeconds(data['duration'])
				durTotNumbers = formatNumbersTime(ex.hourTot, ex.minTot, ex.secTot)
				durTotSheets = formatSheetsTime(ex.hourTot, ex.minTot, ex.secTot)
			
	# 			print("Duration: " + durTotNumbers)
	# 			print("Duration: " + durTotSheets)
				ex.durTot = durTotSheets	
			
				ex.avgHeartRate = data['avgHeartrate']
	
				ex.calTot = data['calories']
				
				if (runGapConfigs['print_data'] == 'Y'):
					print("Start Date Time: " + 
						ex.startTime.strftime('%Y-%m-%d %H:%M:%S %Z'))
					print("Distance: " + str(ex.distTot))
					print("Duration: " + ex.durTot)
					print('Avg Heartrate: ' + str(ex.avgHeartRate))
					print('Calories Burned: ' + str(ex.calTot))

				exLst.append(ex)

	# Save Exercise to spreadsheet then remove files
	for ex in exLst:
		startDateTime = ex.startTime.strftime(dateTimeSheetFormat)
		distance = "%.2f" % ex.distTot
		duration = formatNumbersTime(ex.hourTot, ex.minTot, ex.secTot)
		scpt.call('addExercise',ex.eDate, 'Running', duration, distance, ex.distUnit, ex.avgHeartRate, ex.calTot, ex.userNotes, startDateTime, ex.gear)

		# Remove files from temp folder then monitor folder
		fileNameChunks = ex.originLoc.split('.')
		fileNameStart = fileNameChunks[0]
		for fl in glob.glob(tempDir + fileNameStart + '*'):
			os.remove(fl)
		if (runGapConfigs['remove_files'] == 'Y'):
			for fl in glob.glob(monitorDir + fileNameStart + '*'):
				os.remove(fl)


	# Process files from ExtraRunData Directory
	for filename in os.listdir(extraDetailsDir):
		print(filename)
		# Skip files that do not have the json extension
		if not (jsonExtRegex.search(filename)):
			continue

		with open(extraDetailsDir + filename) as data_file:
			data = json.load(data_file)
		try:
			outfit = data['Outfit']
		except:
			outfit = ''
		try:
			shoes = data['Shoes']
		except:
			shoes = ''
		try:
			exDateStr = data['Date']
			print(exDateStr)
			exDate = datetime.datetime.strptime(exDateStr, '%b %d, %Y, %I:%M %p')
		except:
			exDate = datetime.datetime.now()
		try:
			notes = data['Notes']
		except:
			notes = ''

		# Find entry on spreadsheet for date
		exerciseLst = scpt.call('getExercises','10')
		bestMatchEx = ''
	
		for ex in exerciseLst:
	# 		print('New Exercise Details Date: ' + exDate.strftime('%Y-%m-%d %H:%M:%S %Z') + ' Existing Exercise Date: ' + ex['dateVal'].strftime('%Y-%m-%d %H:%M:%S %Z') )
			existingExDate = ex['dateVal']
	# 		if exDate.date() == existingExDate.date():
			timeDiff = abs(exDate - existingExDate)
			if bestMatchEx == '':
				ex['timeDiff'] = timeDiff
				bestMatchEx = ex
			else:
				if bestMatchEx['timeDiff'] > timeDiff:
					ex['timeDiff'] = timeDiff
					bestMatchEx = ex
		print('Best Match: row ' + str(bestMatchEx['rowVal']) + ' date ' + bestMatchEx['dateVal'].strftime('%Y-%m-%d %H:%M:%S %Z'))

		# Update values on spreadsheet
		if bestMatchEx['gearVal'] == '':
			bestMatchEx['gearVal'] = shoes
		if bestMatchEx['noteVal'] == '':
			bestMatchEx['noteVal'] = outfit + '. ' + notes

		print(str(bestMatchEx))

		h, m, s = breakTimeFromSeconds(bestMatchEx['durationVal'])

		scpt.call('updateExercise', bestMatchEx['rowVal'], bestMatchEx['dateVal'],
			bestMatchEx['typeVal'],
			formatNumbersTime(h, m, s),
			bestMatchEx['distanseVal'],'MI', 
			bestMatchEx['hrVal'], bestMatchEx['caloriesVal'], 
			bestMatchEx['noteVal'], bestMatchEx['gearVal'])
		
		if (runGapConfigs['remove_files'] == 'Y'):
			os.remove(extraDetailsDir + filename)

if __name__ == '__main__':
	main()

