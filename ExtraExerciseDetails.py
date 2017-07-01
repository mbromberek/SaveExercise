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



def main():
	# Get config details
	progDir = os.path.dirname(os.path.abspath(__file__))	
	config = configparser.ConfigParser()
	config.read(progDir + "/../configs/newExerciseConfig.txt")
	
	pathToAppleScript = config['applescript']['script_path']
	appleScriptName = config['applescript']['sheet_name']

	runGapConfigs = config['rungap']
	extraDetailsDir = runGapConfigs['extra_details_dir']

	compressFileRegex = re.compile(r'(.zip|.gz)$')
	jsonFileRegex = re.compile(r'(metadata.json)$')
	jsonExtRegex = re.compile(r'(.json)$')

	# ) Read applescript file for reading and updating exercise spreadseeht
	scptFile = open(pathToAppleScript + 'AddExercise.txt')
	scptTxt = scptFile.read()
	scpt = applescript.AppleScript(scptTxt)
	scpt.call('initialize',appleScriptName)



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

