#! /Users/mikeyb/Applications/python3

# TODO Need to setup ability to control web browser


import pyperclip, os
import applescript
from selenium import webdriver
from ExerciseInfo_Class import ExerciseInfo
import RunkeeperSiteControls
from datetime import datetime
import configparser

import GoogleConnection
from apiclient import discovery
import httplib2
import time

SLEEP_TIME = 0

#######################################################
# Break up the passed in time and return results, 
#  returns 0 for values not in the passed string
#######################################################
def splitTime(t):
	times = t.split(':')
	h = '0'
	m = '0'
	s = '0'
	if len(times) > 0:
		s = times[-1]
		if len(times) > 1:
			m = times[-2]
			if len(times) > 2:
				h = times[-3]
	return h,m,s

#######################################################
# Makes a web service call to passed spreadsheet to add
#  the passed exercise below the last line in the sheet
# service - Google service to connect to
# spreadsheetId - ID of spreadsheet to update
# sheetName - tab of spreadsheet to update
# ex - exercise data to load into the spreadsheet
#######################################################
def addExerciseToGoogleSheet(service, spreadsheetId, sheetName, ex):
	# 1) Get last populated row in sheet
	rangeName = sheetName+'!A:B'
	result = service.spreadsheets().values().get(
		spreadsheetId=spreadsheetId, range=rangeName).execute()
	values = result.get('values', [])

	# 2) If nothing found use row 1
	if not values:
		print('No data found.')
		lastRowNum = 1
		lastRowStr = str(lastRowNum)			
	else:
		lastRowNum = len(values)+1
		lastRowStr = str(lastRowNum)
		print('Number rows: ' + lastRowStr)

	# 3) Generate web service JSON to send
	batchUpdateData = [
		{
			'range': sheetName+'!A'+lastRowStr+':I'+lastRowStr, 
			'majorDimension': 'ROWS', 
			'values': [
				[
					ex.eDate, ex.type, 
					ex.hourTot + ':' + ex.minTot + ':' + ex.secTot,
					'=E'+lastRowStr+'*1.60934', ex.distTot, ex.avgHeartRate,
					'=C'+lastRowStr+'/E'+lastRowStr,
					ex.calTot,
					ex.userNotes
				]
			]
		}
	]
	batchUpdateBody = {'valueInputOption': 'USER_ENTERED', 'data': batchUpdateData}

	service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId,
	 body=batchUpdateBody).execute() 


#######################################################
# MAIN
#######################################################
def main():
	# Get config details
	config = configparser.ConfigParser()
	config.read("../configs/newExerciseConfig.txt")

	# Enter user name and password for runkeeper
	runKeeperConfigs = config['runkeeper']
	uName = runKeeperConfigs['user_name']
	pWord = runKeeperConfigs['password']
	exerciseType = 'Running'
	
	pathToAppleScript = config['applescript']['script_path']
	appleScriptName = config['applescript']['sheet_name']

	rk = RunkeeperSiteControls
	gc = GoogleConnection

	#Connect to Google Sheet
	credentials = gc.get_credentials()#might need to pass application name
	http = credentials.authorize(httplib2.Http())
	discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
					'version=v4')
	service = discovery.build('sheets', 'v4', http=http,
							  discoveryServiceUrl=discoveryUrl)

	#ID for Google spreadsheet to update
	spreadsheetId = config['google_sheet']['sheet_id']
	sheetName= config['google_sheet']['sheet_name']

	# 1) Determine type to save and clean up Monitor folder if any files in there
	monitorDir = config['directories']['monitor_directory']
	for filename in os.listdir(monitorDir):
		if filename.startswith('Run'):
			exerciseType = 'Running'
		elif filename.startswith('Cycl'):
			exerciseType = 'Cycling'
		os.unlink(monitorDir + filename)

	print('Exercise Type to save: ' + exerciseType)

	# 2) Read applescript file for reading and updating exercise spreadseeht
	scptFile = open(pathToAppleScript + 'AddExercise.txt')
	scptTxt = scptFile.read()
	scpt = applescript.AppleScript(scptTxt)
	scpt.call('initialize',appleScriptName)

	# 3) Read last entered date/time for exercise type
	lastExDtTm = scpt.call('readLastExercise', exerciseType, 'desc')
	print('Last recorded run was on: ' + str(lastExDtTm))
	leDate = lastExDtTm.date()

	# 4) Login to Runkeeper web page
	browser = rk.loginToRunkeeper(uName, pWord)

	# 5) Get list of exercises including and since last run date
	exLst = rk.getExerciseSince(leDate, exerciseType, browser)

	# 6) Loop through returned list backwards
	for i in range(len(exLst)-1, -1, -1):

	# 7) Get element for current date and type and get exercise details
		nxtElem = browser.find_element_by_partial_link_text(exLst[i])
		nxtElem.click()
		ex = rk.getExerciseDetails(browser, exerciseType)

	# 8) Append zero to the hour if hour is one digit
		if (len(ex.startTime.split(':')[0]) == 1):
			ex.startTime = '0' + ex.startTime
		exDateTime = datetime.strptime(ex.eDate + ' ' + ex.startTime, '%m/%d/%Y %I:%M %p')

	# 9) If exercise date/time is greater than last date/time
		if (exDateTime > lastExDtTm):
			ex.hourTot,ex.minTot,ex.secTot = splitTime(ex.durTot)

	# 10) Add exercise entry to spreadsheet
			scpt.call('addExercise',ex.eDate, exerciseType, 
				ex.hourTot + 'h ' + ex.minTot + 'm ' + ex.secTot + 's', ex.distTot, 
				ex.distUnit, ex.avgHeartRate, ex.calTot, ex.userNotes, ex.startTime, 
				ex.gear)
	# 11) Add exercise entry to Google Sheet
			addExerciseToGoogleSheet(service, spreadsheetId, sheetName, ex)

		else:
			print('Exercise at ' + ex.eDate + ' ' + ex.startTime + 
			' is already in the past')

	browser.quit()


if __name__ == "__main__":
	main()

