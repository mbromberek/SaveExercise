#! /Users/mikeyb/Applications/python3

# pip3 install -r requirements.txt
# pip3 install --upgrade google-api-python-client

import os, platform
import datetime, time

import applescript
import GoogleConnection, GoogleConnectionEmail 
from apiclient import discovery
import httplib2, base64
import math
import configparser, json
import pandas as pd

# Custom Classes
from summaries.ActivitySummary_Class import ActivitySummary
from summaries.ExerciseInfo_Class import ExerciseInfo
from util.Util import Util
import PrintData


#needed for attachment
import smtplib  
import mimetypes
from email import encoders
from email.message import Message
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


config = configparser.ConfigParser()

#######################################################
# Makes connection to Google spreadsheet
# gets all Activity data from google sheet, and converts
# the data into a DataFrame returning it. 
#######################################################
def getActivityData(config):
	gc = GoogleConnection
	credentials = gc.get_credentials()#might need to pass application name
	http = credentials.authorize(httplib2.Http())
	discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
					'version=v4')
	serviceSheet = discovery.build('sheets', 'v4', http=http,
							  discoveryServiceUrl=discoveryUrl)

	spreadsheetId = config['google_sheet']['spreadsheet_id']
	sheetId = config['google_sheet']['sheet_id']
	sheetName=config['google_sheet']['sheet_name']

	activityHdr, activityData = getActivities(serviceSheet, spreadsheetId, sheetName)	
	actvDf = generateActivityDataFrame(activityHdr, activityData)
	
	return actvDf

#######################################################
# Get activity details from Google Sheet
# 	Split out the header and data from results
# Return: header and data
#######################################################
def getActivities(service, spreadsheetId, sheetName):
	rangeName = sheetName+'!A:I'
	result = service.spreadsheets().values().get(
		spreadsheetId=spreadsheetId, range=rangeName).execute()
	values = result.get('values', [])
	
	# Convert headers to lowercase and replace spaces with _
	activityHdr = [x.lower().replace(' ','_') for x in values[0]]
	activityData = values[1:]
		
	return activityHdr, activityData

#######################################################
# Using past in header and data create DataFrame
# 	Convert columns to correct formats
# 	Generate Index
# 
# Return: generated DataFrame
#######################################################
def generateActivityDataFrame(hdr, data):
	numericCols = ['distance', 'steps','floors_climbed', 'tot_cal_burned','active_calories', 'weight_(lb)']
	dateCols = ['date']
	df = pd.DataFrame(data,columns=hdr)
	
	# Set number and date columns to numeric and datetime formats
	df[numericCols] = df[numericCols].apply(pd.to_numeric, errors='coerce')
	df[dateCols] = df[dateCols].apply(pd.to_datetime, errors='coerce')

	# Create dateStr column as a string of the date column and make it the index
	df['dateStr'] = df['date'].dt.strftime('%Y-%m-%d')
	df = df.set_index('dateStr')
	
	return df

#######################################################
# Uses AppleScript to get contents of Exercise spreadsheet. 
# Converts the data into a DataFrame and returns it. 
#######################################################
def getExerciseData(config):
	pathToAppleScript = config['applescript']['script_path']
	spreadsheetName = config['exercise']['spreadsheet_name']
	# ) Read applescript file for reading and updating exercise spreadseeht
	scptFile = open(pathToAppleScript + 'AddExercise.txt')
	scptTxt = scptFile.read()
	scpt = applescript.AppleScript(scptTxt)
	scpt.call('initialize',spreadsheetName)

	exLst = scpt.call('getExercises',config['exercise']['number_exercises_to_read'])
	exDf = generateExerciseDataFrame(exLst)
	
	return exDf

#######################################################
# Creates DataFrame based on list of exercise data
# Sets which fields are numeric and which are date types.
# Sorts data by date
#######################################################
def generateExerciseDataFrame(exLst):
	numericCols = ['distanseVal', 'caloriesVal','durationVal', 'hrVal']
	dateCols = ['dateVal']

	exDf = pd.DataFrame(exLst)	
	
	exDf[numericCols] = exDf[numericCols].apply(pd.to_numeric, errors='coerce')
	exDf[dateCols] = exDf[dateCols].apply(pd.to_datetime, errors='coerce')

	# Create dateStr column as a string of the date column and make it the index
	exDf['dateStr'] = exDf['dateVal'].dt.strftime('%Y-%m-%d')
	exDf = exDf.set_index('dateStr')
	
	# Need to sort for the loc function to work in exSummary
	exDf = exDf.sort_values(by=['dateVal'])

	return exDf

#######################################################
# 
#######################################################
def exSummary(df,  dtStart, dtEnd):
	dfRange = df.loc['{:%Y-%m-%d}'.format(dtStart):'{:%Y-%m-%d}'.format(dtEnd)]
	
	ex = {}
	ex['startDate'] = dtStart
	ex['endDate'] = dtEnd
	ex['totDays'] = (dtEnd - dtStart).days
	
	# TODO make the types totally configurable so adding elliptical would just be config
	runTypLst = config['types']['run'].split(',')
	cycleTypLst = config['types']['cycle'].split(',')
	swimTypLst = config['types']['swim'].split(',')

	ex['run'] = ExerciseInfo('Running')
	ex['run'].startTime = dtStart
	ex['run'].endTime = dtEnd
	ex['run'].distTot = dfRange[dfRange['typeVal'].isin(runTypLst)].sum()['distanseVal']
	ex['run'].durTot = dfRange[dfRange['typeVal'].isin(runTypLst)].sum()['durationVal']

	ex['run'].ct = dfRange[dfRange['typeVal'].isin(runTypLst)].count()['typeVal']
	ex['run'].avgEasyPace = dfRange[(dfRange['typeVal'].isin(runTypLst)) &	(dfRange['catVal']=='Easy')].mean()['paceVal']
	ex['run'].avgEasyHr = dfRange[(dfRange['typeVal'].isin(runTypLst)) &	(dfRange['catVal']=='Easy')].mean()['hrVal']

	ex['run'].avgLongDist = dfRange[(dfRange['typeVal'].isin(runTypLst)) &	(dfRange['catVal']=='Long Run')].mean()['distanseVal']
	ex['run'].avgLongPace = dfRange[(dfRange['typeVal'].isin(runTypLst)) &	(dfRange['catVal']=='Long Run')].mean()['paceVal']
	ex['run'].avgLongHr = dfRange[(dfRange['typeVal'].isin(runTypLst)) &	(dfRange['catVal']=='Long Run')].mean()['hrVal']
	ex['run'].avgLongDur = dfRange[(dfRange['typeVal'].isin(runTypLst)) &	(dfRange['catVal']=='Long Run')].mean()['durationVal']
	
	
	ct = dfRange[dfRange['typeVal'].isin(swimTypLst)].count()['typeVal']
	if (ct > 0):
		ex['swim'] = ExerciseInfo('Swimming')
		ex['swim'].distTot = dfRange[dfRange['typeVal'].isin(swimTypLst)].sum()['distanseVal']
		ex['swim'].durTot = dfRange[dfRange['typeVal'].isin(swimTypLst)].sum()['durationVal']
		ex['swim'].ct = ct

	ct = dfRange[dfRange['typeVal'].isin(cycleTypLst)].count()['typeVal']
	if (ct > 0):
		ex['cycle'] = ExerciseInfo('Cycling')
		ex['cycle'].distTot = dfRange[dfRange['typeVal'].isin(cycleTypLst)].sum()['distanseVal']
		ex['cycle'].durTot = dfRange[dfRange['typeVal'].isin(cycleTypLst)].sum()['durationVal']
		ex['cycle'].ct = ct
	
	return ex

#######################################################
# Gets totals for passed activity for the data between
# passed dtStart and dtEnd
# Then gets summaries of exercises split into run, swim,
#  and cycle
#######################################################
def calcSummary(actvDf, exDf, dtStart, dtEnd):
	dfRange = actvDf.loc['{:%Y-%m-%d}'.format(dtStart):'{:%Y-%m-%d}'.format(dtEnd)]

	activ = ActivitySummary(dtStart, dtEnd)	
	activ.totDays = (dtEnd - dtStart).days
	
	activ.totSteps = dfRange['steps'].sum()
	activ.totDist = dfRange['distance'].sum()
	activ.totActiveCal = dfRange['active_calories'].sum()
	activ.totFloors = dfRange['floors_climbed'].sum()
	
	exerciseSummaries = exSummary(exDf, dtStart, dtEnd)
	if ('run' in exerciseSummaries):
		activ.exRun = exerciseSummaries['run']
	if ('swim' in exerciseSummaries):
		activ.exSwim = exerciseSummaries['swim']
	if ('cycle' in exerciseSummaries):
		activ.exCycle = exerciseSummaries['cycle']
	
	return activ

#######################################################
# Get start and end dates of previous week and 
# number of weeks past the previous week based on 
# passed numWeeks value
# Return: List of Dictionary for start and end of weeks 
#######################################################
def getWeeksStartEnd(numWeeks):
	actvWeeks = []
	actvWeeks.append({})
	actvWeeks[0]['end'] = Util.getPreviousSunday(datetime.date.today())
	actvWeeks[0]['start'] = actvWeeks[0]['end'] - datetime.timedelta(days=6)
	for i in range(1,numWeeks):
		actvWeeks.append({})
		actvWeeks[i]['end'] = actvWeeks[i-1]['start'] - datetime.timedelta(days=1)
		actvWeeks[i]['start'] = actvWeeks[i]['end'] - datetime.timedelta(days=6)

	return actvWeeks


#######################################################
# Makes email connection with Google
# Sends the passed email message
#######################################################
def sendEmail(config, msg):
	# Setup Email connection
	gcEmail = GoogleConnectionEmail
	credentialsEmail = gcEmail.get_credentials()
	httpEmail = credentialsEmail.authorize(httplib2.Http())
	serviceEmail = discovery.build('gmail', 'v1', http=httpEmail)
	
	# Send Email
# 	subj = 'Activity Analysis ' + '{:%Y-%m-%d}'.format(datetime.date.today())
	subj = config['email']['subject'].replace('~date~', '{:%Y-%m-%d}'.format(datetime.date.today()))
	
	if (config['email']['type'] == 'HTML'):
		msgRaw = createMessageHtml(config['email']['srcEmail'],config['email']['destEmail'], subj, msg)
	else:
		msgRaw = gcEmail.create_message(config['email']['srcEmail'],config['email']['destEmail'], subj, msg)

	
	sentMsg = gcEmail.send_message(serviceEmail, config['email']['srcEmail'], msgRaw)
	
	return sentMsg

def createMessageHtml(sender, to, subj, msg):
	message = MIMEMultipart('alternative') # needed for both plain & HTML (the MIME type is multipart/alternative)
	message['Subject'] = subj
	message['From'] = sender
	message['To'] = to

    #Create the body of the message (a plain-text and an HTML version)
#     message.attach(MIMEText(message_text_plain, 'plain'))
	message.attach(MIMEText(msg, 'html'))

	raw_message_no_attachment = base64.urlsafe_b64encode(message.as_bytes())
	raw_message_no_attachment = raw_message_no_attachment.decode()
	body  = {'raw': raw_message_no_attachment}
	return body


#######################################################
# MAIN
#######################################################
def main():
	# Get config details
	progDir = os.path.dirname(os.path.abspath(__file__))	
	config.read(progDir + "/../configs/activityAnalysisConfig.txt")
	
	
	numWeeks = int(config['activity']['number_weeks'])
	if (numWeeks < 2):
		numWeeks = 2
	actvWeeks = getWeeksStartEnd(numWeeks)

	actvDf = getActivityData(config)	
	exDf = getExerciseData(config)
		
	actvSummaryLst = []
	for i in range(len(actvWeeks)):
		actvSummaryLst.append(calcSummary(actvDf, exDf, actvWeeks[i]['start'], actvWeeks[i]['end']))

	if (config['email']['type'] == 'HTML'):
		msg = PrintData.generateHtmlSummary(actvSummaryLst)
	else:
		msg = PrintData.generateSummary(actvSummaryLst)
	
	print(msg)
		
	sentMsg = sendEmail(config, msg)
	print ('Message Id: %s' % sentMsg['id'])
	
	return actvSummaryLst


if __name__ == '__main__':
	main()

