#! /Users/mikeyb/Applications/python3

#######################################################
#RunkeeperSiteControls
#######################################################

from selenium import webdriver
from datetime import datetime
import time
from selenium.webdriver.chrome.options import Options

monthMap = {'JAN':'01', 'FEB':'02', 'MAR':'03', 'APR':'04', 'MAY':'05', 'JUN': '06', 'JUL': '07', 'AUG':'08', 'SEP':'09', 'OCT':'10', 'NOV':'11','DEC':'12'}
SLEEP_TIME = 5


#######################################################
# Login to Runkeeper website using the past user name 
#  and password
#######################################################
def loginToRunkeeper(uName, pWord):
# 	browser = webdriver.Firefox()
	chrome_options = Options()
	chrome_options.add_experimental_option('prefs', {
		'credentials_enable_service': False,
		'profile': {
			'password_manager_enabled': False
		}
	})

	browser = webdriver.Chrome(
	executable_path='/Users/mikeyb/Dropbox/code/python/WebDrivers/chromedriver',
	chrome_options=chrome_options)

	browser.get('https://runkeeper.com/login')
	emailElem = browser.find_element_by_id('loginEmail')
	emailElem.send_keys(uName)
	passElem = browser.find_element_by_id('loginPassword')
	passElem.send_keys(pWord)
	passElem.submit()
	
	return browser

#load the exercise feed page
def gotoExerciseFeed(browser):
	# Select the first entry in Feed
	# div class feedItemContainer
	# a class "usernameLinkNoSpace user-name"
	linkElem = browser.find_element_by_class_name('feedItemContainer')
	type(linkElem)
	linkElem.click()


#######################################################
# getExerciseDetails
# Get the cycle details from the past in web page
# Returns CycleInfo class of read in data
#######################################################
def getExerciseDetails(browser, exerciseType):
	from ExerciseInfo_Class import ExerciseInfo
	exc = ExerciseInfo()
	time.sleep(1)
	
	# 1) Read date and time of run from page
	try:
		headerElem = browser.find_element_by_class_name('userHeader')
		type(headerElem)
		dateTimeElem = headerElem.find_element_by_class_name('activitySubTitle')
		type(dateTimeElem)
		rDateTimeTxt = dateTimeElem.text.split(' - ')
		rDateTxt = rDateTimeTxt[0].strip()
		rStartTime = rDateTimeTxt[1].strip()
	except:
		#NOT SURE CAUSE OF THIS BUT MIGHT BE PAGE LOAD TIME, SETTING UP A LOOP TO TRY ONE OR TWO TIMES WITH A time.sleep MIGHT FIX THE ISSUE
		print('Date and Time elements missing')
		rDateTxt = 'JAN 01, 2020'
		rStartTime = '1:00 AM'
	
	# 2) Break up the read in date and format it
	rDateLst = rDateTxt.split(' ')
	runMonthTxt = rDateLst[0]
	runMonth = monthMap[runMonthTxt]
	runDay = rDateLst[1].strip(',')
	runYear = rDateLst[2]

	exc.eDate = '/'.join([runMonth,runDay,runYear])
	exc.startTime = rStartTime
	exc.type = exerciseType
	
	# 3) Read run measurements
	try:
		totDistanceElem = browser.find_element_by_id('totalDistance')
		type(totDistanceElem)
		exc.distTot = totDistanceElem.find_element_by_tag_name('span').text
		exc.distUnit = totDistanceElem.find_element_by_tag_name('h5').text
	except:
		print('Distance Element missing')
		exc.distTot = 0
		exc.distUnit = 0
		
	try:
		totDurationElem = browser.find_element_by_id('totalDuration')
		type(totDurationElem)
		exc.durTot = totDurationElem.find_element_by_tag_name('span').text
	except:
		print('Duration Element missing')
		exc.durTot = 0
	

	try:
		avgPaceElem = browser.find_element_by_id('averagePace')
	except:
		try:
			avgPaceElem = browser.find_element_by_id('averageSpeed')
		except:
			print('Average Pace/Speed element missing')

	try:
		totCaloriesElem = browser.find_element_by_id('totalCalories')
		type(totCaloriesElem)
		exc.calTot = totCaloriesElem.find_element_by_tag_name('span').text
	except:
		print('Total Calories element missing')
		exc.calTot = 0
	
	try:
		hrElem = browser.find_element_by_id('heartRate')
		type(hrElem)
		exc.avgHeartRate = hrElem.find_element_by_class_name('value').text
	except:
		print('Heart Rate Elemenet missing')
		exc.avgHeartRate = 0
	
	exc.rating = '3'
	exc.gear = ''

	userNotesElem = browser.find_element_by_id('userNotes')
	exc.userNotes = userNotesElem.text
	
	return exc



def getExerciseDate(browser):
	#Read date of run
	headerElem = browser.find_element_by_class_name('userHeader')
	type(headerElem)
	divs = headerElem.find_elements_by_tag_name('div')
	dateTm = divs[-1].text
	return dateTm


#######################################################
# Get passed exercises including and since passed date
# Returns: list of elements with exercises
#######################################################
def getExerciseSince(lastExDt, eType, browser):
	
	excToReturnLst = []
	
	currYr = str(datetime.now().year)
	
	# 1: Get to the activities page
	getActivitiesPage(browser)
	
	# 2: Loop through the activities, if date is after 
	leftSectionElem = browser.find_element_by_class_name('leftContentColumn')
	type(leftSectionElem)
	historyElem = leftSectionElem.find_element_by_id('activityHistoryMenu')
	type(historyElem)
	latestExElems = historyElem.find_elements_by_partial_link_text(eType)
	for i in range(len(latestExElems)):
		eDateStr = latestExElems[i].find_element_by_class_name('startDate').text + '/' + currYr
		eDate = datetime.strptime(eDateStr, '%m/%d/%Y')
		if eDate.date() >= lastExDt:
			excToReturnLst.append(latestExElems[i].text.rsplit('\n')[0])
		else:
			break
	print(excToReturnLst)
	return excToReturnLst

#######################################################
# Get to the activities page by selecting ME tab
#  than selecting the history element
#######################################################
def getActivitiesPage(browser):
	time.sleep(SLEEP_TIME) #pause to let site finish loading
	# 1) Select the ME Tab
	meDivElem = browser.find_element_by_class_name('me')
	type(meDivElem)
	meDivElem.click()

	time.sleep(SLEEP_TIME)
	# 2) Select Activities Tab
	activitiesDivElem = browser.find_element_by_class_name('history')
	type(activitiesDivElem)
	activitiesDivElem.click()
	time.sleep(SLEEP_TIME)
	

#######################################################
# Load the latest run page
#######################################################
def getLatestExercisePageOfType(exercise, browser):

	getActivitiesPage(browser)

	# 1) Select first activity with Running in the text
	# This gets the latest run since the exercises are in descending order
	leftSectionElem = browser.find_element_by_class_name('leftContentColumn')
	type(leftSectionElem)
	historyElem = leftSectionElem.find_element_by_id('activityHistoryMenu')
	type(historyElem)
	latestRunElem = historyElem.find_element_by_partial_link_text(exercise)
	type(latestRunElem)
	latestRunElem.click()
