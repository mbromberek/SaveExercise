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

import ExtraExerciseDetails
import ProcessRunGap
import RunKeeperSaveNewExercises

def main():
	ProcessRunGap.main()
	RunKeeperSaveNewExercises.main()
	ExtraExerciseDetails.main()

if __name__ == '__main__':
	main()

