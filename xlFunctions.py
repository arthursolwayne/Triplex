import xlwings as xw
from datetime import date

#for getImpMove
import pyautogui
import numpy as np
from PIL import ImageGrab
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
from textblob import TextBlob
pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

@xw.func
def convertToRTD (attr, ticker):
	return '=RTD("tos.rtd", , "'+attr+'", "'+ticker+'")'

@xw.func
def getShortStrike (ticker):
	if 'C' in ticker:
		return ticker[ticker.find('C')+1:ticker.find('-')]
	else:
		return ticker[ticker.find('-')-3:ticker.find('-')]

@xw.func
def getLongStrike(ticker):
	if 'C' in ticker:
		return ticker[ticker.rfind('C')+1:]
	else:
		return ticker[ticker.rfind('P')+1:]

@xw.func
def getSide(ticker):
	if 'C' in ticker:
		return 'Call'
	else:
		return 'Put'

@xw.func
def getShortStrikeTicker (ticker):
	return ticker[ticker.find('.'):ticker.find('-')]

@xw.func
def getLongStrikeTicker (ticker):
	return ticker[ticker.find('-')+1:]

@xw.func
def getDTE (expir):
	exp = date(int(expir[0:4]),int(expir[5:7]),int(expir[8:]))
	today = date.today()
	dte = exp - today
	return dte.days

@xw.func
def getExp (ticker):
	return '20'+ticker[4:6]+'-'+ticker[6:8]+'-'+ticker[8:10]

@xw.func
def getImpMove ():
	# Grab some screen
	screen =  pyautogui.screenshot(region = (3592,1088,246,38))
	# Make greyscale
	w = screen.convert('L')
	# Save so we can see what we grabbed
	w.save('grabbed.png')
	w.show()
	text = pytesseract.image_to_string(w)
	correctedText = TextBlob(text).correct()
	print(correctedText)

 #  tl 3592 1088
 #  tr 3838 1088
 #  bl 3592 1050
 #  br 3838 1050
