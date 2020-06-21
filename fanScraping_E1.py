import PyPDF2, glob, os, csv, sys
from pandas import ExcelWriter
import pandas as pd

# Function to extract the pageContent ---------------------
def extractContent(pageNumber):
	pageObj = pdfReader.getPage(pageNumber)
	pageContent = pageObj.extractText()
	return pageContent
# ---------------------------------------------------------

# Function to get the range of pages of each unit ---------
def pagesFunction(keyword = 'Unit no.:'):
	aPageStart, aPageEnd = [], []
	last_page = number_of_pages - 1

	# Loop through all the pages
	for pageNumber in range(last_page):
		pageContent = extractContent(pageNumber)

		# Get the number of starting and ending pages
		if keyword in pageContent:
			if len(aPageStart) == 0:
				aPageStart.append(pageNumber)
			else:
				aPageEnd.append(pageNumber - 1)
				aPageStart.append(pageNumber)

	# Get the very last page
	aPageEnd.append(last_page)
	return aPageStart, aPageEnd

# Get SINGLE value function---------------------------------
def get_value_function(pageContent, wordStart, wordEnd, min_len = 1, max_len = 45):
	# wordStart and wordEnd are unique values, not lists
	posStart = pageContent.index(wordStart) + len(wordStart)
	newContent = pageContent[posStart:]

	#posEnd = indexFunction(wordEnd, newContent)
	posEnd = newContent.index(wordEnd)
	#unitFeature = newContent[:posEnd].strip()
	unitFeature = newContent[:posEnd]

	# Check the lenght in order to avoid errors
	if (len(unitFeature) > min_len) & (len(unitFeature) < max_len):
		return unitFeature
	else:
		print('Not valid feature. Content lenght is not ok!')
		return 'Error flag!'

# First page function -------------------------------------
def fpFunction():
	print('Starting first page function--------------------')
	inner_list, outter_list = [], []
	for page, pageEnd in zip(aPageStart, aPageEnd):
		print('Looking at', page, 'page')
		pageContent = extractContent(page)
		print('\n')

		# Get line
		wordStart = 'Unit no.'
		wordEnd = 'Fecha'
		line = get_value_function(pageContent, wordStart, wordEnd)

		# Reset ahu_value for each pageStart
		ahu_value = []

		# Get AHU type
		for ahu in ahus:
			if ahu in pageContent:
				ahu_value.append(ahu)

		# In case of DV10 or DV100, always get the longest one
		if len(ahu_value) == 1:
			ahu = ahu_value[0]
		elif len(ahu_value) > 1:
			print('Possible conflict!')
			final_ahu = ''
			for value in ahu_value:
				if len(value) > len(final_ahu):
					final_ahu = value
			ahu = final_ahu
			print('Final value:', ahu)
			print('\n')
		else:
			ahu = '---'

		# Get reference
		wordStart = 'Planta no.'
		wordEnd = 'Unit no.'
		ref = get_value_function(pageContent, wordStart, wordEnd)

		# Airflow
		wordStart = ')'
		wordEnd = 'm'
		airflow = get_value_function(pageContent, wordStart, wordEnd)

		inner_list = [page, pageEnd, line, ahu, ref, airflow]
		outter_list.append(inner_list)

	return outter_list

# Possible main--------------------------------------------
def extractFeatures(aWordStart, aWordEnd, pageStart, pageEnd, allowed_pages = 1):
	outter_list = []
	for page in range(pageStart, pageEnd):
		# Initiate the inner_list and get the page number
		inner_list = []
		
		# Extract page content
		pageContent = extractContent(page)
		print('Checking at page number', page+1)

		for wordStart, wordEnd in zip(aWordStart, aWordEnd):
			print('Looking for ', wordStart, 'and', wordEnd)

			# Work in starting and ending pairs, page by page
			if (wordStart in pageContent) and (wordEnd in pageContent):
				print('Found on page', page+1)
				print('\n')
				print(pageContent)
				print('\n')
				unitFeature = get_value_function(pageContent, wordStart, wordEnd)

				# Important in case the next wordStart is above the previos one
				print('Feature found:', unitFeature)
				split_word = unitFeature + wordEnd
				print('Split_word:', split_word)
				if split_word in pageContent:
					print('Split_word in pageContent')
					print('\n')
				posEnd = pageContent.index(split_word)
				pageContent = pageContent[posEnd:]

				if unitFeature == 'Error flag!':
					print('Error flag! Length not correct.')
					break
				else:
					inner_list.append(unitFeature)

			# If cheking of additional pages is allowed
			elif allowed_pages > 0:
				if len(inner_list) == 0:
					# Reset inner list
					inner_list = []
					print('No luck this time')
					print('\n')
					# Exit loop and go for the next page
					break
				elif len(inner_list) > 0:
					until_page = page + allowed_pages + 1
					for new_page in range(page + 1, until_page):
						pageContent = extractContent(new_page)
						try:
							print('\n')
							print('New page being checked')
							unitFeature = get_value_function(pageContent, wordStart, wordEnd)
							inner_list.append(unitFeature)
						except:
							print('No luck even in the next page')
							print('\n')
							# Exit loop and go for the next page
							break

		# Check the lenght and append to the outter list
		if len(inner_list) == len(aWordStart):
			print('New entry for the outter list!')
			print('\n')
			# Add the number of page in the inner list
			inner_list = [page + 1, *inner_list] # In order to show real page number
			outter_list.append(inner_list)
			# Reset inner list for next feature
			inner_list = []

	try:
		return outter_list
	except:
		print('No outter_list found!')
#----------------------------------------------------------

# Main ----------------------------------------------------
# Open file and read it
path = os.path.dirname(os.path.realpath(__file__))
num_files = len(glob.glob1(path,'*.pdf'))

extList = []
newList = []

# EC_FANS
excel_file = 'EC_FANS.xlsx'
df_ec = pd.read_excel(excel_file, dtype={'Item': str, 'Gross price': float})

# Empty dataframe
cols = ['Page', 'Airflow', 'Static Press.', 'Motor Power', 'RPM', 'Consump. kW',
		'Line', 'AHU', 'Ref', 'No Fans', 'ID', 'Gross price', 'File name']

df_outter = pd.DataFrame(columns=cols)

for fileName in glob.glob('*.pdf'):
	# Initialize ----------------------------------------------
	aDVSize = []
	aDVLine = []
	aPageStart = []

	ahus = ['DV10', 'DV15', 'DV20', 'DV25', 'DV30', 'DV40', 'DV50',
	'DV60', 'DV80', 'DV100', 'DV120', 'DV150', 'DV190', 'DV240',
	'Geniox 10', 'Geniox 11', 'Geniox 12', 'Geniox 14', 'Geniox 16',
	'Geniox 18', 'Geniox 20', 'Geniox 22', 'Geniox 24', 'Geniox 27',
	'Geniox 29']

	# Read the pdf --------------------------------------------
	fileName = fileName[:-4]
	pdfFileObj = open(fileName + '.pdf', 'rb')
	pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

	# Get number of pages
	number_of_pages = pdfReader.getNumPages()
	print('Number of pages:', number_of_pages)
	print('\n')

	# Get the range of the pages ------------------------------
	aPageStart, aPageEnd = pagesFunction()
	last_page = aPageEnd[-1:][0]
	print(aPageStart, aPageEnd)
	print('\n')

	units = fpFunction()
	columns_units = ['Page start', 'Page end', 'Line', 'AHU', 'Ref', 'Airflow']
	df_units = pd.DataFrame(units, columns = columns_units)

	# To get an accurate reading in the Excel file
	cols = ['Page start', 'Page end']
	for col in cols:
		df_units[col] = df_units[col].astype(int)
		df_units[col] = df_units[col] + 1

	'''
	name = 'Units.xlsx'
	writer = pd.ExcelWriter(name)
	df_units.to_excel(writer, index = False)
	writer.save()
	'''
	
	#---------------------------------------------------------------------------------------------------------------#
	# Fan scraping
	aWordStart = ['-fancaudal de aire', 'húmedas)', 'Potencia', 'Velocidad (nominal)', 'incl. el control de velocidad']
	aWordEnd = ['m', 'Pa', 'kW', 'RPM', 'kW']
	columns = ['Page', 'Airflow', 'Static Press.', 'Motor Power', 'RPM', 'Consump. kW']
	outter = extractFeatures(aWordStart, aWordEnd, 0, last_page, 1)
	df = pd.DataFrame(outter, columns = columns)
	df['Page'] = df['Page'].astype(int)

	#---------------------------------------------------------------------------------------------------------------#
	# Merge the fans with the size of the units
	cols = ['Line', 'AHU', 'Ref']
	line_list, ahu_list, ref_list = [], [], []
	list_of_lists = [line_list, ahu_list, ref_list]

	for fan_page in df['Page'].values:
		for col, list in zip(cols, list_of_lists):
			value = df_units.loc[(df_units['Page start'].values < fan_page) & (df_units['Page end'] > fan_page), col].values[0]
			list.append(value)

	for col, list in zip(cols, list_of_lists):
		df[col] = list
	#---------------------------------------------------------------------------------------------------------------#
	# Dataframe cleaning for several motors
	# Number of fans
	df['No Fans'] = 1
	print('df before fans:')
	print(df['Motor Power'])

	print('\n')
	df.loc[df['Motor Power'].str.contains('total'), 'No Fans'] = df['Motor Power'].str.slice(7, 8, 1)
	df['No Fans'] = df['No Fans'].astype(int)

	# Motor Power
	df.loc[df['Motor Power'].str.contains('nominal'), 'Motor Power'] = df['Motor Power'].str.slice(8)
	df.loc[df['Motor Power'].str.contains('total'), 'Motor Power'] = df['Motor Power'].str.slice(11)

	# ID - cleaning
	cols = ['Motor Power', 'RPM']

	for col in cols:
		df[col] = df[col].str.replace(" ", "")

	df['ID'] = df['Motor Power'] + '-' + df['RPM']

	# Show results
	print('df after fans:')
	print('\n')
	print(df)

	# Old Gross
	df = pd.merge(df, df_ec.loc[:, ['ID', 'Gross price']], on='ID')
	df['Gross price'] = df['Gross price'].values * df['No Fans'].values
	#---------------------------------------------------------------------------------------------------------------#

	# Last tweaks
	df['File name'] = fileName
	df_outter = df_outter.append(df)

	print('\n')

name = 'Fans per AHU.xlsx'
writer = pd.ExcelWriter(name)
df_outter.to_excel(writer, index = False)
writer.save()

pdfFileObj.close()

import fan_selection