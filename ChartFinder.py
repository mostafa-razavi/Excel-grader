import os
import shutil
import re
import zipfile
import ntpath

#import Excellent

def get_match_number(fileName,string,mode):
	count = 0
	file = open(fileName, 'r')
	if mode == 1:	# *word*
		regex = "" + string + ""
	for line in file.readlines():
		if re.search( regex, line, re.I):
			count = count + len(re.findall(regex, line))
	return count

def unzip(bookAddress):
	unzipped_directory = os.path.dirname(os.path.abspath(bookAddress)) + "/temp"
	unzipped_directory = str(unzipped_directory)
	zip_ref = zipfile.ZipFile(bookAddress, 'r')
	zip_ref.extractall(unzipped_directory)
	zip_ref.close()
	return unzipped_directory

def delete_directory(directory):
	shutil.rmtree(directory)
	return

def get_charts_list_in_sheet(bookAddress,sheet_number,unzipped_directory):

	#unzipped_directory = unzip(bookAddress)

	xl_dir = unzipped_directory + "/xl"

	charts_dir = xl_dir + "/charts"
	drawings_dir = xl_dir + "/drawings"
	worksheets_dir = xl_dir + "/worksheets"

	drawingrels_dir = drawings_dir + "/_rels"
	sheetrels_dir = worksheets_dir + "/_rels"

	charts_paths_list = []
	for subdir, dirs, files in os.walk(charts_dir):
	    for file in sorted(files):
		ext = os.path.splitext(file)[-1].lower()
		if ext == ".xml":
			if file[0:5] == "chart":
				chart_path_name = subdir + "/" + file
				charts_paths_list.append(chart_path_name)

	drawingrels_paths_list = []
	for subdir, dirs, files in os.walk(drawingrels_dir):
	    for file in sorted(files):
		ext = os.path.splitext(file)[-1].lower()
		if ext == ".rels":
			if file[0:7] == "drawing":
				drawingrels_path_name = subdir + "/" + file
				drawingrels_paths_list.append(drawingrels_path_name)

	sheetrels_paths_list = []
	for subdir, dirs, files in os.walk(sheetrels_dir):
	    for file in sorted(files):
		ext = os.path.splitext(file)[-1].lower()
		if ext == ".rels":
			if file[0:5] == "sheet":
				sheetrels_path_name = subdir + "/" + file
				sheetrels_paths_list.append(sheetrels_path_name)

	sheets_paths_list = []
	for subdir, dirs, files in os.walk(worksheets_dir):
	    for file in sorted(files):
		ext = os.path.splitext(file)[-1].lower()
		if ext == ".xml":
			if file[0:5] == "sheet":
				sheet_path_name = subdir + "/" + file
				sheets_paths_list.append(sheet_path_name)

	charts_list = []
	sheet_number = str(sheet_number)
	sheet_name = sheetrels_dir + "/sheet" + sheet_number + ".xml.rels"
	if os.path.isfile(sheet_name) == True:
		for i in range(0,len(charts_paths_list)):
			for j in range(0,len(drawingrels_paths_list)):
				doesMatchExist = get_match_number(drawingrels_paths_list[j],ntpath.basename(charts_paths_list[i]),1)
				if doesMatchExist != 0:
					drawingX_dot_xml = drawingrels_paths_list[j][:-4]
					doesMatchExist2 = get_match_number(sheet_name,ntpath.basename(drawingX_dot_xml),1)
					if doesMatchExist2 != 0:
						charts_list.append(ntpath.basename(charts_paths_list[i]))
		
	

	return charts_list







def extract_between(string,start,end):
	expression = start + "(.*?)" + end
	result = re.findall(expression, string)

	return result

def get_element(string,tag):
	expression = beg(tag) + "(.*?)" + end(tag)
	result = re.findall(expression, string)



	return result

def get_attribute_list(string,tag):
	expression = "<" + tag + "(.*?)" + ">"
	result = re.findall(expression, string)
	return result

def get_attribute_value(string,tag,attribute):
	expression0 = "<" + tag + "(.*?)" + ">"
	result0 = re.findall(expression0, string)

	expression = attribute + "=\"" + "(.*?)" + "\""
	result = re.findall(expression, result0[0])

	return result

def beg(string):
	return "<" + string + ">"

def end(string):
	return "</" + string + ">"




def xvalues(fileName,index):
	txt = get_element(file_to_str(fileName),"c:xVal")
	return get_element(txt[index-1],"c:v")

def yvalues(fileName,index):
	txt = get_element(file_to_str(fileName),"c:yVal")
	return get_element(txt[index-1],"c:v")

def xtitle(fileName):
	txt1 = get_element(file_to_str(fileName),"c:valAx")

	txt2 = get_element(txt1[0],"a:t")
	return txt2[0]

def ytitle(fileName):
	txt1 = get_element(file_to_str(fileName),"c:valAx")
	txt2 = get_element(txt1[1],"a:t")
	return txt2[0]

def xfont(fileName):
	txt1 = get_element(file_to_str(fileName),"c:valAx")
	txt2 = get_element(txt1[0],"a:pPr")
	txt3 = get_attribute_value(txt2[0],"a:latin","typeface")
	return txt3

def yfont(fileName):
	txt1 = get_element(file_to_str(fileName),"c:valAx")
	print len(txt1)
	print len(txt1)
	txt2 = get_element(txt1[1],"a:pPr")	
	txt3 = get_attribute_value(txt2[0],"a:latin","typeface")
	return txt3

def xfontsz(fileName):
	txt1 = get_element(file_to_str(fileName),"c:valAx")
	txt2 = get_attribute_value(txt1[0],"a:defRPr","sz")
	return txt2

def yfontsz(fileName):
	txt1 = get_element(file_to_str(fileName),"c:valAx")
	txt2 = get_attribute_value(txt1[1],"a:defRPr","sz")
	return txt2

def series_symbol(fileName,index):
	txt1 = get_element(file_to_str(fileName),"c:ser")
	txt2 = get_element(txt1[index-1],"c:marker")
	txt3 = get_attribute_value(txt2[0],"c:symbol","val")
	return txt3

def number_of_series(fileName):
	txt1 = get_element(file_to_str(fileName),"c:ser")
	return len(txt1)

def accumulate_array(array):
	sum = 0.0
	for i in range(0,len(array)):
		try:
			float_element = float(array[i])
			sum = sum + float_element
		except ValueError:
			print "Warning: one of the elements is not convertable to float!"
		
		
	return sum

def grade_chart(reference_book,book,sheet_number):
	chart_grade = 100.0
	number_of_series_r = 0
	xvalues_r = 0.0
	yvalues_r = 0.0 
	number_of_series_s = 0

	book_folder = unzip(book)
	ref_book_folder = unzip(reference_book)

	student_charts = get_charts_list_in_sheet(book,sheet_number,book_folder)
	referece_charts = get_charts_list_in_sheet(reference_book,sheet_number,ref_book_folder)

	if len(referece_charts) > len(student_charts):
		chart_grade = 0.0
	else: 
		book_charts_dir = book_folder + '/xl/charts/' + student_charts[0]
		ref_charts_dir = ref_book_folder + '/xl/charts/' + referece_charts[0]
		


		number_of_series_r = number_of_series(ref_charts_dir)
		'''xtitle_r = xtitle(ref_charts_dir)
		ytitle_r = xtitle(ref_charts_dir)
		xfont_r = xfont(ref_charts_dir)
		yfont_r = yfont(ref_charts_dir)
		xfontsz_r = xfont(ref_charts_dir)
		yfontsz_r = yfont(ref_charts_dir)'''

		xvalues_r = xvalues(ref_charts_dir,1)
		yvalues_r = yvalues(ref_charts_dir,1)
		#series_symbol_r = series_symbol(ref_charts_dir,1)

		number_of_series_s = number_of_series(book_charts_dir)
		'''xtitle_s = xtitle(book_charts_dir)
		ytitle_s = xtitle(book_charts_dir)
		xfont_s = xfont(book_charts_dir)
		yfont_s = yfont(book_charts_dir)
		xfontsz_s = xfont(book_charts_dir)
		yfontsz_s = yfont(book_charts_dir)'''

		xvalues_s = xvalues(book_charts_dir,1)
		yvalues_s = yvalues(book_charts_dir,1)
		#series_symbol_s = series_symbol(book_charts_dir,1)

		if ( number_of_series_r - number_of_series_s ) == 0:
			'''if xtitle_r.lower() != xtitle_s.lower():
				chart_grade = chart_grade - 0.0
			if ytitle_r.lower() != ytitle_s.lower():
				chart_grade = chart_grade - 0.0	
			if xfont_r != xfont_s:
				chart_grade = chart_grade - 0.0
			if yfont_r != yfont_s:
				chart_grade = chart_grade - 0.0			
			if xfontsz_r != xfontsz_s:
				chart_grade = chart_grade - 0.0
			if yfontsz_r != yfontsz_s:
				chart_grade = chart_grade - 0.0'''

			for iSeries in range(1,number_of_series(ref_charts_dir)+1):
				sum_ref_x = accumulate_array(xvalues(ref_charts_dir,iSeries))
				sum_ref_y = accumulate_array(yvalues(ref_charts_dir,iSeries))
				sum_book_x = accumulate_array(xvalues(book_charts_dir,iSeries))
				sum_book_y = accumulate_array(yvalues(book_charts_dir,iSeries))
				if is_around(sum_ref_x+sum_ref_y,sum_book_x+sum_book_y,1.0) == False:
					chart_grade = chart_grade - 50.0
				#if series_symbol_r != series_symbol_s:
				#	chart_grade = chart_grade - 5.0
		else:
			chart_grade = chart_grade - 100.0

		delete_directory(book_folder)
		delete_directory(ref_book_folder)
	#print "Chart Grade  : ",chart_grade,"/100.0"
	
	return chart_grade


def is_around(target,value,tolrnc_percent):
	if (type(value) == float or type(value) == int or type(value) == long ) and (type(target) == float or type(target) == int or type(target) == long ):
		value = float(value)
		target = float(target)
		if abs((value-target)/(target+0.0000000001)*100.0) < tolrnc_percent:
			return True
		else:
			return False
	elif type(value) == unicode and type(target) == unicode:
		if value == target:
			return True
		else:
			return False

def grade_chart2(chart_address):
	#print file_to_str(chart_address)
	print(type(chart_address))
	print xfont(chart_address)  ##
	return




def file_to_str(fileName):
	#with open(fileName) as myfile:
	#	data = myfile.read().replace('\n', '')
	#	return data
	data = open(fileName, 'r') 
	string = data.read()
	data.close()
	return string

#file = "/home/mostafa/Git/Excellent/chart1.xml"
#print file_to_str(fileName)
#print(type(file))
#grade_chart2(file)
#grade_chart2("chart1.xml")

#file = "/home/mostafa/Git/Excellent/chart-test/Submissions/11/temp/xl/charts"
#file = file + "/chart1.xml"


#print xvalues(file,1)
#print yvalues(file,1)

#print xvalues(file,2)
#print yvalues(file,2)

#print xtitle(file)
#print ytitle(file)

#print xfont(file)
#print yfont(file)

#print xfontsz(file)
#print yfontsz(file)

#print series_symbol(file,1)
#print series_symbol(file,2)