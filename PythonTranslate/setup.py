#!/usr/bin/python 
# coding: UTF-8

import io
import os
import glob
import xlwt
import sys, getopt
import xlrd
import re
import operator
import time
import ConfigParser
import codecs
import struct
import shutil
import xlutils.copy
# import chardet

global lang_package_name
lang_package_name = ""
db_xls_name = 'db'
extract_dir_name = "source"

path = os.getcwd()
global source_path
source_path = path + "/" + extract_dir_name
output_dir_path = path + "/output"
backup_dir_path = path + "/backup"
db_path = path + "/" + db_xls_name + ".xls"
fixdb_path = path + "/fixdb.xls"
translate_path = path + "/translate_path"
translate_path_txt = translate_path + "/translate_path.txt"
translate_path_server_txt = translate_path + "/translate_path_server.txt"
finish_path = translate_path + "/finish"

root_path = os.path.abspath(os.path.join(os.getcwd(), "../..")) + "/"

extract_cn_list = []
extract_cn_dic = {}
once_extract_cn_dic_num = {}
file_list = []
db_dic = {}
exclude_file_dic = {}
cn_pattern = re.compile(u'[\u4e00-\u9fa5]+')
global is_vn
is_vn = False
vn_pattern = re.compile(u'[\u0041-\u1ef9]+')
vn_pattern2 = re.compile(u'[\u0080-\u1ef9]+')	#不包含英文和数字

pattern_A_to_Z = re.compile(u"[A-Z]+")
pattern_0_to_9 = re.compile(u"[0-9]+")
pattern_special = re.compile(u"\.php")
pattern_sign = re.compile(u"[\u0021-\u002f]+|[\u003a-\u0040]+|[\u005b-\u0060]+|[\u007b-\u007e]+|[\u00b7]+|[％★；.–—‘’“”…、。〈〉《》「」『』【】〔〕！（），．：？]+")	#匹配标点符号，不包含空格
# pattern_china_sign = re.compile(u"")
pattern_word_jump_over = re.compile(u"{wordcolor;[\s\S]*?;|{openLink;")

str_list = []
temp_find_pos = 0

#client_or_server：0客户端，1服务端
def CopyNeedTranslateFile(client_or_server):
	if not os.path.exists(finish_path):
		os.makedirs(finish_path)
	shutil.rmtree(finish_path + "/")
	
	global translate_path
	if client_or_server == 0:
		translate_path = translate_path_txt
	elif client_or_server == 1:
		translate_path = translate_path_server_txt
	
	print ("复制文件中...").decode('UTF-8')
	with open(translate_path, 'r') as file_to_read:
		while True:
			line = file_to_read.readline()
			if not line:
				break
			global line_temp
			line = line.strip('\n')
			line = line.strip()
			line_temp = line
			if client_or_server == 0:
				line_temp = re.sub(u'/[^/]*.lua', "", line_temp)
			elif client_or_server == 1:
				line_temp = re.sub(u'/[^/]*.xml', "", line_temp)
			else:
				break
			global finish_path_temp
			finish_path_temp = finish_path + "/" + line_temp
			# print "--------1----------"
			# print line_temp
			# print root_path + line
			# print translate_path
			# print translate_path + "/" + line_temp
			# print finish_path_temp
			# print "--------2----------"
			
			if not os.path.exists(finish_path_temp):
				os.makedirs(finish_path_temp)
			if os.path.exists(root_path + line):
				shutil.copy(root_path + line, finish_path_temp)
			else:
				print ("不存在的文件:").decode('UTF-8') + root_path + line

def ReadConfig():
	config = ConfigParser.ConfigParser()
	config.readfp(codecs.open("config.ini", "r", "UTF-8"))
	exclude_file_str = config.get("Default", "ExcludeFileNames")
	exclude_file_list = exclude_file_str.split(';')
	for item in exclude_file_list:
		exclude_file_dic[item] = 1

	global lang_package_name
	lang_package_name = config.get("Default", "LangPackageName")

	print(config.get("Default", "SourcePath"))
	global source_path
	if config.get("Default", "SourcePath") != "":
		source_path = config.get("Default", "SourcePath")

	global excel_path
	if config.get("Default", "ExcelPath") != "":
		excel_path = config.get("Default", "ExcelPath")

def ReadExcel():
	#创建样式----------------------------
	#原中文
	pattern = xlwt.Pattern() # Create the Pattern
	pattern.pattern = xlwt.Pattern.SOLID_PATTERN # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
	pattern.pattern_fore_colour = 5 # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
	style = xlwt.XFStyle() # Create the Pattern
	style.pattern = pattern # Add Pattern to Style
	
	#未翻译的中文
	pattern_red = xlwt.Pattern() # Create the Pattern
	pattern_red.pattern = xlwt.Pattern.SOLID_PATTERN # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
	pattern_red.pattern_fore_colour = 2 #说好的2是红色呢？结果还是黄色...
	style_red = xlwt.XFStyle() # Create the Pattern
	style_red.pattern = pattern # Add Pattern to Style
	#-----------------------------------------

	print ("开始读取Excel并翻译").decode('UTF-8')
	print excel_path
	for file_path in glob.glob(excel_path + os.sep + "*"):
		if os.path.isdir(file_path):
			print ("文件夹").decode('UTF-8') + file_path
		else:
			if exclude_file_dic.has_key(os.path.basename(file_path)) or (os.path.splitext(file_path)[1] != ".xls" and os.path.splitext(file_path)[1] != ".xlsx"):
				continue
			print ("文件名：").decode('UTF-8') + file_path
			rb = xlrd.open_workbook(file_path, formatting_info=True) #formatting_info是否完全保留excel内容读取
			wb = xlutils.copy.copy(rb)
			ext_name = os.path.splitext(file_path)[1]
			if ext_name == ".xls" or ext_name == ".xlsx":
				one_sheet_table = {}
				once_extract_cn_dic_num.clear()
				sheet_index = 1
				#查找每个标签下的中文
				sheets = rb.sheet_names()
				for i in range(len(sheets)):
					if i != 0 and not one_sheet_table.has_key(i):
						continue
				
					sheet = rb.sheet_by_name(sheets[i])			#同时获取sheet名称、行数、列数（sheet.name,sheet.nrows,sheet.ncols）
					
					#需要特殊处理的配置
					[dirname,filename]=os.path.split(file_path)
					if filename == ("G-怪物技能.xls").decode('UTF-8'):
						one_sheet_table[0] = 11
						ws = wb.get_sheet(i)
						for col_num in range(sheet.ncols):
							if is_over_generate or "" == sheet.cell(one_sheet_table[0] - 3, col_num).value or "" == sheet.cell(one_sheet_table[0] - 1, col_num).value or "TEMP" == sheet.cell(one_sheet_table[0] - 3, col_num).value:
								is_over_generate = True
								if "TEMP" == sheet.cell(one_sheet_table[0] - 1, col_num).value:
									for row_num in range(len(sheet.col_values(col_num))):
										ws.write(row_num, col_num, sheet.cell(row_num, col_num).value, style)
								continue
							for j in range (len(sheet.col_values(col_num))):
								if j < one_sheet_table[0]:
									continue
								if None == sheet.cell(j, 0).value or "" == sheet.cell(j, 0).value:
									break
								#当前列有一个单元格存在中文就翻译
								if IsTranslateByContent(sheet.cell(j, col_num).value):
									WriteExcel(ws, col_num, sheet, one_sheet_table[0], style, style_red, file_path)
									break
						break
					
					#遍历每一行、每一列，第一个标签获取参数，第二个标签开始判断、翻译
					if i == 0:
						for j in range (len(sheet.row_values(1))):
							if j == 2:
								one_sheet_table[0] = int(sheet.cell(1, j).value) + 2	#需要翻译的起始行数，如果设置的起始行数是1，+2实际是从第4行开始遍历
								
							elif j >= 3 and "" != sheet.cell(1, j).value:
								one_sheet_table[sheet_index] = sheet_index
								sheet_index = sheet_index + 1
					else:
						#获取sheet对象，通过sheet_by_index()获取的sheet对象没有write()方法
						ws = wb.get_sheet(i)
						is_over_generate = False
						for col_num in range(sheet.ncols):
							if is_over_generate or "" == sheet.cell(one_sheet_table[0] - 3, col_num).value or "" == sheet.cell(one_sheet_table[0] - 1, col_num).value or "TEMP" == sheet.cell(one_sheet_table[0] - 3, col_num).value:
								is_over_generate = True
								if col_num < sheet.ncols and "TEMP" == sheet.cell(one_sheet_table[0] - 3, col_num).value:
									for row_num in range(len(sheet.col_values(col_num))):
										ws.write(row_num, col_num, sheet.cell(row_num, col_num).value, style)
								continue
							for j in range (len(sheet.col_values(col_num))):
								if j < one_sheet_table[0]:
									continue
								if None == sheet.cell(j, 0).value or "" == sheet.cell(j, 0).value:
									break
								#当前列有一个单元格存在中文就翻译
								if IsTranslateByContent(sheet.cell(j, col_num).value):
									WriteExcel(ws, col_num, sheet, one_sheet_table[0], style, style_red, file_path)
									wb.save(file_path)
									rb = xlrd.open_workbook(file_path)
									sheet = rb.sheet_by_name(sheets[i])
									break

				wb.save(file_path)
				print "该表未翻译的数量：".decode('UTF-8'), len(once_extract_cn_dic_num)
				print("--------------")
				
	print "读取Excel并翻译完毕".decode('UTF-8')
	ExtractExcelChinese()

	
#翻译，并将需要翻译的那一列复制插入到最后一列；如果当前单元格已翻译则不移动
#col_target：需要翻译的那一列的列数
#sheet：当前标签
#blank_rows：需要留空的行数
def WriteExcel(ws, col_target, sheet, blank_rows, style, style_red, file_path):
	insert_col_num, is_once_insert = GetInsertColNum(col_target, sheet, blank_rows)
	#print insert_col_num, is_once_insert
	for j in range (len(sheet.col_values(col_target))):
		target_text_temp = sheet.cell(j, col_target).value
		if j < blank_rows:
			if insert_col_num < 256:
				if j >= blank_rows - 2:
					ws.write(j, insert_col_num, target_text_temp, style)
				elif j >= blank_rows - 3:
					ws.write(j, insert_col_num, "TEMP", style)
			continue
		if None == sheet.cell(j, 0) or "" == sheet.cell(j, 0).value:
			break
		
		translate_target_text = TranslateExcelContent(target_text_temp)
		if "" == target_text_temp or IsTranslateByContent(target_text_temp):
			if insert_col_num < 256:
				ws.write(j, insert_col_num, sheet.cell(j, col_target).value, style)
			if "" != sheet.cell(j, col_target).value and IsTranslateByContent(translate_target_text):
				once_extract_cn_dic_num[sheet.cell(j, col_target).value] = 1
				#if not extract_cn_dic.has_key(sheet.cell(j, col_target).value):
					#print sheet.cell(j, col_target).value.encode("GB18030");
				extract_cn_dic[sheet.cell(j, col_target).value] = 1
		elif is_once_insert and insert_col_num < 256:
			ws.write(j, insert_col_num, sheet.cell(j, insert_col_num).value, style)
			
		if IsTranslateByContent(translate_target_text):
			ws.write(j, col_target, translate_target_text, style_red)
		else:
			ws.write(j, col_target, translate_target_text)
				
#
def GetInsertColNum(col_target, sheet, blank_rows):
	for col_num in range(sheet.ncols):
		if col_num > col_target:
			if sheet.cell(blank_rows - 1, col_target).value == sheet.cell(blank_rows - 1, col_num).value:
				return col_num, True
			elif IsInsertForCurrCol(col_num, sheet, blank_rows):
				break

	if col_num + 1 >= sheet.ncols and "TEMP" != sheet.cell(blank_rows - 3, col_num).value:
		return col_num + 2, False
	elif "" != sheet.cell(blank_rows - 3, col_num).value and "TEMP" == sheet.cell(blank_rows - 3, col_num).value:
		return col_num + 1, False
	return col_num, False
	
#当前列是否能够插入数据（当前列全是空单元格，且前一列起始行数为空单元格 或 前一列起始行数是TEMP）
def IsInsertForCurrCol(col_num, sheet, blank_rows):
	for row_num in range (len(sheet.col_values(col_num))):
		if row_num > blank_rows and "" == sheet.cell(row_num, 0).value:
			break
		if ("" != sheet.cell(blank_rows - 3, col_num - 1).value and "TEMP" != sheet.cell(blank_rows - 3, col_num - 1).value) or "" != sheet.cell(row_num, col_num).value:
			return False
	return True


def TranslateExcelContent(content):
	if(db_dic.has_key(content)):
		en = db_dic[content]
		if en != "" and en != None:
			return en
	return content
				
def ExtractExcelChinese():
	print "开始导出Excel中未翻译的内容~~~".decode('UTF-8')
	wb = xlwt.Workbook()
	ws = wb.add_sheet('cn')
	row_num = 0
	for i in extract_cn_dic:
		 ws.write(row_num, 0, i)
		 ws.col(0).width = 8000
		 row_num += 1
	wb.save(output_dir_path +"/%s%s%d%s" % (lang_package_name, "_excel", time.time(), ".xls"))
	print "complete export xls, xls name is " + lang_package_name + "_excel, item total count is %d" % (row_num)

			
def WriteDb():
	print "start write db"
	# row_num = 0
	# wb = xlwt.Workbook()
	# ws = wb.add_sheet('cn')
	# ws.col(0).width = 8000
	# ws.col(1).width = 8000
	# for i in db_dic:
	#     if i != "" and db_dic[i] != "":
	#         ws.write(row_num, 0, i)
	#         ws.write(row_num, 1, db_dic[i])
	#         row_num +=1

	# backup_path = backup_dir_path + "/%s%s%d%s" % (db_xls_name, "_", time.time(), ".xls")
	# wb.save(backup_path)
	# print "complete backup, path is " + backup_path
	# wb.save(db_path)
	# print "complete write db"

def DepthExtract(path):
	for file_path in glob.glob(path + os.sep + '*' ):
		if os.path.isdir(file_path):
			DepthExtract(file_path) 
		else:
			if exclude_file_dic.has_key(os.path.basename(file_path)):
				continue
			print "try extract" + file_path
			ext_name = os.path.splitext(file_path)[1]
			if ext_name == ".lua":
				temp_file = open(file_path, 'r')
				content = temp_file.read()
				content = content.decode('UTF-8')
				file_list.append(file_path)
				ExtractFromLua(file_path, content)
			elif ext_name == ".xml":
				temp_file = open(file_path, 'r')
				content = temp_file.read()
				content = content.decode('UTF-8')
				file_list.append(file_path)
				ExtractFromXml(file_path, content)
			elif ext_name == ".php":
				temp_file = open(file_path, 'r')
				content = temp_file.read()
				content = content.decode('UTF-8')
				file_list.append(file_path)
				ExtractFromPHP(file_path, content)
			elif ext_name == ".txt":
				temp_file = open(file_path, 'r')
				content = temp_file.read()
				content = content.decode('UTF-8')
				file_list.append(file_path)
				ExtractFromTxt(file_path, content)

def ExtractFromTxt(file_path, content):
	#查找中文
	pattern = re.compile(u'\"([^\"]*)\"')
	results = pattern.findall(content)
	for result in results:
		if len(cn_pattern.findall(result)) > 0 and not extract_cn_dic.has_key(result):
			#print "extract " + result.encode('GBK', 'ignore')  #Win7中的cmd，默认codepage是CP936，即GBK的编码，所以需要先将的Unicode先编码为GBK
			extract_cn_dic[result] = 1
			extract_cn_list.append([source_path, result, "", len(result)])

	pattern = re.compile(u'\'([^\']*)\'')
	results = pattern.findall(content)
	for result in results:
		if len(cn_pattern.findall(result)) > 0 and not extract_cn_dic.has_key(result):
			#print "extract" + result.encode('GBK', 'ignore')
			extract_cn_dic[result] = 1
			extract_cn_list.append([source_path, result, "", len(result)])

	#查找越南文
	if is_vn:
		# pattern = re.compile(u'\"([^\"]*)\"')
		# results = pattern.findall(content)
		# for result in results:
			if not extract_cn_dic.has_key(content):
				extract_cn_dic[content] = 1
				extract_cn_list.append([source_path, content, "", len(content)])

		# pattern = re.compile(u'\'([^\']*)\'')
		# results = pattern.findall(content)
		# for result in results:
			if not extract_cn_dic.has_key(content):
				extract_cn_dic[content] = 1
				extract_cn_list.append([source_path, content, "", len(content)])

def ExtractFromLua(file_path, content):
	content = re.sub(u'--\"([^\"]*)\"', "", content)
	content = re.sub(u'--\s+\"([^\"]*)\"', "", content)
	content = re.sub(u'print\([\s\S]*?\)', "", content)
	content = re.sub(u'print_warning\([\s\S]*?\)', "", content)
	content = re.sub(u'print_error\([\s\S]*?\)', "", content)

	#check " legal
	match = re.search('\"', content)
	pattern = re.compile(u'\"')
	results = pattern.findall(content)
	if len(results) % 2 == 1 :
		print "-----Error:check legal 1, file path is " + file_path
		return
	
	#check ' legal
	pattern = re.compile(u'\'')
	results = pattern.findall(content)
	if len(results) % 2 == 1 :
		print "Error:check legal 2, file path is " + file_path
		return
	
	#查找中文
	pattern = re.compile(u'\"([^\"]*)\"')
	results = pattern.findall(content)
	for result in results:
		if len(cn_pattern.findall(result)) > 0 and not extract_cn_dic.has_key(result):
			#print "extract " + result.encode('GBK', 'ignore')  #Win7中的cmd，默认codepage是CP936，即GBK的编码，所以需要先将的Unicode先编码为GBK
			extract_cn_dic[result] = 1
			extract_cn_list.append([source_path, result, "", len(result)])

	pattern = re.compile(u'\'([^\']*)\'')
	results = pattern.findall(content)
	for result in results:
		if len(cn_pattern.findall(result)) > 0 and not extract_cn_dic.has_key(result):
			#print "extract" + result.encode('GBK', 'ignore')
			extract_cn_dic[result] = 1
			extract_cn_list.append([source_path, result, "", len(result)])

	#查找越南文
	if is_vn:
		pattern = re.compile(u'\"([^\"]*)\"')
		results = pattern.findall(content)
		for result in results:
			if len(vn_pattern.findall(result)) > 0 and not extract_cn_dic.has_key(result):
				extract_cn_dic[result] = 1
				extract_cn_list.append([source_path, result, "", len(result)])

		pattern = re.compile(u'\'([^\']*)\'')
		results = pattern.findall(content)
		for result in results:
			if len(vn_pattern.findall(result)) > 0 and not extract_cn_dic.has_key(result):
				extract_cn_dic[result] = 1
				extract_cn_list.append([source_path, result, "", len(result)])

def ExtractFromXml(file_path, content):
	content = re.sub(u'<!--[\s\S]*?-->', "", content)
  
	pattern = re.compile(u'\<\!\[CDATA\[[\s\S]*?\]\]\>')
	results = pattern.findall(content)
	for result in results:
		if len(cn_pattern.findall(result)) > 0:
			result = result.replace("<![CDATA[", "")
			result = result.replace("]]>", "")
			if not extract_cn_dic.has_key(result):
				#print "extract" + result.encode('GBK', 'ignore')
				extract_cn_dic[result] = 1
				extract_cn_list.append([source_path, result, "", len(result)])
	pattern = re.compile(u'\<.*?=.*?\>')
	results = pattern.findall(content)
	for result in results:
		cn_results = cn_pattern.findall(result)
		for cn_result in cn_results:
			if not extract_cn_dic.has_key(cn_result):
				#print "extract" + cn_result.encode('GBK', 'ignore')
				extract_cn_dic[cn_result] = 1
				extract_cn_list.append([source_path, cn_result, "", len(cn_result)])
	pattern = re.compile(u'\<.*\>.*\<\/.*\>')
	results = pattern.findall(content)
	for result in results:
		cn_result = cn_pattern.findall(result)
		if len(cn_result) > 0:
			if not extract_cn_dic.has_key(result):
				#print "extract" + result.encode('GBK', 'ignore')
				extract_cn_dic[result] = 1
				extract_cn_list.append([source_path, result, "", len(result)])

	if is_vn:
		for result in results:
			if len(vn_pattern.findall(result)) > 0:
				result = result.replace("<![CDATA[", "")
				result = result.replace("]]>", "")
				if not extract_cn_dic.has_key(result):
					extract_cn_dic[result] = 1
					extract_cn_list.append([source_path, result, "", len(result)])
		pattern = re.compile(u'\<.*?=.*?\>')
		results = pattern.findall(content)
		for result in results:
			cn_results = vn_pattern.findall(result)
			for cn_result in cn_results:
				if not extract_cn_dic.has_key(cn_result):
					extract_cn_dic[cn_result] = 1
					extract_cn_list.append([source_path, cn_result, "", len(cn_result)])
		pattern = re.compile(u'\<.*\>.*\<\/.*\>')
		results = pattern.findall(content)
		for result in results:
			cn_result = vn_pattern.findall(result)
			if len(cn_result) > 0:
				if not extract_cn_dic.has_key(result):
					extract_cn_dic[result] = 1
					extract_cn_list.append([source_path, result, "", len(result)])

def ExtractFromPHP(file_path, content):
	content = re.sub(u'[^\"]*--[^\"]*', "", content)
	#check " legal
	match = re.search('\"', content)
	pattern = re.compile(u'\"')
	results = pattern.findall(content)
	if len(results) % 2 == 1 :
		print "Error:check legal 1, file path is " + file_path
		return
	
	#check ' legal
	pattern = re.compile(u'\'')
	results = pattern.findall(content)
	if len(results) % 2 == 1 :
		print "Error:check legal 2, file path is " + file_path
		return
	
	#search chinese
	pattern = re.compile(u'\"([^\"]*)\"')
	results = pattern.findall(content)
	for result in results:
		if len(cn_pattern.findall(result)) > 0 and not extract_cn_dic.has_key(result):
			#print "extract " + result.encode('GBK', 'ignore')  #Win7中的cmd，默认codepage是CP936，即GBK的编码，所以需要先将的Unicode先编码为GBK
			extract_cn_dic[result] = 1
			extract_cn_list.append([source_path, result, "", len(result)])

	pattern = re.compile(u'\'([^\']*)\'')
	results = pattern.findall(content)
	for result in results:
		if len(cn_pattern.findall(result)) > 0 and not extract_cn_dic.has_key(result):
			#print "extract" + result.encode('GBK', 'ignore')
			extract_cn_dic[result] = 1
			extract_cn_list.append([source_path, result, "", len(result)])
	
	pattern = re.compile(u'\<.*\>.*\<\/.*\>')
	results = pattern.findall(content)
	for result in results:
		cn_result = cn_pattern.findall(result)
		if len(cn_result) > 0:
			if not extract_cn_dic.has_key(result):
				#print "extract" + result.encode('GBK', 'ignore')
				extract_cn_dic[result] = 1
				extract_cn_list.append([source_path, result, "", len(result)])

#extract chinese
def ExtractChinese():
	print "start extract chinese, source path is " + source_path
	del extract_cn_list[:]
	del file_list[:]
	extract_cn_dic.clear()
	db_dic.clear() 

	ReadConfig()
	DepthExtract(source_path) 
	print "extract completed, file count is %d" % (len(file_list))

	print "start export to xls"
	wb = xlwt.Workbook()
	ws = wb.add_sheet('cn')
	row_num = 0
	for item in extract_cn_list:
		# if not db_dic.has_key(item[1]):
			ws.write(row_num, 0, item[1])
			ws.write(row_num, 1, item[2])
			ws.col(0).width = 8000
			ws.col(1).width = 8000
			row_num +=1

	wb.save(output_dir_path +"/%s%s%d%s" % (lang_package_name, "_", time.time(), ".xls"))
	print "complete export xls, xls name is " + lang_package_name + ", item count is %d" % (row_num)

##translate file
def TranslateFile(path):
	print "translate file " + path 
	a_file = open(path, 'r')
	content = "content"
	content = a_file.read()
	content = content.decode('UTF-8')

	ext_name = os.path.splitext(path)[1]
	new_content = content
	if ext_name == ".lua":
		new_content = TranslateLuaContent(content) 
	elif ext_name == ".xml":
		new_content = TranslateXmlContent(content)
	elif ext_name == ".txt":
		new_content = TranslateTxtContent(content)

	if content != new_content:
		e_file = open(path, 'w')
		e_file.write(new_content.encode('UTF-8'))
		e_file.close()

def TranslateLuaContent(content):
	#new_content = content
	for item in extract_cn_list:
		cn = item[1]
		if(db_dic.has_key(cn)):
			en = db_dic[cn]
			if en != "" and en != None:
				content = content.replace('"' + cn + '"', '"' + en + '"')
				content = content.replace("'" + cn + "'", "'" + en + "'")
	return content

def TranslateXmlContent(content):
	for item in extract_cn_list:
		cn = item[1]
		if(db_dic.has_key(cn)):
			en = db_dic[cn]
			if en != "" and en != None:
				content = content.replace('"' + cn + '"', '"' + en + '"')
				content = content.replace("'" + cn + "'", "'" + en + "'")
				content = content.replace(cn, en)
	return content
	
def TranslateTxtContent(content):
	for item in extract_cn_list:
		cn = item[1]
		if(db_dic.has_key(cn)):
			en = db_dic[cn]
			if en != "" and en != None:
				content = content.replace(cn, en)
	return content

def IsTranslateByContent(content):
	if content == None or content == "":
		return False

	new_content = content
	if type(content) != unicode:
		new_content = repr(content)
	
	#查找中文
	return len(cn_pattern.findall(new_content)) > 0

def Translate():
	print "start extract chinese, source path is " + source_path
	del extract_cn_list[:]
	del file_list[:]
	extract_cn_dic.clear()
	db_dic.clear() 

	ReadConfig()
	DepthExtract(source_path)
	print "extract completed, file count is %d" % (len(file_list))

	lang_pack_path = path + "/" + lang_package_name
	print "start import lang package, path is " + lang_pack_path + ".xls"
	wb = xlrd.open_workbook(lang_pack_path + ".xls")
	sheet = wb.sheet_by_name('cn')
	for rownum in range(sheet.nrows):
		cn = sheet.cell(rownum, 0).value
		en = sheet.cell(rownum, 1).value
		if(type(en) == float):   
		   en = repr(int(en)).rstrip("\r\n")
		elif(type(en) == int):
			en = repr(int(en)).rstrip("\r\n")
		else:
			en = en.rstrip("\r\n")
		#print chardet.detect(cn)
		db_dic[cn] = en
		# for item in extract_cn_list:
		#     if item[1] == cn:
		#         #在有些语言翻译中，如‘零’会直接只翻译成0,读取后python认为是float型，用replace时会引起报错
		#         if(type(en) == float):   
		#             item[2] = repr(int(en)).rstrip("\r\n")
		#         else:
		#             item[2] = en.rstrip("\r\n")

	extract_cn_list.sort(lambda x,y:cmp(y[3],x[3]))

	for item in file_list:
		TranslateFile(item)

	#WriteDb()
	
	print "translate complete!"

def SetDbDic():
	print "start extract chinese, source path is " + source_path
	del extract_cn_list[:]
	del file_list[:]
	extract_cn_dic.clear()
	db_dic.clear() 

	ReadConfig()

	print "extract completed, file count is %d" % (len(file_list))

	lang_pack_path = path + "/" + lang_package_name
	print "start import lang package, path is " + lang_pack_path + ".xls"
	wb = xlrd.open_workbook(lang_pack_path + ".xls")
	sheet = wb.sheet_by_name('cn')
	for rownum in range(sheet.nrows):
		cn = sheet.cell(rownum, 0).value
		en = sheet.cell(rownum, 1).value
		if(type(en) == float):   
		   en = repr(int(en)).rstrip("\r\n")
		elif(type(en) == int):
			en = repr(int(en)).rstrip("\r\n")
		else:
			en = en.rstrip("\r\n")
		db_dic[cn] = en
	extract_cn_list.sort(lambda x,y:cmp(y[3],x[3]))

def Clear():
	remain_count = 5
	print "start remove file ,but remain count %d" % remain_count
	file_list = []
	#remove oldder backup ,but remain some
	del file_list[:]
	for file_path in glob.glob(output_dir_path + os.sep + '*' ):
		file_list.append([file_path, os.path.getctime(file_path)])

	file_list.sort(lambda x,y:cmp(x[1],y[1])) 
	del_num = len(file_list) - remain_count
	for item in file_list:
		if del_num <= 0:
			break
		del_num -= 1
		os.remove(item[0])
		print "remove file " + item[0]
	
	#remove oldder output ,but remain some
	del file_list[:]
	for file_path in glob.glob(backup_dir_path + os.sep + '*' ):
		file_list.append([file_path, os.path.getctime(file_path)])

	file_list.sort(lambda x,y:cmp(x[1],y[1])) 
	del_num = len(file_list) -  remain_count
	for item in file_list:
		if del_num <= 0:
			break
		del_num -= 1
		os.remove(item[0])
		print "remove file " + item[0]
	
	print "clear complete!"

def TranslateToTw():
	print "start extract chinese, source path is " + source_path
	del extract_cn_list[:]
	del file_list[:]
	extract_cn_dic.clear()
	db_dic.clear() 

	ReadConfig()
	DepthExtract(source_path) 
	print "extract completed, file count is %d" % (len(file_list))

	zh_file = open(path + "/ZhConVersion.php", 'r')
	content = zh_file.read()
	content = content.decode('UTF-8')
	pattern = re.compile(u'\'.*\' => \'.*\',')
	results = pattern.findall(content)
	t_pattern = re.compile(u'\'.*?\'')
	zh_list = []
	for item in results:
		t_items = t_pattern.findall(item) 
		t_item = []
		t_item.append(t_items[0].replace("'", ""))
		t_item.append(t_items[1].replace("'", ""))
		t_item.append(len(t_item[0]))
		zh_list.append(t_item)

	zh_list.sort(lambda x,y:cmp(y[2],x[2])) 
	
	for cn_item in extract_cn_list:
		cn = cn_item[1]
		tw = cn
		for item in zh_list:
			tw = tw.replace(item[0], item[1])
		db_dic[cn] = tw
		#print "cn to tw:%s%s" % (cn_item[1], cn_item[2])

	for item in file_list:
		TranslateFile(item)
	
	print "complete translate CN to TW"

def FixFiles():
	print "start fix files"
	ReadConfig()
	wb = xlrd.open_workbook(fixdb_path)
	sheet = wb.sheet_by_name('cn')
	#check
	for rownum in range(sheet.nrows):
		file_name = sheet.cell(rownum, 0).value 
		cn = sheet.cell(rownum, 1).value
		en = sheet.cell(rownum, 2).value
		path = source_path + "/" + file_name
		if not os.path.exists(path):
			print("fix error: not find " + path)	
			print("Error! please remove it from fixdb file")
			return

		a_file = open(path, 'r')
		content = "content"
		content = a_file.read()
		content = content.decode('UTF-8')
		new_content = content.replace(cn, en)
		if new_content == content:
			print("fix warning, not find match :" + en.encode('GBK', 'ignore'))

	for rownum in range(sheet.nrows):
		file_name = sheet.cell(rownum, 0).value 
		cn = sheet.cell(rownum, 1).value
		en = sheet.cell(rownum, 2).value
		path = source_path + "/" + file_name
		if os.path.exists(path):
			print("fix files:" + path)
			a_file = open(path, 'r')
			content = "content"
			content = a_file.read()
			content = content.decode('UTF-8')
			new_content = content.replace(cn, en)
			e_file = open(path, 'w')
			e_file.write(new_content.encode('UTF-8'))
			e_file.close()
	
	print "fix files finish"

def PackUi(path):
	for file_path in glob.glob(path + os.sep + '*' ):
		if os.path.isdir(file_path):
			PackUi(file_path) 
		else:
			ext_name = os.path.splitext(file_path)[1]
			if ext_name == ".rb":
				print "try PackUi" + file_path
				os.chdir(os.path.dirname(file_path))
				packer_name = os.path.basename(file_path)
				print packer_name
				os.system("ruby %s" % packer_name)

def AdditionalNumSpace():
	print "start additional num space, source path is " + source_path
	del extract_cn_list[:]
	del file_list[:]
	extract_cn_dic.clear()
	db_dic.clear() 

	ReadConfig()
	DepthExtractAndAdditionalNumSpace(source_path)
	print "extract completed, file count is %d" % (len(file_list))

	for item in file_list:
		TranslateNumSpaceToFile(item)

	#WriteDb()
	
	print "translate complete!"

def DepthExtractAndAdditionalNumSpace(path):
	for file_path in glob.glob(path + os.sep + '*' ):
		if os.path.isdir(file_path):
			DepthExtractAndAdditionalNumSpace(file_path) 
		else:
			if exclude_file_dic.has_key(os.path.basename(file_path)):
				continue
			print "try extract" + file_path
			ext_name = os.path.splitext(file_path)[1]
			if ext_name == ".lua":
				temp_file = open(file_path, 'r')
				content = temp_file.read()
				content = content.decode('UTF-8')
				file_list.append(file_path)
				ExtractNumSpaceFromLua(file_path, content)
			elif ext_name == ".xml":
				temp_file = open(file_path, 'r')
				content = temp_file.read()
				content = content.decode('UTF-8')
				file_list.append(file_path)
				ExtractNumSpaceFromXml(file_path, content)
			elif ext_name == ".prefab":
				temp_file = open(file_path, 'r')
				content = temp_file.read()
				content = content.decode('utf8')
				# content = content.encode('unicode-escape')
				# content = content.decode('GBK')
				# print(content).decode('GBK', 'ignore')
				file_list.append(file_path)
				ExtractNumSpaceFromPrefab(file_path, content)

def ExtractNumSpaceFromLua(file_path, content):
	content = re.sub(u'--\"([^\"]*)\"', "", content)
	content = re.sub(u'--\s+\"([^\"]*)\"', "", content)
	content = re.sub(u'print\([\s\S]*?\)', "", content)
	content = re.sub(u'print_log\([\s\S]*?\)', "", content)
	content = re.sub(u'print_warning\([\s\S]*?\)', "", content)
	content = re.sub(u'print_error\([\s\S]*?\)', "", content)
	content = re.sub(u'TestPrint\([\s\S]*?\)', "", content)
	content = re.sub(u'bundle_name = \"[\s\S]*?\"', "", content)
	content = re.sub(u'asset_name = \"[\s\S]*?\"', "", content)
	content = re.sub(u'mask = \"[\s\S]*?\"', "", content)
	content = re.sub(u'string.format\([\s\S]*?\)', "", content)

	#check " legal
	match = re.search('\"', content)
	pattern = re.compile(u'\"')
	results = pattern.findall(content)
	if len(results) % 2 == 1 :
		print "Error:check legal 1, file path is " + file_path
		return
	
	#check ' legal
	pattern = re.compile(u'\'')
	results = pattern.findall(content)
	if len(results) % 2 == 1 :
		print "Error:check legal 2, file path is " + file_path
		return

	#查找越南文
	pattern = re.compile(u'\"([^\"]*)\"')
	results = pattern.findall(content)
	for result in results:
		if len(vn_pattern2.findall(result)) > 0:
			extract_cn_list.append([source_path, result, "", len(result), os.path.splitext(file_path)[0]])

	pattern = re.compile(u'\'([^\']*)\'')
	results = pattern.findall(content)
	for result in results:
		if len(vn_pattern2.findall(result)) > 0:
			extract_cn_list.append([source_path, result, "", len(result), os.path.splitext(file_path)[0]])
			
def ExtractNumSpaceFromPrefab(file_path, content):
	#查找越南文
	pattern = re.compile(u'm_Text: \"([^\"]*)\"')
	results = pattern.findall(content)
	for result in results:
		if len(vn_pattern2.findall(result)) > 0 or len(pattern_A_to_Z.findall(result)) > 0:
			extract_cn_list.append([source_path, result, "", len(result), os.path.splitext(file_path)[0]])

	pattern = re.compile(u'm_Text: \'([^\']*)\'')
	results = pattern.findall(content)
	for result in results:
		if len(vn_pattern2.findall(result)) > 0 or len(pattern_A_to_Z.findall(result)) > 0:
			extract_cn_list.append([source_path, result, "", len(result), os.path.splitext(file_path)[0]])

def ExtractNumSpaceFromXml(file_path, content):
	print ""

def TranslateNumSpaceToFile(path):
	print "translate file " + path 
	a_file = open(path, 'r')
	content = "content"
	content = a_file.read()
	content = content.decode('UTF-8')

	file_name = os.path.splitext(path)[0]
	ext_name = os.path.splitext(path)[1]
	new_content = content
	if ext_name == ".lua":
		new_content = TranslateNumSpaceLuaContent(content, file_name) 
	elif ext_name == ".prefab":
		new_content = TranslateNumSpacePrefabContent(content, file_name)

	if content != new_content:
		e_file = open(path, 'w')
		e_file.write(new_content.encode('UTF-8'))
		e_file.close()

def TranslateNumSpaceLuaContent(content, file_name):
	#new_content = content
	for item in extract_cn_list:
		if item[4] != file_name:
			continue
		cn = item[1]
		if len(pattern_special.findall(cn)) > 0:
			continue
		global temp_find_pos
		temp_find_pos = 0
		global str_list
		str_list = list(cn)
		is_force_continue = False
		jump_by_index_continue = 0
		for index in range(len(cn)):
			if is_force_continue != False:
				if cn[index] == is_force_continue:
					is_force_continue = False
				continue
			if jump_by_index_continue != 0:
				if index >= jump_by_index_continue:
					jump_by_index_continue = 0
				continue
			is_force_continue, jump_by_index_continue = GetIsContinue(cn, index)
			# if cn[index] == "[":
			# 	print cn[index:index+3].encode('GBK', 'ignore')
			# print("?????", index, cn[index], cn[index:index+6])
			
		en = ''.join(str_list)
		content = content.replace('"' + cn + '"', '"' + en + '"')
		content = content.replace("'" + cn + "'", "'" + en + "'")
	return content

def TranslateNumSpacePrefabContent(content, file_name):
	#new_content = content
	for item in extract_cn_list:
		if item[4] != file_name:
			continue
		cn = item[1]
		if len(pattern_special.findall(cn)) > 0:
			continue
		global temp_find_pos
		temp_find_pos = 0
		global str_list
		str_list = list(cn)
		is_force_continue = False
		jump_by_index_continue = 0
		for index in range(len(cn)):
			if is_force_continue != False:
				if cn[index] == is_force_continue:
					is_force_continue = False
				continue
			if jump_by_index_continue != 0:
				if index >= jump_by_index_continue:
					jump_by_index_continue = 0
				continue
			is_force_continue, jump_by_index_continue = GetIsContinue(cn, index)
			# if cn[index] == "[":
			# 	print cn[index:index+3].encode('GBK', 'ignore')
			# print("?????", index, cn[index], is_force_continue, jump_by_index_continue)
			
		en = ''.join(str_list)
		content = content.replace('"' + cn + '"', '"' + en + '"')
		content = content.replace("'" + cn + "'", "'" + en + "'")
	return content

def TranslateNumSpaceXmlContent(content):
	for item in extract_cn_list:
		cn = item[1]
		content = content.replace('"' + cn + '"', '"' + en + '"')
		content = content.replace("'" + cn + "'", "'" + en + "'")
		content = content.replace(cn, en)
	return content

def GetIsContinue(cn, index):
	global str_list
	global temp_find_pos
	if cn[index] == "<" and index + 6 <= len(cn) and cn[index:index+6] == "<color":
		is_force_continue = ">"
		if index > 0 and cn[index - 1] != " " and len(pattern_sign.findall(cn[index-1])) == 0:
			str_list.insert(index + temp_find_pos, " ")
			temp_find_pos += 1
		return is_force_continue, 0
	if cn[index] == "<" and index + 8 <= len(cn) and cn[index:index+8] == "</color>":
		jump_by_index_continue = index + 7
		if index + 8 < len(cn) and cn[index + 8] != " " and len(pattern_sign.findall(cn[index+8])) == 0:
			str_list.insert(index + 8 + temp_find_pos, " ")
			temp_find_pos += 1
		return False, jump_by_index_continue
	if index + 10 <= len(cn) and cn[index:index+10] == "font color":
		is_force_continue = ">"
		if index > 0 and cn[index - 1] != " " and len(pattern_sign.findall(cn[index-1])) == 0:
			str_list.insert(index + temp_find_pos, " ")
			temp_find_pos += 1
		return is_force_continue, 0
	if cn[index] == "<" and index + 7 <= len(cn) and cn[index:index+7] == "</font>":
		jump_by_index_continue = index + 6
		if index + 7 < len(cn) and cn[index + 7] != " " and len(pattern_sign.findall(cn[index+7])) == 0:
			str_list.insert(index + 7 + temp_find_pos, " ")
			temp_find_pos += 1
		return False, jump_by_index_continue
	if cn[index] == "[" and index + 3 <= len(cn) and cn[index:index+3] == "[p_":
		is_force_continue = "]"
		return is_force_continue, 0
	if cn[index:index+2].lower() == "\\x":
		jump_by_index_continue = index + 3
		return False, jump_by_index_continue
	if cn[index:index+2].lower() == "\\u":
		jump_by_index_continue = index + 5
		return False, jump_by_index_continue
	if cn[index] == "{" and index + 18 <= len(cn) and len(pattern_word_jump_over.findall(cn[index:index+18])) != 0:
		jump_by_index_continue = index + 17
		return False, jump_by_index_continue
	if cn[index] == "{" and index + 10 <= len(cn) and len(pattern_word_jump_over.findall(cn[index:index+10])) != 0:
		jump_by_index_continue = index + 9
		return False, jump_by_index_continue
	if (index > 0) and (cn[index - 1].lower() == "x" or cn[index].lower() == "x"):
		return False, 0
	if cn[index] == "%":
		if index + 4 <= len(cn) and (cn[index+1:index+4] == "02d"):
			jump_by_index_continue = index + 3
			return False, jump_by_index_continue
		if index + 5 <= len(cn) and (cn[index+1:index+5] == "0.1f" or cn[index+1:index+5] == "0.2f"):
			jump_by_index_continue = index + 4
			return False, jump_by_index_continue
		if index + 2 <= len(cn) and (cn[index+1:index+3]) == "2d":
			jump_by_index_continue = index + 1
			return False, jump_by_index_continue
		return False, 0

	uppercase_letter_len = len(pattern_A_to_Z.findall(cn[index]))
	on_uppercase_letter_len = index > 0 and len(pattern_A_to_Z.findall(cn[index - 1])) or 0

	#当前字符的前4个字符或后4个字符是boss
	if index - 4 >= 0 and cn[index-4:index].lower() == "boss" and len(pattern_sign.findall(cn[index])) == 0:
		if cn[index] != " ":
			str_list.insert(index + temp_find_pos, " ")
			temp_find_pos += 1
		return False, 0
	if index + 5 <= len(cn) and cn[index+1:index+5].lower() == "boss" and len(pattern_sign.findall(cn[index])) == 0:
		if cn[index] != " ":
			str_list.insert(index + 1 + temp_find_pos, " ")
			temp_find_pos += 1
		jump_by_index_continue = index + 4
		return False, jump_by_index_continue
	if index + 4 <= len(cn) and cn[index+1:index+4].lower() == "pvp" and len(pattern_sign.findall(cn[index])) == 0:
		if cn[index] != " ":
			str_list.insert(index + 1 + temp_find_pos, " ")
			temp_find_pos += 1
		jump_by_index_continue = index + 3
		return False, jump_by_index_continue
	# print("-------", index, cn[index], cn[index+1:index+4], pattern_sign.findall(cn[index]))
	if index + 4 <= len(cn) and cn[index+1:index+4].lower() == "1v1" and len(pattern_sign.findall(cn[index])) == 0:
		if cn[index] != " ":
			str_list.insert(index + 1 + temp_find_pos, " ")
			temp_find_pos += 1
		jump_by_index_continue = index + 3
		return False, jump_by_index_continue
	if index + 4 <= len(cn) and cn[index+1:index+4].lower() == "3v3" and len(pattern_sign.findall(cn[index])) == 0:
		if cn[index] != " ":
			str_list.insert(index + 1 + temp_find_pos, " ")
			temp_find_pos += 1
		jump_by_index_continue = index + 3
		return False, jump_by_index_continue
	if index - 2 >= 0 and cn[index-2:index].lower() == "lv":
		return False, 0
	# print("------,", index, cn[index], cn[index+4], cn[index+1:index+4])
	if index + 4 <= len(cn) and cn[index+1:index+4].lower() == "vip" and len(pattern_sign.findall(cn[index])) == 0:
		if cn[index] != " ":
			str_list.insert(index + 1 + temp_find_pos, " ")
			temp_find_pos += 1
		jump_by_index_continue = index + 3
		return False, jump_by_index_continue
	if index - 3 >= 0 and cn[index-3:index].lower() == "vip" and len(pattern_sign.findall(cn[index])) == 0:
		if cn[index] != " ":
			str_list.insert(index + temp_find_pos, " ")
			temp_find_pos += 1
		return False, 0
	#当前字符的前2个字符是\n或\t
	if (index - 2 >= 0) and (cn[index-2:index] == "\\n" or cn[index-2:index] == "\\t"):
		return False, 0

	#当前字符是大写字母，上一个字符不是标点符号和大写字母
	if uppercase_letter_len > 0 and index > 0 and len(pattern_sign.findall(cn[index - 1])) == 0 and on_uppercase_letter_len == 0 and cn[index - 1] != " ":
		# print("----- 111 ----- %s", cn[index], cn[index - 1], pattern_china_sign.findall(cn[index-1]))
		str_list.insert(index + temp_find_pos, " ")
		temp_find_pos += 1
		#下一个字符不是空格、大写字母和数字
		# if index + 1 < len(cn) and cn[index + 1] != " " and len(pattern_A_to_Z.findall(cn[index + 1])) == 0 and len(pattern_0_to_9.findall(cn[index + 1])) == 0:
		# 	print("----- %222 ----- %s", cn[index], cn[index + 1])
		# 	str_list.insert(index + 1 + temp_find_pos, " ")
		# 	temp_find_pos += 1
		return False, 0

	results_len = len(pattern_0_to_9.findall(cn[index]))
	on_results_len = index > 0 and len(pattern_0_to_9.findall(cn[index - 1])) or 0
	on_sign_len = index > 0 and len(pattern_sign.findall(cn[index - 1])) or 0
	next_results_len = 0
	next_uppercase_letter_len = 0
	next_sign_len = 0
	if index + 1 < len(cn):
		next_results_len = len(pattern_0_to_9.findall(cn[index + 1]))
		next_uppercase_letter_len = len(pattern_A_to_Z.findall(cn[index + 1]))
		next_sign_len = len(pattern_sign.findall(cn[index + 1]))
	#当前字符是数字
	if results_len > 0 and index > 0:
		#上一个字符不是符号（包括空格）、数字、x
		if cn[index - 1] != " " and on_sign_len == 0 and on_results_len == 0 and cn[index - 1] != "x":
			str_list.insert(index + temp_find_pos, " ")
			temp_find_pos += 1
		#下一个字符是越南文
		if index + 1 < len(cn) and cn[index + 1] != " " and cn[index + 1].lower() != "h" and cn[index + 1].lower() != "s" and next_sign_len == 0 and next_results_len == 0 and next_uppercase_letter_len == 0:
		# if (index + 1 < len(cn)) and len(vn_pattern2.findall(cn[index + 1])) != 0:
			str_list.insert(index + 1 + temp_find_pos, " ")
			temp_find_pos += 1
	return False, 0

opts, args = getopt.getopt(sys.argv[1:], "etwayk")
if len(opts) == 0:
	print "Error:must have agument"
	print ("-e 导出中文").decode('UTF-8')
	print ("-t 翻译").decode('UTF-8')
	print ("-w 台湾翻译").decode('UTF-8')
	print ("-a 直接对excel表翻译").decode('UTF-8')
	print ("-y 翻译或替换越南文").decode('UTF-8')
	print ("-k 处理越南文粘在一起的问题").decode('UTF-8')

for op, value in opts:
	if op == "-e":
		ExtractChinese()
	elif op == "-t":
		#CopyNeedTranslateFile(0)
		Translate()
		#FixFiles()
	elif op == "-s":
		#CopyNeedTranslateFile(1)
		Translate()
	elif op == "-w":
		TranslateToTw()
	elif op == "-v":
		is_vn = True
		Translate()
	elif op == "-a":
		SetDbDic()
		ReadExcel()
	elif op == "-y":
		is_vn = True
		Translate()
	elif op == "-k":
		is_vn = True
		AdditionalNumSpace()