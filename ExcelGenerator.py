# Instruction:

# 根据 Folder 中的sub_folder 和 子文件来建立excel表格
# 并且会根据reference excel来添加对应的col
# -------------------------------------------------
# For Generator Excel:

# column 				Value
# ------		----------------------
# 	0				itemCode/midea_erp_code    		
#	1				image_name
#	2				item_id
#	3				link
#	4				is_video 
#	5				key
#	6				video_flag 
#	7				detail

# For 总表:

# column 				Value
# ------		----------------------
# 	0				item_id    		
#	1				itemCode/midea_erp_code
#	2				link
#	3				image_name
#	4				detail


# Remove 'DS_Store': 
# find . -name '.DS_Store' -type f -delete

import sys
import os
import xlrd
import xlsxwriter

# New Col Access
idArr = ['item_id']
linkArr = ['link']

isVideoArr = ['is_video']
keyArr = ['key']

# Data Retrieve from Reference 
dataArr = [['itemCode/midea_erp_code','image_name']]
product_path = "./冰箱总文件夹"

def writeFile(arr):	# Open and write arr into file
	workbook = xlsxwriter.Workbook(sys.argv[1] + '.xlsx')
	worksheet = workbook.add_worksheet()

	for col, data in enumerate(arr):
		row = 0
		worksheet.write_column(row, col, data)

	workbook.close()

def convert2number(arr, col):
	for i in range(1, len(arr[col])):
		arr[col][i] = int(arr[col][i])


def reshapeArr(arr): # Reshape array
	return list(map(list, zip(*arr)))


def accessDir(path):
	for path, subdirs, files in os.walk(path):
		for name in files:

			value = os.path.join(path, name)

			dataArr.append([path.replace('./冰箱总文件夹/',''),name])
			
			
def filter(arr):
	for i in range(len(arr)-1):
		if arr[i][1] == '.DS_Store':
			arr.remove(arr[i])


def readRef(file): # Retrieve all Data from excel
	loc = (file)
	
	wb = xlrd.open_workbook(loc)
	sheet = wb.sheet_by_index(0)
	sheet.cell_value(0, 0)	# For row 0 and column 0

	row = sheet.nrows 
	col = sheet.ncols

	oldArray = [] # Exist data from excel

	for i in range(col): # Data excel -> array
		pairs = [] # Existed attr in the file
		
		for j in range(row):
			pairs.append(sheet.cell_value(j,i))

		oldArray.append(pairs)

	return oldArray


def compare(data,ref):  # Retrieve all data from spec col
	for i in range(1, len(data[0])):
		j = 1

		while j < len(data[0]):
			if data[0][i] == str(ref[1][j]).split('.')[0]:
				# ref[x][j]: x 会根据具体reference 表格中的col来更改
				idArr.append(int(str(ref[0][j]).split('.')[0]))
				linkArr.append(ref[2][j])

				j = len(data[0]) # Terminate
			j+=1 


def videoChecker(arr, col): # Mark by '1' if has .mp4
	for i in range(1, len(arr[col])):
		if '.mp4' in arr[col][i]:
			isVideoArr.append(1)
		else:
			isVideoArr.append('')


def videoPrefix(arr, imageCol, codeCol): # Replace with itemcode to prefix of .mp4
	for i in range(1, len(arr[imageCol])):
		if '.mp4' in arr[imageCol][i]:
			itemCode = arr[codeCol][i].split('.mp4')[0]
			arr[imageCol][i] = itemCode + '.mp4'


def keyGenerator(arr, imageCol, codeCol): # Generate Key
	for i in range(1, len(arr[0])): # Trace all rows in a col
		keyArr.append(arr[codeCol][i] + '/' + arr[imageCol][i])


def videoFlagGenerator(arr):
	videoFlagArr = ['video_flag']
	index = []
	itemID = []

	for i in range(1, len(arr[4])): # Get index that is_video = 1
		if arr[4][i] == 1:	# Retrieve is_video
			index.append(i)

	for i in index:					
		itemID.append(arr[2][i])	# Retrieve item_id

	for i in range(1, len(arr[2])):	
		if arr[2][i] in itemID:
			videoFlagArr.append(1)
		else:
			videoFlagArr.append('')

	return videoFlagArr


def detailGenerator(arr, file):
	imageName = readRef(file)[3] # image_name from '总表''
	detail = readRef(file)[4]	# detail from '总表''
	itemDict = {}
	detailArr = ['detail']

	for i in range(1, len(imageName)): # Retrieve all image_name without duplicate
		if imageName[i] not in itemDict:
			itemDict[imageName[i]] = detail[i] # Dictionary to store "image_name: detail"

	for i in range(1, len(arr[1])): # Retrieve image_name from 'Product'
		if arr[1][i] in itemDict:
			detailArr.append(itemDict[arr[1][i]])
		else:
			detailArr.append('')

	return detailArr


def featureIdGenerator(arr):
	featureIdArr = ['feature_id']
	feature_id = 1
	pre_index = 1
	
	if arr[7][pre_index] != '':
		featureIdArr.append(feature_id)
		feature_id+=1
	else:
		featureIdArr.append('')

	for i in range(2,len(arr[0])):
		cur_index = i

		if arr[0][cur_index] == arr[0][pre_index]:	# Compare cur and prev elements are same
			pre_index = cur_index 

			if arr[7][cur_index] != '':	
				featureIdArr.append(feature_id)
				feature_id+=1
			else:
				featureIdArr.append('')
		else:
			feature_id = 1
			pre_index = cur_index

			if arr[7][cur_index] != '':	
				featureIdArr.append(feature_id)
				feature_id+=1
			else:
				featureIdArr.append('')

	return featureIdArr


def combine(): # Combine all references to array
	compare(reshapeArr(dataArr) ,readRef('./总表.xlsx'))
	finalArr = reshapeArr(dataArr)

	finalArr.extend((idArr,linkArr)) # item_id & link

	videoChecker(finalArr, 1) 
	finalArr.append(isVideoArr) # is_video

	videoPrefix(finalArr, 1, 0) 

	keyGenerator(finalArr, 1, 0) 
	finalArr.append(keyArr) # key
	finalArr.append(videoFlagGenerator(finalArr)) # video_flag
	finalArr.append(detailGenerator(finalArr, './总表.xlsx')) # detail
	finalArr.append(featureIdGenerator(finalArr))

	# convert2number(finalArr, 0) # Convert item_code to number

	# print(finalArr)
	return finalArr


# main
accessDir(product_path)

filter(dataArr)

writeFile(combine())

