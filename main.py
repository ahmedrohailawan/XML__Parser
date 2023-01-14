# XML parser libraray and Regex library
# although you have to (pip install openpyxl) to read or write in excel files

import xml.etree.ElementTree as ET	
import re							
import openpyxl						
from openpyxl.styles import Font	

# XML parser function that will accept xml file, then translate data and convert it into excel file
def XMLParser(file_path):
	# condition to check if file exist or not
	# if path is correct then whether it is xml file or not
	try:
		global tree
		tree = ET.parse(file_path)
	except:
		print("1 Either file is not xml file or")
		print("2 No such file or directory exist!")
		print("3 Enter path again... ")
		main()
	root = tree.getroot()
	wb = openpyxl.Workbook()
	sheet = wb.worksheets[0]
	excel_filename = 'compiler.xlsx'

	# initializing excel file, then fetching node tage of xml file which will be columns in excel file
	st_font = Font(bold=True)
	node_tags = []
	col = 'A'
	for node in root.iter():
		if node.tag in node_tags:
			break
		elif node.tag != 'catalog':
			sheet[col+'1'].font = st_font
			sheet[col+'1'] = node.tag.capitalize()
			node_tags.append(node.tag)
			col = chr(ord(col)+1)

	# Regex Patterns for each data item
	catalog_pattern = '^catalog$'
	book_pattern = '^book$'
	author_pattern = '^author$'
	title_pattern = '^title$'
	genre_pattern = '^genre$'
	price_pattern = '^price$'
	publish_date_pattern = '^publish_date$'
	description_pattern = '^description$'

	# Parses the XML file iteratively and uses regex patterns to match and maps the data onto the Excel file
	current_row = 1	
	for node in root.iter():
		if re.findall(catalog_pattern, node.tag):
			sheet.title = 'Catalog'
		elif re.findall(book_pattern, node.tag):
			current_row += 1
			sheet['A'+str(current_row)] = node.get('id')
		elif re.findall(author_pattern, node.tag):
			sheet['B'+str(current_row)] = node.text
		elif re.findall(title_pattern, node.tag):
			sheet['C'+str(current_row)] = node.text
		elif re.findall(genre_pattern, node.tag):
			sheet['D'+str(current_row)] = node.text
		elif re.findall(price_pattern, node.tag):
			sheet['E'+str(current_row)] = node.text
		elif re.findall(publish_date_pattern, node.tag):
			sheet['F'+str(current_row)] = node.text
		elif re.findall(description_pattern, node.tag):
			sheet['G'+str(current_row)] = node.text
		else:
			print("Did not matched to anything!")
	
	# saving excel file if operation run successfully
	try:
		wb.save(excel_filename)
		print("Converted to Excel File Successfully")
	except:
		print("Some Error Occured, Try again !")
		

def main():
	print("\n*************************************")
	file_path = input("Enter file path : ")
	XMLParser(file_path)
	print("*************************************\n")

	# if you dont want to take user input comment above and uncomment below code
	# print("\n*************************************")
	# file_path = 'compiler.xml'
	# XMLParser(file_path)
	# print("\n*************************************")

if __name__ == "__main__":
	main()