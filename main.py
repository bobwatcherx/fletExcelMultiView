from  flet import *
import openpyxl
import os


def main(page:Page):
	listfile = Row(wrap=True)
	# CREATE TABLE
	dt = DataTable(
		columns=[],
		rows=[]
		)

	# NOW I WILL SCAN FOLDER MYEXCEL AND FIND ALL FILE
	# AND THEN SHOW IT IN CONTAINER

	def scan_now(folder_path):
		# NOW IF FOLDER NOT FOUND THEN PRINT NOTFOUND
		if not os.path.exists(folder_path):
			print("folder myexcel not found")
			return
		# CHECK NAME is myexcel but not folder
		if not os.path.isdir(folder_path):
			print("folder myexcel not FOlder")
			return
		files = os.listdir(folder_path)
		# IF NOT FOUND ALL FILE IN FOLDER MYEXCEL
		if not files:
			print("folder myexcel empty")
			return
		print("myexcel folder found here guys :")

		for file_name in files:
			print(file_name)
			listfile.controls.append(
				# CREATE DRAG ABLE
				Draggable(
				content=Container(
					bgcolor="blue200",
					height=150,
					padding=10,
					content=Column([
					Icon(name="folder",size=30),
					Text(file_name,size=30,weight="bold")
						])
					)


					)
				)
		page.update()
	folder_path = "myexcel"
	scan_now(folder_path)






	def youaccept(e):
		src = page.get_control(e.src_id)
		# CLEAR TABLE COLUMN AND ROW
		dt.columns.clear()
		dt.rows.clear()
		page.update()

		folder_path = "myexcel"
		# GET FILE NAME IF FOUND 
		file_name = src.content.content.controls[1].value
		file_path = os.path.join(folder_path,file_name)

		# OPEN EXCEL
		workbook = openpyxl.load_workbook(file_path)
		sheet = workbook.active

		# NOW GET COLUMN NAMES
		# YOU EXCEL FILE
		column_names = [cell.value for cell in sheet[1]]
		data = []
		for row in sheet.iter_rows(min_row=2):
			row_data = {}
			for index,cell in enumerate(row):
				column_name = column_names[index]
				row_data[column_name] = cell.value
			data.append(row_data)
		print("Data : ")
		print(data)
		column_set = set()
		for item in data:
			column_set.update(item.keys())
		# AND GET NAME COLUMN IN YOU EXCEL FILE
		for column_name in column_set:
			print(column_name)
			dt.columns.append(DataColumn(Text(column_name)))
		page.update()

		# AND NOW SET FOR ROW YOU EXCEL FILE
		for item in data:
			row_cells = []
			for column_name in column_names:
				cell_value = item.get(column_name,"")
				row_cells.append(DataCell(Text(str(cell_value))))
			dt.rows.append(DataRow(cells=row_cells))
		page.update()





	page.add(
		Column([
		Text("Excel Viewer",size=30),
		listfile,
		# CREATE DRAG FOR UPLOAD FILE 
		# AREA
		DragTarget(
		on_accept=youaccept,
		content=Container(
			bgcolor="yellow",
			padding=10,
			content=Text("upload Excel YOu Here",
				size=30,weight="bold"
				)
			)
			),
		dt



			])
		)

flet.app(target=main)
