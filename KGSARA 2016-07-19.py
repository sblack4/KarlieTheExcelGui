#!/usr/bin/python
# -*- coding: iso-8859-1 -*-
# http://sebsauvage.net/python/gui/ 

import tkinter as Tkinter
import xlsxwriter as xlw
import re 
import pandas as pd

class simpleapp_tk(Tkinter.Tk):
	def __init__(self,parent):
		Tkinter.Tk.__init__(self,parent)
		self.parent = parent
		self.minsize(width=600, height=300)
		self.initialize()

	def initialize(self):
		self.grid()

		self.entryVariable = Tkinter.StringVar()
		self.entry = Tkinter.Entry(self,textvariable=self.entryVariable)
		self.entry.grid(column=0,row=0,sticky='EW')
		self.entry.bind("<Return>", self.OnPressEnter)
		self.entryVariable.set(u"2016 Master GSA Tracking.xlsx")

		button = Tkinter.Button(self,text=u"Click me !",
								command=self.OnButtonClick)
		button.grid(column=1,row=0)

		self.labelVariable = Tkinter.StringVar()
		label = Tkinter.Label(self,textvariable=self.labelVariable,
							  anchor="w",fg="white",bg="blue")
		label.grid(column=0,row=1,columnspan=2,sticky='EW')
		self.labelVariable.set(u"Hello !")

		self.grid_columnconfigure(0,weight=1)
		self.resizable(True,True)
		self.update()
		self.geometry(self.geometry())		 
		self.entry.focus_set()
		self.entry.selection_range(0, Tkinter.END)

	def OnButtonClick(self):
		self.labelVariable.set(self.entryVariable.get() + " (You clicked the button)")
		self.entry.focus_set()
		self.entry.selection_range(0, Tkinter.END)
		self.getNotebook(str(self.entryVariable.get()))

	def OnPressEnter(self,event):
		self.labelVariable.set(self.entryVariable.get() + " (You pressed ENTER)")
		self.entry.focus_set()
		self.entry.selection_range(0, Tkinter.END)
		self.getNotebook(str(self.entryVariable.get()))
	
	def getNotebook(self, notebook):
		# print('\n FIRST!  Make sure the excel file is in the same folder as this file, the output will be placed here too \n \n Second, enter notebook name as in Karlies great workbook.xlsx') # 'GSA_TrackMaster.xlsm'
		#notebook = input('Enter Notebook Name: ') 

		# ABOVE IS FROM CMD LINE APP - this def gets files and run the rest of the program

		self.labelVariable.set('Working on :'+self.entryVariable.get())
		self.entry.focus_set()

		months = ['January IFF',	'February IFF',	'March IFF',	'April IFF',	'May IFF',	'June IFF',	'July IFF',	'August IFF',	'September IFF',	'October IFF',	'November IFF',	'December IFF']

		noteboook = self.entryVariable.get()

		reg = re.compile('.*[\.]')
		new_notebook = re.search(reg, notebook)
		try: 
			xls = pd.ExcelFile(notebook)
		except Exception as e:
			self.labelVariable.set(str(e)+"\n Oops! make sure it's spelled right and in the same folder")
			self.entry.focus_set()

		try:		# to get names 
			
			new_notebook = str(new_notebook.group(0)) 
			new_notebook = str(new_notebook.replace('.','')) + ' FeesReport.xlsx'
			wb = xlw.Workbook(new_notebook)
			wb.close()
			for month in months:
				wb.add_worksheet(month)
			wb.close()
			xls2 = pd.ExcelWriter(new_notebook)

		except Exception as e: 
			self.labelVariable.set(str(e)+' Oops! make sure to enter it like: "2016 Master GSA Tracking.xlsx"')
			self.entry.focus_set()
		try: 
			self.contractDF(xls, xls2, months)			# Run the rest of the program
		except Exception as e:
			self.labelVariable.set(str(e)+' Oops! make sure to enter it like: "2016 Master GSA Tracking.xlsx"')
			self.entry.focus_set()
			
	
	def stringToCurrency(self, df, MonthCol):
		try:
			#df[MonthCol].convert_objects(convert_numeric=True).dropna()
		
			df[MonthCol] = df[MonthCol].replace('X', 0).astype(float).fillna(0.0)
			#df = df.loc[(not isinstance(df[MonthCol], str))]
			df = df.loc[(df[MonthCol] > 0)]
			bool = True 
		except Exception as e: 
			self.labelVariable.set(str(e)+' Make sure the charecters in the IFF fields are only numeric or "X"')
			self.entry.focus_set()

		
			bool = False 
			df = pd.DataFrame()
		return df, bool 


	def insertColumns(self, df, month, contract):
		df.insert(0,'Order#',['' for i in range(0, len(df.index))])
		df.insert(2, 'Time Frame', [month for i in range(0,len(df.index))])
		df.insert(4, 'Contract', [contract for i in range(0,len(df.index))])
		df.insert(5,'PO #',['' for i in range(0, len(df.index))])
		colOrder = ['Order#', 'Teaming Partner', 'Time Frame',month, 'Contract', 'PO #', 'Product']
		df = df[colOrder]
		return df
		
		
	def contractsDict(self, xls):
		contracts = dict()
		for sheet in xls.sheet_names:
			contracts[sheet] = dict()
			try: 
				df = xls.parse(sheet, header=-1)
				custs = df[df[df.columns[0]].str.contains('Cust') == True].index.tolist()
				#print(df.columns)
				#print(str(sheet)+ " \n "+ str(custs) +" \n ")

				for i in range(0,len(custs)):
					if i < 1:
						#print(df.columns[1])
						contracts[sheet].update({custs[i]:df.columns[1]})
					else:
						#print(df[df.columns[1]][custs[i]-2])
						contracts[sheet].update({custs[i]:df[df.columns[1]][custs[i] - 2]})
			except Exception as e: 
				self.labelVariable.set(str(e)+' Error in finding headers (rows where Column A is Cust. No.)')
				self.entry.focus_set()
		return(contracts)

		
	def contractDF(self, xls, xls2, months):
		monthsD = dict(zip(months, [0 for i in range(0,len(months))]))
		sheets = xls.sheet_names
		contDict = self.contractsDict(xls)
		for sheet in sheets: 
			#print(sheet)
			for monthIFF in months:
				#print(monthIFF)
				contracts = contDict[sheet]
				contracts_id = list(contDict[sheet].keys())
				
				contracts_id.sort()
				
				for i in range(0, len(contracts_id)): 
					if i < len(contracts_id) - 1:
						end = contracts_id[i + 1] - (contracts_id[i] + 5)
					else:
						end = 1000
					
					dfm = xls.parse(sheet, header=contracts_id[i] + 1)
					dfm = dfm.iloc[(dfm.index < end)] 			# df = df.loc[(df[MonthCol] > 0)&(df.index<end)]
					#print(dfm, contracts[contracts_id[i]],contracts_id[i],end, contracts_id)
					df, bool = self.stringToCurrency(dfm, monthIFF)
					if any(monthIFF in s for s in dfm.columns) and bool:
				
						df = df.loc[:, ['Teaming Partner', 'Product', monthIFF]]
						df = self.insertColumns(df, monthIFF, contracts[contracts_id[i]])
						df.to_excel(xls2, sheet_name=monthIFF, startrow=monthsD[monthIFF], header=(monthsD[monthIFF] == 0))
						monthsD[monthIFF] += len(df.index) + 1
					else:
						pass
		self.labelVariable.set('Done.')
		self.entry.focus_set()
		# TODO: add module to get rid of null rows
		# TODO: add module to format sheet - autofit columnwidth, highlighting
		# TODO: add module for NO FEE and ADOBE handling

if __name__ == "__main__":
	app = simpleapp_tk(None)
	app.title('Enter Excel File Name (Make sure it is in the same folder!)')
	app.mainloop()
