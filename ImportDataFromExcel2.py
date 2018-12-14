# -*- coding: UTF-8 -*-

import wx
import os
import wx.xrc as xrc
import wx.grid as  gridlib
#from win32com.client import Dispatch
import win32com.client
import ResMgr
import customDialog
from EasyExcel import EasyExcel
import time

#特殊数据表导入数据起始行、字段对应行、名称行
ENTITY_EXCEL_DATA_LINE = {
	"Sound": (1, 0, 0), 
}

from UI.ImportDataFromExcel_Dialog import ID_IMPORTDATAFROMEXCEL


class ImportDataFromExcel(ID_IMPORTDATAFROMEXCEL):
	def __init__(self, parent,entName):
		ID_IMPORTDATAFROMEXCEL.__init__(self, parent)
		#定义模块属性
		font = wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.NORMAL, False, "") #字体
		#self._xlApp   = win32com.client.Dispatch('Excel.Application')  #连接COM
		self._xlApp = None
		#self._xlApp.Visible  = False
		#self._xlApp.DisplayAlerts  = False
		self._xlBook  = None	 #打开的BOOK
		self._easyexcel = None
		self._entName = entName  #数据定义名称
		self._entDefn = None		 #数据定义结构
		self._title   = []
		self._RangeData = None   #从EXCEL读取的数据
		self._beginLine = 2	  #开始的行数
		self._sheet = {} 	#记录Excel表分页对应序号
		self._scriptName = None #记录函数脚本名
		self.ID_TEXTCTRL2.SetValue(str(self._beginLine))
		
		#修改特殊Excel表的起始行、字段对应行、名称行
		if entName in ENTITY_EXCEL_DATA_LINE:
			entLineData = ENTITY_EXCEL_DATA_LINE[entName]
			self.ID_TEXTCTRL2.SetValue(str(entLineData[0]))
			self.ID_TEXTCTRL5.SetValue(str(entLineData[1]))
			self.ID_TEXTCTRL6.SetValue(str(entLineData[2]))
			self._beginLine = entLineData[0]
		
		self._selModName = ""	#选择的脚本名称
		#属性选择控件
		box = wx.BoxSizer(wx.VERTICAL)
		self.ID_PANEL1.SetSizer(box)
		self.idgrid = gridlib.Grid(self.ID_PANEL1,-1,(0,0),(-1,-1))
		box.Add(self.idgrid,1, wx.GROW|wx.ALL, 1)
		self.idgrid.SetRowLabelSize(20)   
		self.idgrid.SetColLabelSize(17) 
		self.idgrid.SetSize(self.ID_PANEL1.GetSize())
		
		#得到属性列表
		self._entDefn = ResMgr.getEntityDefine(self._entName)
		self.idgrid.CreateGrid(len(self._entDefn), 3) 
		self.idgrid.SetColLabelValue(0,u"属性")
		self.idgrid.SetColLabelValue(1,u"类型")
		self.idgrid.SetColLabelValue(2,u"对应列")
		self.idgrid.SetLabelFont(font)

		#写入数据
		#i = 0
		#for k, v in self._entDefn.iteritems():
		#	self.idgrid.SetCellValue(i,0,k)
		#	self.idgrid.SetCellValue(i,1,v[0])
		#	i = i + 1
		entData = ResMgr.openEntity(self._entName)
		keys = entData.keys()
		for i in range(len(keys)):
			type = self._entDefn[keys[i]][0]
			self.idgrid.SetCellValue(i,0,keys[i])
			self.idgrid.SetCellValue(i,1,type)

		self.ID_CHECKBOX2.SetValue(1)

		#设置第一和二行为只读
		attred = wx.grid.GridCellAttr()
		attred.SetReadOnly(1)
		attred.SetBackgroundColour(wx.Colour(red=240, green=240, blue=240))
		self.idgrid.SetColAttr(0,attred)
		attred = wx.grid.GridCellAttr()
		attred.SetReadOnly(1)
		attred.SetBackgroundColour(wx.Colour(red=240, green=240, blue=240))
		self.idgrid.SetColAttr(1,attred)
		#设置第三行为下拉控件
		attr = wx.grid.GridCellAttr()
		attr.SetEditor(gridlib.GridCellChoiceEditor(self._title, False))
		self.idgrid.SetColAttr(2,attr)
		#绑定事件
		self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)					   #关闭事件
		self.Bind(wx.EVT_BUTTON, self.OnOpenExcelFile, self.ID_BUTTON1)   #打开EXCEL文件
		self.Bind(wx.EVT_BUTTON, self.OnSelectScriptFile, self.ID_BUTTON2)#选择脚本
		self.Bind(wx.EVT_BUTTON, self.OnOk, self.ID_BUTTON3)			  #确定
		self.Bind(wx.EVT_BUTTON, self.OnExit, self.ID_BUTTON4)			#退出
		self.Bind(wx.EVT_CHOICE, self.OnChoiceSheet, self.ID_CHOICE1)	 #选择Sheet事件	 
		#self.ID_TEXTCTRL1.Bind(wx.EVT_KILL_FOCUS, self.OnKillFocus_TEXTCTRL1)
		self.ID_TEXTCTRL2.Bind(wx.EVT_KILL_FOCUS, self.OnKillFocus_TEXTCTRL2)
		#self.idgrid.Bind(gridlib.EVT_GRID_SELECT_CELL, self.OnSelectCell) #单元格改变
		self.Bind(wx.EVT_CHECKBOX, self.EvtCheckBox, self.ID_CHECKBOX1)   #脚本选择
		self.ID_TEXTCTRL5.Bind(wx.EVT_KILL_FOCUS, self.OnKillFocus_TEXTCTRL5) #如果有选择对应行
		self.ID_TEXTCTRL6.Bind(wx.EVT_KILL_FOCUS, self.OnKillFocus_TEXTCTRL6) #名称行选择
		self.Bind(wx.EVT_BUTTON, self.startBat, self.ID_BUTTON)
		
		section = ResMgr.openSection("../paths.xml")  #获取路径
		if not section[ "tableFilePaths" ]:
			section.createSection( "tableFilePaths")
			section.createSection( "tableFilePaths/abPath").asString = ""
			section.createSection( "tableFilePaths/serverPath").asString = ""
			section.saveToXML("./paths.xml")
		path = section["tableFilePaths"]["abPath"].asString
		if not path or not os.path.isdir( path ):
			dlg = wx.MessageDialog(self, u"请正确配置paths.xml文件的tableFilePaths路径", u"提示",wx.OK)
			dlg.ShowModal()
			dlg.Destroy()
			return
		exportPath = ""
		sheet = ""
		sql = "SELECT sm_key, sm_importPath, sm_sheet FROM tbl_AssetPath"
		msgData = ResMgr.getDataRecord( sql )
		for data in msgData:
			if data[ 0 ].strip() == entName:
				exportPath = data[1].strip()
				sheet = data[ 2 ].strip().decode("utf-8")
				break
		if exportPath and path:
			path = path + "\\" + exportPath
		else:
			path = ""
		path = path.replace( "/", "\\" )
		if path and not os.path.exists( path.decode("utf-8") ):
			dlg = wx.MessageDialog(self, u"指定文件不存在：" + path.decode("utf-8") , u"提示",wx.OK)
			dlg.ShowModal()
			dlg.Destroy()
			return
		self.ID_TEXTCTRL1.SetValue( path.decode( "utf-8" ) ) 
		if os.path.exists( path.decode("utf-8") ):
			self.openExcelFile()
			if self._sheet.has_key( sheet ):
				index = self._sheet[ sheet ]
				self.ID_CHOICE1.Select( index )
				self.OnChoiceSheet( None )
				self.setCheckBoxByScriptName()
			else :
				dlg = wx.MessageDialog(self, u"指定文件分页不存在：" + sheet , u"提示",wx.OK)
				dlg.ShowModal()
				dlg.Destroy()
				return
		
	def setCheckBoxByScriptName(self):
		self._scriptName = self._entName + "Handle"
		funlist = self.findScriptFunction('ImportExcelScript.py')
		templist = self.setLower(funlist)
		if self._scriptName.lower() in templist:
			#print "slef._scriptName=", self._scriptName
			self.ID_CHECKBOX1.SetValue(True)
			self.setScriptStatue(funlist)
			#wx.PostEvent(self.ID_CHECKBOX1, wx.CommandEvent(wx.wxEVT_COMMAND_CHECKBOX_CLICKED))
			
	#将字符串全部转化为小写,容错处理
	def setLower(self,funlist):
		templist = []
		for i in range(len(funlist)):
			templist.append(funlist[i].lower())
		return templist
	
	def setScriptStatue(self,funlist):
		self.ID_TEXTCTRL3.SetValue('ImportExcelScript.py')
		self.ID_COMBOCTRL1.Clear()
		templist = self.setLower(funlist)
		for i in range(len(funlist)):
			self.ID_COMBOCTRL1.Append(funlist[i])
			#此处比较是和转化后的小写比较,最终选择的还是转化前的脚本函数
			if templist[i]==self._scriptName.lower():
				self.ID_COMBOCTRL1.SetSelection(i)
			#print "funlist[i]=",funlist[i]

	#脚本选择事件
	def EvtCheckBox(self, event):
		print "EvtCheckBox"
		if self.ID_CHECKBOX1.IsChecked():
			funlist = self.findScriptFunction('ImportExcelScript.py')
			self.setScriptStatue(funlist)
		else:
			self.ID_TEXTCTRL3.SetValue('')
			self.ID_COMBOCTRL1.Clear()
			
	#打开EXCEL文件
	def OnSelectScriptFile(self, event):
		fileName = ""
		filePath = ""
		dlg = customDialog.FileDialog(
				self, message=u"选择一个脚本文件",
				defaultDir = "./res/entities/plugins",
				defaultFile=self._entName + ".py",
				wildcard="xls File (*.py)|*.py",
				style=wx.OPEN | wx.MULTIPLE
				)
		if dlg.ShowModal() == wx.ID_OK:
			filePath = dlg.GetPath()
			fileName = dlg.GetFilename()
		dlg.Destroy()

		#写到函数列表
		if fileName != "":
			fileName = fileName.encode('utf-8')
			funlist  = self.findScriptFunction(fileName)
			self.ID_COMBOCTRL1.Clear()
			for i in range(len(funlist)):
				self.ID_COMBOCTRL1.Append(funlist[i])
			self.ID_TEXTCTRL3.SetValue(fileName)

	#找到脚本中的函数
	def findScriptFunction(self,nameScript):
		name = nameScript.split(".")[0]
		funlist = []
		self._selModName = name
		m = __import__(name)
		reload(m)
		molist = dir(m)
		molist.sort()
		#找到模块中的函数 
		for i in range(len(molist)):
			modnum  = getattr(m,molist[i]) 
			if str(type(modnum)) == "<type 'function'>":
				funlist.append(molist[i])
		return funlist

	# 获取所有字段名称
	def getColumnTitle( self ):
		titleRow = int(self.ID_TEXTCTRL5.GetValue())
		return list(self._RangeData[titleRow])

	def getExportSheetNames(self):
		"""
		获取需要导入数据的表名
		"""
		exportnames = []
		if self._easyexcel:
			sheetnames = self._easyexcel.getAllSheetNames()
			for i in sheetnames:
				if "export" in i.lower():
					exportnames.append(i)
		return exportnames
					

	#单元格改变
	def OnSelectCell(self, evt = None):
		pass
	
	#EXCEL文件地址改变
	def OnKillFocus_TEXTCTRL1(self,event):
		self.openExcelFile()

	#开始行数改变
	def OnKillFocus_TEXTCTRL2(self,event):
		self._beginLine = int(self.ID_TEXTCTRL2.GetValue())

	#对应行
	def OnKillFocus_TEXTCTRL5(self,event):
		value = self.ID_TEXTCTRL5.GetValue()
		if value != "" and self._RangeData != None:
			acc	 = int(value)					  #得到行号
			acclist = list(self._RangeData[acc])	  #得到数据
			for i in range(len(acclist)):
				title = self.idgrid.GetCellValue(i,0) #得到标题
				title = title.encode('utf-8')
				try:
					num   = acclist.index(title) + 1  #找在列表的号
					numABC= self.Numb2ABC(num)		#转成ABC的表示
					self.idgrid.SetCellValue(i,2,self._title[num])	#写入值
				except ValueError :
					title = ""

	#名称行
	def OnKillFocus_TEXTCTRL6(self,event):
		#得到数据定义的excel名称
		titNmae = self.ID_TEXTCTRL6.GetValue()
		if titNmae != "" and self._RangeData != None :
			tieNumb = int(titNmae)					#得到行号
			TitList = list(self._RangeData[tieNumb])  #得到数据

			#如果不是空就重新写单元格的下拉单元
			self._title = []
			if self._RangeData != None:
				tData = self._RangeData[0]
				for i in range(len(tData) + 1):
					value = self.Numb2ABC(i) 
					if titNmae != "" and i != 0:
						value = value + ":" + str(TitList[i - 1])
					self._title.append(value)

			attr = wx.grid.GridCellAttr()
			attr.SetEditor(gridlib.GridCellChoiceEditor(self._title, False))
			self.idgrid.SetColAttr(2,attr)
			self.OnKillFocus_TEXTCTRL5(None)

	#打开EXCEL文件对话框
	def OnOpenExcelFile(self, event):
		name = self._entName
		path = os.getcwd()
		dlg = customDialog.FileDialog(
				self, message=u"选择一个EXCEL文件",
				defaultDir=path, 
				defaultFile=name,
				wildcard="EXCEL File (*.xls,*xlsx,*xlsm)|*.xls;*.xlsx;*.xlsm",
				style=wx.OPEN | wx.MULTIPLE
				)
		if dlg.ShowModal() == wx.ID_OK:
			filePaths = dlg.GetPaths()
			fileName  = filePaths[0]
			self.ID_TEXTCTRL1.SetValue(fileName)
			path = dlg.GetPath()
		dlg.Destroy()
		self.openExcelFile()

	#打开EXCEL文件
	def openExcelFile(self):
		#读文件名称
		fileName = self.ID_TEXTCTRL1.GetValue()
		#print(fileName.encode("gbk"))
		#fileName = fileName.encode('utf-8')
		if fileName == "":
			return
		#如果原来已经打开了一个Excel文件，需要先关闭，否则再打开其他文件时，可能会出现不能打开新的文件的问题
		if self._xlBook:
			self._xlBook.Close(SaveChanges=0) 
			del self._xlApp
		#打开选择的文件
		self._easyexcel = EasyExcel(fileName)
		self._xlApp = self._easyexcel.xlApp
		self._xlBook = self._easyexcel.xlBook
		#差错提示
		if self._xlBook == None:
			dlg = wx.MessageDialog(self,u"提示",
							   u"打开文件错误",
							   wx.OK | wx.ICON_INFORMATION
							   )
			dlg.ShowModal()
			dlg.Destroy()

		#得到所有的Sheet名称，写到self.ID_CHOICE1
		self._sheet = {} 	#记录Excel表分页对应序号
		count = self._xlApp.Worksheets.Count
		self.ID_CHOICE1.Clear()
		for i in range(count):
			sheetName = self._xlApp.Worksheets(i+1).name
			self.ID_CHOICE1.Append(sheetName)
			self._sheet[ sheetName.strip() ] = i		

	#确定按键
	def OnOk(self, event):
		if self._xlBook == None:
			dlg = wx.MessageDialog(self,
					   u"请打开EXCEL表!", u"提示",
					   wx.OK | wx.ICON_ERROR
					   )
			dlg.ShowModal()
			dlg.Destroy()			 
			return #没有打开EXCEL文件
		selString = self.ID_CHOICE1.GetStringSelection()
		if selString == "":
			dlg = wx.MessageDialog(self,
							   u"请选择SHEET!", u"提示",
							   wx.OK | wx.ICON_ERROR
							   )
			dlg.ShowModal()
			dlg.Destroy() 
			return #没有选择SHEET
		if self._RangeData == None:
			print "ERROR self._RangeData is None"
			dlg = wx.MessageDialog(self,
							   u"EXCEL记录的数据是空!", u"提示",
							   wx.OK | wx.ICON_ERROR
							   )
			dlg.ShowModal()
			dlg.Destroy()
			return #EXCEL记录的数据是空
		
		#print "self._RangeData = ", str( self._RangeData  )
		print "start to deal sheet[%s]"%selString.encode("utf-8")
		print "ready to do lines = %d"%(len(self._RangeData) - self._beginLine)
		
		if self.ID_CHECKBOX1.IsChecked():
			self.importUseScript()
		else:
			self.importUseCFunc()
		
		self.OnCloseWindow(None)
	
	def importUseScript( self ):
		"""
		使用导入脚本导入
		"""
		entData = []
		for j in range(len(self._RangeData) - self._beginLine):
			excelData = self._RangeData[j + self._beginLine]
			
			isEmpty = True	#用来判断这个记录是否全部是空
			for i in range(len(excelData)):
				dataStr = str( excelData[i].encode('utf-8') )
				if dataStr and "#[" not in dataStr:
					isEmpty = False
			if isEmpty:
				continue
			
			#直接append数据，让导入脚本自己去解析
			entData.append(excelData)
		
		m = __import__(self._selModName)
		funName = self.ID_COMBOCTRL1.GetStringSelection()
		funName = funName.encode('utf-8')
		fun = getattr(m,funName) 
		sheetName = self.ID_CHOICE1.GetStringSelection()
		sheetName = sheetName.encode('utf-8')
		fun(self,self._entName,self._entDefn,entData,sheetName,self._xlApp)
	
	def importUseCFunc( self ):
		"""
		使用c++接口处理
		"""
		entData = []
		colTitleList = self.getColumnTitle()
		for j in range(len(self._RangeData) - self._beginLine):
			excelData = self._RangeData[j + self._beginLine]
			
			isEmpty = True	#用来判断这个记录是否全部是空
			for i in range(len(excelData)):
				dataStr = str( excelData[i].encode('utf-8') )
				if dataStr and "#[" not in dataStr:
					isEmpty = False
			if isEmpty:
				continue
			
			#需要按def文件打包数据
			recData = []
			for i in range(self.idgrid.GetNumberRows()):
				value = self.idgrid.GetCellValue(i,2)
				defType = self.idgrid.GetCellValue(i,1)
				title = self.idgrid.GetCellValue(i,0)
				if defType == "ARRAY":
					recData.append((''))
				else:
					index = colTitleList.index( title )
					data = ( excelData[index].encode('utf-8') )
					recData.append( data )
			
			if len(recData):
				entData.append( recData )
		
		overkey = self.ID_CHECKBOX2.IsChecked()
		arg = (self._entName,entData,self._entDefn,overkey)
		ResMgr.importFromExcel(arg)
		
		#提示
		dlg = wx.MessageDialog(self,
				   u"导入完成", u"提示",
				   wx.OK
				   )
		dlg.ShowModal()
		dlg.Destroy()

	def startBat(self, event):
		"""
		开始批处理
		"""
		if self._xlBook == None:
			dlg = wx.MessageDialog(self,
					   u"请打开EXCEL表!", u"提示",
					   wx.OK | wx.ICON_ERROR
					   )
			dlg.ShowModal()
			dlg.Destroy()			 
			return #没有打开EXCEL文件
		
		sheetnum = len(self.ID_CHOICE1.Items)
		if not sheetnum:
			dlg = wx.MessageDialog(self,
					   u"表为空!", u"提示",
					   wx.OK | wx.ICON_ERROR
					   )
			dlg.ShowModal()
			dlg.Destroy()					
			return 
		if not len(self.getExportSheetNames()):
			dlg = wx.MessageDialog(self,
					   u"没有找到需要批量导入的export表!", u"提示",
					   wx.OK | wx.ICON_ERROR
					   )
			dlg.ShowModal()
			dlg.Destroy()					
			return 			
		"""
		dlg = wx.MessageDialog(self,
				   u"批处理的表各列的含义应当是一致的才可以用批处理!", u"提示",
				   wx.OK | wx.ICON_INFORMATION
				   )
		dlg.ShowModal()
		dlg.Destroy()			
		"""	
		#清空表，不可恢复
		sql = "truncate table tbl_%s"%self._entName
		ResMgr.getDataRecord(sql)		#清空数据
		
		for i in xrange(sheetnum):
			self.ID_CHOICE1.Select(i)
			self.OnChoiceSheet(None)
			sheetname = self.ID_CHOICE1.Items[i]
			time.sleep(1)
			if "export" in sheetname.lower():
				self.OnOk(None)
				time.sleep(1)
				self.openExcelFile()
		self.OnCloseWindow(None)
		dlg = wx.MessageDialog(self,
				   u"批处理完毕！", u"提示",
				   wx.OK | wx.ICON_INFORMATION
				   )
		dlg.ShowModal()
		dlg.Destroy()		


	#退出
	def OnExit(self, event):
		self.OnCloseWindow(None)

	#把数字变成A-Z的字符表示
	def Numb2ABC(self,numb):
		i  = numb
		c  = ''
		ch = ''
		l = 0
		while i != 0 :
			j = (i-1)%26 + 0x41
			i = int((i-1)/26)
			c = c + chr(j)
		l = len(c)
		for i in range(l):
			ch = ch + c[l - i - 1]		
		return ch		

	#改变Sheet选择
	def OnChoiceSheet(self, event):
		#print "OnChoiceSheet=", self.ID_CHOICE1.GetStringSelection()
		#取得Sheet的最大列
		self.ID_TEXTCTRL4.SetValue(u"EXCEL数据行数:")
		selString = self.ID_CHOICE1.GetStringSelection()
		#设置对应的列
		for i in range(self.idgrid.GetNumberRows()):
			self.idgrid.SetCellValue(i,2,"")

		if selString == "":
			return #没有选择SHEET

		#打开SHEET
		useArea = self._xlApp.Worksheets(selString).UsedRange.Address
		openSheet = self._xlBook.Sheets.Add()
		formStr = "=\"\"&\'" + selString + "\'!RC"  #By GS: 如果公式中引用了其他工作表或工作簿中的值或单元格，并且那些工作簿或工作表的名称中包含非字母字符，那么您必须用单引号 (') 将这个名称括起来。
		openSheet.Range(useArea).Formula = formStr
		self._RangeData = openSheet.Range(useArea).Value

		if len(self._RangeData) < 2:
			return
		#得到excel字段显示的excel名称
		titNmae = self.ID_TEXTCTRL6.GetValue()
		indFiel = int(self.ID_TEXTCTRL5.GetValue())
		if titNmae != "" :
			tieNumb = int(titNmae)					#得到行号
			TitList = list(self._RangeData[tieNumb])  #得到数据
		#如果不是空就重新写单元格的下拉单元
		self._title = []
		tData = self._RangeData[indFiel]
		for i in range(len(tData) + 1):
			value = self.Numb2ABC(i) 
			if titNmae != "" and i != 0:
				value = value + ":" + TitList[i - 1]
			self._title.append(value)
		
		attr = wx.grid.GridCellAttr()
		attr.SetEditor(gridlib.GridCellChoiceEditor(self._title, False))
		self.idgrid.SetColAttr(2,attr)

		#设置对应的列
		for i in range(self.idgrid.GetNumberRows()):
			#读取变量和检索顺序
			steValue = self.idgrid.GetCellValue(i,0)
			steValue = steValue.encode("utf-8")
			if steValue in tData:
				indValue = list(tData).index(steValue)
				self.idgrid.SetCellValue(i,2,self._title[indValue + 1])
			else:
				self.idgrid.SetCellValue(i,2,"")

		self.ID_TEXTCTRL4.SetValue(u"EXCEL数据行数:" + str(len(self._RangeData)))

	def OnCloseMe(self, event):
		self.Close(True)

	def OnCloseWindow(self, event):
		#关闭EXCEL
		if self._xlBook != None:
			self._xlApp.Visible = True
			"""
			dlg = wx.MessageDialog(self,
					   u"按下确认键后，\r\n请在刚打开的EXCEL进行相应操作!", u"提示",
					   wx.OK | wx.ICON_ERROR
					   )
			dlg.ShowModal()
			dlg.Destroy()  
			"""
			self._xlBook.Close(SaveChanges=0) ####????语法问题？
			del self._xlApp
		self.Destroy()
