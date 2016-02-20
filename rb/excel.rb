# -*- coding: utf-8 -*-
=begin
*******************************************************
エクセル読込　written by ぱぱのすけ
*******************************************************
=end
require 'win32ole'
#*************************************************************
#サブルーチン
#*************************************************************
#*****************
#*excelフルパス取得
#*****************
def getAbsolutePath filename
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  return fso.GetAbsolutePathName(filename)
end
#*****************
#*ExcelFileオープン
#*****************
def openExcelWorkbook filename
  filename = getAbsolutePath(filename)
  xl = WIN32OLE.new('Excel.Application')
  xl.Visible = false
  xl.DisplayAlerts = false
  book = xl.Workbooks.Open(filename)
  begin
    yield book
  ensure
    xl.Workbooks.Close
    xl.Quit
  end
end
#*****************
#*Excelfile作成
#*****************
def createExcelWorkbook
  xl = WIN32OLE.new('Excel.Application')
  xl.Visible = false
  xl.DisplayAlerts = false
  book = xl.Workbooks.Add()
  begin
    yield book
  ensure
    xl.Workbooks.Close
    xl.Quit
  end
end
