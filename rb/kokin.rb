# -*- coding: utf-8 -*-
=begin
*******************************************************
公金振替　written by ぱぱのすけ
*******************************************************
=end
require 'date'
require 'win32ole'
require_relative 'hantei'
require_relative 'excel'
require 'kconv'
#*************************************************************
#メインルーチン
#*************************************************************
#date.txtを読み込む
file = open("date.txt")
while infile = file.gets do
  date_w = infile.split(";")
  syorinen = date_w[0].to_i
  syorituki = date_w[1].to_i
end
file.close
#
print"**********************************************************************\n"
print"スケジュール作成システム②　公金振替　written by ぱぱのすけ\n"
print"**********************************************************************\n"
#
#
#年月入力
while true
  print "*----------------------------------------------------------------------------*\n"
  print " 業務予定を作成する年を指定してください 西暦9(4)\n"
  print "*----------------------------------------------------------------------------*\n"
  print "==> "
#  syorinen = STDIN.gets.chomp!
  print syorinen,"\n"
  print "*----------------------------------------------------------------------------*\n"
  print " 業務予定を作成する月を指定してください 9(2)\n"
  print "*----------------------------------------------------------------------------*\n"
  print "==> "
# syorituki = STDIN.gets.chomp!
  print syorituki,"\n"
  print "*----------------------------------------------------------------------------*\n"
  print " 作成する業務を指定してください 9(1)\n"
  print "*----------------------------------------------------------------------------*\n"
  openExcelWorkbook('支払業務予定.xls') do |book|
    ws = book.Worksheets.Item('gyomu')
    x = 1
    ws.UsedRange.Rows.each do |row|
      record = []
      row.Columns.each do |cell|
        record << Kconv.toutf8(cell.Value.to_s)
      end
      print record[0].to_i ,"  " , record[1], "\n"
    end
  end
  print "==> "
# gyomu_w =gets.chomp!.to_i
  gyomu_w = 2
  print gyomu_w,"\n"

  break
end
#日付クラス作成（日付-1で最終日取得）
date_w = Date.new(syorinen.to_i,syorituki.to_i,-1)

#エクセル出力
openExcelWorkbook('支払業務予定.xls') do |book|
#
#メインループ------------------------------↓↓↓↓↓↓
  ws = book.Worksheets.Item(date_w.month.to_s)
  ws.Cells(1,4).Value = Date.new(date_w.year,date_w.month,1).to_s

#
#サブループ------------------------------↓↓↓↓↓↓
  gyo = 26
  (1..date_w.day).each do | day |
#
#日付表示用dateクラス
    date_r = Date.new(date_w.year,date_w.month,day)
#
##print "処理 ",date_r," 曜日 ",date_r.wday,"\n"
    hantei = syoribi_hantei(book,date_r,gyomu_w)
#   print "処理日=",date_r," 処理=",hantei,"\n"
    if hantei == ""
      print "処理日=",date_r," 処理=",hantei,"\n"
    else
      w_area = hantei.split(";")
      date_h = Date.new(w_area[0][0,4].to_i,w_area[0][5,2].to_i,w_area[0][8,2].to_i)
      if date_r.month == date_h.month
        print "処理日=",date_r," 処理=",date_h.day.to_s+"日:"+w_area[1],"\n"
      else
        print "処理日=",date_r," 処理=",date_h.month.to_s+"/"+date_h.day.to_s+w_area[1],"\n"
      end
    end
#
    if hantei != ""
      ws.Cells(gyo,8).Value = date_r.year.to_s + "/" + date_r.month.to_s  + "/" + date_r.day.to_s
      gyo += 1
    end
#
  end
#サブループ------------------------------↑↑↑↑↑↑
#
  ws.select
  book.save
#
end
#メインループ------------------------------↑↑↑↑↑↑
print "\n"
print "**********************************************","\n"
print "    ",date_w.year,"年",date_w.month,"月の公金振替の作成終了しました。","\n"
print "**********************************************","\n"
