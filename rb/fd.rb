# -*- coding: utf-8 -*-
=begin
*******************************************************
フロッピー・ディスク取扱日程表　written by ぱぱのすけ
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
file = open("date.txt","r")
while infile = file.gets do
  date_w = infile.split(";")
  syorinen = date_w[0].to_i
  syorituki = date_w[1].to_i
end
file.close
#
print"**********************************************************************\n"
print"スケジュール作成システム③　FD取扱日程表　written by ぱぱのすけ\n"
print"**********************************************************************\n"
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
#
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
# gyomu_w =STDIN.gets.chomp!.to_i
  gyomu_w = 3
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
  gyo = 4
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
        print "処理日=",date_r," 処理=",date_h.month.to_s+"/"+date_h.day.to_s+":"+w_area[1],"\n"
      end
    end
#
#   print hantei
#
    if hantei != ""
      ws.Cells(gyo,7).Value = date_r.year.to_s + "/" + date_r.month.to_s  + "/" + date_r.day.to_s
      w_area = hantei.split(";")
      ws.Cells(gyo,8).Value = w_area[0][0,4] + "/" + w_area[0][5,2]  + "/" + w_area[0][8,2]
#
      if w_area[1].slice(14,1).to_i == 1
        date_yokujitu = Date.new(w_area[0][0,4].to_i,w_area[0][5,2].to_i,w_area[0][8,2].to_i) + 1
      else
        date_yokujitu = Date.new(w_area[0][0,4].to_i,w_area[0][5,2].to_i,w_area[0][8,2].to_i)
      end

      date_yokujitu = holiday_hantei(date_yokujitu,2)
      ws.Cells(gyo,9).Value = date_yokujitu.to_s
#

      ws.Cells(gyo,10).Value = w_area[1]
      gyo += 1
    end
#
  end
#サブループ------------------------------↑↑↑↑↑↑
  (gyo..13).each do | day |
    ws.Cells(day,7).Value = ""
    ws.Cells(day,8).Value = ""
    ws.Cells(day,9).Value = ""
    ws.Cells(day,10).Value = ""
  end
#
  ws.select
  book.save
#
end
#メインループ------------------------------↑↑↑↑↑↑
print "\n"
print "**************************************************","\n"
print "    ",date_w.year,"年",date_w.month,"月のFD取扱日程表の作成終了しました。","\n"
print "**************************************************","\n"
