9# -*- coding: utf-8 -*-
=begin
*************************************************************
スケジュール作成システム①　予定表　written by ぱぱのすけ\n"
*************************************************************
=end
require 'date'
require 'win32ole'
require_relative './TCalendar'
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
#comment　テストコミット用
#
print"**********************************************************************\n"
print"スケジュール作成システム①　予定表　written by ぱぱのすけ\n"
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
  gyomu_w = 1
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
  (1..date_w.day).each do | day |
#
#日付表示用dateクラス
    date_r = Date.new(date_w.year,date_w.month,day)
#
##print "処理 ",date_r," 曜日 ",date_r.wday,"\n"
    hantei = syoribi_hantei(book,date_r,gyomu_w)
####print "処理日=",date_r," 処理=",hantei,"\n"
    if hantei == ""
      print "処理日=",date_r," 処理=",hantei,"\n"
    else
      w_area = hantei.split(";")
      w_area_all = ""
      w_area.each do |str|
        if str[4,1] == "-"
          date_h = Date.new(str[0,4].to_i,str[5,2].to_i,str[8,2].to_i)
          if date_r.month == date_h.month
            w_area_all << date_h.day.to_s+"日:"
          else
            w_area_all << date_h.month.to_s+"/"+date_h.day.to_s+":"
          end
        else
          w_area_all << str + " "
        end
      end
      print "処理日=",date_r," 処理=",w_area_all,"\n"
    end
#
    ws.Cells(day+2,1).Value = day
    ws.Cells(day+2,3).Value = w_area_all
#
    date_s = TCalendar.new( date_w.year, date_w.month )
    if date_s.status(day) == "holiday"  or date_s.status(day) == "nenmatsu"
      ws.Cells(day+2,3).Value = "（"+date_s.holiday_name[day] +"）"
    end
  end
#サブループ------------------------------↑↑↑↑↑↑
#
#29～31日は表示される月とそうでない月があるため、その対応
  if date_w.day != 31
    (date_w.day+1..31).each do | day |
      ws.Cells(day+2,1).Value = ""
      ws.Cells(day+2,3).Value = ""
      ws.Cells(day+2,6).Value = ""
    end
  end
#
  ws.select
  book.save
#
end
#メインループ------------------------------↑↑↑↑↑↑
print "\n"
print "*****************************************","\n"
print "    ",date_w.year,"年",date_w.month,"月予定の作成終了しました。","\n"
print "*****************************************","\n"

