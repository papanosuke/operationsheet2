# -*- coding: utf-8 -*-
=begin
*******************************************************
スケジュール出力システム　written by ぱぱのすけ
*******************************************************
=end
require 'date'
require 'win32ole'
require_relative 'TCalendar'
#*************************************************************
#サブルーチン
#*************************************************************
#*******************
#休日を除外する
#*******************
def holiday_hantei(syoribi,kyujitu)
#
#祝日確認用TCalendarHolidayクラス
  while true
    date_h = TCalendar.new( syoribi.year, syoribi.month )
    if date_h.status(syoribi.day) == "holiday"  or date_h.status(syoribi.day) == "nenmatsu" or syoribi.wday == 0 or syoribi.wday == 6
      syoribi -= 1 if kyujitu == 1
      syoribi += 1 if kyujitu == 2
    else
      break
    end
  end
#
  return syoribi
#
end
#*******************
#パターン設定
#*******************
def pertern_settei(syoribi,pertern,perterndays)
#
#祝日確認用TCalendarHolidayクラス
  days = 0
  days_max = perterndays
  while days < days_max
    syoribi -= 1 if pertern == 1
    syoribi += 1 if pertern == 2
    date_h = TCalendar.new( syoribi.year, syoribi.month )
    if date_h.status(syoribi.day) == "holiday"  or date_h.status(syoribi.day) == "nenmatsu" or syoribi.wday == 0 or syoribi.wday == 6
    else
      days += 1
    end
  end
#
  return syoribi
#
end
#*******************
#処理日を判定する
#*******************
def syoribi_hantei(book,syoribi,gyomu)
#
#メインループ------------------------------↓↓↓↓↓↓
  syoribi_hantei_w =''
  ws = book.Worksheets.Item('syori')
  ws.UsedRange.Rows.each do |row|
    record = []
    row.Columns.each do |cell|
      record << Kconv.toutf8(cell.Value.to_s)
    end
#
#処理基本日
    date_r = Date.new( syoribi.year, syoribi.month, syoribi.day)
    gyomucd = record[0].to_i
    syoricd = record[1].to_i
    syoriname = record[2]
    syorituki = record[3].to_i
    syoricycle = record[4].to_i
    pertern = record[5].to_i
    perterndays = record[6].to_i
    kyujitu = record[7].to_i
    teiki = record[8].to_i
    hyoji = record[9].to_i
    hyojiname = record[10]
#
#処理月判定
    if syorituki == 1
      date_r = date_r.prev_month
    elsif syorituki == 2
      date_r = date_r.next_month
    end
#
#     print"debug 処理②=",date_r," syoriname=",syoriname,"\n"
#
#処理サイクル
#debug start
#該当月のみ選択
    if teiki != 99
      if date_r.month  != teiki
        next
      end
    end
#debug end
#
    if syoricycle == 0
      syorinaiyo[syorisu] = syoriname
      syorisu += 1
    elsif syoricycle == 99
      date_r = Date.new(date_r.year,date_r.month,-1)
    else
#     print"debug year=",date_r.year," month=",date_r.month," syoricycle=",syoricycle,"\n"
#debug start
#==============================================================#
#処理サイクルに29日以降の設定があると、エラーになる場合の回避
#==============================================================#
      if date_r.month == 4 or date_r.month == 6 or date_r.month == 9 or date_r.month == 11
         if syoricycle == 31
     print"確認１ year=",date_r.year," month=",date_r.month," syoricycle=",syoricycle,"\n"
            next
         end
      end
      if date_r.month == 2
         if syoricycle >  29
#    print"確認２ year=",date_r.year," month=",date_r.month," syoricycle=",syoricycle,"\n"
            next
         elsif syoricycle == 29
            if (date_r.year % 100 == 0)
              if (date_r.year % 400 == 0)
#    print"うるう年１ year=",date_r.year," month=",date_r.month," syoricycle=",syoricycle,"\n"
              else
#    print"確認３ year=",date_r.year," month=",date_r.month," syoricycle=",syoricycle,"\n"
                next
              end
            elsif (date_r.year % 4 == 0)
#    print"うるう年２ year=",date_r.year," month=",date_r.month," syoricycle=",syoricycle,"\n"
            else
#    print"確認４ year=",date_r.year," month=",date_r.month," syoricycle=",syoricycle,"\n"
              next
            end
         end
      end
#==============================================================#
#処理サイクルに29日以降の設定があると、エラーになる場合の回避
#==============================================================#
#debug end
      date_r = Date.new(date_r.year,date_r.month,syoricycle)
    end
#
    date_r = holiday_hantei(date_r,kyujitu)
#
    date_s = Date.new(date_r.year,date_r.month,date_r.day)
#     print"debug 処理③=",date_r," syoriname=",syoriname,"\n"
#
#処理パターン
    date_r = pertern_settei(date_r,pertern,perterndays)
#
#     print"debug 処理④=",date_r," syoriname=",syoriname,"\n"
#
    if hyoji == 1
      if gyomucd == gyomu
        if teiki == 99 or syoribi.month == teiki
          if syoribi == date_r
#
#             print "結果 syoribi=",syoribi," syoriname=",syoriname,date_s,"\n"
#
#              if syoribi.month == date_s.month
#                syoribi_hantei_w << date_s.day.to_s + "日:" + hyojiname + " "
#              else
#                syoribi_hantei_w << date_s.month.to_s + "/" + date_s.day.to_s + ":" + hyojiname + " "
#              end
             syoribi_hantei_w << date_s.to_s + ";"+hyojiname + ";"
          end
        end
      end
    end
#
  end
  return syoribi_hantei_w

#メインループ------------------------------↑↑↑↑↑↑
#
end

