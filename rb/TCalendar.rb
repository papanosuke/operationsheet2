# -*- coding: utf-8 -*-
class TCalendar
  require 'kconv'

  def initialize(year = 0, month = 0)
    @now = Time.now
    @year = if year == 0 then @now.year else year end
    @month = if month == 0 then @now.month else month end
    @holiday = []
    @holiday_name = []
    @nenmatsu = []
    @mday_arr = []
    1.upto(31) do |d|
      begin
        @mday_arr[d] = Time.local(@year, @month, d, 0, 0, 0)
        if d > 28 and @mday_arr[d].month != @month
          @mday_arr[d] = nil
        end
      rescue ArgumentError
        @mday_arr[d] = nil
      end
      @holiday[d] = false
#nenmatsu追加
      @nenmatsu[d] = false
    end
    month_name = @mday_arr[1].strftime('%b')

# holiday_str.txtのデータ形式は次のようになっています。
# 月名<space>日付<space>有効年<space>年末判定フラグ<space>コメント
# HM2 Happy Monday(2nd monday)
# HM3 Happy Monday(3rd monday)

#holiday.txtを読み込む
    file = open("holiday_str.txt")
    while infile = file.gets do

      infile.split(/\n/).each do |l|
        next if l == '' or l =~ /^\#/
        l.sub!(/\#.*$/, '')
  #
  #nenmatsu追加 f
  #     m, d, y, c = l.split(/\s+/, 4)
        m, d, y, f, c = l.split(/\s+/, 5)
        c = Kconv.toutf8(c.to_s)

        if y != '0'
          if y[0,1] == '-'
            next if @year > y[1,4].to_i
          elsif y[-1,1] == '-'
            next if @year < y[0,4].to_i
          elsif y[4,1] == '-'
            next if @year < y[0,4].to_i || @year > y[5,4].to_i
          end
        end

        if month_name == m
          case d
            when 'SHUNBUN'
              d = syunbun(@year).to_s
            when 'SYUBUN'
              d = syubun(@year).to_s
            when 'HM2'
              d = nMonday(2).to_s
            when 'HM3'
              d = nMonday(3).to_s
          end

  #debug start
  ########@holiday[d.to_i] = true
          if f == '1'
  #nenmatsu追加 f
            @nenmatsu[d.to_i] = true
          else
            @holiday[d.to_i] = true
          end
          @holiday_name[d.to_i] = c
  #debug end

        end
      end
    end
    file.close

    if @year >= 1986
      i = 0
      while i < 31 - 2
        # 「国民の休日」判定
        # 当日が祝日       次の日が祝日でない           日曜日でない                   次の次の日が祝日
        if @holiday[i] and @holiday[i + 1] == false and @mday_arr[i + 1].wday != 0 and @holiday[i + 2]
          @holiday[i + 1] = true
          @holiday_name[i + 1] = '国民の休日'
          i += 1                # skip
        end
        i += 1
      end
    end

  end
  attr_reader :holiday_name, :year, :month

  def nMonday(n)
    count = 0
    @mday_arr.each_index do |d|
      next if d < 1
      count += 1 if @mday_arr[d].wday == 1
      return d if count == n
    end
  end

  def today?(mday)
    @now.year == @year and @now.month == @month and @now.mday == mday
  end

  def wday(mday)
    return nil unless @mday_arr[mday]
    @mday_arr[mday].wday
  end

  # 2005 からの振替休日 連続する祝日が日曜日にかかると祝日の終りの次の平日を振替休日に
  def furikae2005(mday, wday)
    year = @mday_arr[mday].year
    if year < 2005
      return 'workday'
    end
    if mday <= wday
      return 'workday'
    end
    (1..wday).each do |i|
      if @holiday[mday - i] == false
        return 'workday'
      end
    end
    @holiday_name[mday] = '振替休日'
    'holiday'
  end

  def status(mday)
    return nil unless @mday_arr[mday]
    return 'holiday' if @holiday[mday]
#nenmatsu追加
    return 'nenmatsu' if @nenmatsu[mday]

    wday = @mday_arr[mday].wday

    case wday
      when 1
        if @mday_arr[mday].year >= 1973 and mday > 1 and @holiday[mday - 1]
#*==============================================DEBUG START
          @holiday_name[mday] = '振替休日'
#*==============================================DEBUG END
          'holiday'
        else
          'workday'
        end
      when 2..5
        furikae2005(mday, wday)
      when 0, 6
        'weekend'
    end
  end

  def header
    msg = @mday_arr[1].strftime('%B %Y').center(20)
    msg + "\n" + 'Su Mo Tu We Th Fr Sa' + "\n"
  end

#| From: hajima atmark crimson.gen.u-tokyo.ac.jp (Ryoichi Hajima)
#| Newsgroups: fj.questions.misc
#| Subject: Re: vernal/autumnal equinox
#| Message-ID: <HAJIMA.94Jul13161542@tanelorn.gen.u-tokyo.ac.jp>
#| Date: 13 Jul 94 07:15:42 GMT
#|
#| 春分日　(31y+2213)/128-y/4+y/100    (1851年-1999年通用)
#| 　　　　(31y+2089)/128-y/4+y/100    (2000年-2150年通用)
#|
#| 秋分日　(31y+2525)/128-y/4+y/100    (1851年-1999年通用)
#| 　　　　(31y+2395)/128-y/4+y/100    (2000年-2150年通用)

  def syunbun(year)
    if year > 2150
      STDERR.print "over year's: #{year}\n"  #'
      exit 1
    end
    v = if year < 2000 then 2213 else 2089 end
    (31 * year + v)/128 - year/4 + year/100
  end

  def syubun(year)
    if year > 2150
      STDERR.print "over year's: #{year}\n" #'
      exit 1
    end
    v = if year < 2000 then 2525 else 2395 end
    (31 * year + v)/128 - year/4 + year/100
  end
end

if $0 == __FILE__
  puts 'テスト'

  cal = TCalendar.new(2016, 1)
  print cal.header
  p cal
end
