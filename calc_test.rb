#! ruby -EWindows-31J
# -*- mode:ruby; coding:Windows-31J -*-

require 'win32ole'
require 'Trail_Calc'
require 'fileutils'

STDOUT.sync = true

#
#  Calc操作　テスト　プログラム
#
def wait(n=0)
  sleep(n)
end

def err_check( err_msg,v1,v2)
  if v1 != v2
    print "Error : " + err_msg + "set = #{v1}, read = #{v2}\n"
  end
end

FileUtils.cp("Data1.ods","Data2.ods")

openCalcWorkbook("Data2.ods") do |book|

  print "ActiveSheet切り替え : DataSheet \n"
  book.set_active_sheet("DataSheet")
  
  # Calcドキュメントの読み込み
  sheet = book.get_sheet("DataSheet")

  # セルの読み書き
  sheet[0,3] = 'D1-'
  print "D1セルに'D1-'を書き込み、読み出した値 =「#{sheet[0,3]}」\n"
  err_check("文字列書き込みエラー",'D1-',sheet[0,3])
  
  sheet[0,4] = 100
  print "E1セルに100を書き込み、読み出した値 =「#{sheet[0,4]}」\n"
  err_check("数値書き込みエラー",100,sheet[0,4])
  wait
  
  #セルの背景色の読み書き
  sheet.set_color(0,3,rgb(0xff,0x00,0x00))
  print "D1セルの色 = #{sheet.color(0,3).to_s(16)}\n"
  err_check("D1セルの背景色設定",rgb(0xff,0x00,0x00).to_s(16),sheet.color(0,3).to_s(16))
  
  sheet.set_color(0,4,rgb(0,255,0))
  print "E1セルの色 = #{sheet.color(0,4).to_s(16)}\n"
  err_check("E1セルの背景色設定",rgb(0,255,0).to_s(16),sheet.color(0,4).to_s(16))
  
  sheet.set_color(0,5,rgb(0,0,255))
  print "F1セルの色 = #{sheet.color(0,5).to_s(16)}\n"
  err_check("F1セルの背景色設定",rgb(0,0,255).to_s(16),sheet.color(0,5).to_s(16))
  wait
  
  print "A22:B24に薄いピンクを設定\n"
  sheet.set_range_color(21,0,23,1,rgb(0xff,0xf0,0xf0))
  err_check("A22セルの背景色設定",rgb(0xff,0xf0,0xf0).to_s(16),sheet.color(21,0).to_s(16))
  err_check("B24セルの背景色設定",rgb(0xff,0xf0,0xf0).to_s(16),sheet.color(23,1).to_s(16))
  wait

  print "21行の下に1行を挿入\n"
  sheet.insert_rows(21)
  err_check("行の挿入",'a22',sheet[22,0])
  wait
  
  print "Fカラムの幅を6000に設定\n"
  sheet.set_width(5,6000)
  err_check("カラム幅設定",6000, sheet.get_width(5))
  print "F1に現在時刻を設定\n"
  t = Time.now
  sheet[0,5] = time_ruby2ooo(t)
  print "F1の時刻 = #{time_ooo2ruby(sheet[0,5]).to_s}\n"
  wait
  
  print "A25セルに式 '=1+2+3' を設定\n"
  sheet.set_formula(24,0,"=1+2+3")
  err_check("数式設定","=1+2+3", sheet.get_formula(24,0))
  err_check("式の値",6, sheet[24,0])

  
  # chartドキュメント取得
  print "chartドキュメント取得\n"
  
  # タイトルとサブタイトル
  chartDoc = sheet.get_chartdoc
  print "Title : #{chartDoc.get_title}\n"
  print "SubTitle : #{chartDoc.get_subtitle}\n"
  wait

  print "Title/Subtitle 書き換え\n"
  chartDoc.set_title("Temperature and Pressure")
  print "Title : #{chartDoc.get_title}\n"
  chartDoc.set_subtitle(Time.now.strftime("%Y/%m/%d"))
  print "SubTitle : #{chartDoc.get_subtitle}\n"
  wait

  print "日付に変換 X軸 min = " + time_ooo2ruby(chartDoc.get_Xmin).to_s + "\n"
  print "日付に変換 X軸 max = " + time_ooo2ruby(chartDoc.get_Xmax).to_s + "\n"
  print "X軸 min = #{chartDoc.get_Xmin}, max = #{chartDoc.get_Xmax}\n"
  wait
  
  print "min 書き換え\n"
  chartDoc.set_Xmin(sheet[1,0])
  print "X軸 min = #{chartDoc.get_Xmin}, max = #{chartDoc.get_Xmax}\n"

  print "max 書き換え\n"
  chartDoc.set_Xmax(sheet[19,0])
  print "X軸 min = #{chartDoc.get_Xmin}, max = #{chartDoc.get_Xmax}\n"
  print "日付に変換 X軸 min = " + time_ooo2ruby(chartDoc.get_Xmin).to_s + "\n"
  print "日付に変換 X軸 max = " + time_ooo2ruby(chartDoc.get_Xmax).to_s + "\n"
  wait
  
  print "ChartType : #{chartDoc.get_chartType}\n"
  print "Range : #{chartDoc.get_Range}\n"
  print "表示範囲を2-20行に変更\n"
  err_check("グラフ範囲変更",true, chartDoc.change_Xrange(2,20))
  print "Range : #{chartDoc.get_Range}\n"
  wait(2)
  
#------------------sheet切り替え-----------------------------------
  print "ActiveSheet切り替え : X^2 \n"
  book.set_active_sheet("X^2")
  sheet2 = book.get_sheet("X^2")
  wait

  chartDoc1 = sheet2.get_chartdoc(0)
  print "Title : #{chartDoc1.get_title}\n"
  print "SubTitle : #{chartDoc1.get_subtitle}\n"
  wait
  
  print "ChartType : #{chartDoc1.get_chartType}\n"
  print "Range : #{chartDoc1.get_Range}\n"
  print "表示範囲を2-11行に変更\n"
  err_check("グラフ範囲変更",true, chartDoc1.change_Xrange(2,11))
  print "Range : #{chartDoc1.get_Range}\n"
  wait

  chartDoc2 = sheet2.get_chartdoc(1)
  print "Title : #{chartDoc2.get_title}\n"
  print "SubTitle : #{chartDoc2.get_subtitle}\n"
  wait

  print "ChartType : #{chartDoc2.get_chartType}\n"
  print "Range : #{chartDoc2.get_Range}\n"
  print "表示範囲を2-11行に変更\n"
  err_check("グラフ範囲変更",true, chartDoc2.change_Xrange(2,11))
  print "Range : #{chartDoc2.get_Range}\n"
  
  print "チャート書き出し\n"
  chartDoc2.save("chart.png")
  
  wait(2)

#------------------sheet切り替え-----------------------------------
  print "ActiveSheet切り替え : 表1 \n"
  book.set_active_sheet("表1")
  sheet3 = book.get_sheet("表1")
  wait
  
  print "2行目から3行目までをグループ化\n"
  sheet3.group_row(1,2)
  wait

  print "2行目から5行目までをグループ化\n"
  sheet3.group_row(1,4)
  wait

  print "3列目から4列目までをグループ化\n"
  sheet3.group_column(2,3)
  wait

  print "A8:E10にwrap設定\n"
  sheet3.wrap(7,0,9,5)
  wait

  print "A1:E10に罫線枠\n"
  sheet3.box(0,0,9,4)
  wait
  
  
  print "A1:C1をマージ\n"
  sheet3.set_color(0,0,rgb(255,100,100))
  sheet3.merge(0,0,0,2)
  wait

  print "A6:C6をマージ\n"
  sheet3.set_color(5,0,rgb(255,50,50))
  sheet3.merge(5,0,5,2)
  wait

  print "A6:C6をマージ解除\n"
  sheet3.merge_off(5,0,5,2)
  wait

  print "A10=center\n"
  sheet3.center(9,0)
  print "A10=TOP\n"
  sheet3.v_top(9,0)
  wait

  print "B10:C10=center\n"
  sheet3.centers(9,1,9,2)
  print "B10:C10=TOP\n"
  sheet3.v_tops(9,1,9,2)
  wait

  print "D10:E10=right\n"
  sheet3.horizontals(9,3,9,4,3)
  print "B10:C10=bottom\n"
  sheet3.verticals(9,3,9,4,3)
  wait
  
  print "書式コピー A10 → A11\n"
  sheet3.format_copy1(9,0,10,0)
  print "書式コピー B10 → B11:C11\n"
  sheet3.format_copy(9,1,10,1,10,2)
  print "書式コピー A10:E10 → A0 2行\n"
  sheet3.format_range_copy(9,0,4,  0,0, 2)
  wait
  
  print "コピーA10:E10 → A12\n"
  sheet3.copy(9,0,9,4,  11,0)

  print "コピーA10:E10 → A12\n"
  sheet3.copy(9,0,9,4,  11,0)
  
  print "行削除 $13:$16\n"
  sheet3.remove_rows(12,4)
  wait(2)

#------------------------------------------------------------------
  book.save
  book.save('Data3.xls')
  book.save('Data3.xlsx')
  
end
