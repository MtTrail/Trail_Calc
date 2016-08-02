#! ruby -EWindows-31J
# -*- mode:ruby; coding:Windows-31J -*-
 
require "Trail_Calc/version"
require 'win32ole'
require 'date'

#= Trail_Calc
#== OpenOffice.org Calc用 ruby拡張モジュール
#
#*  OpenOffice.org(LibreOffice)のCalcをrubyから操作するためのruby拡張モジュールです。
#
#Authors:: Mt.Trail
#Version:: Ver. 1.03 2016/08/02 Mt.Trail
#* Ver. 1.03  : gem化
#* Ver. 1.02  : excel形式等での保存機能を追加、互換性には問題あり、チャートの範囲変更のbugfix
#* ver. 1.01  : チャートの画像出力機能追加
#* Ver. 1.00  : 初回公開版
#
#Copyright:: Copyright (C) Mt.Trail 2015 All rights reserved.
#License:: GPL version 2
#
#
#==目的
#* rubyからOpenOffice.orgのCalcを簡単に操作すること。
# 
#* Excelの無い環境で表やグラフを自動更新するために作成しました。
#* 基本的にはシートのレイアウトやチャートの形式等はCalcで作成しデータの追加や変更を自動化する目的で作成しています。
#* ですから、詳細なプロパティの設定や、新しいチャートの作成等はサポートしていません。
#* 書式を設定したテンプレート領域を用意しておき書式のコピーで書式を設定してください。
#
#= 利用方法
#*  rubyスクリプトでrequireして利用してください。
#
#== 使用例
#     require 'trail_calc'
#     
#     openCalcWorkbook("file.ods") do |book|
#       sheet = book.get_sheet("DataSheet")
#       sheet[0,3] = 1
#       ...
#       book.save
#     end
#
#==ドキュメント生成
#
#  rdoc --title Trail_Calc --main Trail_calc  --all Trail_calc.rb

module TrailCalc
  # Dumy module for RDOC
end


#----------------------------------------------------
#== 汎用関数
#=== 色コード生成
# 赤、緑、青を8bit値(0から255)で指定して24bitの色コードを作成します。
#
# _red_   :: 赤(0-255)
# _green_ :: 緑(0-255)
# _blue_  :: 青(0-255)
#
def rgb(red, green, blue)
  return red << 16 | green << 8 | blue
end

#=== OpenOffice.orgの時刻からrubyの時刻へ変換
#
# _t_:: OOoの日付時刻を表すdoubleの数値
#
def time_ooo2ruby(t)  ## t はOOoの日付時刻を表すdoubleの数値
  #OOoの基準はデフォルトだと1899/12/30日
  #Appleは1904/1/1, StartOfficeは1900/1/1
  #rubyは1970//1/1 9:0:0 (JSTだから+9時間)
  diff = -2209194000  # Time.at(Time.new(1899,12,30)).to_i
  Time.at(t * 24 * 60 * 60 + diff +0.1) ## 0.1は微妙な変換誤差の補正用
end

#=== rubyの時刻からOpenOffice.orgの時刻へ変換
#
# _t_:: rubyのTimeオブジェクト
#
def time_ruby2ooo(t)  ## t rubyのTimeオブジェクト
  diff = -2209194000  # Time.at(Time.new(1899,12,30)).to_i
  d = (t.to_i - diff)/(24.0 * 60.0 * 60.0)
end

#=== 絶対パスの取得
#
# _filename_ :: ファイル名
#
def getAbsolutePath filename
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  fn = fso.GetAbsolutePathName(filename)
  fn.gsub!(/\\/,'/')
  fn
end


#----------------------------------------------------
#==  チャートドキュメント操作
#
# チャートドキュメントクラスの拡張モジュール
# チャートドキュメント取り出し時に自動的に組み込まれる。

module CalcChartDoc

  #=== チャートのタイトル取得
  #
  # ret::文字列
  def get_title
    self.getTitle.String
  end
  
  #=== チャートのタイトル設定
  #
  def set_title(name)
    t = self.getTitle
    t.String = name
  end

  #=== チャートのサブタイトル取得
  #
  # ret::文字列
  def get_subtitle
    self.getSubTitle.String
  end

  #=== チャートのサブタイトル設定
  #
  def set_subtitle(name)
    t = self.getSubTitle
    t.String = name
  end

  #=== 変更されたか調べる
  # OOo標準
  #
  # isModified()
  # Modified
  # ret::Bool
  
  #=== 保存
  # OOo標準
  #
  # store()

  #=== X軸の最小値を取得
  #
  # ret::double
  def get_Xmin
    self.Diagram.XAxis.Min
  end

  #=== X軸の最大値を取得
  #
  # ret::double
  def get_Xmax
    self.Diagram.XAxis.Max
  end

  #=== X軸の最小値を設定
  #
  # _t_:: OOoでのdoubleの値、日時の場合もdoubleで表現される。
  def set_Xmin(t)  ## t はOOoでの値
    self.Diagram.XAxis.Min = t
  end

  #=== X軸の最大値を設定
  #
  # _t_:: OOoでのdoubleの値、日時の場合もdoubleで表現される。
  def set_Xmax(t)  ## t はOOoでの値
    self.Diagram.XAxis.Max = t
  end


#== チャートのテンプレート
#
# チャートのテンプレートの種類を表す文字列の配列。
# 先頭をデフォルト値とする。

#=== 棒グラフ
#  Role
#    categories
#    values-y
ChartTemplateBar = [
  "com.sun.star.chart2.template.Column",                          #  0 # 縦
  "com.sun.star.chart2.template.StackedColumn",                   #  1 # 縦積み上げ
  "com.sun.star.chart2.template.PercentStackedColumn",            #  2 # 縦積み上げパーセント
  "com.sun.star.chart2.template.ThreeDColumnDeep",                #  3 # 3D  縦奥行きあり
  "com.sun.star.chart2.template.ThreeDColumnFlat",                #  4 # 3D  縦奥行きなし
  "com.sun.star.chart2.template.PercentStackedThreeDColumnFlat",  #  5 # 3D  縦積み上げ
  "com.sun.star.chart2.template.StackedThreeDColumnFlat",         #  6 # 3D  縦積み上げパーセント
  "com.sun.star.chart2.template.Bar",                             #  7 # 横
  "com.sun.star.chart2.template.StackedBar",                      #  8 # 横積み上げ
  "com.sun.star.chart2.template.PercentStackedBar",               #  9 # 横積み上げパーセント
  "com.sun.star.chart2.template.ThreeDBarDeep",                   # 10 # 3D  横奥行きあり
  "com.sun.star.chart2.template.ThreeDBarFlat",                   # 11 # 3D  横奥行きなし
  "com.sun.star.chart2.template.StackedThreeDBarFlat",            # 12 # 3D  横積み上げ
  "com.sun.star.chart2.template.PercentStackedThreeDBarFlat"      # 13 # 3D  横積み上げパーセント
]
#
#=== 円グラフ
#  Role
#    categories
#    values-y
ChartTemplatePie = [
  "com.sun.star.chart2.template.Pie",                      #  0 # 扇型
  "com.sun.star.chart2.template.PieAllExploded",           #  1 # 扇型分解
  "com.sun.star.chart2.template.ThreeDPie",                #  2 # 3D 扇型
  "com.sun.star.chart2.template.ThreeDPieAllExploded"      #  3 # 3D 扇型分解
]
#=== ドーナツグラフ
ChartTemplateDonut = [
  "com.sun.star.chart2.template.Donut",                    #  0 # ドーナツ
  "com.sun.star.chart2.template.DonutAllExploded",         #  1 # ドーナツ分解
  "com.sun.star.chart2.template.ThreeDDonut",              #  2 # 3D ドーナツ
  "com.sun.star.chart2.template.ThreeDDonutAllExploded"    #  3 # 3D ドーナツ分解
]
#
#=== エリアグラフ
#  Role
#    categories
#    values-y
ChartTemplateArea = [
  "com.sun.star.chart2.template.Area",                     #  0 #   エリア
  "com.sun.star.chart2.template.StackedArea",              #  1 #   積み上げ
  "com.sun.star.chart2.template.ThreeDArea",               #  2 #   3D
  "com.sun.star.chart2.template.StackedThreeDArea",        #  3 #   3D 積み上げ
  "com.sun.star.chart2.template.PercentStackedArea",       #  4 #   積み上げパーセント
  "com.sun.star.chart2.template.PercentStackedThreeDArea"  #  5 #   3D 積み上げパーセント
]
#
#=== 折れ線
ChartTemplateLine = [
  "com.sun.star.chart2.template.Line",                     #  0 #   線
  "com.sun.star.chart2.template.Symbol",                   #  1 #   点
  "com.sun.star.chart2.template.LineSymbol",               #  2 #   点と線
  "com.sun.star.chart2.template.ThreeDLine",               #  3 #   3D 線
  "com.sun.star.chart2.template.ThreeDLineDeep",           #  4 #   3D 線奥行きあり
  "com.sun.star.chart2.template.StackedLine",              #  6 #   線積み上げ
  "com.sun.star.chart2.template.StackedSymbol",            #  5 #   点積み上げ
  "com.sun.star.chart2.template.StackedLineSymbol",        #  7 #   点と線積み上げ
  "com.sun.star.chart2.template.StackedThreeDLine",        #  8 #   3D 線積み上げ
  "com.sun.star.chart2.template.PercentStackedSymbol",     #  9 #   点積み上げパーセント
  "com.sun.star.chart2.template.PercentStackedLine",       # 10 #   線積み上げパーセント
  "com.sun.star.chart2.template.PercentStackedLineSymbol", # 11 #   点と線積み上げパーセント
  "com.sun.star.chart2.template.PercentStackedThreeDLine"  # 12 #   3D 線積み上げパーセント
]
#
#=== 散布図
#  Role
#    categories: 系列のラベル
#    values-x : X データ
#    values-y: Y データ
ChartTemplateScatter = [
  "com.sun.star.chart2.template.ScatterLine",         #  0 # ラインのみ
  "com.sun.star.chart2.template.ScatterLineSymbol",   #  1 # ラインとデータ点
  "com.sun.star.chart2.template.ScatterSymbol",       #  2 # データ点
  "com.sun.star.chart2.template.ThreeDScatter"        #  3 # 3D
]
#
#=== レーダー網
#  Role
#    categories
#    values-y
ChartTemplateNet = [
  "com.sun.star.chart2.template.NetLine",                #  0 # 線
  "com.sun.star.chart2.template.Net",                    #  1 # 点と線
  "com.sun.star.chart2.template.NetSymbol",              #  2 # 点
  "com.sun.star.chart2.template.StackedNet",             #  3 # 積み上げ点と線
  "com.sun.star.chart2.template.StackedNetLine",         #  4 # 積み上げ線
  "com.sun.star.chart2.template.StackedNetSymbol",       #  5 # 積み上げ点
  "com.sun.star.chart2.template.PercentStackedNet",      #  6 # 点と線積み上げパーセント
  "com.sun.star.chart2.template.PercentStackedNetLine",  #  7 # 線積み上げパーセント
  "com.sun.star.chart2.template.PercentStackedNetSymbol" #  8 # 点積み上げパーセント
]

#=== レーダー網(塗りつぶし)
ChartTemplateFilledNet = [
  "com.sun.star.chart2.template.FilledNet",              #  0 # 3.2
  "com.sun.star.chart2.template.StackedFilledNet",       #  1 # 3.2
  "com.sun.star.chart2.template.PercentStackedFilledNet" #  2 # 3.2
]

#=== ストックチャート
ChartTemplateStock = [
  "com.sun.star.chart2.template.StockLowHighClose",           #  0 #
  "com.sun.star.chart2.template.StockOpenLowHighClose",       #  1 #
  "com.sun.star.chart2.template.StockVolumeLowHighClose",     #  2 #
  "com.sun.star.chart2.template.StockVolumeOpenLowHighClose"  #  3 #
]

#=== バブルチャート
ChartTemplateBubble = [
  "com.sun.star.chart2.template.Bubble"  #  0 #
]


#== チャートの種類
#
# チャートの種類とテンプレートの関連付け
#
ChartTypeList = {
  "com.sun.star.chart.AreaDiagram"      => ChartTemplateArea,       # 表面
  "com.sun.star.chart.BarDiagram"       => ChartTemplateBar,        # 列
  "com.sun.star.chart.BubbleDiagram"    => ChartTemplateBubble,     # バブル
  "com.sun.star.chart.DonutDiagram"     => ChartTemplateDonut,      # 扇（ドーナツ）
  "com.sun.star.chart.FilledNetDiagram" => ChartTemplateFilledNet,  # レーダー網
  "com.sun.star.chart.LineDiagram"      => ChartTemplateLine,       # 線
  "com.sun.star.chart.NetDiagram"       => ChartTemplateNet,        # レーダー網
  "com.sun.star.chart.PieDiagram"       => ChartTemplatePie,        # 扇
  "com.sun.star.chart.StockDiagram"     => ChartTemplateStock,      # 株価
  "com.sun.star.chart.XYDiagram"        => ChartTemplateScatter     # 散布図
}

  #=== チャートの種類を取得
  # ret::文字列
  def get_chartType
    self.Diagram.getDiagramType
  end
  
  #=== データ列の範囲を表す文字列を取得
  #
  # _n_:: 何番目のデータ列かを指定、指定されない場合には 0。
  # ret::文字列
  def get_Range(n=0)
    self.DataSequences[n].Values.SourceRangeRepresentation
  end
  
  #=== X軸の範囲を変更する
  #
  # データ配列の範囲を指定された最小値と最大値へ変更する。
  # _min_index_:: 最小の行番号(1始まり)
  # _max_index_:: 最大の行番号(1始まり)
  # _chartTypeIndex_:: チャートテンプレート配列でのインデックス、指定されない場合には 0。
  # ret::成功のときtrue、失敗のときfalse
  #
  def change_Xrange(min_index,max_index,chartTypeIndex=0)
    min_index = min_index.to_i
    max_index = max_index.to_i
    tName = ChartTypeList[self.Diagram.getDiagramType][chartTypeIndex]
    ret = false
    if tName
      ret = true
      n = self.DataSequences.size
      prov = self.getDataProvider
      dt = []
      j = 0
      x_was_taken = false
      
      (1..n).each do |i|
        seq = self.DataSequences[i-1]
        role = seq.Values.Role
        range =  seq.Values.SourceRangeRepresentation
        range.sub!(/(.*\$)\d+(\:.*\$)\d+/,"\\1#{min_index}\\2#{max_index}")
        new_seq = prov.createDataSequenceByRangeRepresentation(range)
        new_seq.Role = role
        if role == "categories"
          labeldSeq = $manager.createInstance("com.sun.star.chart2.data.LabeledDataSequence")
          labeldSeq.setValues(new_seq)
          dt << labeldSeq
          x_was_taken = true
          j += 1
        elsif (role == "values-x")
          if ! x_was_taken
            dt[j] = $manager.createInstance("com.sun.star.chart2.data.LabeledDataSequence")
            dt[j].setValues(new_seq)

            if seq.Label
              label_role = seq.Label.Role
              label_range =  seq.Label.SourceRangeRepresentation
              new_label_seq = prov.createDataSequenceByRangeRepresentation(label_range)
              new_label_seq.Role = seq.Label.Role
              dt[j].setLabel(new_label_seq)
            end
            x_was_taken = true
            j += 1
#          else
#            print "ignore values-x\n"
          end
        else # (role == "values-y" or else )
          dt[j] = $manager.createInstance("com.sun.star.chart2.data.LabeledDataSequence")
          dt[j].setValues(new_seq)

          if seq.Label
            label_role = seq.Label.Role
            label_range =  seq.Label.SourceRangeRepresentation
            new_label_seq = prov.createDataSequenceByRangeRepresentation(label_range)
            new_label_seq.Role = seq.Label.Role
            dt[j].setLabel(new_label_seq)
          end
          j += 1
        end
      end
      source = self.getUsedData
      source.setData(dt)
      tManager = self.getChartTypeManager
      tTemplate = tManager.createInstance(tName)
      tTemplate.changeDiagramData(self.Diagram.getDiagram, source, [])
    end
    ret
  end
  
  #=== 画像タイプ定義
  #
  Graphic_Types = {
    'gif'=>'image/gif',
    'jpeg'=>'image/jpeg',
    'jpg'=>'image/jpeg',
    'png'=>'image/png',
    'fh'=>'image/x-freehand' ,
    'cgm'=>'image/cgm',
    'tiff'=>'image/tiff',
    'dxf'=>'image/vnd.dxf',
    'emf'=>'image/x-emf',
    'tga'=>'image/x-targa',
    'sgf'=>'image/x-sgf',
    'svm'=>'image/x-svm',
    'wmf'=>'image/x-wmf',
    'pict'=>'image/x-pict',
    'cmx'=>'image/x-cmx',
    'svg'=>'image/svg+xml',
    'bmp'=>'image/x-MS-bmp',
    'wpg'=>'image/x-wpg',
    'eps'=>'image/x-eps',
    'met'=>'image/x-met',
    'pbm'=>'image/x-portable-bitmap',
    'pcd'=>'image/x-photo-cd',
    'pcx'=>'image/x-pcx',
    'pgm'=>'image/x-portable-graymap',
    'ppm'=>'image/x-portable-pixmap',
    'psd'=>'image/vnd.adobe.photoshop',
    'ras'=>'image/x-cmu-raster',
    'ras'=>'image/x-sun-raster',
    'xbm'=>'image/x-xbitmap',
    'xpm'=>'image/x-xpixmap'
  }
  
  #=== チャートの保存
  #
  # _filename_::出力する画像ファイル名、拡張子で画像タイプを判定する。(bmp,jpg,png...)
  #
  def save(filename)
    done = false
    if filename != ''
      if filename !~ /^file\:\/\/\//
        filename = 'file:///'+getAbsolutePath(filename)
      end
      filename =~ /\.([^\/\.\\]+?)$/
      ext = $1
      ext.downcase! if ext
      t = Graphic_Types[ext]
    else
      print "画像のファイル名を指定してください\n"
      return false
    end
    if t == nil
      print "拡張子から画像タイプが判別できません。\n"
      return false
    end
    
    done = true
    begin
      f = $manager.createInstance("com.sun.star.drawing.GraphicExportFilter")
      f.setSourceDocument(self.getDrawPage)
      opt = _opts($manager,{'URL'=>filename,'MediaType'=>t})
      f.filter(opt)
    rescue 
      print "書き込みできませんでした。#{$!}\n"
      done = false
    end
    done
  end
  
end

#----------------------------------------------------
#=  シートドキュメント操作
#
# シートクラスの拡張モジュール
# シート取り出し時に自動的に組み込まれる。
module CalcWorksheet

  #=== チャートドキュメント取り出し
  #
  # _n_::何番目のチャートを取り出すかの指定(0始まり)、指定されない場合には 0。
  # ret::シートオブジェクト
  def get_chartdoc(n=0)
    charts = self.getCharts
    if n.class == String
      chart = charts.getByName(n)
    else
      chart = charts.getByIndex(n)
    end
    chartDoc = chart.EmbeddedObject
    chartDoc.extend(CalcChartDoc)
  end

  #=== セルの背景色を取得
  #
  # _y_::行番号(0始まり)
  # _x_::カラム番号(0始まり)
  # ret::RGB24bitでの色コード
  #
  def color(y,x)  #
      self.getCellByPosition(x,y).CellBackColor
  end

  #=== セルの背景色を設定
  #
  # _y_::行番号(0始まり)
  # _x_::カラム番号(0始まり)
  # _color_::RGB24bitでの色コード
  #
  def set_color(y,x,color)  #
      self.getCellByPosition(x,y).CellBackColor = color
  end

  #=== セル範囲の背景色を設定
  #
  # _y1_::行番号(0始まり)
  # _x1_::カラム番号(0始まり)
  # _y2_::行番号(0始まり)
  # _x2_::カラム番号(0始まり)
  # _color_::RGB24bitでの色コード
  #
  def set_range_color(y1,x1,y2,x2,color) #
    self.getCellRangeByPosition(x1,y1,x2,y2).CellBackColor = color
  end

  #=== カラム幅取得
  #
  # _x_::カラム番号(0始まり)
  # ret::幅(1/100mm単位)
  def get_width(x)  #
      self.Columns.getByIndex(x).Width
  end
  
  #=== カラム幅設定
  #
  # _x_::カラム番号(0始まり)
  # _width_::幅(1/100mm単位)
  def set_width(x,width)  #
      self.Columns.getByIndex(x).Width = width
  end
  
  #=== セルの値取り出し
  # sheet[行番号,カラム番号] でセルを参照する。
  #
  # _y_::行番号(0始まり)
  # _x_::カラム番号(0始まり)
  #
  def [] y,x    #
    cell = self.getCellByPosition(x,y)
    if cell.Type == 2 #CellCollectionType::TEXT
      cell.String
    else
      cell.Value
    end
  end

  #=== セルの値設定
  # sheet[行番号,カラム番号] でセルを参照する。
  #
  # _y_::行番号(0始まり)
  # _x_::カラム番号(0始まり)
  # _value_::設定値
  #
  def []= y,x,value   #
    cell = self.getCellByPosition(x,y)
    if value.class == String #CellCollectionType::TEXT
      cell.String = value
    else
      cell.Value= value
    end
  end

  $a2z = ('A'..'Z').to_a
  
  #=== 範囲指定の文字列作成
  #
  # _y_::行番号(0始まり)
  # _x_::カラム番号(0始まり)
  # ret::範囲指定文字列
  #
  def r_str(y,x)
    r = ''
    x -= 1
    if x > 26*26
      return "ZZ#{y}"
    else
      r = $a2z[((x/26)-1).to_i] if x > 25
      r += $a2z[(x%26).to_i]
      r += y.to_s
    end
    r
  end

  #=== セルの式を取得
  #
  # _y_::行番号(0始まり)
  # _x_::カラム番号(0始まり)
  # ret::式を表す文字列
  #
  def get_formula( y,x)  #
    cell = self.getCellByPosition(x,y)
    cell.Formula
  end

  #=== セルへ式を設定
  #
  # _y_::行番号(0始まり)
  # _x_::カラム番号(0始まり)
  # _f_::式を表す文字列
  #
  def set_formula( y,x,f)  #
    cell = self.getCellByPosition(x,y)
    cell.Formula = f
  end

  #=== 行をグループ化
  #
  # _y1_::開始行番号(0始まり)
  # _y2_::終了行番号(0始まり)
  #
  # Excelのグループ化と+-のアイコンの位置が異なります。
  def group_row(y1,y2)
    r = self.getCellRangeByPosition(0,y1,0,y2).RangeAddress
    self.group(r,1)
  end
  
  #=== 列をグループ化
  #
  # _y1_::開始列番号(0始まり)
  # _y2_::終了列番号(0始まり)
  #
  # Excelのグループ化と+-のアイコンの位置が異なります。
  def group_column(x1,x2)
    r = self.getCellRangeByPosition(x1,0,x2,0).RangeAddress
    self.group(r,0)
  end
  
  #=== セルのマージ設定
  #
  # _y1_::左上行番号(0始まり)
  # _x1_::左上カラム番号(0始まり)
  # _y2_::右下行番号(0始まり)
  # _x2_::右下カラム番号(0始まり)
  def merge(y1,x1,y2,x2)
    self.getCellRangeByPosition(x1,y1,x2,y2).merge(true)
  end

  #=== セルのマージ解除
  #
  # _y1_::左上行番号(0始まり)
  # _x1_::左上カラム番号(0始まり)
  # _y2_::右下行番号(0始まり)
  # _x2_::右下カラム番号(0始まり)
  def merge_off(y1,x1,y2,x2)
    self.getCellRangeByPosition(x1,y1,x2,y2).merge(false)
  end
  
  #=== 罫線枠設定
  #
  # _y1_::左上行番号(0始まり)
  # _x1_::左上カラム番号(0始まり)
  # _y2_::右下行番号(0始まり)
  # _x2_::右下カラム番号(0始まり)
  def box(y1,x1,y2,x2)
    r = self.getCellRangeByPosition(x1,y1,x2,y2)
    b = r.RightBorder
    b.InnerLineWidth = 10

    r.BottomBorder = b
    r.TopBorder = b
    r.LeftBorder = b
    r.RightBorder = b
  end


  #=== Wrap表示設定
  #
  # _y1_::左上行番号(0始まり)
  # _x1_::左上カラム番号(0始まり)
  # _y2_::右下行番号(0始まり)
  # _x2_::右下カラム番号(0始まり)
  def wrap(y1,x1,y2,x2)
    self.getCellRangeByPosition(x1,y1,x2,y2).IsTextWrapped = true
  end

  #=== 垂直方向の表示設定
  #
  # _y1_::左上行番号(0始まり)
  # _x1_::左上カラム番号(0始まり)
  # _y2_::右下行番号(0始まり)
  # _x2_::右下カラム番号(0始まり)
  # _v_:: 0:STANDARD,1:TOP,2:CENTER,3:BOTTOM
  #
  def verticals(y1,x1,y2,x2,v=0)
    self.getCellRangeByPosition(x1,y1,x2,y2).VertJustify  = v
  end

  def vertical(y1,x1,v=0)
    self.getCellByPosition(x1,y1).VertJustify  = v
  end
  
  
  #=== 上付き表示設定
  #
  # _y1_::左上行番号(0始まり)
  # _x1_::左上カラム番号(0始まり)
  # _y2_::右下行番号(0始まり)
  # _x2_::右下カラム番号(0始まり)
  #
  def v_tops(y1,x1,y2,x2)
    self.verticals(y1,x1,y2,x2,1)
  end

  def v_top(y1,x1)
    self.vertical(y1,x1,1)
  end

  #=== 水平方向の表示設定(Range)
  #
  # _y1_::左上行番号(0始まり)
  # _x1_::左上カラム番号(0始まり)
  # _y2_::右下行番号(0始まり)
  # _x2_::右下カラム番号(0始まり)
  # _h_:: 0:STANDARD,1:LEFT,2:CENTER,3:RIGHT,4:BLOCK,5:REPEAT
  #
  def horizontals(y1,x1,y2,x2,h=0)
    self.getCellRangeByPosition(x1,y1,x2,y2).HoriJustify  = h
  end

  #=== 水平方向の表示設定(Cell)
  #
  # _y1_::左上行番号(0始まり)
  # _x1_::左上カラム番号(0始まり)
  # _h_:: 0:STANDARD,1:LEFT,2:CENTER,3:RIGHT,4:BLOCK,5:REPEAT
  #
  def horizontal(y1,x1,h=0)
    self.getCellByPosition(x1,y1).HoriJustify  = h
  end

  #=== センター表示設定(Range)
  #
  # _y1_::左上行番号(0始まり)
  # _x1_::左上カラム番号(0始まり)
  # _y2_::右下行番号(0始まり)
  # _x2_::右下カラム番号(0始まり)
  def centers(y1,x1,y2,x2)
    self.horizontals(y1,x1,y2,x2,2)
  end

  #=== センター表示設定(Cell)
  #
  # _y1_::左上行番号(0始まり)
  # _x1_::左上カラム番号(0始まり)
  def center(y1,x1)
    self.horizontal(y1,x1,2)
  end

  #=== 書式コピー(Range)
  #
  # _sy_::コピー元 行番号(0始まり)
  # _sx_::コピー元 カラム番号(0始まり)
  # _ty1_::コピー先 左上行番号(0始まり)
  # _tx1_::コピー先 左上カラム番号(0始まり)
  # _ty2_::コピー先 右下行番号(0始まり)
  # _tx2_::コピー先 右下カラム番号(0始まり)
  def format_copy(sy,sx,ty1,tx1,ty2,tx2)
    s = self.getCellByPosition(sx,sy)
    sp = s.getPropertySetInfo.getProperties
    names = sp.each.map{|p| p.Name}
    ps = s.getPropertyValues(names)
    self.getCellRangeByPosition(tx1,ty1,tx2,ty2).setPropertyValues(names,ps)
  end

  #=== 書式コピー2(Range)
  #
  # コピー元行の書式をコピー先にn行コピーする。
  # コピー先はコピー元と同じカラム数とする。
  #
  # _sy1_::コピー元 行番号(0始まり)
  # _sx1_::コピー元 開始カラム番号(0始まり)
  # _sx2_::コピー元 終了カラム番号(0始まり)
  # _ty_::コピー先 左上行番号(0始まり)
  # _tx_::コピー先 左上カラム番号(0始まり)
  # _n_:: コピー行数
  def format_range_copy(sy,sx1,sx2,  ty,tx,n=1)
    return if n < 1
    (sx1..sx2).each do |x|
      self.format_copy(sy,x,  ty,tx+(x-sx1),ty+n-1,tx+(x-sx1))
    end
  end

  #=== 書式コピー(Cell)
  #
  # _sy_::コピー元 行番号(0始まり)
  # _sx_::コピー元 カラム番号(0始まり)
  # _ty_::コピー先 左上行番号(0始まり)
  # _tx_::コピー先 左上カラム番号(0始まり)
  def format_copy1(sy,sx,ty,tx)
    s = self.getCellByPosition(sx,sy)
    sp = s.getPropertySetInfo.getProperties
    names = sp.each.map{|p| p.Name}
    ps = s.getPropertyValues(names)
    self.getCellByPosition(tx,ty).setPropertyValues(names,ps)
  end

  #=== コピー(Range)
  #
  # _sy1_::コピー元 左上行番号(0始まり)
  # _sx1_::コピー元 左上カラム番号(0始まり)
  # _sy2_::コピー元 右下行番号(0始まり)
  # _sx2_::コピー元 右下カラム番号(0始まり)
  # _ty_::コピー先 左上行番号(0始まり)
  # _tx_::コピー先 左上カラム番号(0始まり)
  def copy(sy1,sx1,sy2,sx2,ty,tx)
    r = self.getCellRangeByPosition(sx1,sy1,sx2,sy2).getRangeAddress
    c = self.getCellByPosition(tx,ty).getCellAddress
    self.copyRange(c,r)
  end


  #=== 行の挿入
  #
  # _n_:: 行番号(0始まり)、この行の前に挿入する。
  # _count_:: 何行挿入するかの指定、指定しない場合には1。
  #
  def insert_rows(n,count=1)   #
    self.Rows.insertByIndex(n,count)
  end

  #=== 行の削除
  #
  # _n_:: 行番号(0始まり)、この行から下を削除する。
  # _count_:: 何行削除するかの指定、指定しない場合には1。
  #
  def remove_rows(n,count=1)   #
    r = self.getCellRangeByPosition(0,n,0,n+count-1).getRangeAddress
    self.removerange(r,3)
  end


end


#----------------------------------------------------
#== OpenOfficeドキュメント操作
#
# Calcドキュメントの拡張モジュール。
# ドキュメント取り出し時に自動的に組み込まれる。
module OOoDocument

  #=== シートの取り出し
  #
  # _s_:: シート名文字列またはシートのインデックス番号
  # ret:: シートオブジェクト
  #
  def get_sheet(s)
    if s.class == String
      sheet = self.sheets.getByName(s)
    else
      sheet = self.sheets.getByIndex(s)
    end
    sheet.extend(CalcWorksheet)
  end

  #=== Activeシートの切り替え
  #
  # _s_:: シート名文字列またはシートのインデックス番号
  #
  def set_active_sheet(s)
    if s.class == String
      self.getCurrentController.setActiveSheet(self.Sheets.getByName(s))
    else
      self.getCurrentController.setActiveSheet(self.Sheets.getByIndex(s))
    end
  end

  #=== Calcドキュメントの保存(書き出し)
  #
  # _filename_:: 名前を変えて保存する場合にファイル名を指定する。
  # ret::成功したらtrue、失敗したらfalse
  #
  def save(filename=nil)
    done = true
    begin
      if filename
        if filename !~ /^file\:\/\/\//
          filename = 'file:///'+getAbsolutePath(filename)
        end
        
        options = []
        options += get_filtername_option(filename)
        self.storeToURL(filename,options)
        sleep(2)
      else
        self.store()
        sleep(2)
      end
    rescue 
      print "書き込みできませんでした。\n"
      done = false
    end
    done
  end
end


#----------------------------------------------------
#== Calcドキュメント操作

#=== オプション指定用配列の作成
#
# _manager_:: com.sun.star.ServiceManager
# _hash_:: オプション指定(名前と値の連想配列)
# ret::オプション指定用の配列
#
def _opts(manager,hash)
  hash.inject(opts = []) {|x,y|
    opt = manager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
    opt[0].Name = y[0]
    opt[0].Value = y[1]
    x << opt
  }
  opts
end

Document_Types = {
    "csv"  =>"Text - txt - csv (StarCalc)",
    "fods" =>"OpenDocument Spreadsheet Flat XML",
    "html" =>"HTML (StarCalc)",
    "ods"  =>"calc8",
    "ots"  =>"calc8_template",
    "pdf"  =>"calc_pdf_Export",
    "xhtml"=>"XHTML Calc File",
    "xls"  =>"MS Excel 97",
    "xlsx" =>"Calc MS Excel 2007 XML"
}

#=== ドキュメント形式名取り出し
#
#  拡張子がodsなら空の配列を返す。
#  その他なら対応するフィルター名のプロパティ配列を返す。
# =filename_::ファイル名
# ret::フィルターオプション(ドキュメント形式名) 
#
def get_filtername_option( filename)
  filename =~ /\.([^\/\.\\]+?)$/
  ext = $1
  return [] if (ext == nil) or (ext == '')
  ext.downcase!

  return [] if ext == 'ods'
  t = Document_Types[ext]
  return [] if !t
  return _opts($manager,{"FilterName"=>t})

end

#----------------------------------------------------
#=== 既存ドキュメントのOpen
#
#  処理ブロックを受け取って実行する。
#
# _filename:: Calcドキュメントのファイル名
# _visible_:: Calcのウインドウを表示するときtrue、指定されない場合にはtrue
#
def openCalcWorkbook filename, visible=true
  if filename !~ /^file\:\/\/\//
    filename = 'file:///'+getAbsolutePath(filename)
  end
  manager = WIN32OLE.new("com.sun.star.ServiceManager")
  $manager = manager
  desktop = manager.createInstance("com.sun.star.frame.Desktop")
  
  options = []
#  options += get_filtername_option(filename)
  options += _opts(manager,{"Hidden" => true}) if !visible
  
  book     = desktop.loadComponentFromURL(filename, "_blank", 0,options)

  book.extend(OOoDocument)

  begin
    yield book
  ensure
    book.close(false)
  end
##  desktop.terminate   ## 開いている他のOOoドキュメントも道ずれにしてすべて終了してしまう
  
end


#----------------------------------------------------
#=== 新規ドキュメントの作成
#
#  処理ブロックを受け取って実行する。
#  saveを呼びだすときにファイル名を指定してを保存する。
#
# _visible_:: Calcのウインドウを表示するときtrue、指定されない場合にはtrue
#
def createCalcWorkbook visible=true
  manager = WIN32OLE.new("com.sun.star.ServiceManager")
  $manager = manager
  desktop = manager.createInstance("com.sun.star.frame.Desktop")
  if visible
    book    = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, [])
  else
    book    = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, _opts(manager,{"Hidden" => true}))
  end
  book.extend(OOoDocument)
  
  begin
    yield book
  ensure
    book.close(false)
  end
##  desktop.terminate   ## 開いている他のOOoドキュメントも道ずれにしてすべて終了してしまう

end


if __FILE__ == $0

  print "OpenOffice.org Calc用 ruby拡張モジュール\n"
  print "\nrubyのRDocでドキュメントを作成してください。\n"
  print "rdoc --title Trail_Calc --main Trail_calc  --all Trail_calc.rb\n"

end
