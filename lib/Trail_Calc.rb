#! ruby -EWindows-31J
# -*- mode:ruby; coding:Windows-31J -*-
 
require "Trail_Calc/version"
require 'win32ole'
require 'date'

#= Trail_Calc
#== OpenOffice.org Calc�p ruby�g�����W���[��
#
#*  OpenOffice.org(LibreOffice)��Calc��ruby���瑀�삷�邽�߂�ruby�g�����W���[���ł��B
#
#Authors:: Mt.Trail
#Version:: Ver. 1.03 2016/08/02 Mt.Trail
#* Ver. 1.03  : gem��
#* Ver. 1.02  : excel�`�����ł̕ۑ��@�\��ǉ��A�݊����ɂ͖�肠��A�`���[�g�͈͕̔ύX��bugfix
#* ver. 1.01  : �`���[�g�̉摜�o�͋@�\�ǉ�
#* Ver. 1.00  : ������J��
#
#Copyright:: Copyright (C) Mt.Trail 2015 All rights reserved.
#License:: GPL version 2
#
#
#==�ړI
#* ruby����OpenOffice.org��Calc���ȒP�ɑ��삷�邱�ƁB
# 
#* Excel�̖������ŕ\��O���t�������X�V���邽�߂ɍ쐬���܂����B
#* ��{�I�ɂ̓V�[�g�̃��C�A�E�g��`���[�g�̌`������Calc�ō쐬���f�[�^�̒ǉ���ύX������������ړI�ō쐬���Ă��܂��B
#* �ł�����A�ڍׂȃv���p�e�B�̐ݒ��A�V�����`���[�g�̍쐬���̓T�|�[�g���Ă��܂���B
#* ������ݒ肵���e���v���[�g�̈��p�ӂ��Ă��������̃R�s�[�ŏ�����ݒ肵�Ă��������B
#
#= ���p���@
#*  ruby�X�N���v�g��require���ė��p���Ă��������B
#
#== �g�p��
#     require 'trail_calc'
#     
#     openCalcWorkbook("file.ods") do |book|
#       sheet = book.get_sheet("DataSheet")
#       sheet[0,3] = 1
#       ...
#       book.save
#     end
#
#==�h�L�������g����
#
#  rdoc --title Trail_Calc --main Trail_calc  --all Trail_calc.rb

module TrailCalc
  # Dumy module for RDOC
end


#----------------------------------------------------
#== �ėp�֐�
#=== �F�R�[�h����
# �ԁA�΁A��8bit�l(0����255)�Ŏw�肵��24bit�̐F�R�[�h���쐬���܂��B
#
# _red_   :: ��(0-255)
# _green_ :: ��(0-255)
# _blue_  :: ��(0-255)
#
def rgb(red, green, blue)
  return red << 16 | green << 8 | blue
end

#=== OpenOffice.org�̎�������ruby�̎����֕ϊ�
#
# _t_:: OOo�̓��t������\��double�̐��l
#
def time_ooo2ruby(t)  ## t ��OOo�̓��t������\��double�̐��l
  #OOo�̊�̓f�t�H���g����1899/12/30��
  #Apple��1904/1/1, StartOffice��1900/1/1
  #ruby��1970//1/1 9:0:0 (JST������+9����)
  diff = -2209194000  # Time.at(Time.new(1899,12,30)).to_i
  Time.at(t * 24 * 60 * 60 + diff +0.1) ## 0.1�͔����ȕϊ��덷�̕␳�p
end

#=== ruby�̎�������OpenOffice.org�̎����֕ϊ�
#
# _t_:: ruby��Time�I�u�W�F�N�g
#
def time_ruby2ooo(t)  ## t ruby��Time�I�u�W�F�N�g
  diff = -2209194000  # Time.at(Time.new(1899,12,30)).to_i
  d = (t.to_i - diff)/(24.0 * 60.0 * 60.0)
end

#=== ��΃p�X�̎擾
#
# _filename_ :: �t�@�C����
#
def getAbsolutePath filename
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  fn = fso.GetAbsolutePathName(filename)
  fn.gsub!(/\\/,'/')
  fn
end


#----------------------------------------------------
#==  �`���[�g�h�L�������g����
#
# �`���[�g�h�L�������g�N���X�̊g�����W���[��
# �`���[�g�h�L�������g���o�����Ɏ����I�ɑg�ݍ��܂��B

module CalcChartDoc

  #=== �`���[�g�̃^�C�g���擾
  #
  # ret::������
  def get_title
    self.getTitle.String
  end
  
  #=== �`���[�g�̃^�C�g���ݒ�
  #
  def set_title(name)
    t = self.getTitle
    t.String = name
  end

  #=== �`���[�g�̃T�u�^�C�g���擾
  #
  # ret::������
  def get_subtitle
    self.getSubTitle.String
  end

  #=== �`���[�g�̃T�u�^�C�g���ݒ�
  #
  def set_subtitle(name)
    t = self.getSubTitle
    t.String = name
  end

  #=== �ύX���ꂽ�����ׂ�
  # OOo�W��
  #
  # isModified()
  # Modified
  # ret::Bool
  
  #=== �ۑ�
  # OOo�W��
  #
  # store()

  #=== X���̍ŏ��l���擾
  #
  # ret::double
  def get_Xmin
    self.Diagram.XAxis.Min
  end

  #=== X���̍ő�l���擾
  #
  # ret::double
  def get_Xmax
    self.Diagram.XAxis.Max
  end

  #=== X���̍ŏ��l��ݒ�
  #
  # _t_:: OOo�ł�double�̒l�A�����̏ꍇ��double�ŕ\�������B
  def set_Xmin(t)  ## t ��OOo�ł̒l
    self.Diagram.XAxis.Min = t
  end

  #=== X���̍ő�l��ݒ�
  #
  # _t_:: OOo�ł�double�̒l�A�����̏ꍇ��double�ŕ\�������B
  def set_Xmax(t)  ## t ��OOo�ł̒l
    self.Diagram.XAxis.Max = t
  end


#== �`���[�g�̃e���v���[�g
#
# �`���[�g�̃e���v���[�g�̎�ނ�\��������̔z��B
# �擪���f�t�H���g�l�Ƃ���B

#=== �_�O���t
#  Role
#    categories
#    values-y
ChartTemplateBar = [
  "com.sun.star.chart2.template.Column",                          #  0 # �c
  "com.sun.star.chart2.template.StackedColumn",                   #  1 # �c�ςݏグ
  "com.sun.star.chart2.template.PercentStackedColumn",            #  2 # �c�ςݏグ�p�[�Z���g
  "com.sun.star.chart2.template.ThreeDColumnDeep",                #  3 # 3D  �c���s������
  "com.sun.star.chart2.template.ThreeDColumnFlat",                #  4 # 3D  �c���s���Ȃ�
  "com.sun.star.chart2.template.PercentStackedThreeDColumnFlat",  #  5 # 3D  �c�ςݏグ
  "com.sun.star.chart2.template.StackedThreeDColumnFlat",         #  6 # 3D  �c�ςݏグ�p�[�Z���g
  "com.sun.star.chart2.template.Bar",                             #  7 # ��
  "com.sun.star.chart2.template.StackedBar",                      #  8 # ���ςݏグ
  "com.sun.star.chart2.template.PercentStackedBar",               #  9 # ���ςݏグ�p�[�Z���g
  "com.sun.star.chart2.template.ThreeDBarDeep",                   # 10 # 3D  �����s������
  "com.sun.star.chart2.template.ThreeDBarFlat",                   # 11 # 3D  �����s���Ȃ�
  "com.sun.star.chart2.template.StackedThreeDBarFlat",            # 12 # 3D  ���ςݏグ
  "com.sun.star.chart2.template.PercentStackedThreeDBarFlat"      # 13 # 3D  ���ςݏグ�p�[�Z���g
]
#
#=== �~�O���t
#  Role
#    categories
#    values-y
ChartTemplatePie = [
  "com.sun.star.chart2.template.Pie",                      #  0 # ��^
  "com.sun.star.chart2.template.PieAllExploded",           #  1 # ��^����
  "com.sun.star.chart2.template.ThreeDPie",                #  2 # 3D ��^
  "com.sun.star.chart2.template.ThreeDPieAllExploded"      #  3 # 3D ��^����
]
#=== �h�[�i�c�O���t
ChartTemplateDonut = [
  "com.sun.star.chart2.template.Donut",                    #  0 # �h�[�i�c
  "com.sun.star.chart2.template.DonutAllExploded",         #  1 # �h�[�i�c����
  "com.sun.star.chart2.template.ThreeDDonut",              #  2 # 3D �h�[�i�c
  "com.sun.star.chart2.template.ThreeDDonutAllExploded"    #  3 # 3D �h�[�i�c����
]
#
#=== �G���A�O���t
#  Role
#    categories
#    values-y
ChartTemplateArea = [
  "com.sun.star.chart2.template.Area",                     #  0 #   �G���A
  "com.sun.star.chart2.template.StackedArea",              #  1 #   �ςݏグ
  "com.sun.star.chart2.template.ThreeDArea",               #  2 #   3D
  "com.sun.star.chart2.template.StackedThreeDArea",        #  3 #   3D �ςݏグ
  "com.sun.star.chart2.template.PercentStackedArea",       #  4 #   �ςݏグ�p�[�Z���g
  "com.sun.star.chart2.template.PercentStackedThreeDArea"  #  5 #   3D �ςݏグ�p�[�Z���g
]
#
#=== �܂��
ChartTemplateLine = [
  "com.sun.star.chart2.template.Line",                     #  0 #   ��
  "com.sun.star.chart2.template.Symbol",                   #  1 #   �_
  "com.sun.star.chart2.template.LineSymbol",               #  2 #   �_�Ɛ�
  "com.sun.star.chart2.template.ThreeDLine",               #  3 #   3D ��
  "com.sun.star.chart2.template.ThreeDLineDeep",           #  4 #   3D �����s������
  "com.sun.star.chart2.template.StackedLine",              #  6 #   ���ςݏグ
  "com.sun.star.chart2.template.StackedSymbol",            #  5 #   �_�ςݏグ
  "com.sun.star.chart2.template.StackedLineSymbol",        #  7 #   �_�Ɛ��ςݏグ
  "com.sun.star.chart2.template.StackedThreeDLine",        #  8 #   3D ���ςݏグ
  "com.sun.star.chart2.template.PercentStackedSymbol",     #  9 #   �_�ςݏグ�p�[�Z���g
  "com.sun.star.chart2.template.PercentStackedLine",       # 10 #   ���ςݏグ�p�[�Z���g
  "com.sun.star.chart2.template.PercentStackedLineSymbol", # 11 #   �_�Ɛ��ςݏグ�p�[�Z���g
  "com.sun.star.chart2.template.PercentStackedThreeDLine"  # 12 #   3D ���ςݏグ�p�[�Z���g
]
#
#=== �U�z�}
#  Role
#    categories: �n��̃��x��
#    values-x : X �f�[�^
#    values-y: Y �f�[�^
ChartTemplateScatter = [
  "com.sun.star.chart2.template.ScatterLine",         #  0 # ���C���̂�
  "com.sun.star.chart2.template.ScatterLineSymbol",   #  1 # ���C���ƃf�[�^�_
  "com.sun.star.chart2.template.ScatterSymbol",       #  2 # �f�[�^�_
  "com.sun.star.chart2.template.ThreeDScatter"        #  3 # 3D
]
#
#=== ���[�_�[��
#  Role
#    categories
#    values-y
ChartTemplateNet = [
  "com.sun.star.chart2.template.NetLine",                #  0 # ��
  "com.sun.star.chart2.template.Net",                    #  1 # �_�Ɛ�
  "com.sun.star.chart2.template.NetSymbol",              #  2 # �_
  "com.sun.star.chart2.template.StackedNet",             #  3 # �ςݏグ�_�Ɛ�
  "com.sun.star.chart2.template.StackedNetLine",         #  4 # �ςݏグ��
  "com.sun.star.chart2.template.StackedNetSymbol",       #  5 # �ςݏグ�_
  "com.sun.star.chart2.template.PercentStackedNet",      #  6 # �_�Ɛ��ςݏグ�p�[�Z���g
  "com.sun.star.chart2.template.PercentStackedNetLine",  #  7 # ���ςݏグ�p�[�Z���g
  "com.sun.star.chart2.template.PercentStackedNetSymbol" #  8 # �_�ςݏグ�p�[�Z���g
]

#=== ���[�_�[��(�h��Ԃ�)
ChartTemplateFilledNet = [
  "com.sun.star.chart2.template.FilledNet",              #  0 # 3.2
  "com.sun.star.chart2.template.StackedFilledNet",       #  1 # 3.2
  "com.sun.star.chart2.template.PercentStackedFilledNet" #  2 # 3.2
]

#=== �X�g�b�N�`���[�g
ChartTemplateStock = [
  "com.sun.star.chart2.template.StockLowHighClose",           #  0 #
  "com.sun.star.chart2.template.StockOpenLowHighClose",       #  1 #
  "com.sun.star.chart2.template.StockVolumeLowHighClose",     #  2 #
  "com.sun.star.chart2.template.StockVolumeOpenLowHighClose"  #  3 #
]

#=== �o�u���`���[�g
ChartTemplateBubble = [
  "com.sun.star.chart2.template.Bubble"  #  0 #
]


#== �`���[�g�̎��
#
# �`���[�g�̎�ނƃe���v���[�g�̊֘A�t��
#
ChartTypeList = {
  "com.sun.star.chart.AreaDiagram"      => ChartTemplateArea,       # �\��
  "com.sun.star.chart.BarDiagram"       => ChartTemplateBar,        # ��
  "com.sun.star.chart.BubbleDiagram"    => ChartTemplateBubble,     # �o�u��
  "com.sun.star.chart.DonutDiagram"     => ChartTemplateDonut,      # ��i�h�[�i�c�j
  "com.sun.star.chart.FilledNetDiagram" => ChartTemplateFilledNet,  # ���[�_�[��
  "com.sun.star.chart.LineDiagram"      => ChartTemplateLine,       # ��
  "com.sun.star.chart.NetDiagram"       => ChartTemplateNet,        # ���[�_�[��
  "com.sun.star.chart.PieDiagram"       => ChartTemplatePie,        # ��
  "com.sun.star.chart.StockDiagram"     => ChartTemplateStock,      # ����
  "com.sun.star.chart.XYDiagram"        => ChartTemplateScatter     # �U�z�}
}

  #=== �`���[�g�̎�ނ��擾
  # ret::������
  def get_chartType
    self.Diagram.getDiagramType
  end
  
  #=== �f�[�^��͈̔͂�\����������擾
  #
  # _n_:: ���Ԗڂ̃f�[�^�񂩂��w��A�w�肳��Ȃ��ꍇ�ɂ� 0�B
  # ret::������
  def get_Range(n=0)
    self.DataSequences[n].Values.SourceRangeRepresentation
  end
  
  #=== X���͈̔͂�ύX����
  #
  # �f�[�^�z��͈̔͂��w�肳�ꂽ�ŏ��l�ƍő�l�֕ύX����B
  # _min_index_:: �ŏ��̍s�ԍ�(1�n�܂�)
  # _max_index_:: �ő�̍s�ԍ�(1�n�܂�)
  # _chartTypeIndex_:: �`���[�g�e���v���[�g�z��ł̃C���f�b�N�X�A�w�肳��Ȃ��ꍇ�ɂ� 0�B
  # ret::�����̂Ƃ�true�A���s�̂Ƃ�false
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
  
  #=== �摜�^�C�v��`
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
  
  #=== �`���[�g�̕ۑ�
  #
  # _filename_::�o�͂���摜�t�@�C�����A�g���q�ŉ摜�^�C�v�𔻒肷��B(bmp,jpg,png...)
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
      print "�摜�̃t�@�C�������w�肵�Ă�������\n"
      return false
    end
    if t == nil
      print "�g���q����摜�^�C�v�����ʂł��܂���B\n"
      return false
    end
    
    done = true
    begin
      f = $manager.createInstance("com.sun.star.drawing.GraphicExportFilter")
      f.setSourceDocument(self.getDrawPage)
      opt = _opts($manager,{'URL'=>filename,'MediaType'=>t})
      f.filter(opt)
    rescue 
      print "�������݂ł��܂���ł����B#{$!}\n"
      done = false
    end
    done
  end
  
end

#----------------------------------------------------
#=  �V�[�g�h�L�������g����
#
# �V�[�g�N���X�̊g�����W���[��
# �V�[�g���o�����Ɏ����I�ɑg�ݍ��܂��B
module CalcWorksheet

  #=== �`���[�g�h�L�������g���o��
  #
  # _n_::���Ԗڂ̃`���[�g�����o�����̎w��(0�n�܂�)�A�w�肳��Ȃ��ꍇ�ɂ� 0�B
  # ret::�V�[�g�I�u�W�F�N�g
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

  #=== �Z���̔w�i�F���擾
  #
  # _y_::�s�ԍ�(0�n�܂�)
  # _x_::�J�����ԍ�(0�n�܂�)
  # ret::RGB24bit�ł̐F�R�[�h
  #
  def color(y,x)  #
      self.getCellByPosition(x,y).CellBackColor
  end

  #=== �Z���̔w�i�F��ݒ�
  #
  # _y_::�s�ԍ�(0�n�܂�)
  # _x_::�J�����ԍ�(0�n�܂�)
  # _color_::RGB24bit�ł̐F�R�[�h
  #
  def set_color(y,x,color)  #
      self.getCellByPosition(x,y).CellBackColor = color
  end

  #=== �Z���͈͂̔w�i�F��ݒ�
  #
  # _y1_::�s�ԍ�(0�n�܂�)
  # _x1_::�J�����ԍ�(0�n�܂�)
  # _y2_::�s�ԍ�(0�n�܂�)
  # _x2_::�J�����ԍ�(0�n�܂�)
  # _color_::RGB24bit�ł̐F�R�[�h
  #
  def set_range_color(y1,x1,y2,x2,color) #
    self.getCellRangeByPosition(x1,y1,x2,y2).CellBackColor = color
  end

  #=== �J�������擾
  #
  # _x_::�J�����ԍ�(0�n�܂�)
  # ret::��(1/100mm�P��)
  def get_width(x)  #
      self.Columns.getByIndex(x).Width
  end
  
  #=== �J�������ݒ�
  #
  # _x_::�J�����ԍ�(0�n�܂�)
  # _width_::��(1/100mm�P��)
  def set_width(x,width)  #
      self.Columns.getByIndex(x).Width = width
  end
  
  #=== �Z���̒l���o��
  # sheet[�s�ԍ�,�J�����ԍ�] �ŃZ�����Q�Ƃ���B
  #
  # _y_::�s�ԍ�(0�n�܂�)
  # _x_::�J�����ԍ�(0�n�܂�)
  #
  def [] y,x    #
    cell = self.getCellByPosition(x,y)
    if cell.Type == 2 #CellCollectionType::TEXT
      cell.String
    else
      cell.Value
    end
  end

  #=== �Z���̒l�ݒ�
  # sheet[�s�ԍ�,�J�����ԍ�] �ŃZ�����Q�Ƃ���B
  #
  # _y_::�s�ԍ�(0�n�܂�)
  # _x_::�J�����ԍ�(0�n�܂�)
  # _value_::�ݒ�l
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
  
  #=== �͈͎w��̕�����쐬
  #
  # _y_::�s�ԍ�(0�n�܂�)
  # _x_::�J�����ԍ�(0�n�܂�)
  # ret::�͈͎w�蕶����
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

  #=== �Z���̎����擾
  #
  # _y_::�s�ԍ�(0�n�܂�)
  # _x_::�J�����ԍ�(0�n�܂�)
  # ret::����\��������
  #
  def get_formula( y,x)  #
    cell = self.getCellByPosition(x,y)
    cell.Formula
  end

  #=== �Z���֎���ݒ�
  #
  # _y_::�s�ԍ�(0�n�܂�)
  # _x_::�J�����ԍ�(0�n�܂�)
  # _f_::����\��������
  #
  def set_formula( y,x,f)  #
    cell = self.getCellByPosition(x,y)
    cell.Formula = f
  end

  #=== �s���O���[�v��
  #
  # _y1_::�J�n�s�ԍ�(0�n�܂�)
  # _y2_::�I���s�ԍ�(0�n�܂�)
  #
  # Excel�̃O���[�v����+-�̃A�C�R���̈ʒu���قȂ�܂��B
  def group_row(y1,y2)
    r = self.getCellRangeByPosition(0,y1,0,y2).RangeAddress
    self.group(r,1)
  end
  
  #=== ����O���[�v��
  #
  # _y1_::�J�n��ԍ�(0�n�܂�)
  # _y2_::�I����ԍ�(0�n�܂�)
  #
  # Excel�̃O���[�v����+-�̃A�C�R���̈ʒu���قȂ�܂��B
  def group_column(x1,x2)
    r = self.getCellRangeByPosition(x1,0,x2,0).RangeAddress
    self.group(r,0)
  end
  
  #=== �Z���̃}�[�W�ݒ�
  #
  # _y1_::����s�ԍ�(0�n�܂�)
  # _x1_::����J�����ԍ�(0�n�܂�)
  # _y2_::�E���s�ԍ�(0�n�܂�)
  # _x2_::�E���J�����ԍ�(0�n�܂�)
  def merge(y1,x1,y2,x2)
    self.getCellRangeByPosition(x1,y1,x2,y2).merge(true)
  end

  #=== �Z���̃}�[�W����
  #
  # _y1_::����s�ԍ�(0�n�܂�)
  # _x1_::����J�����ԍ�(0�n�܂�)
  # _y2_::�E���s�ԍ�(0�n�܂�)
  # _x2_::�E���J�����ԍ�(0�n�܂�)
  def merge_off(y1,x1,y2,x2)
    self.getCellRangeByPosition(x1,y1,x2,y2).merge(false)
  end
  
  #=== �r���g�ݒ�
  #
  # _y1_::����s�ԍ�(0�n�܂�)
  # _x1_::����J�����ԍ�(0�n�܂�)
  # _y2_::�E���s�ԍ�(0�n�܂�)
  # _x2_::�E���J�����ԍ�(0�n�܂�)
  def box(y1,x1,y2,x2)
    r = self.getCellRangeByPosition(x1,y1,x2,y2)
    b = r.RightBorder
    b.InnerLineWidth = 10

    r.BottomBorder = b
    r.TopBorder = b
    r.LeftBorder = b
    r.RightBorder = b
  end


  #=== Wrap�\���ݒ�
  #
  # _y1_::����s�ԍ�(0�n�܂�)
  # _x1_::����J�����ԍ�(0�n�܂�)
  # _y2_::�E���s�ԍ�(0�n�܂�)
  # _x2_::�E���J�����ԍ�(0�n�܂�)
  def wrap(y1,x1,y2,x2)
    self.getCellRangeByPosition(x1,y1,x2,y2).IsTextWrapped = true
  end

  #=== ���������̕\���ݒ�
  #
  # _y1_::����s�ԍ�(0�n�܂�)
  # _x1_::����J�����ԍ�(0�n�܂�)
  # _y2_::�E���s�ԍ�(0�n�܂�)
  # _x2_::�E���J�����ԍ�(0�n�܂�)
  # _v_:: 0:STANDARD,1:TOP,2:CENTER,3:BOTTOM
  #
  def verticals(y1,x1,y2,x2,v=0)
    self.getCellRangeByPosition(x1,y1,x2,y2).VertJustify  = v
  end

  def vertical(y1,x1,v=0)
    self.getCellByPosition(x1,y1).VertJustify  = v
  end
  
  
  #=== ��t���\���ݒ�
  #
  # _y1_::����s�ԍ�(0�n�܂�)
  # _x1_::����J�����ԍ�(0�n�܂�)
  # _y2_::�E���s�ԍ�(0�n�܂�)
  # _x2_::�E���J�����ԍ�(0�n�܂�)
  #
  def v_tops(y1,x1,y2,x2)
    self.verticals(y1,x1,y2,x2,1)
  end

  def v_top(y1,x1)
    self.vertical(y1,x1,1)
  end

  #=== ���������̕\���ݒ�(Range)
  #
  # _y1_::����s�ԍ�(0�n�܂�)
  # _x1_::����J�����ԍ�(0�n�܂�)
  # _y2_::�E���s�ԍ�(0�n�܂�)
  # _x2_::�E���J�����ԍ�(0�n�܂�)
  # _h_:: 0:STANDARD,1:LEFT,2:CENTER,3:RIGHT,4:BLOCK,5:REPEAT
  #
  def horizontals(y1,x1,y2,x2,h=0)
    self.getCellRangeByPosition(x1,y1,x2,y2).HoriJustify  = h
  end

  #=== ���������̕\���ݒ�(Cell)
  #
  # _y1_::����s�ԍ�(0�n�܂�)
  # _x1_::����J�����ԍ�(0�n�܂�)
  # _h_:: 0:STANDARD,1:LEFT,2:CENTER,3:RIGHT,4:BLOCK,5:REPEAT
  #
  def horizontal(y1,x1,h=0)
    self.getCellByPosition(x1,y1).HoriJustify  = h
  end

  #=== �Z���^�[�\���ݒ�(Range)
  #
  # _y1_::����s�ԍ�(0�n�܂�)
  # _x1_::����J�����ԍ�(0�n�܂�)
  # _y2_::�E���s�ԍ�(0�n�܂�)
  # _x2_::�E���J�����ԍ�(0�n�܂�)
  def centers(y1,x1,y2,x2)
    self.horizontals(y1,x1,y2,x2,2)
  end

  #=== �Z���^�[�\���ݒ�(Cell)
  #
  # _y1_::����s�ԍ�(0�n�܂�)
  # _x1_::����J�����ԍ�(0�n�܂�)
  def center(y1,x1)
    self.horizontal(y1,x1,2)
  end

  #=== �����R�s�[(Range)
  #
  # _sy_::�R�s�[�� �s�ԍ�(0�n�܂�)
  # _sx_::�R�s�[�� �J�����ԍ�(0�n�܂�)
  # _ty1_::�R�s�[�� ����s�ԍ�(0�n�܂�)
  # _tx1_::�R�s�[�� ����J�����ԍ�(0�n�܂�)
  # _ty2_::�R�s�[�� �E���s�ԍ�(0�n�܂�)
  # _tx2_::�R�s�[�� �E���J�����ԍ�(0�n�܂�)
  def format_copy(sy,sx,ty1,tx1,ty2,tx2)
    s = self.getCellByPosition(sx,sy)
    sp = s.getPropertySetInfo.getProperties
    names = sp.each.map{|p| p.Name}
    ps = s.getPropertyValues(names)
    self.getCellRangeByPosition(tx1,ty1,tx2,ty2).setPropertyValues(names,ps)
  end

  #=== �����R�s�[2(Range)
  #
  # �R�s�[���s�̏������R�s�[���n�s�R�s�[����B
  # �R�s�[��̓R�s�[���Ɠ����J�������Ƃ���B
  #
  # _sy1_::�R�s�[�� �s�ԍ�(0�n�܂�)
  # _sx1_::�R�s�[�� �J�n�J�����ԍ�(0�n�܂�)
  # _sx2_::�R�s�[�� �I���J�����ԍ�(0�n�܂�)
  # _ty_::�R�s�[�� ����s�ԍ�(0�n�܂�)
  # _tx_::�R�s�[�� ����J�����ԍ�(0�n�܂�)
  # _n_:: �R�s�[�s��
  def format_range_copy(sy,sx1,sx2,  ty,tx,n=1)
    return if n < 1
    (sx1..sx2).each do |x|
      self.format_copy(sy,x,  ty,tx+(x-sx1),ty+n-1,tx+(x-sx1))
    end
  end

  #=== �����R�s�[(Cell)
  #
  # _sy_::�R�s�[�� �s�ԍ�(0�n�܂�)
  # _sx_::�R�s�[�� �J�����ԍ�(0�n�܂�)
  # _ty_::�R�s�[�� ����s�ԍ�(0�n�܂�)
  # _tx_::�R�s�[�� ����J�����ԍ�(0�n�܂�)
  def format_copy1(sy,sx,ty,tx)
    s = self.getCellByPosition(sx,sy)
    sp = s.getPropertySetInfo.getProperties
    names = sp.each.map{|p| p.Name}
    ps = s.getPropertyValues(names)
    self.getCellByPosition(tx,ty).setPropertyValues(names,ps)
  end

  #=== �R�s�[(Range)
  #
  # _sy1_::�R�s�[�� ����s�ԍ�(0�n�܂�)
  # _sx1_::�R�s�[�� ����J�����ԍ�(0�n�܂�)
  # _sy2_::�R�s�[�� �E���s�ԍ�(0�n�܂�)
  # _sx2_::�R�s�[�� �E���J�����ԍ�(0�n�܂�)
  # _ty_::�R�s�[�� ����s�ԍ�(0�n�܂�)
  # _tx_::�R�s�[�� ����J�����ԍ�(0�n�܂�)
  def copy(sy1,sx1,sy2,sx2,ty,tx)
    r = self.getCellRangeByPosition(sx1,sy1,sx2,sy2).getRangeAddress
    c = self.getCellByPosition(tx,ty).getCellAddress
    self.copyRange(c,r)
  end


  #=== �s�̑}��
  #
  # _n_:: �s�ԍ�(0�n�܂�)�A���̍s�̑O�ɑ}������B
  # _count_:: ���s�}�����邩�̎w��A�w�肵�Ȃ��ꍇ�ɂ�1�B
  #
  def insert_rows(n,count=1)   #
    self.Rows.insertByIndex(n,count)
  end

  #=== �s�̍폜
  #
  # _n_:: �s�ԍ�(0�n�܂�)�A���̍s���牺���폜����B
  # _count_:: ���s�폜���邩�̎w��A�w�肵�Ȃ��ꍇ�ɂ�1�B
  #
  def remove_rows(n,count=1)   #
    r = self.getCellRangeByPosition(0,n,0,n+count-1).getRangeAddress
    self.removerange(r,3)
  end


end


#----------------------------------------------------
#== OpenOffice�h�L�������g����
#
# Calc�h�L�������g�̊g�����W���[���B
# �h�L�������g���o�����Ɏ����I�ɑg�ݍ��܂��B
module OOoDocument

  #=== �V�[�g�̎��o��
  #
  # _s_:: �V�[�g��������܂��̓V�[�g�̃C���f�b�N�X�ԍ�
  # ret:: �V�[�g�I�u�W�F�N�g
  #
  def get_sheet(s)
    if s.class == String
      sheet = self.sheets.getByName(s)
    else
      sheet = self.sheets.getByIndex(s)
    end
    sheet.extend(CalcWorksheet)
  end

  #=== Active�V�[�g�̐؂�ւ�
  #
  # _s_:: �V�[�g��������܂��̓V�[�g�̃C���f�b�N�X�ԍ�
  #
  def set_active_sheet(s)
    if s.class == String
      self.getCurrentController.setActiveSheet(self.Sheets.getByName(s))
    else
      self.getCurrentController.setActiveSheet(self.Sheets.getByIndex(s))
    end
  end

  #=== Calc�h�L�������g�̕ۑ�(�����o��)
  #
  # _filename_:: ���O��ς��ĕۑ�����ꍇ�Ƀt�@�C�������w�肷��B
  # ret::����������true�A���s������false
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
      print "�������݂ł��܂���ł����B\n"
      done = false
    end
    done
  end
end


#----------------------------------------------------
#== Calc�h�L�������g����

#=== �I�v�V�����w��p�z��̍쐬
#
# _manager_:: com.sun.star.ServiceManager
# _hash_:: �I�v�V�����w��(���O�ƒl�̘A�z�z��)
# ret::�I�v�V�����w��p�̔z��
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

#=== �h�L�������g�`�������o��
#
#  �g���q��ods�Ȃ��̔z���Ԃ��B
#  ���̑��Ȃ�Ή�����t�B���^�[���̃v���p�e�B�z���Ԃ��B
# =filename_::�t�@�C����
# ret::�t�B���^�[�I�v�V����(�h�L�������g�`����) 
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
#=== �����h�L�������g��Open
#
#  �����u���b�N���󂯎���Ď��s����B
#
# _filename:: Calc�h�L�������g�̃t�@�C����
# _visible_:: Calc�̃E�C���h�E��\������Ƃ�true�A�w�肳��Ȃ��ꍇ�ɂ�true
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
##  desktop.terminate   ## �J���Ă��鑼��OOo�h�L�������g��������ɂ��Ă��ׂďI�����Ă��܂�
  
end


#----------------------------------------------------
#=== �V�K�h�L�������g�̍쐬
#
#  �����u���b�N���󂯎���Ď��s����B
#  save���Ăт����Ƃ��Ƀt�@�C�������w�肵�Ă�ۑ�����B
#
# _visible_:: Calc�̃E�C���h�E��\������Ƃ�true�A�w�肳��Ȃ��ꍇ�ɂ�true
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
##  desktop.terminate   ## �J���Ă��鑼��OOo�h�L�������g��������ɂ��Ă��ׂďI�����Ă��܂�

end


if __FILE__ == $0

  print "OpenOffice.org Calc�p ruby�g�����W���[��\n"
  print "\nruby��RDoc�Ńh�L�������g���쐬���Ă��������B\n"
  print "rdoc --title Trail_Calc --main Trail_calc  --all Trail_calc.rb\n"

end
