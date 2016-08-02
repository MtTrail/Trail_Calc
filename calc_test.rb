#! ruby -EWindows-31J
# -*- mode:ruby; coding:Windows-31J -*-

require 'win32ole'
require 'Trail_Calc'
require 'fileutils'

STDOUT.sync = true

#
#  Calc����@�e�X�g�@�v���O����
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

  print "ActiveSheet�؂�ւ� : DataSheet \n"
  book.set_active_sheet("DataSheet")
  
  # Calc�h�L�������g�̓ǂݍ���
  sheet = book.get_sheet("DataSheet")

  # �Z���̓ǂݏ���
  sheet[0,3] = 'D1-'
  print "D1�Z����'D1-'���������݁A�ǂݏo�����l =�u#{sheet[0,3]}�v\n"
  err_check("�����񏑂����݃G���[",'D1-',sheet[0,3])
  
  sheet[0,4] = 100
  print "E1�Z����100���������݁A�ǂݏo�����l =�u#{sheet[0,4]}�v\n"
  err_check("���l�������݃G���[",100,sheet[0,4])
  wait
  
  #�Z���̔w�i�F�̓ǂݏ���
  sheet.set_color(0,3,rgb(0xff,0x00,0x00))
  print "D1�Z���̐F = #{sheet.color(0,3).to_s(16)}\n"
  err_check("D1�Z���̔w�i�F�ݒ�",rgb(0xff,0x00,0x00).to_s(16),sheet.color(0,3).to_s(16))
  
  sheet.set_color(0,4,rgb(0,255,0))
  print "E1�Z���̐F = #{sheet.color(0,4).to_s(16)}\n"
  err_check("E1�Z���̔w�i�F�ݒ�",rgb(0,255,0).to_s(16),sheet.color(0,4).to_s(16))
  
  sheet.set_color(0,5,rgb(0,0,255))
  print "F1�Z���̐F = #{sheet.color(0,5).to_s(16)}\n"
  err_check("F1�Z���̔w�i�F�ݒ�",rgb(0,0,255).to_s(16),sheet.color(0,5).to_s(16))
  wait
  
  print "A22:B24�ɔ����s���N��ݒ�\n"
  sheet.set_range_color(21,0,23,1,rgb(0xff,0xf0,0xf0))
  err_check("A22�Z���̔w�i�F�ݒ�",rgb(0xff,0xf0,0xf0).to_s(16),sheet.color(21,0).to_s(16))
  err_check("B24�Z���̔w�i�F�ݒ�",rgb(0xff,0xf0,0xf0).to_s(16),sheet.color(23,1).to_s(16))
  wait

  print "21�s�̉���1�s��}��\n"
  sheet.insert_rows(21)
  err_check("�s�̑}��",'a22',sheet[22,0])
  wait
  
  print "F�J�����̕���6000�ɐݒ�\n"
  sheet.set_width(5,6000)
  err_check("�J�������ݒ�",6000, sheet.get_width(5))
  print "F1�Ɍ��ݎ�����ݒ�\n"
  t = Time.now
  sheet[0,5] = time_ruby2ooo(t)
  print "F1�̎��� = #{time_ooo2ruby(sheet[0,5]).to_s}\n"
  wait
  
  print "A25�Z���Ɏ� '=1+2+3' ��ݒ�\n"
  sheet.set_formula(24,0,"=1+2+3")
  err_check("�����ݒ�","=1+2+3", sheet.get_formula(24,0))
  err_check("���̒l",6, sheet[24,0])

  
  # chart�h�L�������g�擾
  print "chart�h�L�������g�擾\n"
  
  # �^�C�g���ƃT�u�^�C�g��
  chartDoc = sheet.get_chartdoc
  print "Title : #{chartDoc.get_title}\n"
  print "SubTitle : #{chartDoc.get_subtitle}\n"
  wait

  print "Title/Subtitle ��������\n"
  chartDoc.set_title("Temperature and Pressure")
  print "Title : #{chartDoc.get_title}\n"
  chartDoc.set_subtitle(Time.now.strftime("%Y/%m/%d"))
  print "SubTitle : #{chartDoc.get_subtitle}\n"
  wait

  print "���t�ɕϊ� X�� min = " + time_ooo2ruby(chartDoc.get_Xmin).to_s + "\n"
  print "���t�ɕϊ� X�� max = " + time_ooo2ruby(chartDoc.get_Xmax).to_s + "\n"
  print "X�� min = #{chartDoc.get_Xmin}, max = #{chartDoc.get_Xmax}\n"
  wait
  
  print "min ��������\n"
  chartDoc.set_Xmin(sheet[1,0])
  print "X�� min = #{chartDoc.get_Xmin}, max = #{chartDoc.get_Xmax}\n"

  print "max ��������\n"
  chartDoc.set_Xmax(sheet[19,0])
  print "X�� min = #{chartDoc.get_Xmin}, max = #{chartDoc.get_Xmax}\n"
  print "���t�ɕϊ� X�� min = " + time_ooo2ruby(chartDoc.get_Xmin).to_s + "\n"
  print "���t�ɕϊ� X�� max = " + time_ooo2ruby(chartDoc.get_Xmax).to_s + "\n"
  wait
  
  print "ChartType : #{chartDoc.get_chartType}\n"
  print "Range : #{chartDoc.get_Range}\n"
  print "�\���͈͂�2-20�s�ɕύX\n"
  err_check("�O���t�͈͕ύX",true, chartDoc.change_Xrange(2,20))
  print "Range : #{chartDoc.get_Range}\n"
  wait(2)
  
#------------------sheet�؂�ւ�-----------------------------------
  print "ActiveSheet�؂�ւ� : X^2 \n"
  book.set_active_sheet("X^2")
  sheet2 = book.get_sheet("X^2")
  wait

  chartDoc1 = sheet2.get_chartdoc(0)
  print "Title : #{chartDoc1.get_title}\n"
  print "SubTitle : #{chartDoc1.get_subtitle}\n"
  wait
  
  print "ChartType : #{chartDoc1.get_chartType}\n"
  print "Range : #{chartDoc1.get_Range}\n"
  print "�\���͈͂�2-11�s�ɕύX\n"
  err_check("�O���t�͈͕ύX",true, chartDoc1.change_Xrange(2,11))
  print "Range : #{chartDoc1.get_Range}\n"
  wait

  chartDoc2 = sheet2.get_chartdoc(1)
  print "Title : #{chartDoc2.get_title}\n"
  print "SubTitle : #{chartDoc2.get_subtitle}\n"
  wait

  print "ChartType : #{chartDoc2.get_chartType}\n"
  print "Range : #{chartDoc2.get_Range}\n"
  print "�\���͈͂�2-11�s�ɕύX\n"
  err_check("�O���t�͈͕ύX",true, chartDoc2.change_Xrange(2,11))
  print "Range : #{chartDoc2.get_Range}\n"
  
  print "�`���[�g�����o��\n"
  chartDoc2.save("chart.png")
  
  wait(2)

#------------------sheet�؂�ւ�-----------------------------------
  print "ActiveSheet�؂�ւ� : �\1 \n"
  book.set_active_sheet("�\1")
  sheet3 = book.get_sheet("�\1")
  wait
  
  print "2�s�ڂ���3�s�ڂ܂ł��O���[�v��\n"
  sheet3.group_row(1,2)
  wait

  print "2�s�ڂ���5�s�ڂ܂ł��O���[�v��\n"
  sheet3.group_row(1,4)
  wait

  print "3��ڂ���4��ڂ܂ł��O���[�v��\n"
  sheet3.group_column(2,3)
  wait

  print "A8:E10��wrap�ݒ�\n"
  sheet3.wrap(7,0,9,5)
  wait

  print "A1:E10�Ɍr���g\n"
  sheet3.box(0,0,9,4)
  wait
  
  
  print "A1:C1���}�[�W\n"
  sheet3.set_color(0,0,rgb(255,100,100))
  sheet3.merge(0,0,0,2)
  wait

  print "A6:C6���}�[�W\n"
  sheet3.set_color(5,0,rgb(255,50,50))
  sheet3.merge(5,0,5,2)
  wait

  print "A6:C6���}�[�W����\n"
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
  
  print "�����R�s�[ A10 �� A11\n"
  sheet3.format_copy1(9,0,10,0)
  print "�����R�s�[ B10 �� B11:C11\n"
  sheet3.format_copy(9,1,10,1,10,2)
  print "�����R�s�[ A10:E10 �� A0 2�s\n"
  sheet3.format_range_copy(9,0,4,  0,0, 2)
  wait
  
  print "�R�s�[A10:E10 �� A12\n"
  sheet3.copy(9,0,9,4,  11,0)

  print "�R�s�[A10:E10 �� A12\n"
  sheet3.copy(9,0,9,4,  11,0)
  
  print "�s�폜 $13:$16\n"
  sheet3.remove_rows(12,4)
  wait(2)

#------------------------------------------------------------------
  book.save
  book.save('Data3.xls')
  book.save('Data3.xlsx')
  
end
