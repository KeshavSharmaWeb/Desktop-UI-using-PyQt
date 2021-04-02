# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'mains_1.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!
import os
from typing import Dict, Any
import csv
from matplotlib import pyplot as plt
import pandas
from PyQt5 import QtCore, QtGui, QtWidgets, QtWebEngineWidgets
from PyQt5.QtWidgets import QTableWidgetItem, QListWidgetItem, QMainWindow, QApplication
import plotly
import chart_studio.plotly as py
import MyDtModule as my
import chart_studio.tools as tls
import numpy as np
from PyQt5.QtWebEngineWidgets import QWebEngineView

output_type = ".XML"
source_excel = pandas.read_excel(
    "static\Input\Input Data + Sample output.xlsx", sheet_name="Sheet1")


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(1760, 661)
        Form.setStyleSheet("background-color: rgb(121, 121, 121)")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(Form)
        self.verticalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_2.setSpacing(0)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setSpacing(0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.mainframe = QtWidgets.QFrame(Form)
        self.mainframe.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.mainframe.setFrameShadow(QtWidgets.QFrame.Raised)
        self.mainframe.setObjectName("mainframe")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.mainframe)
        self.verticalLayout_3.setContentsMargins(1, 1, 1, 1)
        self.verticalLayout_3.setSpacing(1)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.bottomframe = QtWidgets.QFrame(self.mainframe)
        self.bottomframe.setFrameShape(QtWidgets.QFrame.WinPanel)
        self.bottomframe.setFrameShadow(QtWidgets.QFrame.Raised)
        self.bottomframe.setObjectName("bottomframe")
        self.verticalLayout_3.addWidget(self.bottomframe)
        self.middleframe = QtWidgets.QFrame(self.mainframe)
        self.middleframe.setFrameShape(QtWidgets.QFrame.WinPanel)
        self.middleframe.setFrameShadow(QtWidgets.QFrame.Raised)
        self.middleframe.setObjectName("middleframe")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.middleframe)
        self.horizontalLayout.setContentsMargins(10, 5, 10, 5)
        self.horizontalLayout.setSpacing(5)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.cateogoryframe = QtWidgets.QFrame(self.middleframe)
        self.cateogoryframe.setFrameShape(QtWidgets.QFrame.WinPanel)
        self.cateogoryframe.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.cateogoryframe.setObjectName("cateogoryframe")
        self.verticalLayout_4 = QtWidgets.QVBoxLayout(self.cateogoryframe)
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.label_2 = QtWidgets.QLabel(self.cateogoryframe)
        font = QtGui.QFont()
        font.setFamily("Cooper Black")
        font.setPointSize(14)
        font.setBold(False)
        font.setWeight(50)
        self.label_2.setFont(font)
        self.label_2.setAlignment(QtCore.Qt.AlignCenter)
        self.label_2.setObjectName("label_2")
        self.verticalLayout_4.addWidget(self.label_2)
        self.PPCategory_listWidget = QtWidgets.QListWidget(self.cateogoryframe)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.PPCategory_listWidget.setFont(font)
        self.PPCategory_listWidget.setStyleSheet(
            "  QListView {\n"
            "      show-decoration-selected: 0;\n"
            "      color:white;\n"
            "      background-color:rgb(56, 56, 56);\n"
            "  }\n"
            "\n"
            "  QListView::item:alternate {\n"
            "      background: #EEEEEE;\n"
            "  }\n"
            "\n"
            "  QListView::item:selected {\n"
            "      border: 1px solid #6a6ea9;\n"
            "  }\n"
            "\n"
            "  QListView::item:selected {\n"
            "      background:yellow;\n"
            "      color:black;\n"
            "  }\n"
            "\n"
            "  QListView::item:hover {\n"
            "      color:blue;\n"
            "      background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,\n"
            "                                  stop: 0 #FAFBFE, stop: 1 #DCDEF1);\n"
            "  }")
        self.PPCategory_listWidget.setFrameShape(QtWidgets.QFrame.WinPanel)
        self.PPCategory_listWidget.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.PPCategory_listWidget.setLineWidth(5)
        self.PPCategory_listWidget.setMidLineWidth(5)
        self.PPCategory_listWidget.setObjectName("PPCategory_listWidget")
        item = QtWidgets.QListWidgetItem()
        self.PPCategory_listWidget.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.PPCategory_listWidget.addItem(item)
        self.verticalLayout_4.addWidget(self.PPCategory_listWidget)
        self.OtherCategory_listWidget = QtWidgets.QListWidget(
            self.cateogoryframe)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.OtherCategory_listWidget.setFont(font)
        self.OtherCategory_listWidget.setStyleSheet(
            "  QListView {\n"
            "      show-decoration-selected: 0;\n"
            "      color:white;\n"
            "      background-color:rgb(56, 56, 56);\n"
            "  }\n"
            "\n"
            "  QListView::item:alternate {\n"
            "      background: #EEEEEE;\n"
            "  }\n"
            "\n"
            "  QListView::item:selected {\n"
            "      border: 1px solid #6a6ea9;\n"
            "  }\n"
            "\n"
            "  QListView::item:selected {\n"
            "      background:yellow;\n"
            "      color:black;\n"
            "  }\n"
            "\n"
            "  QListView::disabled {\n"
            "      background:grey;\n"
            "      color:silver;\n"
            "  }\n"
            "\n"
            "  QListView::item:hover {\n"
            "      color:blue;\n"
            "      background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,\n"
            "                                  stop: 0 #FAFBFE, stop: 1 #DCDEF1);\n"
            "  }")
        self.OtherCategory_listWidget.setFrameShape(QtWidgets.QFrame.WinPanel)
        self.OtherCategory_listWidget.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.OtherCategory_listWidget.setLineWidth(5)
        self.OtherCategory_listWidget.setMidLineWidth(5)
        self.OtherCategory_listWidget.setObjectName("OtherCategory_listWidget")
        item = QtWidgets.QListWidgetItem()
        self.OtherCategory_listWidget.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.OtherCategory_listWidget.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.OtherCategory_listWidget.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.OtherCategory_listWidget.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.OtherCategory_listWidget.addItem(item)
        self.verticalLayout_4.addWidget(self.OtherCategory_listWidget)
        self.Reset_pushButton = QtWidgets.QPushButton(self.cateogoryframe)
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.Reset_pushButton.setFont(font)
        self.Reset_pushButton.setStyleSheet("color:white;\n"
                                            "background-color:red;")
        self.Reset_pushButton.setObjectName("Reset_pushButton")
        self.verticalLayout_4.addWidget(self.Reset_pushButton)
        self.horizontalLayout.addWidget(self.cateogoryframe)
        self.processframe = QtWidgets.QFrame(self.middleframe)
        self.processframe.setFrameShape(QtWidgets.QFrame.WinPanel)
        self.processframe.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.processframe.setLineWidth(1)
        self.processframe.setObjectName("processframe")
        self.gridLayout = QtWidgets.QGridLayout(self.processframe)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout.setObjectName("gridLayout")
        self.stackedWidget = QtWidgets.QStackedWidget(self.processframe)
        self.stackedWidget.setStyleSheet("background-color:rgb(56,56,56)")
        self.stackedWidget.setObjectName("stackedWidget")
        self.PP_ext = QtWidgets.QWidget()
        self.PP_ext.setAutoFillBackground(False)
        self.PP_ext.setStyleSheet("")
        self.PP_ext.setObjectName("PP_ext")
        self.verticalLayout_11 = QtWidgets.QVBoxLayout(self.PP_ext)
        self.verticalLayout_11.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_11.setSpacing(0)
        self.verticalLayout_11.setObjectName("verticalLayout_11")
        self.verticalLayout_5 = QtWidgets.QVBoxLayout()
        self.verticalLayout_5.setSpacing(0)
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.pp_main = QtWidgets.QFrame(self.PP_ext)
        self.pp_main.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.pp_main.setFrameShadow(QtWidgets.QFrame.Raised)
        self.pp_main.setObjectName("pp_main")
        self.verticalLayout_6 = QtWidgets.QVBoxLayout(self.pp_main)
        self.verticalLayout_6.setContentsMargins(0, 30, 0, 15)
        self.verticalLayout_6.setSpacing(0)
        self.verticalLayout_6.setObjectName("verticalLayout_6")
        self.output = QtWidgets.QFrame(self.pp_main)
        self.output.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.output.setFrameShadow(QtWidgets.QFrame.Raised)
        self.output.setObjectName("output")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout(self.output)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.label_3 = QtWidgets.QLabel(self.output)
        font = QtGui.QFont()
        font.setFamily("Courier")
        font.setPointSize(13)
        font.setBold(True)
        font.setWeight(75)
        self.label_3.setFont(font)
        self.label_3.setAutoFillBackground(False)
        self.label_3.setStyleSheet("color:white")
        self.label_3.setTextFormat(QtCore.Qt.RichText)
        self.label_3.setAlignment(QtCore.Qt.AlignRight
                                  | QtCore.Qt.AlignTrailing
                                  | QtCore.Qt.AlignVCenter)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_3.addWidget(self.label_3)
        self.PPInput_lineEdit = QtWidgets.QLineEdit(self.output)
        font = QtGui.QFont()
        font.setPointSize(11)
        self.PPInput_lineEdit.setFont(font)
        self.PPInput_lineEdit.setAutoFillBackground(False)
        self.PPInput_lineEdit.setStyleSheet("background-color:silver;\n"
                                            "color:black;")
        self.PPInput_lineEdit.setReadOnly(True)
        self.PPInput_lineEdit.setObjectName("PPInput_lineEdit")
        self.horizontalLayout_3.addWidget(self.PPInput_lineEdit)
        self.PPInput_pushButton = QtWidgets.QPushButton(self.output)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.PPInput_pushButton.setFont(font)
        self.PPInput_pushButton.setStyleSheet(
            "background-color:rgb(255, 177, 41);")
        self.PPInput_pushButton.setObjectName("PPInput_pushButton")
        self.horizontalLayout_3.addWidget(self.PPInput_pushButton)
        self.verticalLayout_6.addWidget(self.output)
        self.input = QtWidgets.QFrame(self.pp_main)
        self.input.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.input.setFrameShadow(QtWidgets.QFrame.Raised)
        self.input.setObjectName("input")
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout(self.input)
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.label_4 = QtWidgets.QLabel(self.input)
        font = QtGui.QFont()
        font.setFamily("Courier")
        font.setPointSize(13)
        font.setBold(True)
        font.setWeight(75)
        self.label_4.setFont(font)
        self.label_4.setAutoFillBackground(False)
        self.label_4.setStyleSheet("color:white")
        self.label_4.setTextFormat(QtCore.Qt.RichText)
        self.label_4.setAlignment(QtCore.Qt.AlignRight
                                  | QtCore.Qt.AlignTrailing
                                  | QtCore.Qt.AlignVCenter)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout_4.addWidget(self.label_4)
        self.PPOutput_lineEdit = QtWidgets.QLineEdit(self.input)
        font = QtGui.QFont()
        font.setPointSize(11)
        self.PPOutput_lineEdit.setFont(font)
        self.PPOutput_lineEdit.setAutoFillBackground(False)
        self.PPOutput_lineEdit.setStyleSheet("background-color:silver;\n"
                                             "color:black;")
        self.PPOutput_lineEdit.setReadOnly(True)
        self.PPOutput_lineEdit.setObjectName("PPOutput_lineEdit")
        self.horizontalLayout_4.addWidget(self.PPOutput_lineEdit)
        self.PPOutput_pushButton = QtWidgets.QPushButton(self.input)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.PPOutput_pushButton.setFont(font)
        self.PPOutput_pushButton.setAutoFillBackground(False)
        self.PPOutput_pushButton.setStyleSheet(
            "background-color:rgb(255, 177, 41)")
        self.PPOutput_pushButton.setObjectName("PPOutput_pushButton")
        self.horizontalLayout_4.addWidget(self.PPOutput_pushButton)
        self.verticalLayout_6.addWidget(self.input)
        self.convert = QtWidgets.QFrame(self.pp_main)
        self.convert.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.convert.setFrameShadow(QtWidgets.QFrame.Raised)
        self.convert.setObjectName("convert")
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout(self.convert)
        self.horizontalLayout_5.setContentsMargins(120, 2, 10, 2)
        self.horizontalLayout_5.setSpacing(2)
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        self.output_format = QtWidgets.QFrame(self.convert)
        self.output_format.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.output_format.setFrameShadow(QtWidgets.QFrame.Raised)
        self.output_format.setObjectName("output_format")
        self.horizontalLayout_6 = QtWidgets.QHBoxLayout(self.output_format)
        self.horizontalLayout_6.setObjectName("horizontalLayout_6")
        self.label_12 = QtWidgets.QLabel(self.output_format)
        font = QtGui.QFont()
        font.setFamily("Courier")
        font.setPointSize(13)
        font.setBold(True)
        font.setWeight(75)
        self.label_12.setFont(font)
        self.label_12.setAutoFillBackground(False)
        self.label_12.setStyleSheet("color:white")
        self.label_12.setTextFormat(QtCore.Qt.RichText)
        self.label_12.setAlignment(QtCore.Qt.AlignCenter)
        self.label_12.setObjectName("label_12")
        self.horizontalLayout_6.addWidget(self.label_12)
        ##################My Fixed####################
        self.PPXML_CheckBox_2 = QtWidgets.QRadioButton(self.output_format)
        ##############################################
        font = QtGui.QFont()
        font.setPointSize(12)
        self.PPXML_CheckBox_2.setFont(font)
        self.PPXML_CheckBox_2.setLayoutDirection(QtCore.Qt.RightToLeft)
        self.PPXML_CheckBox_2.setAutoFillBackground(False)
        self.PPXML_CheckBox_2.setStyleSheet("color:white;")
        self.PPXML_CheckBox_2.setObjectName("PPXML_CheckBox_2")
        self.horizontalLayout_6.addWidget(self.PPXML_CheckBox_2, 0,
                                          QtCore.Qt.AlignHCenter)
        ##################My Fixed####################
        self.PPJSON_CheckBox_2 = QtWidgets.QRadioButton(self.output_format)
        ##############################################
        font = QtGui.QFont()
        font.setPointSize(12)
        self.PPJSON_CheckBox_2.setFont(font)
        self.PPJSON_CheckBox_2.setLayoutDirection(QtCore.Qt.RightToLeft)
        self.PPJSON_CheckBox_2.setAutoFillBackground(False)
        self.PPJSON_CheckBox_2.setStyleSheet("color:white;")
        self.PPJSON_CheckBox_2.setObjectName("PPJSON_CheckBox_2")
        #################### My Fixed####################
        ##################
        self.PPXML_CheckBox_2.setChecked(True)
        self.PPXML_CheckBox_2.toggled.connect(
            lambda: self.isCheckedXMLORJSON(self.PPXML_CheckBox_2))
        self.PPJSON_CheckBox_2.toggled.connect(
            lambda: self.isCheckedXMLORJSON(self.PPJSON_CheckBox_2))
        #################
        #################################################
        self.horizontalLayout_6.addWidget(self.PPJSON_CheckBox_2)
        self.horizontalLayout_5.addWidget(self.output_format, 0,
                                          QtCore.Qt.AlignHCenter)
        self.convert_2 = QtWidgets.QFrame(self.convert)
        self.convert_2.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.convert_2.setFrameShadow(QtWidgets.QFrame.Raised)
        self.convert_2.setObjectName("convert_2")
        self.horizontalLayout_7 = QtWidgets.QHBoxLayout(self.convert_2)
        self.horizontalLayout_7.setContentsMargins(70, 2, 2, 2)
        self.horizontalLayout_7.setSpacing(2)
        self.horizontalLayout_7.setObjectName("horizontalLayout_7")
        self.PPConvert_pushButton_2 = QtWidgets.QPushButton(self.convert_2)
        self.PPConvert_pushButton_2.setMinimumSize(QtCore.QSize(158, 0))
        self.PPConvert_pushButton_2.setMaximumSize(QtCore.QSize(158, 16777215))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.PPConvert_pushButton_2.setFont(font)
        self.PPConvert_pushButton_2.setAutoFillBackground(False)
        self.PPConvert_pushButton_2.setStyleSheet(
            "background-color:rgb(255, 177, 41)")
        self.PPConvert_pushButton_2.setFlat(False)
        self.PPConvert_pushButton_2.setObjectName("PPConvert_pushButton_2")
        self.horizontalLayout_7.addWidget(self.PPConvert_pushButton_2)
        self.horizontalLayout_5.addWidget(self.convert_2)
        self.horizontalLayout_5.setStretch(0, 10)
        self.horizontalLayout_5.setStretch(1, 5)
        self.verticalLayout_6.addWidget(self.convert)
        self.result = QtWidgets.QFrame(self.pp_main)
        self.result.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.result.setFrameShadow(QtWidgets.QFrame.Raised)
        self.result.setObjectName("result")
        self.horizontalLayout_8 = QtWidgets.QHBoxLayout(self.result)
        self.horizontalLayout_8.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_8.setSpacing(0)
        self.horizontalLayout_8.setObjectName("horizontalLayout_8")
        self.frame = QtWidgets.QFrame(self.result)
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.verticalLayout_7 = QtWidgets.QVBoxLayout(self.frame)
        self.verticalLayout_7.setObjectName("verticalLayout_7")
        self.label_13 = QtWidgets.QLabel(self.frame)
        font = QtGui.QFont()
        font.setFamily("Courier")
        font.setPointSize(13)
        font.setBold(True)
        font.setWeight(75)
        self.label_13.setFont(font)
        self.label_13.setAutoFillBackground(False)
        self.label_13.setStyleSheet("color:white")
        self.label_13.setTextFormat(QtCore.Qt.RichText)
        self.label_13.setAlignment(QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft
                                   | QtCore.Qt.AlignVCenter)
        self.label_13.setObjectName("label_13")
        self.verticalLayout_7.addWidget(self.label_13)
        self.PPFile_listWidget_2 = QtWidgets.QListWidget(self.frame)
        self.PPFile_listWidget_2.setAutoFillBackground(True)
        self.PPFile_listWidget_2.setStyleSheet("background-color:silver;\n"
                                               "color:black;\n"
                                               "")
        self.PPFile_listWidget_2.setObjectName("PPFile_listWidget_2")
        self.verticalLayout_7.addWidget(self.PPFile_listWidget_2)
        self.horizontalLayout_8.addWidget(self.frame)
        self.frame_2 = QtWidgets.QFrame(self.result)
        self.frame_2.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_2.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_2.setObjectName("frame_2")
        self.verticalLayout_8 = QtWidgets.QVBoxLayout(self.frame_2)
        self.verticalLayout_8.setObjectName("verticalLayout_8")
        self.label_14 = QtWidgets.QLabel(self.frame_2)
        font = QtGui.QFont()
        font.setFamily("Courier")
        font.setPointSize(13)
        font.setBold(True)
        font.setWeight(75)
        self.label_14.setFont(font)
        self.label_14.setAutoFillBackground(False)
        self.label_14.setStyleSheet("color:white")
        self.label_14.setTextFormat(QtCore.Qt.RichText)
        self.label_14.setAlignment(QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft
                                   | QtCore.Qt.AlignVCenter)
        self.label_14.setObjectName("label_14")
        self.verticalLayout_8.addWidget(self.label_14)
        self.PPContent_textEdit_2 = QtWidgets.QTextEdit(self.frame_2)
        self.PPContent_textEdit_2.setAutoFillBackground(True)
        self.PPContent_textEdit_2.setStyleSheet("background-color:silver;\n"
                                                "color:black;")
        self.PPContent_textEdit_2.setObjectName("PPContent_textEdit_2")
        self.verticalLayout_8.addWidget(self.PPContent_textEdit_2)
        self.horizontalLayout_8.addWidget(self.frame_2)
        self.horizontalLayout_8.setStretch(0, 4)
        self.horizontalLayout_8.setStretch(1, 7)
        self.verticalLayout_6.addWidget(self.result)
        self.progress = QtWidgets.QFrame(self.pp_main)
        self.progress.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.progress.setFrameShadow(QtWidgets.QFrame.Raised)
        self.progress.setObjectName("progress")
        self.verticalLayout_9 = QtWidgets.QVBoxLayout(self.progress)
        self.verticalLayout_9.setObjectName("verticalLayout_9")
        self.PP_progressBar_2 = QtWidgets.QProgressBar(self.progress)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.PP_progressBar_2.setFont(font)
        self.PP_progressBar_2.setAutoFillBackground(False)
        self.PP_progressBar_2.setStyleSheet(
            " QProgressBar {\n"
            "      border: 1px solid rgba(90, 172, 142, 10);\n"
            "      border-radius: 5px;\n"
            "      text-align: center;\n"
            "      background-color:rgb(217, 217, 217);\n"
            "      color:black;\n"
            "  }\n"
            " QProgressBar::chunk {\n"
            "      background:qlineargradient(spread:reflect, x1:0, y1:0, x2:0.688, y2:0.267045, stop:0 rgba(121, 75, 255, 255), stop:0.107345 rgba(2, 79, 0, 255), stop:0.638418 rgba(91, 153, 77, 255)) ;\n"
            "      border-radius: 9px;\n"
            "      border: 1px solid rgb(156, 255, 156,50);\n"
            "  }")
        self.PP_progressBar_2.setProperty("value", 0)
        self.PP_progressBar_2.setAlignment(QtCore.Qt.AlignCenter)
        self.PP_progressBar_2.setOrientation(QtCore.Qt.Horizontal)
        self.PP_progressBar_2.setTextDirection(
            QtWidgets.QProgressBar.TopToBottom)
        self.PP_progressBar_2.setObjectName("PP_progressBar_2")
        self.verticalLayout_9.addWidget(self.PP_progressBar_2)
        self.verticalLayout_6.addWidget(self.progress)
        self.verticalLayout_6.setStretch(0, 3)
        self.verticalLayout_6.setStretch(1, 3)
        self.verticalLayout_6.setStretch(2, 3)
        self.verticalLayout_6.setStretch(3, 18)
        self.verticalLayout_6.setStretch(4, 4)
        self.verticalLayout_5.addWidget(self.pp_main)
        self.verticalLayout_11.addLayout(self.verticalLayout_5)
        self.stackedWidget.addWidget(self.PP_ext)
        self.RCM_ext = QtWidgets.QWidget()
        self.RCM_ext.setStyleSheet("")
        self.RCM_ext.setObjectName("RCM_ext")
        self.verticalLayout_21 = QtWidgets.QVBoxLayout(self.RCM_ext)
        self.verticalLayout_21.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_21.setSpacing(0)
        self.verticalLayout_21.setObjectName("verticalLayout_21")
        self.verticalLayout_10 = QtWidgets.QVBoxLayout()
        self.verticalLayout_10.setContentsMargins(-1, 30, -1, 15)
        self.verticalLayout_10.setSpacing(0)
        self.verticalLayout_10.setObjectName("verticalLayout_10")
        self.frame_6 = QtWidgets.QFrame(self.RCM_ext)
        self.frame_6.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_6.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_6.setObjectName("frame_6")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.frame_6)
        self.horizontalLayout_2.setContentsMargins(20, 0, 20, 0)
        self.horizontalLayout_2.setSpacing(6)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_15 = QtWidgets.QLabel(self.frame_6)
        font = QtGui.QFont()
        font.setFamily("Courier")
        font.setPointSize(13)
        font.setBold(True)
        font.setWeight(75)
        self.label_15.setFont(font)
        self.label_15.setAutoFillBackground(False)
        self.label_15.setStyleSheet("color:white")
        self.label_15.setTextFormat(QtCore.Qt.RichText)
        self.label_15.setAlignment(QtCore.Qt.AlignCenter)
        self.label_15.setObjectName("label_15")
        self.horizontalLayout_2.addWidget(self.label_15)
        self.RCMInput_lineEdi_2 = QtWidgets.QLineEdit(self.frame_6)
        font = QtGui.QFont()
        font.setPointSize(11)
        self.RCMInput_lineEdi_2.setFont(font)
        self.RCMInput_lineEdi_2.setAutoFillBackground(False)
        self.RCMInput_lineEdi_2.setStyleSheet("background-color:silver;\n"
                                              "color:black;")
        self.RCMInput_lineEdi_2.setReadOnly(True)
        self.RCMInput_lineEdi_2.setObjectName("RCMInput_lineEdi_2")
        self.horizontalLayout_2.addWidget(self.RCMInput_lineEdi_2)
        self.RCMInput_pushButton_2 = QtWidgets.QPushButton(self.frame_6)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.RCMInput_pushButton_2.setFont(font)
        self.RCMInput_pushButton_2.setStyleSheet(
            "background-color:rgb(255, 177, 41);")
        self.RCMInput_pushButton_2.setObjectName("RCMInput_pushButton_2")
        self.horizontalLayout_2.addWidget(self.RCMInput_pushButton_2)
        self.verticalLayout_10.addWidget(self.frame_6)
        self.frame_5 = QtWidgets.QFrame(self.RCM_ext)
        self.frame_5.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_5.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_5.setObjectName("frame_5")
        self.verticalLayout_14 = QtWidgets.QVBoxLayout(self.frame_5)
        self.verticalLayout_14.setContentsMargins(3, 0, 20, 0)
        self.verticalLayout_14.setObjectName("verticalLayout_14")
        self.frame_7 = QtWidgets.QFrame(self.frame_5)
        self.frame_7.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_7.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_7.setObjectName("frame_7")
        self.horizontalLayout_9 = QtWidgets.QHBoxLayout(self.frame_7)
        self.horizontalLayout_9.setContentsMargins(7, 0, 0, 0)
        self.horizontalLayout_9.setSpacing(6)
        self.horizontalLayout_9.setObjectName("horizontalLayout_9")
        self.label_24 = QtWidgets.QLabel(self.frame_7)
        font = QtGui.QFont()
        font.setFamily("Courier")
        font.setPointSize(13)
        font.setBold(True)
        font.setWeight(75)
        self.label_24.setFont(font)
        self.label_24.setAutoFillBackground(False)
        self.label_24.setStyleSheet("color:white")
        self.label_24.setTextFormat(QtCore.Qt.RichText)
        self.label_24.setAlignment(QtCore.Qt.AlignCenter)
        self.label_24.setObjectName("label_24")
        self.horizontalLayout_9.addWidget(self.label_24)
        self.RCMOutput_LineText = QtWidgets.QLineEdit(self.frame_7)
        font = QtGui.QFont()
        font.setPointSize(11)
        self.RCMOutput_LineText.setFont(font)
        self.RCMOutput_LineText.setAutoFillBackground(False)
        self.RCMOutput_LineText.setStyleSheet("background-color:silver;\n"
                                              "color:black;")
        self.RCMOutput_LineText.setReadOnly(True)
        self.RCMOutput_LineText.setObjectName("RCMOutput_LineText")
        self.horizontalLayout_9.addWidget(self.RCMOutput_LineText)
        self.RCMOutput_pushButton_4 = QtWidgets.QPushButton(self.frame_7)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.RCMOutput_pushButton_4.setFont(font)
        self.RCMOutput_pushButton_4.setStyleSheet(
            "background-color:rgb(255, 177, 41);")
        self.RCMOutput_pushButton_4.setObjectName("RCMOutput_pushButton_4")
        self.horizontalLayout_9.addWidget(self.RCMOutput_pushButton_4)
        self.verticalLayout_14.addWidget(self.frame_7)
        self.RCMConvert_pushButton_2 = QtWidgets.QPushButton(self.frame_5)
        self.RCMConvert_pushButton_2.setMaximumSize(QtCore.QSize(75, 16777215))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.RCMConvert_pushButton_2.setFont(font)
        self.RCMConvert_pushButton_2.setAutoFillBackground(False)
        self.RCMConvert_pushButton_2.setStyleSheet(
            "background-color:rgb(255, 177, 41)")
        self.RCMConvert_pushButton_2.setObjectName("RCMConvert_pushButton_2")
        self.verticalLayout_14.addWidget(self.RCMConvert_pushButton_2, 0,
                                         QtCore.Qt.AlignRight)
        self.verticalLayout_10.addWidget(self.frame_5)
        self.frame_3 = QtWidgets.QFrame(self.RCM_ext)
        self.frame_3.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_3.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_3.setObjectName("frame_3")
        self.verticalLayout_13 = QtWidgets.QVBoxLayout(self.frame_3)
        self.verticalLayout_13.setContentsMargins(-1, 0, -1, 0)
        self.verticalLayout_13.setObjectName("verticalLayout_13")
        self.label_16 = QtWidgets.QLabel(self.frame_3)
        font = QtGui.QFont()
        font.setFamily("Courier")
        font.setPointSize(13)
        font.setBold(True)
        font.setWeight(75)
        self.label_16.setFont(font)
        self.label_16.setAutoFillBackground(False)
        self.label_16.setStyleSheet("color:white")
        self.label_16.setTextFormat(QtCore.Qt.RichText)
        self.label_16.setAlignment(QtCore.Qt.AlignCenter)
        self.label_16.setObjectName("label_16")
        self.verticalLayout_13.addWidget(self.label_16, 0, QtCore.Qt.AlignLeft)
        self.RCMContent_tableWidget_2 = QtWidgets.QTableWidget(self.frame_3)
        self.RCMContent_tableWidget_2.setStyleSheet(
            "background-color:silver;\n"
            "color:black;")
        self.RCMContent_tableWidget_2.setObjectName("RCMContent_tableWidget_2")
        self.RCMContent_tableWidget_2.setColumnCount(0)
        self.RCMContent_tableWidget_2.setRowCount(0)
        self.verticalLayout_13.addWidget(self.RCMContent_tableWidget_2)
        self.verticalLayout_10.addWidget(self.frame_3)
        self.frame_4 = QtWidgets.QFrame(self.RCM_ext)
        self.frame_4.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_4.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_4.setObjectName("frame_4")
        self.horizontalLayout_10 = QtWidgets.QHBoxLayout(self.frame_4)
        self.horizontalLayout_10.setObjectName("horizontalLayout_10")
        self.RCM_progressBar_3 = QtWidgets.QProgressBar(self.frame_4)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.RCM_progressBar_3.setFont(font)
        self.RCM_progressBar_3.setStyleSheet(
            " QProgressBar {\n"
            "      border: 1px solid rgba(90, 172, 142, 10);\n"
            "      border-radius: 5px;\n"
            "      text-align: center;\n"
            "      background-color:rgb(217, 217, 217);\n"
            "      color:black;\n"
            "  }\n"
            " QProgressBar::chunk {\n"
            "      background:qlineargradient(spread:reflect, x1:0, y1:0, x2:0.688, y2:0.267045, stop:0 rgba(121, 75, 255, 255), stop:0.107345 rgba(2, 79, 0, 255), stop:0.638418 rgba(91, 153, 77, 255)) ;\n"
            "      border-radius: 9px;\n"
            "      border: 1px solid rgb(156, 255, 156,50);\n"
            "  }")
        self.RCM_progressBar_3.setProperty("value", 0)
        self.RCM_progressBar_3.setAlignment(QtCore.Qt.AlignCenter)
        self.RCM_progressBar_3.setOrientation(QtCore.Qt.Horizontal)
        self.RCM_progressBar_3.setTextDirection(
            QtWidgets.QProgressBar.TopToBottom)
        self.RCM_progressBar_3.setObjectName("RCM_progressBar_3")
        self.horizontalLayout_10.addWidget(self.RCM_progressBar_3)
        self.verticalLayout_10.addWidget(self.frame_4)
        self.verticalLayout_10.setStretch(0, 4)
        self.verticalLayout_10.setStretch(1, 8)
        self.verticalLayout_10.setStretch(2, 30)
        self.verticalLayout_10.setStretch(3, 5)
        self.verticalLayout_21.addLayout(self.verticalLayout_10)
        self.stackedWidget.addWidget(self.RCM_ext)
        self.RCM_stat = QtWidgets.QWidget()
        self.RCM_stat.setObjectName("RCM_stat")
        self.verticalLayout_12 = QtWidgets.QVBoxLayout(self.RCM_stat)
        self.verticalLayout_12.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_12.setSpacing(0)
        self.verticalLayout_12.setObjectName("verticalLayout_12")
        self.verticalLayout_15 = QtWidgets.QVBoxLayout()
        self.verticalLayout_15.setObjectName("verticalLayout_15")
        self.frame_8 = QtWidgets.QFrame(self.RCM_stat)
        self.frame_8.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_8.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_8.setObjectName("frame_8")
        self.verticalLayout_16 = QtWidgets.QVBoxLayout(self.frame_8)
        self.verticalLayout_16.setObjectName("verticalLayout_16")
        self.frame_9 = QtWidgets.QFrame(self.frame_8)
        self.frame_9.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_9.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_9.setObjectName("frame_9")
        self.verticalLayout_17 = QtWidgets.QVBoxLayout(self.frame_9)
        self.verticalLayout_17.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_17.setObjectName("verticalLayout_17")
        self.frame_11 = QtWidgets.QFrame(self.frame_9)
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.frame_11.setFont(font)
        self.frame_11.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_11.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_11.setObjectName("frame_11")
        self.horizontalLayout_12 = QtWidgets.QHBoxLayout(self.frame_11)
        self.horizontalLayout_12.setContentsMargins(-1, -1, -1, 2)
        self.horizontalLayout_12.setObjectName("horizontalLayout_12")
        self.label = QtWidgets.QLabel(self.frame_11)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setStyleSheet("color:white")
        self.label.setObjectName("label")
        self.horizontalLayout_12.addWidget(self.label)
        self.Stat_Total_Number_Label = QtWidgets.QLabel(self.frame_11)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.Stat_Total_Number_Label.setFont(font)
        self.Stat_Total_Number_Label.setStyleSheet("color:white")
        self.Stat_Total_Number_Label.setObjectName("Stat_Total_Number_Label")
        self.horizontalLayout_12.addWidget(self.Stat_Total_Number_Label)
        self.horizontalLayout_12.setStretch(0, 4)
        self.horizontalLayout_12.setStretch(1, 6)
        self.verticalLayout_17.addWidget(self.frame_11)
        self.frame_12 = QtWidgets.QFrame(self.frame_9)
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.frame_12.setFont(font)
        self.frame_12.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_12.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_12.setObjectName("frame_12")
        self.horizontalLayout_13 = QtWidgets.QHBoxLayout(self.frame_12)
        self.horizontalLayout_13.setContentsMargins(-1, 2, -1, 2)
        self.horizontalLayout_13.setObjectName("horizontalLayout_13")
        self.label_6 = QtWidgets.QLabel(self.frame_12)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_6.setFont(font)
        self.label_6.setStyleSheet("color:white")
        self.label_6.setObjectName("label_6")
        self.horizontalLayout_13.addWidget(self.label_6)
        self.Stat_RCM_Names_listWidget = QtWidgets.QListWidget(self.frame_12)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.Stat_RCM_Names_listWidget.setFont(font)
        self.Stat_RCM_Names_listWidget.setStyleSheet("color:white")
        self.Stat_RCM_Names_listWidget.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.Stat_RCM_Names_listWidget.setLineWidth(0)
        self.Stat_RCM_Names_listWidget.setSelectionMode(
            QtWidgets.QAbstractItemView.NoSelection)
        self.Stat_RCM_Names_listWidget.setObjectName(
            "Stat_RCM_Names_listWidget")
        self.horizontalLayout_13.addWidget(self.Stat_RCM_Names_listWidget)
        self.horizontalLayout_13.setStretch(0, 4)
        self.horizontalLayout_13.setStretch(1, 6)
        self.verticalLayout_17.addWidget(self.frame_12)
        self.frame_14 = QtWidgets.QFrame(self.frame_9)
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.frame_14.setFont(font)
        self.frame_14.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_14.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_14.setObjectName("frame_14")
        self.horizontalLayout_14 = QtWidgets.QHBoxLayout(self.frame_14)
        self.horizontalLayout_14.setContentsMargins(-1, 2, -1, 2)
        self.horizontalLayout_14.setObjectName("horizontalLayout_14")
        self.label_7 = QtWidgets.QLabel(self.frame_14)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_7.setFont(font)
        self.label_7.setStyleSheet("color:white")
        self.label_7.setObjectName("label_7")
        self.horizontalLayout_14.addWidget(self.label_7)
        self.Stat_RCM_Location_Label = QtWidgets.QLabel(self.frame_14)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.Stat_RCM_Location_Label.setFont(font)
        self.Stat_RCM_Location_Label.setStyleSheet("color:white")
        self.Stat_RCM_Location_Label.setObjectName("Stat_RCM_Location_Label")
        self.horizontalLayout_14.addWidget(self.Stat_RCM_Location_Label)
        self.horizontalLayout_14.setStretch(0, 4)
        self.horizontalLayout_14.setStretch(1, 6)
        self.verticalLayout_17.addWidget(self.frame_14)
        self.frame_13 = QtWidgets.QFrame(self.frame_9)
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.frame_13.setFont(font)
        self.frame_13.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_13.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_13.setObjectName("frame_13")
        self.horizontalLayout_15 = QtWidgets.QHBoxLayout(self.frame_13)
        self.horizontalLayout_15.setContentsMargins(9, 0, 9, 2)
        self.horizontalLayout_15.setObjectName("horizontalLayout_15")
        self.verticalLayout_17.addWidget(self.frame_13)
        self.verticalLayout_17.setStretch(0, 1)
        self.verticalLayout_17.setStretch(1, 3)
        self.verticalLayout_17.setStretch(2, 1)
        self.verticalLayout_17.setStretch(3, 1)
        self.verticalLayout_16.addWidget(self.frame_9)
        self.frame_10 = QtWidgets.QFrame(self.frame_8)
        self.frame_10.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_10.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_10.setObjectName("frame_10")
        self.horizontalLayout_16 = QtWidgets.QHBoxLayout(self.frame_10)
        self.horizontalLayout_16.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_16.setSpacing(6)
        self.horizontalLayout_16.setObjectName("horizontalLayout_16")
        self.frame_16 = QtWidgets.QFrame(self.frame_10)
        self.frame_16.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_16.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_16.setObjectName("frame_16")
        self.verticalLayout_18 = QtWidgets.QVBoxLayout(self.frame_16)
        self.verticalLayout_18.setObjectName("verticalLayout_18")
        self.label_8 = QtWidgets.QLabel(self.frame_16)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_8.setFont(font)
        self.label_8.setStyleSheet("color:white")
        self.label_8.setObjectName("label_8")
        self.verticalLayout_18.addWidget(self.label_8)
        self.Stat_Missing_Columns_tableWidget_3 = QtWidgets.QTableWidget(
            self.frame_16)
        self.Stat_Missing_Columns_tableWidget_3.setStyleSheet(
            "background-color:silver;\n"
            "color:black;")
        self.Stat_Missing_Columns_tableWidget_3.setObjectName(
            "Stat_Missing_Columns_tableWidget_3")
        self.Stat_Missing_Columns_tableWidget_3.setColumnCount(0)
        self.Stat_Missing_Columns_tableWidget_3.setRowCount(0)
        self.verticalLayout_18.addWidget(
            self.Stat_Missing_Columns_tableWidget_3)
        self.horizontalLayout_16.addWidget(self.frame_16)
        self.frame_15 = QtWidgets.QFrame(self.frame_10)
        self.frame_15.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_15.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_15.setObjectName("frame_15")
        self.verticalLayout_19 = QtWidgets.QVBoxLayout(self.frame_15)
        self.verticalLayout_19.setObjectName("verticalLayout_19")
        self.label_9 = QtWidgets.QLabel(self.frame_15)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_9.setFont(font)
        self.label_9.setStyleSheet("color:white")
        self.label_9.setObjectName("label_9")
        self.verticalLayout_19.addWidget(self.label_9)
        self.Stat_PoMi_Data_tableWidget_4 = QtWidgets.QTableWidget(
            self.frame_15)
        self.Stat_PoMi_Data_tableWidget_4.setStyleSheet(
            "background-color:silver;\n"
            "color:black;")
        self.Stat_PoMi_Data_tableWidget_4.setObjectName(
            "Stat_PoMi_Data_tableWidget_4")
        self.Stat_PoMi_Data_tableWidget_4.setColumnCount(0)
        self.Stat_PoMi_Data_tableWidget_4.setRowCount(0)
        self.verticalLayout_19.addWidget(self.Stat_PoMi_Data_tableWidget_4)
        self.horizontalLayout_16.addWidget(self.frame_15)
        self.verticalLayout_16.addWidget(self.frame_10)
        self.verticalLayout_16.setStretch(0, 3)
        self.verticalLayout_16.setStretch(1, 10)
        self.verticalLayout_15.addWidget(self.frame_8)
        self.verticalLayout_12.addLayout(self.verticalLayout_15)
        self.stackedWidget.addWidget(self.RCM_stat)
        self.PRC_stat = QtWidgets.QWidget()
        self.PRC_stat.setStyleSheet("")
        self.PRC_stat.setObjectName("PRC_stat")
        self.verticalLayout_22 = QtWidgets.QVBoxLayout(self.PRC_stat)
        self.verticalLayout_22.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_22.setSpacing(0)
        self.verticalLayout_22.setObjectName("verticalLayout_22")
        self.verticalLayout_20 = QtWidgets.QVBoxLayout()
        self.verticalLayout_20.setObjectName("verticalLayout_20")
        self.frame_17 = QtWidgets.QFrame(self.PRC_stat)
        self.frame_17.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_17.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_17.setObjectName("frame_17")
        self.verticalLayout_23 = QtWidgets.QVBoxLayout(self.frame_17)
        self.verticalLayout_23.setContentsMargins(9, 9, 9, 9)
        self.verticalLayout_23.setSpacing(9)
        self.verticalLayout_23.setObjectName("verticalLayout_23")

        self.frame_18 = QtWidgets.QFrame(self.frame_17)
        self.frame_18.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_18.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_18.setObjectName("frame_18")
        # self.horizontalLayout_11 = QtWidgets.QHBoxLayout(self.frame_18)
        # self.horizontalLayout_11.setContentsMargins(0, -1, 0, -1)
        # self.horizontalLayout_11.setSpacing(0)
        # self.horizontalLayout_11.setObjectName("horizontalLayout_11")

        self.process_domain = QtWidgets.QLabel(self.frame_18)
        self.process_domain.setGeometry(QtCore.QRect(15, 40, 251, 91))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.process_domain.setFont(font)
        self.process_domain.setAcceptDrops(False)
        self.process_domain.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.process_domain.setStyleSheet(
            "background-color:rgb(160, 160, 160);\n"
            "")
        self.process_domain.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.process_domain.setFrameShadow(QtWidgets.QFrame.Plain)
        self.process_domain.setAlignment(QtCore.Qt.AlignCenter)
        self.process_domain.setObjectName("process_domain")
        self.process_domain.setText("Number of process domains\n\n")

        self.process_domain_value = QtWidgets.QLabel(self.process_domain)
        self.process_domain_value.setGeometry(QtCore.QRect(15, 40, 240, 35))
        font = QtGui.QFont()
        font.setPointSize(13)
        font.setBold(True)
        font.setWeight(75)
        self.process_domain_value.setFont(font)
        self.process_domain_value.setAlignment(QtCore.Qt.AlignCenter)
        self.process_domain_value.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.process_domain_value.setStyleSheet(
            "\n"
            "background-color:rgb(160, 160, 160);")
        self.process_domain_value.setObjectName("process_domain_value")
        self.process_domain_value.setText(self.open_sheet("Process Domain"))

        self.owners = QtWidgets.QLabel(self.frame_18)
        self.owners.setGeometry(QtCore.QRect(786, 40, 251, 91))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.owners.setFont(font)
        self.owners.setAcceptDrops(False)
        self.owners.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.owners.setStyleSheet("\n" "background-color:rgb(160, 160, 160);")
        self.owners.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.owners.setFrameShadow(QtWidgets.QFrame.Plain)
        self.owners.setAlignment(QtCore.Qt.AlignCenter)
        self.owners.setObjectName("owners")
        self.owners.setText("Number of control owners\n\n")
        self.owners_value = QtWidgets.QLabel(self.owners)
        self.owners_value.setGeometry(QtCore.QRect(15, 40, 240, 35))
        font = QtGui.QFont()
        font.setPointSize(13)
        font.setBold(True)
        font.setWeight(75)
        self.owners_value.setFont(font)
        self.owners_value.setAlignment(QtCore.Qt.AlignCenter)
        self.owners_value.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.owners_value.setStyleSheet("\n"
                                        "background-color:rgb(160, 160, 160);")
        self.owners_value.setObjectName("owners_value")
        self.owners_value.setText(self.open_sheet("Control Owner"))

        self.controls = QtWidgets.QLabel(self.frame_18)
        self.controls.setGeometry(QtCore.QRect(529, 40, 251, 91))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.controls.setFont(font)
        self.controls.setAcceptDrops(False)
        self.controls.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.controls.setStyleSheet("background-color:rgb(160, 160, 160);")
        self.controls.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.controls.setFrameShadow(QtWidgets.QFrame.Plain)
        self.controls.setAlignment(QtCore.Qt.AlignCenter)
        self.controls.setObjectName("controls")
        self.controls.setText("Number of controls\n\n")

        self.controls_value = QtWidgets.QLabel(self.controls)
        self.controls_value.setGeometry(QtCore.QRect(15, 40, 240, 35))
        font = QtGui.QFont()
        font.setPointSize(13)
        font.setBold(True)
        font.setWeight(75)
        self.controls_value.setFont(font)
        self.controls_value.setAlignment(QtCore.Qt.AlignCenter)
        self.controls_value.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.controls_value.setStyleSheet(
            "\n"
            "background-color:rgb(160, 160, 160);")
        self.controls_value.setObjectName("controls_value")
        self.controls_value.setText(self.open_sheet("Control ID"))

        self.risks = QtWidgets.QLabel(self.frame_18)
        self.risks.setGeometry(QtCore.QRect(272, 40, 251, 91))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.risks.setFont(font)
        self.risks.setAcceptDrops(False)
        self.risks.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.risks.setStyleSheet("background-color: rgb(160, 160, 160);")
        self.risks.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.risks.setFrameShadow(QtWidgets.QFrame.Plain)
        self.risks.setAlignment(QtCore.Qt.AlignCenter)
        self.risks.setObjectName("risks")
        self.risks.setText("Number of risks\n\n")

        self.risks_value = QtWidgets.QLabel(self.risks)
        self.risks_value.setGeometry(QtCore.QRect(15, 40, 240, 35))
        font = QtGui.QFont()
        font.setPointSize(13)
        font.setBold(True)
        font.setWeight(75)
        self.risks_value.setFont(font)
        self.risks_value.setAlignment(QtCore.Qt.AlignCenter)
        self.risks_value.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.risks_value.setStyleSheet("\n"
                                       "background-color:rgb(160, 160, 160);")
        self.risks_value.setObjectName("risks_value")
        self.risks_value.setText(self.open_sheet("Risk ID"))

        self.business_units = QtWidgets.QLabel(self.frame_18)
        self.business_units.setGeometry(QtCore.QRect(1043, 40, 251, 91))
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(True)
        font.setWeight(75)
        self.business_units.setFont(font)
        self.business_units.setAcceptDrops(False)
        self.business_units.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.business_units.setStyleSheet(
            "background-color:rgb(160, 160, 160);")
        self.business_units.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.business_units.setFrameShadow(QtWidgets.QFrame.Plain)
        self.business_units.setAlignment(QtCore.Qt.AlignCenter)
        self.business_units.setObjectName("business_units")
        self.business_units.setText("Number of business units\n\n")

        self.business_units_value = QtWidgets.QLabel(self.business_units)
        self.business_units_value.setGeometry(QtCore.QRect(15, 40, 240, 35))
        font = QtGui.QFont()
        font.setPointSize(13)
        font.setBold(True)
        font.setWeight(75)
        self.business_units_value.setFont(font)
        self.business_units_value.setAlignment(QtCore.Qt.AlignCenter)
        self.business_units_value.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.business_units_value.setStyleSheet(
            "\n"
            "background-color:rgb(160, 160, 160);")
        self.business_units_value.setObjectName("business_units_value")
        self.business_units_value.setText(self.open_sheet("Business Unit"))

        self.verticalLayout_28 = QtWidgets.QVBoxLayout(self.business_units)
        self.verticalLayout_28.setContentsMargins(4, 4, 4, 4)
        self.verticalLayout_28.setSpacing(4)
        self.verticalLayout_28.setObjectName("verticalLayout_28")

        # self.horizontalLayout_11.addWidget(self.business_units)
        self.verticalLayout_23.addWidget(self.frame_18)
        self.frame_19 = QtWidgets.QFrame(self.frame_17)
        self.frame_19.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_19.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_19.setObjectName("frame_19")

        def genSankey(df, cat_cols=[], value_cols='', title='Sankey Diagram'):
            # maximum of 6 value cols -> 6 colors
            colorPalette = ['#FFD43B', '#646464', '#4B8BBE', '#306998']
            labelList = []
            colorNumList = []
            for catCol in cat_cols:
                labelListTemp = list(set(df[catCol].values))
                colorNumList.append(len(labelListTemp))
                labelList = labelList + labelListTemp

            # remove duplicates from labelList
            labelList = list(dict.fromkeys(labelList))

            # define colors based on number of levels
            colorList = []
            link_color = ["#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", "#E2CFCF", "#E7DDDC", "#D9C3B7", "#E6B426", "#DDCB76", "#FFF1A9", ]
            for idx, colorNum in enumerate(colorNumList):
                colorList = colorList + [colorPalette[idx]] * colorNum

            # transform df into a source-target pair
            for i in range(len(cat_cols) - 1):
                if i == 0:
                    sourceTargetDf = df[[
                        cat_cols[i], cat_cols[i + 1], value_cols
                    ]]
                    sourceTargetDf.columns = ['source', 'target', 'count']
                else:
                    tempDf = df[[cat_cols[i], cat_cols[i + 1], value_cols]]
                    tempDf.columns = ['source', 'target', 'count']
                    sourceTargetDf = pandas.concat([sourceTargetDf, tempDf])
                    sourceTargetDf = sourceTargetDf.groupby(
                        ['source', 'target']).agg({
                            'count': 'sum'
                        }).reset_index()

            # add index for source-target pair
            sourceTargetDf['sourceID'] = sourceTargetDf['source'].apply(
                lambda x: labelList.index(x))
            sourceTargetDf['targetID'] = sourceTargetDf['target'].apply(
                lambda x: labelList.index(x))

            # creating the sankey diagram
            data = dict(type='sankey',
                        node=dict(pad=15,
                                  thickness=10,
                                  line=dict(color="black", width=1),
                                  label=labelList,
                                  color=colorList),
                        link=dict(
                            source=sourceTargetDf['sourceID'],
                            target=sourceTargetDf['targetID'],
                            value=sourceTargetDf['count'],
                            color = link_color
                        ))

            layout = dict(title="", font=dict(size=10), paper_bgcolor="rgb(56, 56, 56)")

            fig = dict(data=[data], layout=layout)
            return fig

        fig = genSankey(source_excel,
                        cat_cols=['Process Domain', 'Risk ID', 'Control ID'],
                        value_cols='count_data',
                        title='Sankey Diagram(Process VS Risk VS Control)')

        html = '<html><body>'
        html += plotly.offline.plot(fig,
                                    output_type='div',
                                    include_plotlyjs='cdn')
        html += '</body></html>'

        self.webEngineView = QtWebEngineWidgets.QWebEngineView(self.frame_18)
        self.webEngineView.setHtml(html)
        self.verticalLayout_23.addWidget(self.webEngineView)

        self.verticalLayout_30 = QtWidgets.QVBoxLayout(self.frame_19)
        self.verticalLayout_30.setContentsMargins(4, 4, 4, 4)
        self.verticalLayout_30.setSpacing(4)
        self.verticalLayout_30.setObjectName("verticalLayout_30")

        self.horizontalLayout_17 = QtWidgets.QHBoxLayout(self.frame_19)
        self.horizontalLayout_17.setContentsMargins(50, 4, 50, 4)
        self.horizontalLayout_17.setSpacing(50)
        self.horizontalLayout_17.setObjectName("horizontalLayout_17")
        self.horizontalLayout_17.addWidget(self.frame_19)
        self.verticalLayout_23.addWidget(self.frame_19)
        self.verticalLayout_23.setStretch(0, 4)
        self.verticalLayout_23.setStretch(1, 7)
        self.verticalLayout_20.addWidget(self.frame_17)
        self.verticalLayout_22.addLayout(self.verticalLayout_20)
        self.stackedWidget.addWidget(self.PRC_stat)
        self.Control_stat = QtWidgets.QWidget()
        self.Control_stat.setObjectName("Control_stat")
        self.verticalLayout_32 = QtWidgets.QVBoxLayout(self.Control_stat)
        self.verticalLayout_32.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_32.setSpacing(0)
        self.verticalLayout_32.setObjectName("verticalLayout_32")
        self.verticalLayout_31 = QtWidgets.QVBoxLayout()
        self.verticalLayout_31.setObjectName("verticalLayout_31")
        self.frame_27 = QtWidgets.QFrame(self.Control_stat)
        self.frame_27.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_27.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_27.setObjectName("frame_27")
        self.verticalLayout_33 = QtWidgets.QVBoxLayout(self.frame_27)
        self.verticalLayout_33.setObjectName("verticalLayout_33")
        self.frame_29 = QtWidgets.QFrame(self.frame_27)
        self.frame_29.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_29.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_29.setObjectName("frame_29")
        self.horizontalLayout_19 = QtWidgets.QHBoxLayout(self.frame_29)
        self.horizontalLayout_19.setObjectName("horizontalLayout_19")
        self.frame_30 = QtWidgets.QFrame(self.frame_29)
        self.frame_30.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_30.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_30.setObjectName("frame_30")
        self.verticalLayout_38 = QtWidgets.QVBoxLayout(self.frame_30)
        self.verticalLayout_38.setObjectName("verticalLayout_38")
        self.label_21 = QtWidgets.QLabel(self.frame_30)
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_21.setFont(font)
        self.label_21.setStyleSheet("color:white")
        self.label_21.setObjectName("label_21")
        self.verticalLayout_38.addWidget(self.label_21)
        self.Stat_Control_1_Chart = QtWidgets.QWidget(self.frame_30)
        self.Stat_Control_1_Chart.setStyleSheet("background-color:white")
        self.Stat_Control_1_Chart.setObjectName("Stat_Control_1_Chart")
        self.verticalLayout_38.addWidget(self.Stat_Control_1_Chart)
        self.verticalLayout_38.setStretch(0, 3)
        self.verticalLayout_38.setStretch(1, 24)
        self.horizontalLayout_19.addWidget(self.frame_30)
        self.frame_31 = QtWidgets.QFrame(self.frame_29)
        self.frame_31.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_31.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_31.setObjectName("frame_31")
        self.verticalLayout_37 = QtWidgets.QVBoxLayout(self.frame_31)
        self.verticalLayout_37.setObjectName("verticalLayout_37")
        self.label_22 = QtWidgets.QLabel(self.frame_31)
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_22.setFont(font)
        self.label_22.setStyleSheet("color:white")
        self.label_22.setObjectName("label_22")
        self.verticalLayout_37.addWidget(self.label_22)
        self.Stat_Control_2_Chart = QtWidgets.QWidget(self.frame_31)
        self.Stat_Control_2_Chart.setStyleSheet("background-color:white")
        self.Stat_Control_2_Chart.setObjectName("Stat_Control_2_Chart")
        self.verticalLayout_37.addWidget(self.Stat_Control_2_Chart)
        self.verticalLayout_37.setStretch(0, 3)
        self.verticalLayout_37.setStretch(1, 24)
        self.horizontalLayout_19.addWidget(self.frame_31)
        self.frame_32 = QtWidgets.QFrame(self.frame_29)
        self.frame_32.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_32.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_32.setObjectName("frame_32")
        self.verticalLayout_36 = QtWidgets.QVBoxLayout(self.frame_32)
        self.verticalLayout_36.setObjectName("verticalLayout_36")
        self.label_23 = QtWidgets.QLabel(self.frame_32)
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_23.setFont(font)
        self.label_23.setStyleSheet("color:white")
        self.label_23.setObjectName("label_23")
        self.verticalLayout_36.addWidget(self.label_23)
        self.Stat_Control_3_Chart = QtWidgets.QWidget(self.frame_32)
        self.Stat_Control_3_Chart.setStyleSheet("background-color:white")
        self.Stat_Control_3_Chart.setObjectName("Stat_Control_3_Chart")
        self.verticalLayout_36.addWidget(self.Stat_Control_3_Chart)
        self.verticalLayout_36.setStretch(0, 3)
        self.verticalLayout_36.setStretch(1, 24)
        self.horizontalLayout_19.addWidget(self.frame_32)
        self.frame_33 = QtWidgets.QFrame(self.frame_29)
        self.frame_33.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_33.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_33.setObjectName("frame_33")
        self.verticalLayout_34 = QtWidgets.QVBoxLayout(self.frame_33)
        self.verticalLayout_34.setObjectName("verticalLayout_34")
        self.label_25 = QtWidgets.QLabel(self.frame_33)
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_25.setFont(font)
        self.label_25.setStyleSheet("color:white")
        self.label_25.setObjectName("label_25")
        self.verticalLayout_34.addWidget(self.label_25)
        self.Stat_Control_4_Chart = QtWidgets.QWidget(self.frame_33)
        self.Stat_Control_4_Chart.setStyleSheet("background-color:white")
        self.Stat_Control_4_Chart.setObjectName("Stat_Control_4_Chart")
        self.verticalLayout_34.addWidget(self.Stat_Control_4_Chart)
        self.verticalLayout_34.setStretch(0, 3)
        self.verticalLayout_34.setStretch(1, 24)
        self.horizontalLayout_19.addWidget(self.frame_33)
        self.verticalLayout_33.addWidget(self.frame_29)
        self.frame_28 = QtWidgets.QFrame(self.frame_27)
        self.frame_28.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_28.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_28.setObjectName("frame_28")
        self.horizontalLayout_18 = QtWidgets.QHBoxLayout(self.frame_28)
        self.horizontalLayout_18.setObjectName("horizontalLayout_18")
        self.frame_34 = QtWidgets.QFrame(self.frame_28)
        self.frame_34.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_34.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_34.setObjectName("frame_34")
        self.verticalLayout_35 = QtWidgets.QVBoxLayout(self.frame_34)
        self.verticalLayout_35.setObjectName("verticalLayout_35")
        self.label_29 = QtWidgets.QLabel(self.frame_34)
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_29.setFont(font)
        self.label_29.setStyleSheet("color:white")
        self.label_29.setObjectName("label_29")
        self.verticalLayout_35.addWidget(self.label_29)
        self.Stat_Control_5_Chart = QtWidgets.QWidget(self.frame_34)
        self.Stat_Control_5_Chart.setStyleSheet("background-color:white")
        self.Stat_Control_5_Chart.setObjectName("Stat_Control_5_Chart")
        self.verticalLayout_35.addWidget(self.Stat_Control_5_Chart)
        self.verticalLayout_35.setStretch(0, 1)
        self.verticalLayout_35.setStretch(1, 20)
        self.horizontalLayout_18.addWidget(self.frame_34)
        self.frame_35 = QtWidgets.QFrame(self.frame_28)
        self.frame_35.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_35.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_35.setObjectName("frame_35")
        self.verticalLayout_41 = QtWidgets.QVBoxLayout(self.frame_35)
        self.verticalLayout_41.setObjectName("verticalLayout_41")
        self.label_28 = QtWidgets.QLabel(self.frame_35)
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_28.setFont(font)
        self.label_28.setStyleSheet("color:white")
        self.label_28.setObjectName("label_28")
        self.verticalLayout_41.addWidget(self.label_28)
        self.Stat_Control_6_Chart = QtWidgets.QWidget(self.frame_35)
        self.Stat_Control_6_Chart.setStyleSheet("background-color:white")
        self.Stat_Control_6_Chart.setObjectName("Stat_Control_6_Chart")
        self.verticalLayout_41.addWidget(self.Stat_Control_6_Chart)
        self.verticalLayout_41.setStretch(0, 1)
        self.verticalLayout_41.setStretch(1, 20)
        self.horizontalLayout_18.addWidget(self.frame_35)
        self.frame_36 = QtWidgets.QFrame(self.frame_28)
        self.frame_36.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_36.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_36.setObjectName("frame_36")
        self.verticalLayout_40 = QtWidgets.QVBoxLayout(self.frame_36)
        self.verticalLayout_40.setObjectName("verticalLayout_40")
        self.label_27 = QtWidgets.QLabel(self.frame_36)
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_27.setFont(font)
        self.label_27.setStyleSheet("color:white")
        self.label_27.setObjectName("label_27")
        self.verticalLayout_40.addWidget(self.label_27)
        self.Stat_Control_7_Chart = QtWidgets.QWidget(self.frame_36)
        self.Stat_Control_7_Chart.setStyleSheet("background-color:white")
        self.Stat_Control_7_Chart.setObjectName("Stat_Control_7_Chart")
        self.verticalLayout_40.addWidget(self.Stat_Control_7_Chart)
        self.verticalLayout_40.setStretch(0, 1)
        self.verticalLayout_40.setStretch(1, 20)
        self.horizontalLayout_18.addWidget(self.frame_36)
        self.frame_37 = QtWidgets.QFrame(self.frame_28)
        self.frame_37.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_37.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_37.setObjectName("frame_37")
        self.verticalLayout_39 = QtWidgets.QVBoxLayout(self.frame_37)
        self.verticalLayout_39.setObjectName("verticalLayout_39")
        self.label_26 = QtWidgets.QLabel(self.frame_37)
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_26.setFont(font)
        self.label_26.setStyleSheet("color:white")
        self.label_26.setObjectName("label_26")
        self.verticalLayout_39.addWidget(self.label_26)
        self.Stat_Control_8_Chart = QtWidgets.QWidget(self.frame_37)
        self.Stat_Control_8_Chart.setStyleSheet("background-color:white")
        self.Stat_Control_8_Chart.setObjectName("Stat_Control_8_Chart")
        self.verticalLayout_39.addWidget(self.Stat_Control_8_Chart)
        self.verticalLayout_39.setStretch(0, 3)
        self.verticalLayout_39.setStretch(1, 24)
        self.horizontalLayout_18.addWidget(self.frame_37)
        self.verticalLayout_33.addWidget(self.frame_28)
        self.verticalLayout_33.setStretch(0, 1)
        self.verticalLayout_33.setStretch(1, 1)
        self.verticalLayout_31.addWidget(self.frame_27)
        self.verticalLayout_32.addLayout(self.verticalLayout_31)
        self.stackedWidget.addWidget(self.Control_stat)
        self.Risk_stat = QtWidgets.QWidget()
        self.Risk_stat.setObjectName("Risk_stat")
        self.horizontalLayout_21 = QtWidgets.QHBoxLayout(self.Risk_stat)
        self.horizontalLayout_21.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_21.setSpacing(0)
        self.horizontalLayout_21.setObjectName("horizontalLayout_21")
        self.horizontalLayout_20 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_20.setSpacing(50)
        self.horizontalLayout_20.setObjectName("horizontalLayout_20")
        self.frame_38 = QtWidgets.QFrame(self.Risk_stat)
        self.frame_38.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_38.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_38.setObjectName("frame_38")
        self.horizontalLayout_22 = QtWidgets.QHBoxLayout(self.frame_38)
        self.horizontalLayout_22.setContentsMargins(30, 0, 30, 0)
        self.horizontalLayout_22.setSpacing(40)
        self.horizontalLayout_22.setObjectName("horizontalLayout_22")
        self.frame_41 = QtWidgets.QFrame(self.frame_38)
        self.frame_41.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_41.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_41.setObjectName("frame_41")
        self.verticalLayout_42 = QtWidgets.QVBoxLayout(self.frame_41)
        self.verticalLayout_42.setSpacing(10)
        self.verticalLayout_42.setObjectName("verticalLayout_42")
        self.label_30 = QtWidgets.QLabel(self.frame_41)
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_30.setFont(font)
        self.label_30.setStyleSheet("color:white")
        self.label_30.setObjectName("label_30")
        self.verticalLayout_42.addWidget(self.label_30)
        self.Stat_Risk_Chart_1_1 = QtWidgets.QWidget(self.frame_41)
        self.Stat_Risk_Chart_1_1.setStyleSheet("background-color:white\n" "")
        self.Stat_Risk_Chart_1_1.setObjectName("Stat_Risk_Chart_1_1")
        self.verticalLayout_42.addWidget(self.Stat_Risk_Chart_1_1)
        self.Stat_Risk_Chart_1_2 = QtWidgets.QWidget(self.frame_41)
        self.Stat_Risk_Chart_1_2.setStyleSheet("background-color:white\n" "")
        self.Stat_Risk_Chart_1_2.setObjectName("Stat_Risk_Chart_1_2")
        self.verticalLayout_42.addWidget(self.Stat_Risk_Chart_1_2)
        self.verticalLayout_42.setStretch(0, 2)
        self.verticalLayout_42.setStretch(1, 10)
        self.verticalLayout_42.setStretch(2, 16)
        self.horizontalLayout_22.addWidget(self.frame_41)
        self.frame_40 = QtWidgets.QFrame(self.frame_38)
        self.frame_40.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_40.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_40.setObjectName("frame_40")
        self.verticalLayout_43 = QtWidgets.QVBoxLayout(self.frame_40)
        self.verticalLayout_43.setSpacing(10)
        self.verticalLayout_43.setObjectName("verticalLayout_43")
        self.label_31 = QtWidgets.QLabel(self.frame_40)
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_31.setFont(font)
        self.label_31.setStyleSheet("color:white")
        self.label_31.setObjectName("label_31")
        self.verticalLayout_43.addWidget(self.label_31)
        self.Stat_Risk_Chart_2_1 = QtWidgets.QWidget(self.frame_40)
        self.Stat_Risk_Chart_2_1.setStyleSheet("background-color:white\n" "")
        self.Stat_Risk_Chart_2_1.setObjectName("Stat_Risk_Chart_2_1")
        self.verticalLayout_43.addWidget(self.Stat_Risk_Chart_2_1)
        self.Stat_Risk_Chart_2_2 = QtWidgets.QWidget(self.frame_40)
        self.Stat_Risk_Chart_2_2.setStyleSheet("background-color:white\n" "")
        self.Stat_Risk_Chart_2_2.setObjectName("Stat_Risk_Chart_2_2")
        self.verticalLayout_43.addWidget(self.Stat_Risk_Chart_2_2)
        self.verticalLayout_43.setStretch(0, 2)
        self.verticalLayout_43.setStretch(1, 10)
        self.verticalLayout_43.setStretch(2, 16)
        self.horizontalLayout_22.addWidget(self.frame_40)
        self.frame_39 = QtWidgets.QFrame(self.frame_38)
        self.frame_39.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_39.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_39.setObjectName("frame_39")
        self.verticalLayout_44 = QtWidgets.QVBoxLayout(self.frame_39)
        self.verticalLayout_44.setSpacing(10)
        self.verticalLayout_44.setObjectName("verticalLayout_44")
        self.label_32 = QtWidgets.QLabel(self.frame_39)
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_32.setFont(font)
        self.label_32.setStyleSheet("color:white")
        self.label_32.setObjectName("label_32")
        self.verticalLayout_44.addWidget(self.label_32)
        self.Stat_Risk_Chart_3_1 = QtWidgets.QWidget(self.frame_39)
        self.Stat_Risk_Chart_3_1.setStyleSheet("background-color:white\n" "")
        self.Stat_Risk_Chart_3_1.setObjectName("Stat_Risk_Chart_3_1")
        self.verticalLayout_44.addWidget(self.Stat_Risk_Chart_3_1)
        self.Stat_Risk_Chart_3_2 = QtWidgets.QWidget(self.frame_39)
        self.Stat_Risk_Chart_3_2.setStyleSheet("background-color:white\n" "")
        self.Stat_Risk_Chart_3_2.setObjectName("Stat_Risk_Chart_3_2")
        self.verticalLayout_44.addWidget(self.Stat_Risk_Chart_3_2)
        self.verticalLayout_44.setStretch(0, 2)
        self.verticalLayout_44.setStretch(1, 10)
        self.verticalLayout_44.setStretch(2, 16)
        self.horizontalLayout_22.addWidget(self.frame_39)
        self.horizontalLayout_20.addWidget(self.frame_38)
        self.horizontalLayout_21.addLayout(self.horizontalLayout_20)
        self.stackedWidget.addWidget(self.Risk_stat)
        self.Process_stat = QtWidgets.QWidget()
        self.Process_stat.setObjectName("Process_stat")
        self.horizontalLayout_24 = QtWidgets.QHBoxLayout(self.Process_stat)
        self.horizontalLayout_24.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_24.setSpacing(0)
        self.horizontalLayout_24.setObjectName("horizontalLayout_24")
        self.horizontalLayout_23 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_23.setObjectName("horizontalLayout_23")
        self.frame_42 = QtWidgets.QFrame(self.Process_stat)
        self.frame_42.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_42.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_42.setObjectName("frame_42")
        self.horizontalLayout_25 = QtWidgets.QHBoxLayout(self.frame_42)
        self.horizontalLayout_25.setObjectName("horizontalLayout_25")
        self.frame_46 = QtWidgets.QFrame(self.frame_42)
        self.frame_46.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_46.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_46.setObjectName("frame_46")
        self.verticalLayout_45 = QtWidgets.QVBoxLayout(self.frame_46)
        self.verticalLayout_45.setObjectName("verticalLayout_45")
        self.label_36 = QtWidgets.QLabel(self.frame_46)
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_36.setFont(font)
        self.label_36.setStyleSheet("color:white")
        self.label_36.setObjectName("label_36")
        self.verticalLayout_45.addWidget(self.label_36)
        self.Stat_Process_Chart_1_1 = QtWidgets.QWidget(self.frame_46)
        self.Stat_Process_Chart_1_1.setStyleSheet("background-color:white\n"
                                                  "")
        self.Stat_Process_Chart_1_1.setObjectName("Stat_Process_Chart_1_1")
        self.verticalLayout_45.addWidget(self.Stat_Process_Chart_1_1)
        self.Stat_Process_Chart_1_2 = QtWidgets.QWidget(self.frame_46)
        self.Stat_Process_Chart_1_2.setStyleSheet("background-color:white\n"
                                                  "")
        self.Stat_Process_Chart_1_2.setObjectName("Stat_Process_Chart_1_2")
        self.verticalLayout_45.addWidget(self.Stat_Process_Chart_1_2)
        self.verticalLayout_45.setStretch(0, 2)
        self.verticalLayout_45.setStretch(1, 10)
        self.verticalLayout_45.setStretch(2, 16)
        self.horizontalLayout_25.addWidget(self.frame_46)
        self.frame_45 = QtWidgets.QFrame(self.frame_42)
        self.frame_45.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_45.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_45.setObjectName("frame_45")
        self.verticalLayout_46 = QtWidgets.QVBoxLayout(self.frame_45)
        self.verticalLayout_46.setObjectName("verticalLayout_46")
        self.label_34 = QtWidgets.QLabel(self.frame_45)
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_34.setFont(font)
        self.label_34.setStyleSheet("color:white")
        self.label_34.setObjectName("label_34")
        self.verticalLayout_46.addWidget(self.label_34)
        self.Stat_Process_Chart_2_1 = QtWidgets.QWidget(self.frame_45)
        self.Stat_Process_Chart_2_1.setStyleSheet("background-color:white\n"
                                                  "")
        self.Stat_Process_Chart_2_1.setObjectName("Stat_Process_Chart_2_1")
        self.verticalLayout_46.addWidget(self.Stat_Process_Chart_2_1)
        self.Stat_Process_Chart_2_2 = QtWidgets.QWidget(self.frame_45)
        self.Stat_Process_Chart_2_2.setStyleSheet("background-color:white\n"
                                                  "")
        self.Stat_Process_Chart_2_2.setObjectName("Stat_Process_Chart_2_2")
        self.verticalLayout_46.addWidget(self.Stat_Process_Chart_2_2)
        self.verticalLayout_46.setStretch(0, 2)
        self.verticalLayout_46.setStretch(1, 10)
        self.verticalLayout_46.setStretch(2, 16)
        self.horizontalLayout_25.addWidget(self.frame_45)
        self.frame_43 = QtWidgets.QFrame(self.frame_42)
        self.frame_43.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_43.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_43.setObjectName("frame_43")
        self.verticalLayout_48 = QtWidgets.QVBoxLayout(self.frame_43)
        self.verticalLayout_48.setObjectName("verticalLayout_48")
        self.label_35 = QtWidgets.QLabel(self.frame_43)
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_35.setFont(font)
        self.label_35.setStyleSheet("color:white")
        self.label_35.setObjectName("label_35")
        self.verticalLayout_48.addWidget(self.label_35)
        self.Stat_Process_Chart_3_1 = QtWidgets.QWidget(self.frame_43)
        self.Stat_Process_Chart_3_1.setStyleSheet("background-color:white\n"
                                                  "")
        self.Stat_Process_Chart_3_1.setObjectName("Stat_Process_Chart_3_1")
        self.verticalLayout_48.addWidget(self.Stat_Process_Chart_3_1)
        self.Stat_Process_Chart_3_2 = QtWidgets.QWidget(self.frame_43)
        self.Stat_Process_Chart_3_2.setStyleSheet("background-color:white\n"
                                                  "")
        self.Stat_Process_Chart_3_2.setObjectName("Stat_Process_Chart_3_2")
        self.verticalLayout_48.addWidget(self.Stat_Process_Chart_3_2)
        self.verticalLayout_48.setStretch(0, 2)
        self.verticalLayout_48.setStretch(1, 10)
        self.verticalLayout_48.setStretch(2, 16)
        self.horizontalLayout_25.addWidget(self.frame_43)
        self.frame_44 = QtWidgets.QFrame(self.frame_42)
        self.frame_44.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_44.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_44.setObjectName("frame_44")
        self.verticalLayout_47 = QtWidgets.QVBoxLayout(self.frame_44)
        self.verticalLayout_47.setObjectName("verticalLayout_47")
        self.label_37 = QtWidgets.QLabel(self.frame_44)
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        self.label_37.setFont(font)
        self.label_37.setStyleSheet("color:white")
        self.label_37.setObjectName("label_37")
        self.verticalLayout_47.addWidget(self.label_37)
        self.Stat_Process_Chart_4_1 = QtWidgets.QWidget(self.frame_44)
        self.Stat_Process_Chart_4_1.setStyleSheet("background-color:white\n"
                                                  "")
        self.Stat_Process_Chart_4_1.setObjectName("Stat_Process_Chart_4_1")
        self.verticalLayout_47.addWidget(self.Stat_Process_Chart_4_1)
        self.Stat_Process_Chart_4_2 = QtWidgets.QWidget(self.frame_44)
        self.Stat_Process_Chart_4_2.setStyleSheet("background-color:white\n"
                                                  "")
        self.Stat_Process_Chart_4_2.setObjectName("Stat_Process_Chart_4_2")
        self.verticalLayout_47.addWidget(self.Stat_Process_Chart_4_2)
        self.verticalLayout_47.setStretch(0, 2)
        self.verticalLayout_47.setStretch(1, 10)
        self.verticalLayout_47.setStretch(2, 16)
        self.horizontalLayout_25.addWidget(self.frame_44)
        self.horizontalLayout_23.addWidget(self.frame_42)
        self.horizontalLayout_24.addLayout(self.horizontalLayout_23)
        self.stackedWidget.addWidget(self.Process_stat)
        self.gridLayout.addWidget(self.stackedWidget, 0, 0, 1, 1)
        self.horizontalLayout.addWidget(self.processframe)
        self.horizontalLayout.setStretch(0, 3)
        self.horizontalLayout.setStretch(1, 10)
        self.verticalLayout_3.addWidget(self.middleframe)
        self.topframe = QtWidgets.QFrame(self.mainframe)
        self.topframe.setFrameShape(QtWidgets.QFrame.WinPanel)
        self.topframe.setFrameShadow(QtWidgets.QFrame.Raised)
        self.topframe.setObjectName("topframe")
        self.verticalLayout_3.addWidget(self.topframe)
        self.verticalLayout_3.setStretch(0, 1)
        self.verticalLayout_3.setStretch(1, 40)
        self.verticalLayout_3.setStretch(2, 1)
        self.verticalLayout.addWidget(self.mainframe)
        self.verticalLayout_2.addLayout(self.verticalLayout)

        self.retranslateUi(Form)
        self.PPCategory_listWidget.item(0).setSelected(True)
        self.PPCategory_listWidget.setEnabled(True)
        ###################My Fixed######################
        ################################
        self.OtherCategory_listWidget.setEnabled(True)
        self.connect2allwidgetlistener()
        #################################################
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.label_2.setText(_translate("Form", "Choose category"))
        __sortingEnabled = self.PPCategory_listWidget.isSortingEnabled()
        self.PPCategory_listWidget.setSortingEnabled(False)
        item = self.PPCategory_listWidget.item(0)
        item.setText(_translate("Form", "PP Data Extractor"))
        item = self.PPCategory_listWidget.item(1)
        item.setText(_translate("Form", "RCM Data EXtractor"))
        self.PPCategory_listWidget.setSortingEnabled(__sortingEnabled)
        __sortingEnabled = self.OtherCategory_listWidget.isSortingEnabled()
        self.OtherCategory_listWidget.setSortingEnabled(False)
        self.OtherCategory_listWidget.setSortingEnabled(__sortingEnabled)
        item = self.OtherCategory_listWidget.item(0)
        item.setText(_translate("Form", "RCM Aggregator Stat"))
        item = self.OtherCategory_listWidget.item(1)
        item.setText(_translate("Form", "Combined PRC Stat"))
        item = self.OtherCategory_listWidget.item(2)
        item.setText(_translate("Form", "Control Statistics"))
        item = self.OtherCategory_listWidget.item(3)
        item.setText(_translate("Form", "Risk Statistics"))
        item = self.OtherCategory_listWidget.item(4)
        item.setText(_translate("Form", "Process Statistics"))
        self.OtherCategory_listWidget.setSortingEnabled(__sortingEnabled)
        self.label_3.setText(_translate("Form", "Input Path"))
        self.PPInput_pushButton.setText(_translate("Form", "Browser"))
        self.Reset_pushButton.setText("Reset")
        self.label_4.setText(_translate("Form", "Output Path"))
        self.PPOutput_pushButton.setText(_translate("Form", "Browser"))
        self.label_12.setText(_translate("Form", "Output Format:"))
        self.PPXML_CheckBox_2.setText(_translate("Form", "XML"))
        self.PPJSON_CheckBox_2.setText(_translate("Form", "JSON"))
        self.PPConvert_pushButton_2.setText(_translate("Form", "Convert"))
        self.label_13.setText(_translate("Form", "Output File(s)"))
        self.label_14.setText(_translate("Form", "File Content"))
        self.label_15.setText(_translate("Form", "Input Path"))
        self.RCMInput_pushButton_2.setText(_translate("Form", "Browse"))
        self.label_24.setText(_translate("Form", "Output Path"))
        self.RCMOutput_pushButton_4.setText(_translate("Form", "Browse"))
        self.RCMConvert_pushButton_2.setText(_translate("Form", "Convert"))
        self.label_16.setText(_translate("Form", "Combined File Content"))
        self.label.setText(
            _translate("Form", "Total Number of RCM file(s) combined"))
        self.Stat_Total_Number_Label.setText(_translate("Form", "0"))
        self.label_6.setText(_translate("Form", "RCM Names"))
        __sortingEnabled = self.Stat_RCM_Names_listWidget.isSortingEnabled()
        self.Stat_RCM_Names_listWidget.setSortingEnabled(False)
        self.Stat_RCM_Names_listWidget.setSortingEnabled(__sortingEnabled)
        self.label_7.setText(_translate("Form", "Combined RCM Location"))
        self.label_8.setText(_translate("Form",
                                        "Combined RCM Missing Columns"))
        self.label_9.setText(
            _translate("Form", "Combined RCM(Populated / Missing Data)"))
        self.label_21.setText(
            _translate("Form", "Controls Associated to Process"))
        self.label_22.setText(_translate("Form",
                                         "Controls Associated to Risk"))
        self.label_23.setText(_translate("Form", "Controls Per Control Owner"))
        self.label_25.setText(_translate("Form", "Controls Per Business Unit"))
        self.label_29.setText(
            _translate(
                "Form", "Controls per Control type (Manual, \n"
                "Semi-automated, Automated)"))
        self.label_28.setText(
            _translate(
                "Form", "Controls per Control \n"
                "type (Preventive , Detective) "))
        self.label_27.setText(
            _translate("Form", "Controls per Control \n"
                       "frequency type"))
        self.label_26.setText(_translate("Form", "Key and non Key controls"))
        self.label_30.setText(
            _translate("Form", "Risks Associated to Each Process"))
        self.label_31.setText(
            _translate("Form", "Risks Associated to Each Control"))
        self.label_32.setText(
            _translate("Form", "Risks Associated to Each Busines Domain"))
        self.label_36.setText(_translate("Form", "Process By Process Domain"))
        self.label_34.setText(_translate("Form", "Process By Business Unit"))
        self.label_35.setText(_translate("Form", "Control Per Process Domain"))
        self.label_37.setText(_translate("Form", "Risk Per Process Domain"))

    def inputFilePP(self):
        global input_folder_path
        filename = QtWidgets.QFileDialog.getExistingDirectory()
        self.PPInput_lineEdit.setText(filename)
        self.PP_progressBar_2.setValue(0)
        input_folder_path = filename

    def inputFileInRCM(self):
        global input_folder_path_rcm
        filename = QtWidgets.QFileDialog.getExistingDirectory()
        self.RCMInput_lineEdi_2.setText(filename)
        self.RCM_progressBar_3.setValue(0)
        input_folder_path_rcm = filename

    def outputFilePP(self):
        global output_folder_path
        filename = QtWidgets.QFileDialog.getExistingDirectory()
        self.PPOutput_lineEdit.setText(filename)
        output_folder_path = filename

    def outputFileRCM(self):
        global output_folder_path_rcm
        filename = QtWidgets.QFileDialog.getExistingDirectory()
        self.RCMOutput_LineText.setText(filename)
        output_folder_path_rcm = filename

    def convertPP(self):
        global output_type
        global cwd
        global input_folder_path
        global output_folder_path
        global output_file_path
        resfiles = []
        self.PPFile_listWidget_2.clear()

        cwd = os.getcwd()
        os.chdir(input_folder_path)
        isSuccess = True
        self.PP_progressBar_2.setMinimum(0)
        self.PP_progressBar_2.setValue(0)
        self.PP_progressBar_2.setMaximum(len(os.listdir()))

        for input_file_path in os.listdir():
            if os.path.exists(input_folder_path):
                os.chdir(input_folder_path)

            if os.path.isfile(input_file_path) and os.access(
                    input_file_path, os.R_OK):
                # find the document type
                name, extension = os.path.splitext(input_file_path)
                if extension == ".docx":
                    # get file name
                    head, tail = os.path.split(input_file_path)

                    output_file_path = tail.split(".")[0] + output_type
                    # extract data from doc
                    data = my.read_docs(input_file_path)
                    data = my.applyruledoc(data)
                    if output_type in [".JSON", ".json"]:
                        # save as json
                        if not os.path.exists(output_folder_path):
                            os.makedirs(output_folder_path)
                        os.chdir(output_folder_path)
                        state = my.save_json(data, output_file_path)
                        self.PPFile_listWidget_2.addItem(
                            QListWidgetItem(output_file_path))
                        os.chdir(cwd)
                    elif output_type in [".XML", ".xml"]:
                        # save as xml
                        if not os.path.exists(output_folder_path):
                            os.makedirs(output_folder_path)
                        os.chdir(output_folder_path)
                        state = my.save_xml(data, output_file_path)
                        self.PPFile_listWidget_2.addItem(
                            QListWidgetItem(output_file_path))
                        os.chdir(cwd)
                    else:
                        print("INVALID OUTPUT FORMAT")
                    # print("SUCCESS")
                elif extension == ".pdf":
                    # get file name
                    head, tail = os.path.split(input_file_path)

                    output_file_path = tail.split(".")[0] + output_type

                    # extract pdf data
                    data = my.pdf_extracter(input_file_path)

                    data = my.applyrulepdf(data)

                    if (data == []):
                        data = my.cleannrtdata(
                            my.split_into_sentences(
                                str(my.nrtpdfextract(input_file_path))))

                        if output_type in [".JSON", ".json"]:
                            # save as json
                            if not os.path.exists(output_folder_path):
                                os.makedirs(output_folder_path)
                            os.chdir(output_folder_path)
                            state = my.save_json(data, output_file_path)
                            self.PPFile_listWidget_2.addItem(
                                QListWidgetItem(output_file_path))
                            os.chdir(cwd)
                        elif output_type in [".XML", ".xml"]:
                            # save as xml
                            if not os.path.exists(output_folder_path):
                                os.makedirs(output_folder_path)
                            os.chdir(output_folder_path)
                            state = my.save_xml(data, output_file_path)
                            self.PPFile_listWidget_2.addItem(
                                QListWidgetItem(output_file_path))
                            os.chdir(cwd)
                        else:
                            print("INVALID OUTPUT FORMAT")
                        # print("SUCCESS")
                    else:
                        if output_type in [".JSON", ".json"]:
                            # save as json
                            if not os.path.exists(output_folder_path):
                                os.makedirs(output_folder_path)
                            os.chdir(output_folder_path)
                            state = my.save_json(data, output_file_path)
                            self.PPFile_listWidget_2.addItem(
                                QListWidgetItem(output_file_path))
                            os.chdir(cwd)
                        elif output_type in [".XML", ".xml"]:
                            # save as xml
                            if not os.path.exists(output_folder_path):
                                os.makedirs(output_folder_path)
                            os.chdir(output_folder_path)
                            state = my.save_xml(data, output_file_path)
                            self.PPFile_listWidget_2.addItem(
                                QListWidgetItem(output_file_path))
                            os.chdir(cwd)
                        else:
                            print("INVALID OUTPUT FORMAT")
                        # print("SUCCESS")
                elif extension == ".csv":
                    head, tail = os.path.split(input_file_path)
                    if output_type in [".JSON", '.json']:
                        csvFilePath = input_file_path
                        jsonFilePath = tail.split('.')[0] + output_type
                        if not os.path.exists(output_folder_path):
                            os.makedirs(output_folder_path)

                        my.make_json_csv(csvFilePath, jsonFilePath,
                                         output_folder_path)
                        self.PPFile_listWidget_2.addItem(
                            QListWidgetItem(jsonFilePath))
                        os.chdir(cwd)
                        # print("SUCCESS")
                    elif output_type in [".XML", ".xml"]:
                        csvFilePath = input_file_path
                        jsonFilePath = tail.split('.')[0] + output_type
                        if not os.path.exists(output_folder_path):
                            os.makedirs(output_folder_path)

                        my.make_json_xml(csvFilePath, jsonFilePath,
                                         output_folder_path)
                        self.PPFile_listWidget_2.addItem(
                            QListWidgetItem(jsonFilePath))
                        os.chdir(cwd)
                    else:
                        print("INVALID OUTPUT FORMAT")
                elif extension == ".xlsx":
                    head, tail = os.path.split(input_file_path)
                    if output_type in [".JSON", '.json']:
                        xlsxfilepath = input_file_path
                        resfiles = []
                        jsonFilePath = tail.split('.')[0]
                        if not os.path.exists(output_folder_path):
                            os.makedirs(output_folder_path)

                        my.xlsx_to_json(xlsxfilepath, jsonFilePath,
                                        output_folder_path, resfiles)
                        for resfname in resfiles:
                            self.PPFile_listWidget_2.addItem(
                                QListWidgetItem(resfname))

                        os.chdir(cwd)
                        # print("SUCCESS")
                    elif output_type in [".XML", ".xml"]:
                        xlsxfilepath = input_file_path
                        resfiles = []
                        xmlfilepath = tail.split('.')[0]
                        if not os.path.exists(output_folder_path):
                            os.makedirs(output_folder_path)

                        my.xlsx_to_xml(xlsxfilepath, xmlfilepath,
                                       output_folder_path, resfiles)
                        for resfname in resfiles:
                            self.PPFile_listWidget_2.addItem(
                                QListWidgetItem(resfname))
                        os.chdir(cwd)
                    else:
                        print("INVALID OUTPUT FORMAT")
                else:
                    print("Only accept docx, pdf, excel and csv")
            else:
                print("")
            if isSuccess:
                self.PP_progressBar_2.setValue(self.PP_progressBar_2.value() +
                                               1)

    def convertInRCM(self):

        global input_folder_path_rcm
        global output_folder_path_rcm
        global cwd
        origalfiles = []
        output_filecontent = [{
            'Process Domain': [],
            'Process ID': [],
            'Process Description': [],
            'Risk ID': [],
            'Risk Description': [],
            'Control ID': [],
            'Control Objective': [],
            'Control Description': [],
            'Control Owner': [],
            'Business Unit': [],
            'Control Type (Preventive , Detective)': [],
            'Control Type (Manual, Semi-automated, Automated)': [],
            'Control Frequency': [],
            'Key control': [],
            'Original File': []
        }]
        cwd = os.getcwd()
        os.chdir(input_folder_path_rcm)
        self.RCM_progressBar_3.setMinimum(0)
        self.RCM_progressBar_3.setValue(0)
        self.RCM_progressBar_3.setMaximum(len(os.listdir()))

        for input_file_path in os.listdir():
            if os.path.exists(input_folder_path_rcm):
                os.chdir(input_folder_path_rcm)
            if os.path.isfile(input_file_path) and os.access(
                    input_file_path, os.R_OK):
                name, extension = os.path.splitext(input_file_path)
                if extension == ".xlsx":
                    if not os.path.exists(output_folder_path_rcm):
                        os.makedirs(output_folder_path_rcm)
                    my.xlsxsToxlsxInFilter(input_file_path, output_filecontent,
                                           origalfiles)
            self.RCM_progressBar_3.setValue(self.RCM_progressBar_3.value() + 1)
        os.chdir(cwd)
        my.fixtocorrecttype(output_filecontent)
        my.write_output_content_to_xlsx(
            output_filecontent, origalfiles,
            output_folder_path_rcm + '/combinedFile.xlsx')
        input_result_xlsx_filepath = output_folder_path_rcm + '/combinedFile.xlsx'
        self.Stat_RCM_Location_Label.setText(output_folder_path_rcm +
                                             '\n/combinedFile.xlsx')
        self.OtherCategory_listWidget.setEnabled(True)
        self.displayfilecontentIntable(input_result_xlsx_filepath)

    def on_row_changed(self, current):
        global output_folder_path
        try:
            filename = self.PPFile_listWidget_2.item(current).text()
            filepath = output_folder_path + "/" + filename
            with open(filepath, 'r') as content_file:
                self.PPContent_textEdit_2.setText(content_file.read())
        except:
            print("filename occurred")

    def isCheckedXMLORJSON(self, rb):
        global output_type
        if rb.text() == "XML":
            if rb.isChecked() == True:
                output_type = ".XML"
            else:
                print(rb.text() + " is deselected")

        if rb.text() == "JSON":
            if rb.isChecked() == True:
                output_type = ".JSON"
            else:
                print(rb.text() + " is deselected")

    def display_part1(self, i):
        self.stackedWidget.setCurrentIndex(i)

    def display_part2(self, i):
        self.stackedWidget.setCurrentIndex(i + 2)

    def open_sheet(self, column):
        df = pandas.read_excel("static\Input\Input Data + Sample output.xlsx",
                               sheet_name="Sheet1")
        _uniq = df[column].unique()
        len_uniq = len(_uniq)
        return str(len_uniq)


#     def piechart(self):
#             df = pandas.read_excel("static\Input\Input Data + Sample output.xlsx", sheet_name="Sheet1")
#             _uniq = df["Process Domain"].unique()
#             controls_to_process = df["Control ID"].unique()
#             data = ""

#             #Creating Plot

#             fig = plt.figure(figsize=(10, 7))
#             plt.pie(data, labels = _uniq)

#             plt.show()

    def displayfilecontentIntable(self, inputfilepath):
        data = pandas.read_excel(inputfilepath)
        columns = len(data.columns)
        rows = len(data.values)
        #########################RCM Aggregetor Stat Module###########################
        missing_column_content = {
            "Column Header": [
                'Process Domain',
                'Process ID',
                'Process Description',
                'Risk ID',
                'Risk Description',
                'Control ID',
                'Control Objective',
                'Control Description',
                'Control Owner',
                'Business Unit',
                'Control Type (Preventive , Detective)',
                'Control Type (Manual, Semi-automated, Automated)',
                'Control Frequency',
                'Key control',
            ]
        }
        populed_counts = {
            'Process Domain': 0,
            'Process ID': 0,
            'Process Description': 0,
            'Risk ID': 0,
            'Risk Description': 0,
            'Control ID': 0,
            'Control Objective': 0,
            'Control Description': 0,
            'Control Owner': 0,
            'Business Unit': 0,
            'Control Type (Preventive , Detective)': 0,
            'Control Type (Manual, Semi-automated, Automated)': 0,
            'Control Frequency': 0,
            'Key control': 0
        }
        populated_content = {
            "Column Header": [
                'Process Domain',
                'Process ID',
                'Process Description',
                'Risk ID',
                'Risk Description',
                'Control ID',
                'Control Objective',
                'Control Description',
                'Control Owner',
                'Business Unit',
                'Control Type (Preventive , Detective)',
                'Control Type (Manual, Semi-automated, Automated)',
                'Control Frequency',
                'Key control',
            ],
            "# of populated rows": [],
            "# of empty rows": []
        }
        #####################################
        countsoffile = []
        for filename in data['Original File']:
            missing_column_content.update({filename: []})
        ###############################

        originalfilenames = []
        for key in missing_column_content:
            ii = 0
            if key != 'Column Header':
                originalfilenames.append(key)
                for filename in data['Original File']:
                    if key == filename:
                        ii += 1
            countsoffile.append(ii)
        ######################################
        ################################
        temp = 0
        for key in populed_counts:
            valc = data.get(key)
            templist = []
            for v in valc:
                templist.append(str(v))
                if 'nan' != str(v):
                    populed_counts[key] += 1
            for ii in range(1, len(countsoffile)):
                count = countsoffile[ii]
                if my.isAllNanInList(templist[temp:temp + count]):
                    missing_column_content[originalfilenames[ii -
                                                             1]].append("N")
                else:
                    missing_column_content[originalfilenames[ii -
                                                             1]].append("Y")
                temp = count
        ############################################

        for key in populed_counts:
            pcount = populed_counts[key]
            emcount = rows - pcount
            populated_content['# of populated rows'].append(pcount)
            populated_content['# of empty rows'].append(emcount)

        #############################################################
        self.RCMContent_tableWidget_2.setColumnCount(columns)
        self.RCMContent_tableWidget_2.setRowCount(rows)
        # self.setAutoFillBackground(True)

        self.RCMContent_tableWidget_2.setHorizontalHeaderLabels(data.columns)

        for i in range(columns):
            for j in range(len(data.values)):
                item = QTableWidgetItem()
                txt = data.values[j][i]
                item.setText(str(txt))
                self.RCMContent_tableWidget_2.setItem(j, i, item)
        ##########################Stat Missing Columns TableWidget###########################
        self.Stat_Missing_Columns_tableWidget_3.setColumnCount(
            len(missing_column_content))
        self.Stat_Missing_Columns_tableWidget_3.setHorizontalHeaderLabels(
            missing_column_content)
        i = 0
        for key in missing_column_content:
            cv = missing_column_content[key]
            self.Stat_Missing_Columns_tableWidget_3.setRowCount(len(cv))
            for j in range(0, len(cv)):
                item = QTableWidgetItem()
                item.setText(cv[j])
                self.Stat_Missing_Columns_tableWidget_3.setItem(j, i, item)
            i += 1

        ###################################################################################
        ###################################################################################
        self.Stat_PoMi_Data_tableWidget_4.setColumnCount(
            len(populated_content))
        self.Stat_PoMi_Data_tableWidget_4.setHorizontalHeaderLabels(
            populated_content)
        i = 0
        for key in populated_content:
            cv = populated_content[key]
            self.Stat_PoMi_Data_tableWidget_4.setRowCount(len(cv))
            for j in range(0, len(cv)):
                item = QTableWidgetItem()
                item.setText(str(cv[j]))
                self.Stat_PoMi_Data_tableWidget_4.setItem(j, i, item)
            i += 1
        #####################################################################################
        ###############################################################
        total_num = len(originalfilenames)
        self.Stat_Total_Number_Label.setText(str(total_num))
        for fn in originalfilenames:
            self.Stat_RCM_Names_listWidget.addItem(QListWidgetItem(fn))

        ################################################################
    def reset(self):
        self.OtherCategory_listWidget.setEnabled(False)
        self.OtherCategory_listWidget.setStyleSheet(
            "QListView::item {\n"
            "      color:silver;\n"
            "      background-color:grey;\n"
            "  }")
        self.RCMOutput_LineText.clear()
        self.RCMInput_lineEdi_2.clear()
        self.PPOutput_lineEdit.clear()
        self.PP_progressBar_2.setValue(0)
        self.RCM_progressBar_2.setValue(0)
        self.PPInput_lineEdit.clear()
        self.PPFile_listWidget_2.clear()
        self.PPContent_textEdit_2.clear()
        self.RCMContent_tableWidget_2.clear()
        self.PPCategory_listWidget.item(0).setSelected(True)
        self.stackedWidget.setCurrentIndex(0)
        self.Stat_PoMi_Data_tableWidget_4.clear()
        self.Stat_Missing_Columns_tableWidget_3.clear()
        self.Stat_RCM_Names_listWidget.clear()
        self.Stat_RCM_Location_Label.clear()
        self.Stat_Total_Number_Label.setText('0')

    def connect2allwidgetlistener(self):
        #########Category###########
        self.PPCategory_listWidget.currentRowChanged.connect(
            self.display_part1)
        self.OtherCategory_listWidget.currentRowChanged.connect(
            self.display_part2)

        ############PP Extractor##############
        self.PPInput_pushButton.clicked.connect(self.inputFilePP)
        self.PPOutput_pushButton.clicked.connect(self.outputFilePP)
        self.PPConvert_pushButton_2.clicked.connect(self.convertPP)
        self.PPFile_listWidget_2.currentRowChanged.connect(self.on_row_changed)

        ############RCM Extractor##############
        self.RCMInput_pushButton_2.clicked.connect(self.inputFileInRCM)
        self.RCMOutput_pushButton_4.clicked.connect(self.outputFileRCM)
        self.RCMConvert_pushButton_2.clicked.connect(self.convertInRCM)
        ################Reset PushButton###############
        self.Reset_pushButton.clicked.connect(self.reset)

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())
    app.exec_()