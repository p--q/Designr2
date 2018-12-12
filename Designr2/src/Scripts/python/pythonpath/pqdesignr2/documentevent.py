#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# ドキュメントイベントについて。
from . import ichiran, points
MODIFYLISTENERS = []  # ModifyListenerのサブジェクトとリスナーのタプルのリスト。
def documentOnLoad(xscriptcontext):  # ドキュメントを開いた時。リスナー追加後。
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	namedranges = doc.getPropertyValue("NamedRanges")  # ドキュメントのNamedRangesを取得。
	for i in namedranges.getElementNames():  # namedrangesをイテレートするとfor文中でnamedrangesを操作してはいけない。
		if not namedranges[i].getReferredCells():
			namedranges.removeByName(i)  # 参照範囲がエラーの名前を削除する。	
	pointsvars = points.VARS  # 点数シートの固有値。
	splittedrow = pointsvars.splittedrow
	splittedcolumn = pointsvars.startcolumn
	sheets = doc.getSheets()
	addModifyListener(doc, (i[splittedrow:, splittedcolumn:].getRangeAddress() for i in sheets if i.getName().isdigit()), points.PointsModifyListener(xscriptcontext))  # 点数シートの点数の変更を検知するリスナー。	
	sheet = sheets["一覧"]			
	ichiranvars = ichiran.VARS	
	addModifyListener(doc, [sheet[ichiranvars.splittedrow:, ichiranvars.idcolumn:ichiranvars.enddaycolumn+1].getRangeAddress()], ichiran.DataModifyListener(xscriptcontext))  # 一覧シートの固定行以下,かつ、ID列から終了日列、の変更を検知するリスナー。
	doc.getCurrentController().setActiveSheet(sheet)  # 一覧シートをアクティブにする。	
	ichiran.initSheet(sheet, xscriptcontext)
def addModifyListener(doc, rangeaddresses, modifylistener):	
	cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
	cellranges.addRangeAddresses(rangeaddresses, False)
	cellranges.addModifyListener(modifylistener)
	MODIFYLISTENERS.append((cellranges, modifylistener))	
def documentUnLoad(xscriptcontext):  # ドキュメントを閉じた時。リスナー削除後。
	for subject, modifylistener in MODIFYLISTENERS:
		subject.removeModifyListener(modifylistener)
