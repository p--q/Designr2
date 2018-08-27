#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# 一覧シートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
# import os, unohelper, glob
# from itertools import chain
from indoc import commons, datedialog, points
from com.sun.star.awt import MouseButton, MessageBoxButtons, MessageBoxResults # 定数
from com.sun.star.awt.MessageBoxType import QUERYBOX  # enum
# from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.i18n.TransliterationModulesNew import FULLWIDTH_HALFWIDTH  # enum
from com.sun.star.lang import Locale  # Struct
from com.sun.star.sheet import CellFlags  # 定数
# from com.sun.star.sheet.CellDeleteMode import ROWS as delete_rows  # enum
# from com.sun.star.table.CellHoriJustify import LEFT  # enum
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
class Ichiran():  # シート固有の値。
	def __init__(self):
		self.menurow = 0
		self.splittedrow = 2  # 分割行インデックス。
		self.idcolumn = 0  # ID列インデックス。	
		self.kanjicolumn = 1  # 漢字列インデックス。	
		self.startdaycolumn = 2 # 開始日列インデックス。
		self.enddaycolumn = 3  # 終了日列インデックス。
	def setSheet(self, sheet):  # 逐次変化する値。
		self.sheet = sheet
		cellranges = sheet[self.splittedrow:, self.idcolumn].queryContentCells(CellFlags.STRING)  # ID列の文字列が入っているセルに限定して抽出。
		backcolors = commons.COLORS["black"], # ジェネレーターに使うので順番が重要。
		gene = (i.getCellAddress().Row for i in cellranges.getCells() if i.getPropertyValue("CellBackColor") in backcolors)
		self.blackrow = next(gene)  # 黒行インデックス。
		cellranges = sheet[:, self.idcolumn].queryContentCells(CellFlags.STRING+CellFlags.VALUE)  # ID列の文字列が入っているセルに限定して抽出。数値の時もありうる。
		self.emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # ID列の最終行インデックス+1を取得。
VARS = Ichiran()
def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
	if enhancedmouseevent.ClickCount==2 and enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ダブルクリックの時。まずselectionChanged()が発火している。
		selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			celladdress = selection.getCellAddress()
			r = celladdress.Row  # selectionの行インデックスを取得。	
			if r==VARS.menurow:  # メニュー行の時。:
				return wClickMenu(enhancedmouseevent, xscriptcontext)
			if r>=VARS.splittedrow or r !=VARS.blackrow:  # 分割行以下、かつ、区切り行でない、時。
				return wClickPt(enhancedmouseevent, xscriptcontext)
	return True  # セル編集モードにする。シングルクリックは必ずTrueを返さないといけない。		
def wClickMenu(enhancedmouseevent, xscriptcontext):
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	txt = selection.getString()  # クリックしたセルの文字列を取得。	
	if txt=="終了を消去":
		msg = "黒行上の行をすべて削除しますか?"
		componentwindow = xscriptcontext.getDocument().getCurrentController().ComponentWindow
		msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_OK_CANCEL+MessageBoxButtons.DEFAULT_BUTTON_CANCEL, "myRs", msg)
		if msgbox.execute()==MessageBoxResults.OK:	
			
				
			pass
	elif txt=="全印刷":
		
		pass
	elif txt=="過去月":
		
		# 同じフォルダにあるファイル一覧を取得してstaticdialogで開く。
		
		
		pass

def wClickPt(enhancedmouseevent, xscriptcontext):
	selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
	sheet = VARS.sheet
	celladdress = selection.getCellAddress()
	r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。
	idtxt, kanjitxt, datevalue = sheet[r, VARS.idcolumn:VARS.enddaycolumn].getDataArray()[0]
	if c==VARS.idcolumn:  # ID列の時。
		if idtxt:  # 空セルでない時。
			ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
			smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
			systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。
			systemclipboard.setContents(commons.TextTransferable(idtxt), None)  # クリップボードにIDをコピーする。
		else:
			return True  # セル編集モードにする。
	elif c==VARS.kanjicolumn:  # 漢字列の時。IDシートをアクティブにする、なければ作成する。シート名はIDと一致。	
		doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 	
		sheets = doc.getSheets()  # シートコレクションを取得。			
		if idtxt in sheets:  # 経過列があり、かつ、IDシートが存在する時。
			doc.getCurrentController().setActiveSheet(sheets[idtxt])  # ID名のシートをアクティブにする。
		else:  # ID名シートがない時。
			if all((idtxt, kanjitxt, datevalue)):  # ID、漢字名、開始日、すべてが揃っている時。	
				colors = commons.COLORS
				ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
				smgr = ctx.getServiceManager()  # サービスマネージャーの取得。				
				functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。		
				daycount = int(functionaccess.callFunction("DAYSINMONTH", (datevalue,)))  # 開始月の日数を取得。
				startdatevalue = int(functionaccess.callFunction("EOMONTH", (datevalue, -1))) + 1  # 開始月の開始日のシリアル値を取得。
				sheets.copyByName("00000000", idtxt, len(sheets))  # テンプレートシートをコピーしてID名のシートにして最後に挿入。	
				idsheet = sheets[idtxt]  # IDシートを取得。  
				pointsvars = points.VARS
				datarows = [(idtxt,), (kanjitxt,)]
				datarows.extend((i,) for i in range(startdatevalue, startdatevalue+daycount))
				splittedrow = pointsvars.splittedrow
				emptyrow = splittedrow + daycount
				idsheet[:emptyrow, pointsvars.daycolumn].setDataArray(datarows)
				idsheet[splittedrow+1:emptyrow, :pointsvars.mincolumn].setPropertyValue("CellBackColor", colors["silver"])  # 背景色をつける
				idsheet[splittedrow:emptyrow, pointsvars.daycolumn].setPropertyValue("NumberFormat", commons.formatkeyCreator(doc)("YYYY-M-DD"))
				y, m = [int(functionaccess.callFunction(i, (startdatevalue,))) for i in ("YEAR", "MONTH")]
				holidays = commons.HOLIDAYS	
				holidayindexes = set()
				if y in holidays:
					holidayindexes.update(holidays[y][m-1])
				startweekday = int(functionaccess.callFunction("WEEKDAY", (startdatevalue, 3)))  # 開始日の曜日を取得。月=0。
				n = 6  # 日曜日の曜日番号。
				sunindexes = set(range(splittedrow+(n-startweekday)%7, emptyrow, 7))  # 日曜日の列インデックスの集合。祝日と重ならないようにあとで使用する。	
				holidayindexes.difference_update(sunindexes)  # 祝日インデックスから日曜日インデックスを除く。
				n = 5  # 土曜日の曜日番号。
				satindexes = set(range(splittedrow+(n-startweekday)%7, emptyrow, 7))  # 土曜日の列インデックスの集合。
				setRangesProperty = createSetRangesProperty(doc, idsheet, pointsvars.daycolumn)
				setRangesProperty(holidayindexes, ("CellBackColor", colors["red3"]))
				setRangesProperty(sunindexes, ("CharColor", colors["red3"]))
				setRangesProperty(satindexes, ("CharColor", colors["skyblue"]))
				doc.getCurrentController().setActiveSheet(idsheet)  # IDシートをアクティブにする。	
			else:
				return True  # セル編集モードにする。						
	elif c==VARS.startdaycolumn:  # 開始日列の時。
		datedialog.createDialog(enhancedmouseevent, xscriptcontext, "開始日", "YYYY-M")			
	elif c==VARS.enddaycolumn:  # 終了日列の時。
		datedialog.createDialog(enhancedmouseevent, xscriptcontext, "終了日", "YYYY-M")		
	return False  # セル編集モードにしない。	
def createSetRangesProperty(doc, sheet, c): 
	def setRangesProperty(rowindexes, prop):  # c列のrowindexesの行のプロパティを変更。prop: プロパティ名とその値のリスト。
		cellranges = doc.createInstance("com.sun.star.sheet.SheetCellRanges")  # セル範囲コレクション。
		if rowindexes:  
			cellranges.addRangeAddresses([sheet[i, c].getRangeAddress() for i in rowindexes], False)  # セル範囲コレクションを取得。rowindexesが空要素だとエラーになる。
			if len(cellranges):  # sheetcellrangesに要素がないときはsetPropertyValue()でエラーになるので要素の有無を確認する。
				cellranges.setPropertyValue(*prop)  # セル範囲コレクションのプロパティを変更。
	return setRangesProperty	
def selectionChanged(eventobject, xscriptcontext):  # 矢印キーでセル移動した時も発火する。
	selection = eventobject.Source.getSelection()
	if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 選択範囲がセル範囲の時。
		VARS.setSheet(selection.getSpreadsheet())			
		drowBorders(selection)  # 枠線の作成。
def drowBorders(selection):  # ターゲットを交点とする行列全体の外枠線を描く。
	celladdress = selection[0, 0].getCellAddress()  # 選択範囲の左上端のセルアドレスを取得。
	r = celladdress.Row  # selectionの行と列のインデックスを取得。	
	sheet = VARS.sheet
	noneline, tableborder2, topbottomtableborder, leftrighttableborder = commons.createBorders()
	sheet[:, :].setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。	
	if r<VARS.splittedrow or r==VARS.blackrow:  # 分割行より上か黒行の時。
		return  # 罫線を引き直さない。
	rangeaddress = selection.getRangeAddress()  # 選択範囲のセル範囲アドレスを取得。
	sheet[rangeaddress.StartRow:rangeaddress.EndRow+1, :].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く
	sheet[:, rangeaddress.StartColumn:rangeaddress.EndColumn+1].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。
	selection.setPropertyValue("TableBorder2", tableborder2)  # 選択範囲の消えた枠線を引き直す。		
def changesOccurred(changesevent, xscriptcontext):  # Sourceにはドキュメントが入る。マクロで変更した時は発火しない模様。	
	selection = None
	for change in changesevent.Changes:
		if change.Accessor=="cell-change":  # セルの値が変化した時。
			selection = change.ReplacedElement  # 値を変更したセルを取得。	
			break
	if selection:
		sheet = selection.getSpreadsheet()
		VARS.setSheet(sheet)
		celladdress = selection.getCellAddress()
		r, c = celladdress.Row, celladdress.Column
		if r>=VARS.splittedrow:  # 分割行以降の時。
			ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
			smgr = ctx.getServiceManager()  # サービスマネージャーの取得。					
			transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。		
			transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))	
			txt = selection.getString()  # セルの文字列を取得。			
			if c==VARS.idcolumn:  # ID列の時。
				txt = transliteration.transliterate(txt, 0, len(txt), [])[0]  # 半角に変換。
				if txt.isdigit():  # 数値の時のみ。空文字の時0で埋まってしまう。
					selection.setString("{:0>8}".format(txt))  # 数値を8桁にして文字列として代入し直す。
			elif c==VARS.kanjicolumn:
				selection.setString(txt.replace("　", " "))  # 全角スペースを半角スペースに置換。
def notifyContextMenuExecute(contextmenuexecuteevent, xscriptcontext):  # 右クリックメニュー。	
	controller = contextmenuexecuteevent.Selection  # コントローラーは逐一取得しないとgetSelection()が反映されない。
	sheet = controller.getActiveSheet()  # アクティブシートを取得。
	VARS.setSheet(sheet)
	contextmenu = contextmenuexecuteevent.ActionTriggerContainer  # コンテクストメニューコンテナの取得。
	contextmenuname = contextmenu.getName().rsplit("/")[-1]  # コンテクストメニューの名前を取得。
	addMenuentry = commons.menuentryCreator(contextmenu)  # 引数のActionTriggerContainerにインデックス0から項目を挿入する関数を取得。
	baseurl = commons.getBaseURL(xscriptcontext)  # ScriptingURLのbaseurlを取得。
	del contextmenu[:]  # contextmenu.clear()は不可。
	selection = controller.getSelection()  # 現在選択しているセル範囲を取得。
	celladdress = selection[0, 0].getCellAddress()  # 選択範囲の左上角のセルのアドレスを取得。
	r = celladdress.Row  # selectionの行と列のインデックスを取得。		
	if r<VARS.splittedrow or r==VARS.blackrow:  # 固定行より上、または黒行の時はコンテクストメニューを表示しない。
		return EXECUTE_MODIFIED
	elif contextmenuname=="cell":  # セルのとき。セル範囲も含む。
		commons.cutcopypasteMenuEntries(addMenuentry)
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:PasteSpecial"})		
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
		addMenuentry("ActionTrigger", {"Text": "クリア", "CommandURL": baseurl.format("entry1")}) 
	elif contextmenuname=="rowheader" and len(selection[0, :].getColumns())==len(sheet[0, :].getColumns()):  # 行ヘッダーのとき、かつ、選択範囲の列数がシートの列数が一致している時。	
		if r>=VARS.splittedrow:
			if r<VARS.blackrow:
				addMenuentry("ActionTrigger", {"Text": "使用中最上行へ", "CommandURL": baseurl.format("entry15")})  # 黒行上から使用中最上行へ
				addMenuentry("ActionTrigger", {"Text": "使用中最下行へ", "CommandURL": baseurl.format("entry16")})  # 黒行上から使用中最下行へ
			elif r>VARS.blackrow:  # 黒行以外の時。
				addMenuentry("ActionTrigger", {"Text": "黒行上へ", "CommandURL": baseurl.format("entry17")})  
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
				addMenuentry("ActionTrigger", {"Text": "使用中最上行へ", "CommandURL": baseurl.format("entry18")})  # 使用中から使用中最上行へ  
				addMenuentry("ActionTrigger", {"Text": "使用中最下行へ", "CommandURL": baseurl.format("entry19")})  # 使用中から使用中最下行へ		
			if r!=VARS.blackrow:
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
				commons.cutcopypasteMenuEntries(addMenuentry)
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
				commons.rowMenuEntries(addMenuentry)		
	elif contextmenuname=="colheader":  # 列ヘッダーの時。
		pass
	elif contextmenuname=="sheettab":  # シートタブの時。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
	return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。	
def contextMenuEntries(entrynum, xscriptcontext):  # コンテクストメニュー番号の処理を振り分ける。引数でこれ以上に取得できる情報はない。		
	controller = xscriptcontext.getDocument().getCurrentController()  # コントローラの取得。
	selection = controller.getSelection()  # 選択範囲を取得。
	if entrynum==1:  # クリア。書式設定とオブジェクト以外を消去。
		selection.clearContents(511)  # 範囲をすべてクリアする。
	elif 14<entrynum<20:
		sheet = controller.getActiveSheet()  # アクティブシートを取得。
		VARS.setSheet(sheet)
		rangeaddress = selection.getRangeAddress()  # 選択範囲のアドレスを取得。
		if entrynum==15:  # 黒行上から使用中最上行へ
			commons.toOtherEntry(sheet, rangeaddress, VARS.blackrow, VARS.blackrow+1)
		elif entrynum==16:  # 黒行上から使用中最下行へ
			commons.toNewEntry(sheet, rangeaddress, VARS.blackrow, VARS.emptyrow) 
		elif entrynum==17:  # 黒行上へ
			commons.toOtherEntry(sheet, rangeaddress, VARS.emptyrow, VARS.blackrow)  
		elif entrynum==18:  # 使用中から使用中最上行へ 
			commons.toOtherEntry(sheet, rangeaddress, VARS.emptyrow, VARS.blackrow+1)
		elif entrynum==19:  # 使用中から使用中最下行へ		
			commons.toNewEntry(sheet, rangeaddress, VARS.emptyrow, VARS.emptyrow) 		
