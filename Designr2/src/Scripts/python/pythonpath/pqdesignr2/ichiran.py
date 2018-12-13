#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# 一覧シートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
import os, unohelper, glob
from . import commons, datedialog, points, menudialog
from com.sun.star.accessibility import AccessibleRole  # 定数
from com.sun.star.awt import MouseButton, MessageBoxButtons, MessageBoxResults, ScrollBarOrientation # 定数
from com.sun.star.awt.MessageBoxType import INFOBOX, QUERYBOX  # enum
from com.sun.star.beans import PropertyValue  # Struct
from com.sun.star.i18n.TransliterationModulesNew import FULLWIDTH_HALFWIDTH  # enum
from com.sun.star.lang import Locale  # Struct
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.sheet.CellDeleteMode import ROWS as delete_rows  # enum
from com.sun.star.table import BorderLine2  # Struct
from com.sun.star.table import BorderLineStyle  # 定数
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum
from com.sun.star.util import XModifyListener
class Ichiran():  # シート固有の値。
	def __init__(self):
		self.splittedrow = 2  # 分割行インデックス。
		self.sumicolumn = 0  # 済列インデックス。
		self.idcolumn = 1  # ID列インデックス。	
		self.kanjicolumn = 2  # 漢字列インデックス。	
		self.startdaycolumn = 3 # 開始日列インデックス。
		self.enddaycolumn = 4  # 終了日列インデックス。
	def setSheet(self, sheet):  # 逐次変化する値。
		self.sheet = sheet
		cellranges = sheet[:, self.idcolumn].queryContentCells(CellFlags.STRING+CellFlags.VALUE)  # ID列の文字列か数値が入っているセルに限定して抽出。
		self.emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # ID列の最終行インデックス+1を取得。	
		idcolumn = self.idcolumn	
		backcolor = commons.COLORS["black"] # 行インデックスを取得したいセルの背景色。
		for i in range(self.splittedrow, self.emptyrow):  # 固定行から最終行まで行インデックスをイテレート。
			if sheet[i, idcolumn].getPropertyValue("CellBackColor")==backcolor:
				self.blackrow = i  # 黒行インデックスを取得。
				sheet[i, idcolumn].clearContents(CellFlags.VALUE+CellFlags.DATETIME+CellFlags.STRING+CellFlags.FORMULA)  # 黒行のID列のセルの値はクリアする。
				break
VARS = Ichiran()
def activeSpreadsheetChanged(activationevent, xscriptcontext):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。
	initSheet(activationevent.ActiveSheet, xscriptcontext)
def initSheet(sheet, xscriptcontext):	
	datarows = ("", "メニュー", "済をﾘｾｯﾄ"),
	sheet[0, :len(datarows[0])].setDataArray(datarows)
	accessiblecontext = xscriptcontext.getDocument().getCurrentController().ComponentWindow.getAccessibleContext()  # コントローラーのアトリビュートからコンポーネントウィンドウを取得。
	for i in range(accessiblecontext.getAccessibleChildCount()): 
		childaccessiblecontext = accessiblecontext.getAccessibleChild(i).getAccessibleContext()
		if childaccessiblecontext.getAccessibleRole()==AccessibleRole.SCROLL_PANE:
			for j in range(childaccessiblecontext.getAccessibleChildCount()): 
				child2 = childaccessiblecontext.getAccessibleChild(j)
				childaccessiblecontext2 = child2.getAccessibleContext()
				if childaccessiblecontext2.getAccessibleRole()==AccessibleRole.SCROLL_BAR:  # スクロールバーの時。
					if child2.getOrientation()==ScrollBarOrientation.VERTICAL:  # 縦スクロールバーの時。
						if childaccessiblecontext2.getBounds().Height>0:  # 右上枠の縦スクロールバーのHeghtが0になっている。
							child2.setValue(0)  # 縦スクロールバーを一番上にする。
							return  # breakだと二重ループは抜けれない。
def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
	if enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左クリックの時。
		selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			celladdress = selection.getCellAddress()
			r, c = celladdress.Row, celladdress.Column  # selectionの行と列インデックスを取得。	
			if enhancedmouseevent.ClickCount==1:  # 左シングルクリックの時。
				VARS.setSheet(selection.getSpreadsheet())  # VARS.sheetがまだ取得出来ていない時がある。
				if c==VARS.sumicolumn and VARS.splittedrow<=r<VARS.emptyrow:  # 済列の時。
					txt = selection.getString()
					if not txt:  # まだ空セルの時は未として扱う。
						txt = "未"
					items = [("待", "skyblue"), ("済", "silver"), ("未", "black")]
					items.append(items[0])  # 最初の要素を最後の要素に追加する。
					dic = {items[i][0]: items[i+1] for i in range(len(items)-1)}  # 順繰り辞書の作成。				
					newtxt = dic[txt][0]							
					selection.setString(newtxt)
					VARS.sheet[r, :].setPropertyValue("CharColor", commons.COLORS[dic[txt][1]])		
			elif enhancedmouseevent.ClickCount==2:  # 左ダブルクリックの時。まずselectionChanged()が発火している。
				if r<VARS.splittedrow:  # 固定行より上の時。
					txt = selection.getString()	
					if txt=="済をﾘｾｯﾄ":
						componentwindow = xscriptcontext.getDocument().getCurrentController().ComponentWindow
						querybox = lambda x: componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO+MessageBoxButtons.DEFAULT_BUTTON_YES, "WEntryBook", x)
						msgbox = querybox("済列をリセットします。")
						if msgbox.execute()!=MessageBoxResults.YES:  # Yes以外の時はここで終わる。		
							return							
						VARS.sheet[VARS.splittedrow:VARS.emptyrow, :].setPropertyValue("CharColor", commons.COLORS["black"])  # 文字色を黒色にする。
						VARS.sheet[VARS.splittedrow:VARS.emptyrow, VARS.sumicolumn].setDataArray([("未",)]*(VARS.emptyrow-VARS.splittedrow))  # 済列をリセット。
					elif txt=="メニュー":
						defaultrows = "継続患者のみ印刷", "全患者印刷", "月末まで埋めて全患者印刷",  "全部位終了患者を一覧から消去", "------", "過去月ファイル一覧を表示"
						menudialog.createDialog(xscriptcontext, txt, defaultrows, enhancedmouseevent=enhancedmouseevent, callback=callback_menuCreator(xscriptcontext))
					return False  # セル編集モードにしない。					
				elif r!=VARS.blackrow:  # 分割行以下、かつ、黒行でない、時。
					return wClickPt(enhancedmouseevent, xscriptcontext)
	return True  # セル編集モードにする。シングルクリックは必ずTrueを返さないといけない。
def callback_menuCreator(xscriptcontext):  # 内側のスコープでクロージャの変数を再定義するとクロージャの変数を参照できなくなる。	
	doc = xscriptcontext.getDocument()
	controller = doc.getCurrentController()
	componentwindow = controller.ComponentWindow
	querybox = lambda x: componentwindow.getToolkit().createMessageBox(componentwindow, QUERYBOX, MessageBoxButtons.BUTTONS_YES_NO+MessageBoxButtons.DEFAULT_BUTTON_YES, "WEntryBook", x)
	def callback_menu(gridcelltxt):			
		if gridcelltxt=="継続患者のみ印刷":	
			printername = getPrinterName(xscriptcontext.getDocument())			
			msgbox = querybox("{}{}します。".format(printername, gridcelltxt))
			if msgbox.execute()!=MessageBoxResults.YES:  # Yes以外の時はここで終わる。		
				return	
			printsheetnames = [i[0] for i in VARS.sheet[VARS.splittedrow:VARS.emptyrow, VARS.idcolumn:VARS.enddaycolumn+1].getDataArray() if i[0] and not i[-1]]  # 黒行でなく、かつ、終了日セルが空セル、のIDのリスト。
			if printsheetnames:  # 印刷するシートがあるとき。
				printPointsSheets(xscriptcontext, printername, printsheetnames)
		elif gridcelltxt=="全患者印刷":	
			msgbox = querybox("{}します。".format(gridcelltxt))
			if msgbox.execute()!=MessageBoxResults.YES:  # Yes以外の時はここで終わる。		
				return	
			printsheetnames = [i[0] for i in VARS.sheet[VARS.splittedrow:VARS.emptyrow, VARS.idcolumn:VARS.enddaycolumn+1].getDataArray() if i[0]]  # IDのリスト。
			if printsheetnames:  # 印刷するシートがあるとき。
				printPointsSheets(xscriptcontext, printername, printsheetnames)
		elif gridcelltxt=="月末まで埋めて全患者印刷":
			msgbox = querybox("{}します。".format(gridcelltxt))
			if msgbox.execute()!=MessageBoxResults.YES:  # Yes以外の時はここで終わる。		
				return	
			printsheetnames = [i[0] for i in VARS.sheet[VARS.splittedrow:VARS.emptyrow, VARS.idcolumn:VARS.enddaycolumn+1].getDataArray() if i[0]]  # IDのリスト。
			if printsheetnames:  # 印刷するシートがあるとき。
				printPointsSheets(xscriptcontext, printername, printsheetnames, True)			
		elif gridcelltxt=="全部位終了患者を一覧から消去":
			msg = "全部位終了しているシートを年月.odsファイルに移動し、\nこの一覧から消去します"
			msgbox = querybox(msg)
			if msgbox.execute()!=MessageBoxResults.YES:  # Yes以外の時はここで終わる。		
				return	
			sheet = VARS.sheet
			sheets = doc.getSheets()
			ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
			smgr = ctx.getServiceManager()  # サービスマネージャーの取得。				
			functionaccess = smgr.createInstanceWithContext("com.sun.star.sheet.FunctionAccess", ctx)  # シート関数利用のため。	
			pointsvars = points.VARS	
			startcolumnidx = pointsvars.startcolumn + 7
			splittedrow = pointsvars.splittedrow
			daycolumn = pointsvars.daycolumn
			for i, datarow in enumerate(sheet[VARS.splittedrow:VARS.emptyrow, VARS.idcolumn].getDataArray()[::-1], start=1):  # IDの行をイテレート。行を削除するので逆順にする。sheetsのイテレートではsheetsの操作ができない。
				if datarow[0]:  # 空セルでない時。
					sheetname = datarow[0]  # シート名を取得。
					pointssheet = sheets[sheetname]  # IDのシートを取得。
					pointsvars.setSheet(pointssheet)  # シートによって変化する値を取得。
					for j in range(startcolumnidx, pointsvars.emptycolumn, 8):  # 部位別合計列インデックスをイテレート。			
						if pointssheet[pointsvars.emptyrow-1, j].getPropertyValue("CellBackColor")==-1:  # 最終日の部位別合計列セルに背景色がない時。
							break
					else:  # for文中でbreakしなかった時は最終日の部位別合計のすべてに背景色があるか、部位が一つもない時。
						y, m = [int(functionaccess.callFunction(j, (pointssheet[splittedrow, daycolumn].getValue(),))) for j in ("YEAR", "MONTH")]  # IDシートの日付セルの年と月を取得。	
						points.createCopySheet(xscriptcontext, y)(sheetname, m)  # IDシートを年月名のファイルにコピーする。
						sheets.removeByName(sheetname)  # コピーしたシートは削除する。
						sheet.removeRange(sheet[VARS.emptyrow-i, 0].getRangeAddress(), delete_rows)  # 削除したシートのID行を削除。			
		elif gridcelltxt=="過去月ファイル一覧を表示":
			dirpath = os.path.dirname(unohelper.fileUrlToSystemPath(doc.getURL()))  # このドキュメントのあるディレクトリのフルパスを取得。
			defaultrows = [os.path.basename(i).split(".")[0] for i in glob.iglob(os.path.join(dirpath, "*", "*年*月.ods"), recursive=True)]  # *年*月のみリストに取得。
			if defaultrows:
				defaultrows.sort(key=lambda x: "{}{:0>2}".format(*x[:-1].split(x[4])))  # 年４桁固定、桁不定月との間に区切り文字が一文字、最後に月数でない文字列が一つあると決めつけて昇順でソートしている。
				menudialog.createDialog(xscriptcontext, "過去月", defaultrows, callback=callback_wClickGridCreator(xscriptcontext, dirpath))
			else:
				msg = "過去のファイルはありません。"
				commons.showErrorMessageBox(controller, msg)
	return callback_menu
def getPrinterName(doc):  # プリンター名を取得。
	for i in doc.getPrinter():  # 現在のプリンターのPropertyValueをイテレート。
		if i.Name=="Name":  # プリンター名の時。
			return "プリンター「{}」で\n".format(i.Value)
	return ""			
def printPointsSheets(xscriptcontext, printername, printsheetnames, fillToEnd=None):  # printsheetnames: 印刷するシート名のイテラブル。fillToEndがTrueの時は月末まで埋める。
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	sheets = doc.getSheets()
	pointsvars = points.VARS
	endpage = 1  # 印刷終了ページ番号。
	noneline = BorderLine2(LineStyle=BorderLineStyle.NONE)
	for i in printsheetnames[::-1]:  # 印刷するシート名を逆順にイテレート。sheetsをイテレートするとsheetsが操作できない。
		if i in sheets:  # シート名がシートコレクショにある時。
			printsheet = sheets[i]  # 印刷するシートを取得。
			pointsvars.setSheet(printsheet)  # シートによって変化する値を取得。
			printsheet[0, :pointsvars.daycolumn].clearContents(CellFlags.STRING)  # ボタンセルを消去する。印刷しないので。シートをアクティブしたときに再度ボタンセルに文字列を代入する。
			printsheet[:, :].setPropertyValue("TopBorder2", noneline)  # 枠線を消す。1辺をNONEにするだけですべての枠線が消える。	
			if fillToEnd is not None:
				points.fillToEndDayRow(doc, pointsvars.emptyrow-1)  # 最終日まで埋める。
			printsheet.setPrintAreas((printsheet[:pointsvars.emptyrow, :pointsvars.emptycolumn].getRangeAddress(),))  # 印刷範囲を設定。			
			sheets.moveByName(i, 0)  # 先頭に持ってくる。
			endpage += 1  # 印刷終了ページ番号を増やす。
	sheets.moveByName("一覧", 0)  # 一覧シートを一番先頭にする。	
	VARS.sheet.setPrintAreas((VARS.sheet[0, 1].getRangeAddress(),))  # 一覧シートの印刷範囲を設定。印刷しないページは1ページで収まるようにする。Windowsでは空セルを指定すると印刷ページにカウントされない。
	controller = doc.getCurrentController()
	if endpage>1:  # 印刷するページがある時。
		doc.getStyleFamilies()["PageStyles"]["Default"].setPropertyValues(("HeaderIsOn", "FooterIsOn", "IsLandscape", "ScaleToPages"), (False, False, True, 1))  # ヘッダーとフッターを付けない、用紙方向を横に、ページ数に合わせて縮小印刷。
		ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
		smgr = ctx.getServiceManager()  # サービスマネージャーの取得。		
		dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)		
		dispatcher.executeDispatch(controller.getFrame(), ".uno:TableSelectAll", "", 0, ())  # すべてのシートを選択。
		propertyvalues = PropertyValue(Name="Pages", Value="2-{}".format(endpage)),  # 印刷ページの指定。	
		doc.print(propertyvalues)  # startpage以降のみ印刷。
		msg = "{}印刷しました。".format(printername)
		componentwindow = controller.ComponentWindow
		msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, INFOBOX, MessageBoxButtons.BUTTONS_OK, "Designr", msg)
		msgbox.execute()
	else:
		commons.showErrorMessageBox(controller, "印刷するシートがありません。")	
	[i.setPrintAreas([]) for i in sheets]  # すべてのシートの印刷範囲をクリアする。  
def callback_wClickGridCreator(xscriptcontext, dirpath):
	def callback_wClickGrid(gridcelldata):  # gridcelldata: グリッドコントロールのダブルクリックしたセルのデータ。	
		systempath = next(glob.iglob(os.path.join(dirpath, "*", "{}.ods".format(gridcelldata)), recursive=True))  # ファイルパスを取得。	
		fileurl = unohelper.systemPathToFileUrl(systempath)	
		xscriptcontext.getDesktop().loadComponentFromURL(fileurl, "_blank", 0, ())  # ファイルを開く。
	return callback_wClickGrid
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
			transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。		
			transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))			
			txt = selection.getString()  # セルの文字列を取得。			
			txt = transliteration.transliterate(txt, 0, len(txt), [])[0]  # 半角に変換。
			if txt.isdigit():  # 数値の時のみ。空文字の時0で埋まってしまう。
				selection.setString("{:0>8}".format(txt))  # 数値を8桁にして文字列として代入し直す。
			systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。
			systemclipboard.setContents(commons.TextTransferable(idtxt), None)  # クリップボードにIDをコピーする。
		else:
			return True  # セル編集モードにする。
	elif c==VARS.kanjicolumn:  # 漢字列の時。IDシートをアクティブにする、なければ作成する。シート名はIDと一致。	
		doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 	
		sheets = doc.getSheets()  # シートコレクションを取得。	
		selection.setString(selection.getString().replace("　", " "))  # 全角スペースを半角スペースに置換。	
		if idtxt in sheets:
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
				pointsvars.setSheet(idsheet)  # 日付代入後に変化する値を取得する。
				points.colorizeDays(doc, functionaccess, startdatevalue)
				doc.getCurrentController().setActiveSheet(idsheet)  # IDシートをアクティブにする。	
			else:
				return True  # セル編集モードにする。						
	elif c==VARS.startdaycolumn:  # 開始日列の時。
		datedialog.createDialog(enhancedmouseevent, xscriptcontext, "開始日", "YYYY-M-D")			
	elif c==VARS.enddaycolumn:  # 終了日列の時。
		datedialog.createDialog(enhancedmouseevent, xscriptcontext, "終了日", "YYYY-M-D")		
	return False  # セル編集モードにしない。	
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
class DataModifyListener(unohelper.Base, XModifyListener):  # 固定行以下,かつ、ID列から終了日列、が変更された時に発火する。
	def __init__(self, xscriptcontext):
		self.formatkey = commons.formatkeyCreator(xscriptcontext.getDocument())("#,##0;[BLUE]-#,##0")
	def modified(self, eventobject):  # 固定行以下固定列右のセルが変化すると発火するメソッド。サブジェクトのどこが変化したかはわからない。eventobject.Sourceは対象全シートのセル範囲コレクション。
		sheet = VARS.sheet
		eventobject.Source.removeModifyListener(self)  # 値を変更するセル範囲のModifyListnerを外しておく。
		VARS.setSheet(sheet)  # 最終行と黒行を取得し直す。
		
		
		sheet[VARS.splittedrow:VARS.emptyrow, VARS.idcolumn:VARS.enddaycolumn+1]("HoriJustify", "NumberFormat"), (LEFT, createFormatKey("M/D"))
			
			# ID列、開始日列、終了日列の書式設定。
			
			
			
		eventobject.Source.addModifyListener(self)  # ModifyListnerを付け直す。
			
			
	def disposing(self, eventobject):
		eventobject.Source.removeModifyListener(self)
# def changesOccurred(changesevent, xscriptcontext):  # Sourceにはドキュメントが入る。マクロで変更した時は発火しない模様。	
# 	selection = None
# 	for change in changesevent.Changes:
# 		if change.Accessor=="cell-change":  # セルの値が変化した時。
# 			selection = change.ReplacedElement  # 値を変更したセルを取得。	
# 			break
# 	if selection:  # セルとは限らずセル範囲のときもある。シートからペーストしたときなど。テキストをペーストした時は発火しない。
# 		sheet = VARS.sheet
# 		splittedrow = VARS.splittedrow
# 		idcolumn = VARS.idcolumn
# 		kanjicolumn = VARS.kanjicolumn
# 		ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
# 		smgr = ctx.getServiceManager()  # サービスマネージャーの取得。		
# 		rangeaddress = selection.getRangeAddress()
# 		transliteration = smgr.createInstanceWithContext("com.sun.star.i18n.Transliteration", ctx)  # Transliteration。		
# 		transliteration.loadModuleNew((FULLWIDTH_HALFWIDTH,), Locale(Language = "ja", Country = "JP"))			
# 		for r in range(rangeaddress.StartRow, rangeaddress.EndRow+1):
# 			for c in range(rangeaddress.StartColumn, rangeaddress.EndColumn+1):
# 				if r>=splittedrow:  # 分割行以降の時。
# 					txt = sheet[r, c].getString()  # セルの文字列を取得。			
# 					if c==idcolumn:  # ID列の時。
# 						txt = transliteration.transliterate(txt, 0, len(txt), [])[0]  # 半角に変換。
# 						if txt.isdigit():  # 数値の時のみ。空文字の時0で埋まってしまう。
# 							sheet[r, c].setString("{:0>8}".format(txt))  # 数値を8桁にして文字列として代入し直す。
# 					elif c==kanjicolumn:
# 						sheet[r, c].setString(txt.replace("　", " "))  # 全角スペースを半角スペースに置換。
def notifyContextMenuExecute(contextmenuexecuteevent, xscriptcontext):  # 右クリックメニュー。	
	contextmenuname, addMenuentry, baseurl, selection = commons.contextmenuHelper(VARS, contextmenuexecuteevent, xscriptcontext)
	celladdress = selection[0, 0].getCellAddress()  # 選択範囲の左上角のセルのアドレスを取得。
	r = celladdress.Row  # selectionの行と列のインデックスを取得。	
	if r<VARS.splittedrow or r==VARS.blackrow:  # 固定行より上、または黒行の時はコンテクストメニューを表示しない。
		return EXECUTE_MODIFIED
	elif contextmenuname=="cell":  # セルのとき。セル範囲も含む。
		addMenuentry("ActionTrigger", {"Text": "選択患者のみ印刷", "CommandURL": baseurl.format("entry2")}) 
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
		commons.cutcopypasteMenuEntries(addMenuentry)
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:PasteSpecial"})		
		addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
		addMenuentry("ActionTrigger", {"Text": "クリア", "CommandURL": baseurl.format("entry1")}) 
	elif contextmenuname=="rowheader" and len(selection[0, :].getColumns())==len(VARS.sheet[0, :].getColumns()):  # 行ヘッダーのとき、かつ、選択範囲の列数がシートの列数が一致している時。	
		if r>=VARS.splittedrow:
			if r<VARS.blackrow:  # 黒行より上の時。
				addMenuentry("ActionTrigger", {"Text": "黒行上へ", "CommandURL": baseurl.format("entry15")})  # 黒行上へ移動。
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
				addMenuentry("ActionTrigger", {"Text": "最下行へ", "CommandURL": baseurl.format("entry16")})  # 最下行へ移動。
			elif r>VARS.blackrow:  # 黒行以外の時。
				addMenuentry("ActionTrigger", {"Text": "黒行上へ", "CommandURL": baseurl.format("entry17")})  # 黒行上へ移動。  
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})
				addMenuentry("ActionTrigger", {"Text": "黒行下へ", "CommandURL": baseurl.format("entry18")})  # 黒行下へ移動。  
				addMenuentry("ActionTrigger", {"Text": "最下行へ", "CommandURL": baseurl.format("entry19")})  # 最下行へ。		
			if r!=VARS.blackrow:  # 黒行でないときのみ。
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
	elif entrynum==2:  # 選択患者のみ印刷
		
		pass
		
	elif 14<entrynum<20:
		sheet = controller.getActiveSheet()  # アクティブシートを取得。
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
