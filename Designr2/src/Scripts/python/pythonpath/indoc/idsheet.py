#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# IDシートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from indoc import commons, datedialog, staticdialog
from com.sun.star.awt import MouseButton, MessageBoxButtons, MessageBoxResults # 定数
from com.sun.star.awt.MessageBoxType import WARNINGBOX  # enum
from com.sun.star.sheet import CellFlags  # 定数
from com.sun.star.sheet.CellDeleteMode import COLUMNS as delete_columns  # enum
from com.sun.star.table.CellHoriJustify import CENTER  # enum
from com.sun.star.ui import ActionTriggerSeparatorType  # 定数
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED  # enum

class IDsheet():  # シート固有の値。
	def __init__(self):
		self.splittedrow = 2  # 分割行インデックス。
		self.daycolumn = 4  # 日付列インデックス。
		self.startcolumn = 5  # 開始列インデックス。	
		self.dic = {\
			"深さ": ("0: 皮膚損傷・発赤なし", "1: 持続する発赤", "2: 真皮までの損傷", "3: 皮下組織までの損傷", "4: 皮下組織を越える損傷", "5: 関節腔、体腔に至る損傷"),\
			"浸出液": ("0: なし", "1: 少量:毎日のドレッシング交換を要しない", "3:  中等量:1日1回のドレッシング交換を要する", "6: 量:1日2回以上のドレッシング交換を要する"),\
			"大きさ": ("0: 皮膚損傷なし", "3: 4未満", "6: 4以上16未満", "8: 16以上36未満", "9: 36以上64未満", "12: 64以上100未満", "15: 100以上"),\
			"炎症・感染": ("0: 局所の炎症徴候なし", "1: 局所の炎症徴候あり(創周囲の発赤、腫脹、熱感、疼痛)", "3: 局所の明らかな感染徴候あり(炎症徴候、膿、悪臭など)", "9: 全身的影響あり(発熱など)"),\
			"肉芽形成": ("0: 治癒あるいは創が浅いため肉芽形成の評価ができない", "1: 良性肉芽が創面の90%以上を占める", "3: 良性肉芽が創面の50%以上90%未満を占める", "4: 良性肉芽が、創面の10%以上50%未満を占める", "5: 良性肉芽が、創面の10%未満を占める", "6: 良性肉芽が全く形成されていない"),\
			"壊死組織": ("0: 壊死組織なし", "3: 柔らかい壊死組織あり", "6: 硬く厚い密着した壊死組織あり"),\
			"ポケット": ("0: ケットなし", "6: 4未満", "9: 4以上16未満", "12: 16以上36未満", "24: 36以上")}
	def setSheet(self, sheet):  # 逐次変化する値。
		self.sheet = sheet
		cellranges = sheet[:, self.daycolumn].queryContentCells(CellFlags.STRING+CellFlags.VALUE+CellFlags.DATETIME)  # ID列の文字列、数値、日付が入っているセルに限定して抽出。
		self.emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # 日付列の最終行インデックス+1を取得。
		cellranges = sheet[self.splittedrow-1, :].queryContentCells(CellFlags.STRING)  # 文字列が入っているセルに限定して抽出。
		self.emptycolumn = cellranges.getRangeAddresses()[-1].EndColumn + 1  # 分割行の上行の最終列インデックス+1を取得。
VARS = IDsheet()
def activeSpreadsheetChanged(activationevent, xscriptcontext):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。



	# 本日の行までの数値をコピーする。
	
	
	

	pass
def mousePressed(enhancedmouseevent, xscriptcontext):  # マウスボタンを押した時。controllerにコンテナウィンドウはない。
	if enhancedmouseevent.ClickCount==2 and enhancedmouseevent.Buttons==MouseButton.LEFT:  # 左ダブルクリックの時。まずselectionChanged()が発火している。
		selection = enhancedmouseevent.Target  # ターゲットのセルを取得。
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			celladdress = selection.getCellAddress()
			r, c = celladdress.Row, celladdress.Column  # selectionの行インデックスと列インデックスを取得。
			if r==0:  # 行インデックス0の時。
				txt = selection.getString()
				if txt=="一覧へ":	
					doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 				
					doc.getCurrentController().setActiveSheet(doc.getSheets()[txt])  # 一覧シートをアクティブにする。
				elif txt=="次月更新":
					
					
					
					pass
				elif txt=="部位追加":
					if (c-VARS.startcolumn)%8==0:  # 部位の先頭列であることを確認する。
						sheet = selection.getSpreadsheet()
						datarows = [("", "", "", "", "", "", "", "", "部位追加")]
						datarows.append(list(VARS.dic.keys()))
						datarows[-1].extend(("部位別合計", "")) 		
						endedge = c + len(datarows[0]) - 1
						sheet[:VARS.splittedrow, c:endedge+1].setDataArray(datarows)
						sheet[0, c:endedge].getColumns().setPropertyValue("Width", 680)
						sheet[0, c:endedge].merge(True)
						sheet[0, endedge].setString("部位追加")	
						VARS.setSheet(sheet)  # 逐次変化する値を取得し直す。
					else:  # 部位の先頭列でないときはエラーメッセージを出す。
						msg = "部位の先頭列ではありません。"
						controller = xscriptcontext.getDocument().getCurrentController()  # コントローラの取得。
						commons.showErrorMessageBox(controller, msg)
				elif c==VARS.daycolumn:  # IDセルの時。IDをコピーする。
					ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
					smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
					systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。
					systemclipboard.setContents(commons.TextTransferable(txt), None)  # クリップボードにIDをコピーする。
				elif VARS.startcolumn<=c<VARS.emptycolumn:
					if (c-VARS.startcolumn)%8==0:  # 部位の先頭列の時。部位が入る。上下左右はコンテクストメニューで追加する。
						defaultrows = "肩", "腰椎部", "仙骨部", "坐骨部", "大転子部", "腓骨部", "腓腹部", "足関節外側", "踵", 
						staticdialog.createDialog(enhancedmouseevent, xscriptcontext, "部位選択", defaultrows)  # 列ヘッダー毎に定型句ダイアログを作成。
						selection.setPropertyValue("HoriJustify", CENTER)
			elif VARS.splittedrow<=r<VARS.emptyrow and VARS.startcolumn<=c<VARS.emptycolumn:  # 点数セルの時。
				if (c-VARS.startcolumn)%8!=7:  # 部位の最終行以外の時。
					headertxt = VARS.sheet[VARS.splittedrow-1, c].getString()
					defaultrows = VARS.dic.get(headertxt, None)
					gridcontrol1, datarows = staticdialog.createDialog(enhancedmouseevent, xscriptcontext, headertxt, defaultrows, callback=callback_wClickPoints)  # 列ヘッダー毎に定型句ダイアログを作成。	
					txt = "{}:".format(enhancedmouseevent.Target.getString())  # セルの入っている数字を文字列で取得。
					for i, d in enumerate(datarows):
						if d[0].startswith(txt):
							gridcontrol1.selectRow(i)  # 先頭が一致するグリッドコントロールの行をハイライト。
							break	
		return False  # セル編集モードにしない。	
	return True  # セル編集モードにする。	シングルクリックは必ずTrueを返さないといけない。	
def callback_wClickPoints(mouseevent, xscriptcontext):
	selection = xscriptcontext.getDocument().getCurrentSelection()  # シート上で選択しているオブジェクトを取得。
	selection.setValue(int(selection.getString().split(":", 1)[0]))  # 点数のみにして数値としてセルに代入し直す。
	selection.setPropertyValue("CellBackColor", -1)  # セルの背景色をクリアする。
def selectionChanged(eventobject, xscriptcontext):  # 矢印キーでセル移動した時も発火する。
	selection = eventobject.Source.getSelection()
	if selection.supportsService("com.sun.star.sheet.SheetCellRange"):  # 選択範囲がセル範囲の時。
		VARS.setSheet(selection.getSpreadsheet())			
		drowBorders(selection)  # 枠線の作成。			
def drowBorders(selection):  # ターゲットを交点とする行列全体の外枠線を描く。
	rangeaddress = selection.getRangeAddress()  # 選択範囲のセル範囲アドレスを取得。	
	sheet = VARS.sheet
	noneline, tableborder2, topbottomtableborder, leftrighttableborder = commons.createBorders()
	sheet[:, :].setPropertyValue("TopBorder2", noneline)  # 1辺をNONEにするだけですべての枠線が消える。	
	startrow = VARS.splittedrow if rangeaddress.StartRow<VARS.splittedrow else rangeaddress.StartRow
	edgerow = rangeaddress.EndRow+1 if rangeaddress.EndRow<VARS.emptyrow else VARS.emptyrow
	edgecolmun = rangeaddress.EndColumn+1 if rangeaddress.EndColumn<VARS.emptycolumn else VARS.emptycolumn
	if startrow<edgerow and rangeaddress.StartColumn<edgecolmun:
		sheet[startrow:edgerow, :VARS.emptycolumn].setPropertyValue("TableBorder2", topbottomtableborder)  # 行の上下に枠線を引く
		sheet[VARS.splittedrow-1:VARS.emptyrow, rangeaddress.StartColumn:edgecolmun].setPropertyValue("TableBorder2", leftrighttableborder)  # 列の左右に枠線を引く。
		selection.setPropertyValue("TableBorder2", tableborder2)  # 選択範囲の消えた枠線を引き直す。		
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
	r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。		
	if contextmenuname=="cell":  # セルのとき	
		if selection.supportsService("com.sun.star.sheet.SheetCell"):  # ターゲットがセルの時。
			if r==0 and (c-VARS.startcolumn)%8==0:  # 部位セルの時。
				addMenuentry("ActionTrigger", {"Text": "左", "CommandURL": baseurl.format("entry2")}) 		
				addMenuentry("ActionTrigger", {"Text": "右", "CommandURL": baseurl.format("entry3")}) 		
				addMenuentry("ActionTrigger", {"Text": "左右なし", "CommandURL": baseurl.format("entry8")}) 		
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
				addMenuentry("ActionTrigger", {"Text": "--上", "CommandURL": baseurl.format("entry4")}) 		
				addMenuentry("ActionTrigger", {"Text": "--下", "CommandURL": baseurl.format("entry5")}) 		
				addMenuentry("ActionTrigger", {"Text": "--左", "CommandURL": baseurl.format("entry6")}) 		
				addMenuentry("ActionTrigger", {"Text": "--右", "CommandURL": baseurl.format("entry7")}) 
				addMenuentry("ActionTrigger", {"Text": "--なし", "CommandURL": baseurl.format("entry9")}) 
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
				commons.cutcopypasteMenuEntries(addMenuentry)					
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
				addMenuentry("ActionTrigger", {"Text": "セル内容をクリア", "CommandURL": baseurl.format("entry1")}) 	
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
				addMenuentry("ActionTrigger", {"Text": "この部位を削除", "CommandURL": baseurl.format("entry12")}) 
			elif VARS.startcolumn<=c<VARS.emptycolumn:  # 点数セルの時。
				thisc = c-(c-VARS.startcolumn)%8	
				ptxt = VARS.sheet[0, thisc].getString()
				if VARS.splittedrow<r<VARS.emptyrow:
					addMenuentry("ActionTrigger", {"Text": "{} 開始日にする".format(ptxt), "CommandURL": baseurl.format("entry10")})
					addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
				if VARS.splittedrow<=r<VARS.emptyrow-1:
					addMenuentry("ActionTrigger", {"Text": "{} 終了日にする".format(ptxt), "CommandURL": baseurl.format("entry11")})	
	elif contextmenuname=="sheettab":  # シートタブの時。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Remove"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:RenameTable"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
	return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。	
def contextMenuEntries(entrynum, xscriptcontext):  # コンテクストメニュー番号の処理を振り分ける。引数でこれ以上に取得できる情報はない。	
	controller = xscriptcontext.getDocument().getCurrentController()  # コントローラの取得。
	selection = controller.getSelection()  # 選択範囲を取得。
	if entrynum==1:  # セル内容をクリア。書式設定とオブジェクト以外を消去。
		selection.clearContents(CellFlags.STRING+CellFlags.VALUE+CellFlags.DATETIME)  # セル内容をクリアする。	
	elif entrynum in (2, 3, 8):  # 前に左右をつける。
		if entrynum==2:  # 左。
			prefix = "左"
		elif entrynum==3:  # 右。
			prefix = "右"			
		elif entrynum==8:  # 左右なし
			prefix = ""
		selection.setString("{}{}".format(prefix, selection.getString().lstrip("左右")))
	elif entrynum in (4, 5, 6, 7, 9):  # 後に上下左右をつける。	
		if entrynum==4:  # 上。
			suffix = "上"			
		elif entrynum==5:  # 下。
			suffix = "下"			
		if entrynum==6:  # 左。
			suffix = "左"			
		elif entrynum==7:  # 右。
			suffix = "右"	
		elif entrynum==9:  # --なし
			suffix = ""				
		selection.setString("{}{}".format(selection.getString().rstrip("上下左右"), suffix))
	elif entrynum in (10, 11):
		sheet = VARS.sheet
		celladdress = selection.getCellAddress()
		r, c = celladdress.Row, celladdress.Column  # selectionの行インデックスと列インデックスを取得。
		thisc = c-(c-VARS.startcolumn)%8  # 部位の開始列インデックスを取得。	
		ptxt = sheet[0, thisc].getString()
		txt = sheet[r, VARS.daycolumn].getString()	
		stxt = txt.split(txt[4])
		datetxt = "{}月{}日".format(stxt[1].lstrip("0"), stxt[2].lstrip("0")) if len(stxt)==3 else ""	
		if entrynum==10:  # 開始日にする。これより上行をクリアする
			msg = "部位: {}\n{}より前の点数をクリアします。\n元には戻せません。".format(ptxt, datetxt)
			if showWarningMessageBox(controller, msg)==MessageBoxResults.OK:		
				datarange = sheet[VARS.splittedrow:r, thisc:thisc+8]
				datarange.setPropertyValue("CellBackColor", commons.COLORS["silver"])  # 背景色をつける
				datarange.clearContents(CellFlags.STRING+CellFlags.VALUE+CellFlags.DATETIME)  # セル内容をクリアする。	
		elif entrynum==11:  # 終了日にする。これより下の行をクリア。
			msg = "部位: {}\n{}より後の点数をクリアします。\n元には戻せません。".format(ptxt, datetxt)
			if showWarningMessageBox(controller, msg)==MessageBoxResults.OK:		
				datarange = sheet[r+1:VARS.emptyrow, thisc:thisc+8]
				datarange.setPropertyValue("CellBackColor", commons.COLORS["silver"])  # 背景色をつける	
				datarange.clearContents(CellFlags.STRING+CellFlags.VALUE+CellFlags.DATETIME)  # セル内容をクリアする。	
		clearCellBackColor(thisc)		
	elif entrynum==12:  # この部位を削除
		msg = "この部位({})をすべて削除します。\n元には戻せません。".format(selection.getString())
		if showWarningMessageBox(controller, msg)==MessageBoxResults.OK:			
			celladdress = selection.getCellAddress()
			c = celladdress.Column  # selectionの列インデックスを取得。
			VARS.sheet.removeRange(VARS.sheet[:, c:c+8].getRangeAddress(), delete_columns )  # 列を削除。	
def showWarningMessageBox(controller, msg):	
	componentwindow = controller.ComponentWindow
	msgbox = componentwindow.getToolkit().createMessageBox(componentwindow, WARNINGBOX, MessageBoxButtons.BUTTONS_OK_CANCEL+MessageBoxButtons.DEFAULT_BUTTON_CANCEL, "myRs", msg)
	return msgbox.execute()
def clearCellBackColor(c):  # 列インデックスcのある部位の値のあるセルの背景色をクリアする。
	searchdescriptor = VARS.sheet.createSearchDescriptor()
	searchdescriptor.setPropertyValue("SearchRegularExpression", True)  # 正規表現を有効にする。
	searchdescriptor.setSearchString("[:digit:]+")  # 戻り値はない。数値の入っているセルを検出。
	cellranges = VARS.sheet[VARS.splittedrow:VARS.emptyrow, c:c+8].findAll(searchdescriptor)  # 見つからなかった時はNoneが返る。
	if cellranges:
		cellranges.setPropertyValue("CellBackColor", -1)  # 数値の入っているセルの背景色をクリアする。
def changesOccurred(changesevent, xscriptcontext):  # Sourceにはドキュメントが入る。	
	selection = None
	for change in changesevent.Changes:
		if change.Accessor=="cell-change":  # セルの値が変化した時。マクロで変更したときはセル範囲が入ってくる時がある。
			selection = change.ReplacedElement  # 値を変更したセルを取得。
			break
	if selection and selection.supportsService("com.sun.star.sheet.SheetCell"):  # セルの時。
		celladdress = selection.getCellAddress()
		r, c = celladdress.Row, celladdress.Column  # selectionの行と列のインデックスを取得。		
		if VARS.splittedrow<=r<VARS.emptyrow and VARS.startcolumn<=c<VARS.emptycolumn:  # 点数セルの時。
			sheet = VARS.sheet
			datarange = sheet[r, :VARS.emptycolumn]
			datarow = datarange.getDataArray()[0]
			thisc = c-(c-VARS.startcolumn)%8  # 部位の開始列インデックスを取得。
			datarow[thisc+7] = sum(datarow[thisc:thisc+7])
			datarow[VARS.daycolumn-1] = min(datarow[i] for i in range(VARS.startcolumn+7, VARS.emptycolumn, 8))
			datarange.setDataArray((datarow,))
		
		
			prevs = sheet[VARS.splittedrow, :VARS.daycolumn-1].getDataArray()[0]
		
		
		
		# 変化した日の計算値を入力する。
		
		# マイナス日の最低値の背景色を変える。
		

		
		
# 		
# 		
# 		offdayc = VARS.templatestartcolumn - 1  # 休日設定のある列インデックスを取得。
# 		if VARS.datarow<=r<VARS.emptyrow:  # 予定セルまたはテンプレートセルのある行の時。
# 			if VARS.datacolumn-1<c<VARS.firstemptycolumn or offdayc<c<VARS.templateendcolumnedge:  # 予定セルまたはテンプレートセルのある列の時。
# 				setCellProp(selection)
# 		elif celladdress.Column==offdayc and selection.getValue()>0:  # 選択セルが休日設定のある列、かつ、選択セルに0より大きい数値が入っている。の時。 
# 			sheet = selection.getSpreadsheet()
# 			searchdescriptor = sheet.createSearchDescriptor()
# 			searchdescriptor.setSearchString("休日設定")  # 戻り値はない。
# 			searchedcell = sheet[VARS.emptyrow:, offdayc].findFirst(searchdescriptor)  # 休日設定の開始セルを取得。見つからなかった時はNoneが返る。
# 			if searchedcell:  # 休日設定の開始セルがある時。
# 				if celladdress.Row>searchedcell.getCellAddress().Row+1:  # 休日設定の開始行より下の時。
# 					selection.setPropertyValues(("NumberFormat", "HoriJustify"), (commons.formatkeyCreator(xscriptcontext.getDocument())('YYYY-M-D'), LEFT))
		


