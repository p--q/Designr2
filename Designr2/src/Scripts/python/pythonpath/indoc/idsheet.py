#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# IDシートについて。import pydevd; pydevd.settrace(stdoutToServer=True, stderrToServer=True)
from indoc import commons, datedialog, staticdialog
from com.sun.star.awt import MouseButton  # 定数
from com.sun.star.sheet import CellFlags  # 定数
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
			"壊死組織": ("0: 死組織なし", "3: 柔らかい壊死組織あり", "6: 硬く厚い密着した壊死組織あり"),\
			"ポケット": ("0: ケットなし", "6: 4未満", "9: 4以上16未満", "12: 16以上36未満", "24: 36以上")}
	def setSheet(self, sheet):  # 逐次変化する値。
		self.sheet = sheet
		cellranges = sheet[:, self.daycolumn].queryContentCells(CellFlags.STRING+CellFlags.VALUE+CellFlags.DATETIME)  # ID列の文字列、数値、日付が入っているセルに限定して抽出。
		self.emptyrow = cellranges.getRangeAddresses()[-1].EndRow + 1  # 日付列の最終行インデックス+1を取得。
		cellranges = sheet[self.splittedrow-1, :].queryContentCells(CellFlags.STRING)  # ID列の文字列が入っているセルに限定して抽出。
		self.emptycolumn = cellranges.getRangeAddresses()[-1].EndColumn + 1  # 分割行の上行の最終列インデックス+1を取得。
VARS = IDsheet()
def activeSpreadsheetChanged(activationevent, xscriptcontext):  # シートがアクティブになった時。ドキュメントを開いた時は発火しない。よく誤入力されるセルを修正する。つまりボタンになっているセルの修正。


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
					
					
					datarows = list(VARS.dic.keys()).append("部位別合計"),
					datarows.append(["0"]*7)
					
					
					
				elif c==VARS.daycolumn:  # IDセルの時。IDをコピーする。
					ctx = xscriptcontext.getComponentContext()  # コンポーネントコンテクストの取得。
					smgr = ctx.getServiceManager()  # サービスマネージャーの取得。
					systemclipboard = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)  # SystemClipboard。クリップボードへのコピーに利用。
					systemclipboard.setContents(commons.TextTransferable(txt), None)  # クリップボードにIDをコピーする。
				elif VARS.startcolumn<=c<VARS.emptycolumn:
					if (c-VARS.startcolumn)%8==0:  # 部位の先頭列の時。部位が入る。上下左右はコンテクストメニューで追加する。
						defaultrows = "肩", "腰椎部", "仙骨部", "坐骨部", "大転子部", "腓骨部", "腓腹部", "足関節外側", "踵", 
						staticdialog.createDialog(enhancedmouseevent, xscriptcontext, "部位選択", defaultrows)  # 列ヘッダー毎に定型句ダイアログを作成。
					elif (c-VARS.startcolumn)%8==7:  # 開始日セルの時。
						datedialog.createDialog(enhancedmouseevent, xscriptcontext, "開始日", "D")	
						
						# 開始日より上をクリア
						
			elif VARS.splittedrow<=r<VARS.emptyrow and VARS.startcolumn<=c<VARS.emptycolumn:  # 点数セルの時。
				if (c-VARS.startcolumn)%8!=7:  # 部位の最終行以外の時。
					headertxt = VARS.sheet[VARS.splittedrow-1, c].getString()
					defaultrows = VARS.dic.get(headertxt, None)
					staticdialog.createDialog(enhancedmouseevent, xscriptcontext, headertxt, defaultrows, callback=callback_wClickPoints)  # 列ヘッダー毎に定型句ダイアログを作成。
# 			elif r==VARS.splittedrow and c==VARS.daycolumn:  # 日付列の先頭行の時。
# 				datedialog.createDialog(xscriptcontext, enhancedmouseevent, "月の選択", "YYYY-MM-DD")		
		return False  # セル編集モードにしない。	
	return True  # セル編集モードにする。	シングルクリックは必ずTrueを返さないといけない。	
def callback_wClickPoints(mouseevent, xscriptcontext):
	selection = xscriptcontext.getDocument().getCurrentSelection()  # シート上で選択しているオブジェクトを取得。	
	selection.setValue(int(selection.getString().split(":", 1)[0]))
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
				addMenuentry("ActionTrigger", {"Text": "クリア", "CommandURL": baseurl.format("entry1")}) 	
			elif VARS.splittedrow<=r<VARS.emptyrow and VARS.startcolumn<=c<VARS.emptycolumn:  # 点数セルの時。
				addMenuentry("ActionTrigger", {"Text": "開始日にする", "CommandURL": baseurl.format("entry10")})
				addMenuentry("ActionTriggerSeparator", {"SeparatorType": ActionTriggerSeparatorType.LINE})  # セパレーターを挿入。
				addMenuentry("ActionTrigger", {"Text": "終了日にする", "CommandURL": baseurl.format("entry11")})	
	elif contextmenuname=="sheettab":  # シートタブの時。
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Remove"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:RenameTable"})
		addMenuentry("ActionTrigger", {"CommandURL": ".uno:Move"})
	return EXECUTE_MODIFIED  # このContextMenuInterceptorでコンテクストメニューのカスタマイズを終わらす。	
def contextMenuEntries(entrynum, xscriptcontext):  # コンテクストメニュー番号の処理を振り分ける。引数でこれ以上に取得できる情報はない。	
	controller = xscriptcontext.getDocument().getCurrentController()  # コントローラの取得。
	selection = controller.getSelection()  # 選択範囲を取得。
	if entrynum==1:  # クリア。書式設定とオブジェクト以外を消去。
		selection.clearContents(511)  # 範囲をすべてクリアする。	
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
	elif entrynum==10:  # 開始日にする
		
		# 開始日セルの変更。
		# これより上の行をクリア。
		
		
		
		
		pass
	elif entrynum==11:  # 終了日にする
		

		# これより下の行をクリア。
		
		
		
		
		pass	

		
		
		