#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
import platform
# ドキュメントイベントについて。
def documentOnLoad(xscriptcontext):  # ドキュメントを開いた時。リスナー追加前。
	doc = xscriptcontext.getDocument()  # ドキュメントのモデルを取得。 
	controller = doc.getCurrentController()  # コントローラの取得。
	sheets = doc.getSheets()
	if platform.system()=="Windows":  # Windowsの時はすべてのシートのフォントをMS Pゴシックにする。
		[i.setPropertyValues(("CharFontName", "CharFontNameAsian"), ("ＭＳ Ｐゴシック", "ＭＳ Ｐゴシック")) for i in sheets]
	controller.setActiveSheet(sheets["一覧"])  # 一覧シートをアクティブにする。
	namedranges = doc.getPropertyValue("NamedRanges")  # ドキュメントのNamedRangesを取得。
	for i in namedranges.getElementNames():  # namedrangesをイテレートするとfor文中でnamedrangesを操作してはいけない。
		if not namedranges[i].getReferredCells():
			namedranges.removeByName(i)  # 参照範囲がエラーの名前を削除する。	
def documentUnLoad(xscriptcontext):  # ドキュメントを閉じた時。リスナー削除後。
	pass
