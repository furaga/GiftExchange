# 以下のページを参考に実装
# https://davis.gfd-dennou.org/rubygadgets/ja/?(Tips)+PowerPoint+%A5%D5%A5%A1%A5%A4%A5%EB%A4%CE%BC%AB%C6%B0%C0%B8%C0%AE
require 'win32ole'

Encoding.default_external = 'UTF-8'
output = "GiftExchange.ppt"

# 参加者の名前を読み込み
f = open("nameList.txt")
names = f.read.split(/(\n|\r)/).select { |line| !line.match(/\s/) }
randNames = names.sort_by {rand} # シャッフル
f.close

def getAbsolutePath(path)
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  return fso.GetAbsolutePathname(path)
end

# PowerPoint アプリケーションを開く
pp =  WIN32OLE.new('PowerPoint.Application')
# 既にファイルを開いている場合は処理をしない
if pp.Presentations.Count > 0 then
  puts "ERROR: Close all PowerPoint file and rerun this script."
  exit 1
end
# アプリケーションウィンドウを表示
pp.Visible = true

# スライドファイルの新規作成
presen = pp.Presentations.add

def genNewSlide(presen, title, text)
	p ipage = presen.Slides.Count
	slidewidth  = presen.pageSetup.slideWidth
	slideheight = presen.pageSetup.slideHeight

	# 新しいスライド作成
	newSlide = presen.Slides.Add(ipage + 1, 11)
	slideShape = newSlide.Shapes
	
	# タイトル
	ttr = slideShape.Title.TextFrame.TextRange
	ttr.Text = title
	ttr.font.size = 60
	ttr.ParagraphFormat.Alignment = 1

	# 本文
	if text == "" then
		return slideShape
	end	
	width = slidewidth
	height = 100
	orientation = 1
	tb = slideShape.addTextBox(orientation, 0, (slideheight - height) / 2, width, height)
	tr = tb.textFrame.textRange
	tr.text = text
	tr.font.size = height
	tr.ParagraphFormat.Alignment = 2

	return slideShape
end

# 1でサムネイルビューを隠す。9で通常表示
pp.activewindow.viewtype = 1

for i in 0...randNames.length
	title = randNames[i] + "から"
	text = randNames[(i + 1) % randNames.length] + "へ"
	genNewSlide(presen, title, "")
	genNewSlide(presen, title, text)
end

output = getAbsolutePath(output)
presen.SaveAs(output)