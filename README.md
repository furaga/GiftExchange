概要
---
プレゼント交換で誰から誰にプレゼントを渡すかランダムに決定し、PowerPointファイル(.ppt)として結果を出力するRubyスクリプト。  
以下のページを参考に実装しています。  
https://davis.gfd-dennou.org/rubygadgets/ja/?(Tips)+PowerPoint+%A5%D5%A5%A1%A5%A4%A5%EB%A4%CE%BC%AB%C6%B0%C0%B8%C0%AE

使い方  
---
1. nameList.txtに参加者の名前を列挙（一行あたり一人）  
2. run.batをダブルクリック。またはコマンドラインで以下のコマンドを実行  
ruby gen_ppt.rb　　
3. GiftExchange.pptが出力されてPowerPointが起動するので、下キーなどでスライドをめくりながらプレゼント交換をしてください

必要な環境
---
Windows上でのみ動作します（Windows 8上で動作確認）  
・Ruby  
・win32ole (Ruby 1.8以降なら標準でインストールされています)  
・PowerPoint

