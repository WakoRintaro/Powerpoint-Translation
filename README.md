Powerpoint-Translation
====

Overview

## Description
Google TranslateとYahooルビ振りAPIを組み合わせたパワーポイント翻訳ツールです。

## Demo
スライド中にある文字を抜き出して，ルビ，英語訳，タイ語訳をノートに表示します。
例）
今日の演習 / (きょう)の(えんしゅう)
Exercise of the day
ออกกำลังกายประจำวัน

## VS. 

## Requirement
Python3系で動きます。Python 3.6.9，3.7.3では動作しています。その他ライブラリが必要です。

## Usage
引数で翻訳対象のパワーポイントのファイル名を渡します。
```
python3 transpptx.py pptx/target.pptx 
```

## Install
ダウンロードして実行するまでの手順を示します。
ダウンロード先のディレクトリーを作成し，Githubからダウンロードします。
```
rintarowako@W35-37ET:~/dev$ mkdir transpptx
rintarowako@W35-37ET:~/dev$ cd transpptx/
rintarowako@W35-37ET:~/dev/transpptx$ git clone https://github.com/WakoRintaro/Powerpoint-Translation.git .
```

翻訳対象のパワポを保存するディレクトリと翻訳済みのパワポを保存するディレクトリを作成します。
翻訳キャッシュを保存するからのcsvファイルを作成します。
```
rintarowako@W35-37ET:~/dev/transpptx$ mkdir pptx
rintarowako@W35-37ET:~/dev/transpptx$ mkdir pptx_translated
rintarowako@W35-37ET:~/dev/transpptx$ touch dictionary.csv
```

python-pptxライブラリがインストールされていなければ，インストールします。
```
rintarowako@W35-37ET:~/dev/transpptx$ sudo pip3 install python-pptx
```

翻訳対象のパワポを指定して実行します。
'''
rintarowako@W35-37ET:~/dev/transpptx$ python3 transpptx.py pptx/test.pptx 
'''

## Licence
Google Translate, Yahoo APIの規約に従ってください。
このパワポ翻訳システム自体は教育目的であれば自由に改変して使って頂いて結構です。

## Author
Rintaro Wako