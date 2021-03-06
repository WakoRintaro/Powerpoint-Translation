Powerpoint-Translation
====

Overview
Google TranslateとYahooルビ振りAPIを組み合わせたパワーポイント翻訳ツールです。  
スライド中にある漢字を含む文を抜き出してルビ，英語訳，英語訳をそのスライドのノートに書きます。

## Description
models.py - パワポ翻訳機能をモジュール化したものです。  
transpptx.py - 実行ファイルです。

## Demo
スライド中に「今日の演習」という文字列がある場合，ノートには下記のように表示されます。  
例）  
今日の演習 / (きょう)の(えんしゅう)
Exercise of the day
ออกกำลังกายประจำวัน

## Requirement
Python3系で動くと思います。  
少なくともPython 3.6.9，3.7.3では動作しています。  
その他ライブラリが必要です。

## Usage
引数で翻訳対象のパワポのファイルパスを渡します。  
翻訳済みのパワポは./pptx_translated/に保存されます。
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
```
rintarowako@W35-37ET:~/dev/transpptx$ python3 transpptx.py pptx/test.pptx 
```

## Licence
本システムの翻訳にはGoogle Translate, Yahoo APIを利用しています。それぞれの規約に触れないように使用してください。
このパワポ翻訳システム自体は教育目的に限り自由に改変して使って頂いて結構です。


## Author
Rintaro Wako