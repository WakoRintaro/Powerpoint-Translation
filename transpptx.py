import sys
from models import Transpptx

if __name__ == "__main__":
    # 引数に翻訳対象のパワポまでのパスを
    args = sys.argv

    # 翻訳対象のパワポ
    target = args[1]
    # 保存先のフォルダ
    save_to = "pptx_translated/"
    # 翻訳キャッシュ
    dict_csv = "dictionary.csv"

    # 翻訳対象のパワポ，保存先のフォルダ，キャッシュ場所を
    # 指定してインスタンスを生成
    transpptx= Transpptx(target, save_to, dict_csv)
    
    # 翻訳開始
    transpptx.pptx_translation()
    