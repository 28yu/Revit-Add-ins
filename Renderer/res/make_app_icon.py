# 描画プログラム（Tools28Render.exe）と、書き出したビューア exe のアイコンを作るスクリプト。
# 28 Tools のサイトのファビコン（favicon_28.png、256×256・背景透明）を、Windows のアイコン（16〜256 画素）にする。
# 使い方: python3 make_app_icon.py  → 同じフォルダに Tools28Render.ico を作る（res/app.rc が exe に入れる）
import os
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
im = Image.open(os.path.join(HERE, 'favicon_28.png')).convert('RGBA')
sizes = [16, 24, 32, 48, 64, 128, 256]
im.resize((256, 256), Image.LANCZOS).save(os.path.join(HERE, 'Tools28Render.ico'), sizes=[(s, s) for s in sizes])
print('ok')
