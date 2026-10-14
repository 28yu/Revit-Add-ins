# v2.6 の新機能9個のアイコンを 96x96 / 288DPI で描き直すスクリプト
# 使い方: python3 draw_v26_icons_96.py <出力フォルダ>
# ⚠ 288DPI（pHYs）が無いと WPF でリボンからはみ出す（Docs/DEVLOG.md 参照）
import math, sys
from PIL import Image, ImageDraw
OUT=sys.argv[1]
import os; os.makedirs(OUT, exist_ok=True)
S=24  # 32単位 → 768px で描いて 96px に縮小（アンチエイリアス）
def canvas():
    im=Image.new('RGBA',(32*S,32*S),(0,0,0,0)); return im, ImageDraw.Draw(im)
def p(v): return v*S
def rect(d,x0,y0,x1,y1,fill=None,outline=None,w=1.0,r=0):
    d.rounded_rectangle([p(x0),p(y0),p(x1),p(y1)],radius=p(r),fill=fill,outline=outline,width=max(1,round(p(w))))
def line(d,pts,col,w):
    d.line([(p(x),p(y)) for x,y in pts],fill=col,width=round(p(w)),joint='curve')
    for x,y in pts: d.ellipse([p(x-w/2),p(y-w/2),p(x+w/2),p(y+w/2)],fill=col)
def poly(d,pts,col): d.polygon([(p(x),p(y)) for x,y in pts],fill=col)
def save(im,name):
    im=im.resize((96,96),Image.LANCZOS); im.save(f'{OUT}/{name}.png',dpi=(288,288))

def auto_backup():
    im,d=canvas(); c='#505050'
    for k in range(12):
        a=math.radians(k*30); s,co=math.sin(a),-math.cos(a)
        line(d,[(16+10.6*s,16+10.6*co),(16+15*s,16+15*co)],c,1.6)
    line(d,[(10.8,15.8),(15,19.8),(22.6,11.2)],'#404040',2.2); save(im,'auto_backup')

def dwg_layer_transfer():
    im,d=canvas()
    for x0,col in ((0.3,'#67b1e6'),(20.4,'#a9d2f0')):
        rect(d,x0,4.5,x0+11.2,27,fill=col,r=0.6)
        for y in (11,15.6,20.2): rect(d,x0+1.9,y-0.8,x0+9.3,y+0.8,fill='#ffffff')
    rect(d,12.6,14.3,17.2,17,fill='#7fbde9'); poly(d,[(16.8,12.2),(20.2,15.65),(16.8,19.1)],'#7fbde9')
    save(im,'dwg_layer_transfer')

def filter_management():
    im,d=canvas()
    poly(d,[(4,4.5),(28,4.5),(19.8,18.2),(19.8,24.6),(12.3,29),(12.3,18.2)],'#67b1e6')
    poly(d,[(4,4.5),(28,4.5),(26.5,7),(5.5,7)],'#5a9fd1'); save(im,'filter_management')

def generic_model_merge():
    im,d=canvas(); b='#0066cc'
    for x in (1.4,7.4):
        for y in (10.4,16.6):
            rect(d,x,y,x+4.8,y+4.8,fill='#ffffff',outline=b,w=1.1)
            rect(d,x+1.6,y+1.6,x+3.2,y+3.2,fill='#81b6eb')
    rect(d,13,15.0,17.6,16.4,fill='#1d5fa8'); poly(d,[(17.2,13.4),(19.8,15.7),(17.2,18)],'#1d5fa8')
    rect(d,20.6,9.6,31,22.2,fill=b,outline='#005bb5',w=0.9); save(im,'generic_model_merge')

def parameter_cleanup():
    im,d=canvas()
    rect(d,4,1.8,28,29.6,fill='#ffffff',outline='#4f93e0',w=1.0,r=1)
    for y,col in ((8.9,'#787878'),(13.7,'#787878'),(18.9,'#e06666'),(24.3,'#606060')):
        rect(d,6,y-0.55,24.8,y+0.55,fill=col)
    save(im,'parameter_cleanup')

def doc(d,x0,y0,x1,y1,col,lines,lcol,w=1.3):
    rect(d,x0,y0,x1,y1,fill='#ffffff',outline=col,w=w,r=1)
    for y in lines: rect(d,x0+2.2,y-0.55,x1-2.2,y+0.55,fill=lcol)

def parameter_value_move():
    im,d=canvas()
    doc(d,1.4,7,12,25,'#848484',(12,16,20),'#888888')
    doc(d,19.3,7,30.4,25,'#1673d0',(12,16,20),'#1673d0')
    rect(d,11.5,14.5,17.4,17.2,fill='#f08a24'); poly(d,[(17,12.4),(21,15.85),(17,19.3)],'#f08a24')
    save(im,'parameter_value_move')

def parameter_value_replace():
    im,d=canvas()
    doc(d,1.5,5,12.6,15.8,'#8a8a8a',(10.4,),'#707070')
    doc(d,18.3,16.4,29.5,26.4,'#1673d0',(21.4,),'#1673d0')
    d.arc([p(8.5),p(8.4),p(25.1),p(25)],start=-90,end=-8,fill='#f08a24',width=round(p(2.0)))
    poly(d,[(21.9,13.6),(27.3,13.6),(24.6,17.8)],'#f08a24'); save(im,'parameter_value_replace')

def room_3d_color():
    im,d=canvas()
    for (x,y),col in (((3,3),'#90afc5'),((17,3),'#a5c498'),((3,17),'#d2a578'),((17,17),'#b496b4')):
        rect(d,x,y,x+12,y+12,fill=col,outline='#5a5a5a',w=1.2)
    save(im,'room_3d_color')

def view_template():
    im,d=canvas()
    rect(d,9.2,4,27.6,24.8,fill='#bae0fc',r=1)
    rect(d,6.6,6,25.5,27,fill='#4996cc',r=1)
    rect(d,4.1,8,23.4,29,fill='#1d7cbf',r=1)
    for x1,y,col in ((17.6,12,'#ffffff'),(15,16.4,'#a9d0ee'),(19,20,'#ffffff'),(15,23.6,'#ffffff')):
        rect(d,7,y-0.6,x1,y+0.6,fill=col)
    save(im,'view_template')

for f in (auto_backup,dwg_layer_transfer,filter_management,generic_model_merge,parameter_cleanup,
          parameter_value_move,parameter_value_replace,room_3d_color,view_template): f()
