# GUI関係
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox as mb

import csv

# PDF表示関係
# pip install Pillow
from PIL import Image, ImageTk  

import sys
from pathlib import Path
# pip install pdf2image
from pdf2image import convert_from_path
import os

# PDF画像編集関係
# import glob
# import PyPDF2

from pdfrw import PdfReader
from pdfrw.buildxobj import pagexobj
from pdfrw.toreportlab import makerl
from reportlab.pdfgen import canvas as canvas2
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import A4, landscape, portrait
from reportlab.lib.units import mm, cm

pdfpages = 8
now_page = 0
checkbtn = []

def pdf_changer():
    global checkbtn
    # popplerへの環境変数PATHを一時的に付与
    poppler_dir = Path(sys.argv[0]).parent.absolute() / "poppler-23.01.0/Library/bin"
    os.environ["PATH"] += os.pathsep + str(poppler_dir)
    ### PDFをpng画像に変更する ###
    path = Path(sys.argv[0]).parent.absolute() / "checksheet.pdf"
    pages = convert_from_path(Path(path),200)
    print(len(pages))
    for i in range(len(pages)):
        print(i)
        text = "image/image"+str(i)+".png"
        print(text)
        pages[i].save(Path(sys.argv[0]).parent.absolute() / text, "png")
    print("OK")
    pass

def TL_checksheet(hisou = "",souju= "", sousi = "",soute= ""):
    global checkbtn
    global now_page
    sub_win = tk.Tk()
    sub_win.geometry("800x1000+0+0")
    sub_win.configure(bg = "lightblue")
    sub_win.title("チェックん")
    if not os.path.isfile("image/image0.png"):
        pdf_changer()
    img = []
    canvas = []
    btn = []

    for i in range(pdfpages):
        # 画像の読み込み
        path = "image/image" + str(i) + ".png"
        img.append("")
        img[i] = Image.open(path)
        img[i] = img[i].resize((706,1000))
        # キャンバスを作成 #
        canvas.append("")
        
        canvas[i] = tk.Canvas(sub_win, width = 706, height = 1000, bg="white")

        # キャンバスを設置 #
        # canvas[i].place(x = 0 , y = 0)

        # キャンバスに描画 #
        canvas[i].Photo = ImageTk.PhotoImage(img[i])
        canvas[i].create_image(0, 0, image=canvas[i].Photo,anchor=tk.NW)

    switch0 = [[True,False,True,False, True,False, False,False,False,False,False,False, False,False,False,False, True,True,False,False,False],
               [False,True, True,True, False, False, False,False,False,False,False,False,False,False,False,False, False, False],
               [False,True,False,False,False,False, True,True,True,True,True,True,True, False,False,False,False,False,False,False,False],
               [False,False,False, False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False,False],
               [False,False, False, True,False,False,False,True,False,False,False,False,True,False,False,False],
               [False,False,True,True, False,False,False,False, False, False,False,False,False,False,False,False, False,False,False,False],
               [False,False,False, False,False,False,False, False,False,False,False,False,False,False],
               [False,False,False,False,False]]
    # ボタンのX座標
    switch0X = [[643,643, 643,643, 643,643, 643,643,643,643,643,643, 643,643,643,643, 643,643,643,643,643],
               [643,643, 643,643, 643, 643, 643,643,643,643,643,643,643,643,643,643, 643, 643],
               [643,643,643,643,643,643, 643,643,643,643,643,643,643, 643,643,643,643,643,643,643,643],
               [643,643,643, 643,643,643,643,643,643,643,643,643,643,643,643,643,643,643,643,643,643,643,643],
               [643,643, 643, 643,643,643,643,643,643,643,643,643,643,643,643,643],
               [643,643,643,643, 643,643,643,643, 643, 643,643,643,643,643,643,643, 643,643,643,643],
               [643,643,643, 643,643,643,643, 643,643,643,643,643,643,643],
               [643,643,643,643,643]]
        # ボタンのY座標
    switch0Base = [[535,620, 677,734, 790,847, 903,960,1017,1073,1130,1187, 1272,1385,1442,1527, 1583,1806,1895,1980,2036],
               [220,333, 361,418, 475, 531, 597,654,711,767,875,932,988,1102,1158,1243, 1328, 1497],
               [220,276,333,418,531,616, 673,758,814,871,928,984,1126, 1211,1296,1353,1466,1636,1720,1894,2012],
               [220,333,417, 531,590,649,705,814,984,1040,1069,1125,1239,1324,1381,1437,1494,1664,1720,1888,1947,2004,2117],
               [220,305, 417, 562,663,748,776,861,1026,1139,1309,1422,1618,1731,1870,1927],
               [220,446,513,569, 644,673,728,781, 838, 911,1034,1090,1175,1260,1373,1458, 1571,1712,1768,1876],
               [219,407,596, 731,875,988,1130, 1243,1300,1357,1442,1611,1781,1941],
               [220,332,472,612,752]]
    # ボタンのY座標
    switch0Y = [[535,620, 677,734, 790,847, 903,960,1017,1073,1130,1187, 1272,1385,1442,1527, 1583,1806,1895,1980,2036],
               [220,333, 361,418, 475, 531, 597,654,711,767,875,932,988,1102,1158,1243, 1328, 1497],
               [220,276,333,418,531,616, 673,758,814,871,928,984,1126, 1211,1296,1353,1466,1636,1720,1894,2012],
               [220,333,417, 531,590,649,705,814,984,1040,1069,1125,1239,1324,1381,1437,1494,1664,1720,1888,1947,2004,2117],
               [220,305, 417, 562,663,748,776,861,1026,1139,1309,1422,1618,1731,1870,1927],
               [220,446,513,569, 644,673,728,781, 838, 911,1034,1090,1175,1260,1373,1458, 1571,1712,1768,1876],
               [219,407,596, 731,875,988,1130, 1243,1300,1357,1442,1611,1781,1941],
               [220,332,472,612,752]]
    # 座標調整
    for j in range(pdfpages):
        for i in range(len(switch0Y[j])):
            switch0Y[j][i] = switch0Base[j][i]*0.4277-8
    # チェックボタン作成
    for j in range(pdfpages):
        checkbtn.append("")
        checkbtn[j] = []
        btn.append("")
        btn[j] = []
        for i in range(len(switch0[j])):
            checkbtn[j].append("")
            checkbtn[j][i] = tk.BooleanVar()
            btn[j].append("")
            btn[j][i] = tk.Checkbutton(sub_win, variable = checkbtn[j][i], text = "             ")
            checkbtn[j][i].set(switch0[j][i])
            
            # btn[j][i].place(x = 559, y = 30 + i*24)

    ent0y = [1750,1850,1900,1971,2026,1855,1980]
    ent0x = [324,478,903,1255]
    for i in range(len(ent0y)):
        ent0y[i] = ent0y[i]*0.4277
    for i in range(len(ent0x)):
        ent0x[i] = ent0x[i]*0.4277
    
    Ent_hisou = tk.Entry(sub_win, font = ("",16), width = 14)
    # Ent_hisou.place(x = ent0x[0], y = ent0y[0])

    Ent_souju1 = tk.Entry(sub_win, font = ("",10), width = 30)
    # Ent_souju1.place(x = ent0x[0], y = ent0y[1])

    Ent_souju2 = tk.Entry(sub_win, font = ("",10), width = 30)
    # Ent_souju2.place(x = ent0x[0], y = ent0y[2])

    Ent_sousi = tk.Entry(sub_win, font = ("",10), width = 20)
    # Ent_sousi.place(x = ent0x[0], y = ent0y[3])

    Ent_soute1 = tk.Entry(sub_win, font = ("",10), width = 8)
    # Ent_soute1.place(x = ent0x[0], y = ent0y[4])

    Ent_soute2 = tk.Entry(sub_win, font = ("",10), width = 8)
    # Ent_soute2.place(x = ent0x[1], y = ent0y[4])

    Ent_soute3 = tk.Entry(sub_win, font = ("",10), width = 8)
    # Ent_soute3.place(x = ent0x[1]+65.8658, y = ent0y[4])

    Ent_zeiju1 = tk.Entry(sub_win, font = ("",10), width = 30)
    # Ent_zeiju1.place(x = ent0x[2], y = ent0y[5])

    Ent_zeiju2 = tk.Entry(sub_win, font = ("",10), width = 30)
    # Ent_zeiju2.place(x = ent0x[2], y = ent0y[5]+21.385)

    Ent_zeisi = tk.Entry(sub_win, font = ("",10), width = 30)
    # Ent_zeisi.place(x = ent0x[2], y = ent0y[6])

    Ent_zeite1 = tk.Entry(sub_win, font = ("",8), width = 5)
    # Ent_zeite1.place(x = ent0x[3]-65.8658, y = ent0y[6])

    Ent_zeite2 = tk.Entry(sub_win, font = ("",8), width = 5)
    # Ent_zeite2.place(x = ent0x[3], y = ent0y[6])

    Ent_zeite3 = tk.Entry(sub_win, font = ("",8), width = 5)
    # Ent_zeite3.place(x = ent0x[3]+65.8658, y = ent0y[6])
    # ページ切り替えーすべて削除後ー該当のページを表示
    def page_change(p = 0):
        global now_page
        # すべて削除
        for i in range(pdfpages):
            for j in range(len(switch0[i])):
                btn[i][j].place_forget()
            canvas[i].place_forget()
        Ent_hisou.place_forget()
        Ent_souju1.place_forget()
        Ent_souju2.place_forget()
        Ent_sousi.place_forget()
        Ent_soute1.place_forget()
        Ent_soute2.place_forget()
        Ent_soute3.place_forget()
        Ent_zeiju1.place_forget()
        Ent_zeiju2.place_forget()
        Ent_zeisi.place_forget()
        Ent_zeite1.place_forget()
        Ent_zeite2.place_forget()
        Ent_zeite3.place_forget()
        # ページ最初と最後をつなげる
        if p == -1:
            now_page -= 1
            if now_page == -1:
                now_page = 7
        elif p == -2:
            now_page += 1
            if now_page == 8:
                now_page = 0
        else:
            now_page = p
        print(now_page)
        # 今のページを表示
        canvas[now_page].place(x = 0, y = 0)
        for i in range(len(switch0[now_page])):
            btn[now_page][i].place(x = switch0X[now_page][i], y = switch0Y[now_page][i])
            # checkbtn[0][0].set(switch0[now_page][i])
        # 最後のページの時いろいろ表示
        if now_page == 7:
            Ent_hisou.place(x = ent0x[0], y = ent0y[0])
            Ent_souju1.place(x = ent0x[0], y = ent0y[1])
            Ent_souju2.place(x = ent0x[0], y = ent0y[2])
            Ent_sousi.place(x = ent0x[0], y = ent0y[3])
            Ent_soute1.place(x = ent0x[0], y = ent0y[4])
            Ent_soute2.place(x = ent0x[1], y = ent0y[4])
            Ent_soute3.place(x = ent0x[1]+65.8658, y = ent0y[4])
            Ent_zeiju1.place(x = ent0x[2], y = ent0y[5])
            Ent_zeiju2.place(x = ent0x[2], y = ent0y[5]+21.385)
            Ent_zeisi.place(x = ent0x[2], y = ent0y[6])
            Ent_zeite1.place(x = ent0x[3]-35.8658, y = ent0y[4])
            Ent_zeite2.place(x = ent0x[3], y = ent0y[4])
            Ent_zeite3.place(x = ent0x[3]+35.8658, y = ent0y[4])
            pass
    page_change()


    print(checkbtn[0][0].get())
    def test():
        global now_page
        # checkbtn[0][0].set(True)
        for j in range(pdfpages):
            for i in range(len(switch0[0])):
                checkbtn[j][i].set(switch0[j][i])
    # test()
    xxx = 20
    tk.Button(sub_win, text = "次",font = ("",25), command = lambda:page_change(-2)).place(x =400-xxx,y = 5)
    tk.Button(sub_win, text = "戻",font = ("",25), command = lambda:page_change(-1)).place(x =300-xxx,y = 5)
    tk.Button(sub_win, text = "始",font = ("",25), command = lambda:page_change(0)).place(x =200-xxx,y = 5)
    tk.Button(sub_win, text = "終",font = ("",25), command = lambda:page_change(7)).place(x =500-xxx,y = 5)

    def GOGOGO():
        ret = mb.askyesno("確認","作成完了ですか？")
        if ret == True:
            print("OK")
            # reader1 = PyPDF2.PdfReader('checksheet.pdf')
            # writer = PyPDF2.PdfWriter()
            # writer.add_page(reader1.pages[0])
            # with open('checksheet2.pdf', 'wb') as f:
            #     writer.write(f)
            HisouName = Ent_hisou.get()
            if HisouName == "":
                out_path = "完成品/税理士法第３３条の２の書面添付に係るチェックシート.pdf"
            else:
                HisouName = HisouName.replace(" ","")
                HisouName = HisouName.replace("　","")
                out_path = "完成品/故"+HisouName+"様税理士法第３３条の２の書面添付に係るチェックシート.pdf"
            font_name = "HeiseiKakuGo-W5"
            pdfmetrics.registerFont(UnicodeCIDFont(font_name))

            in_path = 'checksheet.pdf'
            # out_path = "完成品/kanseitest.pdf"
            # 保存先PDFデータを作成
            cc = canvas2.Canvas(out_path)
            # PDFを読み込む
            pdf = PdfReader(in_path, decompress=False)
            # PDFのページデータを取得
            # page = pdf.pages[0]
            # # PDFデータへのページデータの展開
            # pp = pagexobj(page) #ページデータをXobjへの変換
            # rl_obj = makerl(cc, pp) # ReportLabオブジェクトへの変換  
            # cc.doForm(rl_obj) # 展開
            # # 長方形の描画
            # cc.setFillColor("yellow", 0.3)
            # # cc.setStrokeColorRGB(1.0, 0, 0)
            # cc.rect(32, 50, 57, 10, fill=1, stroke = 0)
            # cc.circle(100 * mm, 120 * mm, 40 * mm, 1, 1)
            # # ページデータの確定
            # cc.showPage()
            # cc.circle(100 * mm, 120 * mm, 40 * mm, 1, 1)
            # cc.showPage()
            # cc.circle(100 * mm, 120 * mm, 40 * mm, 1, 1)
            # cc.showPage()
            # cc.circle(100 * mm, 120 * mm, 40 * mm, 1, 1)
            # cc.showPage()
            # cc.circle(100 * mm, 120 * mm, 40 * mm, 1, 1)
            # cc.showPage()
            # cc.circle(100 * mm, 120 * mm, 40 * mm, 1, 1)
            # cc.showPage()
            # cc.circle(100 * mm, 120 * mm, 40 * mm, 1, 1)
            
            # 図形設定時のpdfの縦横ピクセル？
            xtx = 1652
            yty = 2338-12
            # ピクセル？からｃｍの変換
            xxx = 78.7
            cksize = 8
            # 確認チェックボックスのY座標
            kakuni = [[535, 677,734, 790,847, 903,960,1017,1073,1130,1187, 1272,1385,1442,1527, 1583,1806,1895,1980,2036],
                           [220,333, 361,418, 475, 531, 597,654,711,767,875,932,988,1102,1158,1243, 1328, 1441, 1497,1639],
                           [220,276,333,418,531,616, 673,758,814,871,928,984,1126, 1211,1296,1353,1466,1636,1720,1894,2012],
                           [220,333,417, 531,649,705,814,984,1069,1125,1239,1324,1437,1494,1664,1720,1888,1947,2004,2117],
                           [220,305, 417, 562,663,861,1026,1422,1618,1731,1927],
                           [220,446,513,569, 644,728,781, 838, 911,1005,1034,1090,1175,1260,1373,1458, 1571,1684,1712,1768],
                           [219,407,596, 731,875,960,988,1130, 1243,1357,1442,1611,1781,1941],
                           [220,332,472,612,752]]
            # 添付だけの場合1 両方の場合0 有無だけの場合-1
            umutenpu = [[1,1,1,0,0 ,0, 0,0,0,0,0,0,0 ,0,0,0,0 ,1,0,0,0],
               [0,0,0 ,1,0 ,0 ,0 ,0,0,0,0,0,0,0,0,0,0 ,0 ],
               [0,0,0,1,1,1,1,1,1,1,1,0,1, 0,1,0,0,0,0,0,0],
               [0,1,0, 0,1,1,1,0,1,1,1,1,1,0,1,1,0,1,0,1,1,0,0],
               [0,0, 0, 0,0,1,1,0,0,1,1,0,0,0,1,1],
               [1,1,1,1, 0,1,1,1, 1, 0,0,0,0,0,1,0, 0,0,0,1],
               [0,-1,0, 0,0,0,0, 1,1,1,0,0,0,1],
               [0,0,0,0,1]]
            def check(x = 0, y = 0):
                cc.line((x/xxx)*cm,((yty-y+cksize)/xxx)*cm, ((x+cksize)/xxx)*cm, ((yty-y)/xxx)*cm)
                cc.line(((x+cksize*3)/xxx)*cm,((yty-y+cksize*2)/xxx)*cm, ((x+cksize)/xxx)*cm, ((yty-y)/xxx)*cm)
            
            # setent = ["","","","","","","","","","","","","","","","","",""]
            for i in range(pdfpages):
                page = pdf.pages[i]
                pp = pagexobj(page) #ページデータをXobjへの変換
                rl_obj = makerl(cc, pp) # ReportLabオブジェクトへの変換  
                cc.doForm(rl_obj) # 展開
                for j in range(len(kakuni[i])):
                    check(1327,kakuni[i][j])
                for k in range(len(umutenpu[i])):
                    p = checkbtn[i][k].get()
                    if umutenpu[i][k] == 1:
                        if p == True:
                            check(1471,switch0Base[i][k])
                    elif umutenpu[i][k] == 0:
                        if p == True:
                            check(1369,switch0Base[i][k])
                            check(1471,switch0Base[i][k])
                        else:
                            check(1410,switch0Base[i][k])
                    else:
                        if p == True:
                            check(1369,switch0Base[i][k])
                        else:
                            check(1410,switch0Base[i][k])
                # ページデータの確定
                if i == 7:
                    print("ss")
                    cc.setFont(font_name, 20)
                    cc.drawString(4.5 * cm, 6.8 * cm, str(Ent_hisou.get()))
                    cc.setFont(font_name, 10)
                    cc.drawString(4.5 * cm, 5.5 * cm, str(Ent_souju1.get()))
                    cc.drawString(4.5 * cm, 5.0 * cm, str(Ent_souju2.get()))
                    cc.setFont(font_name, 12)
                    cc.drawString(4.5 * cm, 4.25 * cm, str(Ent_sousi.get()))
                    cc.setFont(font_name, 10)
                    cc.drawString(4.5 * cm, 3.5 * cm, str(Ent_soute1.get()))
                    cc.drawString(6.5 * cm, 3.5 * cm, str(Ent_soute2.get()))
                    cc.drawString(8 * cm, 3.5 * cm, str(Ent_soute3.get()))
                    cc.drawString(12 * cm, 5.5 * cm, str(Ent_zeiju1.get()))
                    cc.drawString(12 * cm, 5.0 * cm, str(Ent_zeiju2.get()))
                    cc.drawString(12 * cm, 4.1 * cm, str(Ent_zeisi.get()))
                    cc.setFont(font_name, 8)
                    cc.drawString(15 * cm, 3.5 * cm, str(Ent_zeite1.get()))
                    cc.drawString(16 * cm, 3.5 * cm, str(Ent_zeite2.get()))
                    cc.drawString(17 * cm, 3.5 * cm, str(Ent_zeite3.get()))
                cc.showPage()

            # PDFの保存
            cc.save()
            GOBTN["state"]= "disable"
            # file_name = 'OFFICE54.pdf'    # ファイル名を設定
            # pdf = canvas.Canvas(file_name, pagesize=portrait(A4))    # PDFを生成、サイズはA4
            # pdf.saveState()    # セーブ

            # pdf.setAuthor('OFFICE54')
            # pdf.setTitle('TEST')
            # pdf.setSubject('TEST')

            # ### 図形の描画 ###
            # pdf.setFillColorRGB(54, 54, 0)
            # pdf.rect(1*cm, 1*cm, 4*cm, 4*cm, stroke=1, fill=1)
            # pdf.setFillColorRGB(0, 0, 0)

            # ### 線の描画 ###
            # pdf.setLineWidth(1)
            # pdf.line(10.5*cm, 27*cm, 10.5*cm, 1*cm)

            # ### フォント、サイズを設定 ###
            # pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
            # pdf.setFont('HeiseiKakuGo-W5', 12)

            # ### 文字を描画 ###
            # pdf.drawString(1*cm, 26*cm, 'テストテストテストtesttesttest')

            # pdf.setFont('HeiseiKakuGo-W5', 20)    # フォントサイズの変更
            # width, height = A4  # A4用紙のサイズ
            # pdf.drawCentredString(width / 2, height - 2*cm, 'タイトル')

            # pdf.restoreState()
            # pdf.save()
        pass
    def zeirisi():
        zeirisiInfo = []
        with open('zeirisi.csv', 'r',encoding="utf-8") as f:
            reader = csv.reader(f)
            header = next(reader)
            print(reader)
            for row1 in reader:
                print(row1)
                zeirisiInfo.append(row1)
            print(zeirisiInfo)
            Ent_zeiju1.delete(0,tk.END)
            Ent_zeiju2.delete(0,tk.END)
            Ent_zeisi.delete(0,tk.END)
            Ent_zeite1.delete(0,tk.END)
            Ent_zeite2.delete(0,tk.END)
            Ent_zeite3.delete(0,tk.END)

            Ent_zeiju1.insert(0,zeirisiInfo[0][0])
            Ent_zeiju2.insert(0,zeirisiInfo[0][1])
            Ent_zeisi.insert(0,zeirisiInfo[0][2])
            Ent_zeite1.insert(0,zeirisiInfo[0][3])
            Ent_zeite2.insert(0,zeirisiInfo[0][4])
            Ent_zeite3.insert(0,zeirisiInfo[0][5])
            

    zeirisibtn = tk.Button(sub_win, text = "関与税理士",command = zeirisi)
    zeirisibtn.place(x = 720, y = 850)
    GOBTN = tk.Button(sub_win, text = "完了",font = ("",20), command = GOGOGO)
    GOBTN.place(x = 720, y = 900)
    sub_win.mainloop()

TL_checksheet()
