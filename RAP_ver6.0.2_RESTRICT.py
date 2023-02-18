### 6.0.1 버전과 다른 점 
### 부문별 분석지표 추가, 지표별 데이터 다운로드 등

from tkinter import *
import tkinter as tk
import tkinter.messagebox as msgbox
import tkinter.ttk as ttk   # 콤보박스용 ttk 불러오기
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import pandas as pd
import matplotlib
from pandastable import Table, TableModel
import random
import sys, os
import numpy as np
import time
from tkinter import filedialog
import datetime, shutil
from openpyxl import load_workbook
from openpyxl import Workbook
from pandas import ExcelWriter
from openpyxl.styles import Font
# from PIL import Image, ImageTk
# from scipy.ndimage.filters import gaussian_filter1d


sido = ['전국','서울특별시','부산광역시','대구광역시','인천광역시','광주광역시','대전광역시', '울산광역시', \
    '세종특별자치시','경기도','강원도','충청북도','충청남도','전라북도','전라남도','경상북도','경상남도','제주특별자치도']

maincolor = '#89BBF1'  # 기본배경, 회색: #F0F0F0, 파랑: #64A5EC, 연파랑: #89BBF1
backcolor = 'aliceblue' # 차트배경
chartcolor = 'darkblue' # 'darkblue'
now = datetime.datetime.now()
month_t = now.month
day_t = now.day
datenow = str(month_t) + "." +str(day_t)

# 그래프 기본 설정
matplotlib.rcParams['font.family'] = 'Malgun Gothic' # Windows 용 '맑은고딕' 폰트 설정
matplotlib.rcParams['font.size'] = 9 # 글자크기 설정
matplotlib.rcParams['axes.unicode_minus'] = False  # 한글 폰트 사용 시, 마이너스 글자 깨짐 방지
matplotlib.rcParams.update({'text.color' : 'black', 'axes.labelcolor' : 'black'})
matplotlib.rcParams['axes.edgecolor'] = 'black'  ## 차트외곽선 색
matplotlib.rcParams['axes.facecolor'] = 'aliceblue'  ## 차트 안 영역 색
matplotlib.rcParams['figure.facecolor'] = 'aliceblue'  # 차트 밖 영역 색
matplotlib.rcParams['xtick.color'] = 'black'
matplotlib.rcParams['ytick.color'] = 'black'

gungucbox_values = {
    '전국'           : ['전체'],
    '서울특별시'     : ['전체', '강남구', '강동구', '강북구', '강서구', '관악구', '광진구', '구로구', '금천구', '노원구', '도봉구', '동대문구', '동작구', \
        '마포구', '서대문구', '서초구', '성동구', '성북구', '송파구', '양천구', '영등포구', '용산구', '은평구', '종로구', '중구', '중랑구'],
    '부산광역시'     : ['전체', '강서구', '금정구', '기장군', '남구', '동구', '동래구', '부산진구', '북구', '사상구', '사하구', '서구', '수영구', '연제구', '영도구', '중구', '해운대구'],
    '대구광역시'     : ['전체', '남구', '달서구', '달성군', '동구', '북구', '서구', '수성구', '중구'],
    '인천광역시'     : ['전체', '강화군', '계양구', '남동구', '동구', '미추홀구', '부평구', '서구', '연수구', '옹진군', '중구'],
    '광주광역시'     : ['전체', '광산구', '남구', '동구', '북구', '서구'],
    '대전광역시'     : ['전체','동구','중구','서구','유성구','대덕구'],        
    '울산광역시'     : ['전체', '남구', '동구', '북구', '울주군', '중구'],
    '세종특별자치시' : ['전체'],
    '경기도'         : ['전체', '가평군', '고양시', '과천시', '광명시', '광주시', '구리시', '군포시', '김포시', '남양주시', '동두천시', '부천시', \
        '성남시', '수원시', '시흥시', '안산시', '안성시', '안양시', '양주시', '양평군', '여주시', '연천군', '오산시', '용인시', '의왕시', \
        '의정부시', '이천시', '파주시', '평택시', '포천시', '하남시', '화성시'],
    '강원도'         : ['전체', '강릉시', '고성군', '동해시', '삼척시', '속초시', '양구군', '양양군', '영월군', '원주시', '인제군', '정선군', '철원군', \
        '춘천시', '태백시', '평창군', '홍천군', '화천군', '횡성군'],
    '충청북도' : ['전체','청주시','충주시','제천시','보은군','옥천군','영동군','증평군','진천군','괴산군','음성군','단양군'],
    '충청남도' : ['전체','천안시','공주시','보령시','아산시','서산시','논산시','계룡시',\
        '당진시','금산군','부여군','서천군','청양군','홍성군','예산군','태안군'],
    '전라북도'  : ['전체', '고창군', '군산시', '김제시', '남원시', '무주군', '부안군', '순창군', '완주군', '익산시', '임실군', '장수군', '전주시', '정읍시', '진안군'],
    '전라남도'  : ['전체', '강진군', '고흥군', '곡성군', '광양시', '구례군', '나주시', '담양군', '목포시', '무안군', '보성군', '순천시', '신안군', \
        '여수시', '영광군', '영암군', '완도군', '장성군', '장흥군', '진도군', '함평군', '해남군', '화순군'],
    '경상북도'  : ['전체', '경산시', '경주시', '고령군', '구미시', '군위군', '김천시', '문경시', '봉화군', '상주시', '성주군', '안동시', '영덕군', \
        '영양군', '영주시', '영천시', '예천군', '울릉군', '울진군', '의성군', '청도군', '청송군', '칠곡군', '포항시'],
    '경상남도'  : ['전체', '거제시', '거창군', '고성군', '김해시', '남해군', '밀양시', '사천시', '산청군', '양산시', '의령군', '진주시', '창녕군', \
        '창원시', '통영시', '하동군', '함안군', '함양군', '합천군'],
    '제주특별자치도' : ['전체', '서귀포시', '제주시']
}

mainw = Tk()
mainw.title("★ 함께봐You! 우리 지역 통계  ★ 같이해You! 더 좋은 일자리 만들기") ### 원클릭✔, 우리지역 통계 

# menu = Menu(mainw)
mainw.resizable(True, True) # 창 너비높이 변경 가능 여부
mainw.geometry("340x800+150+100")    


### 파일 절대경로 지정 함수
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstoaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASSp
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def activate_info():
    msgbox.showinfo("안내", "☞ 이 자료는 지자체 및 일자리 유관기관 업무담당자의 통계분석을 돕기 위해," + "\n" \
        + "고용노동부 청주지청 지역협력과에서 제작(담당: 정연욱, ☎043-299-1333)하였습니다." + "\n" \
        + "'모든 이용자의 임의 편집, 복제 및 재배포를 적극 권장'하오니 자유로이 활용하시기 바랍니다.")

# 플랫폼 소개 메뉴
# menu_file = Menu(menu, tearoff = 0)
# menu_file.add_command(label = "분석폼 소개", command = activate_info) 
# menu_file.add_command(label = "끝내기", command = mainw.quit)
# menu_file.add_command(label = "끝내기", command = exit, state = "disable")  # 메뉴를 비활성화시킬 때
# menu_file.add_separator()   # 구분선 넣기

def on_select(event):
    selected = event.widget.get()
    global gungucbox
    if (sidocbox.get() == "전국") | (sidocbox.get() == "세종특별자치시"):
        gungucbox.set("전체")
    else:
        gungucbox.set("선택")
    gungucbox['values'] = gungucbox_values[selected]
    global values
    values = gungucbox_values[selected]     

def Ton_select(event):
    selected = event.widget.get()
    global Tgungucbox
    if (Tsidocbox.get() == "전국") | (Tsidocbox.get() == "세종특별자치시"):
        Tgungucbox.set("전체")
    else:
        Tgungucbox.set("선택")
    Tgungucbox['values'] = gungucbox_values[selected]
    global values
    values = gungucbox_values[selected] 

def search():
    # print("광역시도 : ", sidocbox.get())
    # print("기초시군구 : ", gungucbox.get())
    global sido, sigu, tido, tigu, oreg, treg, n_sido, n_tido
    sido = sidocbox.get()
    sigu = gungucbox.get()
    tido = Tsidocbox.get()
    tigu = Tgungucbox.get()
    selcol = random.choice(['lightblue1', 'lightcyan','lightyellow2','thistle1','lightgoldenrodyellow','lightsteelblue1'])

    if sido == "선택" or sigu == "선택" or tido == "선택" or tigu == "선택":
        msgbox.showinfo("확인", "지역 선택을 완료하세요!")
        return
    elif sido == tido and sigu == tigu:
        msgbox.showinfo("확인", "분석지역과 비교지역은 서로 다른 지역을 선택하세요!")
        return

    if sigu == "전체":
        oreg = sido
        n_sido = sido
    else:
        oreg = sido + " " + sigu
        n_sido = sigu
    if tigu == "전체":
        treg = tido
        n_tido = tido
    else:
        treg = tido + " " + tigu
        n_tido = tigu

    main_label.configure(text = "[ '" + oreg + "' 고용노동 지표 분석(비교: " + treg + ") ]", bg = selcol, \
        font = ("arial", 17, "bold"), pady = 30) # 분석지역 표시 라벨
   
    ### tkinter 테마 사용
    global selcol_tree, selcol_headbar
    selcol_tree = random.choice(['dimgray','darkmagenta', 'darkred','darkolivegreen','darkcyan','dodgerblue','darkslateblue'])
    selcol_headbar = random.choice(['whitesmoke','floralwhite', 'mintcream','aliceblue','lavender','lightcyan','seashell'])
    
    style = ttk.Style()
    style.theme_use('clam') # 'clam', 'alt', 'default', 'classic'
    style.configure('Treeview.Heading', background = selcol_headbar) ## 제목줄 색깔
    style.map('Treeview', background=[('selected', selcol_tree)]) ## 선택한셀 색깔 ['dimgray','darkmagneta', 'darkred','darkolivegreen','darkcyan','royalblue','darkslateblue'])
    style.configure("red.Horizontal.TProgressbar", foreground='red', background='red')

    btnphoto.configure(image=photo, bg= selcol)  # button 컬러만 변경

    #### progressbar 적용, 팝업 등
    popup = tk.Toplevel()
    tk.Label(popup, font = ("arial", 11, "normal"), text = "------- 자료 조회 중 입니다 -------").grid(row = 1, column = 1, padx = 10, pady = 10)

    p_var = DoubleVar()
    p_var = ttk.Progressbar(popup, style = 'red.Horizontal.TProgressbar', orient = "horizontal", variable=p_var, maximum=100, length = 500)
    p_var.grid(row = 2, column = 1, padx = 10, pady = 5)
    tk.Label(popup, font = ("arial", 9, "normal"), anchor = 'e', text = " ").grid(row = 3, column = 1, padx = 10, pady = 0, sticky = 'e')
    popup.pack_slaves()

    debug()
    time.sleep(0.05)
    maincolor = '#89BBF1'  # 기본배경, 회색: #F0F0F0, 파랑: #64A5EC, 연파랑: #89BBF1
    mainw.geometry("1920x1020+0+0")
    # mainw.state('zoomed')

    for iframe in [ad1_frame, first_frame]:
        iframe.configure(bg= 'white', fg = 'midnightblue', relief = "raised")
    Pop_frame.configure(bg = 'white')
    popchart_plot()
    popchart_pyramid()
    popchart_pie()
    p_var['value'] = 25
    popup.update()
    time.sleep(0.05)

    for iframe in [ad2_frame, second_frame]:
        iframe.configure(bg= 'white', fg = 'midnightblue', relief = "raised")
    emp_frame.configure(bg = 'white')
    emp_plot()
    emp_sex_plot()
    emp_youngage_plot()
    spyder_plot_a()
    spyder_plot_b()
    p_var['value'] = 50
    popup.update()
    time.sleep(0.05)

    for iframe in [ad3_frame, third_frame]:
        iframe.configure(bg= 'white', fg = 'midnightblue', relief = "raised")
    sanup_frame.configure(bg = 'white')
    sanup_bar()
    worker_bar()
    p_var['value'] = 75
    popup.update()
    time.sleep(0.05)

    for iframe in [ad4_frame, forth_frame]:
        iframe.configure(bg= 'white', fg = 'midnightblue', relief = "raised")
    guin_frame.configure(bg = 'white')
    guin_plot()
    guinjob_plot()
    guinmatch_plot()
    p_var['value'] = 100
    popup.update()
    tk.Label(popup, font = ("arial", 11, "bold"), text = "♣♣♣ 자료 업데이트 완료. 이 창은 곧 닫힙니다. ♣♣♣").grid(row = 1, column = 1, padx = 10, pady = 10)
    tk.Label(popup, font = ("arial", 9, "normal"), anchor = 'e', text = "고용노동부").grid(row = 3, column = 1, padx = 10, pady = 0, sticky = 'e')
    popup.update()
    time.sleep(1.5)
    popup.destroy()
    # mainw.state("normal")
    # mainw.geometry("1920x1020+0+0") 
    btnphoto.pack(side = LEFT, padx = 20, pady = 8)
    adjust_scrollregion()

def debug():

    try:
        # name_frames = [Pop_frame, emp_frame, sanup_frame]
        for widgets in Pop_frame.winfo_children():
            widgets.destroy()
        for widgets in emp_frame.winfo_children():
            widgets.destroy()
        for widgets in sanup_frame.winfo_children():
            widgets.destroy()     
        for widgets in guin_frame.winfo_children():
            widgets.destroy()           
        # canvas.get_tk_widget().destroy()
        # popplot_label.destroy()
        # popplot_label2.destroy()
    except:
        pass

def on_mousewheel(event):
    shift = (event.state & 0x1) != 0
    scroll = -1 if event.delta > 0 else 1
    if shift:
        my_canvas.xview_scroll(scroll, "units")
    else:
        my_canvas.yview_scroll(scroll, "units")    

def adjust_scrollregion():
    my_canvas.configure(scrollregion = my_canvas.bbox("all"))


########################### CHART chapter Ⅰ. 인구 꺾은선 차트 ###########################
def popchart_plot():
    ## csv 파일 열기
    df = pd.read_csv(resource_path('data/1_1_db_1st_popul_2022.csv'), encoding = 'cp949')
    
    ### 인구 라벨_꺾은선차트 아래 ###
    dfa = df.loc[((df['시도'] == sido) & (df['행정구역'] == sigu) & (df['연월'] != "'16.12월")), ['연월', 'value']]
    dft = df.loc[((df['시도'] == tido) & (df['행정구역'] == tigu) & (df['연월'] != "'16.12월")), ['연월', 'value']]
    
    # 인구 라벨 첫번째 문구 작성
    dfa_value = dfa['value'].iloc[-1]
    dfa_value2 = dfa['value'].iloc[-2]
    dfa_plus = ((dfa_value - dfa_value2) / dfa_value2) *100
    # dft_value = dft['value'].iloc[-1]
    if dfa_value > dfa_value2:
        dfa_title = "+" + dfa_plus.round(2).astype(str)
    else:
        dfa_title = dfa_plus.round(2).astype(str)
    dfa_str = format(dfa_value, ",")
    # dft_str = format(dft_value, ",")
    global pop_label1, pop_label2, pop_label3
    pop_label1 = Label(Pop_frame, bg = 'white', \
        relief = "flat", text = "총 인구 : " + dfa_str + "명(전년대비 " + dfa_title + "%)", \
        font = ("arial", 17, "bold"), padx = 20, pady = 0, anchor = "w", fg = 'black')
    pop_label1.grid(row = 1, column = 0, ipadx = 0, ipady = 0, padx = 0, pady = 5, sticky = "nsew")


    ### 조건에 맞는 자료 추출 및 머지
    pd.options.display.float_format = '{:,.2f}'.format  ## 소수첫째자리 옵션설정
    x = dft
    y = dfa
    a = pd.merge(x,y, on = '연월')
    a = a.rename(columns= {'연월': '연월', 'value_x': n_tido, 'value_y': n_sido})
    a = a.reset_index(drop=True)
    b = a.copy()

    b[n_tido] = round(b[n_tido] /10000, 2)
    b[n_sido] = round(b[n_sido] /10000, 2)

    ## 인구 꺾은선 차트 그리기 ##
    x = b['연월']
    y1 = b[n_tido]
    y2 = b[n_sido]
    max_y1 = y1.max() + 0.3  ## y축 최대값
    min_y1 = y1.min() - 0.3  ## y축 최소값
    max_y2 = y2.max() + 0.3  ## y축 최대값
    min_y2 = y2.min() - 0.3  ## y축 최소값

    fig_p1, ax1 = plt.subplots(figsize = (4.53, 3))
    fig_p1.suptitle("[ '" + oreg + "' 인구 추이 ]", fontsize = 11, y=0.95)

    ax1.set_ylabel(oreg + '(만명)')
    ax1.plot(x, y2, color = 'red', linewidth = 2, marker = "o", markeredgewidth = 1, markeredgecolor = "red", markerfacecolor = "yellow", markersize = 6, label = n_sido, markevery = 1)
    ax1.spines['top'].set_visible(False)  # 테두리 여부
    ax1.set_ylim(min_y2, max_y2)

    ax2 = ax1.twinx() # x축을 함께 사용    
    ax2.set_ylabel(treg + ' (만명)')
    ax2.plot(x, y1, color = 'darkblue', linewidth = 1, marker = "o", markeredgewidth = 1, markerfacecolor = backcolor, \
        markersize = 3, label = n_tido, markevery = 1, linestyle = '--')
    ax2.spines['top'].set_visible(False)  # 테두리 여부
    ax2.set_ylim(min_y1, max_y1)
  

    # y2 차트 데이터레이블 표시
    for idx, txt in enumerate(y2):
        ax1.text(x[idx], y2[idx] + 0.03, txt, ha='center')

    
    fig_p1.legend(loc = (0.16, 0.75), fontsize = 9, facecolor = backcolor) # 범례 표시
    fig_p1.tight_layout(pad=1, h_pad=None, w_pad= 0.9)  ## 이미지 타이트하게

    global canvas
    # canvas = Canvas(Pop_frame, width = 1700, height = 300, bg = 'cornsilk', bd = 1, relief = "raised")
    canvas = FigureCanvasTkAgg(fig_p1, master = Pop_frame)
    canvas.draw()
    canvas.get_tk_widget().grid(row = 2, column = 0, ipadx = 20, ipady = 20, padx = 20, pady = 2, sticky=N+E+W+S)
    # canvas.get_tk_widget().pack() 

    ### 차트 편집_a1_plot
    downbt_a1_plot = Button(Pop_frame, text = "편집/저장", bg = "mistyrose", width = 6, relief = "flat", overrelief = "raised", command = a1_plot)
    downbt_a1_plot.grid(row = 2, column = 0, ipadx = 5, ipady = 0, padx = 20, pady = 2, sticky = 'nw')


    ###★★★★★★ Treeview, 인구 꺾은선 Treeview
    a[n_sido + "증감(명)"] = 0
    a[n_sido + "증감률(%)"] = 0.00
    a[n_tido + "대비비중(%)"] = 0.00
    a[n_tido + '증감(명)'] = 0
    a[n_tido + '증감률(%)'] = 0.00
    for i in a.index:
        a.iloc[i,3] = a.iloc[i, 2] - a.iloc[i-1, 2]
        a.iloc[i,4] = round(a.iloc[i, 3] / a.iloc[i-1, 2] *100, 2)
        a.iloc[i,5] = round(a.iloc[i, 2] / a.iloc[i, 1] *100, 2)
        a.iloc[i,6] = a.iloc[i, 1] - a.iloc[i-1, 1]
        a.iloc[i,7] = round(a.iloc[i, 6] / a.iloc[i-1, 1] *100, 2)

    a[n_tido] = a.apply(lambda x: "{:,}".format(x[n_tido]), axis=1) ### 천단위 구분기호
    a[n_sido] = a.apply(lambda x: "{:,}".format(x[n_sido]), axis=1) ### 천단위 구분기호
    a[n_sido + "증감(명)"] = a.apply(lambda x: "{:,}".format(x[n_sido + "증감(명)"]), axis=1) ### 천단위 구분기호
    a[n_tido + '증감(명)'] = a.apply(lambda x: "{:,}".format(x[n_tido + '증감(명)']), axis=1) ### 천단위 구분기호
    a.iloc[0, 3] = "-"
    a.iloc[0, 4] = "-"
    a.iloc[0, 6] = "-"
    a.iloc[0, 7] = "-"    
    a = a[['연월', n_sido, n_sido + "증감(명)", n_sido + "증감률(%)", n_tido + "대비비중(%)", n_tido, n_tido + '증감(명)', n_tido + '증감률(%)']] ## 열 정렬
    a = a.rename(columns= {n_sido: "○ " + n_sido + "(명)", n_sido + "증감(명)": '    ㄴ증감(명)', n_sido + "증감률(%)": '    ㄴ증감률(%)', \
        n_tido: '○ ' + n_tido + '(명)', n_tido + '대비비중(%)': "※" + n_sido + "/" + n_tido + "(%)", n_tido + '증감(명)':"    ㄴ증감(명)", n_tido + '증감률(%)' : '    ㄴ증감률(%)'})
    a = a.set_index('연월').T
    a = a.reset_index(drop=False)

    a = a.rename(columns = ({'index' : "구분"}))

    cols = list(a.columns)
    column_append_cols = ["c"+ str(i+1) for i in range(len(cols))]  ###★★★★★★ 리스트 누적해서 열리스트 한정(대전광역시 자치구 때문에 필요)
    
    treeview_people = ttk.Treeview(Pop_frame, column= column_append_cols, show = 'headings', height = 5)
    ### Treeview heading
    for i, j in enumerate(cols):
        treeview_people.column("# " + str(i+1), anchor=CENTER, width = 20)
        treeview_people.heading("# " + str(i+1), text = j)  
    treeview_people.column("c1", width = 60)  # 지역구분은 좀 넓게 추가 정의

    ### Treeview data insert
    for i in a.index:
        treeview_people.insert('', 'end', text = i, values = (list(a.values[i])))

    treeview_people.column(0, anchor= 'w') ### Treeview 내부 '구분'항목은 왼쪽 정렬
    for i in [1,6]:
        treeview_people.column(i, anchor= 'e') ### Treeview 내부 '숫자'항목은 오른쪽 정렬

    treeview_people.grid(row = 3, rowspan = 2, column = 0, ipadx = 20, ipady = 20, padx = 20, pady = 0, sticky=N+E+W+S)

    ### 자료다운_a1_dt
    global a1_dt
    a1_dt = a.copy()
    downbt_a1_dt = Button(Pop_frame, text = "xls", bg = "mistyrose", width = 2, height = 1, relief = "flat", overrelief = "raised", command = a1_dtdown)
    downbt_a1_dt.grid(row = 4, column = 0, ipadx = 0, ipady = 0, padx = 21, pady = 1, sticky = 'sw')
   

########################### 인구 두번째, 피라미드 차트 그리기 ###########################
def popchart_pyramid():
    #### 피라미드 해설 라벨
    global pop_label4, pop_label5, pop_label6
    pop_label4 = Label(Pop_frame, bg = 'white', \
        relief = "flat", text = " < " + oreg + " 인구피라미드 변화 >", \
                font = ("arial", 15, "bold"), padx = 30, pady = 7, anchor = "n", fg = 'black')
    pop_label4.grid(row = 1, column = 1, columnspan = 6, ipadx = 0, ipady = 3, padx = 0, pady = 0, sticky = "nsew")

    ### ★★★★★★★★★ 과거 첫번째 인구피라미드, csv 파일 열기 ★★★★★★★★★★★★★★★
    df_p = pd.read_csv(resource_path('data/1_1_db_2st_past_popul_2022.csv'))
    
    ### 조건에 맞는 자료 추출 및 머지
    x_p = df_p.loc[(df_p['시도'] == sido) & (df_p['행정구역'] == sigu) & (df_p['성별'] == "남자인구수[명]"), ['5세별', 'value']]
    y_p = df_p.loc[(df_p['시도'] == sido) & (df_p['행정구역'] == sigu) & (df_p['성별'] == "여자인구수[명]"), ['5세별', 'value']]
    past_dataframe = df_p.loc[(df_p['시도'] == sido) & (df_p['행정구역'] == sigu) & (df_p['성별'] == "남자인구수[명]"), ['5세별', '연월', 'value']]
    # past_dataframe['연월'] = past_dataframe['연월'].str.slice(start = 0, stop = 3) + "년말"
    global past
    past = past_dataframe.iloc[0, 1]

    a_p = pd.merge(x_p, y_p, on = '5세별')
    # a = a.reset_index(drop=True)

    ## 바차트 그리기 ##
    fig = Figure(figsize = (3.53, 3), dpi = 100)
    ax = fig.add_subplot(111)
    fig.suptitle("<<  " + past + "  >>", fontsize = 11, x=0.56, y=0.95)
    ax.barh(a_p['5세별'], -a_p.value_x, label = "남(명)", color = 'forestgreen')
    ax.barh(a_p['5세별'], a_p.value_y, label = "여(명)", color = 'darkorange')
    ax.spines['top'].set_visible(False)  # 테두리 여부
    ax.spines['right'].set_visible(False)  # 테두리 여부
    ax.tick_params(right = False, left = False, bottom = False, labelleft=True, labelright=False, pad = 0) ### 눈금선 없애기, 라벨위치여부
    fig.tight_layout()  ## 이미지 타이트하게
    fig.legend(loc = (0.22, 0.71), fontsize = 9) # 범례 표시
    
    global canvas
    # canvas = Canvas(Pop_frame, width = 1700, height = 300, bg = "white")
    canvas = FigureCanvasTkAgg(fig, master = Pop_frame)
    canvas.draw()
    canvas.get_tk_widget().grid(row = 2, column = 1, columnspan= 4, ipadx = 20, ipady = 20, padx = 10, pady = 2)
    # canvas.get_tk_widget().pack()

    ### 차트 편집_a2_plot
    downbt_a2_plot = Button(Pop_frame, text = "편집/저장", bg = "mistyrose", width = 6, relief = "flat", overrelief = "raised", command = a2_plot)
    downbt_a2_plot.grid(row = 2, column = 1, ipadx = 5, ipady = 0, padx = 10, pady = 2, sticky = 'nw')


    ########## 인구 두번째, 현재 인구피라미드, csv 파일 열기
    df = pd.read_csv(resource_path('data/1_1_db_2st_now_popul_2022.csv'))
    global present
    present = df.iloc[0,5]

    ### 조건에 맞는 자료 추출 및 머지
    x = df.loc[(df['시도'] == sido) & (df['행정구역'] == sigu) & (df['성별'] == "남자인구수[명]"), ['5세별', 'value']]
    y = df.loc[(df['시도'] == sido) & (df['행정구역'] == sigu) & (df['성별'] == "여자인구수[명]"), ['5세별', 'value']]

    a = pd.merge(x,y, on = '5세별')
    # a = a.reset_index(drop=True)
    
    ## 바차트 그리기 ##
    fig = Figure(figsize = (3.53, 3), dpi = 100)
    ax = fig.add_subplot(111)
    fig.suptitle("<<  " + present + "  >>", fontsize = 11, x=0.425, y=0.95)
    ax.barh(a['5세별'], -a.value_x, label = "남(명)", color = 'forestgreen')
    ax.barh(a['5세별'], a.value_y, label = "여(명)", color = 'darkorange')
    ax.spines['top'].set_visible(False)  # 테두리 여부
    ax.spines['left'].set_visible(False)  # 테두리 여부
    ax.tick_params(right = False, left = False, bottom = False, labelleft=False, labelright=True, pad = 0) ### 눈금선 없애기, 라벨위치여부
    fig.tight_layout()  ## 이미지 타이트하게
    fig.legend(loc = (0.58, 0.71), fontsize = 9) # 범례 표시
  
    # canvas = Canvas(Pop_frame, width = 1700, height = 300, bg = "white")
    canvas = FigureCanvasTkAgg(fig, master = Pop_frame)
    canvas.draw()
    canvas.get_tk_widget().grid(row = 2, column = 6, ipadx = 20, ipady = 20, padx = 10, pady = 2)
    # canvas.get_tk_widget().pack()

    ### 차트 편집_a3_plot
    downbt_a3_plot = Button(Pop_frame, text = "편집/저장", bg = "mistyrose", width = 6, relief = "flat", overrelief = "raised", command = a3_plot)
    downbt_a3_plot.grid(row = 2, column = 6, ipadx = 6, ipady = 0, padx = 10, pady = 2, sticky = 'ne')

    
    ## 중간 영역 '변화'라벨과 화살표
    pop_label5 = Label(Pop_frame, fg = 'black', bg = 'white', \
        relief = "flat", text = "변화" + "\n" + "→" + "\n" + "→" + "\n" + "→" + "\n", font = ("arial", 11, "bold"), padx = 0, anchor = "c")
    pop_label5.grid(row = 2, column = 5, ipadx = 0, ipady = 0, padx = 0, pady = 0, sticky = "nesw")


    ###★★★★★★ Treeview용 인구피라미드 자료 편집
    a_p = a_p.rename(columns= {'value_x': 'p_m', 'value_y': 'p_w'}) # 과거자료 변수변경 및 인구합계(아래구문)
    a_p['p_tot'] = a_p['p_m'] + a_p['p_w']
    a = a.rename(columns= {'value_x': 'n_m', 'value_y': 'n_w'}) # 현재자료 변수변경 및 인구합계(아래구문)
    a['n_tot'] = a['n_m'] + a['n_w']
    at = pd.merge(a_p, a, on = '5세별')
    at = at.set_index('5세별').T  # 편집용 Transpose

    # 항목 재정의: 5세별 
    under15 = ['0 - 4세', '5 - 9세', '10 - 14세']
    btw1529 = ['15 - 19세', '20 - 24세', '25 - 29세']
    thforty = ['30 - 34세', '35 - 39세', '40 - 44세', '45 - 49세']
    fisixty = ['50 - 54세', '55 - 59세', '60 - 64세', '65 - 69세']
    upper70 = ['70 - 74세', '75 - 79세', '80 - 84세', '85 - 89세', '90 - 94세', '95 - 99세', '100+']
    total = [" ㄴ15세미만", " ㄴ청년층", " ㄴ중년층", " ㄴ장년층", " ㄴ70세이상"]
    list_colname = total + ["○ " + n_sido]
    list_var = (under15, btw1529, thforty, fisixty, upper70, total)
    for k in range(6):
        at[list_colname[k]] = 0
        at[list_colname[k]] = at[list_var[k]].sum(axis = 1)

    at = at[list_colname] # 정의된 변수에 해당하는 컬럼만 남기고 삭제
    cols = list(at.columns) # 컬럼이름 리스트로 받기
    at = (at[[cols[-1]] + cols[0:-1]]).T # 컬럼 순서변경 및 Transpose
    at = at.reset_index(drop=False)
    at["증감_계"] = 0
    at["증감률_계"] = 0.0
    at["증감_남"] = 0
    at["증감률_남"] = 0.0
    at["증감_여"] = 0
    at["증감률_여"] = 0.0
    at["판단_계"] = ''
    at["판단_남"] = ''
    at["판단_여"] = ''
    at = at[["5세별", "판단_계", "p_tot", "n_tot", "증감_계", "증감률_계", "판단_남", \
            "p_m", "n_m", "증감_남", "증감률_남", "판단_여", "p_w", "n_w", "증감_여", "증감률_여"]]

    at["증감_계"] = at["n_tot"] - at["p_tot"]
    at["증감률_계"] = round((at["증감_계"] / at["p_tot"]) *100 , 1)
    at["증감_남"] = at["n_m"] - at["p_m"]
    at["증감률_남"] = round((at["증감_남"] / at["p_m"]) *100 , 1)
    at["증감_여"] = at["n_w"] - at["p_w"]
    at["증감률_여"] = round((at["증감_여"] / at["p_w"]) *100 , 1)
    if_list = ["판단_계", "판단_남", "판단_여"]
    ck_list = ["증감률_계", "증감률_남", "증감률_여"]

    # 판단영역
    for i in range(3):
        at.loc[at[ck_list[i]] < -1, if_list[i]] = "감소"
        at.loc[at[ck_list[i]] >= -1, if_list[i]] = "보합"
        at.loc[at[ck_list[i]] > 1, if_list[i]] = "증가"

    # 원 값 천단위 구분기호 추가
    for i in ["n_tot", "p_tot", "증감_계", "n_m", "p_m", "증감_남", "n_w", "p_w", "증감_여"]:
        at[i] = at.apply(lambda x: "{:,}".format(x[i]), axis=1)
    # print(at.columns)

    at = at[['5세별', '판단_계', 'p_tot', 'n_tot', '증감률_계', '판단_남', 'p_m', 'n_m', \
       '증감률_남', '판단_여', 'p_w', 'n_w', '증감률_여']]
    at = at.rename(columns = {'5세별' : '구분(연령층)', "판단_계":"전체", "p_tot": past, "n_tot" : present, "증감률_계": "증감률", "판단_남":"남자", \
        "p_m": past, "n_m": present, "증감률_남":"증감률", "판단_여":"여자", "p_w": past, "n_w": present, "증감률_여": "증감률"})

    ################## ★★★★★★ Treeview 그리기, dataframe 'at' 계속 사용 ##################
    cols = list(at.columns)
    column_append_cols = ["c"+ str(i+1) for i in range(len(cols))]  ###★★★★★★ 리스트 누적해서 열리스트 한정(대전광역시 자치구 때문에 필요)
    
    treeview_pyramid = ttk.Treeview(Pop_frame, column= column_append_cols, show = 'headings', height = 6)
    ### Treeview heading
    for i, j in enumerate(cols):
        treeview_pyramid.column("# " + str(i+1), anchor=CENTER, width = 20)
        treeview_pyramid.heading("# " + str(i+1), text = j)  
    treeview_pyramid.column("c1", width = 60)  # 지역구분은 좀 넓게 추가 정의

    ### Treeview data insert
    for i in at.index:
        treeview_pyramid.insert('', 'end', text = i, values = (list(at.values[i])))

    treeview_pyramid.column(0, anchor= 'w') ### Treeview 내부 '구분'항목은 왼쪽 정렬
    for i in [2,3,4,6,7,8,10,11,12]:
        treeview_pyramid.column(i, anchor= 'e') ### Treeview 내부 '숫자'항목은 오른쪽 정렬

    treeview_pyramid.grid(row = 3, column = 1, columnspan = 6, ipadx = 20, ipady = 0, padx = 10, pady = 0, sticky='news')   

    ### 연령기준 및 분석 참고
    pop_label6 = Label(Pop_frame, \
        relief = "flat", text = "           ※ 연령층: 청년(15~29세), 중년(30~40대), 장년(50~60대)                          " +"\n" + \
            "              (이용안내) 가로막대 길이를 비교하여 연령대별 증감을 알 수 있습니다." + "\n" + "             단, 가로축이 달라질 경우(예: 세종시) 해석에 유의하셔야 됩니다.       ", \
        font = ("arial", 9, "normal"), padx = 10, pady = 0, anchor = "w", bg = 'white', fg = 'black')
    pop_label6.grid(row = 4, column = 1, columnspan = 6, ipadx = 0, ipady = 0, padx = 0, pady = 0, sticky = "nw")

    # 빈 텍스트로 여백 확보
    pop_label3 = Label(Pop_frame, relief = "flat", text = " ", font = ("arial", 10, "normal"), padx = 20, pady = 0, anchor = "w", bg = 'white', fg = 'black') ## 공백행 삽입
    pop_label3.grid(row = 5, column = 0, ipadx = 0, ipady = 20, padx = 0, pady = 0, sticky = "w")

    ### 자료다운_a2_dt
    global a2_dt
    a2_dt = at.copy()
    downbt_a2_dt = Button(Pop_frame, text = "xls", bg = "mistyrose", width = 2, height = 1, relief = "flat", overrelief = "raised", command = a2_dtdown)
    downbt_a2_dt.grid(row = 4, column = 1, ipadx = 0, ipady = 0, padx = 10, pady = 2, sticky = 'nw')


########################### 인구 세번째, 고령화율 파이 차트 그리기 ###########################
def popchart_pie():
    pop_label4 = Label(Pop_frame, relief = "flat", text = "< " + oreg + " 고령화율 >", font = ("arial", 15, "bold"), padx = 20, pady = 0, bg = 'white', fg = 'black') ## 공백행 삽입
    pop_label4.grid(row = 6, column = 0, columnspan = 5, ipadx = 20, ipady = 10, padx = 10, pady = 2, sticky = "nsew")
    
    df = pd.read_csv(resource_path('data/1_2_db_oldman_2022.csv'))

    ## 시군 고령화율 구하기
    df1 = df.loc[((df['시도'] == sido) & (df['행정구역'] == sigu) & (df['5세코드'] == 0))]
    df2 = df.loc[((df['시도'] == sido) & (df['행정구역'] == sigu) & (df['5세코드'] > 65))]
    df1 = df1[["연월", "value"]]
    df2 = df2.groupby(["연월"]).sum()
    df2 = df2.reset_index(drop= False)
    df2 = df2[['연월', 'value']]
    dfa = pd.merge(df1, df2, on = "연월")
    dfa["고령화율"] = 0.00
    dfa["고령화율"] = round(dfa['value_y'] / dfa['value_x'] *100, 2)
    dfa = dfa[["연월", "고령화율"]]
    ## 비교지역 고령화율 구하기
    df1 = df.loc[((df['시도'] == tido) & (df['행정구역'] == tigu) & (df['5세코드'] == 0))]
    df2 = df.loc[((df['시도'] == tido) & (df['행정구역'] == tigu) & (df['5세코드'] > 65))]
    df1 = df1[["연월", "value"]]
    df2 = df2.groupby(["연월"]).sum()
    df2 = df2.reset_index(drop= False)
    df2 = df2[['연월', 'value']]
    dfb = pd.merge(df1, df2, on = "연월")
    dfb["고령화율"] = 0.00
    dfb["고령화율"] = round(dfb['value_y'] / dfb['value_x'] *100, 2)
    dfb = dfb[["연월", "고령화율"]]
    ## 시군과 비교지역 고령화율 merge
    at = pd.merge(dfa, dfb, on = "연월")
    at = at.rename(columns = {"고령화율_x" : n_sido, "고령화율_y" : n_tido})
    at["차이(%p)"] = round(at[n_sido] - at[n_tido], 2)
    at = at[["연월", n_sido, "차이(%p)", n_tido]]
    
    if n_sido == "청주시":
        a = at.iloc[2,1]
        b = at.iloc[6,1]
        text_title = "      << " + at.iloc[2,0] + " >>                →→                << " + at.iloc[6,0] + " >>                →→                << " + at.iloc[10,0] + " >>"
    else:
        a = at.iloc[0,1]
        b = at.iloc[5,1]
        text_title = "      << " + at.iloc[0,0] + " >>                →→                << " + at.iloc[5,0] + " >>                →→                << " + at.iloc[10,0] + " >>"
    c = at.iloc[10,1]

    colors = ['cornflowerblue', 'crimson']
    explode = [0.05] * 2
    labels = ['65세미만', '65세이상']
    wedgeprops = {'width' : 0.55, 'edgecolor' : 'w', 'linewidth' : 3}
    fig = Figure(figsize = (6, 3), dpi = 100)    
    fig.suptitle(text_title, fontsize = 11, y =0.94)
    ax = fig.add_subplot(131)
    ax.pie((100 - a, a), labels = labels, autopct = '%.1f%% ', textprops = {'fontsize': 11, 'color' : 'black', 'weight' : 'bold'}, pctdistance = 0.7, \
        labeldistance = 1.1, startangle = 90, counterclock = False, colors = colors, explode = explode, wedgeprops = wedgeprops)

    ax2 = fig.add_subplot(132)
    ax2.pie((100 - b, b), autopct = '%.1f%% ', textprops = {'fontsize': 11, 'color' : 'black', 'weight' : 'bold'}, pctdistance = 0.7, \
        labeldistance = 1.1, startangle = 90, counterclock = False, colors = colors, explode = explode, wedgeprops = wedgeprops)

    ax3 = fig.add_subplot(133)
    ax3.pie((100 - c, c), autopct = '%.1f%% ', textprops = {'fontsize': 12, 'color' : 'black', 'weight' : 'bold'}, pctdistance = 0.7, \
        labeldistance = 1.1, startangle = 90, counterclock = False, colors = colors, explode = explode, wedgeprops = wedgeprops)

    fig.tight_layout()  ## 이미지 타이트하게
    # fig.legend(loc = (0.58, 0.71), fontsize = 9) # 범례 표시

    # canvas = Canvas(Pop_frame, width = 1700, height = 300, bg = "white")
    canvas = FigureCanvasTkAgg(fig, master = Pop_frame)
    canvas.draw()
    canvas.get_tk_widget().grid(row = 7, rowspan = 6, column = 0, columnspan = 5, ipadx = 0, ipady = 20, padx = 20, pady = 2, sticky = "nsew")
    # canvas.get_tk_widget().pack()

    ### 차트 편집_a4_plot
    downbt_a4_plot = Button(Pop_frame, text = "편집/저장", bg = "mistyrose", width = 6, relief = "flat", overrelief = "raised", command = a4_plot)
    downbt_a4_plot.grid(row = 7, column = 0, ipadx = 5, ipady = 0, padx = 20, pady = 2, sticky = 'nw')


    ################## ★★★★★★ 고령화율Treeview 그리기, dataframe 'at' 계속 사용 ##################
    cols = list(at.columns)
    column_append_cols = ["c"+ str(i+1) for i in range(len(cols))]  ###★★★★★★ 리스트 누적해서 열리스트 한정(대전광역시 자치구 때문에 필요)
    
    treeview_pyramid = ttk.Treeview(Pop_frame, column= column_append_cols, show = 'headings', height = 2)
    ### Treeview heading
    for i, j in enumerate(cols):
        treeview_pyramid.column("# " + str(i+1), anchor=CENTER, width = 20)
        treeview_pyramid.heading("# " + str(i+1), text = j)  
    # treeview_pyramid.column("c1", width = 60)  # 지역구분은 좀 넓게 추가 정의

    ### Treeview data insert
    for i in at.index:
        treeview_pyramid.insert('', 'end', text = i, values = (list(at.values[i])))

    treeview_pyramid.grid(row = 8, rowspan = 5, column = 5, columnspan = 2, ipadx = 20, ipady = 0, padx = 10, pady = 0, sticky='news') 

    #### 자료 설명
    pop_label2 = Label(Pop_frame, \
        relief = "flat", text = " ※ 용어해설: 고령화율은 전체 인구 중 65세 이상 인구의 비율입니다.        " + "\n" + \
            "    ┖  7%이상: 고령화사회,  14%이상 : 고령사회,  20%이상 : 초고령사회",
        font = ("arial", 9, "normal"), padx = 0, pady = 0, anchor = "sw", bg = 'white', fg = 'black')
    pop_label2.grid(row = 7, column = 5, columnspan = 2, ipadx = 20, ipady = 0, padx = 10, pady = 0, sticky='news')
    #### 자료 설명
    pop_label3 = Label(Pop_frame, \
        relief = "flat", text = "※ 자료: 행정안전부 '주민등록인구'", \
        font = ("arial", 9, "normal"), padx = 20, pady = 0, anchor = "w", bg = 'white', fg = 'black')
    pop_label3.grid(row = 13, column = 0, ipadx = 20, ipady = 0, padx = 10, pady = 0, sticky = "w")
    #### 자료 설명
    if n_sido == "청주시" or n_tido == "청주시":
        pop_label4 = Label(Pop_frame, \
            relief = "flat", text = " ※ 분석유의: 청주시는 '14.7월 청원군과 통합했고, 통합전후는 지역이 다릅니다.",
            font = ("arial", 8, "normal"), padx = 0, pady = 0, anchor = "sw", bg = 'white', fg = 'black')
        pop_label4.grid(row = 13, column = 5, columnspan = 2, ipadx = 20, ipady = 0, padx = 10, pady = 0, sticky='news')

    ### 자료다운_a3_plot
    global a3_dt
    a3_dt = at.copy()
    downbt_a3_dt = Button(Pop_frame, text = "xls", bg = "mistyrose", width = 2, height = 1, relief = "flat", overrelief = "raised", command = a3_dtdown)
    downbt_a3_dt.grid(row = 12, column = 5, ipadx = 0, ipady = 0, padx = 11, pady = 1, sticky = 'sw')        


########################### CHART chapter Ⅱ. 고용 첫번째 꺾은선차트 그리기 ###########################
def emp_plot():
    ## excel 파일 열기
    df = pd.read_csv(resource_path('data/2_1_db_sex_emprate.csv'), encoding = 'cp949')

    ### 조건에 맞는 자료 추출
    z = df.loc[(df['시도'] == tido) & (df['행정구역'] == tigu) & (df['성별'] == "계"), ['연월', '15~64세 고용률 (%)']]
    z = z.rename(columns = {'15~64세 고용률 (%)' : n_tido})
    a = df.loc[(df['시도'] == sido) & (df['행정구역'] == sigu) & (df['성별'] == "계"), ['연월', '15~64세 고용률 (%)']]
    a = pd.merge(a, z, on = '연월')
    a = a.rename(columns = {'15~64세 고용률 (%)' : n_sido})
    x = a['연월']
    y = a[n_sido]
    z = a[n_tido]
    maxmin = df.loc[(((df['시도'] == sido) & (df['행정구역'] == sigu)) | ((df['시도'] == tido) & (df['행정구역'] == tigu))) & (df['성별'] == "계"), ['15~64세 고용률 (%)']]
    max_y = maxmin['15~64세 고용률 (%)'].max() + 0.9  ## y축 최대값
    min_y = maxmin['15~64세 고용률 (%)'].min() - 0.7  ## y축 최소값
    
    ## 꺾은선차트 그리기 ##
    fig = Figure(figsize = (4.6, 2.5), dpi = 100)
    ax = fig.add_subplot(111)
    fig.suptitle("[ '" + n_sido + "' 고용률(15~64세) 추이 ]", fontsize = 11, y=0.95)
    ax.plot(x, y, color = 'red', linewidth = 2, marker = "o", markeredgewidth = 1, \
        markeredgecolor = "red", markerfacecolor = 'yellow', markersize = 6, label = n_sido, markevery = 1)
    ax.plot(x, z, color = 'darkblue', linewidth = 1, marker = "o", markeredgewidth = 1, \
        markeredgecolor = 'darkblue', markerfacecolor = backcolor, markersize = 3, label = n_tido, markevery = 1, linestyle = '--')
    ax.set_ylim(min_y, max_y)
    ax.spines['top'].set_visible(False)  # 테두리 여부
    ax.spines['right'].set_visible(False)  # 테두리 여부
    ax.tick_params(right = False, left = False, bottom = False, labelleft=True, labelright=False, pad = 0) ### 눈금선 없애기, 라벨위치여부
    fig.legend(loc = (0.13, 0.73), fontsize = 9) # 범례 표시
    # y1 차트 데이터레이블 표시
    for idx, txt in enumerate(y):
        ax.text(x[idx], y[idx] + 0.07, txt, ha='center')
        
    global canvas
    canvas = FigureCanvasTkAgg(fig, master = emp_frame)
    canvas.draw()
    canvas.get_tk_widget().grid(row = 2, column = 0,  columnspan = 4, ipadx = 30, ipady = 30, padx = 25, pady = 2)    

    ### 차트 편집_b1_plot
    downbt_b1_plot = Button(emp_frame, text = "편집/저장", bg = "mistyrose", width = 6, relief = "flat", overrelief = "raised", command = b1_plot)
    downbt_b1_plot.grid(row = 2, column = 0, ipadx = 0, ipady = 0, padx = 25, pady = 2, sticky = 'nw')


    ##### ★★★★★★ 고용 첫번째 Treeview 그리기, dataframe 'a' 계속 사용
    a['차이'] = round(a[n_sido] - a[n_tido], 1)
    a = a[["연월", n_sido, "차이", n_tido]]
    a = a.rename(columns = ({n_sido : "○ " + n_sido + "(%)", "차이" : '※' + n_sido + "-" + n_tido + '(%p)', n_tido : "○ " + n_tido + "(%)"}))
    a = a.set_index('연월').T ### 데이터 전치
    a = a.reset_index(drop = False) ### 전치 후 '항목 열'을 'index'로 계속 가지고 감
    a = a.rename(columns = {'index' : '구분'})
    cols = list(a.columns)
    column_append_cols = ["c"+ str(i+1) for i in range(len(cols))]  ###★★★★★★ 리스트 누적해서 열리스트 한정(대전광역시 자치구 때문에 필요)
    
    treeview_empplot1 = ttk.Treeview(emp_frame, column= column_append_cols, show = 'headings', height = 1)
    
    ### Treeview heading
    for i, j in enumerate(cols):
        treeview_empplot1.column("# " + str(i+1), anchor=CENTER, width = 20)
        treeview_empplot1.heading("# " + str(i+1), text = j)  
    treeview_empplot1.column("c1", width = 83)  # 지역구분은 좀 넓게 추가 정의

    ### Treeview data insert
    for i in a.index:
        treeview_empplot1.insert('', 'end', text = i, values = (list(a.values[i])))

    treeview_empplot1.column(0, anchor= 'w') ### Treeview 내부 '구분'항목은 왼쪽 정렬
    treeview_empplot1.grid(row = 3, column = 0, columnspan = 4, ipadx = 0, ipady = 20, padx = 25, pady = 2, sticky=N+E+W+S) 
    
    # for i, j in enumerate(a.index):
    #     treeview_empplot1.insert('', 'end', text = i, values = (a.iloc[i, 0], a.iloc[i, 1], a.iloc[i, 2], a.iloc[i, 3], a.iloc[i, 5], \
    #         a.iloc[i, 6], a.iloc[i, 8], a.iloc[i, 7], a.iloc[i, 10], \
    #         a.iloc[i, 11], a.iloc[i, 13], a.iloc[i, 12], a.iloc[i, 15]))
    
    # ## 트리뷰 내부텍스트 위치 정렬
    # for i in enumerate(cols):
    #     treeview_empplot1.column(i, anchor= 'c')

    if (sigu == "전체" and tigu != "전체") or (sigu != "전체" and tigu == "전체") :
        emp_label2 = Label(emp_frame, \
            relief = "flat", text = " ※ 기초시군구(반기)와 광역시도(분기)간 청년고용률은 조사기준이 달라 서로 비교하지 않음", \
            font = ("arial", 9, "normal"), padx = 20, pady = 0, anchor = "w", bg = 'white', fg = 'black')
        emp_label2.grid(row = 5, column = 6,  columnspan = 3, ipadx = 0, ipady = 0, padx = 0, pady = 0, sticky = "nw")

    ### 자료다운_b1_dt
    global b1_dt
    b1_dt = a.copy()
    downbt_b1_dt = Button(emp_frame, text = "xls", bg = "mistyrose", width = 2, height = 1, relief = "flat", overrelief = "raised", command = b1_dtdown)
    downbt_b1_dt.grid(row = 4, column = 0, ipadx = 0, ipady = 0, padx = 26, pady = 0, sticky = 'nw')              

########################### 고용 두번째, 성별 바차트 그리기 ###########################
def emp_sex_plot():
    ## excel 파일 열기
    df = pd.read_csv(resource_path('data/2_1_db_sex_emprate.csv'), encoding = 'cp949')

    month_emp = df.sort_values('연월', ascending = False)
    month_emp = month_emp.iloc[0, 3]  ## 가져올 자료의 기준시점 설정

    ### 조건에 맞는 자료 추출 및 사전작업
    male = df.loc[(((df['시도'] == sido) & (df['행정구역'] == sigu)) | ((df['시도'] == tido) & (df['행정구역'] == tigu))) & \
        (df['성별'] == "남자") & (df['연월'] == month_emp), ['시도', '행정구역', '성별','연월', '15~64세 고용률 (%)']]
    female = df.loc[(((df['시도'] == sido) & (df['행정구역'] == sigu)) | ((df['시도'] == tido) & (df['행정구역'] == tigu))) & \
        (df['성별'] == "여자") & (df['연월'] == month_emp), ['시도', '행정구역', '성별','연월', '15~64세 고용률 (%)']]
    male['구분'] = male['행정구역']
    male.loc[male['행정구역'] == "전체", '구분'] = male['시도']
    male = male[['구분', '15~64세 고용률 (%)']]
    female['구분'] = female['행정구역']
    female.loc[female['행정구역'] == "전체", '구분'] = female['시도']
    female = female[['구분', '15~64세 고용률 (%)']]
    a = pd.merge(male,female, on = "구분")
    a = a.rename(columns = {"15~64세 고용률 (%)_x" : "남자", "15~64세 고용률 (%)_y" : "여자"})
    a = a[["구분", "남자", "여자"]].set_index("구분").T
    a = a.reset_index(drop = False)
    at = a.copy() ###########################Treeview용 데이터 카피#######
    a = a[['index', n_sido, n_tido]].T
    a = a.reset_index(drop = False)
    a = a.rename(columns = a.iloc[0])
    a = a.drop(a.index[0])
    
    m = a['남자'].reset_index(drop=True)
    f = a['여자'].reset_index(drop=True)
    t = pd.concat([m, f], ignore_index = True)
    max_t = t.max() + 19
    min_t = t.min() - 19

    reg = a['index']  # 'x축 명' 표시용

    x_axis = np.arange(len(reg))

    ## 바차트 그리기 ##
    fig = Figure(figsize = (2.2, 2.5), dpi = 100)
    ax = fig.add_subplot(111)
    fig.suptitle("[ 성별고용률 (" + month_emp +")]", fontsize = 11, y=0.95)
    ax.bar(x_axis -0.2, m, width = 0.4, label = '남(%)', color = 'forestgreen')
    ax.bar(x_axis +0.2, f, width=0.4, label = '여(%)', color = 'darkorange')
    ax.set_ylim(min_t, max_t)
    ax.set_xticks(x_axis, reg)
    fig.legend(loc = (0.13, 0.72), fontsize = 9) # 범례 표시
    ax.spines['bottom'].set_visible(True)  # 테두리 여부
    ax.spines['top'].set_visible(False)  # 테두리 여부
    ax.spines['right'].set_visible(False)  # 테두리 여부
    ax.spines['left'].set_visible(False)  # 테두리 여부
    ax.tick_params(right = False, left = False, bottom = False, labelleft=False, labelright=False, pad = 0) ### 눈금선 없애기, 라벨위치여부
        
    # male과 female 차트 데이터레이블 표시
    for idx, txt in enumerate(m):
        ax.text(x_axis[idx]-0.2, m[idx]+1.5, txt, ha='center')    
    for idx, txt in enumerate(f):
        ax.text(x_axis[idx]+0.2, f[idx]+1.5, txt, ha='center')

    global canvas
    # canvas = Canvas(Pop_frame, width = 1700, height = 300, bg = "white")
    canvas = FigureCanvasTkAgg(fig, master = emp_frame)
    canvas.draw()
    canvas.get_tk_widget().grid(row = 2, column = 4,  columnspan = 2, ipadx = 45, ipady = 20, padx = 10, pady = 2, sticky = 'nsew')

    ### 차트 편집_b2_plot
    downbt_b2_plot = Button(emp_frame, text = "편집/저장", bg = "mistyrose", width = 6, relief = "flat", overrelief = "raised", command = b2_plot)
    downbt_b2_plot.grid(row = 2, column = 4, ipadx = 6, ipady = 0, padx = 10, pady = 2, sticky = 'nw')


    # ########## ★★★★★★  Treeview, 고용 두번째 성별고용률, dataframe 'at' 계속 사용
    at = at[["index", n_sido, n_tido]]
    at["차이"] = round(at[n_sido] - at[n_tido],2)
    at = at[["index", n_sido, "차이", n_tido]].reset_index(drop=True)
    at = at.rename(columns = {"index" : "구분", n_sido : "○ " + n_sido + "(%)", "차이" : "※" + n_sido + "-" + n_tido + '(%p)', n_tido : "○ " + n_tido + "(%)"}).T.reset_index(drop=False)
    at = at.rename(columns = at.iloc[0])  ##### 첫행을 변수로 바꾸기
    at = at.drop(at.index[0])  ###### 바꾼 후 기존 첫행 삭제
    at = at.reset_index(drop = True)  ###### 인덱스 초기화해줘야 정상 작동

    cols = list(at.columns)
    column_append_cols = ["c"+ str(i+1) for i in range(len(cols))]  ###★★★★★★ 리스트 누적해서 열리스트 한정(대전광역시 자치구 때문에 필요)
    
    treeview_empplot2 = ttk.Treeview(emp_frame, column= column_append_cols, show = 'headings', height = 1)
    
    ### Treeview heading
    for i, j in enumerate(cols):
        treeview_empplot2.column("# " + str(i+1), anchor=CENTER, width = 20)
        treeview_empplot2.heading("# " + str(i+1), text = j)  
    treeview_empplot2.column("c1", width = 60)  # 지역구분은 좀 넓게 추가 정의

    ### Treeview data insert
    for i in at.index:
        treeview_empplot2.insert('', 'end', text = i, values = (list(at.values[i])))

    treeview_empplot2.column(0, anchor= 'w') ### Treeview 내부 '구분'항목은 왼쪽 정렬
    treeview_empplot2.grid(row = 3, column = 4,  columnspan = 2, ipadx = 45, ipady = 0, padx = 10, pady = 2, sticky='nesw') 

    ### 성별고용률 엑셀 바로 실행_b2_dt
    global b2_dt
    b2_dt = at.copy()
    downbt_b2_dt = Button(emp_frame, text = "xls", bg = "mistyrose", width = 2, height = 1, relief = "flat", overrelief = "raised", command = b2_dtdown)
    downbt_b2_dt.grid(row = 4, column = 4, ipadx = 0, ipady = 0, padx = 11, pady = 0, sticky = 'nw')      


########################### 고용 세번째 연령별_청년 라인차트 그리기 ###########################
def emp_youngage_plot():
    ## 1. 청년고용률
    ## excel 파일 열기
    df = pd.read_csv(resource_path('data/2_2_db_age_emprate_adapt.csv'), encoding = 'cp949')

    ### 조건에 맞는 자료 추출 및 사전작업
    if gungucbox.get() == '전체' and Tgungucbox.get() == '전체':
        ### 청년 꺾은선 차트 용 dataframe 'c'
        a = df.loc[(df['시도'] == sido) & (df['행정구역'] == sigu), ['시점', '고용률 (%)']]
        b = df.loc[(df['시도'] == tido) & (df['행정구역'] == tigu), ['시점', '고용률 (%)']]
        c = pd.merge(a,b, on = "시점")
        c = c.sort_values('시점', ascending = True)
        c = c.reset_index(drop = True)
        a = c.copy()
        c.rename(columns = {"고용률 (%)_x" : n_sido, "고용률 (%)_y" : n_tido}, inplace = True)
        x = c['시점']
        y = c[n_sido]
        z = c[n_tido]
        maxmin = df.loc[((df['시도'] == sido) & (df['행정구역'] == sigu)) | ((df['시도'] == tido) & (df['행정구역'] == tigu)), ['고용률 (%)']]
        max_y = maxmin['고용률 (%)'].max() + 0.9  ## y축 최대값
        min_y = maxmin['고용률 (%)'].min() - 0.7  ## y축 최소값

        fig = Figure(figsize = (4.6, 2.7), dpi = 100)
        ax = fig.add_subplot(111)
        fig.suptitle("[ '" + n_sido + "' 청년고용률(15-29세) ]", fontsize = 11, y=0.95)
        ax.plot(x, y, color = 'red', linewidth = 2, marker = "o", markeredgewidth = 1, \
            markeredgecolor = "firebrick", markerfacecolor = "yellow", markersize = 6, label = n_sido, markevery = 1)
        ax.plot(x, z, color = 'darkblue', linewidth = 1, marker = "o", markeredgewidth = 1, \
            markeredgecolor = "darkblue", markerfacecolor = "darkblue", markersize = 3, label = n_tido, markevery = 1, linestyle = '--')
        ax.set_ylim(min_y, max_y)
        ax.spines['top'].set_visible(False)  # 테두리 여부
        ax.spines['right'].set_visible(False)  # 테두리 여부
        ax.tick_params(right = False, left = False, bottom = False, labelleft=True, labelright=False, pad = 0) ### 눈금선 없애기, 라벨위치여부
        fig.legend(loc = (0.13, 0.73), fontsize = 9) # 범례 표시
        # y1 차트 데이터레이블 표시
        for idx, txt in enumerate(y):
            ax.text(x[idx], y[idx] + 0.3, txt, ha='center')
            
        global canvas
        canvas = FigureCanvasTkAgg(fig, master = emp_frame)
        canvas.draw()
        canvas.get_tk_widget().grid(row = 2, column = 6,  columnspan = 3, ipadx = 15, ipady = 20, padx = 25, pady = 2) 

        #### Treeview용 데이터 편집 dataframe 'a' 및 Treeview 실행
        a.rename(columns = {"고용률 (%)_x" : "○ " + n_sido + "(%)", "고용률 (%)_y" : "○ " + n_tido + "(%)"}, inplace = True)
        a["차이"] = round(a["○ " + n_sido + "(%)"] - a["○ " + n_tido + "(%)"],2)
        a = a[["시점", "○ " + n_sido + "(%)", "차이", "○ " + n_tido + "(%)"]].set_index("시점")
        a = a.rename(columns = {"index" : "구분", "차이" : "※" + n_sido + "-" + n_tido + "(%p)"})
        a = a.T
        a = a.reset_index(drop = False)
        a = a.rename(columns = {"index" : "구분", "차이" : "※" + n_sido + "-" + n_tido + "(%p)"})

        cols = list(a.columns)
        column_append_cols = ["c"+ str(i+1) for i in range(len(cols))]  ###★★★★★★ 리스트 누적해서 열리스트 한정(대전광역시 자치구 때문에 필요)
        
        treeview_empplot3 = ttk.Treeview(emp_frame, column= column_append_cols, show = 'headings', height = 1)
        
        ### Treeview heading
        for i, j in enumerate(cols):
            treeview_empplot3.column("# " + str(i+1), anchor=CENTER, width = 20)
            treeview_empplot3.heading("# " + str(i+1), text = j)  
        treeview_empplot3.column("c1", width = 80)  # 지역구분은 좀 넓게 추가 정의

        ### Treeview data insert
        for i in a.index:
            treeview_empplot3.insert('', 'end', text = i, values = (list(a.values[i])))

        treeview_empplot3.column(0, anchor= 'w') ### Treeview 내부 '구분'항목은 왼쪽 정렬
        treeview_empplot3.grid(row = 3, column = 6,  columnspan = 3, ipadx = 15, ipady = 20, padx = 25, pady = 2, sticky=N+E+W+S) 

        ### 청년고용률 엑셀 바로 실행 첫번째_tt_b3_dt
        global tt_b3_dt
        tt_b3_dt = a.copy()
        downbt_tt_b3_dt = Button(emp_frame, text = "xls", bg = "mistyrose", width = 2, height = 1, relief = "flat", overrelief = "raised", command = tt_b3_dtdown)
        downbt_tt_b3_dt.grid(row = 4, column = 6, ipadx = 0, ipady = 0, padx = 26, pady = 0, sticky = 'nw') 

    elif gungucbox.get() != '전체' and Tgungucbox.get() != '전체' and sido != ['서울특별시', '부산광역시', '대구광역시', '인천광역시', '광주광역시', \
        '대전광역시', '울산광역시'] and tido != ['서울특별시', '부산광역시', '대구광역시', '인천광역시', '광주광역시', '대전광역시', '울산광역시']:
        ### 청년 꺾은선 차트 용 dataframe 'c'
        a = df.loc[(df['시도'] == sido) & (df['행정구역'] == sigu), ['시점', '고용률 (%)']]
        b = df.loc[(df['시도'] == tido) & (df['행정구역'] == tigu), ['시점', '고용률 (%)']]
        c = pd.merge(a,b, on = "시점")
        c = c.sort_values('시점', ascending = True)
        c = c.reset_index(drop = True)
        a = c.copy()
        c.rename(columns = {"고용률 (%)_x" : n_sido, "고용률 (%)_y" : n_tido}, inplace = True)
        x = c['시점']
        y = c[n_sido]
        z = c[n_tido]
        maxmin = df.loc[((df['시도'] == sido) & (df['행정구역'] == sigu)) | ((df['시도'] == tido) & (df['행정구역'] == tigu)), ['고용률 (%)']]
        max_y = maxmin['고용률 (%)'].max() + 0.9  ## y축 최대값
        min_y = maxmin['고용률 (%)'].min() - 0.7  ## y축 최소값

        fig = Figure(figsize = (4.6, 2.7), dpi = 100)
        ax = fig.add_subplot(111)
        fig.suptitle("[ '" + n_sido + "' 청년고용률(15-29세) ]", fontsize = 11, y=0.95)
        ax.plot(x, y, color = 'red', linewidth = 2, marker = "o", markeredgewidth = 1, \
            markeredgecolor = "firebrick", markerfacecolor = "yellow", markersize = 6, label = n_sido, markevery = 1)
        ax.plot(x, z, color = 'darkblue', linewidth = 1, marker = "o", markeredgewidth = 1, \
            markeredgecolor = "darkblue", markerfacecolor = "darkblue", markersize = 3, label = n_tido, markevery = 1, linestyle = '--')
        ax.set_ylim(min_y, max_y)
        ax.spines['top'].set_visible(False)  # 테두리 여부
        ax.spines['right'].set_visible(False)  # 테두리 여부
        ax.tick_params(right = False, left = False, bottom = False, labelleft=True, labelright=False, pad = 0) ### 눈금선 없애기, 라벨위치여부
        fig.legend(loc = (0.13, 0.73), fontsize = 9) # 범례 표시
        # y1 차트 데이터레이블 표시
        for idx, txt in enumerate(y):
            ax.text(x[idx], y[idx] + 0.3, txt, ha='center')
            
        canvas = FigureCanvasTkAgg(fig, master = emp_frame)
        canvas.draw()
        canvas.get_tk_widget().grid(row = 2, column = 6,  columnspan = 3, ipadx = 15, ipady = 20, padx = 25, pady = 2) 

        #### Treeview용 데이터 편집 dataframe 'a' 및 Treeview 실행
        a.rename(columns = {"고용률 (%)_x" : "○ " + n_sido + "(%)", "고용률 (%)_y" : "○ " + n_tido + "(%)"}, inplace = True)
        a["차이"] = round(a["○ " + n_sido + "(%)"] - a["○ " + n_tido + "(%)"],2)
        a = a[["시점", "○ " + n_sido + "(%)", "차이", "○ " + n_tido + "(%)"]].set_index("시점")
        a = a.rename(columns = {"index" : "구분", "차이" : "※" + n_sido + "-" + n_tido + "(%p)"})
        a = a.T
        a = a.reset_index(drop = False)
        a = a.rename(columns = {"index" : "구분", "차이" : "※" + n_sido + "-" + n_tido + "(%p)"})

        cols = list(a.columns)
        column_append_cols = ["c"+ str(i+1) for i in range(len(cols))]  ###★★★★★★ 리스트 누적해서 열리스트 한정(대전광역시 자치구 때문에 필요)
        
        treeview_empplot3 = ttk.Treeview(emp_frame, column= column_append_cols, show = 'headings', height = 1)
        
        ### Treeview heading
        for i, j in enumerate(cols):
            treeview_empplot3.column("# " + str(i+1), anchor=CENTER, width = 20)
            treeview_empplot3.heading("# " + str(i+1), text = j)  
        treeview_empplot3.column("c1", width = 80)  # 지역구분은 좀 넓게 추가 정의

        ### Treeview data insert
        for i in a.index:
            treeview_empplot3.insert('', 'end', text = i, values = (list(a.values[i])))

        treeview_empplot3.column(0, anchor= 'w') ### Treeview 내부 '구분'항목은 왼쪽 정렬
        treeview_empplot3.grid(row = 3, column = 6,  columnspan = 3, ipadx = 15, ipady = 20, padx = 25, pady = 2, sticky=N+E+W+S) 

        ### 청년고용률 엑셀 바로 실행 두번째_tt_b3_dt: 모두 9개도의 기초시군구 선택
        global rr_b3_dt
        rr_b3_dt = a.copy()
        downbt_rr_b3_dt = Button(emp_frame, text = "xls", bg = "mistyrose", width = 2, height = 1, relief = "flat", overrelief = "raised", command = rr_b3_dtdown)
        downbt_rr_b3_dt.grid(row = 4, column = 6, ipadx = 0, ipady = 0, padx = 26, pady = 0, sticky = 'nw') 
        


    else:
        a = df.loc[(df['시도'] == sido) & (df['행정구역'] == sigu), ['시점', '고용률 (%)']]
        a = a.sort_values('시점', ascending = True)
        a = a.reset_index(drop=True)

        x = a['시점']
        y = a['고용률 (%)']
        
        max_y = a['고용률 (%)'].max() + 0.9  ## y축 최대값
        min_y = a['고용률 (%)'].min() - 0.7  ## y축 최소값

        fig = Figure(figsize = (4.6, 2.5), dpi = 100)
        ax = fig.add_subplot(111)
        fig.suptitle("[ " + n_sido + " 청년고용률(15-29세) ]", fontsize = 11, y=0.95)
        ax.plot(x, y, color = 'red', linewidth = 2, marker = "o", markeredgewidth = 1, \
            markeredgecolor = "firebrick", markerfacecolor = "yellow", markersize = 6, label = n_sido, markevery = 1)
        ax.set_ylim(min_y, max_y)
        ax.spines['top'].set_visible(False)  # 테두리 여부
        ax.spines['right'].set_visible(False)  # 테두리 여부
        ax.tick_params(right = False, left = False, bottom = False, labelleft=True, labelright=False, pad = 0) ### 눈금선 없애기, 라벨위치여부
        # y1 차트 데이터레이블 표시
        for idx, txt in enumerate(y):
            ax.text(x[idx], y[idx] + 0.3, txt, ha='center')
            
        # global canvas
        canvas = FigureCanvasTkAgg(fig, master = emp_frame)
        canvas.draw()
        canvas.get_tk_widget().grid(row = 2, column = 6,  columnspan = 3, ipadx = 15, ipady = 20, padx = 25, pady = 2) 

        #### Treeview, 청년고용률 첫번째 데이터 편집 dataframe 'a' 및 Treeview 실행
        a.rename(columns = {"고용률 (%)" : "○ " + n_sido}, inplace = True)
        a[" -전기대비증감(%p)"] = 0.0
        for i in a.index:
            a.iloc[i,2] = round(a.iloc[i, 1] - a.iloc[i-1, 1],1)
        a.iloc[0, 2] = "-"
        a = a.set_index("시점").T
        a = a.reset_index(drop = False)
        a = a.rename(columns = {"index" : "구분"})

        cols = list(a.columns)
        column_append_cols = ["c"+ str(i+1) for i in range(len(cols))]  ###★★★★★★ 리스트 누적해서 열리스트 한정(대전광역시 자치구 때문에 필요)
        
        treeview_empplot3_1 = ttk.Treeview(emp_frame, column= column_append_cols, show = 'headings', height = 1)
        
        ### Treeview heading
        for i, j in enumerate(cols):
            treeview_empplot3_1.column("# " + str(i+1), anchor=CENTER, width = 20)
            treeview_empplot3_1.heading("# " + str(i+1), text = j)  
        treeview_empplot3_1.column("c1", width = 95)  # 지역구분은 좀 넓게 추가 정의

        ### Treeview data insert
        for i in a.index:
            treeview_empplot3_1.insert('', 'end', text = i, values = (list(a.values[i])))

        treeview_empplot3_1.column(0, anchor= 'w') ### Treeview 내부 '구분'항목은 왼쪽 정렬
        treeview_empplot3_1.grid(row = 3, column = 6,  columnspan = 3, ipadx = 15, ipady = 20, padx = 25, pady = 2, sticky=N+E+W) 
        
        ### 청년고용률 엑셀 바로 실행 세번째_etc1_b3_dt: 그외 시군구 선택
        global etc1_b3_dt
        etc1_b3_dt = a.copy()
        downbt_etc1_b3_dt = Button(emp_frame, text = "xls", bg = "mistyrose", width = 2, height = 1, relief = "flat", overrelief = "raised", command = etc1_b3_dtdown)
        downbt_etc1_b3_dt.grid(row = 4, column = 6, ipadx = 0, ipady = 0, padx = 26, pady = 0, sticky = 'nw')         

        ############################################ Treeview, 청년고용률 두번째 선택시군 데이터 편집 dataframe 'a' 및 Treeview 실행
        tot = df.loc[(df['시도'] == tido) & (df['행정구역'] == tigu), ['시점', '고용률 (%)']]
        tot = tot.sort_values('시점', ascending = True)
        tot = tot.reset_index(drop = True)
        tot.rename(columns = {"고용률 (%)" : "○ " + n_tido}, inplace = True)
        tot[" -전기대비증감(%p)"] = 0.0
        for i in tot.index:
            tot.iloc[i,2] = round(tot.iloc[i, 1] - tot.iloc[i-1, 1],1)
        tot.iloc[0, 2] = "-"
        tot = tot.set_index("시점").T
        tot = tot.reset_index(drop = False)
        tot = tot.rename(columns = {"index" : "구분"})

        cols = list(tot.columns)
        column_append_cols = ["c"+ str(i+1) for i in range(len(cols))]  ###★★★★★★ 리스트 누적해서 열리스트 한정(대전광역시 자치구 때문에 필요)
        
        treeview_empplot3_2 = ttk.Treeview(emp_frame, column= column_append_cols, show = 'headings', height = 2)
        
        ### Treeview heading
        for i, j in enumerate(cols):
            treeview_empplot3_2.column("# " + str(i+1), anchor=CENTER, width = 20)
            treeview_empplot3_2.heading("# " + str(i+1), text = j)  
        treeview_empplot3_2.column("c1", width = 65)  # 지역구분은 좀 넓게 추가 정의

        ### Treeview data insert
        for i in tot.index:
            treeview_empplot3_2.insert('', 'end', text = i, values = (list(tot.values[i])))

        treeview_empplot3_2.column(0, anchor= 'w') ### Treeview 내부 '구분'항목은 왼쪽 정렬
        treeview_empplot3_2.grid(row = 7, column = 6,  columnspan = 3, ipadx = 15, ipady = 0, padx = 25, pady = 2, sticky=N+E+W+S) 

        ### 청년고용률 엑셀 바로 실행 세번째_etc2_b3_dt: 그외 시군구 선택
        global etc2_b3_dt
        etc2_b3_dt = tot.copy()
        downbt_etc2_b3_dt = Button(emp_frame, text = "xls", bg = "mistyrose", width = 2, height = 1, relief = "flat", overrelief = "raised", command = etc2_b3_dtdown)
        downbt_etc2_b3_dt.grid(row = 8, column = 6, ipadx = 0, ipady = 0, padx = 26, pady = 0, sticky = 'nw') 

    ### 차트 편집_b3_plot
    downbt_b3_plot = Button(emp_frame, text = "편집/저장", bg = "mistyrose", width = 6, relief = "flat", overrelief = "raised", command = b3_plot)
    downbt_b3_plot.grid(row = 2, column = 6, ipadx = 6, ipady = 0, padx = 25, pady = 2, sticky = 'nw')


########################### 고용 다섯번째 직업별 방사형차트 그리기 ###########################
def spyder_plot_a():
    global job_month
    job_month = '2022.1/2'
    emp_label0 = Label(emp_frame, relief = "flat", text = " " + "\n" + "< 한걸음 더 ① > '직업별' 취업자 비중(" + str(job_month) + ")            ", font = ("arial", 15, "bold"), padx = 20, pady = 0, bg = 'white', fg = 'black') ## 공백행 삽입
    emp_label0.grid(row = 9, column = 0, columnspan = 5, ipadx = 20, ipady = 10, padx = 10, pady = 2, sticky = "nsew")

    df = pd.read_csv(resource_path('data/2_3_db_job.csv'), encoding = 'cp949')
    df1 = df.loc[((df['시도'] == tido) & (df['행정구역'] == tigu) & (df['직업별'] != "계")), ['직업별', job_month]]
    df1['비중'] = round(df1[job_month] / df1[job_month].sum() * 100, 1)
    df2 = df.loc[((df['시도'] == sido) & (df['행정구역'] == sigu) & (df['직업별'] != "계")), ['직업별', job_month]]
    df2['비중'] = round(df2[job_month] / df2[job_month].sum() * 100, 1)
    df3 = pd.merge(df1, df2, on = "직업별")
    at = df3.copy()
    df3 = df3[['직업별', '비중_x', '비중_y']]
    df3 = df3.rename(columns = {'비중_x' : n_tido, '비중_y' : n_sido})
    df = df3.reset_index(drop = True)
    labels = df['직업별']

    fig = plt.figure(figsize = (5, 2.8), dpi = 100)

    ax = fig.add_subplot(111, projection="polar")
    theta = np.arange(len(df) + 1) / float(len(df)) * 2 * np.pi
    values = df[n_sido].values
    values = np.append(values, values[0])
    ax.plot(theta, values, color="crimson", label=n_sido)
    ax.tick_params(pad=10)
    plt.xticks(theta[:-1], labels, fontsize=10)
    ax.fill(theta, values, 'crimson', alpha=0.1)

    # ax = fig.add_subplot(111, projection="polar")
    # theta = np.arange(len(df) + 1) / float(len(df)) * 2 * np.pi
    values = df[n_tido].values
    values = np.append(values, values[0])
    ax.plot(theta, values, color="cornflowerblue", label= n_tido)
    ax.tick_params(pad=10)
    ax.fill(theta, values, 'cornflowerblue', alpha=0.1)
    plt.legend(bbox_to_anchor=(0.1, 0.98))
    fig.tight_layout()  ## 이미지 타이트하게
    
    global canvas
    # canvas = Canvas(Pop_frame, width = 1700, height = 300, bg = "white")
    canvas = FigureCanvasTkAgg(fig, master = emp_frame)
    canvas.draw()
    canvas.get_tk_widget().grid(row = 10, column = 0,  columnspan = 4, ipadx = 0, ipady = 20, padx = 0, pady = 0, sticky='se')

    ### 차트 편집_b4_plot
    downbt_b4_plot = Button(emp_frame, text = "편집/저장", bg = "mistyrose", width = 6, relief = "flat", overrelief = "raised", command = b4_plot)
    downbt_b4_plot.grid(row = 10, column = 0, ipadx = 6, ipady = 0, padx = 70, pady = 5, sticky='nw')

    ################## ★★★★★★ 방사형차트 Treeview 그리기, dataframe 'at' 계속 사용 ##################
    at = at.rename(columns = {'2021.2/2_x' : n_tido +"(천명)", '2021.2/2_y' : n_sido + "천명", '비중_x' : n_tido + "비중(%)", '비중_y' : n_sido + "비중(%)"})
    at = at[["직업별", n_sido + "비중(%)", n_tido + "비중(%)"]]
    cols = list(at.columns)
    column_append_cols = ["c"+ str(i+1) for i in range(len(cols))]  ###★★★★★★ 리스트 누적해서 열리스트 한정(대전광역시 자치구 때문에 필요)
    
    treeview_spyder = ttk.Treeview(emp_frame, column= column_append_cols, show = 'headings', height = 6)
    ### Treeview heading
    for i, j in enumerate(cols):
        treeview_spyder.column("# " + str(i+1), anchor=CENTER, width = 70)
        treeview_spyder.heading("# " + str(i+1), text = j)  
    treeview_spyder.column("c1", width = 100)  # 지역구분은 좀 넓게 추가 정의

    ### Treeview data insert
    for i in at.index:
        treeview_spyder.insert('', 'end', text = i, values = (list(at.values[i])))

    treeview_spyder.grid(row = 10, column = 3, columnspan = 2, ipadx = 10, ipady = 0, padx = 10, pady = 5, sticky='sew')  
    
    ### 엑셀바로 보기 및 다운: 직업별 취업자 비중_b4_dt
    global b4_dt
    b4_dt = at.copy()
    downbt_b4_dt = Button(emp_frame, text = "xls", bg = "mistyrose", width = 2, height = 1, relief = "flat", overrelief = "raised", command = b4_dtdown)
    downbt_b4_dt.grid(row = 11, column = 4, ipadx = 0, ipady = 0, padx = 11, pady = 0, sticky = 'ne')  


########################### 고용 여섯번째_종사상지위별 방사형차트 그리기 ###########################
def spyder_plot_b():
    emp_label0 = Label(emp_frame, relief = "flat", text = " " + "\n" + "< 한걸음 더 ② > '종사상지위별' 취업자 비중(" + str(job_month) + ")         ", font = ("arial", 15, "bold"), padx = 20, pady = 0, bg = 'white', fg = 'black') ## 공백행 삽입
    emp_label0.grid(row = 9, column = 5, columnspan = 4, ipadx = 20, ipady = 10, padx = 10, pady = 2, sticky = "nsew")

    df = pd.read_csv(resource_path('data/2_3_db_joblevel.csv'), encoding = 'cp949')
    df1 = df.loc[((df['시도'] == tido) &(df['행정구역'] == tigu) & (df['종사상지위별'] != "계")), ['종사상지위별', job_month]]
    df1['비중'] = round(df1[job_month] / df1[job_month].sum() * 100, 1)
    df2 = df.loc[((df['시도'] == sido) & (df['행정구역'] == sigu) & (df['종사상지위별'] != "계")), ['종사상지위별', job_month]]
    df2['비중'] = round(df2[job_month] / df2[job_month].sum() * 100, 1)
    df3 = pd.merge(df1, df2, on = "종사상지위별")
    at = df3.copy()
    df3 = df3[['종사상지위별', '비중_x', '비중_y']]
    df3 = df3.rename(columns = {'비중_x' : n_tido, '비중_y' : n_sido})
    df = df3.reset_index(drop = True)
    labels = df['종사상지위별']

    fig_s2 = plt.figure(figsize = (5, 2.8), dpi = 100)

    ax = fig_s2.add_subplot(111, projection="polar")
    theta = np.arange(len(df) + 1) / float(len(df)) * 2 * np.pi
    values = df[n_sido].values
    values = np.append(values, values[0])
    ax.plot(theta, values, color="crimson", label=n_sido)
    ax.tick_params(pad=10)
    plt.xticks(theta[:-1], labels, fontsize=10)
    ax.fill(theta, values, 'crimson', alpha=0.1)

    values = df[n_tido].values
    values = np.append(values, values[0])
    ax.plot(theta, values, color="cornflowerblue", label= n_tido)
    ax.tick_params(pad=10)
    ax.fill(theta, values, 'cornflowerblue', alpha=0.1)
    plt.legend(bbox_to_anchor=(0.1, 0.98))
    fig_s2.tight_layout()  ## 이미지 타이트하게
    
    global canvas
    canvas = FigureCanvasTkAgg(fig_s2, master = emp_frame)
    canvas.draw()
    canvas.get_tk_widget().grid(row = 10, column = 5,  columnspan = 3, ipadx = 0, ipady = 20, padx = 10, pady = 2, sticky='se')

    ### 차트 편집_b5_plot
    downbt_b5_plot = Button(emp_frame, text = "편집/저장", bg = "mistyrose", width = 6, relief = "flat", overrelief = "raised", command = b5_plot)
    downbt_b5_plot.grid(row = 10, column = 7, ipadx = 6, ipady = 0, padx = 10, pady = 2, sticky='ne')

    ################## ★★★★★★ 방사형두번째차트 Treeview 그리기, dataframe 'at' 계속 사용 ##################
    at = at.rename(columns = {'2021.2/2_x' : n_tido + "(천명)", '2021.2/2_y' : n_sido +"천명", '비중_x' : n_tido + "비중(%)", '비중_y' : n_sido + "비중(%)"})
    at = at[["종사상지위별", n_sido + "비중(%)", n_tido + "비중(%)"]]
    cols = list(at.columns)
    column_append_cols = ["c"+ str(i+1) for i in range(len(cols))]  ###★★★★★★ 리스트 누적해서 열리스트 한정(대전광역시 자치구 때문에 필요)
    
    treeview_spyder = ttk.Treeview(emp_frame, column= column_append_cols, show = 'headings', height = 3)
    ### Treeview heading
    for i, j in enumerate(cols):
        treeview_spyder.column("# " + str(i+1), anchor=CENTER, width = 70)
        treeview_spyder.heading("# " + str(i+1), text = j)  
    treeview_spyder.column("c1", width = 90)  # 지역구분은 좀 넓게 추가 정의

    ### Treeview data insert
    for i in at.index:
        treeview_spyder.insert('', 'end', text = i, values = (list(at.values[i])))

    treeview_spyder.grid(row = 10, column = 7, columnspan = 2, ipadx = 30, ipady = 0, padx = 20, pady = 10, sticky='sew')  

    ### 엑셀 바로 실행 및 자료 보기: 종사상지위별 취업자비중_b5_dt
    global b5_dt
    b5_dt = at.copy()
    downbt_b5_dt = Button(emp_frame, text = "xls", bg = "mistyrose", width = 2, height = 1, relief = "flat", overrelief = "raised", command = b5_dtdown)
    downbt_b5_dt.grid(row = 11, column = 8, ipadx = 0, ipady = 0, padx = 21, pady = 0, sticky = 'ne')      

    ###### 참고자료 문구 ############
    global emp_label1, emp_label2
    emp_label1 = Label(emp_frame, \
        relief = "flat", text = " ※ 자료: 통계청 '경제활동인구조사' 및 '지역별고용조사'", \
        font = ("arial", 9, "normal"), padx = 20, pady = 0, anchor = "w", bg = 'white', fg = 'black')
    emp_label1.grid(row = 12, column = 0, columnspan = 4, ipadx = 0, ipady = 0, padx = 0, pady = 0, sticky = "nw")
    

########################### CHART chapter Ⅲ. 사업체종사자, 첫번째 가로바차트 그리기 ###########################
def sanup_bar():
    ## csv 파일 열기
    df = pd.read_csv(resource_path('data/3_1_db_company.csv'), encoding = 'cp949')
    
    ### 조건에 맞는 자료 추출 및 머지
    x = df.loc[(df['시도'] == tido) & (df['행정구역'] == tigu) & (df['산업대'] != "전체"), ['산업대', '2020 년']]
    y = df.loc[(df['시도'] == sido) & (df['행정구역'] == sigu) & (df['산업대'] != "전체"), ['산업대', '2020 년']]

    a = pd.merge(x,y, on = '산업대')
    a = a.sort_values('산업대', ascending = False)
    # a = a.reset_index(drop=True)

    ## 바차트 그리기 ##
    fig = Figure(figsize = (5.19, 3.5), dpi = 100)
    ax = fig.add_subplot(1,2,1)
    ax.barh(a['산업대'], a["2020 년_y"], color='slateblue')
    # fig.patch.set_visible(False)
    # ax.axis('off')
    ax.grid(False)
    ax.tick_params(labelleft=False, labelright=True, pad = 0)  
    ax.spines['bottom'].set_visible(True)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    # ax.spines['left'].set_visible(False)
    ax.tick_params(axis = 'y', length = 30, right = False, left = False, bottom = False) ### 눈금선 없애기
    # ax.set_xlabel("명", fontweight ='bold', fontsize=13, loc="left")  ## x축 라벨 별도로 표시할 때
    fig.tight_layout()  ## 이미지 타이트하게

    ## 사업체수 두번째 : 비교지역 바차트 그리기 ##
    # fig = Figure(figsize = (2.9, 3.7), dpi = 100)
    ax = fig.add_subplot(1,2,2)
    fig.suptitle("     - " + n_sido + " 사업체수 →                                            ← " + n_tido + " 사업체수-", fontsize = 11, x = 0.5, y = 0.95)
    ax.barh(a['산업대'], -a["2020 년_x"], color='slateblue')
    ax.spines['bottom'].set_visible(True)  # 테두리 여부
    ax.spines['top'].set_visible(False)  # 테두리 여부
    # ax.spines['right'].set_visible(False)  # 테두리 여부
    ax.spines['left'].set_visible(False)  # 테두리 여부
    ax.tick_params(right = False, left = False, bottom = False, labelleft=False, labelright=False, pad = 0) ### 눈금선 없애기, 라벨위치여부

    # ax.set_xlabel("명", fontweight ='bold', fontsize=13, loc="left")  ## x축 라벨 별도로 표시할 때
    fig.tight_layout()  ## 이미지 타이트하게

    # fig.legend(loc = (0.13, 0.71), fontsize = 9) # 범례 표시 안함

    canvas = FigureCanvasTkAgg(fig, master = sanup_frame)
    canvas.draw()
    canvas.get_tk_widget().grid(row = 2, column = 0, ipadx = 70, ipady = 0, padx = 30, pady = 2)
    # canvas.get_tk_widget().pack()

    ### 차트 편집_c1_plot
    downbt_c1_plot = Button(sanup_frame, text = "편집/저장", bg = "mistyrose", width = 6, relief = "flat", overrelief = "raised", command = c1_plot)
    downbt_c1_plot.grid(row = 2, column = 0, ipadx = 6, ipady = 0, padx = 30, pady = 2, sticky = 'nw')

    # #### table 구문
    # global table_i
    # #####★★★★★ 테이블 추가 ★★★★★★★
    # table_i = Table(sanup_frame, showtoolbar = True, showstatusbar = True, dataframe= a)
    # table_i.model.df = a
    # table_i.editable = False
    # # table_i.colheader.bgcolor = 'lightgreen'
    # table_i.grid(row= 2, column =2)

    ######################### Treeview, 사업체수 Treeview  #########################
    a = a.sort_values('산업대', ascending = True)
    a = a.reset_index(drop=True)    
    a.rename(columns = {'2020 년_x': n_tido, '2020 년_y': n_sido}, inplace = True)
    sum_t = a[n_tido].sum()
    sum_r = a[n_sido].sum()
    a.index += 1
    a.loc[0] = ["계", sum_t, sum_r]
    a= a.sort_index(ascending = False)
    a[n_tido + "(%)"] = round(a[n_tido] / sum_t *100, 1)
    a[n_sido + "(%)"] = round(a[n_sido] / sum_r *100, 1)

    # 천단위 구분기호 전국과 keyR 적용
    # (참고) 구분기호 없애기: df["전국"] = df["전국"].str.replace(',', '').astype(int)  ## 1,195,951 -> 1195951로 바뀜
    a[n_tido] = a.apply(lambda x: "{:,}".format(x[n_tido]), axis=1) ### 천단위 구분기호
    a[n_sido] = a.apply(lambda x: "{:,}".format(x[n_sido]), axis=1) ### 천단위 구분기호

    a = a[[n_sido, n_sido + "(%)", "산업대", n_tido + "(%)", n_tido]] ## 열 순서

    listbox_sanup = ttk.Treeview(sanup_frame, column=("c1", "c2", "c3", "c4", "c5"), show = 'headings', height = 17)

    list_up = [n_sido + "사업체수(개)", n_sido + "구성비(%)", "산업대분류", n_tido + "구성비(%)", n_tido + "사업체수(개)"]
    width_num = [40, 20, 60, 20, 40]
    for i in range(5):
        listbox_sanup.column("# " + str(i+1), anchor=CENTER, width = width_num[i])
        listbox_sanup.heading("# " + str(i+1), text = list_up[i])  

    # listbox_sanup.column("# 1", anchor=CENTER, width = 35)
    # listbox_sanup.heading("# 1", text = keyR + "사업체수(개)")
    # listbox_sanup.column("# 2", anchor=CENTER, width = 10)
    # listbox_sanup.heading("# 2", text = keyR + "구성비(%)")
    # listbox_sanup.column("# 3", anchor=CENTER, width = 50)
    # listbox_sanup.heading("# 3", text = "산업대분류")
    # listbox_sanup.column("# 4", anchor=CENTER, width = 10)
    # listbox_sanup.heading("# 4", text = "전국구성비(%)")
    # listbox_sanup.column("# 5", anchor=CENTER, width = 35)
    # listbox_sanup.heading("# 5", text = "전국사업체수(개)")
    
    for i in a.index:
        listbox_sanup.insert('', 'end', text = i, values = (a.iloc[i, 0], a.iloc[i, 1], a.iloc[i, 2], a.iloc[i, 3], a.iloc[i, 4]))
    for i in [0,1,3,4]:
        listbox_sanup.column(i, anchor= 'e') ### Treeview 내부 '구분'항목은 왼쪽 정렬
    listbox_sanup.grid(row = 3, column = 0, ipadx = 20, ipady = 20, padx = 30, pady = 2, sticky=N+E+W+S)

    
    ### 엑셀 바로 실행 및 자료 보기: 노동실태현황 사업체수_c1_dt
    global c1_dt
    c1_dt = a.copy()
    c1_dt = c1_dt.sort_index(ascending = True)
    downbt_c1_dt = Button(sanup_frame, text = "xls", bg = "mistyrose", width = 2, height = 1, relief = "flat", overrelief = "raised", command = c1_dtdown)
    downbt_c1_dt.grid(row = 4, column = 0, ipadx = 0, ipady = 0, padx = 30, pady = 0, sticky = 'nw')      

    global sanup_label1, sanup_label2, sanup_label3
    sanup_label1 = Label(sanup_frame, \
        relief = "flat", text = "< 산업별 사업체수 및 구성비 >", \
        font = ("arial", 13, "bold"), padx = 20, pady = 0, anchor = "n", bg = 'white', fg = 'black')
    sanup_label1.grid(row = 1, column = 0, ipadx = 0, ipady = 0, padx = 0, pady = 0, sticky = "nsew")

    sanup_label2 = Label(sanup_frame, \
        relief = "flat", text = "< 산업별 종사자수 및 구성비 >", \
        font = ("arial", 13, "bold"), padx = 20, pady = 0, anchor = "n", bg = 'white', fg = 'black')
    sanup_label2.grid(row = 1, column = 1, ipadx = 0, ipady = 0, padx = 0, pady = 0, sticky = "nsew")

    sanup_label3 = Label(sanup_frame, \
        relief = "flat", text = "   ※ 자료: 고용노동부 사업체노동실태현황(2020년말 기준)", \
        font = ("arial", 9, "normal"), padx = 20, pady = 0, anchor = "w", bg = 'white', fg = 'black')
    sanup_label3.grid(row = 5, column = 0, ipadx = 0, ipady = 0, padx = 0, pady = 0, sticky = "w")


######################### ★★★★★ 종사자 차트 및 treeview  ####################################
def worker_bar():
    ## csv 파일 열기
    df = pd.read_csv(resource_path('data/3_2_db_company.csv'), encoding = 'cp949')
    
    ### 조건에 맞는 자료 추출 및 머지
    x = df.loc[(df['시도'] == tido) & (df['행정구역'] == tigu) & (df['산업대'] != "전체"), ['산업대', '2020 년']]
    y = df.loc[(df['시도'] == sido) & (df['행정구역'] == sigu) & (df['산업대'] != "전체"), ['산업대', '2020 년']]

    a = pd.merge(x,y, on = '산업대')
    a = a.sort_values('산업대', ascending = False)
    # a = a.reset_index(drop=True)

    ## 바차트 그리기 ##
    fig = Figure(figsize = (5.19, 3.5), dpi = 100)
    ax = fig.add_subplot(1,2,1)
    ax.barh(a['산업대'], a["2020 년_y"], color='mediumOrchid')
    # fig.patch.set_visible(False)
    # ax.axis('off')
    ax.grid(False)
    ax.tick_params(labelleft=False, labelright=True, pad = 0)  
    ax.spines['bottom'].set_visible(True)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    # ax.spines['left'].set_visible(False)
    ax.tick_params(axis = 'y', length = 30, right = False, left = False, bottom = False) ### 눈금선 없애기
    # ax.set_xlabel("명", fontweight ='bold', fontsize=13, loc="left")  ## x축 라벨 별도로 표시할 때
    fig.tight_layout()  ## 이미지 타이트하게

    ## 종사자수 두번째 : 비교지역 바차트 그리기 ##
    ax = fig.add_subplot(1,2,2)
    fig.suptitle("     - " + n_sido + " 종사자수 →                                            ← " + n_tido + " 종사자수-", fontsize = 11, x = 0.5, y = 0.95)
    ax.barh(a['산업대'], -a["2020 년_x"], color='mediumOrchid')
    ax.spines['bottom'].set_visible(True)  # 테두리 여부
    ax.spines['top'].set_visible(False)  # 테두리 여부
    # ax.spines['right'].set_visible(False)  # 테두리 여부
    ax.spines['left'].set_visible(False)  # 테두리 여부
    ax.tick_params(right = False, left = False, bottom = False, labelleft=False, labelright=False, pad = 0) ### 눈금선 없애기, 라벨위치여부
    ax.ticklabel_format(style = 'plain', axis = 'x') # 오류표시가 나타나지 않게 함  
    fig.tight_layout()  ## 이미지 타이트하게

    canvas = FigureCanvasTkAgg(fig, master = sanup_frame)
    canvas.draw()
    canvas.get_tk_widget().grid(row = 2, column = 1, ipadx = 70, ipady = 0, padx = 30, pady = 2)

    ### 차트 편집_c2_plot
    downbt_c2_plot = Button(sanup_frame, text = "편집/저장", bg = "mistyrose", width = 6, relief = "flat", overrelief = "raised", command = c2_plot)
    downbt_c2_plot.grid(row = 2, column = 1, ipadx = 6, ipady = 0, padx = 30, pady = 2, sticky = 'nw')


    ############# Treeview. 종사자수 Treeview 구문 ###############################
    a = a.sort_values('산업대', ascending = True)
    a = a.reset_index(drop=True)    
    a.rename(columns = {'2020 년_x': n_tido, '2020 년_y': n_sido}, inplace = True)
    sum_t = a[n_tido].sum()
    sum_r = a[n_sido].sum()
    a.index += 1
    a.loc[0] = ["계", sum_t, sum_r]
    a= a.sort_index(ascending = False)
    a[n_tido + "(%)"] = round(a[n_tido] / sum_t *100, 1)
    a[n_sido + "(%)"] = round(a[n_sido] / sum_r *100, 1)

    # 천단위 구분기호 전국과 keyR 적용
    # (참고) 구분기호 없애기: df["전국"] = df["전국"].str.replace(',', '').astype(int)  ## 1,195,951 -> 1195951로 바뀜
    a[n_tido] = a.apply(lambda x: "{:,}".format(x[n_tido]), axis=1) ### 천단위 구분기호
    a[n_sido] = a.apply(lambda x: "{:,}".format(x[n_sido]), axis=1) ### 천단위 구분기호

    a = a[[n_sido, n_sido + "(%)", "산업대", n_tido + "(%)", n_tido]] ## 열 정렬

    ### Treeview_종사자수 생성
    listbox_sanup = ttk.Treeview(sanup_frame, column=("c1", "c2", "c3", "c4", "c5"), show = 'headings', height = 17)
    
    list_workerup = [n_sido + "종사자수(명)", n_sido + "구성비(%)", "산업대분류", n_tido + "(%)", n_tido + "종사자수(명)"]
    width_num = [40, 20, 60, 20, 40]
    for i in range(5):
        listbox_sanup.column("# " + str(i+1), anchor=CENTER, width = width_num[i])
        listbox_sanup.heading("# " + str(i+1), text = list_workerup[i])  

    # listbox_sanup.column("# 1", anchor=CENTER, width = 35)
    # listbox_sanup.heading("# 1", text = keyR + "종사자수(명)")
    # listbox_sanup.column("# 2", anchor=CENTER, width = 10)
    # listbox_sanup.heading("# 2", text = keyR + "구성비(%)")
    # listbox_sanup.column("# 3", anchor=CENTER, width = 50)
    # listbox_sanup.heading("# 3", text = "산업대분류")
    # listbox_sanup.column("# 4", anchor=CENTER, width = 10)
    # listbox_sanup.heading("# 4", text = "전국구성비(%)")
    # listbox_sanup.column("# 5", anchor=CENTER, width = 35)
    # listbox_sanup.heading("# 5", text = "전국종사자수(명)")
    
    for i in a.index:
        listbox_sanup.insert('', 'end', text = i, values = (a.iloc[i, 0], a.iloc[i, 1], a.iloc[i, 2], a.iloc[i, 3], a.iloc[i, 4]))
    for i in [0,1,3,4]:
        listbox_sanup.column(i, anchor= 'e') ### Treeview 내부 '숫자'항목은 오른쪽 정렬
    listbox_sanup.grid(row = 3, column = 1, ipadx = 50, ipady = 20, padx = 30, pady = 2, sticky=N+E+W+S)

    ### 엑셀 바로 실행 및 자료 보기: 노동실태현황 종사자비중_c2_dt
    global c2_dt
    c2_dt = a.copy()
    c2_dt = c2_dt.sort_index(ascending = True)
    downbt_c2_dt = Button(sanup_frame, text = "xls", bg = "mistyrose", width = 2, height = 1, relief = "flat", overrelief = "raised", command = c2_dtdown)
    downbt_c2_dt.grid(row = 4, column = 1, ipadx = 0, ipady = 0, padx = 30, pady = 0, sticky = 'ne')      


######################### ★★★★★ 4_1. 월별 구인구직 차트 및 treeview  ####################################
def guin_plot():
    ## csv 파일 열기
    df = pd.read_csv(resource_path('data/4_1_db_guin.csv'), encoding = 'cp949')

    ### 조건에 맞는 자료 추출
    z = df.loc[(df['시도'] == tido) & (df['행정구역'] == tigu), ['기준년월', '취업건수', '구인배율', '월']]
    z = z.rename(columns = {'구인배율' : n_tido, '취업건수' : n_tido + '취업'})
    a = df.loc[(df['시도'] == sido) & (df['행정구역'] == sigu), ['기준년월', '취업건수', '구인배율', '월']]
    a = pd.merge(a, z, on = '기준년월')
    a = a.rename(columns = {'구인배율' : n_sido})
    a['연월'] = a['월_x']
    # for i in [0,7,19,31,36]:   ### Treeview 간격 변형 용도
    #     a.loc[i, '연월'] = a.loc[i,'기준년월']
    # a.loc[0, '연월'] = a.loc[0, '기준년월']
    # a.loc[12, '연월'] = a.loc[12, '기준년월']
    # a.loc[24, '연월'] = a.loc[24, '기준년월']
    # a.loc[36, '연월'] = a.loc[35, '기준년월']
    x = a['기준년월']
    y = round(a[n_sido], 2)
    z = round(a[n_tido], 2)
    maxmin = df.loc[(((df['시도'] == tido) & (df['행정구역'] == tigu)) | ((df['시도'] == sido) & (df['행정구역'] == sigu))), ['구인배율']]
    max_y = maxmin['구인배율'].max() + 0.1  ## y축 최대값
    if maxmin['구인배율'].max() - maxmin['구인배율'].min() > 1:
        min_y = maxmin['구인배율'].min() - 2.0  ## y축 최소값
    elif maxmin['구인배율'].max() - maxmin['구인배율'].min() > 0.5:
        min_y = maxmin['구인배율'].min() - 1.4  ## y축 최소값
    else:
        min_y = maxmin['구인배율'].min() - 0.7
    
    ## 꺾은선차트 그리기 ##
    fig = Figure(figsize = (13.7, 3.8), dpi = 100)
    ax = fig.add_subplot(111)
    fig.suptitle("[ 구인배율(꺾은선)과 취업건수(막대) ]", fontsize = 11, y = 0.95)
    # ysmoothed = gaussian_filter1d(y, sigma=0.5)  # 완만한선, 실패
    ax.plot(x, y, color = 'red', linewidth = 2, marker = "o", markeredgewidth = 1, \
        markeredgecolor = "red", markerfacecolor = 'yellow', markersize = 6, label = n_sido + "구인배율", markevery = 1)
    ax.plot(x, z, color = 'gray', linewidth = 1, marker = "o", markeredgewidth = 1, \
        markeredgecolor = 'gray', markerfacecolor = backcolor, markersize = 3, label = n_tido + "구인배율", markevery = 1, linestyle = '--')
    ax.set_ylim(min_y, max_y)
    ax.set_ylabel('구인배율(p)')
    ax.spines['top'].set_visible(False)  # 테두리 여부
    ax.spines['right'].set_visible(False)  # 테두리 여부
    ax.tick_params(right = False, left = False, bottom = False, labelleft=True, labelright=False, pad = 0) ### 눈금선 없애기, 라벨위치여부
    for idx, txt in enumerate(y):
        ax.text(x[idx], y[idx] + 0.05, txt, ha='center')

    ## 취업건수 바차트 추가 그리기 ##
    ax2 = ax.twinx()
    ax2.set_ylabel('취업건수(건)')
    amax = a["취업건수"].max() * 2
    ax2.set_ylim(0, amax)
    ax2.bar(x, a["취업건수"], color = '#ff812d', label = n_sido + "취업건수")
    ax2.spines['top'].set_visible(False)  # 테두리 여부
    fig.legend(loc = (0.05, 0.85), fontsize = 9) # 범례 표시
    fig.autofmt_xdate(rotation=45)  ## x축 레이블 45도각도 눕히기
    # a["취업건수"] = a.apply(lambda x: "{:,}".format(x["취업건수"]), axis=1)
    for idx, txt in enumerate(a["취업건수"]):
        ax2.text(idx, a["취업건수"][idx] + 12, txt, ha='center')
    fig.tight_layout()  ## 이미지 타이트하게
 
    global canvas
    canvas = FigureCanvasTkAgg(fig, master = guin_frame)
    canvas.draw()
    canvas.get_tk_widget().grid(row = 2, column = 0, columnspan = 7, ipadx = 30, ipady = 20, padx = 5, pady = 2)

    ### 차트 편집_d1_plot
    downbt_d1_plot = Button(guin_frame, text = "편집/저장", bg = "mistyrose", width = 6, relief = "flat", overrelief = "raised", command = d1_plot)
    downbt_d1_plot.grid(row = 2, column = 0, ipadx = 6, ipady = 0, padx = 5, pady = 2, sticky = 'nw')    


    ##### ★★★★★★ 구인구직 첫번째 Treeview 그리기, dataframe 'a' 계속 사용
    a['차이'] = round(a[n_sido] - a[n_tido], 2)
    a[n_sido] =  round(a[n_sido], 2)
    a[n_tido] =  round(a[n_tido], 2)
    a = a[["연월", n_sido, "차이", n_tido]]
    a = a.rename(columns = ({n_sido : "○ " + n_sido, "차이" : n_sido + "-" + n_tido +'(p)', n_tido : "○ " + n_tido}))
    a = a.set_index('연월').T ### 데이터 전치
    a = a.reset_index(drop = False) ### 전치 후 '항목 열'을 'index'로 계속 가지고 감
    a = a.rename(columns = {'index' : '구분(구인배율)'})
    cols = list(a.columns)
    column_append_cols = ["c"+ str(i+1) for i in range(len(cols))]  ###★★★★★★ 리스트 누적해서 열리스트 한정(대전광역시 자치구 때문에 필요)
    
    treeview_guinplot1 = ttk.Treeview(guin_frame, column= column_append_cols, show = 'headings', height = 1)
    
    ### Treeview heading
    for i, j in enumerate(cols):
        treeview_guinplot1.column("# " + str(i+1), anchor=CENTER, width = 20)
        treeview_guinplot1.heading("# " + str(i+1), text = j)  
    for i in range(len(cols)): # 문자열길이 4보다 큰 거(1월)은 좀 넓게 추가 정의
        if len(cols[i]) > 4:
            treeview_guinplot1.column("c"+ str(i+1), width = 35)    
    treeview_guinplot1.column("c1", width = 83)  # 지역구분은 좀 넓게 추가 정의
    

    ### Treeview data insert
    for i in a.index:
        treeview_guinplot1.insert('', 'end', text = i, values = (list(a.values[i])))

    treeview_guinplot1.column(0, anchor= 'w') ### Treeview 내부 '구분'항목은 왼쪽 정렬
    treeview_guinplot1.grid(row = 3, column = 0, columnspan = 7, ipadx = 30, ipady = 20, padx = 5, pady = 2, sticky=N+E+W+S) 

    guin_label = Label(guin_frame, \
        relief = "flat", text = "    (용어해설) '구인배율'은 노동의 수요와 공급을 나타내는 지표로 일자리수를 구직자수로 나누어 구합니다." + "\n" + \
            "       예를 들어 구인배율 0.8은 구직자 10명에 대해 일자리는 8개가 있음을 의미합니다.", \
        font = ("arial", 9, "normal"), padx = 20, pady = 0, anchor = "w", bg = 'white', fg = 'black')
    guin_label.grid(row = 4, column = 0, columnspan = 7, ipadx = 0, ipady = 0, padx = 0, pady = 0, sticky = "nsew")   

    ### 엑셀 바로 실행 및 자료 보기: 구인배율추이_d1_dt
    global d1_dt
    d1_dt = a.copy()
    downbt_d1_dt = Button(guin_frame, text = "xls", bg = "mistyrose", width = 2, height = 1, relief = "flat", overrelief = "raised", command = d1_dtdown)
    downbt_d1_dt.grid(row = 4, column = 0, ipadx = 0, ipady = 0, padx = 6, pady = 0, sticky = 'nw')    


######################### ★★★★★ 4_2. 연도별 구인구직 차트 및 treeview  ####################################
def guinjob_plot():
    ## csv 파일 열기
    df = pd.read_csv(resource_path('data/4_2_db_guin.csv'), encoding = 'cp949')

    a = df.loc[(df['시도'] == sido) & (df['행정구역'] == sigu), ['기준년', '신규구인인원', '신규구직건수']]
    a = a.groupby(['기준년'])[['신규구인인원', '신규구직건수']].sum()
    a = a.T
    a['증감'] = a['22년'] - a['21년']
    a['증감률'] = round((a['증감'] / a['21년']) * 100,1)
    a['행정구역'] = n_sido
    a = a.reset_index(drop = False)
    a = a[["행정구역", "index", "21년", "22년", "증감", "증감률"]]
    a.columns = ["행정구역", "구분", "'21년", "'22년", "증감", "증감률"]
    
    b = df.loc[(df['시도'] == tido) & (df['행정구역'] == tigu), ['기준년', '신규구인인원', '신규구직건수']]
    b = b.groupby(['기준년'])[['신규구인인원', '신규구직건수']].sum()
    b = b.T
    b['증감'] = b['22년'] - b['21년']
    b['증감률'] = round((b['증감'] / b['21년']) * 100,1)
    b['행정구역'] = n_tido
    b = b.reset_index(drop = False)
    b = b[["행정구역", "index", "21년", "22년", "증감", "증감률"]]
    b.columns = ["행정구역", "구분", "'21년", "'22년", "증감", "증감률"]
    
    z = pd.concat([a, b], ignore_index = True)
    
    aa = z.iloc[0,5]  ### 신규구인인원 증감률
    if aa > 0:
        c = "+" + str(aa) + "%"
    elif aa < 0:
        c = "" + str(aa) + "%"
    else:
        c = "-(변동없음)"
    ad = z.iloc[1,5]  ### 신규구직건수 증감률
    if ad > 0:
        e = "+" + str(ad) + "%"
    elif ad < 0:
        e = "" + str(ad) + "%"
    else:
        e = "(변동없음)"


    guinjob_label0 = Label(guin_frame, \
        relief = "flat", text = " " + "\n" + " ", \
        font = ("arial", 10, "bold"), padx = 0, pady = 0, bg = 'white', fg = 'darkviolet', anchor = 'sw')
    guinjob_label0.grid(row = 8, rowspan = 3, column = 0, ipadx = 20, ipady = 30, padx = 0, pady = 0)
    
    guinjob_label = Label(guin_frame, \
        relief = "flat", text = "○ 신규구인인원과 신규구직건수('22년)", \
        font = ("arial", 18, "bold"), padx = 0, pady = 0, bg = 'white', fg = 'darkviolet', anchor = 'w')
    guinjob_label.grid(row = 5, columnspan = 5, column = 1, ipadx = 0, ipady = 30, padx = 0, pady = 0)

    guinjob_label1 = Label(guin_frame, \
        relief = "flat", text = "신규구인인원", \
        font = ("arial", 15, "bold"), padx = 0, pady = 0, bg = 'white', fg = 'midnightblue', anchor = 'w')
    guinjob_label1.grid(row = 6, column = 1, ipadx = 0, ipady = 0, padx = 0, pady = 0)

    dataguin_label2 = Label(guin_frame, \
        relief = "flat", text = format(z.iloc[0,3], ',') + "명", \
        font = ("arial", 20, "bold"), padx = 0, pady = 0, bg = 'white', fg = 'dodgerblue', anchor = 'e')
    dataguin_label2.grid(row = 7, column = 1, ipadx = 0, ipady = 0, padx = 0, pady = 0)

    dataguin_label3 = Label(guin_frame, \
        relief = "flat", text = c, \
        font = ("arial", 20, "bold"), padx = 0, pady = 0, bg = 'dodgerblue', fg = 'white')
    dataguin_label3.grid(row = 7, column = 2, ipadx = 0, ipady = 0, padx = 0, pady = 0, sticky = "ew")

    guin_label4 = Label(guin_frame, fg = 'darkorange', bg = 'white', \
        relief = "flat", text = "|" + "\n" + "|" + "\n" + "|" + "\n" + "|", font = ("arial", 10, "bold"), padx = 0)
    guin_label4.grid(row = 6, rowspan = 2, column = 3, ipadx = 10, ipady = 0, padx = 0, pady = 0)

    guinjob_label5 = Label(guin_frame, \
        relief = "flat", text = "신규구직건수", \
        font = ("arial", 15, "bold"), padx = 0, pady = 0, bg = 'white', fg = 'midnightblue', anchor = 'w')
    guinjob_label5.grid(row = 6, column = 4, ipadx = 0, ipady = 0, padx = 0, pady = 0)

    dataguin_label6 = Label(guin_frame, \
        relief = "flat", text = format(z.iloc[1,3], ',') + "건", \
        font = ("arial", 20, "bold"), padx = 0, pady = 0, bg = 'white', fg = 'dodgerblue', anchor = 'e')
    dataguin_label6.grid(row = 7, column = 4, ipadx = 0, ipady = 0, padx = 0, pady = 0)

    dataguin_label7 = Label(guin_frame, \
        relief = "flat", text = e, \
        font = ("arial", 20, "bold"), padx = 0, pady = 0, bg = 'dodgerblue', fg = 'white')
    dataguin_label7.grid(row = 7, column = 5, ipadx = 0, ipady = 0, padx = 0, pady = 0, sticky = "ew")


    ####### data 정리 및 Treeview 연결
    z = z.reset_index(drop = True)
    cols = list(z.columns)
    column_append_cols = ["c"+ str(i+1) for i in range(len(cols))]  ###★★★★★★ 리스트 누적해서 열리스트 한정
    
    treeview_guinplot1 = ttk.Treeview(guin_frame, column= column_append_cols, show = 'headings', height = 4)

    for i in ["'21년", "'22년", "증감"]:
        z[i] = z.apply(lambda x: "{:,}".format(x[i]), axis=1) ### 천단위 구분기호

    ### Treeview heading
    for i, j in enumerate(cols):
        treeview_guinplot1.column("# " + str(i+1), anchor=CENTER, width = 10)
        treeview_guinplot1.heading("# " + str(i+1), text = j)  
    for i in [2,3,4]:
        treeview_guinplot1.column(i, anchor= 'e') ### Treeview 내부 '숫자'항목은 오른쪽 정렬

    ### Treeview data insert
    for i in z.index:
        treeview_guinplot1.insert('', 'end', text = i, values = (list(z.values[i])))

    treeview_guinplot1.grid(row = 8, rowspan = 3, column = 1, columnspan = 5, ipadx = 0, ipady = 0, padx = 0, pady = 2, sticky=N+E+W+S) 

    ### 엑셀 바로 실행 및 자료 보기: 신규구인구직비교_d2_dt
    global d2_dt
    d2_dt = z.copy()
    downbt_d2_dt = Button(guin_frame, text = "xls", bg = "mistyrose", width = 2, height = 1, relief = "flat", overrelief = "raised", command = d2_dtdown)
    downbt_d2_dt.grid(row = 11, column = 1, ipadx = 0, ipady = 0, padx = 1, pady = 0, sticky = 'nw')    

    guin_label = Label(guin_frame, \
        relief = "flat", text =  "※ 자료: 한국고용정보원 '워크넷'", \
        font = ("arial", 9, "normal"), padx = 0, pady = 0, anchor = "w", bg = 'white', fg = 'black')
    guin_label.grid(row = 13, column = 0, columnspan = 2, ipadx = 10, ipady = 0, padx = 0, pady = 10, sticky = "nsew")

############### ★★★★★ 4_2 추가 미스매치 데이터 및 바차트 #############################################
def guinmatch_plot():
    ## csv 파일 열기
    df = pd.read_csv(resource_path('data/4_2_db_guin.csv'), encoding = 'cp949')

    guinmatch_label = Label(guin_frame, \
        relief = "flat", text =  "○ 직종별 구인구직 미스매치('22년)", \
        font = ("arial", 18, "bold"), padx = 20, pady = 0, anchor = "w", bg = 'white', fg = 'darkviolet')
    guinmatch_label.grid(row = 5, column = 6, ipadx = 50, ipady = 30, padx = 60, pady = 0, sticky = "nsew")

    ############### 미스매치 데이터 편집 #############################################
    a = df.loc[((df['시도'] == sido) & (df['행정구역'] == sigu) & (df['기준년'] == "22년")), ["직종_중분류", "신규구인인원", "신규구직건수", "구인배율"]]
    filt = (a['신규구인인원'] > 10) & (a['신규구직건수'] > 10)
    a = a[filt]
    a = a.sort_values('구인배율', ascending = True)
    a = a.reset_index(drop = True)
    a['구인배율'] = round(a['구인배율'],2)
    a1 = a.iloc[0:5,[0,3]]
    a1 = a1.sort_values('구인배율', ascending = False)  # a1 완료
    a2 = a.sort_values('구인배율', ascending = False)
    a2 = a2.reset_index(drop = True)
    a2 = a2.iloc[0:5, [0,3]]
    a2 = a2.sort_values('구인배율', ascending = True)
    a = pd.concat([a1, a2], ignore_index = True)

    ############### 미스매치 바차트 그리기 #############################################
    fig = Figure(figsize = (6, 2.3), dpi = 100)
    ax = fig.add_subplot()
    ax.barh(a['직종_중분류'], a["구인배율"], color='lightpink')
    # fig.patch.set_visible(False)
    # ax.axis('off')
    ax.grid(False)
    ax.tick_params(labelleft=True, labelright=False, pad = 0)  
    ax.spines['bottom'].set_visible(True)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    # ax.spines['left'].set_visible(False)
    ax.tick_params(axis = 'y', right = False, left = False, bottom = False) ### 눈금선 없애기
    # ax.set_xlabel("명", fontweight ='bold', fontsize=13, loc="left")  ## x축 라벨 별도로 표시할 때
    fig.tight_layout()  ## 이미지 타이트하게

    canvas = FigureCanvasTkAgg(fig, master = guin_frame)
    canvas.draw()
    canvas.get_tk_widget().grid(row = 6, rowspan = 5, column = 6, ipadx = 50, ipady = 0, padx = 80, pady = 0, sticky = "nsew")

    ### 차트 편집_d2_plot
    downbt_d2_plot = Button(guin_frame, text = "편집/저장", bg = "mistyrose", width = 6, relief = "flat", overrelief = "raised", command = d2_plot)
    downbt_d2_plot.grid(row = 6, column = 6, ipadx = 6, ipady = 0, padx = 80, pady = 0, sticky = 'nw')    

    guinmatch_label = Label(guin_frame, relief = "flat", text =  "※ 막대가 긴 경우(구인배율 상위5개)는 구인자의 채용이 상대적으로 어려웠을 수 있고,", font = ("arial", 9, "bold"), padx = 40, pady = 0, anchor = "nw", bg = 'white', fg = 'indigo')
    guinmatch_label.grid(row = 11, column = 6, ipadx = 50, ipady = 0, padx = 40, pady = 0, sticky = "nsew")
    guinmatch_label = Label(guin_frame, relief = "flat", text =  "     막대가 짧은 경우(구인배율 하위5개)는 구직자의 일자리 찾기가 상대적으로 어려웠을 수 있었던 직종입니다.", font = ("arial", 9, "bold"), padx = 40, pady = 0, anchor = "nw", bg = 'white', fg = 'indigo')
    guinmatch_label.grid(row = 12, column = 6, ipadx = 50, ipady = 0, padx = 40, pady = 0, sticky = "nsew")


##############  ★★★★★★★    이미지 및 데이터테이블 다운로드 지원 ★★★★★★★★#############################################
def a1_plot(): # 인구 첫번째 꺾은선 차트
    ## csv 파일 열기
    df = pd.read_csv(resource_path('data/1_1_db_1st_popul_2022.csv'), encoding = 'cp949')
    x = df.loc[((df['시도'] == tido) & (df['행정구역'] == tigu) & (df['연월'] != "'16.12월")), ['연월', 'value']]
    y = df.loc[((df['시도'] == sido) & (df['행정구역'] == sigu) & (df['연월'] != "'16.12월")), ['연월', 'value']]
    a = pd.merge(x,y, on = '연월')
    a = a.rename(columns= {'연월': '연월', 'value_x': n_tido, 'value_y': n_sido})
    a = a.reset_index(drop=True)
    b = a.copy()
    b[n_tido] = round(b[n_tido] /10000, 2)
    b[n_sido] = round(b[n_sido] /10000, 2)
    ## 인구 꺾은선 차트 그리기 ##
    x = b['연월']
    y1 = b[n_tido]
    y2 = b[n_sido]
    max_y1 = y1.max() + 0.3  ## y축 최대값
    min_y1 = y1.min() - 0.3  ## y축 최소값
    max_y2 = y2.max() + 0.3  ## y축 최대값
    min_y2 = y2.min() - 0.3  ## y축 최소값
    fig_a1 = plt.figure(figsize = (6, 3.5), dpi = 100)
    ax1 = fig_a1.add_subplot()
    fig_a1.suptitle("[ '" + oreg + "' 인구 추이 ]", fontsize = 11, y = 0.95)
    ax1.set_ylabel(oreg + '(만명)')
    ax2 = ax1.twinx() # x축을 함께 사용
    ax1.plot(x, y2, color = 'red', linewidth = 2, marker = "o", markeredgewidth = 1, \
        markeredgecolor = "red", markerfacecolor = "yellow", markersize = 6, label = n_sido, markevery = 1)
    ax2.plot(x, y1, color = 'darkblue', linewidth = 1, marker = "D", markerfacecolor = chartcolor, \
        markersize = 3, label = n_tido, markevery = 1, linestyle = '--')
    ax1.spines['top'].set_visible(False)  # 테두리 여부
    ax1.set_ylim(min_y2, max_y2)
    # y2 차트 데이터레이블 표시
    for idx, txt in enumerate(y2):
        ax1.text(x[idx], y2[idx] + 0.05, txt, ha='center')
    ax2.set_ylabel(treg + ' (만명)')
    ax2.spines['top'].set_visible(False)  # 테두리 여부
    ax2.set_ylim(min_y1, max_y1)
    fig_a1.legend(loc = (0.18, 0.35), fontsize = 9, facecolor = backcolor) # 범례 표시
    fig_a1.tight_layout(pad=1, h_pad=None, w_pad= 0.9)  ## 이미지 타이트하게
    # plt.savefig('test.png') 
    fig_a1.show()    
def a2_plot(): # 인구 과거피라미드 차트, 415행
    df_p = pd.read_csv(resource_path('data/1_1_db_2st_past_popul_2022.csv'))    
    ### 조건에 맞는 자료 추출 및 머지
    x_p = df_p.loc[(df_p['시도'] == sido) & (df_p['행정구역'] == sigu) & (df_p['성별'] == "남자인구수[명]"), ['5세별', 'value']]
    y_p = df_p.loc[(df_p['시도'] == sido) & (df_p['행정구역'] == sigu) & (df_p['성별'] == "여자인구수[명]"), ['5세별', 'value']]
    past_dataframe = df_p.loc[(df_p['시도'] == sido) & (df_p['행정구역'] == sigu) & (df_p['성별'] == "남자인구수[명]"), ['5세별', '연월', 'value']]
    # past_dataframe['연월'] = past_dataframe['연월'].str.slice(start = 0, stop = 3) + "년말"
    global past
    past = past_dataframe.iloc[0, 1]
    a_p = pd.merge(x_p, y_p, on = '5세별')
    ## 바차트 그리기 ##
    fig_a2 = plt.figure(figsize = (4, 3.5), dpi = 100)
    ax = fig_a2.add_subplot()
    fig_a2.suptitle("<<  " + past + "  >>", fontsize = 11, x=0.56, y=0.95)
    ax.barh(a_p['5세별'], -a_p.value_x, label = "남(명)", color = 'forestgreen')
    ax.barh(a_p['5세별'], a_p.value_y, label = "여(명)", color = 'darkorange')
    ax.spines['top'].set_visible(False)  # 테두리 여부
    ax.spines['right'].set_visible(False)  # 테두리 여부
    ax.tick_params(right = False, left = False, bottom = False, labelleft=True, labelright=False, pad = 0) ### 눈금선 없애기, 라벨위치여부
    fig_a2.tight_layout()  ## 이미지 타이트하게
    fig_a2.legend(loc = (0.2, 0.71), fontsize = 9) # 범례 표시
    fig_a2.show()
def a3_plot(): # 인구 현재피라미드 차트, 415행
    df_p = pd.read_csv(resource_path('data/1_1_db_2st_now_popul_2022.csv'))
    ### 조건에 맞는 자료 추출 및 머지
    x_p = df_p.loc[(df_p['시도'] == sido) & (df_p['행정구역'] == sigu) & (df_p['성별'] == "남자인구수[명]"), ['5세별', 'value']]
    y_p = df_p.loc[(df_p['시도'] == sido) & (df_p['행정구역'] == sigu) & (df_p['성별'] == "여자인구수[명]"), ['5세별', 'value']]
    a_p = pd.merge(x_p, y_p, on = '5세별')
    ## 바차트 그리기 ##
    fig_a3 = plt.figure(figsize = (4, 3.5), dpi = 100)
    ax = fig_a3.add_subplot()
    fig_a3.suptitle("<<  " + present + "  >>", fontsize = 11, x=0.56, y=0.95)
    ax.barh(a_p['5세별'], -a_p.value_x, label = "남(명)", color = 'forestgreen')
    ax.barh(a_p['5세별'], a_p.value_y, label = "여(명)", color = 'darkorange')
    ax.spines['top'].set_visible(False)  # 테두리 여부
    ax.spines['left'].set_visible(False)  # 테두리 여부
    ax.tick_params(right = False, left = False, bottom = False, labelleft=False, labelright=True, pad = 0) ### 눈금선 없애기, 라벨위치여부
    fig_a3.tight_layout()  ## 이미지 타이트하게
    fig_a3.legend(loc = (0.6, 0.71), fontsize = 9) # 범례 표시
    fig_a3.show()
def a4_plot():
    df = pd.read_csv(resource_path('data/1_2_db_oldman_2022.csv'))
    df1 = df.loc[((df['시도'] == sido) & (df['행정구역'] == sigu) & (df['5세코드'] == 0))]
    df2 = df.loc[((df['시도'] == sido) & (df['행정구역'] == sigu) & (df['5세코드'] > 65))]
    df1 = df1[["연월", "value"]]
    df2 = df2.groupby(["연월"]).sum()
    df2 = df2.reset_index(drop= False)
    df2 = df2[['연월', 'value']]
    dfa = pd.merge(df1, df2, on = "연월")
    dfa["고령화율"] = 0.00
    dfa["고령화율"] = round(dfa['value_y'] / dfa['value_x'] *100, 2)
    dfa = dfa[["연월", "고령화율"]]
    df1 = df.loc[((df['시도'] == tido) & (df['행정구역'] == tigu) & (df['5세코드'] == 0))]
    df2 = df.loc[((df['시도'] == tido) & (df['행정구역'] == tigu) & (df['5세코드'] > 65))]
    df1 = df1[["연월", "value"]]
    df2 = df2.groupby(["연월"]).sum()
    df2 = df2.reset_index(drop= False)
    df2 = df2[['연월', 'value']]
    dfb = pd.merge(df1, df2, on = "연월")
    dfb["고령화율"] = 0.00
    dfb["고령화율"] = round(dfb['value_y'] / dfb['value_x'] *100, 2)
    dfb = dfb[["연월", "고령화율"]]
    at = pd.merge(dfa, dfb, on = "연월")
    at = at.rename(columns = {"고령화율_x" : n_sido, "고령화율_y" : n_tido})
    at["차이(%p)"] = round(at[n_sido] - at[n_tido], 2)
    at = at[["연월", n_sido, "차이(%p)", n_tido]]
    if n_sido == "청주시":
        a = at.iloc[2,1]
        b = at.iloc[6,1]
        text_title = "      << " + at.iloc[2,0] + " >>                →→                << " + at.iloc[6,0] + " >>                →→                << " + at.iloc[10,0] + " >>"
    else:
        a = at.iloc[0,1]
        b = at.iloc[5,1]
        text_title = "      << " + at.iloc[0,0] + " >>                →→                << " + at.iloc[5,0] + " >>                →→                << " + at.iloc[10,0] + " >>"
    c = at.iloc[10,1]
    colors = ['cornflowerblue', 'crimson']
    explode = [0.05] * 2
    labels = ['65세미만', '65세이상']
    wedgeprops = {'width' : 0.55, 'edgecolor' : 'w', 'linewidth' : 2}
    fig_a4 = plt.figure(figsize = (10, 3), dpi = 100)    
    fig_a4.suptitle(text_title, fontsize = 11, y =0.95)
    ax = fig_a4.add_subplot(131)
    ax.pie((100 - a, a), labels = labels, autopct = '%.1f%% ', textprops = {'fontsize': 10, 'color' : 'black', 'weight' : 'bold'}, pctdistance = 0.7, \
        labeldistance = 1.1, startangle = 90, counterclock = False, colors = colors, explode = explode, wedgeprops = wedgeprops)
    ax2 = fig_a4.add_subplot(132)
    ax2.pie((100 - b, b), autopct = '%.1f%% ', textprops = {'fontsize': 11, 'color' : 'black', 'weight' : 'bold'}, pctdistance = 0.7, \
        labeldistance = 1.1, startangle = 90, counterclock = False, colors = colors, explode = explode, wedgeprops = wedgeprops)
    ax3 = fig_a4.add_subplot(133)
    ax3.pie((100 - c, c), autopct = '%.1f%% ', textprops = {'fontsize': 12, 'color' : 'black', 'weight' : 'bold'}, pctdistance = 0.7, \
        labeldistance = 1.1, startangle = 90, counterclock = False, colors = colors, explode = explode, wedgeprops = wedgeprops)
    fig_a4.tight_layout()  ## 이미지 타이트하게
    fig_a4.show()
def b1_plot():
    ## excel 파일 열기
    df = pd.read_csv(resource_path('data/2_1_db_sex_emprate.csv'), encoding = 'cp949')
    ### 조건에 맞는 자료 추출
    z = df.loc[(df['시도'] == tido) & (df['행정구역'] == tigu) & (df['성별'] == "계"), ['연월', '15~64세 고용률 (%)']]
    z = z.rename(columns = {'15~64세 고용률 (%)' : n_tido})
    a = df.loc[(df['시도'] == sido) & (df['행정구역'] == sigu) & (df['성별'] == "계"), ['연월', '15~64세 고용률 (%)']]
    a = pd.merge(a, z, on = '연월')
    a = a.rename(columns = {'15~64세 고용률 (%)' : n_sido})
    x = a['연월']
    y = a[n_sido]
    z = a[n_tido]
    maxmin = df.loc[(((df['시도'] == sido) & (df['행정구역'] == sigu)) | ((df['시도'] == tido) & (df['행정구역'] == tigu))) & (df['성별'] == "계"), ['15~64세 고용률 (%)']]
    max_y = maxmin['15~64세 고용률 (%)'].max() + 0.9  ## y축 최대값
    min_y = maxmin['15~64세 고용률 (%)'].min() - 0.7  ## y축 최소값
    ## 꺾은선차트 그리기 ##
    fig_b1 = plt.figure(figsize = (6, 3), dpi = 100)
    ax = fig_b1.add_subplot()
    fig_b1.suptitle("[ " + n_sido + " 고용률(15~64세) 추이 ]", fontsize = 11, y = 0.95)
    ax.plot(x, y, color = 'red', linewidth = 2, marker = "o", markeredgewidth = 1, \
        markeredgecolor = "red", markerfacecolor = 'yellow', markersize = 6, label = n_sido, markevery = 1)
    ax.plot(x, z, color = 'darkblue', linewidth = 1, marker = "o", markeredgewidth = 1, \
        markeredgecolor = 'darkblue', markerfacecolor = backcolor, markersize = 3, label = n_tido, markevery = 1, linestyle = '--')
    ax.set_ylim(min_y, max_y)
    ax.spines['top'].set_visible(False)  # 테두리 여부
    ax.spines['right'].set_visible(False)  # 테두리 여부
    ax.tick_params(right = False, left = False, bottom = False, labelleft=True, labelright=False, pad = 0) ### 눈금선 없애기, 라벨위치여부
    fig_b1.legend(loc = (0.13, 0.73), fontsize = 9) # 범례 표시
    # y1 차트 데이터레이블 표시
    for idx, txt in enumerate(y):
        ax.text(x[idx], y[idx] + 0.07, txt, ha='center')
    fig_b1.show()
def b2_plot():
    ## excel 파일 열기
    df = pd.read_csv(resource_path('data/2_1_db_sex_emprate.csv'), encoding = 'cp949')
    month_emp = df.sort_values('연월', ascending = False)
    month_emp = month_emp.iloc[0, 3]  ## 가져올 자료의 기준시점 설정
    ### 조건에 맞는 자료 추출 및 사전작업
    male = df.loc[(((df['시도'] == sido) & (df['행정구역'] == sigu)) | ((df['시도'] == tido) & (df['행정구역'] == tigu))) & \
        (df['성별'] == "남자") & (df['연월'] == month_emp), ['시도', '행정구역', '성별','연월', '15~64세 고용률 (%)']]
    female = df.loc[(((df['시도'] == sido) & (df['행정구역'] == sigu)) | ((df['시도'] == tido) & (df['행정구역'] == tigu))) & \
        (df['성별'] == "여자") & (df['연월'] == month_emp), ['시도', '행정구역', '성별','연월', '15~64세 고용률 (%)']]
    male['구분'] = male['행정구역']
    male.loc[male['행정구역'] == "전체", '구분'] = male['시도']
    male = male[['구분', '15~64세 고용률 (%)']]
    female['구분'] = female['행정구역']
    female.loc[female['행정구역'] == "전체", '구분'] = female['시도']
    female = female[['구분', '15~64세 고용률 (%)']]
    a = pd.merge(male,female, on = "구분")
    a = a.rename(columns = {"15~64세 고용률 (%)_x" : "남자", "15~64세 고용률 (%)_y" : "여자"})
    a = a[["구분", "남자", "여자"]].set_index("구분").T
    a = a.reset_index(drop = False)
    a = a[['index', n_sido, n_tido]].T
    a = a.reset_index(drop = False)
    a = a.rename(columns = a.iloc[0])
    a = a.drop(a.index[0])
    m = a['남자'].reset_index(drop=True)
    f = a['여자'].reset_index(drop=True)
    t = pd.concat([m, f], ignore_index = True)
    max_t = t.max() + 19
    min_t = t.min() - 19
    reg = a['index']  # 'x축 명' 표시용
    x_axis = np.arange(len(reg))
    ## 바차트 그리기 ##
    fig_b2 = plt.figure(figsize = (3.5, 3.5), dpi = 100)
    ax = fig_b2.add_subplot()
    fig_b2.suptitle("[ " + n_sido + " 성별고용률(" + month_emp + ")]", fontsize = 11, y = 0.95)
    ax.bar(x_axis -0.2, m, width = 0.4, label = '남(%)', color = 'forestgreen')
    ax.bar(x_axis +0.2, f, width=0.4, label = '여(%)', color = 'darkorange')
    ax.set_ylim(min_t, max_t)
    ax.set_xticks(x_axis, reg)
    fig_b2.legend(loc = (0.13, 0.72), fontsize = 9) # 범례 표시
    ax.spines['bottom'].set_visible(True)  # 테두리 여부
    ax.spines['top'].set_visible(False)  # 테두리 여부
    ax.spines['right'].set_visible(False)  # 테두리 여부
    ax.spines['left'].set_visible(False)  # 테두리 여부
    for idx, txt in enumerate(m):
        ax.text(x_axis[idx]-0.2, m[idx]+1.5, txt, ha='center')    
    for idx, txt in enumerate(f):
        ax.text(x_axis[idx]+0.2, f[idx]+1.5, txt, ha='center')
    ax.tick_params(right = False, left = False, bottom = False, labelleft=False, labelright=False, pad = 0) ### 눈금선 없애기, 라벨위치여부
    # male과 female 차트 데이터레이블 표시    
    fig_b2.show()
def b3_plot():
    df = pd.read_csv(resource_path('data/2_2_db_age_emprate_adapt.csv'), encoding = 'cp949')
    ### 조건에 맞는 자료 추출 및 사전작업
    if gungucbox.get() == '전체' and Tgungucbox.get() == '전체':
        ### 청년 꺾은선 차트 용 dataframe 'c'
        a = df.loc[(df['시도'] == sido) & (df['행정구역'] == sigu), ['시점', '고용률 (%)']]
        b = df.loc[(df['시도'] == tido) & (df['행정구역'] == tigu), ['시점', '고용률 (%)']]
        c = pd.merge(a,b, on = "시점")
        c = c.sort_values('시점', ascending = True)
        c = c.reset_index(drop = True)
        a = c.copy()
        c.rename(columns = {"고용률 (%)_x" : n_sido, "고용률 (%)_y" : n_tido}, inplace = True)
        x = c['시점']
        y = c[n_sido]
        z = c[n_tido]
        maxmin = df.loc[((df['시도'] == sido) & (df['행정구역'] == sigu)) | ((df['시도'] == tido) & (df['행정구역'] == tigu)), ['고용률 (%)']]
        max_y = maxmin['고용률 (%)'].max() + 0.9  ## y축 최대값
        min_y = maxmin['고용률 (%)'].min() - 0.7  ## y축 최소값
        fig_b3 = plt.figure(figsize = (6, 3), dpi = 100)
        ax = fig_b3.add_subplot()
        fig_b3.suptitle("[ " + n_sido + " 청년고용률(15-29세) ]", fontsize = 11, y = 0.95)
        ax.plot(x, y, color = 'red', linewidth = 2, marker = "o", markeredgewidth = 1, \
            markeredgecolor = "firebrick", markerfacecolor = "yellow", markersize = 6, label = n_sido, markevery = 1)
        ax.plot(x, z, color = 'darkblue', linewidth = 1, marker = "o", markeredgewidth = 1, \
            markeredgecolor = "darkblue", markerfacecolor = "darkblue", markersize = 3, label = n_tido, markevery = 1, linestyle = '--')
        ax.set_ylim(min_y, max_y)
        ax.spines['top'].set_visible(False)  # 테두리 여부
        ax.spines['right'].set_visible(False)  # 테두리 여부
        ax.tick_params(right = False, left = False, bottom = False, labelleft=True, labelright=False, pad = 0) ### 눈금선 없애기, 라벨위치여부
        fig_b3.legend(loc = (0.13, 0.73), fontsize = 9) # 범례 표시
        # y1 차트 데이터레이블 표시
        for idx, txt in enumerate(y):
            ax.text(x[idx], y[idx] + 0.3, txt, ha='center')
        fig_b3.show()
    elif gungucbox.get() != '전체' and Tgungucbox.get() != '전체' and sido != ['서울특별시', '부산광역시', '대구광역시', '인천광역시', '광주광역시', \
        '대전광역시', '울산광역시'] and tido != ['서울특별시', '부산광역시', '대구광역시', '인천광역시', '광주광역시', '대전광역시', '울산광역시']:
        ### 청년 꺾은선 차트 용 dataframe 'c'
        a = df.loc[(df['시도'] == sido) & (df['행정구역'] == sigu), ['시점', '고용률 (%)']]
        b = df.loc[(df['시도'] == tido) & (df['행정구역'] == tigu), ['시점', '고용률 (%)']]
        c = pd.merge(a,b, on = "시점")
        c = c.sort_values('시점', ascending = True)
        c = c.reset_index(drop = True)
        a = c.copy()
        c.rename(columns = {"고용률 (%)_x" : n_sido, "고용률 (%)_y" : n_tido}, inplace = True)
        x = c['시점']
        y = c[n_sido]
        z = c[n_tido]
        maxmin = df.loc[((df['시도'] == sido) & (df['행정구역'] == sigu)) | ((df['시도'] == tido) & (df['행정구역'] == tigu)), ['고용률 (%)']]
        max_y = maxmin['고용률 (%)'].max() + 0.9  ## y축 최대값
        min_y = maxmin['고용률 (%)'].min() - 0.7  ## y축 최소값

        fig_b3 = plt.figure(figsize = (6,3), dpi = 100)
        ax = fig_b3.add_subplot()
        fig_b3.suptitle("[ '" + n_sido + "' 청년고용률(15-29세) ]", fontsize = 11, y=0.95)
        ax.plot(x, y, color = 'red', linewidth = 2, marker = "o", markeredgewidth = 1, \
            markeredgecolor = "firebrick", markerfacecolor = "yellow", markersize = 6, label = n_sido, markevery = 1)
        ax.plot(x, z, color = 'darkblue', linewidth = 1, marker = "o", markeredgewidth = 1, \
            markeredgecolor = "darkblue", markerfacecolor = "darkblue", markersize = 3, label = n_tido, markevery = 1, linestyle = '--')
        ax.set_ylim(min_y, max_y)
        ax.spines['top'].set_visible(False)  # 테두리 여부
        ax.spines['right'].set_visible(False)  # 테두리 여부
        ax.tick_params(right = False, left = False, bottom = False, labelleft=True, labelright=False, pad = 0) ### 눈금선 없애기, 라벨위치여부
        fig_b3.legend(loc = (0.13, 0.73), fontsize = 9) # 범례 표시
        # y1 차트 데이터레이블 표시
        for idx, txt in enumerate(y):
            ax.text(x[idx], y[idx] + 0.3, txt, ha='center')
        fig_b3.show()
    else:
        a = df.loc[(df['시도'] == sido) & (df['행정구역'] == sigu), ['시점', '고용률 (%)']]
        a = a.sort_values('시점', ascending = True)
        a = a.reset_index(drop=True)
        x = a['시점']
        y = a['고용률 (%)']
        max_y = a['고용률 (%)'].max() + 0.9  ## y축 최대값
        min_y = a['고용률 (%)'].min() - 0.7  ## y축 최소값
        fig_b3 = plt.figure(figsize = (6, 3), dpi = 100)
        ax = fig_b3.add_subplot()
        fig_b3.suptitle("[ " + n_sido + " 청년고용률(15-29세) ]", fontsize = 11, y = 0.95)
        ax.plot(x, y, color = 'red', linewidth = 2, marker = "o", markeredgewidth = 1, \
            markeredgecolor = "firebrick", markerfacecolor = "yellow", markersize = 6, label = n_sido, markevery = 1)
        ax.set_ylim(min_y, max_y)
        ax.spines['top'].set_visible(False)  # 테두리 여부
        ax.spines['right'].set_visible(False)  # 테두리 여부
        ax.tick_params(right = False, left = False, bottom = False, labelleft=True, labelright=False, pad = 0) ### 눈금선 없애기, 라벨위치여부
        # y1 차트 데이터레이블 표시
        for idx, txt in enumerate(y):
            ax.text(x[idx], y[idx] + 0.3, txt, ha='center')
        fig_b3.show()
def b4_plot():
    df = pd.read_csv(resource_path('data/2_3_db_job.csv'), encoding = 'cp949')
    df1 = df.loc[((df['시도'] == tido) & (df['행정구역'] == tigu) & (df['직업별'] != "계")), ['직업별', job_month]]
    df1['비중'] = round(df1[job_month] / df1[job_month].sum() * 100, 1)
    df2 = df.loc[((df['시도'] == sido) & (df['행정구역'] == sigu) & (df['직업별'] != "계")), ['직업별', job_month]]
    df2['비중'] = round(df2[job_month] / df2[job_month].sum() * 100, 1)
    df3 = pd.merge(df1, df2, on = "직업별")
    at = df3.copy()
    df3 = df3[['직업별', '비중_x', '비중_y']]
    df3 = df3.rename(columns = {'비중_x' : n_tido, '비중_y' : n_sido})
    df = df3.reset_index(drop = True)
    labels = df['직업별']

    fig = plt.figure(figsize = (5, 2.8), dpi = 100)

    ax = fig.add_subplot(111, projection="polar")
    theta = np.arange(len(df) + 1) / float(len(df)) * 2 * np.pi
    values = df[n_sido].values
    values = np.append(values, values[0])
    ax.plot(theta, values, color="crimson", label=n_sido)
    ax.tick_params(pad=10)
    plt.xticks(theta[:-1], labels, fontsize=10)
    ax.fill(theta, values, 'crimson', alpha=0.1)

    # ax = fig.add_subplot(111, projection="polar")
    # theta = np.arange(len(df) + 1) / float(len(df)) * 2 * np.pi
    values = df[n_tido].values
    values = np.append(values, values[0])
    ax.plot(theta, values, color="cornflowerblue", label= n_tido)
    ax.tick_params(pad=10)
    ax.fill(theta, values, 'cornflowerblue', alpha=0.1)
    plt.legend(bbox_to_anchor=(0.1, 0.98))
    fig.tight_layout()  ## 이미지 타이트하게
    fig.show()
def b5_plot():
    df = pd.read_csv(resource_path('data/2_3_db_joblevel.csv'), encoding = 'cp949')
    df1 = df.loc[((df['시도'] == tido) &(df['행정구역'] == tigu) & (df['종사상지위별'] != "계")), ['종사상지위별', job_month]]
    df1['비중'] = round(df1[job_month] / df1[job_month].sum() * 100, 1)
    df2 = df.loc[((df['시도'] == sido) & (df['행정구역'] == sigu) & (df['종사상지위별'] != "계")), ['종사상지위별', job_month]]
    df2['비중'] = round(df2[job_month] / df2[job_month].sum() * 100, 1)
    df3 = pd.merge(df1, df2, on = "종사상지위별")
    at = df3.copy()
    df3 = df3[['종사상지위별', '비중_x', '비중_y']]
    df3 = df3.rename(columns = {'비중_x' : n_tido, '비중_y' : n_sido})
    df = df3.reset_index(drop = True)
    labels = df['종사상지위별']

    fig_s2 = plt.figure(figsize = (5, 2.8), dpi = 100)

    ax = fig_s2.add_subplot(111, projection="polar")
    theta = np.arange(len(df) + 1) / float(len(df)) * 2 * np.pi
    values = df[n_sido].values
    values = np.append(values, values[0])
    ax.plot(theta, values, color="crimson", label=n_sido)
    ax.tick_params(pad=10)
    plt.xticks(theta[:-1], labels, fontsize=10)
    ax.fill(theta, values, 'crimson', alpha=0.1)

    values = df[n_tido].values
    values = np.append(values, values[0])
    ax.plot(theta, values, color="cornflowerblue", label= n_tido)
    ax.tick_params(pad=10)
    ax.fill(theta, values, 'cornflowerblue', alpha=0.1)
    plt.legend(bbox_to_anchor=(0.1, 0.98))
    fig_s2.tight_layout()  ## 이미지 타이트하게
    fig_s2.show()
def c1_plot():
    df = pd.read_csv(resource_path('data/3_1_db_company.csv'), encoding = 'cp949')
    ### 조건에 맞는 자료 추출 및 머지
    x = df.loc[(df['시도'] == tido) & (df['행정구역'] == tigu) & (df['산업대'] != "전체"), ['산업대', '2020 년']]
    y = df.loc[(df['시도'] == sido) & (df['행정구역'] == sigu) & (df['산업대'] != "전체"), ['산업대', '2020 년']]
    a = pd.merge(x,y, on = '산업대')
    a = a.sort_values('산업대', ascending = False)
    ## 바차트 그리기 ##
    fig_c1 = plt.figure(figsize = (7, 4), dpi = 100)
    ax = fig_c1.add_subplot(1,2,1)
    ax.barh(a['산업대'], a["2020 년_y"], color='slateblue')
    fig_c1.patch.set_visible(False)
    ax.grid(False)
    ax.tick_params(labelleft=False, labelright=True, pad = 0)  
    ax.spines['bottom'].set_visible(True)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    # ax.spines['left'].set_visible(False)
    ax.tick_params(axis = 'y', length = 30, right = False, left = False, bottom = False) ### 눈금선 없애기
    fig_c1.tight_layout()  ## 이미지 타이트하게
    ## 사업체수 두번째 : 전국 바차트 그리기 ##
    ax2 = plt.subplot(1,2,2)
    fig_c1.suptitle("- " + n_sido + " 사업체수 →                                                 ← " + n_tido + " 사업체수-", fontsize = 11, x = 0.5, y = 0.95)
    ax2.barh(a['산업대'], -a["2020 년_x"], color='slateblue')
    ax2.spines['bottom'].set_visible(True)  # 테두리 여부
    ax2.spines['top'].set_visible(False)  # 테두리 여부
    # ax.spines['right'].set_visible(False)  # 테두리 여부
    ax2.spines['left'].set_visible(False)  # 테두리 여부
    ax2.tick_params(right = False, left = False, bottom = False, labelleft=False, labelright=False, pad = 0) ### 눈금선 없애기, 라벨위치여부
    fig_c1.tight_layout()  ## 이미지 타이트하게
    fig_c1.show()
def c2_plot():
    ## csv 파일 열기
    df = pd.read_csv(resource_path('data/3_2_db_company.csv'), encoding = 'cp949')
    ### 조건에 맞는 자료 추출 및 머지
    x = df.loc[(df['시도'] == tido) & (df['행정구역'] == tigu) & (df['산업대'] != "전체"), ['산업대', '2020 년']]
    y = df.loc[(df['시도'] == sido) & (df['행정구역'] == sigu) & (df['산업대'] != "전체"), ['산업대', '2020 년']]
    a = pd.merge(x,y, on = '산업대')
    a = a.sort_values('산업대', ascending = False)
    ## 바차트 그리기 ##
    fig_c2 = plt.figure(figsize = (7, 4), dpi = 100)
    ax = fig_c2.add_subplot(1,2,1)
    ax.barh(a['산업대'], a["2020 년_y"], color='mediumOrchid')
    ax.grid(False)
    ax.tick_params(labelleft=False, labelright=True, pad = 0)  
    ax.spines['bottom'].set_visible(True)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.tick_params(axis = 'y', length = 30, right = False, left = False, bottom = False) ### 눈금선 없애기
    fig_c2.tight_layout()  ## 이미지 타이트하게
    ## 종사자수 두번째 : 전국 바차트 그리기 ##
    ax2 = plt.subplot(1,2,2)
    fig_c2.suptitle("- " + n_sido + " 종사자수 →                                                 ← " + n_tido + " 종사자수-", fontsize = 11, x = 0.5, y = 0.95)
    ax2.barh(a['산업대'], -a["2020 년_x"], color='mediumOrchid')
    ax2.spines['bottom'].set_visible(True)  # 테두리 여부
    ax2.spines['top'].set_visible(False)  # 테두리 여부
    # ax.spines['right'].set_visible(False)  # 테두리 여부
    ax2.spines['left'].set_visible(False)  # 테두리 여부
    ax2.tick_params(right = False, left = False, bottom = False, labelleft=False, labelright=False, pad = 0) ### 눈금선 없애기, 라벨위치여부
    ax2.ticklabel_format(style = 'plain', axis = 'x') # 오류표시가 나타나지 않게 함  
    fig_c2.tight_layout()  ## 이미지 타이트하게
    fig_c2.show()
def d1_plot():
    df = pd.read_csv(resource_path('data/4_1_db_guin.csv'), encoding = 'cp949')
    ### 조건에 맞는 자료 추출
    z = df.loc[(df['시도'] == tido) & (df['행정구역'] == tigu), ['기준년월', '취업건수', '구인배율', '월']]
    z = z.rename(columns = {'구인배율' : n_tido, '취업건수' : n_tido + '취업'})
    a = df.loc[(df['시도'] == sido) & (df['행정구역'] == sigu), ['기준년월', '취업건수', '구인배율', '월']]
    a = pd.merge(a, z, on = '기준년월')
    a = a.rename(columns = {'구인배율' : n_sido})
    a['연월'] = a['월_x']
    for i in [0,7,19,31,36]:   ### Treeview 간격 변형
        a.loc[i, '연월'] = a.loc[i,'기준년월']
    x = a['기준년월']
    y = round(a[n_sido], 2)
    z = round(a[n_tido], 2)
    maxmin = df.loc[(((df['시도'] == tido) & (df['행정구역'] == tigu)) | ((df['시도'] == sido) & (df['행정구역'] == sigu))), ['구인배율']]
    max_y = maxmin['구인배율'].max() + 0.1  ## y축 최대값
    if maxmin['구인배율'].max() - maxmin['구인배율'].min() > 1:
        min_y = maxmin['구인배율'].min() - 2.0  ## y축 최소값
    elif maxmin['구인배율'].max() - maxmin['구인배율'].min() > 0.5:
        min_y = maxmin['구인배율'].min() - 1.4  ## y축 최소값
    else:
        min_y = maxmin['구인배율'].min() - 0.7
    ## 꺾은선차트 그리기 ##
    fig_d1 = plt.figure(figsize = (13.7, 4.2), dpi = 100)
    ax = fig_d1.add_subplot(111)
    # fig_d1, ax = plt.subplots()
    fig_d1.suptitle("[ 구인배율(꺾은선)과 취업건수(막대) ]", fontsize = 11, y = 0.95)
    ax.plot(x, y, color = 'red', linewidth = 2, marker = "o", markeredgewidth = 1, \
        markeredgecolor = "red", markerfacecolor = 'yellow', markersize = 6, label = n_sido + "구인배율", markevery = 1)
    ax.plot(x, z, color = 'gray', linewidth = 1, marker = "o", markeredgewidth = 1, \
        markeredgecolor = 'gray', markerfacecolor = backcolor, markersize = 3, label = n_tido + "구인배율", markevery = 1, linestyle = '--')
    ax.set_ylim(min_y, max_y)
    ax.set_ylabel('구인배율(p)')
    ax.spines['top'].set_visible(False)  # 테두리 여부
    ax.spines['right'].set_visible(False)  # 테두리 여부
    ax.tick_params(right = False, left = False, bottom = False, labelleft=True, labelright=False, pad = 0) ### 눈금선 없애기, 라벨위치여부
    for idx, txt in enumerate(y):
        ax.text(x[idx], y[idx] + 0.05, txt, ha='center')
    ## 취업건수 바차트 추가 그리기 ##
    ax2 = ax.twinx()
    ax2.set_ylabel('취업건수(건)')
    amax = a["취업건수"].max() * 2
    ax2.set_ylim(0, amax)
    ax2.bar(x, a["취업건수"], color = '#ff812d', label = n_sido + "취업건수")
    ax2.spines['top'].set_visible(False)  # 테두리 여부
    fig_d1.legend(loc = (0.05, 0.85), fontsize = 9) # 범례 표시
    fig_d1.autofmt_xdate(rotation=45)  ## x축 레이블 45도각도 눕히기
    # a["취업건수"] = a.apply(lambda x: "{:,}".format(x["취업건수"]), axis=1)
    for idx, txt in enumerate(a["취업건수"]):
        ax2.text(idx, a["취업건수"][idx] + 12, txt, ha='center')
    fig_d1.tight_layout()  ## 이미지 타이트하게
    fig_d1.show()
def d2_plot():
    df = pd.read_csv(resource_path('data/4_2_db_guin.csv'), encoding = 'cp949')
    a = df.loc[((df['시도'] == sido) & (df['행정구역'] == sigu) & (df['기준년'] == "21년")), ["직종_중분류", "신규구인인원", "신규구직건수", "구인배율"]]
    filt = (a['신규구인인원'] > 10) & (a['신규구직건수'] > 10)
    a = a[filt]
    a = a.sort_values('구인배율', ascending = True)
    a = a.reset_index(drop = True)
    a['구인배율'] = round(a['구인배율'],2)
    a1 = a.iloc[0:5,[0,3]]
    a1 = a1.sort_values('구인배율', ascending = False)  # a1 완료
    a2 = a.sort_values('구인배율', ascending = False)
    a2 = a2.reset_index(drop = True)
    a2 = a2.iloc[0:5, [0,3]]
    a2 = a2.sort_values('구인배율', ascending = True)
    a = pd.concat([a1, a2], ignore_index = True)
    fig_d2 = plt.figure(figsize = (6, 2.3), dpi = 100)
    ax = fig_d2.add_subplot()
    fig_d2.suptitle("[ " + n_sido + " 직종별 구인구직 미스매치('21년) ]", fontsize = 11, y = 0.95)
    ax.barh(a['직종_중분류'], a["구인배율"], color='lightpink')
    ax.grid(False)
    ax.tick_params(labelleft=True, labelright=False, pad = 0)  
    ax.spines['bottom'].set_visible(True)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.tick_params(axis = 'y', right = False, left = False, bottom = False) ### 눈금선 없애기
    fig_d2.tight_layout()  ## 이미지 타이트하게
    fig_d2.show()


# 결과파일 저장 (폴더)
def browse_dest_path():
    folder_selected = filedialog.askdirectory()
    if folder_selected == "": # 사용자가 취소를 누를 때
        return
    #print(folder_selected)
    txt_dest_path.delete(0, END)
    txt_dest_path.insert(0, folder_selected)
def popdata():
    if len(txt_dest_path.get()) == 0:
        msgbox.showwarning("확인", "'폴더선택' 버튼을 먼저 클릭하세요." + "\n" + \
            "  선택한 폴더에 파일이 저장됩니다.")
        return
    try:
        name_data = 'data_1_인구_'+ datenow +'.xlsx'
        dest_path = os.path.join(txt_dest_path.get(), name_data)
        file_from = resource_path('data/data_1_pop.xlsx')
        shutil.copy(file_from, dest_path)
        time.sleep(0.5)  
        msgbox.showinfo("알림", "'인구 데이터' 다운로드가 완료되었습니다." + "\n" + " 지정 폴더에서 파일을 확인하세요!")
    except:
        msgbox.showwarning("확인", "같은 이름의 엑셀파일이 현재 실행중인지 확인바랍니다!" + "\n" + \
            "실행 중인 파일을 닫고 다시 실행하시기 바랍니다.")
def empdata():
    if len(txt_dest_path.get()) == 0:
        msgbox.showwarning("확인", "'폴더선택' 버튼을 먼저 클릭하세요." + "\n" + \
            "  선택한 폴더에 파일이 저장됩니다.")
        return
    try:
        name_data = 'data_2_고용_'+ datenow +'.xlsx'
        dest_path = os.path.join(txt_dest_path.get(), name_data)
        file_from = resource_path('data/data_2_emp.xlsx')
        shutil.copy(file_from, dest_path)
        time.sleep(0.5)  
        msgbox.showinfo("알림", "'고용 데이터' 다운로드가 완료되었습니다." + "\n" + " 지정 폴더에서 파일을 확인하세요!")
    except:
        msgbox.showwarning("확인", "같은 이름의 엑셀파일이 현재 실행중인지 확인바랍니다!" + "\n" + \
            "실행 중인 파일을 닫고 다시 실행하시기 바랍니다.") 
def sanupdata():
    if len(txt_dest_path.get()) == 0:
        msgbox.showwarning("확인", "'폴더선택' 버튼을 먼저 클릭하세요." + "\n" + \
            "  선택한 폴더에 파일이 저장됩니다.")
        return
    try:
        name_data = 'data_3_산업_'+ datenow +'.xlsx'
        dest_path = os.path.join(txt_dest_path.get(), name_data)
        file_from = resource_path('data/data_3_sanup.xlsx')
        shutil.copy(file_from, dest_path)
        time.sleep(0.5)  
        msgbox.showinfo("알림", "'산업 데이터' 다운로드가 완료되었습니다." + "\n" + " 지정 폴더에서 파일을 확인하세요!")
    except:
        msgbox.showwarning("확인", "같은 이름의 엑셀파일이 현재 실행중인지 확인바랍니다!" + "\n" + \
            "실행 중인 파일을 닫고 다시 실행하시기 바랍니다.")              
def guindata():
    if len(txt_dest_path.get()) == 0:
        msgbox.showwarning("확인", "'폴더선택' 버튼을 먼저 클릭하세요." + "\n" + \
            "  선택한 폴더에 파일이 저장됩니다.")
        return
    try:
        name_data = 'data_4_구인구직_'+ datenow +'.xlsx'
        dest_path = os.path.join(txt_dest_path.get(), name_data)
        file_from = resource_path('data/data_4_guin.xlsx')
        shutil.copy(file_from, dest_path)
        time.sleep(0.5)  
        msgbox.showinfo("알림", "'구인구직 데이터' 다운로드가 완료되었습니다." + "\n" + " 지정 폴더에서 파일을 확인하세요!")
    except:
        msgbox.showwarning("확인", "같은 이름의 엑셀파일이 현재 실행중인지 확인바랍니다!" + "\n" + \
            "실행 중인 파일을 닫고 다시 실행하시기 바랍니다.")  
def totaldata():
    if len(txt_dest_path.get()) == 0:
        msgbox.showwarning("확인", "'폴더선택' 버튼을 먼저 클릭하세요." + "\n" + \
            "  선택한 폴더에 파일이 저장됩니다.")
        return
    try:
        name_data = 'data_전체자료(인구+고용+산업+구인구직)_'+ datenow +'.xlsx'
        dest_path = os.path.join(txt_dest_path.get(), name_data)
        file_from = resource_path('data/data_5_tot.xlsx')
        shutil.copy(file_from, dest_path)
        time.sleep(0.5)  
        msgbox.showinfo("알림", "'전체 데이터' 다운로드가 완료되었습니다." + "\n" + " 지정 폴더에서 파일을 확인하세요!")
    except:
        msgbox.showwarning("확인", "같은 이름의 엑셀파일이 현재 실행중인지 확인바랍니다!" + "\n" + \
            "실행 중인 파일을 닫고 다시 실행하시기 바랍니다.")



################--- 엑셀 자료 바로 실행 ---######################################################################################################

def createFolder():
    try:
        os.makedirs(resource_path('Temp'))
    except:
        pass
def a1_dtdown():
    createFolder()
    a1_dt.to_excel(resource_path('Temp/data_인구시계열_' + oreg + '.xlsx'), sheet_name = oreg, index = False)
    os.startfile(resource_path('Temp/data_인구시계열_' + oreg + '.xlsx'))
def a2_dtdown():
    createFolder()
    a2_dt.to_excel(resource_path('Temp/data_인구피라미드_' + oreg + '.xlsx'), sheet_name = oreg, index = False)
    os.startfile(resource_path('Temp/data_인구피라미드_' + oreg + '.xlsx'))
def a3_dtdown():
    createFolder()
    a3_dt.to_excel(resource_path('Temp/data_고령화율_' + oreg + '.xlsx'), sheet_name = oreg, index = False)
    os.startfile(resource_path('Temp/data_고령화율_' + oreg + '.xlsx'))    
def b1_dtdown():
    createFolder()
    b1_dt.to_excel(resource_path('Temp/data_고용률추이_' + oreg + '.xlsx'), sheet_name = oreg, index = False)
    os.startfile(resource_path('Temp/data_고용률추이_' + oreg + '.xlsx'))    
def b2_dtdown():
    createFolder()
    b2_dt.to_excel(resource_path('Temp/data_성별고용률_' + oreg + '.xlsx'), sheet_name = oreg, index = False)
    os.startfile(resource_path('Temp/data_성별고용률_' + oreg + '.xlsx'))
def tt_b3_dtdown():
    createFolder()
    tt_b3_dt.to_excel(resource_path('Temp/data_청년고용률_광역_' + oreg + '.xlsx'), sheet_name = oreg, index = False)
    os.startfile(resource_path('Temp/data_청년고용률_광역_' + oreg + '.xlsx'))
def rr_b3_dtdown():
    createFolder()
    rr_b3_dt.to_excel(resource_path('Temp/data_청년고용률_기초_' + oreg + '.xlsx'), sheet_name = oreg, index = False)
    os.startfile(resource_path('Temp/data_청년고용률_기초_' + oreg + '.xlsx'))
def etc1_b3_dtdown():
    createFolder()
    etc1_b3_dt.to_excel(resource_path('Temp/data_청년고용률a_' + oreg + '.xlsx'), sheet_name = oreg, index = False)
    os.startfile(resource_path('Temp/data_청년고용률a_' + oreg + '.xlsx'))
def etc2_b3_dtdown():
    createFolder()
    etc2_b3_dt.to_excel(resource_path('Temp/data_청년고용률b_' + n_tido + '.xlsx'), sheet_name = n_tido, index = False)
    os.startfile(resource_path('Temp/data_청년고용률b_' + n_tido + '.xlsx'))    
def b4_dtdown():
    createFolder()
    b4_dt.to_excel(resource_path('Temp/data_직업별취업자_' + oreg + '.xlsx'), sheet_name = oreg, index = False)
    os.startfile(resource_path('Temp/data_직업별취업자_' + oreg + '.xlsx'))
def b5_dtdown():
    createFolder()
    b5_dt.to_excel(resource_path('Temp/data_종사상지위별취업자_' + oreg + '.xlsx'), sheet_name = oreg, index = False)
    os.startfile(resource_path('Temp/data_종사상지위별취업자_' + oreg + '.xlsx'))    
def c1_dtdown():
    createFolder()
    c1_dt.to_excel(resource_path('Temp/data_사업체수_' + oreg + '.xlsx'), sheet_name = oreg, index = False)
    os.startfile(resource_path('Temp/data_사업체수_' + oreg + '.xlsx'))
def c2_dtdown():
    createFolder()
    c2_dt.to_excel(resource_path('Temp/data_종사자수_' + oreg + '.xlsx'), sheet_name = oreg, index = False)
    os.startfile(resource_path('Temp/data_종사자수_' + oreg + '.xlsx'))
def d1_dtdown():
    createFolder()
    d1_dt.to_excel(resource_path('Temp/data_구인배율_' + oreg + '.xlsx'), sheet_name = oreg, index = False)
    os.startfile(resource_path('Temp/data_구인배율_' + oreg + '.xlsx'))
def d2_dtdown():
    createFolder()
    d2_dt.to_excel(resource_path('Temp/data_신규구인구직건수_' + oreg + '.xlsx'), sheet_name = oreg, index = False)
    os.startfile(resource_path('Temp/data_신규구인구직건수_' + oreg + '.xlsx'))


##-------------------------------------분석자료 상단 프레임 -------------------------------------##
# main_frame = Frame(mainw, bg = backcolor) 
# main_frame.pack(side='top', anchor="n", fill = "both")

selcol = random.choice(['lightblue1', 'lightcyan','lightyellow2','thistle1','lightgoldenrodyellow','lightsteelblue1'])

## 분석 메인 라벨 _ 선택한 시군구 표시, search버튼에 연결  # 추천색상: lightcyan, lightyellow2, thistle1, lightgoldenrodyellow, 
main_label = Label(mainw, bg = selcol, \
    relief = "raised", borderwidth = 2, text = "[ 전국 시군구별 고용지표 분석  ]", \
    font = ("arial", 15, "bold"), padx = 5, pady = 20)
main_label.pack(side='top', anchor="n", fill = "x")

photo = PhotoImage(file = resource_path("data/logo_array_fin.png"))
btnphoto = Button(main_label, relief = "flat", overrelief = "raised", command = activate_info)
# btnphoto.pack(side = LEFT, padx = 20, pady = 8)


##------------------------------------- 지역선택(오른쪽) 프레임 -------------------------------------##
## 1. '분석지역' 선택(오른쪽) 프레임
R_frame = Frame(mainw, relief = "raised", borderwidth = 2)
R_frame.pack(side='right', anchor="n", fill = "y")

## 지역 선택 라벨 프레임
choice_frame = LabelFrame(R_frame, text='☞ 분석 지역', relief = "raised", padx=5, pady=5) # padx / pady 내부여백
choice_frame.pack(side=TOP, anchor="ne", padx=15, pady=90) # padx / pady 외부여백

# 지역선택 라벨
label_a = Label(choice_frame, text="☞ 광역")
label_b = Label(choice_frame, text="☞ 기초")

## 지역선택_광역 콤보상자
sidocbox = ttk.Combobox(choice_frame, height = 0, \
    values = sido, state = "readonly")
sidocbox.set('선택')   # 최초 목록 제목 설정
# sidocbox.bind("<<ComboboxSelected>>", gunguchoice)
sidocbox.bind('<<ComboboxSelected>>', on_select)

## 지역선택_기초시군구 콤보상자
gungucbox = ttk.Combobox(choice_frame, height = 0, state = "readonly")
gungucbox.set('선택')
# gungucbox.pack(side = "right", anchor = "ne", padx = 15, pady = 5)

#### 지역선택_검색버튼
searchb = Button(choice_frame, text = "조회", font = ("arial", 10, "bold"), fg = 'darkblue', bg = "#19CE60", command = search)

## 지역선택_그리드
label_a.grid(column = 0, row = 0, sticky = N+E+W+S,  padx = 5, pady = 5)
label_b.grid(column = 0, row = 1, sticky = N+E+W+S,  padx = 5, pady = 5)
sidocbox.grid(column = 1, row =  0, sticky = N+E+W+S, padx = 3, pady = 5)
gungucbox.grid(column = 1, row =1, sticky = N+E+W+S, padx = 3, pady = 5)
searchb.grid(column = 2, row =0, sticky = N+E+W+S, padx = 5, pady = 5, rowspan = 3)


######################################## 2. '비교 지역' 선택(오른쪽) 프레임
## 비교 지역 선택 라벨 프레임
Tchoice_frame = LabelFrame(choice_frame, text='★ 비교 지역', bg = 'lavender', relief = "flat", padx=5, pady=5) # padx / pady 내부여백
# Tchoice_frame.pack(side=TOP, anchor="ne", padx=15, pady=0) # padx / pady 외부여백

# 지역선택 라벨
Tlabel_a = Label(Tchoice_frame, text="★ 광역", bg = 'lavender',)
Tlabel_b = Label(Tchoice_frame, text="★ 기초", bg = 'lavender',)

## 지역선택_광역 콤보상자
Tsidocbox = ttk.Combobox(Tchoice_frame, height = 0, \
    values = sido, state = "readonly")
Tsidocbox.set('선택')   # 최초 목록 제목 설정
# sidocbox.bind("<<ComboboxSelected>>", gunguchoice)
Tsidocbox.bind('<<ComboboxSelected>>', Ton_select)

## 지역선택_기초시군구 콤보상자
Tgungucbox = ttk.Combobox(Tchoice_frame, height = 0, state = "readonly")
Tgungucbox.set('선택')
# gungucbox.pack(side = "right", anchor = "ne", padx = 15, pady = 5)

#### 지역선택_검색버튼
# Tsearchb = Button(choice_frame, text = "조회", bg = "lime green", command = search)

## 지역선택_그리드
Tchoice_frame.grid(column = 0, row = 2, columnspan = 2, sticky = N+E+W+S,  padx = 5, pady = 5)
Tlabel_a.grid(column = 0, row = 3, sticky = N+E+W+S,  padx = 5, pady = 5)
Tlabel_b.grid(column = 0, row = 4, sticky = N+E+W+S,  padx = 5, pady = 5)
Tsidocbox.grid(column = 1, row =  3, sticky = N+E+W+S, padx = 3, pady = 5)
Tgungucbox.grid(column = 1, row = 4, sticky = N+E+W+S, padx = 3, pady = 5)
# Tsearchb.grid(column = 2, row =0, sticky = N+E+W+S, padx = 5, pady = 5, rowspan = 2)



####### 자료다운(오른쪽) 프레임####
downdata_frame = LabelFrame(R_frame, text='※ 부문별 전국 자료 다운로드(.xlsx)', relief = "raised", bd=1, padx=5, pady=5) # padx / pady 내부여백
downdata_frame.pack(side = 'bottom', anchor="nw", fill="x", padx = 15, pady = 90) # padx / pady 외부여백

up_frame = Frame(downdata_frame, relief = "flat")
up_frame.pack(side='top', anchor="w", fill = "x")

txt_dest_path = Entry(up_frame)  ### str 값 반환할 Entry, 선택된 폴더 경로 알려줌
txt_dest_path.pack(side="left", fill="x", expand=True, padx= 5, pady=7, ipady=4) # 높이 변경

btn_dest_path = Button(up_frame, text="폴더선택", command=browse_dest_path) ### 폴더선택 버튼
btn_dest_path.pack(side="right", padx= 5, pady=5)

totaldown_frame = Frame(downdata_frame, relief = "flat")
totaldown_frame.pack(side='bottom', anchor="s", fill = "x")
downbt_total = Button(totaldown_frame, text = "전체자료 하나의 파일로 받기", width = 6, bg= "linen", relief = "raised", activebackground='snow', padx=5, pady=2, command = totaldata)
downbt_total.pack(side = 'bottom', anchor = 's',fill = 'x', padx= 3, pady=3, ipady=4)

down_frame = Frame(downdata_frame, relief = "flat")
down_frame.pack(side='bottom', anchor="s", fill = "x")

downbt_popdata = Button(down_frame, text = "인 구", width = 6, relief = "raised", activebackground='honeydew', padx=7, pady=2, command = popdata)
downbt_popdata.pack(side = 'left', padx= 5, pady=3, ipady=4)

downbt_empdata = Button(down_frame, text = "고 용", width = 6, relief = "raised", activebackground='honeydew', padx=7, pady=2, command = empdata)
downbt_empdata.pack(side = 'left', padx= 5, pady=3, ipady=4)

downbt_sanupdata = Button(down_frame, text = "산 업",  width = 6, relief = "raised", activebackground='honeydew', padx=7, pady=2, command = sanupdata)
downbt_sanupdata.pack(side = 'left', padx= 5, pady=3, ipady=4)

downbt_guindata = Button(down_frame, text = "구인구직", width = 6, relief = "raised", activebackground='honeydew', padx=5, pady=2, command = guindata)
downbt_guindata.pack(side = 'left', padx= 5, pady=3, ipady=4)


##-------------------------------------메인 분석 프레임(왼쪽)과 스크롤바 -------------------------------------##
L_frame = tk.Frame(mainw, bg= maincolor)
L_frame.pack(side='left', anchor="n", fill = "both", expand = 'true')

my_canvas = tk.Canvas(L_frame, bg= maincolor)

vscrollbar = tk.Scrollbar(L_frame, orient = 'vertical', command = my_canvas.yview, bg = 'white')
hscrollbar = tk.Scrollbar(L_frame, orient = 'horizontal', command = my_canvas.xview, bg = 'white')
my_canvas.grid(row=0, column=0, sticky="nsew")
vscrollbar.grid(row=0, column=1, sticky="ns")
hscrollbar.grid(row=1, column=0, sticky="ew")
L_frame.grid_rowconfigure(0, weight=1)  # horizontal, vertical 두개를 grid로 적용하는 핵심 설정
L_frame.grid_columnconfigure(0, weight=1) #             "               "

my_canvas.configure(yscrollcommand = vscrollbar.set)
my_canvas.configure(xscrollcommand = hscrollbar.set)
my_canvas.bind('<Configure>', lambda e: my_canvas.configure(scrollregion = my_canvas.bbox("all")))  #### 스크롤바 연동 제일 중요한 구문 ★★★★★★
my_canvas.bind_all("<MouseWheel>", on_mousewheel)

t_frame = tk.Frame(my_canvas, bg= maincolor)
t_frame.grid(row=0, column = 0, sticky = 'nsew')
# t_frame.pack(side='top', fill = "both", expand = 'true')
my_canvas.create_window((0,0), window = t_frame, anchor = 'nw')

##-------------------------------------# 인구 분석 프레임-------------------------------------## 색상: '#e4e1d2'
ad0_frame = LabelFrame(t_frame, text='', relief = "flat", bg = maincolor, fg = 'white', \
    padx = 7, pady = 30, font = ("arial", 18, "bold")) # padx / pady 내부여백
ad0_frame.pack(side='top', anchor="w", padx = 30, pady = 10, expand = True)

ad1_frame = LabelFrame(t_frame, text='', relief = "flat", bg = maincolor, fg = 'white', \
    padx = 7, pady = 30, font = ("arial", 18, "bold")) # padx / pady 내부여백
ad1_frame.pack(side='top', anchor="w", padx = 30, pady = 20, expand = True)

first_frame = LabelFrame(ad1_frame, text='     인 구', bd = 0, relief = "raised", bg = maincolor, fg = maincolor, \
    padx = 15, pady = 10, font = ("arial", 22, "bold")) # padx / pady 내부여백
first_frame.pack(side='top', anchor="w", padx = 0, pady = 0, expand = True) # padx / pady 외부여백

## 인구 분석 자료 프레임
Pop_frame = Frame(first_frame, width = 1250, bg = maincolor, relief = "raised", padx = 15, pady = 15) 
Pop_frame.pack(side='top', anchor="w", padx = 0, pady = 0, fill = "both", expand = True)


##-------------------------------------# 고용률 분석 프레임-------------------------------------## 색상: '#ac896d'
ad2_frame = LabelFrame(t_frame, text='', relief = "flat", bg = maincolor, fg = 'black', \
    padx = 7, pady = 30, font = ("arial", 18, "bold")) # padx / pady 내부여백
ad2_frame.pack(side='top', anchor="w", padx = 30, pady = 20, expand = True)

second_frame = LabelFrame(ad2_frame, text='     고 용', relief = "flat", bg = maincolor, fg = maincolor, \
    bd = 0, padx = 15, pady = 10, font = ("arial", 22, "bold")) # padx / pady 내부여백
second_frame.pack(side='top', anchor="w", padx = 0, pady = 0, expand = True) # padx / pady 외부여백

## 고용률 분석 자료 프레임
emp_frame = Frame(second_frame, width = 1250, bg = maincolor, relief = "raised", padx = 15, pady = 15) 
emp_frame.pack(side='top', anchor="w", padx = 0, pady = 0, fill = "both", expand = True)

##-------------------------------------# 사업체종사자 분석 프레임-------------------------------------## 색상: '#016a3f'
ad3_frame = LabelFrame(t_frame, text='',  relief = "flat", bg = maincolor, fg = 'black', \
    padx = 7, pady = 30, font = ("arial", 18, "bold")) # padx / pady 내부여백
ad3_frame.pack(side='top', anchor="w", padx = 30, pady = 20, expand = True)

third_frame = LabelFrame(ad3_frame, text='     산 업', relief = "raised", bg = maincolor, fg = maincolor, \
    bd = 0, padx = 15, pady = 10, font = ("arial", 22, "bold")) # padx / pady 내부여백
third_frame.pack(side='top', anchor="w", padx = 0, pady = 0, expand = True) # padx / pady 외부여백

## 사업체 분석 자료 프레임
sanup_frame = Frame(third_frame, width = 1250, bg = maincolor, relief = "raised", padx = 15, pady = 15) 
sanup_frame.pack(side='top', anchor="w", fill = "both", expand = True)

##-------------------------------------# 구인구직 분석 프레임-------------------------------------## 색상: 'darkcyan'
ad4_frame = LabelFrame(t_frame, text='', relief = "flat", bg = maincolor, fg = 'black', \
    padx = 7, pady = 30, font = ("arial", 18, "bold")) # padx / pady 내부여백
ad4_frame.pack(side='top', anchor="w", padx = 30, pady = 20, expand = True)

forth_frame = LabelFrame(ad4_frame, text='     구인구직', relief = "raised", bg = maincolor, fg = maincolor, \
    bd = 0, padx = 15, pady = 10, font = ("arial", 22, "bold")) # padx / pady 내부여백
forth_frame.pack(side='top', anchor="w", padx = 0, pady = 0, expand = True) # padx / pady 외부여백

## 구인구직 분석 자료 프레임
guin_frame = Frame(forth_frame, width = 1250, bg = maincolor, relief = "raised", padx = 15, pady = 15) 
guin_frame.pack(side='top', anchor="w", fill = "both", expand = True)


##-------------------------------------# 프로그레스바------------------------------##
# p_var = DoubleVar()
# progressbar = ttk.Progressbar(mainw, maximum = 100, mode = 'determinate', length = 150, variable = p_var2)
# progressbar.start(10)
# progressbar.pack(pady = 20)



# ## 사업체 table 추가 프레임
# table_sanup_frame = Frame(third_frame) 
# table_sanup_frame.pack(side='right', anchor="e")


# 메뉴 생성
# menu.add_cascade(label = "  Home", menu=menu_file)
# menu.add_cascade(label = " Analysis ")

# mainw.config(menu=menu)

mainw.mainloop()
