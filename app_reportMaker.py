import tkinter.ttk as ttk
import tkinter.messagebox as msgbox
import func_pathInfo as info
import func_report as report
from tkinter import * # __all__
from tkcalendar import Calendar
from datetime import timedelta
import time
import json

stop_command = False

def add_file():
    update_progressbar()
    date = cal.get_date()
    
    sep_dot = True
    if date.__contains__('/'):
        date = date.split('/')
        sep_dot = False
    elif date.__contains__('.'):
        date = date.split('.')
    else:
        msgbox.showerror("에러",'날짜 구분자가 없습니다.')
    
    date = [x.strip() for x in date]
    
    global g_year; 
    global g_month; 
    global g_day
    
    print(date)
    if sep_dot:
        g_year = date[0]
        g_month = date[1]
        g_day = date[2]
    else:
        g_year = date[2]
        g_month = date[0]
        g_day = date[1]
    
    files = info.get_working_id_lists(g_year, g_month, g_day)
 
    if list_file.size() > 0:
        for index in range(list_file.size()):
            list_file.delete(0)        
    
    if len(files) > 0:
        for file in files:
            list_file.insert(END, file)
        
        typeCountlst = report.extract_typeCount(files)[:]
        print(report.extract_typeCount(files)[:])
        print(typeCountlst)
        update_count(countLists, typeCountlst)
    else:
        msgbox.showwarning("경고", "선택 일자에 작업 파일이 존재하지 않습니다.")

# 선택 삭제
def del_file():
    update_progressbar()
    for index in reversed(list_file.curselection()):
        list_file.delete(index)
        
    if list_file.size() > 0:
        typeCountlst = report.extract_typeCount(list_file.get(0,END))[:]
        print(report.extract_typeCount(list_file.get(0,END))[:])
        print(typeCountlst)
        update_count(countLists, typeCountlst)
        
# 선택 추출
def sel_file():
    update_progressbar()
    print(list_file.size())
    print(reversed(range(list_file.size())))
    for index in reversed(range(list_file.size())):
        if index in list_file.curselection():
            continue
        list_file.delete(index)
        
    if list_file.size() > 0:
        typeCountlst = report.extract_typeCount(list_file.get(0,END))[:]
        print(report.extract_typeCount(list_file.get(0,END))[:])
        print(typeCountlst)
        update_count(countLists, typeCountlst)

# 중지
def stop():
    change_btn_bg_color(btn_start,"SystemButtonFace")    
    change_btn_bg_color(btn_stop,"lightgray")
    
    global stop_command
    stop_command = True    
    print('stop:stop_command = True')
# 시작
def start():
    try:
        change_btn_bg_color(btn_start,"lightgray")
        # 파일 목록 확인
        if list_file.size() == 0:
            msgbox.showwarning("경고", "작업 파일이 존재하지 않습니다")
            return

        # 리포트 생성 작업
        contents = list_file.get(0,END)
        
        global stop_command    
        stop_command = False
        
        idx = 0
        start = time.time()
        final = False
        
        for content in contents:
            if not stop_command:
                idx += 1
                progress = (idx) / len(contents) * 100 # 실제 percent 정보를 계산
                p_var.set(progress)
                progress_bar.update()
                end = time.time()
                elapsed = str(timedelta(seconds=end - start))[:7]
                
                label_progress.config(text=f'경과 시간 : {elapsed}, {idx} / {len(contents)} : {content}')
                label_progress.update()
                
                if idx == len(contents) or stop_command:
                    final = True
                
                # report.make_reports(content, g_year, g_month, g_day, final)   
                except_list = data['side_picture_except_index']
                root.after(1000, report.make_reports(content, g_year, g_month, g_day, final, except_list))
                if final:
                    change_btn_bg_color(btn_stop,"SystemButtonFace")                    
                    msgbox.showinfo("안내","작업이 완료되었습니다.")
                
        end = time.time()
        elapsed = str(timedelta(seconds=end - start))[:7]
        
        label_progress.config(text=f'경과 시간 : {elapsed}, {idx} / {len(contents)} : {content}')
        label_progress.update()
    finally:
        change_btn_bg_color(btn_start,"SystemButtonFace")
        

        
        
def change_btn_bg_color(button, color):
    button.config(background = color)

def update_progressbar():
    label_progress.config(text=f'경과 시간 : 00 : 00 : 00, Unkkown')
    label_progress.update()
    p_var.set(0)
    progress_bar.update()
    
def update_count(objLists, cntLists):
    for index in range(len(cntLists)):
        objLists[index].config(text=cntLists[index])
        
    objLists[-1].config(text= sum(cntLists))
    
    for obj in objLists:
        obj.update()

root = Tk()
root.title("Report Maker v_0.1")
root.geometry("1000x850") # 가로 * 세로

root.resizable(width=True, height=True) # x(너비), y(높이) 값 변경 불가 (창 크기 변경 불가)

# 파일 프레임 (파일 추가, 선택 삭제)
file_frame = Frame(root, borderwidth=2, relief="ridge")
file_frame.pack(padx=5, pady=5, fill="x") # 간격 띄우기

cal = Calendar(file_frame, selectmode = 'day')
cal.pack(padx = 2, pady = 2, side="left")

btn_del_file = Button(file_frame, padx=10, pady=15, width=12, text="선택삭제", command=del_file)
btn_del_file.place(x=250,y=1)

btn_sel_file = Button(file_frame, padx=10, pady=15, width=12, text="선택작업", command=sel_file)
btn_sel_file.place(x=250,y=55)

btn_add_file = Button(file_frame, padx=10, pady=15, width=12, text="목록추출", command=add_file)
btn_add_file.place(x=250,y=110)

# 맨 윗줄
# table_frame = Frame(root, borderwidth=2, relief="solid")
table_frame = Frame(root, borderwidth=2, relief="ridge")
table_frame.place(x=375,y=5)

name_width = 12
cnt_width = 7
pady = 2
height = 1

countLists = []

lbl_f16 = Label(table_frame, text="Target", width=name_width, height=height+1, borderwidth=2, relief="ridge")
lbl_f17 = Label(table_frame, text="Count", width=cnt_width, height=height+1, borderwidth=2, relief="ridge")
lbl_f18 = Label(table_frame, text="Target", width=name_width, height=height+1, borderwidth=2, relief="ridge")
lbl_f19 = Label(table_frame, text="Count", width=cnt_width, height=height+1, borderwidth=2, relief="ridge")
lbl_f20 = Label(table_frame, text="Target", width=name_width, height=height+1, borderwidth=2, relief="ridge")
lbl_f21 = Label(table_frame, text="Count", width=cnt_width, height=height+1, borderwidth=2, relief="ridge")
lbl_f16.grid(row=0, column=0, sticky=N+E+W+S, padx=3, pady=pady)
lbl_f17.grid(row=0, column=1, sticky=N+E+W+S, padx=3, pady=pady)
lbl_f18.grid(row=0, column=2, sticky=N+E+W+S, padx=3, pady=pady)
lbl_f19.grid(row=0, column=3, sticky=N+E+W+S, padx=3, pady=pady)
lbl_f20.grid(row=0, column=4, sticky=N+E+W+S, padx=3, pady=pady)
lbl_f21.grid(row=0, column=5, sticky=N+E+W+S, padx=3, pady=pady)

# clear 줄
lbl_clear = Label(table_frame, text="증착 A3", width=name_width, height=height, borderwidth=2, relief="ridge")
lbl_equal = Label(table_frame, text="0", width=cnt_width, height=height, borderwidth=2, relief="ridge")
lbl_div = Label(table_frame, text="증착 A4", width=name_width, height=height, borderwidth=2, relief="ridge")
lbl_mul = Label(table_frame, text="0", width=cnt_width, height=height, borderwidth=2, relief="ridge")
lbl_cls1 = Label(table_frame, text="증착 A6", width=name_width, height=height, borderwidth=2, relief="ridge")
lbl_cls2 = Label(table_frame, text="0", width=cnt_width, height=height, borderwidth=2, relief="ridge")
lbl_clear.grid(row=1, column=0, sticky=N+E+W+S, padx=3, pady=pady)
lbl_equal.grid(row=1, column=1, sticky=N+E+W+S, padx=3, pady=pady)
lbl_div.grid(row=1, column=2, sticky=N+E+W+S, padx=3, pady=pady)
lbl_mul.grid(row=1, column=3, sticky=N+E+W+S, padx=3, pady=pady)
lbl_cls1.grid(row=1, column=4, sticky=N+E+W+S, padx=3, pady=pady)
lbl_cls2.grid(row=1, column=5, sticky=N+E+W+S, padx=3, pady=pady)
countLists.append(lbl_equal)
countLists.append(lbl_mul)
countLists.append(lbl_cls2)

# 7 시작 줄
lbl_7 = Label(table_frame, text="성산/아성 A3", width=name_width, height=height, borderwidth=2, relief="ridge")
lbl_8 = Label(table_frame, text="0", width=cnt_width, height=height, borderwidth=2, relief="ridge")
lbl_9 = Label(table_frame, text="성산/아성 A4", width=name_width, height=height, borderwidth=2, relief="ridge")
lbl_sub = Label(table_frame, text="0", width=cnt_width, height=height, borderwidth=2, relief="ridge")
lbl_cls3 = Label(table_frame, text="성산/아성 A6", width=name_width, height=height, borderwidth=2, relief="ridge")
lbl_cls4 = Label(table_frame, text="0", width=cnt_width, height=height, borderwidth=2, relief="ridge")
lbl_7.grid(row=2, column=0, sticky=N+E+W+S, padx=3, pady=pady)
lbl_8.grid(row=2, column=1, sticky=N+E+W+S, padx=3, pady=pady)
lbl_9.grid(row=2, column=2, sticky=N+E+W+S, padx=3, pady=pady)
lbl_sub.grid(row=2, column=3, sticky=N+E+W+S, padx=3, pady=pady)
lbl_cls3.grid(row=2, column=4, sticky=N+E+W+S, padx=3, pady=pady)
lbl_cls4.grid(row=2, column=5, sticky=N+E+W+S, padx=3, pady=pady)
countLists.append(lbl_8)
countLists.append(lbl_sub)
countLists.append(lbl_cls4)

# 4 시작 줄
lbl_4 = Label(table_frame, text="OPEN A3", width=name_width, height=height, borderwidth=2, relief="ridge")
lbl_5 = Label(table_frame, text="0", width=cnt_width, height=height, borderwidth=2, relief="ridge")
lbl_6 = Label(table_frame, text="OPEN A4", width=name_width, height=height, borderwidth=2, relief="ridge")
lbl_add = Label(table_frame, text="0", width=cnt_width, height=height, borderwidth=2, relief="ridge")
lbl_cls5 = Label(table_frame, text="OPEN A4", width=name_width, height=height, borderwidth=2, relief="ridge")
lbl_cls6 = Label(table_frame, text="0", width=cnt_width, height=height, borderwidth=2, relief="ridge")
lbl_4.grid(row=3, column=0, sticky=N+E+W+S, padx=3, pady=pady)
lbl_5.grid(row=3, column=1, sticky=N+E+W+S, padx=3, pady=pady)
lbl_6.grid(row=3, column=2, sticky=N+E+W+S, padx=3, pady=pady)
lbl_add.grid(row=3, column=3, sticky=N+E+W+S, padx=3, pady=pady)
lbl_cls5.grid(row=3, column=4, sticky=N+E+W+S, padx=3, pady=pady)
lbl_cls6.grid(row=3, column=5, sticky=N+E+W+S, padx=3, pady=pady)
countLists.append(lbl_5)
countLists.append(lbl_add)
countLists.append(lbl_cls6)

# 1 시작 줄
lbl_1 = Label(table_frame, text="CVD A3", width=name_width, height=height, borderwidth=2, relief="ridge")
lbl_2 = Label(table_frame, text="0", width=cnt_width, height=height, borderwidth=2, relief="ridge")
lbl_3 = Label(table_frame, text="CVD A4", width=name_width, height=height, borderwidth=2, relief="ridge")
lbl_enter = Label(table_frame, text="0", width=cnt_width, height=height, borderwidth=2, relief="ridge")
lbl_cls7 = Label(table_frame, text="CVD A6", width=name_width, height=height, borderwidth=2, relief="ridge")
lbl_cls8 = Label(table_frame, text="0", width=cnt_width, height=height, borderwidth=2, relief="ridge")
lbl_1.grid(row=4, column=0, sticky=N+E+W+S, padx=3, pady=pady)
lbl_2.grid(row=4, column=1, sticky=N+E+W+S, padx=3, pady=pady)
lbl_3.grid(row=4, column=2, sticky=N+E+W+S, padx=3, pady=pady)
lbl_enter.grid(row=4, column=3, sticky=N+E+W+S, padx=3, pady=pady) # 현재 위치로부터 아래쪽으로 몇 줄을 더함
lbl_cls7.grid(row=4, column=4, sticky=N+E+W+S, padx=3, pady=pady)
lbl_cls8.grid(row=4, column=5, sticky=N+E+W+S, padx=3, pady=pady) # 현재 위치로부터 아래쪽으로 몇 줄을 더함
countLists.append(lbl_2)
countLists.append(lbl_enter)
countLists.append(lbl_cls8)

# 0 시작 줄
lbl_0 = Label(table_frame, text="Total", width=name_width, height=height, borderwidth=2, relief="ridge")
lbl_point = Label(table_frame, text="0", width=name_width, height=height, borderwidth=2, relief="ridge")
lbl_0.grid(row=5, column=0, columnspan=2, sticky=N+E+W+S, padx=3, pady=pady) # 현재 위치로부터 오른쪽으로 몇 칸 더함
lbl_point.grid(row=5, column=2, columnspan=2, sticky=N+E+W+S, padx=3, pady=pady)
countLists.append(lbl_point)

# 리스트 프레임
list_frame = Frame(root)
list_frame.pack(fill="both", padx=5, pady=5)

scrollbar = Scrollbar(list_frame)
scrollbar.pack(side="right", fill="y")

list_file = Listbox(list_frame, selectmode="extended", height=30, yscrollcommand=scrollbar.set)
list_file.pack(side="left", fill="both", expand=True)
scrollbar.config(command=list_file.yview)



# 진행 상황 Progress Bar
frame_progress = LabelFrame(root, text="진행상황")
frame_progress.pack(fill="x", padx=5, pady=5, ipady=5)

label_progress = Label(frame_progress, text="", anchor="w")
label_progress.pack(fill="x")

p_var = DoubleVar()
progress_bar = ttk.Progressbar(frame_progress, maximum=100, variable=p_var)
progress_bar.pack(fill="x", padx=5, pady=5)


# 실행 프레임
frame_run = Frame(root)
frame_run.pack(fill="x", padx=5, pady=5)

btn_close = Button(frame_run, padx=5, pady=5, text="닫기", width=12, command=root.quit)
btn_close.pack(side="right", padx=5, pady=5)

btn_stop = Button(frame_run, padx=5, pady=5, text="중지", width=12, command=stop)
btn_stop.pack(side="right", padx=5, pady=5)

btn_start = Button(frame_run, padx=5, pady=5, text="시작", width=12, command=start)
btn_start.pack(side="right", padx=5, pady=5)

data = info.read_options()

def save(saveContents):
    values = []
    index = 0
    for chkbox in saveContents["SIDE"]:
        if chkbox.get() == 1:
            values.append(index)
        index += 1
    
    data = {}    
    data['side_picture_except_index'] = values   
    data["NEW_NORMAL"] =  saveContents["NEW_NORMAL"].get()
    data["NEW_OPEN_CVD"] = saveContents["NEW_OPEN_CVD"].get()
    data["NEW_NEW_NORMAL"] = saveContents["NEW_NEW_NORMAL"].get()
    save_options(data)
        
def save_options(data):
    with open(info.path_sample_option, "w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False)
        

saveContents = {}

def setting_options():
    popup= Toplevel(root)
    popup.geometry("800x400")
    popup.title("세부 옵션 설정")
    
    frame_option = LabelFrame(popup, text="옵션 1")
    frame_option.pack(padx=5, pady=5, ipady=5, fill='x')

    
    # 1. [Side]제외 사진 선택 옵션
    lbl_side = Label(frame_option, text="[Side] 제외 사진 선택 : ")
    lbl_side.pack(side="left", padx=5, pady=5)


    frame_option_2 = LabelFrame(popup, text="옵션 2")
    frame_option_2.pack(padx=5, pady=10, ipady=5, fill='x')
    
    lbl_new_normal = Label(frame_option_2, text="[신규 업체명 - NORMAL] : ")
    lbl_new_normal.pack(side="left",padx=5,pady=5)
    e1 = Entry(frame_option_2, width = 100)
    e1.pack(side="left", padx=5, pady=5)
    e1.insert(0, data["NEW_NORMAL"] if "NEW_NORMAL" in data else "업체명을 입력하세요. (예:성산,아성)" )
    
    
    frame_option_3 = LabelFrame(popup, text="옵션 3")
    frame_option_3.pack(padx=5, pady=10, ipady=5, fill='x')
    
    lbl_new_open_cvd = Label(frame_option_3, text="[신규 업체명 - OPEN,CVD] : ")
    lbl_new_open_cvd.pack(side="left",padx=5,pady=5)
    e2 = Entry(frame_option_3, width = 100)
    e2.pack(side="left", padx=5, pady=5)
    e2.insert(0, data["NEW_OPEN_CVD"] if "NEW_OPEN_CVD" in data else "업체명을 입력하세요. (예:세우,풍원,핌스)")

    frame_option_4 = LabelFrame(popup, text="옵션 4")
    frame_option_4.pack(padx=5, pady=10, ipady=5, fill='x')
    
    lbl_new_new_normal = Label(frame_option_4, text="[신규 업체명 - NEW NORMAL] : ")
    lbl_new_new_normal.pack(side="left",padx=5,pady=5)
    e3 = Entry(frame_option_4, width = 100)
    e3.pack(side="left", padx=5, pady=5)
    e3.insert(0, data["NEW_NEW_NORMAL"] if "NEW_NEW_NORMAL" in data else "업체명을 입력하세요. (예:세우,풍원,핌스)")
    
    max_pic = 14
    chkvars = [] 
    chkboxes = []
    for i in range(max_pic):
        chkvars.append(IntVar())
        chkboxes.append(Checkbutton(frame_option, text=str(i+1), variable=chkvars[i]))
        chkboxes[i].pack(side="left", padx=3)
        
        # global data
        if (i) in data['side_picture_except_index']:
            chkboxes[i].select()
            
    
    saveContents["NEW_NORMAL"] = e1
    saveContents["NEW_OPEN_CVD"] = e2
    saveContents["NEW_NEW_NORMAL"] = e3
    saveContents["SIDE"] = chkvars
    
    # 버튼s
    frame_btn = Frame(popup)
    frame_btn.pack(fill="x", padx=5, pady=5)
    
    btn_close = Button(frame_btn, padx=5, pady=5, text="닫기", width=12, command=popup.destroy)
    btn_close.pack(side="right", padx=5, pady=5)
    
    btn_save = Button(frame_btn, padx=5, pady=5, text="저장", width=12, command=lambda: save(saveContents))
    btn_save.pack(side="right", padx=5, pady=5)
       

   
   

btn_option = Button(frame_run, padx=5, pady=5, text="옵션 설정", width=12, command=setting_options)
btn_option.pack(side="left", padx=5, pady=5)

root.mainloop()



