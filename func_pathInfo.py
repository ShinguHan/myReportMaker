import os
import tkinter.messagebox as msgbox
import json

# 성적서 작성 템플릿 경로 및 파일 명


# 복구 필요
# if os.path.exists('\\\\192.168.110.45\\라인컴퓨터'):
#     path_sample_root = '\\\\192.168.110.45\\라인컴퓨터\\ReportAutomation'
# else:
#     path_sample_root = 'E:\\라인컴퓨터\\Samples'
path_sample_root = 'E:\\라인컴퓨터\\Samples'
path_sample_option = os.path.join(path_sample_root, 'options.json')

def read_options():  
    # Opening JSON file
    try:
        with open(path_sample_option, encoding='UTF-8') as file:
            data = json.load(file)
        
        return data
    except:
        msgbox.showerror('에러',f'다음 경로에 options.json 파일이 없습니다.\n경로:{path_sample_option}') 

    
path_sample_files = ['sample_depo_A3_9F.pptx',    # 증착 A3
                    'sample_depo_A4.pptx',        # 증착 A4
                    'sample_new_normal_A3.pptx',  # 신규 A3(성산, 아성)
                    'sample_new_normal_A4.pptx',  # 신규 A4(성산, 아성)          
                    'sample_new_open_A3.pptx',    # 신규 오픈 A3(세우, 풍원, 핌스)
                    'sample_new_open_A4.pptx',    # 신규 오픈 A4(세우, 풍원, 핌스)
                    'sample_new_cvd_A3.pptx',     # 신규 CVD A3(세우, 풍원, 핌스)
                    'sample_new_cvd_A4.pptx']     # 신규 CVD A4(세우, 풍원, 핌스)

comInfo = read_options()

new_open_cvd_com_lists = comInfo["NEW_OPEN_CVD"].split(',') if "NEW_OPEN_CVD" in comInfo else ['핌스','세우','풍원']
new_open_cvd_lists = ['OPEN','CVD']
new_normal_com_lists = (comInfo["NEW_NORMAL"]).split(',') if "NEW_NORMAL" in comInfo else ['성산','아성']

# 복구 필요
# if os.path.exists('\\\\192.168.110.45\\라인컴퓨터'):
#     path_root = '\\\\192.168.110.45\\라인컴퓨터\\20'
# else:
#     path_root = 'E:\\라인컴퓨터\\20'

path_root = 'E:\\라인컴퓨터\\20'

path_year = '년'
path_month = '월'
path_day = '일'
path_save_depo = '1. 증착'

path_depo_eye_1 = '년\\1. 증착\\목시\\'
path_depo_eye_2 = '월\\'
path_depo_eye_A3 = '일\\A3\\세정후'
path_depo_eye_A4 = '일\\A4\\세정후'

path_depo_vision_1 = '년\\1. 증착\\비전\\'
path_depo_vision_2 = '월\\'
path_depo_vision_A3 = '일\\A3\\'
path_depo_vision_A4 = '일\\A4\\'

front = '\\front'
side = '\\side'
hole = '\\hole'
inspect = '\\inspect'

path_new_eye_1 = '년\\2. 신규\\목시\\'
path_new_eye_2 = '월\\'
path_new_eye_3 = '일\\'
# +성산, 아성
path_new_eye_normal_A3 = '\\A3\\세정후'
path_new_eye_normal_A4 = '\\A4\\세정후'

path_new_vision_1 = '년\\2. 신규\\비전'
path_new_vision_2 = '월'
path_new_vision_3 = '일'

#복구 필요
# if os.path.exists('\\\\192.168.110.45\\라인컴퓨터'):
#     path_save_root = '\\\\192.168.110.45\\01. 세정팀) 성적서\\20'
# else:
#     path_save_root = 'E:\\01. 세정팀) 성적서\\20'

path_save_root = 'E:\\01. 세정팀) 성적서\\20'

path_year = '년'
path_month = '월'
path_day = '일'
path_save_depo = '1. 증착'
path_save_new = '2. 신규'




def get_path_save_depo(year, month, day, detail, filetype):
    tempPath = os.path.join(path_save_root+year+path_year,path_save_depo,month+path_month,day+path_day, detail, filetype)
    
    if not os.path.exists(tempPath):
            os.makedirs(tempPath)
            
    return tempPath

def get_path_save_new(year, month, day, company, detail, filetype):
    tempPath = os.path.join(path_save_root+year+path_year,path_save_new,month+path_month,day+path_day, company, detail, filetype)
    
    if not os.path.exists(tempPath):
            os.makedirs(tempPath)
            
    return tempPath

def get_path_save_new_openOrcvd(year, month, day, company, detail, openOrcvd, filetype):
    tempPath = os.path.join(path_save_root+year+path_year,path_save_new,month+path_month,day+path_day, company, detail, openOrcvd, filetype)
    
    if not os.path.exists(tempPath):
            os.makedirs(tempPath)
            
    return tempPath

def get_path_new_vision(year, month, day, company, id, detail):
    return os.path.join(path_root + year+path_new_vision_1,month + path_new_vision_2,day + path_new_vision_3,company,id, detail)
    
path_new_eye_open_cvd_A3 = '\\A3\\'
path_new_eye_open_cvd_A4 = '\\A4\\'

path_new_eye_open_cvd_after = '\\세정후'

path_save_ppt = path_sample_root

def get_working_id_lists(year, month, day):
    temp_lists =[]
    lists = []
    
    
    # 증착 A3
    path = get_path_depo_eye_A3(year, month, day)
    print(path)
    temp_lists = search_files(path)
    if temp_lists != None:
        for file in temp_lists:
            lists.append(path + '\\' + file)            
    
    # 증착 A4
    path = get_path_depo_eye_A4(year, month, day)
    temp_lists = search_files(path)
    if temp_lists != None:    
        for file in temp_lists:
            lists.append(path + '\\' + file)
    
    # 신규 (성산, 아성) A3            
    for com in new_normal_com_lists:
        path = get_path_new_eye_A3(year, month, day, com)
        temp_lists = search_files(path)
        if temp_lists != None:    
            for file in temp_lists:
                lists.append(path + '\\' + file)
                
    # 신규 (성산, 아성) A4
    for com in new_normal_com_lists:
        path = get_path_new_eye_A4(year, month, day, com)
        temp_lists = search_files(path)
        if temp_lists != None:    
            for file in temp_lists:
                lists.append(path + '\\' + file)
                
    # 신규 (풍원, 세우, 핌스) A3
    for com in new_open_cvd_com_lists:
        for oc in new_open_cvd_lists:
            path = get_path_new_open_cvd_A3(year, month, day, com, oc)
            temp_lists = search_files(path)
            if temp_lists != None:    
                for file in temp_lists:
                    lists.append(path + '\\' + file)
                    
    # 신규 (풍원, 세우, 핌스) A4
    for com in new_open_cvd_com_lists:
        for oc in new_open_cvd_lists:
            path = get_path_new_open_cvd_A4(year, month, day, com, oc)
            temp_lists = search_files(path)
            if temp_lists != None:    
                for file in temp_lists:
                    lists.append(path + '\\' + file)
    
    return lists
    
    


def get_path_depo_eye_A3(year, month, day):
    path = path_root + str(year) + path_depo_eye_1 + str(month) + path_depo_eye_2 + str(day) + path_depo_eye_A3
    return path

def get_path_depo_eye_A4(year, month, day):
    path = path_root + str(year) + path_depo_eye_1 + str(month) + path_depo_eye_2 + str(day) + path_depo_eye_A4
    return path

def get_path_new_eye_A3(year, month, day, company):
    path = path_root + str(year) + path_new_eye_1 + str(month) + path_new_eye_2 + str(day) + path_new_eye_3 + company + path_new_eye_normal_A3
    return path

def get_path_new_eye_A4(year, month, day, company):
    path = path_root + str(year) + path_new_eye_1 + str(month) + path_new_eye_2 + str(day) + path_new_eye_3 + company + path_new_eye_normal_A4
    return path

def get_path_new_open_cvd_A3(year, month, day, company, open_cvd):
    path = path_root + str(year) + path_new_eye_1 + str(month) + path_new_eye_2 + str(day) + path_new_eye_3 + company + path_new_eye_open_cvd_A3 + open_cvd + path_new_eye_open_cvd_after
    return path

def get_path_new_open_cvd_A4(year, month, day, company, open_cvd):
    path = path_root + str(year) + path_new_eye_1 + str(month) + path_new_eye_2 + str(day) + path_new_eye_3 + company + path_new_eye_open_cvd_A4 + open_cvd + path_new_eye_open_cvd_after
    return path


def get_path_depo_vision_A3(year, month, day):
    path = path_root + str(year) + path_depo_vision_1 + str(month) + path_depo_vision_2 + str(day) + path_depo_vision_A3
    return path

def get_path_depo_vision_A4(year, month, day):
    path = path_root + str(year) + path_depo_vision_1 + str(month) + path_depo_vision_2 + str(day) + path_depo_vision_A4
    return path

def get_path_front(prefix):
    path = prefix + front
    return path

def get_path_side(prefix):
    path = prefix + side
    return path

def get_path_hole_depo_vision_A3(id, year, month, day):
    path = get_path_depo_vision_A3(year, month, day) + id + hole
    return path

def get_path_hole_depo_vision_A4(id, year, month, day):
    path = get_path_depo_vision_A4(year, month, day) + id + hole
    return path

def get_path_inspect_depo_vision_A3(id, year, month, day):
    path = get_path_depo_vision_A3(year, month, day) + id + inspect
    return path

def get_path_inspect_depo_vision_A4(id, year, month, day):
    path = get_path_depo_vision_A4(year, month, day) + id + inspect
    return path

def get_path_sample(cat):
# path_sample_files = ['sample_depo_A3_9F.pptx',    # 증착 A3
#                     'sample_depo_A4.pptx',        # 증착 A4
#                     'sample_new_normal_A3.pptx',  # 신규 A3(성산, 아성)
#                     'sample_new_normal_A4.pptx',  # 신규 A4(성산, 아성)          
#                     'sample_new_open_A3.pptx',    # 신규 오픈 A3(세우, 풍원, 핌스)
#                     'sample_new_open_A4.pptx',    # 신규 오픈 A4(세우, 풍원, 핌스)
#                     'sample_new_cvd_A3.pptx']  # 신규 CVD A3(세우, 풍원, 핌스)
#                     'sample_new_cvd_A4.pptx']  # 신규 CVD A4(세우, 풍원, 핌스)
    if cat == 'depo_A3':
        path = path_sample_root + '\\templeteSamples\\' + path_sample_files[0]
    elif cat == 'depo_A4':        
        path = path_sample_root + '\\templeteSamples\\' + path_sample_files[1]
    elif cat == 'new_normal_A3':        
        path = path_sample_root + '\\templeteSamples\\' + path_sample_files[2]
    elif cat == 'new_normal_A4':        
        path = path_sample_root + '\\templeteSamples\\' + path_sample_files[3]
    elif cat == 'new_open_A3':        
        path = path_sample_root + '\\templeteSamples\\' + path_sample_files[4]
    elif cat == 'new_open_A4':
        path = path_sample_root + '\\templeteSamples\\' + path_sample_files[5]
    elif cat == 'new_cvd_A3':
        path = path_sample_root + '\\templeteSamples\\' + path_sample_files[6]
    elif cat == 'new_cvd_A4':
        path = path_sample_root + '\\templeteSamples\\' + path_sample_files[7]
    else:
        msgbox.msgbox.showerror("에러", "잘못된 샘플 목록을 입력하였습니다")    
        
    return path


def search_files(path):
    try:
        return os.listdir(path=path)
    except:
        pass
    

