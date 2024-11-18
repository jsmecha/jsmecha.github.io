#import sys
import asyncio
import psutil
import win32gui
import win32process
import time
from datetime import datetime, timedelta
import keyboard
import socket
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import logging

#sys.stdout.reconfigure(encoding='utf-8')

File_info = {} #업데이트 할 파일 정보를 저장
File_name = 'application_log.txt'
Folder_id = "1b0mUpjU7sVB7RBvwyUUDkxQeXsE8ClWC" # 변경 금지. google drive 내의 AppLogger 폴더임


Except_list = ['explorer.exe', 'SearchHost.exe', 'Tastmgr.exe', 'cmd.exe', 'ShellExperienceHost.exe', 'StartMenuExperienceHost.exe']

logging.basicConfig(filename="AppLogger4Windows.log", format='%(asctime)s %(levelname)s:%(message)s',  
    datefmt='%m/%d/%Y %I:%M:%S %p', level=logging.DEBUG)



#Google Drive 인증 설정
def authenticate():
    gauth = GoogleAuth()
    gauth.LoadCredentialsFile('credentials.json')

    if gauth.credentials is None:
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        try:
            gauth.Refresh() #--> 여기서 Refresh 안되고 exception 발생. token access fail
        except:
            gauth.credentials = None
            gauth.LocalWebserverAuth()
    else:
        gauth.Authorize()

    gauth.SaveCredentialsFile('credentials.json')
    return gauth

def get_active_window_name():    
    window_handle = win32gui.GetForegroundWindow() # 최상위 창 핸들 가져오기
    window_title = win32gui.GetWindowText(window_handle) # 가져온 핸들의 창 이름(프로세스 이름) 
    _,process_id = win32process.GetWindowThreadProcessId(window_handle)
    
    if process_id >= 0:        
        try:
            process = psutil.Process(process_id)
            process_name = process.name()

            if process_name not in Except_list:
                start_time = time.ctime(process.create_time())
                return process_name, window_title , start_time
        except(psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    
    
    return None,None,None

def upload_file_to_GoogleDrive(drive, file_name):    
    file_drive = drive.CreateFile({'title':file_name, 'parents':[{'id':Folder_id}]})
    file_drive.SetContentFile(file_name)
    file_drive.Upload()
    print("File uploaded to Google Drive")
    logging.debug("File uploaded to Google Drive")
    return file_drive

def update_file_to_GoogelDrive( file_name):     
    file_drive = File_info
    file_drive.SetContentFile(file_name)
    file_drive.Upload()
    print("File updated on Google Drive")
    logging.debug("File updated to Google Drive")

def delete_old_file_from_drive(drive, cut_off_days):
    file_list = drive.ListFile({'q':f"'{Folder_id}' in parents and trashed=false"}).GetList()
    cutoff_time = datetime.now() - timedelta(days = cut_off_days)
    for file in file_list:
        file_created_time = datetime.strptime(file['createdDate'], '%Y-%m-%dT%H:%M:%S.%fZ')

        if file_created_time < cutoff_time:
            file.Delete()
            logging.debug("File deleted to Google Drive {0}".format(file['title']))

    

async def monitor_and_upload():
    global File_info
    last_app = None

    drive_connect = True

    try:
        gauth = authenticate()
        gdrive = GoogleDrive(gauth)

        delete_old_file_from_drive(gdrive, 7)   
    
    except socket.error as e:
        drive_connect = False
        logging.debug("monitor_and_upload() socket.error")

    except Exception as e:
        drive_connect = False
        print(e) # 네트워크 연결 문제는 파일저장하면 되므로 일단 그냥 넘어가기.
        logging.debug("monitor_and_upload() exception {0}".format(e))


    while True:
        if keyboard.is_pressed('esc'):
            print('ESC pressed. exiting the program')
            logging.debug("ESC pressed.exiting the program")
            break
        
        current_app,window_title,start_time = get_active_window_name()
       
        if current_app and current_app != last_app:
            global File_info

            #print(f"Application:{current_app}, Window Title : {window_title}, Start Time:{start_time}")
            cur_date = time.strftime("%Y%m%d")
            file_name = cur_date + "_" + File_name
            with open(file_name, "a", encoding='utf-8') as file:
                file.write(f"{time.ctime()}-{current_app} [{window_title}] started at {start_time}\n")                

                try:
                    if drive_connect:
                        if File_info == {} or File_info['title'] != file_name: # 빈거면, 프로그램 실행 후 별도로 파일 정보 찾아보지 않음.
                            query = f"title = '{file_name}' and '{Folder_id}' in parents and trashed = false"
                            file_list = gdrive.ListFile({'q': query}).GetList()
                        
                            if len(file_list) > 0:
                                File_info = file_list[0] #해당 파일 존재함
                                update_file_to_GoogelDrive(file_name)
                            else:
                                File_info = upload_file_to_GoogleDrive(gdrive, file_name)                    
                        else:
                            update_file_to_GoogelDrive(file_name)
                                
                except socket.error as e:
                    drive_connect = False  
                    logging.debug("monitor_and_upload() socket.error")                  

                except Exception as e:
                    drive_connect = False
                    print(e)
                    logging.debug("monitor_and_upload() exception {0}".format(e))
                
            last_app = current_app

        await asyncio.sleep(30)


if __name__ == "__main__":  
    logging.debug("start program")
    asyncio.run(monitor_and_upload())
    
   
