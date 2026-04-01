#Relation of function
"""-------------------------------------------------------------------------------------------------------
Diagnosis:-----:set_main:-----:easy_set:-----:start(0):-----:measurement(0):*
                           |--:developmet:-----:parts_1st:*
                           |--:ev_fit:*     |--:start(1):-----:measurement(1):-----:eraser_diagnosis:*
                           |--:get_cons:*   |--:stop:*                          |--:parts_1st:*
                                            |--:record:*
                                            |--:calibration:-----:eraser_measurement:
                                            |                 |--:parts_2st:-----:setup:*
                                            |                                 |--:renew:*
                                            |                                 |--:flash:*
                                            |                                 |--:see_key:*
                                            |                                 |--:see_pass:*
                                            |                                 |--:get_device:*
                                            |                                 |--:add_data(a):*
                                            |                                 |--:add_data(w):*
                                            |                                 |--:ch_aprx:-----:inspect:-----:function_model:*
                                            |--:select_device:-----:option_device:-----:reg:*
                                            |--:select_comb:*
                                            |--:calibration_curve:-----:inspect:-----:function_model:*                                                   
----------------------------------------------------------------------------------------------------------"""
#Recommended folder structure
"""
esp_azure
txt(devicename.txt,measurement_data_for_GUI.xlsx,req.txt)
image(image1.png,image2.png,image3.png,image4.png,image5.png,image6.png,image7.png)
setuper(azure-cli-2.65.0-x64.msi,esp-idf-tools-setup-offline-5.3.exe,Git-2.46.2-32-bit.exe)
"""
#Header files (You need to run the pip commnad as required) 
import tkinter as tk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from tkinter import ttk
import subprocess
import re
import threading
import numpy as np
from scipy.interpolate import make_interp_spline
import math
from scipy.optimize import curve_fit
from sklearn.metrics import r2_score
import pandas as pd
from PIL import Image, ImageTk
from tkinter import scrolledtext
import os
import shutil

#Globals that are not defined at each function 
global mode,record_before,dev
global wcount,acount
wcount=0
acount=0
dev,mode,record_before,record_count,z=0,0,0,0,0
#Semi globals that are not defined at each function
wavelength=[410,435,460,485,510,535,560,585,610,645,680,705,730,760,810,860,900,940]
alphbet=["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE"]
log_con=[]
log_abs=[]
log_abs.append(0)
log_con.append(0)
devicename_path="txt\devicename.txt"
xlsx_path="txt\measurement_data_for_GUI.xlsx"
#Function for deleting wasting data from azure commnication 
def delete_wasting_characters(a):
    a = re.sub(r'^.*?:', '', a)
    pattern = r':(.*?),'
    matches = re.findall(pattern, a)
    result_background_list = [int(match) for match in matches]
    return result_background_list
#Register Button
def reg():
    with open(devicename_path,"r") as f:
        lines=f.readlines()
    del lines[len(lines)-1]
    with open(devicename_path, "w") as f:
        new_device = reg_txt.get()
        f.writelines(lines)
        f.write(new_device + "\nend")
    reg_txt.destroy()
    reg_bot.destroy()
#Initialize,Register,Delete
def option_device(event):
    global result_background_list, process, mes, devices, combobox_option, reg_txt, reg_bot
    if combobox_option.get() == "Initialization":
        # Read spectram data from Azure Hub
        mes = f"az iot hub monitor-events --device-id {devices} --hub-name MeniconPilotHub"
        process = subprocess.Popen(mes, stdout=subprocess.PIPE, shell=True)
        line = process.stdout.readline().decode('utf-8').rstrip()
        print(line)
        # Read initial spectram data from Azure Hub for background data
        for i in range(9):
            line = process.stdout.readline().decode('utf-8').rstrip()
            if i == 6:
                a = line
        result_background_list = delete_wasting_characters(a)
        #changing device order
        with open(devicename_path,"r") as file:
            lines = file.readlines()
        for i in range(len(lines)):
            if lines[i].rstrip("\n")==devices:
                break
            else:
                i+=1
        line_to_move = lines.pop(i)  
        lines.insert(0, line_to_move) 
        # re-writing file
        with open(devicename_path, "w") as file:
            file.writelines(lines)

    elif combobox_option.get() == "Deletion":
        with open(devicename_path, "r") as f:
            read_delete = [line.strip() for line in f if line.strip() != "end"]
        with open(devicename_path, "w") as f:
            f.writelines(line + "\n" for line in read_delete if line != devices)
            f.write("end\n")
    elif combobox_option.get() == "Registration":
        reg_txt = tk.Entry(root, width=6, justify='right')
        reg_txt.place(x=230, y=305)
        reg_bot = tk.Button(root, text="OK", command=reg)
        reg_bot.place(x=280, y=305)
#Select azure device
def select_device(event):
    global devices, combobox_option
    devices = combobox_device.get().strip()
    option_list = ["Initialization", "Deletion", "Registration"]
    combobox_option = ttk.Combobox(root, values=option_list, width=10)
    combobox_option.bind('<<ComboboxSelected>>', option_device)
    combobox_option.place(x=140, y=305)
#Prepare spectral window
def parts_1st():
    global fig,ax,title,canvas_graph,canvas_background_show,lblsubtitle,lbl,wave_lbl,concentlbl,choicelbl,txtleft,txtright,txt,absorbancetxt,selectwavetxt
    root.geometry('1280x720')
    fig = Figure(figsize=(11.5, 7), facecolor="#ACACAC")
    fig.subplots_adjust(hspace=0.5)
    ax = []
    title = ['Background data', 'Current measured data', 'Absorbed light data', 'Absorbance data']
    for i in range(3):
        ax.append(fig.add_subplot(3, 2, i * 2 + 1))
        ax[i].set_title(title[i])
    ax.append(fig.add_subplot(3, 2, 6))
    ax[3].set_title(title[3])
    
    canvas_background_show = tk.Canvas(root, bg="#ACACAC", width=980, height=670, borderwidth=-2)
    canvas_background_show.place(x=300, y=57)
    canvas_graph = FigureCanvasTkAgg(fig, master=canvas_background_show)
    canvas_graph_widget = canvas_graph.get_tk_widget()
    canvas_background_show.create_window(500, 320, window=canvas_graph_widget)
    
    #Output the wavelength txt
    lblsubtitle = tk.Label(text='Absorbance',bg='#ACACAC', font=("MSゴシック","20"))
    lblsubtitle.place(x=830, y=100)
    #Show wavelength txt
    lbl=[]
    for i in range(18):
        wave_lbl=[]
        wave_lbl.append(wavelength[i])
        wave_lbl.append("nm")
    
        if i<5:
            lbl.append(tk.Label(text=wave_lbl, bg='#ACACAC'))
            lbl[i].place(x=818+i*80,y=150)
        elif 4<i<10:
            lbl.append(tk.Label(text=wave_lbl, bg='#ACACAC'))
            lbl[i].place(x=818+(i-5)*80,y=190)
        elif 9<i<15:
            lbl.append(tk.Label(text=wave_lbl, bg='#ACACAC'))
            lbl[i].place(x=818+(i-10)*80,y=230)  
        elif 14<i<18:
            lbl.append(tk.Label(text=wave_lbl, bg='#ACACAC'))
            lbl[i].place(x=818+(i-15)*80,y=270)  
    #Show texts      
    concentlbl = tk.Label(text='Expected Protein concentration', background='#ACACAC',font=("Helvetica","14"))
    concentlbl.place(x=818, y=345)
    
    choicelbl = tk.Label(text='Select wavelength',background='#ACACAC',font=("Helvetica","14"))
    choicelbl.place(x=818,y=305)
    #Meke thing to show absorbance and wavelength 
    txtleft = tk.Entry(width=10, justify='right',background='#ACACAC')
    txtleft.place(x=1116, y=270)
    txtright = tk.Entry(width=6, justify='right')
    txtright.place(x=1180, y=270)
    #Make things to show each absorbance 
    txt=[]
    for i in range(18):
        if i<5:
            txt.append(tk.Entry(width=6, justify='right'))
            txt[i].place(x=860+i*80,y=150)
        elif 4<i<10:
            txt.append(tk.Entry(width=6, justify='right'))
            txt[i].place(x=860+(i-5)*80,y=190)
        elif 9<i<15:
            txt.append(tk.Entry(width=6, justify='right'))
            txt[i].place(x=860+(i-10)*80,y=230)
        elif 14<i<18:
            txt.append(tk.Entry(width=6, justify='right'))
            txt[i].place(x=860+(i-15)*80,y=270)
    
    absorbancetxt = tk.Entry(width=7, justify='right')
    absorbancetxt.place(x=1100, y=350)
    
    selectwavetxt = tk.Entry(width=7, justify='right')
    selectwavetxt.place(x=985, y=310)
#Hide password
def see_pass():
    global pass_count
    if pass_count==1:
        Password_txt.config(show="")
        pass_count=0
    elif pass_count==0:
        Password_txt.config(show="*")
        pass_count=1
#Hide key
def see_key():
    global key_count
    if key_count==1:
        key_txt.config(show="")
        key_count=0
    elif key_count==0:
        key_txt.config(show="*")
        key_count=1
#Install ESP-IDF, Azure CLI, Git
def setup():
    # idf installer path-----------------------------------------------------------------------------------
    idf_path = r"setuper\esp-idf-tools-setup-offline-5.3.exe"
    # Git_path installer path------------------------------------------------------------------------------
    Git_path = r"setuper\Git-2.46.2-32-bit.exe"
    # Azure CLI installer path-----------------------------------------------------------------------------
    CLI_path = r"setuper\azure-cli-2.65.0-x64.msi"
    result = subprocess.run(['wmic', 'product', 'get', 'name'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    result_V = result.stdout.split("\n")
    CLI_count=0
    for app in result_V:
        app = app.strip()  
        if app == "Microsoft Azure CLI (64-bit)":
            CLI_count=1
            break
    IDF_count=0
    Git_count=0
    if os.path.isdir("C:/ESPRESSIF"):
        IDF_count=1
    if os.path.isdir("C:/Program Files (x86)/Git"):
        Git_count=1

    if IDF_count==0:
        text_box.insert(tk.END,"ESP-IDF environment is being installed\n")
        # idf install commnad by silent-------------------------------------------------------------------
        subprocess.run([idf_path, "/verysilent"])
        text_box.insert(tk.END,"Installing is finished\n")
    elif IDF_count==1:
        text_box.insert(tk.END,"ESP-IDF has been installed already\n")
    if Git_count==0:
        text_box.insert(tk.END,"Git environment is being installed\n")
        subprocess.run([Git_path, "/verysilent"])
        text_box.insert(tk.END,"Installing is finished\n")
    elif Git_count==1:
        text_box.insert(tk.END,"Git has been installed already\n")
    if CLI_count==0:
        text_box.insert(tk.END,"Azure CLI environment is being installed\n")
        subprocess.run([CLI_path,"/verysilent"])
    elif CLI_count==1:
        text_box.insert(tk.END,"Azure CLI has been installed it already\n")
    current=os.getcwd()
    src_folder=current+"\esp_azure"
    dst_folder=r"C:\Espressif\frameworks\esp-idf-v5.3\examples\get-started"
    shutil.copytree(src_folder, dst_folder,dirs_exist_ok=True)
#Get azure device name    
def get_device():
    #logout for securely login to azure-------------------------------------------------------------------
    subprocess.run(['powershell','-Command','az logout'],capture_output=True, text=True, check=True)
    #login for using azure CLI----------------------------------------------------------------------------
    subprocess.run(['powershell','-Command','az login'],capture_output=True, text=True, check=True)
    text_box.insert(tk.END,'Login is succesful\n')
    # get device names------------------------------------------------------------------------------------
    result = subprocess.run(['powershell', '-Command', 'az iot hub device-identity list --hub-name MeniconPilotHub --query "[].deviceId" --output tsv'], capture_output=True, text=True, check=True)
    
    if result.returncode == 0:
        result_id = result.stdout.strip() 
        result_id_list=result_id.split("\n")
        device_comb['values']=result_id_list
        text_box.insert(tk.END,'Device names are gotten\n')
    else:
        text_box.insert(tk.END,f'Error{result.stderr}\n')
#Re-writing each requirement(Wi-Fi:{SSID, password}, Azure IoT Hub{device name, device key}) 
def renew():
    def replace_in_file(file_path, search_text, replace_text):
        try:
            #file open for read mode---------------------------------------------------------------------
            for i in range(4):
                with open(file_path, 'r', encoding='utf-8') as file:
                    content = file.read()
                re_sdk=content.replace(search_text[i],replace_text[i])
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(re_sdk)
    
        except FileNotFoundError:
            print(f"'{file_path}' is not found")
        except Exception as e:
            print(f"error occur: {e}")
    req_path=r'txt\req.txt'
    with open(req_path,'r',encoding='utf-8') as file:
        req=[]
        for i in range(4):
            req.append(file.readline().split("\n")[0])
    error_count=0
    SSID=SSID_txt.get()
    if SSID=="":
        error_count+=1
    Password=Password_txt.get()
    if Password=="":
        error_count+=1
    device=device_comb.get()
    if device=="":
        error_count+=1
    # get device key-------------------------------------------------------------------------------------
    if error_count==0:
        result = subprocess.run(['powershell', '-Command', 'az iot hub device-identity show --hub-name MeniconPilotHub --device-id '+device+ ' --query "authentication.symmetricKey.primaryKey" --output tsv'], capture_output=True, text=True, check=True)
        if result.returncode == 0:
            result_key = result.stdout.strip()  
            key=result_key
            key_txt.insert(tk.END,key)
            key_txt.config(show="*")
            replace_text = [SSID,Password,device,key] 
            for i in range(4):
                with open(req_path,'r',encoding='utf-8') as file:
                    req_main=file.read()
                context=req_main.replace(req[i],replace_text[i])
                with open(req_path, 'w', encoding='utf-8') as file:
                    file.write(context)
            
            file_path = r'C:\Espressif\frameworks\esp-idf-v5.3\examples\get-started\iot-middleware-freertos-samples\demos\projects\ESPRESSIF\aziotkit\sdkconfig'
            search_text = req   
            replace_in_file(file_path, search_text, replace_text)
            Password_hided="Password hided"
            Key_hided="key hided"
            text_box.insert(tk.END, f"Finished re-write each requirements{req[0],Password_hided,req[2],Key_hided} -> {SSID,Password_hided,device,Key_hided}\n")
        else:
            text_box.insert(tk.END,f'Error:{result.stderr}\n')
    else:
        text_box.insert(tk.END,'Please input three requirement\n')
#Flash program to ESP          
def flash():
    def run_flash_process():
        try:
            # save current directory path-----------------------------------------------------------------
            original_dir = os.getcwd()
            
            # Add git path at environment variable--------------------------------------------------------
            new_path = r'C:\Program Files (x86)\Git\bin'
            os.environ['PATH'] += os.pathsep + new_path
            print(os.environ['PATH'])
            
            # remove build folder-------------------------------------------------------------------------
            folder_path = r'C:\espbuild'
            if os.path.exists(folder_path):
                shutil.rmtree(folder_path)
                print(f"{folder_path} is erased")
            else:
                print(f"{folder_path} is not found")
            
            # ESP bat path--------------------------------------------------------------------------------
            esp_idf_setup = r'C:\Espressif\frameworks\esp-idf-v5.3\export.bat'
            
            # ESP path------------------------------------------------------------------------------------
            project_dir = r'C:\Espressif\frameworks\esp-idf-v5.3\examples\get-started\iot-middleware-freertos-samples\demos\projects\ESPRESSIF\aziotkit'
            
            # move project directory----------------------------------------------------------------------
            os.chdir(project_dir)
            com = com_txt.get()
            com_number = '-p ' + com
            command_base = ' && idf.py --no-ccache -B "C:\\espbuild" '
            command_erase = command_base + com_number + ' erase-flash'
            command_build = command_base + 'build'
            command_flash = command_base + com_number + ' flash'
            
            # Show log------------------------------------------------------------------------------------
            text_box.insert(tk.END, 'Running flasher process...\n')
            text_box.insert(tk.END, 'Please wait and keep pushing the ESP button.\n')
            text_box.yview(tk.END)
            root.update() 
            
            # Run ESP commands----------------------------------------------------------------------------
            subprocess.run(
                ['cmd', '/c', esp_idf_setup + command_erase + command_build + command_flash],
                shell=True,
                check=True
            )
            
        except Exception as e:
            # Error message-------------------------------------------------------------------------------
            text_box.insert(tk.END, f"Error: {e}\n")
            text_box.yview(tk.END)
            root.update() 
        
        finally:
            os.chdir(original_dir)
            text_box.insert(tk.END, "Finished ESP flash\n")
            text_box.yview(tk.END)
            root.update()  
    
    # Running process in parallel-------------------------------------------------------------------------
    flash_thread = threading.Thread(target=run_flash_process)
    flash_thread.start()
#Adding measurement data to xlsx
def add_data(aorw):
    global wcount,acount
    df = pd.read_excel(xlsx_path)
    if aorw == "a":
        if mode == 0 or mode == 1:
            xlsx_i=4+acount
            xlsx_j=1
            xlsx_range=8
        elif mode == 2:
            xlsx_i=5+acount
            xlsx_j=11
            xlsx_range=9
        elif mode == 3:
            xlsx_i=5+acount
            xlsx_j=22
            xlsx_range=9
        if wcount!=0:
            xlsx_i=acount+2
        acount+=1
        print(acount)
    elif aorw == "w":
        xlsx_i=1
        if mode == 0 or mode == 1:
            xlsx_j=1
            xlsx_range=8
            for i in range(len(df)-4):
                for j in range(xlsx_range):
                    df.iat[i+xlsx_i,j+xlsx_j]=None

        elif mode == 2:
            xlsx_j=11
            xlsx_range=9
            for i in range(len(df)-4):
                for j in range(xlsx_range):
                    df.iat[i+xlsx_i,j+xlsx_j]=None
        elif mode == 3:
            xlsx_j=22
            xlsx_range=9
            for i in range(len(df)-4):
                for j in range(xlsx_range):
                    df.iat[i+xlsx_i,j+xlsx_j]=None
        wcount+=1
    for i in range(xlsx_range):
        df.iat[xlsx_i,i+xlsx_j]=log_abs[i+1]
    
    for i in range(8):
        df.iat[18,i+1]="=AVERAGE({a}3:{a}19)".format(a=alphbet[i+1])
    for i in range(9):
        df.iat[18,i+11]="=AVERAGE({a}3:{a}19)".format(a=alphbet[i+11])
    for i in range(9):
        df.iat[18,i+22]="=AVERAGE({a}3:{a}19)".format(a=alphbet[i+22])    
    df.to_excel(xlsx_path, index=False)
    #get average
    for i in range(xlsx_range):
        log_abs[i+1]=df.iat[18,i+1]
    message_txt.insert(tk.END,"Excel file is re-written\n")
#Make a model of least square method
def function_model(x,a,b,c,d,e):
    global z
    if z==0:
        return a*x
    if z==1:
        return a*x**2+b*x
    if z==2:
        return a*x**3+b*x**2+c*x
    if z==3:
        return a*x**4+b*x**3+c*x**2+d*x
    if z==4:
        return a*x**5+b*x**4+c*x**3+d*x**2+e*x
#Inspect calibration curve
def inspect(newxlist,newylist):
    global z,ylist,mode,coe

    if z != 0 and 'ylist' in globals():
        y_oldlist = ylist  
    else:
        y_oldlist = None
    
    # make medel
    params, covariance = curve_fit(function_model, newxlist, newylist)
    if mode == 0 or mode == 1:
        xlist = np.linspace(0, np.max(np.array(newxlist)), 101)
    elif mode == 2 or mode == 3:
        xlist = np.linspace(0, np.max(np.array(newxlist)), 2001)
    # prepare parameta
    coe = [param for param in params if param != 1]
    coe.reverse()
    coe += [0] * (6 - len(coe)) 
    
    # make approximated formula
    ylist = coe[5]*pow(xlist,6) + coe[4]*pow(xlist,5) + coe[3]*pow(xlist,4) + coe[2]*pow(xlist,3) + coe[1]*pow(xlist,2) + coe[0]*pow(xlist,1)

    # verify a monotonisity
    mono = 1 if all(ylist[i] <= ylist[i+1] for i in range(len(ylist) - 1)) else 2
    
    if mono == 1:
        z += 1
        if z == 5:
            print(z)
            y_sendlist = ylist
            ylist_R2=[]
            Index=newxlist
            for i in range(len(Index)):
                ylist_R2.append(y_sendlist[int(Index[i])])
            coe_det = r2_score(newylist,ylist_R2)
            return xlist, y_sendlist, coe_det
        else:
            return inspect(newxlist, newylist)
    elif mono == 2:
        print(z)
        y_sendlist = y_oldlist if y_oldlist is not None else ylist
        ylist_R2=[]
        Index=newxlist
        for i in range(len(Index)):
            ylist_R2.append(y_sendlist[int(Index[i])])
        coe_det = r2_score(newylist,ylist_R2)
        return xlist, y_sendlist, coe_det
#Changing prediction equation
def ch_aprx():
    df=pd.read_excel(xlsx_path)
    average_absorbance=[]
    based_concentration=[]
    if mode == 0 or mode == 1:
        range_length=8
        start_position=1
    elif mode == 2:
        range_length=9
        start_position=11
    elif mode == 3:
        range_length=9
        start_position=22
    for i in range(range_length):
        average_absorbance.append(df.iat[18,i+start_position])
        based_concentration.append(df.iat[19,i+start_position])
    inspect(average_absorbance,based_concentration)
    new_coe=[]
    for i in range(z):
        new_coe.append(coe[i])
    new_coe.reverse()
    if mode == 0 or mode == 1:
        df.iat[20,1]=z
        for i in range(z):
            df.iat[20,i+2]=round(new_coe[i],3)
    elif mode == 2:
        df.iat[20,11]=z
        for i in range(z):
            df.iat[20,i+12]=round(new_coe[i],3)
    elif mode == 3:
        df.iat[20,22]=z
        for i in range(z):
            df.iat[20,i+23]=round(new_coe[i],3)
    for i in range(8):
        df.iat[18,i+1]="=AVERAGE({a}3:{a}19)".format(a=alphbet[i+1])
    for i in range(9):
        df.iat[18,i+11]="=AVERAGE({a}3:{a}19)".format(a=alphbet[i+11])
    for i in range(9):
        df.iat[18,i+22]="=AVERAGE({a}3:{a}19)".format(a=alphbet[i+22])    
    df.to_excel(xlsx_path, index=False)
    message_txt.insert(tk.END,"Coefficient of calibration curve exchanges completes\n")
#Prepare calibration window
def parts_2nd():
    #Calibration curve globals
    global fig2,ax_cal,txt_gr_title,con_txt,abs_txt,canvas_graph
    #Flasher globals
    global pass_count,key_count,text_box,device_comb,SSID_txt,Password_txt,key_txt,com_txt
    pass_count=1
    key_count=1
    combo_init='NULL'
    #calibration_session---------------------------------------------------------------------------------
    fig2 = Figure(figsize=(6,5),facecolor="#ACACAC")
    ax_cal=fig2.add_subplot(1,1,1)
    canvas_background_show = tk.Canvas(root, bg="#ACACAC", width=980, height=670, borderwidth=-2)
    canvas_background_show.place(x=300, y=57)
    canvas_graph = FigureCanvasTkAgg(fig2, master=canvas_background_show)
    canvas_graph_widget = canvas_graph.get_tk_widget()
    canvas_background_show.create_window(300, 250, window=canvas_graph_widget)
    txt_gr_title=tk.Label(text='Calibration Curve',bg='#ACACAC', font=("MSゴシック","20"))
    txt_gr_title.place(x=500,y=70)
    canvas_background_controll = tk.Canvas(root, bg="gray", width=430, height=280, borderwidth=-2)
    canvas_background_controll.place(x=850, y=57)
    con_txt=[]
    for i in range(9):
        con_txt.append(tk.Entry(width=6, justify='right'))
        con_txt[i].place(x=400+(i*50),y=550)
    abs_txt=[]
    for i in range(9):
        abs_txt.append(tk.Entry(width=6, justify='right'))
        abs_txt[i].place(x=400+(i*50),y=590)
    con_lbl=tk.Label(text='Concetration',bg='#ACACAC')
    con_lbl.place(x=325,y=550)
    abs_lbl=tk.Label(text='Absorbance',bg='#ACACAC')
    abs_lbl.place(x=330,y=590)
    #Flasher session------------------------------------------------------------------------------------------
    # scrollable text box-------------------------------------------------------------------------------------
    text_box = scrolledtext.ScrolledText(root, wrap='none', height=5, width=50)
    text_box.place(x=870,y=260)
    
    device_comb=ttk.Combobox(root,width=17,values=combo_init)
    device_comb.place(x=1125,y=130)
    
    # Make button---------------------------------------------------------------------------------------------
    setup_button = tk.Button(root, width=10,text="setup", command=setup)
    setup_button.place(x=1155,y=220)
    renew_button = tk.Button(root, width=25,text="Re-writing ESP requirement", command=renew)
    renew_button.place(x=870,y=220)
    flash_button = tk.Button(root, width=10,text="flash", command=flash)
    flash_button.place(x=1070,y=220)
    show_key_button = tk.Button(root,width=1,text="*",command=see_key)
    show_key_button.place(x=1250,y=157)
    show_pass_button = tk.Button(root,width=1,text='*',command=see_pass)
    show_pass_button.place(x=1050,y=157)
    get_device_button = tk.Button(root, width=10,text='get_device', command=get_device)
    get_device_button.place(x=870,y=190)
    #text box--------------------------------------------------------------------------------------------------
    SSID_txt=tk.Entry(width=20, justify='left')
    SSID_txt.place(x=925,y=130)
    Password_txt=tk.Entry(width=20, justify='left')
    Password_txt.place(x=925,y=160)
    key_txt=tk.Entry(width=20, justify='right')
    key_txt.place(x=1125,y=160)
    com_txt=tk.Entry(width=13, justify='right')
    com_txt.place(x=1155,y=190)
    Password_txt.config(show="*")
    #key925-894=31,device878-925=52,scroll1070
    #txt message-----------------------------------------------------------------------------------------------
    Flasher_text=tk.Label(text="ESP FLASHER",bg="gray",font=("MSゴシック","18"))
    Flasher_text.place(x=860,y=80)
    SSID_message=tk.Label(text='SSID :',bg="gray")
    SSID_message.place(x=890,y=127)
    password_message=tk.Label(text='Password :',bg="gray")
    password_message.place(x=863,y=157)
    device_message=tk.Label(text='Device :',bg="gray")
    device_message.place(x=1073,y=127)
    key_message=tk.Label(text='Key :',bg="gray")
    key_message.place(x=1094,y=157)
    com_message=tk.Label(text='COM number :',bg="gray")
    com_message.place(x=1070,y=187)
    #Eqation part
    equation_information_text=tk.Label(text="EQUATION INFORMATION",bg="#ACACAC",font=("MSゴシック","18"))
    equation_information_text.place(x=860,y=350)
    add_data_bt=tk.Button(root,text="Superscription",command=lambda:add_data("a"))
    add_data_bt.place(x=900,y=500)
    rewrite_data_bt=tk.Button(root,text="Postscription",command=lambda:add_data("w"))
    rewrite_data_bt.place(x=900,y=525)
    change_approximation=tk.Button(root,text="Change approximation",command=ch_aprx)
    change_approximation.place(x=900,y=550)
#Demolish spectral window
def eraser_measurement():
    lblsubtitle.destroy()
    concentlbl.destroy()
    choicelbl.destroy()
    txtleft.destroy()
    txtright.destroy()
    absorbancetxt.destroy()
    selectwavetxt.destroy()
    for i in range(18):
        lbl[i].destroy()
        txt[i].destroy()
    
    if canvas_background_show.winfo_exists():
        canvas_background_show.destroy()
#Demolish calibration window
def eraser_diagnosis():
    if canvas_background_show.winfo_exists():
        canvas_background_show.destroy()
#Record absorbance with concentration
def record():
    global record_count,record_txt
    abs_txt[record_count].insert(tk.END,round(absorbance_use,3))
    con_txt[record_count].insert(tk.END,record_txt.get())
    log_abs.append(round(absorbance_use,3))
    log_con.append(record_txt.get())
    message_txt.insert(tk.END,"Absorbance is")
    message_txt.insert(tk.END,abs_txt[record_count])
    message_txt.insert(tk.END,"\n")
    message_txt.insert(tk.END,"Concentration is")
    message_txt.insert(tk.END,con_txt[record_count])
    message_txt.insert(tk.END,"\n")
    record_count+=1    
#Start parallel process with measurement function
def start(develop_imi):
    global running
    running=True
    threading.Thread(target=lambda:measurement(develop_imi)).start()
#Measure spectral data
def measurement(develop_imi):
    global mode,absorbance_use,concentration
    if situtxt.get()=="Current Status:Under measurement":
        message_txt.insert(tk.END,"Already measurement mode\n")
    else:   
        while running:
            #separate
            if develop_imi==1:
                if combobox_option.get()!="Initialization":
                    break
                if situtxt.get()=="Current Status:Calibration":
                    eraser_diagnosis()
                    parts_1st()
                situtxt.delete(0,tk.END)
                situtxt.insert(tk.END,"Current Status:Under measurement")
                dele=[]
                for i in range(4):
                    dele.append(ax[i].clear())
                    ax[i].set_title(title[i])
            resultlist_back=[]
            resultlist_cur=[]
            resultlist_ab=[]
            for i in range(9):
                line = process.stdout.readline().decode('utf-8').rstrip()
                if i == 6:
                    line_decode=line
    
            try:
                result_list=delete_wasting_characters(line_decode)
                for i in range(18):
                    resultlist_back.append(int(result_background_list[i]))
                    resultlist_cur.append(int(result_list[i]))
                    resultlist_ab.append(int(result_background_list[i])-int(result_list[i]))
    
            finally:
                process.terminate()
                process.wait()
                process.kill()
            
            y_plot=[]
            
            x  = np.asarray(wavelength)
            y1 = np.asarray(resultlist_back)
            y2 = np.asarray(resultlist_cur)
            y3 = np.asarray(resultlist_ab)
            
            #spline model
            model1 = make_interp_spline(x,y1)
            model2 = make_interp_spline(x,y2)
            model3 = make_interp_spline(x,y3)
            xs  = np.linspace(410,940,531)
            
            #make spline model
            y_plot.append(model1(xs))
            y_plot.append(model2(xs))
            y_plot.append(model3(xs))
            
            #Enable only wavelength between 410 and 940 kk 
            #separate
            if develop_imi==0:
                wavedata=40
            elif develop_imi==1:
                txtdata=[]
                if selectwavetxt.get() == "" or int(selectwavetxt.get())>940 or int(selectwavetxt.get())<410 :
                    wavedata=152
                    txtdata.append("Invalid")
                else :
                    wavedata=int(selectwavetxt.get())-410
                    txtdata.append(selectwavetxt.get())
                    txtdata.append("nm")
            
            #Calculate absorbance
            abs = []
            y4 =[]
    
            for i in range(531):
                if y_plot[0][i] <= 0 or y_plot[1][i] <= 0:
                    y4.append(0)
                else:
                    y4.append(-math.log10(y_plot[1][i]/y_plot[0][i]))
            
            #Make absorbance curve to calculate per 1nm absorbance
            for i in range(531):
                if i==0 or i==25 or i==50 or i==75 or i==100 or i==125 or i==150 or i==175 or i==200 or i==235 or i==270 or i==305 or i==320 or i==350 or i==400 or i==450 or i==490 or i==530:
                    abs.append(y4[i])
    
            y_plot.append(y4)
            absorbance_use=y4[wavedata]
            #separate
            if develop_imi==0:
                coe_X=[]
                df=pd.read_excel(xlsx_path)
                log=int(df.iat[len(df)-1,1])
                concentration=0
                for i in range(log):
                    coe_X.append(df.iat[len(df)-1,i+2]) 
                    concentration+=coe_X[i]*pow(absorbance_use,log-i)
                if absorbance_use<0 :
                    concentration=0
                print(concentration)
            elif develop_imi==1:
                sho =[]
                for i in range(4):
                        sho.append(ax[i].plot(xs,y_plot[i]))
                        if i==2:
                            ax[i].set_yticks([0,50,100,150,200,250])
                            ax[i].set_ylim(0,250)
                        if i==3:
                            ax[i].set_yticks([0,0.2,0.4,0.6,0.8,1.0])
                            ax[i].set_ylim(0,1)
                        else:
                            ax[i].set_yticks([0,200,400,600,800,1000,1200,1400,1600])
                            ax[i].set_ylim(0,1700)
                        ax[i].minorticks_on()         
                        ax[i].grid(which="major", color="black", alpha=0.5)
                        ax[i].grid(which="minor", color="#ACACAC", linestyle=":")                       
                    
                #Deletion of absobance data
                txtleft.delete(0, tk.END)
                txtright.delete(0, tk.END)
                
                for i in range(18):
                    txt[i].delete(0,tk.END)
                
                absorbancetxt.delete(0, tk.END)
                df = pd.read_excel(xlsx_path)
                concentration=0
                #HRP mode task
                if mode == 0 or mode == 1:
                    coe_XH=[]
                    log=int(df.iat[len(df)-1,1])
                    for i in range(log):
                        coe_XH.append(df.iat[len(df)-1,i+2])
                        concentration+=coe_XH[i]*pow(absorbance_use,log-i)
                    if absorbance_use<0 :
                        concentration=0
                    absorbancetxt.insert(tk.END,round(concentration,3))       
                #BSA mode task
                elif mode == 2:
                    coe_XBB=[]
                    log=int(df.iat[len(df)-1,11])
                    for i in range(log):
                        coe_XBB.append(df.iat[len(df)-1,i+12])
                        concentration+=coe_XBB[i]*pow(absorbance_use,log-i)
                    if absorbance_use<0 :
                        concentration=0
                    if concentration>=0 and concentration<=2000:
                        absorbancetxt.insert(tk.END,round(concentration,3))
                    else:
                        absorbancetxt.insert(tk.END,"Need to change mode")
                elif mode == 3:
                    coe_XBI=[]
                    log=int(df.iat[len(df)-1,22])
                    for i in range(log):
                        coe_XBI.append(df.iat[len(df)-1,i+23])
                        concentration+=coe_XBI[i]*pow(absorbance_use,log-i)
                    if absorbance_use<0 :
                        concentration=0
                    if concentration>=0 and concentration<=2000:
                        absorbancetxt.insert(tk.END,round(concentration,3))
                    else:
                        absorbancetxt.insert(tk.END,"Need to change mode")
                #Insert of data
                txtleft.insert(tk.END,txtdata)  
                txtright.insert(tk.END,round(absorbance_use,3))        
                for i in range(18):
                    txt[i].insert(tk.END,round(abs[i],3))       
                canvas_graph.draw()
#Interrupt measurement process
def stop():
    global running
    running=False
    while situtxt.get()!="Current Status:STOP":
        situtxt.delete(0,tk.END)
        situtxt.insert(tk.END,"Current Status:STOP")
#Execute make calibration window
def calibration():
    global running,canvas_background_show,tk_image2,Label2
    running=False
    log=0
    if situtxt.get()!="Current Status:Calibration":
        log=1
    while situtxt.get()!="Current Status:Calibration":
        situtxt.delete(0,tk.END)
        situtxt.insert(tk.END,"Current Status:Calibration")
    if log==1:
        eraser_measurement()
        parts_2nd()
        situtxt.delete(0,tk.END)
        situtxt.insert(tk.END,'Current Status:Calibration')
        for i in range(len(log_abs)-1):
            abs_txt[i].insert(tk.END,log_abs[i+1])
            con_txt[i].insert(tk.END,log_con[i+1])
        if record_before!=0:
            before_txt.insert(tk.END,record_before)
        
    else:
        message_txt.insert(tk.END,"Already calibration mode\n")
#Make calibration curve
def calibration_curve():
    global z
    z=0
    x_cal,y_cal=[],[]
    for i in range(len(log_con)-1):
        x_cal.append(float(log_con[i+1]))
        y_cal.append(float(log_abs[i+1]))
    dele=ax_cal.clear()
    scat=ax_cal.scatter(x_cal,y_cal)
    x_cal,y_cal,coe_det=inspect(x_cal,y_cal)
    plo=ax_cal.plot(x_cal,y_cal)
    canvas_graph.draw()
    equation_R="Coeficient of Determination: "+str(coe_det)
    equation_R_text=tk.Label(text=equation_R,bg="#ACACAC")
    equation_R_text.place(x=890,y=400)
    for i in range(z):
        coe[i]=round(coe[i],4)
    if z==1:
        equation_text=r'y = {a}x'.format(a=coe[0])
    if z==2:
        equation_text=r'y = {a}x^2 + {b}x'.format(a=coe[1],b=coe[0])
    if z==3:
        equation_text=r'y = {a}x^3 + {b}x^2 + {c}x'.format(a=coe[2],b=coe[1],c=coe[0])
    if z==4:
        equation_text=r'y = {a}x^4 + {b}x^3 + {c}x^2 + {d}x'.format(a=coe[3],b=coe[2],c=coe[1],d=coe[0])
    if z==5:
        equation_text=r'y = {a}x^5 + {b}x^4 + {c}x^3 + {d}x^2 + {e}x'.format(a=coe[4],b=coe[3],c=coe[2],d=coe[1],e=coe[0])       
    equation_message=tk.Label(text=equation_text,bg="#ACACAC")
    equation_message.place(x=890,y=425)
    message_txt.insert(tk.END,"Calibration curve is made\n")
#Select sample and assay method
def select_combo(event):
    global mode
    'HRP・IgG','BCA・BSA','BCA・IgG'
    if combobox_mode.get()=='HRP・IgG':
        mode = 1
    elif combobox_mode.get()=='BCA・BSA':
        mode = 2
    elif combobox_mode.get()=='BCA・IgG':
        mode = 3    
#Open developer mode
def development():
    global combobox_device,combobox_mode,tk_image3,message_txt,dev,record_txt
    #read password
    if dev_pass_txt.get()=="Menicon":
        #remake logo
        parts_1st()
        Label.destroy()
        image1_path="image\image1.png"
        pre_image=Image.open(image1_path)
        tk_image3=ImageTk.PhotoImage(pre_image)
        Label3 = tk.Label(root,image=tk_image3,bg="skyblue")
        Label3.place(x=1,y=1)
        #replace text
        ratio_concentration.place(x=114,y=200)
        ratio_concentration_tx.place(x=81,y=200)
        situtxt.place(x=20,y=260)
        before_txt.place(x=72,y=120)
        after_txt.place(x=216,y=120)
        bf_concentration_tx.place(x=32,y=120)
        af_concentration_tx.place(x=181,y=120)
        Label2.place(x=45,y=60)
        canvas_bar.place(x=280,y=135)
        bar_labe1.place(x=210,y=140)
        bar_labe2.place(x=195,y=235)
        #button replace
        ev_fit_bt.place(x=120,y=230)
        get_cons_bt.place(x=10,y=230)
        easy_set_bt.destroy()
        #make button
        bt_start = tk.Button(root, text="START", command=lambda:start(1))
        bt_start.place(x=50, y=385)
        bt_stop = tk.Button(root, text="STOP", command=stop)
        bt_stop.place(x=95, y=385)
        bt_calibration = tk.Button(root, text="Calibration", command=calibration)
        bt_calibration.place(x=140, y=385)
        bt_record = tk.Button(root, text="Record", command=record)
        bt_record.place(x=95,y=455)
        bt_make_fit = tk.Button(root, text="Make Calibration Curve", command=calibration_curve)
        bt_make_fit.place(x=150,y=455)
        #message box
        message_txt=tk.Text(width=35,height=13)
        message_txt.place(x=20,y=505)
        message_txt.insert(tk.END,"Password is autherized\n")
        
        box_txt=tk.Label(text="comboboxes")
        box_txt.place(x=50,y=285)
        with open(devicename_path, "r") as f:
            devicename = [line.strip() for line in f if line.strip() != "end"]
        combobox_device = ttk.Combobox(root, values=devicename, width=10)
        combobox_device.bind('<<ComboboxSelected>>', select_device)
        combobox_device.place(x=50, y=305)
        module = ('HRP・IgG','BCA・BSA','BCA・IgG')
        v = tk.StringVar()
        combobox_mode = ttk.Combobox(root, textvariable= v, values=module, style="office.TCombobox")
        combobox_mode.bind('<<ComboboxSelected>>', select_combo)
        combobox_mode.place(x=50,y=335)
        #Separate Label
        main_buttons=tk.Label(text="Main buttons")
        main_buttons.place(x=50,y=360) 
        Calibration_buttons=tk.Label(text="Calibration_buttons")
        Calibration_buttons.place(x=50,y=430)
        #----
        record_txt=tk.Entry(width=6, justify='right')
        record_txt.place(x=50,y=460)
        dev=1
#Read spectral data easily
def easy_set():
    global result_background_list,process
    situtxt.delete(0,tk.END)
    situtxt.insert(tk.END,"Initialization")
    with open(devicename_path,'r',encoding='utf-8') as file:
        device=file.readline().rstrip("\n")
    mes=f"az iot hub monitor-events --device-id {device} --hub-name MeniconPilotHub"
    process = subprocess.Popen(mes, stdout=subprocess.PIPE, shell=True)
    line = process.stdout.readline().decode('utf-8').rstrip()
    print(line)
    for i in range(9):
        line = process.stdout.readline().decode('utf-8').rstrip()
        if i == 6:
            a = line
    result_background_list = delete_wasting_characters(a)
    #モード設定
    develop_imi=0
    start(develop_imi)
#Get concentration
def get_cons():
    global record_before
    if before_txt.get()=="":
        before_txt.insert(tk.END,concentration)
        record_before=before_txt.get()
    else:
        after_txt.delete(0,tk.END)
        after_txt.insert(tk.END,concentration)
#Evaluation wearability of contact lenses
def ev_fit():
    global tk_image4,Label4
    image_path="image\image3.png","image\image4.png","image\image5.png","image\image6.png","image\image7.png"
    inc=100*(float(after_txt.get())-float(before_txt.get()))/float(before_txt.get())
    ratio_concentration.delete(0,tk.END)
    ratio_concentration.insert(tk.END,inc)
    ratio_concentration.insert(tk.END," %")
    threshold=[]
    #20,40,60,80,100
    for i in range(5):
        threshold.append(20*(i+1))
    for i in range(len(threshold)):
        if inc<threshold[i]:
            image3_path=Image.open(image_path[i])
            original_wi, original_hi = image3_path.size  
            aspect_ratio = 0.3
            new_width = int(original_wi * aspect_ratio)
            new_height = int(original_hi * aspect_ratio)
            resized_image2 = image3_path.resize((new_width, new_height))
            crop_image2=resized_image2.crop((0,20,100,75))
            tk_image4=ImageTk.PhotoImage(crop_image2)
            Label4 = tk.Label(root,image=tk_image4)
            if dev==0:
                Label4.place(x=300,y=77)
            elif dev==1:
                Label4.place(x=80,y=137)
            break
        else:
            None
#Prepare main window
def set_main():
    global dev_pass_txt,before_txt,bf_concentration_tx,af_concentration_tx,after_txt,ratio_concentration,ratio_concentration_tx,tk_image,get_cons_bt,easy_set_bt,ev_fit_bt,Label,situtxt,canvas_bar,bar_labe1,bar_labe2
    canvas_background_controll = tk.Canvas(root, width=450, height=670, borderwidth=-2)
    canvas_background_controll.place(x=0, y=55)
    #make text_box
    before_txt = tk.Entry(width=6, justify='right')
    after_txt = tk.Entry(width=6, justify='right')
    ratio_concentration = tk.Entry(width=6, justify='right')
    before_txt.place(x=72,y=140)
    after_txt.place(x=216,y=140)
    ratio_concentration.place(x=340,y=140)
    dev_pass_txt = tk.Entry(width=16, justify='right')
    dev_pass_txt.place(x=340,y=200)
    
    #make text
    bf_concentration_tx = tk.Label(text="Before")
    af_concentration_tx = tk.Label(text="After")
    ratio_concentration_tx = tk.Label(text="Ratio")
    bf_concentration_tx.place(x=32,y=140)
    af_concentration_tx.place(x=181,y=140)
    ratio_concentration_tx.place(x=307,y=140)

    #make button
    dev_bt=tk.Button(root,text="Development mode",command=development)
    easy_set_bt=tk.Button(root,text="Initialize setting",command=easy_set)
    ev_fit_bt=tk.Button(root,text="Evaluation",command=ev_fit)
    get_cons_bt=tk.Button(root,text="Get concentration",command=get_cons)
    dev_bt.place(x=335,y=225)
    easy_set_bt.place(x=38,y=170)
    ev_fit_bt.place(x=235,y=170)
    get_cons_bt.place(x=130,y=170)

    image1_path="image\image1.png"
    pre_image=Image.open(image1_path)
    pre_hight,pre_width = pre_image.size
    aspect_ratio=0.8
    as_image=pre_image.resize((int(pre_hight*aspect_ratio),int(pre_width*aspect_ratio)))
    tk_image=ImageTk.PhotoImage(as_image)
    Label = tk.Label(root,image=tk_image,bg="skyblue")
    Label.place(x=1,y=1)

    #sitiation box
    situtxt = tk.Entry(width=35, justify='center')
    situtxt.place(x=20, y=220)
    situtxt.insert(tk.END,"Current Status:None")
     # 色コードのリスト
    colors = ["#00FF00", "#80FF00", "#FFFF00", "#FF8000", "#FF0000"]
    
    def create_bars_show(canvas_bars, colors, bar_height=25):
        y_start = 0
        for color in colors:
            # 各バーを描画
            canvas_bars.create_rectangle(0, y_start, 10, y_start + bar_height, fill=color, outline="")
            y_start += bar_height
    
    # キャンバスの作成
    canvas_bar = tk.Canvas(root, width=10, height=len(colors) * 25)
    canvas_bar.place(x=430,y=60)
    # 色のバーを作成
    create_bars_show(canvas_bar, colors)
    bar_labe1=tk.Label(text="Comfortable")
    bar_labe2=tk.Label(text="Uncomfortable")
    bar_labe1.place(x=350,y=60)
    bar_labe2.place(x=340,y=165)
#Comprises all functions
def diagnosis():
    global tk_image2,Label2
    set_main()
    image_path="image\image2.png"
    pre_image=Image.open(image_path)
    pre_hight,pre_width = pre_image.size
    aspect_ratio=0.3
    resized_image=pre_image.resize((int(pre_hight*aspect_ratio),int(pre_width*aspect_ratio)))
    crop_image = resized_image.crop((5,25,240,75))
    tk_image2=ImageTk.PhotoImage(crop_image)
    Label2 = tk.Label(root,image=tk_image2)
    Label2.place(x=45,y=80)
#Execute GUI
root=tk.Tk()
root.geometry('450x250')
root.title("PoC Spectrometer")
root.configure(background="skyblue")
diagnosis()
root.mainloop()