from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import datetime
import time
import os
import sys
import logging
import traceback
from pymongo import MongoClient
import sqlite3
import win32com.client as win32
from jproperties import Properties
#import subprocess

def error_handler(exctype, value, tb):
    logging.error('-------------------------------ERROR-------------------------------')
    logging.error('Type:' + str(exctype))
    logging.error('Value:' + str(value))
    logging.error('Traceback:' + str(tb))
    logging.error('-------------------------------------------------------------------')


def send_mail(fTo, fSubject, fBody):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = fTo
    mail.Subject = fSubject
    mail.Body = fBody
    #mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional
    #attachment  = "Path to the attachment"
    #mail.Attachments.Add(attachment)
    
    #mail.display()
    mail.Send()
    
def convert_date_format(s_date) -> str:
    for i in range(0, 31):
        new_date = datetime.date.today() + datetime.timedelta(days=i)
        if int(s_date[-2:]) == new_date.day:
            return str(new_date.strftime("%d/%m/%Y"))
        

def init_variables():
    global Loop_Timer
    global Initial_Rep
    global Download_Schedule_Timer
    global Download_Reserves_Timer
    global Check_Reserves_Timer
    global Start_Reserve_Time_In_Advance
    global Server_Start_Time
    global Server_Shutdown_Time
    global Chrome_Window_Size_Schedule
    global Chrome_Window_Size_Reserves
    global Admin_Email
    global Shutdown_Path
    global s_User
    global s_Password
    global Contact_Email
    global Fav_Slot_Default
    global Gofit_Url_Login
    global Gofit_Url_Schedule
    global Chrome_Driver_Path
    global DB_Path
    global Application_Path
    Application_Path = os.path.dirname(os.path.abspath(__file__))
    
    #INICIALIZAR EL LOG
    logging.basicConfig(filename=Application_Path + '\\log.log', filemode='a', format='%(asctime)s - %(message)s', level=logging.INFO)
    
    #IMPORTAR ARCHIVO DE PROPIEDADES
    global configs
    
    configs = Properties()
    with open(Application_Path + '\\config.properties', 'rb') as read_prop:
        configs.load(read_prop)
    
    
    Loop_Timer = float(configs.get("Loop_Timer").data)
    Initial_Rep = int(configs.get("Initial_Rep").data)
    Download_Schedule_Timer = float(configs.get("Download_Schedule_Timer").data)
    Download_Reserves_Timer = float(configs.get("Download_Reserves_Timer").data)
    Check_Reserves_Timer = float(configs.get("Check_Reserves_Timer").data)
    Start_Reserve_Time_In_Advance = float(configs.get("Start_Reserve_Time_In_Advance").data)
    Server_Start_Time = datetime.datetime.strptime(str(datetime.date.today()), "%Y-%m-%d") + datetime.timedelta(hours=float(configs.get("Server_Start_Time").data.split(":")[0])) + datetime.timedelta(minutes=float(configs.get("Server_Start_Time").data.split(":")[1]))
    Server_Shutdown_Time = datetime.datetime.strptime(str(datetime.date.today()), "%Y-%m-%d") + datetime.timedelta(hours=float(configs.get("Server_Shutdown_Time").data.split(":")[0])) + datetime.timedelta(minutes=float(configs.get("Server_Shutdown_Time").data.split(":")[1]))
    
    Chrome_Window_Size_Schedule = configs.get("Chrome_Window_Size_Schedule").data
    Chrome_Window_Size_Reserves = configs.get("Chrome_Window_Size_Reserves").data
    
    Admin_Email = configs.get("Admin_Email").data
    Shutdown_Path = configs.get("Shutdown_Path").data
    
    s_User = configs.get("s_User").data
    s_Password = configs.get("s_Password").data
    Contact_Email = configs.get("Contact_Email").data
    
    Gofit_Url_Login = configs.get("Gofit_Url_Login").data
    Gofit_Url_Schedule = configs.get("Gofit_Url_Schedule").data
    
    Chrome_Driver_Path = configs.get("Chrome_Driver_Path").data
    DB_Path = configs.get("DB_Path").data
    
    Fav_Slot_Default = int(configs.get("Fav_Slot_Default").data)

    global Fav_Slot_List
    Fav_Slot_List = {
        "Activity": [],
        "Slot_Number": []
    }

    for i in range(0, int(configs.get("Fav_Slot_List.count").data)):
        Fav_Slot_List["Activity"].append(configs.get("Fav_Slot_List." + str(i) + ".activity").data)
        Fav_Slot_List["Slot_Number"].append(int(configs.get("Fav_Slot_List." + str(i) + ".slot").data))

    global Init_Time
    Init_Time = datetime.datetime.now()
    
    global rep
    rep = Initial_Rep
    
    load_reserves_database()
    
    
def get_fav_slot(r_Activity) -> int:
    for index in range(0, len(Fav_Slot_List["Activity"])):
        if Fav_Slot_List["Activity"][index] == r_Activity:
            return Fav_Slot_List["Slot_Number"][Fav_Slot_List["Activity"].index(r_Activity)]
    return Fav_Slot_Default
            

def load_reserves_database():
    global T_Reserves
    T_Reserves = {
        "Activity_ID": [],
        "Reserve_Start_Date": [],
        "Reserve_Status": []
    }
    
    try:
        conn = sqlite3.connect(DB_Path)
        cur = conn.cursor()
        cur.execute("SELECT * FROM T_Reserves")
        rows = cur.fetchall()
        conn.close()
        
        for row in rows:
            if (datetime.datetime.now()) < (datetime.datetime.strptime(row[1], "%d/%m/%Y %H:%M") + datetime.timedelta(hours=48)):
                T_Reserves["Activity_ID"].append(row[0])
                T_Reserves["Reserve_Start_Date"].append(row[1])
                T_Reserves["Reserve_Status"].append(row[2])
        
        logging.info('LOAD RESERVES FROM DB COMPLETED')
        print("    Load reserves from DB completed")
        
    except:
        e = traceback.format_exc()
        logging.error('Error:' + e)
        print(str(rep) + " rep >>> ************************ERROR************************")
        print(str(rep) + " rep >>> load_reserves_database")


def update_reserves_database():

    #print("    Saving reserves... >>> " + str(datetime.datetime.now().isoformat(sep=' ', timespec='seconds')))
    
    try:
        conn = sqlite3.connect(DB_Path)
        
        for index in range(0, len(T_Reserves["Activity_ID"])):
            
            row = [
                str(T_Reserves["Activity_ID"][index]),
                str(T_Reserves["Reserve_Start_Date"][index]),
                str(T_Reserves["Reserve_Status"][index])
            ]

            conn.execute("INSERT INTO T_Reserves (Activity_ID, Reserve_Start_Date, Reserve_Status) VALUES (?,?,?)", row)
            conn.commit()
        
        conn.close()
        logging.info('UPDATE RESERVES DB COMPLETED')
        print("    Update reserves DB completed")
            
    except:
        e = traceback.format_exc()
        logging.error('Error:' + e)
        print(str(rep) + " rep >>> ************************ERROR************************")
        print(str(rep) + " rep >>> update_reserves_database")


def update_schedule_database():
    
    #print("    Saving schedule... >>> " + str(datetime.datetime.now().isoformat(sep=' ', timespec='seconds')))
    
    try:
        conn = sqlite3.connect(DB_Path)
        
        for index in range(0, len(T_Schedule["Activity_ID"])):
            
            row = [
                str(T_Schedule["Weekday"][index]),
                str(T_Schedule["Day"][index]),
                str(T_Schedule["Start_Time"][index]),
                str(T_Schedule["End_Time"][index]),
                str(T_Schedule["Activity"][index]),
                str(T_Schedule["Room"][index]),
                str(T_Schedule["Monitor"][index]),
                str(T_Schedule["Color"][index]),
                str(T_Schedule["Activity_Start_Date"][index]),
                str(T_Schedule["Reserve_Start_Date"][index]),
                str(T_Schedule["Activity_ID"][index]),
                str(T_Schedule["Status"][index])
            ]

            conn.execute("INSERT INTO T_Schedule (Weekday, Day, Start_Time, End_Time, Activity, Room, Monitor, Color, Activity_Start_Date, Reserve_Start_Date, Activity_ID, Status) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", row)
            conn.commit()
            
        conn.close()
        logging.info('UPDATE SCHEDULE DB COMPLETED')
        print("    Update schedule DB completed")
        
    except:
        e = traceback.format_exc()
        logging.error('Error:' + e)
        print(str(rep) + " rep >>> ************************ERROR************************")
        print(str(rep) + " rep >>> update_schedule_database")

        
def search_in_schedule(r_Day, r_Start_Time, r_Activity) -> str:
    
    list_of_days = []
    index_r_Day = -1
    
    print("    Searching activity in schedule... >>> " + str(datetime.datetime.now().isoformat(sep=' ', timespec='seconds')))
    
    time.sleep(1.2)
    
    try:
        #LEEMOS LA CABECERA DE LAS COLUMNAS
        header_days = browser.find_elements(By.XPATH, "//div[contains(@class, 'contenedor-cabecera-dias padding-9 col-xs-2 cajaNombredias ng-scope')]")
        
        #HAYAMOS EL INDICE DEL DIA QUE BUSCAMOS
        for index, day in enumerate(header_days):
            list_of_days.append(convert_date_format(day.get_attribute('innerText')))
            if list_of_days[index] == r_Day:
                index_r_Day = index
        
        if index_r_Day == -1:
            return "Not found"
        
        #LEEMOS LAS ACTIVIDADES DE LA COLUMNA DEL DIA ELEGIDO
        column = browser.find_element(By.XPATH, "//div[contains(@class, 'contenedor-item-dia maximoEntre')][" + str(index_r_Day + 1) + "]")
        
        activities_list = column.find_elements(By.XPATH, ".//div[@class='div item-dias altoNormal alturaActividadesReservas ng-scope grisClaro']")
            
        for activity in activities_list:
            s_start_time = activity.find_element(By.XPATH, ".//span[@class='label etiquetaHora ng-binding']").get_attribute('innerText').split("/", 2)[0].strip()
            s_activity = activity.find_element(By.XPATH, ".//div[@class='actividad centrarElementos']").get_attribute('innerText')
            
            if s_start_time == r_Start_Time and s_activity == r_Activity:
                r_current_status = activity.find_element(By.XPATH, ".//span[@class='label padding-0 ng-binding ng-scope']").get_attribute('innerText')
                if r_current_status.casefold() == "Reservar ya".casefold():
                    activity.click()
                    return "AVAILABLE"
                elif r_current_status.casefold() == "Completa".casefold():
                    return "FULL"
                elif r_current_status.casefold() == "No disponible".casefold():
                    return "UNAVAILABLE"
                elif r_current_status.casefold() == "Finalizada".casefold():
                    return "FINISHED"
                else:
                    return "ERROR"

                        
    except:
        e = traceback.format_exc()
        logging.error('Error:' + e)
        print(str(rep) + " rep >>> ************************ERROR************************")
        print(str(rep) + " rep >>> search_in_shedule")


def reserve_activity(r_Activity) -> str:
    
    slot_list = []
    selected_slot = "-"
    
    try:
        slots_table = browser.find_element(By.ID, "puestos-horario")
        
        #DETECTAMOS SI LA ACTIVIDAD NO TIENE LISTA DE PUESTOS
        if slots_table.get_attribute('innerText') == "":
            reserve_button = browser.find_element(By.XPATH, "//*[@class='btn-tg btn-tg-modal-actividad tg-centrado btn-tg-modal-salir clicked ng-scope']")
            reserve_button.click()
            selected_slot = "N/A"     
            return selected_slot
        
        #LISTAMOS LOS PUESTOS
        for slot in slots_table.find_elements(By.XPATH, ".//*[@class='puesto tg-centrado clicked padding-0 centrarElementos ng-scope puesto_libre']"):
            slot_list.append(slot.get_attribute('innerText'))
        
        reserve_button = browser.find_element(By.XPATH, "//*[@class='btn-tg btn-tg-modal-plazas col-md-5 col-xs-offset-1 col-xs-10 tg-centrado btn-tg-modal-salir clicked']")
        
        #SELECCIONAMOS EL PUESTO MAS CERCANO
        for slot_nbr in slot_list:
            if int(slot_nbr) >= get_fav_slot(r_Activity):
                slot_button = slots_table.find_elements(By.XPATH, ".//*[@class='puesto tg-centrado clicked padding-0 centrarElementos ng-scope puesto_libre']")[slot_list.index(slot_nbr)]
                slot_button.click()
                time.sleep(0.2)
                reserve_button.click()
                selected_slot = slot_nbr
                time.sleep(0.2)
                test_elements = browser.find_elements(By.ID,"mTitulo") #DETECTAR SI HA FALLADO LA RESERVA
                for element in test_elements:
                    if (element.get_attribute('innerText') == "ATENCIÓN") or (element.get_attribute('innerText') == "atención"):
                        selected_slot = "REPEATED"
                return selected_slot
                
        for slot_nbr in slot_list.reverse():
            if int(slot_nbr) < get_fav_slot(r_Activity):
                slot_button = slots_table.find_elements(By.XPATH, ".//*[@class='puesto tg-centrado clicked padding-0 centrarElementos ng-scope puesto_libre']")[slot_list.index(slot_nbr)]
                slot_button.click()
                time.sleep(0.2)
                reserve_button.click()
                selected_slot = slot_nbr
                time.sleep(0.2)
                test_elements = browser.find_elements(By.ID,"mTitulo") #DETECTAR SI HA FALLADO LA RESERVA
                for element in test_elements:
                    #if (element.get_attribute('innerText') == "ATENCIÓN") or (element.get_attribute('innerText') == "atención"):
                    if element.get_attribute('innerText').casefold() == "ATENCIÓN".casefold():
                        selected_slot = "REPEATED"
                return selected_slot
            
    except:
        e = traceback.format_exc()
        logging.error('Error:' + e)
        print(str(rep) + " rep >>> ************************ERROR************************")
        print(str(rep) + " rep >>> reserve_activity")

def read_schedule():

    dia_en_semana = ("Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo")
    list_of_days = []
    
    print("    Downloading schedule... >>> " + str(datetime.datetime.now().isoformat(sep=' ', timespec='seconds')))
    
    try:
        #LEEMOS LA CABECERA DE LAS COLUMNAS
        header_days = browser.find_elements(By.XPATH, "//div[contains(@class, 'contenedor-cabecera-dias padding-9 col-xs-2 cajaNombredias ng-scope')]")
        
        for day in header_days:
            list_of_days.append(convert_date_format(day.get_attribute('innerText')))
            
        #LEEMOS LAS ACTIVIDADES DE CADA COLUMNA
        column_list = browser.find_elements(By.XPATH, "//div[contains(@class, 'contenedor-item-dia maximoEntre')]")
        
        day_index = 0
        
        for column in column_list:
            activities_list = column.find_elements(By.XPATH, ".//div[@class='div item-dias altoNormal alturaActividadesReservas ng-scope grisClaro']")
            
            for activity in activities_list:
                s_day = list_of_days[day_index]
                T_Schedule["Day"].append(s_day)
                s_weekday = dia_en_semana[datetime.datetime.strptime(s_day, "%d/%m/%Y").weekday()]
                T_Schedule["Weekday"].append(s_weekday)
                s_start_time = activity.find_element(By.XPATH, ".//span[@class='label etiquetaHora ng-binding']").get_attribute('innerText').split("/", 2)[0].strip()
                T_Schedule["Start_Time"].append(s_start_time)
                s_activity = activity.find_element(By.XPATH, ".//div[@class='actividad centrarElementos']").get_attribute('innerText')
                T_Schedule["Activity"].append(s_activity)
                T_Schedule["End_Time"].append(activity.find_element(By.XPATH, ".//span[@class='label etiquetaHora ng-binding']").get_attribute('innerText').split("/", 2)[1].strip())
                T_Schedule["Room"].append(activity.find_elements(By.XPATH, ".//div[@class='salaMonitor centrarElementos']//span")[0].get_attribute('innerText').split(":", 2)[1].strip())
                T_Schedule["Monitor"].append(activity.find_elements(By.XPATH, ".//div[@class='salaMonitor centrarElementos']//span")[1].get_attribute('innerText').split(":", 2)[1].strip())
                T_Schedule["Color"].append(activity.find_element(By.XPATH, ".//div[@class='lineaColorActividad']").value_of_css_property('background-color'))
                T_Schedule["Status"].append(activity.find_element(By.XPATH, ".//span[@class='label padding-0 ng-binding ng-scope']").get_attribute('innerText'))
                
                s_activity_start_date = s_day + " " + s_start_time
                T_Schedule["Activity_Start_Date"].append(s_activity_start_date)
                s_reserve_start_date = datetime.datetime.strptime(s_activity_start_date, "%d/%m/%Y %H:%M") + datetime.timedelta(hours=-49)
                T_Schedule["Reserve_Start_Date"].append(str(s_reserve_start_date.strftime("%d/%m/%Y %H:%M")))
                T_Schedule["Activity_ID"].append(s_day + "_" + s_start_time + "_" + s_activity)
            
            day_index += 1

    except:
        e = traceback.format_exc()
        logging.error('Error:' + e)
        print(str(rep) + " rep >>> ************************ERROR************************")
        print(str(rep) + " rep >>> read_shedule")


def load_schedule(window_size):
    global browser
    
    print("    Conecting to gofit.es... >>> " + str(datetime.datetime.now().isoformat(sep=' ', timespec='seconds')))

    try:
        #START BROWSER
        options = Options()
        #options.headless = True
        options.add_argument("--window-size=" + window_size)
        options.add_argument("--incognito")
        s = Service(Chrome_Driver_Path)
        browser = webdriver.Chrome(service=s, options=options)
        #browser.implicitly_wait(10)
        
        #ABRIR URL GOFIT
        browser.get(Gofit_Url_Login)

        #LOG IN
        browser.find_element(By.NAME, 'tg_login_email').send_keys(s_User)
        browser.find_element(By.NAME, 'tg_login_password').send_keys(s_Password)
        accept_button = browser.find_element(By.XPATH, "//form[@class='wrap-form js-validateUser']//button[@class='button filled orange']")
        accept_button.click()
        try:
            WebDriverWait(browser, 20).until(EC.title_contains('Área privada'))
        except:
            logging.error(str(rep) + " rep >>> load Area privada - TimeoutException")
            print(str(rep) + " rep >>> ************************ERROR************************")
            print(str(rep) + " rep >>> load Area privada - TimeoutException")
        
        #CERRAR VENTANDE DE COOKIES
        try:
            accept_button = browser.find_element(By.ID, 'wt-cli-accept-all-btn')
            accept_button.click()
        except:
            logging.error(str(rep) + " rep >>> Area privada - Cookies sin ventana")
            print(str(rep) + " rep >>> ************************ERROR************************")
            print(str(rep) + " rep >>> Area privada - Cookies sin ventana")
            
        #BUSCAR URL HORARIO
        browser.get(Gofit_Url_Schedule)
        
        try:
            #WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.XPATH, "//article[@class='block-booking block-no-activity']//iframe")))
            WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.XPATH, "//*[contains(@class, 'block-booking')]//iframe")))
        except:
            logging.error(str(rep) + " rep >>> load Gofit_Url_Schedule - TimeoutException")
            print(str(rep) + " rep >>> ************************ERROR************************")
            print(str(rep) + " rep >>> load Gofit_Url_Schedule - TimeoutException")
        
        #Url_Schedule = browser.find_element(By.XPATH, "//article[@class='block-booking block-no-activity']//iframe").get_attribute("src")
        Url_Schedule = browser.find_element(By.XPATH, "//*[contains(@class, 'block-booking')]//iframe").get_attribute("src")

        #NAVEGAR AL HORARIO
        browser.get(Url_Schedule)
        
        try:
            WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'contenedor-item-dia maximoEntre')]")))
        except:
            logging.error(str(rep) + " rep >>> load Url_Schedule - TimeoutException")
            print(str(rep) + " rep >>> ************************ERROR************************")
            print(str(rep) + " rep >>> load Url_Schedule - TimeoutException")
            
        time.sleep(3)
        
        #PROTECCION CONTRA CAJA-LOGIN
        #login_box = browser.find_element(By.XPATH, "//*[@class='caja-login-input']")
                
    except Exception as e:
        e = traceback.format_exc()
        logging.error('Error:' + e)
        print(str(rep) + " rep >>> ************************ERROR************************")
        print(str(rep) + " rep >>> load_schedule")
        
        
def get_schedule_main():
    global T_Schedule
    
    T_Schedule = {
        "Weekday": [],
        "Day": [],
        "Start_Time": [],
        "End_Time": [],
        "Activity": [],
        "Room": [],
        "Monitor": [],
        "Color": [],
        "Activity_Start_Date": [],
        "Reserve_Start_Date": [],
        "Activity_ID": [],
        "Status": []
    }
    
    try:
        load_schedule(Chrome_Window_Size_Schedule)
        
        read_schedule()
        
        #CAMBIO DE SEMANA
        week_button = browser.find_element(By.XPATH, "//*[@class='fa icon-avanzar redIcon30 bolder']")
        week_button.click()
        
        try:
            WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'contenedor-item-dia maximoEntre')]")))
        except:
            logging.error(str(rep) + " rep >>> cambio_semana - TimeoutException")
            print(str(rep) + " rep >>> ************************ERROR************************")
            print(str(rep) + " rep >>> cambio_semana - TimeoutException")
            
        time.sleep(3)
        
        read_schedule()
        
        #LIMPIAR LA DB
        if len(T_Schedule["Day"]) > 2:
            conn = sqlite3.connect(DB_Path)
            conn.execute("DELETE FROM T_Schedule")
            conn.commit()
            conn.close()
                    
            update_schedule_database()
        
    except Exception as e:
        e = traceback.format_exc()
        logging.error('Error:' + e)
        print(str(rep) + " rep >>> ************************ERROR************************")
        print(str(rep) + " rep >>> get_schedule_main")
        
    finally:
        browser.quit()
        T_Schedule.clear()        

        
def make_reserve_main(reserve_ID, r_Day, r_Start_Time, r_Activity):
    
    dia_en_semana = ("Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo")
    r_weekday = dia_en_semana[datetime.datetime.strptime(r_Day, "%d/%m/%Y").weekday()]
    
    try:
        logging.info('Trying to reserve >>> ' + reserve_ID)
        print("    Trying to reserve >>> " + reserve_ID)
        
        Start_Reserve_Time = datetime.datetime.now()
        
        load_schedule(Chrome_Window_Size_Reserves)
        
        reserve_current_status = search_in_schedule(r_Day, r_Start_Time, r_Activity)
        T_Reserves["Reserve_Status"][T_Reserves["Activity_ID"].index(reserve_ID)] = reserve_current_status
        time.sleep(0.5)
        
        if reserve_current_status == "AVAILABLE":
            final_slot = reserve_activity(r_Activity)
            time.sleep(0.5)
            
            if final_slot == "-":
                send_mail(Contact_Email, "Reserve ERROR - " + reserve_ID, reserve_ID + " error during slot selection.\nManual reserve recommended.")
            elif final_slot == "REPEATED":
                send_mail(Contact_Email, "NOT RESERVED - Class is incompatible - " + r_Activity + " - " + r_weekday[0:3] + "-" + r_Day[0:2] + " " + r_Start_Time, r_Activity + " - " + r_weekday + " " + r_Day + " " + r_Start_Time + " impossible to reserve. Class is incompatible with another reserve or daily reserve limit exceeded.")
                #send_mail(Admin_Email, "BEA NOT RESERVED - Class is incompatible - " + r_Activity + " - " + r_weekday[0:3] + "-" + r_Day[0:2] + " " + r_Start_Time, r_Activity + " - " + r_weekday + " " + r_Day + " " + r_Start_Time + " impossible to reserve for BEA. Class is incompatible with another reserve or daily reserve limit exceeded.")
                T_Reserves["Reserve_Status"][T_Reserves["Activity_ID"].index(reserve_ID)] = "INCOMPATIBLE"
                logging.info('NOT RESERVED - Class is incompatible with another reserve or daily reserve limit exceeded >>> ' + reserve_ID)
                print("    NOT RESERVED - Class is incompatible with another reserve or daily reserve limit exceeded >>> " + reserve_ID)
            else:
                T_Reserves["Reserve_Status"][T_Reserves["Activity_ID"].index(reserve_ID)] = "RESERVED"
                send_mail(Contact_Email, "Reserve COMPLETED - " + r_Activity + " - " + r_weekday[0:3] + "-" + r_Day[0:2] + " " + r_Start_Time, r_Activity + " - " + r_weekday + " " + r_Day + " " + r_Start_Time + " reserve COMPLETED.\nSlot: " + final_slot)
                #send_mail(Admin_Email, " BEA Reserve COMPLETED - " + r_Activity + " - " + r_weekday[0:3] + "-" + r_Day[0:2] + " " + r_Start_Time, r_Activity + " - " + r_weekday + " " + r_Day + " " + r_Start_Time + " reserve COMPLETED for BEA.\nBEA slot: " + final_slot)
                logging.info('Reserve COMPLETED >>> ' + reserve_ID)
                print("    Reserve COMPLETED >>> " + reserve_ID)

        elif reserve_current_status == "UNAVAILABLE":
            send_mail(Contact_Email, "NOT RESERVED - Class is UNAVAILABLE yet - " + r_Activity + " - " + r_weekday[0:3] + "-" + r_Day[0:2] + " " + r_Start_Time, r_Activity + " - " + r_weekday + " " + r_Day + " " + r_Start_Time + " impossible to reserve. Class is UNAVAILABLE yet.\n\nStart reserve time: " + str(Start_Reserve_Time.isoformat(sep=' ', timespec='seconds')) + "\nClass is UNAVAILABLE at: " + str(datetime.datetime.now().isoformat(sep=' ', timespec='seconds')))
            logging.info('NOT RESERVED - Class is UNAVAILABLE yet >>> ' + reserve_ID)
            print("    NOT RESERVED - Class is UNAVAILABLE yet >>> " + reserve_ID)
        elif reserve_current_status == "FULL":
            send_mail(Contact_Email, "NOT RESERVED - Class is FULL - " + r_Activity + " - " + r_weekday[0:3] + "-" + r_Day[0:2] + " " + r_Start_Time, r_Activity + " - " + r_weekday + " " + r_Day + " " + r_Start_Time + " impossible to reserve. Class is FULL.\nTry to enter the wating list via GoFit app on your smartphone.\n\nStart reserve time: " + str(Start_Reserve_Time.isoformat(sep=' ', timespec='seconds')) + "\nClass is full at: " + str(datetime.datetime.now().isoformat(sep=' ', timespec='seconds')))
            logging.info('NOT RESERVED - Class is FULL >>> ' + reserve_ID)
            print("    NOT RESERVED - Class is FULL >>> " + reserve_ID)
        elif reserve_current_status == "FINISHED":
            send_mail(Contact_Email, "NOT RESERVED - Class is FINISHED - " + reserve_ID, reserve_ID + " impossible to reserve. Class is FINISHED at " + str(datetime.datetime.now().isoformat(sep=' ', timespec='seconds')))
            logging.info('NOT RESERVED - Class is FINISHED >>> ' + reserve_ID)
            print("    NOT RESERVED - Class is FINISHED >>> " + reserve_ID)
        else:
            send_mail(Contact_Email, "Reserve ERROR - " + reserve_ID, reserve_ID + " not found in schedule.\nManual reserve recommended.")
            logging.info('RESERVE ERROR >>> ' + reserve_ID)
            print("    RESERVE ERROR >>> " + reserve_ID)
            
        #LIMPIAR LA DB
        conn = sqlite3.connect(DB_Path)
        conn.execute("DELETE FROM T_Reserves")
        conn.commit()
        conn.close()
                
        update_reserves_database()

    except Exception as e:
        e = traceback.format_exc()
        logging.error('Error:' + e)
        print(str(rep) + " rep >>> ************************ERROR************************")
        print(str(rep) + " rep >>> make_reserve_main")
        
    finally:
        browser.quit()
        final_slot = "None"
    
        
def check_reserves():
    for r_reserve_ID in T_Reserves["Activity_ID"]:
        r_reserve_status = T_Reserves["Reserve_Status"][T_Reserves["Activity_ID"].index(r_reserve_ID)]
        r_reserve_time = datetime.datetime.strptime(T_Reserves["Reserve_Start_Date"][T_Reserves["Activity_ID"].index(r_reserve_ID)], "%d/%m/%Y %H:%M")
        #if (r_reserve_time + datetime.timedelta(seconds=-Start_Reserve_Time_In_Advance) < datetime.datetime.now()) and (r_reserve_time + datetime.timedelta(seconds=180) > datetime.datetime.now()) and (r_reserve_status != "RESERVED") and (r_reserve_status != "FULL"):
        if (r_reserve_time + datetime.timedelta(seconds=-Start_Reserve_Time_In_Advance) < datetime.datetime.now()) and (r_reserve_status != "RESERVED") and (r_reserve_status != "FULL") and (r_reserve_status != "INCOMPATIBLE"):
            r_Day = r_reserve_ID.split("_", 3)[0]
            r_Start_Time = r_reserve_ID.split("_", 3)[1]
            r_Activity = r_reserve_ID.split("_", 3)[2]
            
            make_reserve_main(r_reserve_ID, r_Day, r_Start_Time, r_Activity)


def start_loop():

    rep = Initial_Rep
    while rep >= 0:
    #for rep in range(0, 5):
    
        try:
            #print(str(rep) + " rep >>> " + str(datetime.datetime.now().isoformat(sep=' ', timespec='seconds')))
            
            if float(rep) % Check_Reserves_Timer == 0:
                print(str(rep) + " rep >>> CHECK RESERVES >>> " + str(datetime.datetime.now().isoformat(sep=' ', timespec='seconds')))
                check_reserves()
                #print(str(rep) + " rep >>> CHECK RESERVES DONE>>> " + str(datetime.datetime.now().isoformat(sep=' ', timespec='seconds')))#######
                
            if float(rep) % Download_Schedule_Timer == 0:
                print(str(rep) + " rep >>> DOWNLOAD SCHEDULE >>> " + str(datetime.datetime.now().isoformat(sep=' ', timespec='seconds')))
                get_schedule_main()
                #print(str(rep) + " rep >>> DOWNLOAD SHEDULE DONE>>> " + str(datetime.datetime.now().isoformat(sep=' ', timespec='seconds')))#######
                
            if float(rep) % Download_Reserves_Timer == 0:
                print(str(rep) + " rep >>> DOWNLOAD RESERVES DB >>> " + str(datetime.datetime.now().isoformat(sep=' ', timespec='seconds')))
                load_reserves_database()
                #print(str(rep) + " rep >>> DOWNLOAD RESERVES DONE>>> " + str(datetime.datetime.now().isoformat(sep=' ', timespec='seconds')))#######
  
            if ((Init_Time + datetime.timedelta(minutes=15)) < (datetime.datetime.now())) and ((datetime.datetime.now() > Server_Shutdown_Time) or (datetime.datetime.now() < Server_Start_Time)):
                server_shutdown()
                break


            #ESPERAR X SEC
            time.sleep(Loop_Timer)
            
        except Exception as e:
            e = traceback.format_exc()
            logging.error('Error:' + e)
            print(str(rep) + " rep >>> ************************ERROR************************")
            print(str(rep) + " rep >>> start_loop")
            
        rep += 1
        

def server_shutdown():
    logging.info('SERVER SHUTDOWN')
    send_mail(Admin_Email, "GoFit_Tool SHUTTING DOWN", "Server shutting down at " + datetime.datetime.now().isoformat(sep=' ', timespec='seconds'))
    #subprocess.call([Shutdown_Path])
    os.system("C:\Windows\System32\cmd.exe /c " + Shutdown_Path)


def main():

    sys.excepthook = error_handler

    init_variables()
    
    print("INIT_TIME >>> " + str(datetime.datetime.now().isoformat(sep=' ', timespec='seconds')))
    logging.info('SCRIPT START')

    send_mail(Admin_Email, "GoFit_Tool STARTING", "GoFit_Tool has started at " + datetime.datetime.now().isoformat(sep=' ', timespec='seconds'))
    
    start_loop()

    print("END_TIME >>> " + str(datetime.datetime.now().isoformat(sep=' ', timespec='seconds')))
    logging.info('SCRIPT END')


if __name__ == "__main__":
     main()