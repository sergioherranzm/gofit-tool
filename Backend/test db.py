import datetime
import os
import traceback
from pymongo import MongoClient
from jproperties import Properties

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
    global DB_URL
    global DB_Name
    global Application_Path
    Application_Path = os.path.dirname(os.path.abspath(__file__))
    
    
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
    
    
    
    
    
    
    DB_URL = configs.get("DB_URL").data
    DB_Name = configs.get("DB_Name").data
    test_database()
    
         

def test_database():
    global DB
    global DB_Reserves
    global DB_Activities
    
    #https://zetcode.com/python/pymongo/
    #https://pymongo.readthedocs.io/en/stable/index.html#
    
        
    try:

        client = MongoClient(DB_URL)
        
        DB = client[DB_Name]
        
        #print(DB.collection_names())
        
        DB_Reserves = DB.categories
        DB_Activities = DB.posts
        
        cat = DB_Reserves.find_one()
        
        print(cat)

        

        
    except:
        e = traceback.format_exc()
        print('Error:' + e)


        

def main():
    
    print('SCRIPT START')

    init_variables()
    
    


    print('SCRIPT END')


if __name__ == "__main__":
     main()