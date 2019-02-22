# -*- coding: utf-8 -*-
"""

Windows Service Monitor (tested \w Python 3.7 on Win7x64 SP1)
(Author Mathew.Brown.mls@gmail.com)

I created this program to monitor/control changes to Windows services for a 
Win 7 x64 machine.  I was tired of finding new services quietly installed
via group policy domain objects, some of which caused performance problems
or scanned my home network when I worked out of my home office.

Better to know what your machine is running than not!

This program is designed to run as a windows service.  It enumerates
all windows services periodically, saves them to a SQLite db file as well as
a state table.  Changes queue updates to the SQLite db, as well as output
to a plain text log file.  If 'forceexpectedstate' is set to 'yes' in the
db for a given service, this program will try to start/stop the service as
appropriate.

To control the behavior of this program/service, you will need a SQLite db
editor.  By default, any new services found default to 'ignore_this_service' =
'yes'.  Change this to NULL in order to log unexpected running states or to
automatically stop/start the service if 'forceexpectedstate' = 'yes' and
'expectedstate' is set to the desired running/nonrunning state.

***

To compile python -> service .exe
- download/extract Winpython 3.7
- Set system environment variables PYTHONHOME and PYTHONPATH
    set PYTHONHOME=C:\Scripts\python\WinPython-64bit-3.6.2.0Qt5\python-3.6.2.amd64\
    set PYTHONPATH=C:\Scripts\python\WinPython-64bit-3.6.2.0Qt5\python-3.6.2.amd64\Lib\
- Make sure pyinstaller available:
    pip install pyinstaller
- compile to .exe (run in python cmd in folder \w this file)
    pyinstaller -F --hidden-import=win32timezone windows_service_monitor_svc.py
- .exe here:
    dist\windows_service_monitor_svc.exe
                   
To install service:
    
open command prompt (as administrator)
ex: windows_service_monitor_svc.exe install

You need to set the location of the sqlite db file and log file in the windows
registry (from cmd):
    
REG ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\Windows Service Monitor" /v sqlite_dbfile /d "C:\scripts\w32services.db"
REG ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\Windows Service Monitor" /v logfile /d "C:\scripts\w32services.log"

Navigate to windows services, adjust service to start on start-up!
    
"""

#sqlite_dbfile = 'C:\\scripts\\w32services.db'
#logfile = 'C:\\scripts\\w32services.log'
check_services_interval = 5     # in seconds
logging = 1                     # only for debugging

import win32event
import win32api
import win32con
import win32service
import win32serviceutil
import servicemanager
import os
import sqlite3
import datetime
import time
import sys
import socket


""" Setup some constants in dictionaries from win32service """
serviceStates = {0:'Unknown',
                 win32service.SERVICE_STOP:'SERVICE_STOP',
                 win32service.SERVICE_STOPPED:'SERVICE_STOPPED',
                 win32service.SERVICE_START_PENDING:'SERVICE_START_PENDING',
                 win32service.SERVICE_START:'SERVICE_START',
                 win32service.SERVICE_RUNNING:'SERVICE_RUNNING'
                 }
serviceTypes = {0:'Unknown',
                win32service.SERVICE_KERNEL_DRIVER:'SERVICE_KERNEL_DRIVER',
                win32service.SERVICE_FILE_SYSTEM_DRIVER:'SERVICE_FILE_SYSTEM_DRIVER',
                win32service.SERVICE_INTERACTIVE_PROCESS:'SERVICE_INTERACTIVE_PROCESS',
                win32service.SERVICE_DRIVER:'SERVICE_DRIVER',
                win32service.SERVICE_WIN32:'SERVICE_WIN32',
                win32service.SERVICE_WIN32_OWN_PROCESS:'SERVICE_WIN32_OWN_PROCESS',
                win32service.SERVICE_WIN32_SHARE_PROCESS:'SERVICE_WIN32_SHARE_PROCESS'  
                }

""" Tries to write to log file + db log table """
def write_to_log(logentry, service_short_name, logfile, sqlite_dbfile):
    try:
        F = open(logfile,'a')
        F.write(logentry + '\n')
        F.close()    
    except:
        servicemanager.LogErrorMsg("Error, cannot open logfile: " + logfile)
    try:    
        conn = sqlite3.connect(sqlite_dbfile)
        c = conn.cursor()
        
        sql = "INSERT INTO win32service_log "
        sql += "(win32service_short_name, logentry, established, established_by) "
        sql += " VALUES ("
        sql += "'" + service_short_name + "', "
        sql += "'" + logentry + "', "
        sql += "" + "datetime('now','localtime')" + ", "
        sql += "'" + "windows_service_monitor_svc.py" + "')"
        
        c.execute(sql)
        conn.commit()
        conn.close()
    except:
        servicemanager.LogErrorMsg("Error, cannot add log entry to db: " + sqlite_dbfile)

""" Checks windows registry keys for values to use for sqlite_dbfile and logfile """
def init_local_vars():
    # sqlite_dbfile is in HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\Windows Service Monitor    
    access = win32con.KEY_READ | win32con.KEY_ENUMERATE_SUB_KEYS | win32con.KEY_QUERY_VALUE
    hkey_base = "SYSTEM\\CurrentControlSet\\Services"
    hkey_key = "\\" + "Windows Service Monitor"
    try:
        hkey = win32api.RegOpenKey(win32con.HKEY_LOCAL_MACHINE, (hkey_base + hkey_key), 0, access)
        regkey = str(win32api.RegQueryValueEx(hkey, "sqlite_dbfile")[0])
        sqlite_dbfile = regkey
    except:
        servicemanager.LogErrorMsg("Cannot open regkey: " + hkey_base + hkey_key + "\\sqlite_dbfile")
        return "", "", 0
    
    # logfile is in HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\Windows Service Monitor    
    access = win32con.KEY_READ | win32con.KEY_ENUMERATE_SUB_KEYS | win32con.KEY_QUERY_VALUE
    hkey_base = "SYSTEM\\CurrentControlSet\\Services"
    hkey_key = "\\" + "Windows Service Monitor"
    try:
        hkey = win32api.RegOpenKey(win32con.HKEY_LOCAL_MACHINE, (hkey_base + hkey_key), 0, access)
        regkey = str(win32api.RegQueryValueEx(hkey, "logfile")[0])
        logfile = regkey
    except:
        servicemanager.LogErrorMsg("Cannot open regkey: " + hkey_base + hkey_key + "\\logfile")
        return "", "", 0
    
    # Verify sqlite_dbfile exists, if not, initialize
    try:
        fh = open(sqlite_dbfile, 'r')
        fh.close()
        # Store configuration file values
    except:
        servicemanager.LogErrorMsg("sqlite_dbfile does not exist!  Attempting to initialize new file")
        init_w32services_db(sqlite_dbfile)
    try:
        conn = sqlite3.connect(sqlite_dbfile)
        c = conn.cursor()
        conn.close()
        # Store configuration file values
    except:
        servicemanager.LogErrorMsg("Cannot initialize/open sqlite_dbfile: " + sqlite_dbfile)
        return "", "", 0
    
    # Verify we can write to logfile
    try:
        F = open(logfile,'a')
        F.close()    
    except:
        servicemanager.LogErrorMsg("Cannot append to file: " + logfile)
        return "", "", 0
    
    return sqlite_dbfile, logfile, 1

""" Loads dict of dicts statetable {{}} from data in sqlite_dbfile """
def read_state_table_from_db_file(statetable, logfile, sqlite_dbfile):
    try:
        conn = sqlite3.connect(sqlite_dbfile)
        c = conn.cursor()
        
        sql = "SELECT short_name, description, laststate, expectedstate, forceexpectedstate, servicetype, "
        sql += "ImagePath, ObjectName, ignore_this_service, established, established_by "
        sql += "FROM win32service"
        #print(sql)
        
        c.execute(sql)
        rows = c.fetchall()
        total_rows = 0
        for row in rows:
            #print(str(row))
            statetable[row[0]] = {}
            statetable[row[0]]['service_short_name'] = row[0]
            statetable[row[0]]['service_description'] = row[1]
            statetable[row[0]]['laststate'] = row[2]
            statetable[row[0]]['expectedstate'] = row[3]
            statetable[row[0]]['forceexpectedstate'] = row[4]
            statetable[row[0]]['servicetype'] = row[5]
            statetable[row[0]]['ImagePath'] = row[6]
            statetable[row[0]]['ObjectName'] = row[7]
            statetable[row[0]]['ignore_this_service'] = row[8]
            statetable[row[0]]['established'] = row[9]
            statetable[row[0]]['established_by'] = row[10]
            #print(str(row[4]));
            total_rows = total_rows + 1
            
        write_to_log(("Loaded " + str(total_rows) + " definitions from " + sqlite_dbfile), "", logfile, sqlite_dbfile)
        conn.commit()
        conn.close()
        return statetable
    except:
        servicemanager.LogErrorMsg("Failed to load statetable from db file")

""" Checks the state table to see if we need to attempt to start or stop
    a service that is not in the epectedstate """
def force_state_if_necessary(statetable, short_name, serviceState, logfile, sqlite_dbfile):
    if (statetable[short_name]['forceexpectedstate'] != None):
        if (statetable[short_name]['forceexpectedstate'] == 'yes'): 
           # We really only worry about SERVICE_RUNNING and SERVICE_STOPPED
           if ((serviceState == 'SERVICE_RUNNING') and (statetable[short_name]['expectedstate'] == 'SERVICE_STOPPED')):
                write_to_log(("Attempting to stop service: " + short_name), short_name, logfile, sqlite_dbfile)
                try:
                    win32serviceutil.StopService(short_name)
                except:
                    write_to_log(("Failed to stop service: " + short_name), short_name, logfile, sqlite_dbfile)
           if ((serviceState == 'SERVICE_STOPPED') and (statetable[short_name]['expectedstate'] == 'SERVICE_RUNNING')):
                write_to_log(("Attempting to start service: " + short_name), short_name, logfile, sqlite_dbfile)
                try:
                    win32serviceutil.StartService(short_name)
                except:
                    write_to_log(("Failed to start service: " + short_name), short_name, logfile, sqlite_dbfile)
                    
""" Routine to create a new sqlite db file/schema """
def init_w32services_db(sqlite_dbfile):
    try:
        conn = sqlite3.connect(sqlite_dbfile)
        c = conn.cursor()
        
        sql = 'CREATE TABLE IF NOT EXISTS '
        sql = sql + "'" + 'win32service' + "' ("
        sql = sql + "'win32service_key' INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL UNIQUE, "
        sql = sql + "'short_name' TEXT, "
        sql = sql + "'description' TEXT, "
        sql = sql + "'laststate' TEXT, "
        sql = sql + "'expectedstate' TEXT, "
        sql = sql + "'forceexpectedstate' TEXT, "
        sql = sql + "'servicetype' TEXT, "
        sql = sql + "'servicestarttype' TEXT, "
        sql = sql + "'serviceerrorcontrol' TEXT, "
        sql = sql + "'ImagePath' TEXT, "
        sql = sql + "'ObjectName' TEXT, "
        sql = sql + "'ignore_this_service' TEXT, "
        sql = sql + "'notes' TEXT, "
        sql = sql + "'established' DATETIME, "
        sql = sql + "'established_by' TEXT, "
        sql = sql + "'edited' DATETIME, "
        sql = sql + "'edited_by' TEXT, "
        sql = sql + "'inactivated' DATETIME, "
        sql = sql + "'inactivated_by' TEXT"
        sql = sql + ");"
        
        # print(sql)
        c.execute(sql)
        
        sql = 'CREATE TABLE IF NOT EXISTS '
        sql = sql + "'" + 'win32service_log' + "' ("
        sql = sql + "'win32service_log_key' INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL UNIQUE, "
        sql = sql + "'win32service_short_name' TEXT, "
        sql = sql + "'logentry' TEXT, "
        sql = sql + "'established' DATETIME, "
        sql = sql + "'established_by' TEXT"
        sql = sql + ");"
        
        # print(sql)
        c.execute(sql)
        
        conn.commit()
        conn.close()
    except:
        servicemanager.LogErrorMsg("Error creating: " + sqlite_dbfile)

    
""" routine to use win32service/api/etc to check current status of windows
    services against statetable{} 
    
    services with forceexpectedstate = 'yes' will cue force_state_if_necessary()
    """
def check_services(statetable, logfile, sqlite_dbfile):
    accessSCM = win32con.GENERIC_READ
    #Open Service Control Manager
    hscm = win32service.OpenSCManager(None, None, accessSCM)
    #Enumerate Service Control Manager DB
    typeFilter = win32service.SERVICE_WIN32
    stateFilter = win32service.SERVICE_STATE_ALL
    statuses = win32service.EnumServicesStatus(hscm, typeFilter, stateFilter)
    for (short_name, desc, status) in statuses:
        if short_name in statetable:
            # if the service is recorded in the statetable, check to see
            # whether ServiceState (running/stopped/etc) matches expectedstate or not
            if status[1] in serviceStates:
                serviceState = serviceStates[status[1]]
            else:
                serviceState = status[1]
            if (statetable[short_name]['ignore_this_service'] != None):
                # do nothing if it's on the ignore list
                pass
            elif (serviceState != statetable[short_name]['expectedstate']):
                logentry = 'Windows service: ' + short_name + '(' + statetable[short_name]['service_description'] + ') is ' + serviceState + ' - not in expectedstate (' + statetable[short_name]['expectedstate'] + ')'
                write_to_log(logentry, short_name, logfile, sqlite_dbfile)
                force_state_if_necessary(statetable, short_name, serviceState, logfile, sqlite_dbfile)
        else:
            # if not in the statetable, must be a new service, create in
            # statetable, log to file/db
            if status[1] in serviceStates:
                serviceState = serviceStates[status[1]]
            else:
                serviceState = status[1]
            
            if status[0] in serviceTypes:
                serviceType = serviceTypes[status[0]]
            else:
                serviceType = status[1]
            # define the new service
            new_svc = {}
            new_svc['service_short_name'] = short_name
            new_svc['service_description'] = desc
            new_svc['expectedstate'] = serviceState # We default to taking new services @ expected state
            new_svc['laststate'] = serviceState
            new_svc['servicetype'] = serviceType
            new_svc['ignore_this_service'] = 'yes'
            
            # Some facets of services only seem to be available via the registry
            # we try and get the executable + run-with-credentials here
            access = win32con.KEY_READ | win32con.KEY_ENUMERATE_SUB_KEYS | win32con.KEY_QUERY_VALUE
            hkey_base = "SYSTEM\\CurrentControlSet\\Services"
            hkey_key = "\\" + short_name
            try:
                hkey = win32api.RegOpenKey(win32con.HKEY_LOCAL_MACHINE, (hkey_base + hkey_key), 0, access)
                regkey = str(win32api.RegQueryValueEx(hkey, "ImagePath")[0])
                new_svc['ImagePath'] = regkey
            except:
                new_svc['ImagePath'] = ''
            try:
                hkey = win32api.RegOpenKey(win32con.HKEY_LOCAL_MACHINE, (hkey_base + hkey_key), 0, access)
                regkey = str(win32api.RegQueryValueEx(hkey, "ObjectName")[0])
                new_svc['ObjectName'] = regkey
            except:
                new_svc['ObjectName'] = ''
            
            # Expand out any env variables used in the impagepath values
            check_for_os_environ = new_svc['ImagePath'].split('%')[0]
            if (check_for_os_environ is ''):
                try:
                   envvar = os.environ[new_svc['ImagePath'].split('%')[1]]
                   rest_of_path = new_svc['ImagePath'].split('%')[2]
                   new_svc['ImagePath'] = (envvar + rest_of_path)
                except:
                    pass
                
            statetable[short_name] = new_svc
            
            # Send new service to db
            add_new_service_to_db(new_svc, sqlite_dbfile, logfile)
            
            # Send output to the log file
            curtime = datetime.datetime.isoformat(datetime.datetime.now())
            log_message = curtime + ' : New windows service discovered : ' + short_name
            write_to_log(log_message, short_name, logfile, sqlite_dbfile)
            log_message = curtime + ' :    - ' + desc
            write_to_log(log_message, short_name, logfile, sqlite_dbfile)
            if ('ImagePath' in new_svc):
                log_message = curtime + ' :    - path: ' + new_svc['ImagePath']
                write_to_log(log_message, short_name, logfile, sqlite_dbfile)
            if ('ObjectName' in new_svc):
                log_message = curtime + ' :    - running as user: ' + new_svc['ObjectName']
                write_to_log(log_message, short_name, logfile, sqlite_dbfile)
    return statetable

""" fix this! """
def add_new_service_to_db(new_svc, sqlite_dbfile, logfile):
    
    try:
        conn = sqlite3.connect(sqlite_dbfile)
        c = conn.cursor()
        sql = "INSERT INTO win32service (short_name, description, laststate, "
        sql += "expectedstate, servicetype, ImagePath, ObjectName, established, "
        sql += "established_by, ignore_this_service) VALUES ("
        sql += "'" + new_svc['service_short_name'].replace("'", "''") + "', "
        sql += "'" + new_svc['service_description'].replace("'", "''") + "', "
        sql += "'" + new_svc['laststate'].replace("'", "''") + "', "
        sql += "'" + new_svc['expectedstate'].replace("'", "''") + "', "
        sql += "'" + str(new_svc['servicetype']).replace("'", "''") + "', "
        sql += "'" + str(new_svc['ImagePath']).replace("'", "''").replace('"', '') + "', "
        sql += "'" + str(new_svc['ObjectName']).replace("'", "''") + "', "
        sql += "" + "datetime('now','localtime')" + ", "
        sql += "'" + "windows_service_monitor" + "', "
        sql += "'" + new_svc['ignore_this_service'].replace("'", "''") + "')"
        c.execute(sql)
        conn.commit()
        conn.close()
    except:
        write_to_log("Error adding new Service to db", short_name, logfile, sqlite_dbfile)

class windows_service_monitor(win32serviceutil.ServiceFramework):
    _svc_name_ = "Windows Service Monitor"
    _svc_display_name_ = "Windows Service Monitor"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        rc = None
        x = 0
        # Each time the service is started we check for registry config vars
        sqlite_dbfile, logfile, regvar_success = init_local_vars()
        if (regvar_success == 0):
            servicemanager.LogErrorMsg("sqlite_dbfile or logfile registry key error.  Please verify these keys are set to valid file locations.")
            self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
            win32event.SetEvent(self.hWaitStop)
            return
        # Each time the service is started we init the statetable from the db
        statetable = {}
        start_time = datetime.datetime.isoformat(datetime.datetime.now())
        write_to_log(("*** Starting Windows Service Monitor @ " + str(start_time)), "", logfile, sqlite_dbfile)
        
        statetable = read_state_table_from_db_file(statetable, logfile, sqlite_dbfile)
    
        while rc != win32event.WAIT_OBJECT_0:
            if ((x % check_services_interval) == 0):
                """ Put in all fucntional code here! """
                check_services(statetable, logfile, sqlite_dbfile)
            if x == 60:
                x = 0;
            time.sleep(1)
            x = x + 1
            rc = win32event.WaitForSingleObject(self.hWaitStop, 0)
        
        try:
            end_time = datetime.datetime.isoformat(datetime.datetime.now())
            write_to_log(("*** Ending Windows Service Monitor @ " + str(end_time)), "", logfile, sqlite_dbfile)
        except:
            pass
        
if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(windows_service_monitor)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        #print(str(win32api.FindFiles(fname)))
        win32serviceutil.HandleCommandLine(windows_service_monitor)