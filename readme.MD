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
    