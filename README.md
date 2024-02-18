# UnlockPDF Python Windows service
A windows service which monitors a given directory and it's subdirectories for file changes and subsequently unlocks (removes set access restrictions) newly added PDFs.

> [!NOTE]
> The service creates and monitors, by default, the **C:/Documents/Unlock PDF** directory
> 
> The directory to be monitored can also be set as a system environment variable under the name: **UPDF_MONITOR_PATH**

## Requirements
- Python 3.*
- Visual C++ Build Tools 2015
- Pyinstaller 6.3

## Build into an executable
```
pyinstaller -F --hidden-import=win32timezone unlockPdfService.py
```
The executable will now be present under a new directory called **dist**

## Run the service

> [!IMPORTANT]
> The user needs to have Administrative rights(Administrator privileges) to install the service into the Windows Service Manager

### Installation
```
dist\unlockPdfService.exe install
```
One can view the installed service in the **Services** Windows utility

### Start the service
```
dist\unlockPdfService.exe start
```
This can also be achieved through the **Services** utility

### Stop the service
```
dist\unlockPdfService.exe stop
```

### Remove the service
```
dist\unlockPdfService.exe remove
```

### Debug the service
```
dist\unlockPdfService.exe debug
```

## References
- Python Programming On Win32  *By Andy Robinson, Mark Hammond*
- [Time Golden's Python Stuff: Watch a Directory for Changes](https://timgolden.me.uk/python/win32_how_do_i/watch_directory_for_changes.html)
