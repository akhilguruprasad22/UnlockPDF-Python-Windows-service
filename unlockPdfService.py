import os
import sys
import win32serviceutil
import servicemanager
import win32event
import win32service

import win32file
import win32con

import pikepdf


class UnlockPdfService(win32serviceutil.ServiceFramework):

    _svc_name_ = "UnlockPdfService"
    _svc_display_name_ = "Unlock PDF Service"
    _svc_description_ = "This service monitors the set UPDF_MONITOR_PATH environment variable or C:/Documents/Unlock PDF folder(Default) and unlocks newly added PDFs."

    _monitor_default_path_ = os.path.normpath('C:/Documents/Unlock PDF')
    _monitor_directory_path_ = os.path.normpath(os.environ.get("UPDF_MONITOR_PATH", _monitor_default_path_))

    def __init__(self, args):
        try:
            win32serviceutil.ServiceFramework.__init__(self, args)

            if not os.path.isdir(self._monitor_directory_path_):
                os.makedirs(self._monitor_directory_path_)

            self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
            self.file_change_handle = win32file.FindFirstChangeNotification(
                self._monitor_directory_path_,
                True,
                win32con.FILE_NOTIFY_CHANGE_FILE_NAME)
            self.isAlive = True
        
        except Exception as ex:
            self.log_message("Error initializing service: {}".format(ex))

    def SvcStop(self):
        try:
            self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
            win32event.SetEvent(self.hWaitStop)

        except Exception as ex:
            self.log_message("Error stopping service: {}".format(ex))

    def SvcDoRun(self):
        try:
            self.isAlive = True

            servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                                    servicemanager.PYS_SERVICE_STARTED,
                                    (self._svc_name_, "\nMonitoring path: " + self._monitor_directory_path_))    
            self.main()

        except Exception as ex:
            self.log_message("Error running service: {}".format(ex))

    def main(self):
        try:
            self.old_path_contents = self.getPdfFiles(self._monitor_directory_path_)

            while self.isAlive:
                monitoring_result = win32event.WaitForSingleObject(self.file_change_handle, 1000)
                result = win32event.WaitForSingleObject(self.hWaitStop, 1000)
                
                if monitoring_result == win32con.WAIT_OBJECT_0:
                    self.new_path_contents = self.getPdfFiles(self._monitor_directory_path_)
                    self.added_pdfs = [file for file in self.new_path_contents if file not in self.old_path_contents]

                    self.unlockPdfs(self.added_pdfs)

                    self.old_path_contents = self.new_path_contents
                    win32file.FindNextChangeNotification(self.file_change_handle)

                if result == win32event.WAIT_OBJECT_0:
                    self.isAlive = False

        except Exception as ex:
            self.log_message("Error running main service: {}".format(ex))

        finally:
            win32file.FindCloseChangeNotification(self.file_change_handle)

    def getPdfFiles(self, path):
        try:
            pdfs = []

            for dirPath, folders, files in os.walk(path):
                for file in files:
                    if file.endswith(".pdf") and not file.endswith("_unlocked.pdf"):
                        pdfs.append(os.path.join(os.path.normpath(dirPath), file))
            return pdfs

        except Exception as ex:
            self.log_message("Error fetching PDFs: {}".format(ex))

    def unlockPdfs(self, pdfs = []):
        try:
            for pdf in pdfs:
                try:
                    with pikepdf.open(pdf, allow_overwriting_input=True) as openPdf:
                        openPdf.save(os.path.splitext(pdf)[0]+"_unlocked.pdf")
                except pikepdf.PasswordError as pikePwdEx:
                    self.log_message("Error unlocking PDF: {}".format(pikePwdEx), type = "WARN")
                except pikepdf.PdfError as pikeEx:
                    self.log_message("Error unlocking PDF: {}".format(pikeEx), type = "WARN")
        except Exception as ex:
                    self.log_message("Error unlocking PDFs: {}".format(ex))

    def log_message(self, message, type = "ERROR"):
        try:
            if type == "INFO":
                servicemanager.LogInfoMsg(message)
            elif type == "ERROR":
                self.SvcStop()
                servicemanager.LogMsg(servicemanager.EVENTLOG_ERROR_TYPE,
                                servicemanager.PYS_SERVICE_STOPPED,
                                (self._svc_name_, "\n"+message))
            elif type == "WARN":
                servicemanager.LogWarningMsg(message)
        
        except Exception as ex:
            print("Error logging message: {}".format(ex))

if __name__ == '__main__':
    try:
        if len(sys.argv) == 1:
            servicemanager.Initialize()
            servicemanager.PrepareToHostSingle(UnlockPdfService)
            servicemanager.StartServiceCtrlDispatcher()
        else:
            win32serviceutil.HandleCommandLine(UnlockPdfService)

    except Exception as ex:
        print("Error: {}".format(ex))
