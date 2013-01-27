#!/usr/bin/python
# -*- coding: utf-8 -*-
# watchfolder.py

import os
import os.path
import shutil
import time
import errno
from settings import Settings
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer
import sip
sip.setapi('QVariant', 2)
from PyQt4 import QtCore
from PyQt4.QtCore import pyqtSlot, pyqtSignal

class FolderWatcher(QtCore.QObject):
    """
    Watches the webprint folder for changes.
    Creates the webprint folder if it does not exist.
    emits:
        addJob(path, printer)
    """
    addJob = pyqtSignal(str, str)
    def __init__(self, parent=None):
        super(FolderWatcher, self).__init__(parent)
        settings = Settings()
        self.event_handler = _NewFileEventHandler()
        self.event_handler.addJob.connect(self.addJob)
        if not os.path.exists(settings.getWebprintFolder()):
            self.build_webprint_dir()
        self.observer = Observer()
        self.observer.start()
        self.start_watching()

    def stop_watching(self):
        """Stops folder watching.

        Detaches all watches. To actually stop the daemon thread, call
        self.observer.stop_watching()
        """
        self.observer.unschedule_all()

    def start_watching(self):
        """Starts folder watching.

        Loads folders to watch from settings.
        """
        settings = Settings()
        webprint_dir = str(settings.getWebprintFolder())
        for (name, location) in settings.getInstalledPrinters():
            path = os.path.join(webprint_dir, str(location))
            self.observer.schedule(self.event_handler, path)

    @pyqtSlot()
    def build_webprint_dir(self):
        """Create print folders in the 'webprint' directory.

        Delete files and folders not listed in settings.
        Creates the webprint directory if it does not exist.
        Folder watching is temporarily stopped during this operation.
        """
        self.stop_watching()
        settings = Settings()
        webprint_dir = settings.getWebprintFolder()
        if not os.path.exists(webprint_dir):
            os.makedirs(webprint_dir)
        folders = [location for (name, location) in
                    settings.getInstalledPrinters()]
        ls = os.listdir(webprint_dir)
        for folder in folders:
            if folder not in ls:
                os.mkdir(os.path.join(webprint_dir, folder))
            else:
                ls.remove(folder)
        for item in ls:
            shutil.rmtree(os.path.join(webprint_dir, item))
        self.start_watching()

    @pyqtSlot(str)
    def move_webprint_dir(self, dest):
        """Moves the webprint folder to a new location.

        Destination folder should not exist.
        If it does exist, it should be empty.

        If the webprint folder does not exist, creates it.
        Folder watching is temporarily stopped during this operation.
        """
        dest = str(dest)
        settings = Settings()
        webprint_dir = settings.getWebprintFolder()
        if os.path.exists(dest):
            #TODO: Handel errors on this delete
            shutil.rmtree(dest)
        os.mkdir(dest)
        if os.path.exists(webprint_dir):
            self.stop_watching()
            for f in os.listdir(webprint_dir):
                #TODO: Handel errors here
                shutil.move(os.path.join(webprint_dir, f), dest)
            os.rmdir(webprint_dir)
            settings.setWebprintFolder(dest)
            self.start_watching()
        else:
            #webprint folder does not exist
            settings.setWebprintFolder(dest)
            self.build_webprint_dir()


class _NewFileEventHandler(FileSystemEventHandler):
    """
    Emits addJob(path, printer)
    """
    def __init__(self):
        super(_NewFileEventHandler, self).__init__()
        self._mediator = _EventHandlerMediator()

    def on_created(self, event):
        """Called when a file or directory is created.

        Adds files to the print queue; does not check their filetype.
        Filetype checking is the responsibility of the printer.
        """
        if not event.is_directory:
            try:
                f = os.open(event.src_path, os.O_RDONLY | os.O_BINARY)
            except OSError as e:
                if e.errno == errno.ENOENT:
                    return
                elif e.errno == errno.EACCES:
                    #sleep until it can be accessed
                    time_slept = 0
                    sleep_duration = 0.5
                    while(True):
                        if time_slept > 30:
                            #file won't unlock, send job anyway and 
                            #allow it to fail with an error
                            break
                        time.sleep(sleep_duration)
                        time_slept += sleep_duration
                        try:
                            f = os.open(event.src_path, os.O_RDONLY |
                                        os.O_BINARY)
                        except OSError as e:
                            if e.errno == errno.ENOENT:
                                return
                        else:
                            os.close(f)
                            break
            else:
                os.close(f)
            printer = os.path.split(os.path.split(event.src_path)[0])[1]
            self.addJob.emit(event.src_path, printer)

    def addJob():
        def fget(self):
            return self._mediator.addJob
        return locals()
    addJob = property(**addJob())


class _EventHandlerMediator(QtCore.QObject):
    addJob = pyqtSignal(str, str) #path, printer
    def __init__(self):
        super(_EventHandlerMediator, self).__init__()