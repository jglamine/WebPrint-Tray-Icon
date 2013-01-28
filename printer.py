#!/usr/bin/python
# -*- coding: utf-8 -*-
# printer.py

import papercut
import time
import os
import os.path
import json
from settings import Settings
from requests.exceptions import ConnectionError

import sip
sip.setapi('QVariant', 2)

from PyQt4.QtCore import pyqtSlot, pyqtSignal, QThreadPool
from PyQt4 import QtCore


class ThreadedPrinter(QtCore.QObject):
    """Contains the user facing API for communicating with webprint.

    messages:
        See source code.
    slots:
        onReceiveCredentials(username, password)
            logs in with the given credentials
        login() - login with credentials from config file.
        login(username, password) - login with the given credentials.
        logout() - logout
    """
    #TODO: Split this class into ThreadedPrinter and PrintSpooler - Refactor
    jobNeedsLogin = pyqtSignal(str, str) #filename, printer
    jobAdded = pyqtSignal(str, str) #filename, printer
    jobUploaded = pyqtSignal(str, str) #filename, printer
    jobStatusUpdate = pyqtSignal(str, str, str) #filename, printer, json dict of status information
    jobFinished = pyqtSignal(str, str, str) #filename, printer, json dict of status information
    jobFailed = pyqtSignal(str, str, str) #filename, printer, Error message
    loginComplete = pyqtSignal(str) #sessionId - also used for icon status
    loginFailed = pyqtSignal(str) #error code - NoSavedPassword, AuthError, UnexpectedError, ext...
    RequestCredentials = pyqtSignal()
    connectionLost = pyqtSignal() #used to update icon status to failed
    connectionWarning = pyqtSignal() #update icon to warning - connection may or may not be lost
    unusedSessionId = pyqtSignal(str) #unusedSessionId
    updateToolTip = pyqtSignal() #sent to signal that the tooltip should be updated
    loginRequired = pyqtSignal() #sent to prompt to user to sign in

    def __init__(self, url=None, test=False, parent=None):
        """args:
        """
        super(ThreadedPrinter, self).__init__(parent)
        self.url = url
        self.runningJobs = {}
        self.savedJobs = [] # jobs waiting to print on login
        #self.balance = None #number of cents (int)
        self.sessionId = None
        self.sessionExpired = True
        #TODO: Turn sessionExpired into a function which uses the current time to check if the session expired.
        self.test = test

    def addJob(self, filepath, printerName):
        """Add a print job to the queue.
        """
        settings = Settings()
        filepath = str(filepath)
        printerName = str(printerName)
        extension = filepath.split('.')[-1].lower()
        if extension not in papercut.VALID_FILETYPES:
            self.jobFailed.emit(os.path.split(filepath)[1], printerName,
                                'Invalid Filetype: WebPrint does not support \
                                files of type \'%s\'.' % extension)
            return
        job = {'filepath':filepath, 'filename':os.path.split(filepath)[1],
        'printerName':printerName}
        if not settings.getHasValidCredentials():
            self.jobNeedsLogin.emit(job['filename'], job['printerName'])
            self._queueJobForLogin(job)
        elif not self.sessionExpired:
            self.runningJobs[(job['filename'], job['printerName'])] = job
            self._printFile(filepath, printerName)
            self.jobAdded.emit(job['filename'], job['printerName'])
        else:
            #sign in and then try to resubmit the job
            self.loginAndPrint(job)

    def _queueJobForLogin(self, job):
        self.savedJobs.append(job)
        self.updateToolTip.emit()
        self.loginComplete.connect(self.runSavedJobs)
        self.loginFailed.connect(self.failSavedJobs)

    def loginAndPrint(self, job):
        """Attempts to sign in and print a job.
        Deletes the job if sign in fails.
        """
        self._queueJobForLogin(job)
        self.login()

    def runSavedJobs(self):
        for job in self.savedJobs:
            self.addJob(job['filepath'], job['printerName'])
        self.savedJobs = []
        self.loginComplete.disconnect(self.runSavedJobs)
        self.loginFailed.disconnect(self.failSavedJobs)

    def failSavedJobs(self):
        for job in self.savedJobs:
            self.jobFailed.emit(job['filename'], job['printerName'],
                                'Unable to connect to WebPrint. Check network connection.')
        self.savedJobs = []
        self.loginComplete.disconnect(self.runSavedJobs)
        self.loginFailed.disconnect(self.failSavedJobs)
        self.updateToolTip.emit()

    def deleteJobFile(self, filename, printerName):
        settings = Settings()
        try:
            path = os.path.join(settings.getWebprintFolder(), printerName, filename)
            os.remove(path)
        except:
            pass

    def _printFile(self, filepath, printerName):
        """Print a file to the given printer.

        Assumes we are logged in with a valid sessionId.

        messages:
            printed(filename, printerName, jobId)
            printFaild(filename, printerName)
        """
        filename = os.path.split(str(filepath))[1]
        task = _PrintFileTask(filepath, filename, printerName, self.sessionId,
                              test=self.test)
        task.finished.connect(self._onUploaded)
        task.failed.connect(self._onUploadFailed)
        QThreadPool.globalInstance().start(task)

    def monitorJob(self, filename, printerName, jobId):
        """Monitor the status of a print job.

        messages:
            jobFinished(filename, printerName, json_dict_of_status_information)
            jobStatusUpdate(json_dict_of_status_information)
            jobFailed(filename, printerName, errorCode)
        """
        task = _MonitorJobTask(filename, printerName, jobId, self.sessionId,
                               test=self.test)
        task.finished.connect(self._onJobFinished)
        task.updated.connect(self.jobStatusUpdate)
        task.failed.connect(self._onMonitorJobFailed)
        QThreadPool.globalInstance().start(task)

    def login(self, username=None, password=None):
        """Login to webprint.

        messages:
            loginComplete(sessionId)
            loginFailed()
        """
        if username is None or password is None:
            settings = Settings()
            if settings.getHasValidCredentials():
                username = settings.getUsername()
                password = settings.getPassword()
            if password is None:
                self.sessionExpired = True
                self.sessionId = None
                self.loginFailed.emit('NoSavedPassword')
                return
        if isinstance(username, QtCore.QString):
            username = unicode(username, 'utf-8')
        if isinstance(password, QtCore.QString):
            password = unicode(password, 'utf-8')
        task = _LoginTask(username, password)
        task.finished.connect(self._onLogin)
        task.failed.connect(self.loginFailed)
        QThreadPool.globalInstance().start(task)

    def logout(self):
        """Ends the current session.

            Deletes the session cookie, does not actually end the session
            or communicate with the server.
            All currently running jobs and requests to the server are still
            sent. This method only prevents more requests from being sent.
        """
        self.sessionId = None
        self.sessionExpired = True

    def requestUnusedSessionId(self):
        """Begins downloading a new, unused session id, to be used in a web
        browser.

        Because it is to be used in a browser, it could become
        invalid at any time, if the user clicks 'logout' on the website.

        Sends unusedSessionId(str) signal when the action is complete.
        If the action fails, sends unusedSessionId(str) with the empty string
        as the sessionId argument

        The current session is still kept for internal use.
        """
        if self.sessionId is None:
            self.unusedSessionId.emit('')
            return
        settings = Settings()
        username, password = None, None
        if settings.getHasValidCredentials():
            username = settings.getUsername()
            password = settings.getPassword()
        if password is None:
            self.unusedSessionId.emit('')
            return
        if isinstance(username, QtCore.QString):
            username = unicode(username, 'utf-8')
        if isinstance(password, QtCore.QString):
            password = unicode(password, 'utf-8')
        task = _LoginTask(username, password, reportErrorsAsEmptyString = True)
        task.finished.connect(self.unusedSessionId)
        task.failed.connect(self.unusedSessionId)
        QThreadPool.globalInstance().start(task)


    @pyqtSlot(str)
    def _onLogin(self, sessionId):
        self.sessionId = sessionId
        self.sessionExpired = False
        self.loginComplete.emit(sessionId)

    def _onError(self, errorCode):
        """Should be called any time a Task fails.
        """
        if errorCode == 'LoginError':
            self.sessionExpired = True
            self.connectionLost.emit()
            self.login()
        elif errorCode == 'UnexpectedResponse':
            self.connectionWarning.emit()
        else:
            self.connectionWarning.emit()

    @pyqtSlot(str, str, str)
    def _onUploaded(self, filename, printerName, jobId):
        self.jobUploaded.emit(filename, printerName)
        self.monitorJob(filename, printerName, jobId)

    def pendingJobsCount(self):
        return len(self.runningJobs) + len(self.savedJobs)

    @pyqtSlot(str, str, str)
    def _onMonitorJobFailed(self, filename, printerName, errorCode):
        #if errorCode == 'JobIdNotFound' or errorCode == 'ConnectionError':
        #don't call _onError
        #ignore this error, it usually only happens if you sleep the computer
        # in the middle of monitoring a job. The job probably printed fine.

        filename = str(filename)
        printerName = str(printerName)
        self.deleteJobFile(filename, printerName)
        del self.runningJobs[filename, printerName]
        self.updateToolTip.emit()
        self._onError(errorCode)

    @pyqtSlot(str, str, str)
    def _onUploadFailed(self, filename, printerName, errorCode):
        filename = str(filename)
        printerName = str(printerName)
        errorCode = str(errorCode)
        if errorCode == 'LoginError':
            if Settings().getHasValidCredentials():
                key = (filename, printerName)
                job = self.runningJobs[key]
                del self.runningJobs[key]
                self.loginAndPrint(job)
            else:
                self._onError(errorCode)
                self._onJobFailed(filename, printerName, 
                                    'Not signed in.')
                self.loginRequired.emit()
        elif errorCode == 'InvalidPrinter':
            self._onError(errorCode)
            self._onJobFailed(filename, printerName,
                        'Invalid Printer: No such printer %s' % printerName)
        elif errorCode == 'UnexpectedResponse':
            self._onError(errorCode)
            self._onJobFailed(filename, printerName, '')
        elif errorCode == 'OpenFileError':
            self._onError(errorCode)
            self._onJobFailed(filename, printerName,
                              'Unable to open file \'%s\'' % filename)
        elif errorCode == 'ConnectionError':
            self._onJobFailed(filename, printerName,
                              'Unable to connect to WebPrint. Check network connection.')
        else:
            self._onError(errorCode)
            self._onJobFailed(filename, printerName,
                              'Unknown Error: %s' % errorCode)

    @pyqtSlot(str, str, str)
    def _onJobFinished(self, filename, printerName, statusDict):
        key = (str(filename), str(printerName))
        del self.runningJobs[key]
        self.jobFinished.emit(filename, printerName, statusDict)
        # #update balance
        # stats = json.loads(str(statusDict))
        # cost = int(float(stats['cost']))
        # if self.balance is not None:
        #     self.balance -= cost

    @pyqtSlot(str, str, str)
    def _onJobFailed(self, filename, printerName, errorMessage):
        key = (str(filename), str(printerName))
        del self.runningJobs[key]
        self.jobFailed.emit(filename, printerName, errorMessage)

    # def updateBalance(self):
    #     """Gets the current balance from the server.
    #     """
    #     task = _UpdateBalanceTask(self.sessionId)
    #     task.finished.connect(self._setBalance)
    #     task.failed.connect(self._onError)
    #     QThreadPool.globalInstance().start(task)

    # @pyqtSlot(str)
    # def _setBalance(self, balance):
    #     """Slot to set the balance.
    #     """
    #     self.balance = balance


class _LoginTask(QtCore.QRunnable):
    def __init__(self, username, password, reportErrorsAsEmptyString=False):
        super(_LoginTask, self).__init__()
        self.username = username
        self.password = password
        self._mediator = _LoginMediator()
        self.reportErrorsAsEmptyString = reportErrorsAsEmptyString

    def run(self):
        try:
            sessionId = papercut.login(self.username, self.password)
            self.finished.emit(sessionId)
        except papercut.AuthError:
            if self.reportErrorsAsEmptyString:
                self.failed.emit('')
            else:
                self.failed.emit('AuthError')
        except papercut.UnexpectedResponse:
            if self.reportErrorsAsEmptyString:
                self.failed.emit('')
            else:
                self.failed.emit('UnexpectedResponse')
        except ConnectionError:
            if self.reportErrorsAsEmptyString:
                self.failed.emit('')
            else:
                self.failed.emit('ConnectionError')
        except Exception as e:
            if self.reportErrorsAsEmptyString:
                self.failed.emit('')
            else:
                self.failed.emit(str(e))

    def failed():
        def fget(self):
            return self._mediator.failed
        return locals()
    failed = property(**failed())

    def finished():
        def fget(self):
            return self._mediator.finished
        return locals()
    finished = property(**finished())


class _LoginMediator(QtCore.QObject):
    failed = pyqtSignal(str) #errorCode
    finished = pyqtSignal(str) #sessionId
    def __init__(self):
        super(_LoginMediator, self).__init__()


class _PrintFileTask(QtCore.QRunnable):
    def __init__(self, path, filename, printerName, sessionId, test=False):
        super(_PrintFileTask, self).__init__()
        self.path = path
        self.filename = filename
        self.printerName = printerName
        self.sessionId = sessionId
        self._mediator = _PrintFileMediator()
        self.test = test

    def run(self):
        try:
            jobId = papercut.printFile(self.path, self.printerName,
                               self.sessionId, test=self.test)
        except papercut.LoginError:
            self.failed.emit(self.filename, self.printerName, 'LoginError')
        except papercut.UnexpectedResponse:
            self.failed.emit(self.filename, self.printerName,
                             'UnexpectedResponse')
        except OSError as e:
            self.failed.emit(self.filename, self.printerName, 'OpenFileError')
        except ConnectionError:
            self.failed.emit(self.filename, self.printerName, 'ConnectionError')
        except Exception as e:
            self.failed.emit(self.filename, self.printerName, str(e))
        else:
            self.finished.emit(self.filename, self.printerName, str(jobId))

    def failed():
        def fget(self):
            return self._mediator.failed
        return locals()
    failed = property(**failed())

    def finished():
        def fget(self):
            return self._mediator.finished
        return locals()
    finished = property(**finished())


class _PrintFileMediator(QtCore.QObject):
    failed = pyqtSignal(str, str, str) #filename, printerName, errorCode
    finished = pyqtSignal(str, str, str) #filename, printerName, jobId
    def __init__(self):
        super(_PrintFileMediator, self).__init__()


class _MonitorJobTask(QtCore.QRunnable):
    def __init__(self, filename, printerName, jobId, sessionId,
                 updateInterval=1, test=False):
        super(_MonitorJobTask, self).__init__()
        self.jobId = str(jobId)
        self.sessionId = str(sessionId)
        self.filename = str(filename)
        self.printerName = str(printerName)
        self.updateInterval = updateInterval #in seconds
        self.test = test
        self._mediator = _MonitorJobMediator()

    def run(self):
        try:
            statusDict = papercut.getPrintStatus(self.jobId, self.sessionId,
                                                 test=self.test)
            while(not statusDict['complete']):
                self.updated.emit(self.filename, self.printerName,
                                  json.dumps(statusDict))
                time.sleep(self.updateInterval)
                statusDict = papercut.getPrintStatus(self.jobId, self.sessionId,
                                                     test=self.test)
            self.finished.emit(self.filename, self.printerName,
                               json.dumps(statusDict))
        except papercut.JobIdNotFound:
            self.failed.emit(self.filename, self.printerName, 'JobIdNotFound')
        except ConnectionError: #handled
            self.failed.emit(self.filename, self.printerName, 'ConnectionError')
        except Exception as e:
            self.failed.emit(self.filename, self.printerName, str(e))

    def failed():
        def fget(self):
            return self._mediator.failed
        return locals()
    failed = property(**failed())

    def finished():
        def fget(self):
            return self._mediator.finished
        return locals()
    finished = property(**finished())

    def updated():
        def fget(self):
            return self._mediator.updated
        return locals()
    updated = property(**updated())


class _MonitorJobMediator(QtCore.QObject):
    finished = pyqtSignal(str, str, str) #filename, printerName, statusDict
    updated = pyqtSignal(str, str, str) #filename, printerName, statusDict
    failed = pyqtSignal(str, str, str) #filename, printerName, errorCode
    def __init__(self):
        super(_MonitorJobMediator, self).__init__()


# class _UpdateBalanceTask(QtCore.QRunnable):
#     def __init__(self, sessionId):
#         super(_UpdateBalanceTask, self).__init__()
#         self.sessionId = sessionId
#         self._mediator = _UpdateBalanceMediator()

#     def run(self):
#         try:
#             balance = papercut.getBalance(self.sessionId)
#             self.updateExpires.emit()
#             self.finished.emit(balance)
#         except papercut.LoginError:
#             self.failed.emit('LoginError')
#         except papercut.UnexpectedResponse:
#             self.failed.emit('UnexpectedResponse')
#         except ConnectionError:
#             self.failed.emit('ConnectionError')
#         except Exception as e:
#             self.failed.emit(str(e))

#     def updateExpires():
#         def fget(self):
#             return self._mediator.updateExpires
#         return locals()
#     updateExpires = property(**updateExpires())

#     def failed():
#         def fget(self):
#             return self._mediator.failed
#         return locals()
#     failed = property(**failed())

#     def finished():
#         def fget(self):
#             return self._mediator.finished
#         return locals()
#     finished = property(**finished())


# class _UpdateBalanceMediator(QtCore.QObject):
#     finished = pyqtSignal(str) #balance
#     failed = pyqtSignal(str) #errorCode
#     updateExpires = pyqtSignal()
#     def __init__(self):
#         super(_UpdateBalanceMediator, self).__init__()
