#!/usr/bin/python
# -*- coding: utf-8 -*-
#webprintfolder
# System tray application to print documents at Calvin College using webprint
# Uses QT with the pySide framework

from printer import ThreadedPrinter
from watchfolder import FolderWatcher
import os
import os.path
import json #for sending lists and dicts as thread safe strings in messages.
from settings import Settings
from py2exetools import is_exe, log_without_popup

import sip
sip.setapi('QVariant', 2)

from PyQt4.QtCore import pyqtSlot, QUrl, pyqtSignal
from PyQt4.QtGui import QDesktopServices
from PyQt4 import QtCore, QtGui, Qt
import resources

class TrayIcon(QtGui.QSystemTrayIcon):
    _showSettingsWindow = pyqtSignal()
    def __init__(self, parent=None):
        super(TrayIcon, self).__init__(parent)
        self.setIcon(QtGui.QIcon(':/resources/printer-icon.png'))
        self.setVisible(True)
        self.onMessageClick = lambda: None
        self.messageClicked.connect(self.runOnMessageClick)
        self.activated.connect(self.onIconClicked)
        self.settingsWindow = None
        self.initPrinter()
        self.updateToolTip()
        self.folderWatcher = FolderWatcher(self)
        self.folderWatcher.addJob.connect(self.printer.addJob)
        self.createMenu()
        if is_exe():
            self.hack = None
        else:
            self.hack = QtGui.QWidget() #hack to allow settingsWindow to display. Not sure why it's needed.
            self.hack.show()

    def updateToolTip(self):
        pendingJobs = self.printer.pendingJobsCount()
        if pendingJobs == 1:
            self.setToolTip("1 Document Pending - WebPrint")
        else:
            self.setToolTip("%d Documents Pending - WebPrint" %
                        pendingJobs)

    def showSettingsWindow(self):
        if self.settingsWindow is None:
            self.settingsWindow = SettingsWindow()
            self.settingsWindow.closed.connect(self.onSettingsWindowDestroyed)
            self.settingsWindow.sendLogin.connect(self.printer.login)
            self.settingsWindow.build_webprint_dir.connect(
                                        self.folderWatcher.build_webprint_dir)
            self.settingsWindow.move_webprint_dir.connect(
                                        self.folderWatcher.move_webprint_dir)
            if self.hack is not None:
                sip.delete(self.hack)
                self.hack = None
        self.settingsWindow.focus()

    def showAccountSettings(self):
        self.showSettingsWindow()
        self.settingsWindow.setCurrentToAccount()

    def onSettingsWindowDestroyed(self):
        sip.delete(self.settingsWindow)
        self.settingsWindow = None

    def initPrinter(self):
        settings = Settings()
        self.printer = ThreadedPrinter(url=settings.getWebprintUrl(),
                                       test=False, parent=self) #TODO: set test to False
        self.printer.loginComplete.connect(self.onLogin)
        self.printer.loginFailed.connect(self.onLoginFailed)
        self.printer.jobAdded.connect(self.onJobAdded)
        self.printer.jobUploaded.connect(self.onJobUploaded)
        self.printer.jobFinished.connect(self.onJobFinished)
        self.printer.jobFailed.connect(self.onJobFailed)
        self.printer.jobNeedsLogin.connect(self.onJobNeedsLogin)
        self.printer.updateToolTip.connect(self.updateToolTip)
        self.printer.loginRequired.connect(self.onLoginRequired)
        self.printer.login()

    def onLoginRequired(self):
        self.display('Sign in required', 'Click here to sign in to WebPrint.',
                         onClick=self.showAccountSettings,
                         icon=self.Warning)

    def onJobAdded(self, filename, printer):
        self.display('File Detected', 'Sending \'%s\' to \'%s\'...' % \
                     (filename, printer))
        self.updateToolTip()

    def onJobUploaded(self, filename, printer):
        self.display('Sent to Printer', 'Uploaded \'%s\' to \'%s\'.' % \
                     (filename, printer))

    def onJobFinished(self, filename, printer, status_dict):
        self.display('Finished Printing', 'Document \'%s\' was successfully sent to \'%s\'.' % \
                     (filename, printer))
        try:
            os.remove(os.path.join(os.path.join(
                      str(Settings().getWebprintFolder()),
                      str(printer)), str(filename)))
        except:
            pass
        self.updateToolTip()

    def onJobFailed(self, filename, printer, errorMessage):
        self.display('Printing Failed', 'Document \'%s\' failed to print to \'%s\'.\n%s' % (filename, printer, errorMessage), icon=self.Warning)
        filepath = os.path.join(os.path.join(
                  str(Settings().getWebprintFolder()),
                  str(printer)), str(filename))
        try:
            os.remove(filepath)
        except Exception:
            logging.debug('Unable to remove failed print job file %s' % filepath)
            pass
        self.updateToolTip()

    def onJobNeedsLogin(self, filename, printer):
        self.display('Not Logged In', 'Document \'%s\' is waiting to print.' %\
                                        filename + ' Click here to sign in.',
                                        icon=self.Warning,
                                        onClick=self.showAccountSettings)

    def onIconClicked(self, reason):
        if reason in (QtGui.QSystemTrayIcon.Trigger,
                      QtGui.QSystemTrayIcon.DoubleClick):
            self.contextMenu().popup(QtGui.QCursor.pos())

    def onLogin(self, sessionId):
        """Test function"""
        settings = Settings()
        settings.setHasValidCredentials(True)
        if self.settingsWindow is not None:
            if not self.settingsWindow.accountTab.usernameBox.isEnabled():
                self.display('Signed in', 'Sign in complete.',
                             onClick=self.showAccountSettings)
            self.settingsWindow.accountTab.onLoginSucess()

    def onLoginFailed(self, errorCode):
        if self.settingsWindow is not None:
            self.settingsWindow.accountTab.onLoginFailed()
        if errorCode == 'AuthError':
            settings = Settings()
            settings.setHasValidCredentials(False)
            self.display('Login Error', 'Invalid username or password.',
                         onClick=self.showAccountSettings,
                         icon=self.Warning)
        elif errorCode == 'NoSavedPassword':
            settings = Settings()
            settings.setHasValidCredentials(False)
            self.display('Sign in', 'Click here to sign in to WebPrint.',
                         onClick=self.showAccountSettings,
                         icon=self.Warning)
        elif errorCode == 'UnexpectedResponse':
            self.display('Login Error', 'Unexpected response.',
                         icon=self.Warning)
        elif errorCode == 'ConnectionError':
            self.display('Login Error', 'Unable to connect to WebPrint. Check network connection.')
        else:
            self.display('Login Error', 'Error: %s' % errorCode,
                         icon=self.Warning)

    def display(self, title, message, onClick=lambda: None,
                icon=None, ms=1000):
        if Settings().getShowNotifications():
            if icon is None:
                icon = self.Information
            self.onMessageClick = onClick
            self.showMessage(title, message, icon, ms)

    def runOnMessageClick(self):
        """Runs the appropriate function when a tray message is clicked.
        """
        self.onMessageClick()

    def createMenu(self):
        exitAction = QtGui.QAction("&Exit", self, triggered=QtGui.qApp.quit)
        webprintFolderAction = QtGui.QAction('Open Webprint folder',
                                            self,
                                            triggered=self.openWebprintFolder)
        helpFileAction = QtGui.QAction('Help', self,
                                       triggered=self.openHelp)
        preferencesAction = QtGui.QAction('Preferences...', self,
                                          triggered=self.showSettingsWindow)
        launchWebprintAction = QtGui.QAction('Launch WebPrint website', self,
                                             triggered=self.launchWebsite)
        font = launchWebprintAction.font()
        font.setBold(True)
        launchWebprintAction.setFont(font)

        menu = QtGui.QMenu()
        menu.addAction(launchWebprintAction)
        menu.addAction(webprintFolderAction)
        menu.addSeparator()
        menu.addAction(preferencesAction)
        menu.addAction(helpFileAction) #TODO: Write help file.
        menu.addAction(exitAction)
        self.setContextMenu(menu)

    def launchWebsite(self):
        def launchSite(sessionId):
            if sessionId == '':
                sessionStr = ''
            else:
                sessionStr = ';jsessionid=%s' % sessionId
            url = QUrl('https://webprint.calvin.edu:9192/app%s?service=page/UserSummary'
                       % sessionStr)
            QDesktopServices.openUrl(url)

        self.printer.unusedSessionId.connect(self.onUnusedSessionId)
        self.runOnUnusedSessionId = launchSite
        self.printer.requestUnusedSessionId()

    @pyqtSlot(str)
    def onUnusedSessionId(self, sessionId):
        self.runOnUnusedSessionId(sessionId)
        self.printer.unusedSessionId.disconnect(self.onUnusedSessionId)

    @pyqtSlot()
    def openWebprintFolder(self):
        settings = Settings()
        url = QUrl('file:///' + os.path.abspath(settings.getWebprintFolder()),
                   QUrl.TolerantMode)
        QDesktopServices.openUrl(url)

    @pyqtSlot()
    def openHelp(self):
        url = QUrl('https://github.com/jglamine/WebPrint-Tray-Icon#webprint-tray-icon', QUrl.TolerantMode)
        QDesktopServices.openUrl(url)


class SettingsWindow(QtGui.QDialog):
    closed = pyqtSignal()
    sendLogin = pyqtSignal(str, str)
    sendLogout = pyqtSignal()
    build_webprint_dir = pyqtSignal() #sync webprint dir contents with printers listed in config file
    move_webprint_dir = pyqtSignal(str) #destination
    def __init__(self, parent=None):
        super(SettingsWindow, self).__init__(parent)
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.accountChanges = False
        self.printersChanges = False
        self.generalChanges = False
        self.initUi()

    def initUi(self):
        """Initialize user interface.
        """
        self.changesMade = False
        self.setWindowFlags(self.windowFlags() ^ 
                            QtCore.Qt.WindowContextHelpButtonHint)
        self.tabWidget = QtGui.QTabWidget()
        self.generalTab = GeneralTab()
        self.printersTab = PrintersTab()
        self.accountTab = AccountTab()
        self.tabWidget.addTab(self.generalTab, 'General')
        self.tabWidget.addTab(self.printersTab, 'Printers')
        self.tabWidget.addTab(self.accountTab, 'Account')

        buttonBox = QtGui.QDialogButtonBox(QtGui.QDialogButtonBox.Ok |
                                           QtGui.QDialogButtonBox.Cancel |
                                           QtGui.QDialogButtonBox.Apply)
        buttonBox.accepted.connect(self.accept)
        buttonBox.rejected.connect(self.reject)
        self.applyButton = buttonBox.button(QtGui.QDialogButtonBox.Apply)
        self.applyButton.clicked.connect(self.apply)
        self.applyButton.setDisabled(True)
        self.okButton = buttonBox.button(QtGui.QDialogButtonBox.Ok)
        self.okButton.setDefault(True)
        self.okButton.setAutoDefault(True)
        mainLayout = QtGui.QVBoxLayout()
        mainLayout.addWidget(self.tabWidget)
        mainLayout.addWidget(buttonBox)
        self.setLayout(mainLayout)
        self.setWindowTitle('WebPrint Preferences')
        self.setFixedSize(self.sizeHint())
        self.okButton.setFocus(True)

        self.generalTab.changesMade.connect(self.onGeneralChanges)
        self.printersTab.changesMade.connect(self.onPrintersChanges)
        self.accountTab.changesMade.connect(self.onAccountChanges)
        self.accountTab.sendLogin.connect(self.sendLogin)
        self.accountTab.sendLogout.connect(self.sendLogout)

        self.tabWidget.currentChanged.connect(self.onTabChanged)

    def setCurrentToAccount(self):
        self.tabWidget.setCurrentWidget(self.accountTab)

    def onTabChanged(self, index):
        if (self.tabWidget.currentWidget() == self.accountTab and
        self.accountTab.group.title() == 'Sign in'):
            self.okButton.setFocus(False)
            self.okButton.setAutoDefault(False)
            self.okButton.setDefault(False)
        else:
            self.okButton.setFocus(True)
            self.okButton.setAutoDefault(True)
            self.okButton.setDefault(True)

    def keyPressEvent(self, keyEvent):
        if keyEvent.key() == QtCore.Qt.Key_Return:
            if self.accountTab.group.title() == 'Sign in':
                return
        super(SettingsWindow, self).keyPressEvent(keyEvent)

    def onGeneralChanges(self):
        self.generalChanges = True
        self.onChangesMade()

    def onAccountChanges(self):
        self.accountChanges = True
        self.onChangesMade()

    def onPrintersChanges(self):
        self.printersChanges = True
        self.onChangesMade()

    def onChangesMade(self):
        self.changesMade = True
        self.applyButton.setDisabled(False)

    def accept(self):
        self.apply()
        super(SettingsWindow, self).accept()
        self.closed.emit()

    def reject(self):
        super(SettingsWindow, self).reject
        self.closed.emit()

    def apply(self):
        if self.changesMade:
            settings = Settings(self)
            if self.generalChanges:
                newWebprintFolder = self.generalTab.folderLabel.text()
                oldWebprintFolder = settings.getWebprintFolder()
                if newWebprintFolder != oldWebprintFolder:
                    self.move_webprint_dir.emit(newWebprintFolder)
                oldDesktopNotifications = settings.getShowNotifications()
                newDesktopNotifications = (self.generalTab.
                                           showDesktopCheck.isChecked())
                if oldDesktopNotifications != newDesktopNotifications:
                    settings.setShowNotifications(newDesktopNotifications)
                oldRunOnStartup = settings.getRunOnStartup()
                newRunOnStartup = self.generalTab.autostartCheck.isChecked()
                if oldRunOnStartup != newRunOnStartup:
                    settings.setRunOnStartup(newRunOnStartup)
                self.generalChanges = False
            if self.accountChanges:
                self.accountChanges = False
            if self.printersChanges:
                self.printersChanges = False
                installedPrinters = []
                uninstalledPrinters = []
                grid = self.printersTab.grid
                for row in xrange(1, grid.rowCount()):
                    name = grid.itemAtPosition(row, 0).widget().text()
                    location = grid.itemAtPosition(row, 1).widget().text()
                    isInstalled = grid.itemAtPosition(row, 2).widget().isChecked()
                    if isInstalled:
                        installedPrinters.append((name, location))
                    else:
                        uninstalledPrinters.append((name, location))
                settings.setInstalledPrinters(installedPrinters)
                settings.setUninstalledPrinters(uninstalledPrinters)
                self.build_webprint_dir.emit()
            self.changesMade = False
            self.applyButton.setDisabled(True)

    def focus(self):
        self.raise_()
        self.activateWindow()
        self.showNormal()


class GeneralTab(QtGui.QWidget):
    changesMade = pyqtSignal()
    def __init__(self, parent=None):
        super(GeneralTab, self).__init__(parent)
        self.initUi()

    def initUi(self):
        settings = Settings(self)
        checkGroup = QtGui.QGroupBox()
        self.showDesktopCheck = QtGui.QCheckBox('Show desktop notifications')
        self.showDesktopCheck.setChecked(settings.getShowNotifications())
        self.autostartCheck = QtGui.QCheckBox('Start webprint on system startup')
        self.autostartCheck.setChecked(settings.getRunOnStartup())
        vbox = QtGui.QVBoxLayout()
        vbox.addWidget(self.showDesktopCheck)
        vbox.addWidget(self.autostartCheck)
        vbox.addStretch(1)
        checkGroup.setLayout(vbox)

        folderLocationGroup = QtGui.QGroupBox('Webprint folder location')
        webprintFolder = settings.getWebprintFolder()
        folderInstructionsLabel = QtGui.QLabel('Select a new place to put your webprint folder.\nA folder called "webprint" will be created inside the folder.\n\nWarning: This will delete everything in your webprint folder.')
        self.folderLabel = QtGui.QLineEdit(webprintFolder)
        self.folderLabel.setDisabled(True)
        folderButton = QtGui.QPushButton('Move...')
        hbox = QtGui.QHBoxLayout()
        hbox.addWidget(self.folderLabel)
        hbox.addWidget(folderButton)
        vbox = QtGui.QVBoxLayout()
        vbox.addWidget(folderInstructionsLabel)
        vbox.addLayout(hbox)
        folderLocationGroup.setLayout(vbox)

        mainLayout = QtGui.QVBoxLayout()
        mainLayout.addWidget(checkGroup)
        mainLayout.addWidget(folderLocationGroup)
        mainLayout.addStretch(1)
        self.setLayout(mainLayout)

        folderButton.clicked.connect(self.onFolderButtonClick)
        self.showDesktopCheck.stateChanged.connect(self.changesMade)
        self.autostartCheck.stateChanged.connect(self.changesMade)

    def onFolderButtonClick(self):
        settings = Settings(self)
        webprintFolder = settings.getWebprintFolder()
        labelFolder = str(self.folderLabel.text())
        if labelFolder != webprintFolder:
            browseFolder = os.path.split(labelFolder)[0]
        else:
            browseFolder = webprintFolder
        dirName = str(QtGui.QFileDialog.getExistingDirectory(self,
                                                            'Browse For Folder',
                                                            browseFolder))
        if os.path.exists(dirName):
            if (dirName == webprintFolder
                or dirName == os.path.split(webprintFolder)[0]):
                #user selected the current webprint folder
                self.folderLabel.setText(webprintFolder)
            else:
                #user did not select the webprint folder
                newFolder = os.path.join(dirName, 'webprint')
                self.folderLabel.setText(newFolder)
                self.changesMade.emit()


class PrintersTab(QtGui.QScrollArea):
    changesMade = pyqtSignal()
    def __init__(self, parent=None):
        super(PrintersTab, self).__init__(parent)
        self.initUi()

    def initUi(self):
        def addToGrid(widget1, widget2, widget3, row):
            self.grid.addWidget(widget1, row, 0)
            self.grid.addWidget(widget2, row, 1)
            self.grid.addWidget(widget3, row, 2)
            return row+1
        def addPrinterListToGrid(printers, installed, row):
            for name, location in printers:
                checkBox = QtGui.QCheckBox()
                checkBox.setChecked(installed)
                checkBox.stateChanged.connect(self.changesMade)
                row = addToGrid(QtGui.QLabel(name), QtGui.QLabel(location),
                                checkBox, row)
            return row
        settings = Settings()
        self.grid = QtGui.QGridLayout()
        row = 0
        row = addToGrid(QtGui.QLabel('Printer'), QtGui.QLabel('Location'),
                        QtGui.QLabel('Installed    '), row)
        row = addPrinterListToGrid(sorted(settings.getInstalledPrinters()),
                                   True, row)
        row = addPrinterListToGrid(sorted(settings.getUninstalledPrinters()),
                                   False, row)

        printersWidget = QtGui.QWidget()
        printersWidget.setLayout(self.grid)
        self.setWidget(printersWidget)
        self.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)


class AccountTab(QtGui.QWidget):
    changesMade = pyqtSignal()
    sendLogin = pyqtSignal(str, str) #sent when the user clicks 'Sign in'
    sendLogout = pyqtSignal() #sent when the user clicks 'Sign out'
    def __init__(self, parent=None):
        super(AccountTab, self).__init__(parent)
        self.initUi()

    def initUi(self):
        settings = Settings()
        if settings.getHasValidCredentials():
            #signed in group
            self.group = self.makeSignedInGroup()
        else:
            self.group = self.makeSignedOutGroup()
        self.mainLayout = QtGui.QVBoxLayout()
        self.mainLayout.addWidget(self.group)
        self.setLayout(self.mainLayout)

    def makeSignedInGroup(self):
        settings = Settings()
        group = QtGui.QGroupBox('Signed in')
        caption = QtGui.QLabel('Signed in as %s' % settings.getUsername())
        self.logoutButton = QtGui.QPushButton('Sign out')
        vbox = QtGui.QVBoxLayout()
        vbox.addWidget(caption)
        hbox = QtGui.QHBoxLayout()
        hbox.addWidget(self.logoutButton)
        hbox.addStretch(1)
        vbox.addLayout(hbox)
        vbox.addStretch(1)
        group.setLayout(vbox)

        self.logoutButton.clicked.connect(self.onLogout)

        return group

    def makeSignedOutGroup(self):
        group = QtGui.QGroupBox('Sign in')
        usernameLabel = QtGui.QLabel('Username:')
        passwordLabel = QtGui.QLabel('Password:')
        self.usernameBox = QtGui.QLineEdit()
        self.passwordBox = QtGui.QLineEdit()
        self.passwordBox.setEchoMode(QtGui.QLineEdit.Password)
        self.loginButton = QtGui.QPushButton('Sign in')
        grid = QtGui.QGridLayout()
        grid.addWidget(usernameLabel, 0, 0)
        grid.addWidget(self.usernameBox, 0, 1)
        grid.addWidget(passwordLabel, 1, 0)
        grid.addWidget(self.passwordBox, 1, 1)
        grid.addWidget(self.loginButton, 2, 1)
        grid.setRowStretch(3, 1)
        grid.setColumnStretch(2, 1)
        group.setLayout(grid)

        self.loginButton.clicked.connect(self.onLoginSubmit)
        self.usernameBox.returnPressed.connect(self.onLoginSubmit)
        self.passwordBox.returnPressed.connect(self.onLoginSubmit)

        return group

    def onLogout(self):
        self.logoutButton.clicked.disconnect()
        self.mainLayout.removeWidget(self.group)
        sip.delete(self.group)
        self.group = self.makeSignedOutGroup()
        self.mainLayout.addWidget(self.group)
        settings = Settings()
        settings.setHasValidCredentials(False)
        settings.deletePassword(settings.getUsername())
        settings.setUsername('')
        self.sendLogout.emit()

    def onLoginSubmit(self):
        self.usernameBox.setDisabled(True)
        self.passwordBox.setDisabled(True)
        self.loginButton.setDisabled(True)
        self.sendLogin.emit(unicode(self.usernameBox.text(), 'utf-8'),
                            unicode(self.passwordBox.text(), 'utf-8'))

    def onLoginFailed(self):
        if not self.loginButton.isEnabled():
            self.usernameBox.setEnabled(True)
            self.passwordBox.setEnabled(True)
            self.loginButton.setEnabled(True)

    def onLoginSucess(self):
        if not self.loginButton.isEnabled():
            settings = Settings()
            settings.setUsername(self.usernameBox.text())
            settings.setPassword(self.passwordBox.text())
            settings.setHasValidCredentials(True)
            self.mainLayout.removeWidget(self.group)
            sip.delete(self.group)
            self.group = self.makeSignedInGroup()
            self.mainLayout.addWidget(self.group)


if __name__ == "__main__":
    import sys

    os.environ['REQUESTS_CA_BUNDLE'] = '..\\cacert.pem'
    if is_exe():
        log_without_popup()

    app = QtGui.QApplication(sys.argv)

    if not QtGui.QSystemTrayIcon.isSystemTrayAvailable():
        QtGui.QMessageBox.critical(None, "Systray",
                "I couldn't detect any system tray on this system.")
        sys.exit(1)

    QtGui.QApplication.setQuitOnLastWindowClosed(False)
    app.setWindowIcon(QtGui.QIcon(':/resources/printer-icon.png'))

    trayIcon = TrayIcon()
    trayIcon.show()
    sys.exit(app.exec_())