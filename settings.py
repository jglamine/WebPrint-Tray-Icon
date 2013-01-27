#!/usr/bin/python
# -*- coding: utf-8 -*-
import sip
sip.setapi('QVariant', 2)

import keyring
import os
import sys
from PyQt4 import QtCore
from PyQt4.QtCore import QSettings
from win32com.client import Dispatch #used to save shortcut in windows

_APP_NAME = 'webprintFolder'
_ORGANIZATION_NAME = 'James Lamine'

#Note: getters and setters are not pythonic. Possibly do this a better way.

class Settings(QtCore.QObject):
    def __init__(self, parent=None):
        super(Settings, self).__init__(parent)
        self.config = self._getQSettings()

    def getHasValidCredentials(self):
        return self.config.value('hasValidCredentials', defaultValue=False,
                                 type=bool)

    def setHasValidCredentials(self, hasValidCredentials):
        self.config.setValue('hasValidCredentials', hasValidCredentials)

    def getWebprintFolder(self):
        if not self.config.contains('webprintFolder'):
            # Platform Dependant: Windows
            webprintFolder = os.path.expanduser('~\Documents\webprint')
            if not os.path.exists(webprintFolder):
                os.mkdir(webprintFolder)
            self.config.setValue('webprintFolder', webprintFolder)
        else:
            webprintFolder = self.config.value('webprintFolder')
        return str(webprintFolder)

    def getUsername(self):
        return str(self.config.value('username', defaultValue=''))

    def getInstalledPrinters(self):
        """Returns a list of printers in format (Name, Location).
        """
        installedPrinters = []
        size = self.config.beginReadArray('installedPrinters')
        for i in xrange(size):
            self.config.setArrayIndex(i)
            name = self.config.value('name')
            location = self.config.value('location')
            installedPrinters.append((str(name), str(location)))
        self.config.endArray()
        return installedPrinters

    def getUninstalledPrinters(self):
        uninstalledPrinters = []
        size = self.config.beginReadArray('uninstalledPrinters')
        for i in xrange(size):
            self.config.setArrayIndex(i)
            name = self.config.value('name')
            location = self.config.value('location')
            uninstalledPrinters.append((str(name), str(location)))
        self.config.endArray()
        if uninstalledPrinters == []:
            uninstalledPrinters = [(u'dash\\BB_PUB', u'Computer Lab (Boer Bennink)'), (u'dash\\BHT_PUB', u'Computer Lab (Bolt Heynes Timmer)'), (u'dash\\BV_PUB', u'Computer Lab (Beets Veenstra)'), (u'dash\\CF229_PUB', u'Covenant Fine Arts Center (CF229)'), (u'dash\\CFAC_WEBPRINT_PUB', u'CFAC'), (u'dash\\DC110_PUB', u'Computer Lab (DC 110)'), (u'dash\\DRC1_PUB', u'Library (HL 206)'), (u'dash\\GAMMA_PUB', u'Computer Lab (Gamma)'), (u'dash\\ITC_HL102_COLORA_PUB', u'Information Technology Center (HL 102)'), (u'dash\\ITC_HL102A_PUB', u'Information Technology Center (HL 102)'), (u'dash\\KH_PUB', u'Computer Lab (Kalsbeek Huizenga)'), (u'dash\\LIBRARY_4TH_FLOOR', u'HL 4th Floor'), (u'dash\\LIBRARY_5TH_FLOOR', u'Library 5th Floor'), (u'dash\\NV_PUB', u'Computer Lab (Noordewier VanderWerp)'), (u'dash\\PC_PUB', u'Computer Lab (Phi Chi)'), (u'dash\\RV_PUB', u'Computer Lab (Rooks Vandellen)'), (u'dash\\SB157_WEBPRINT_PUB', u'SB 157'), (u'dash\\SB176_PUB', u'Lab (SB 176)'), (u'dash\\SB177_PUB', u'Lab (SB 177)'), (u'dash\\SE_PUB', u'Computer Lab (Schultze Eldersveld)'), (u'dash\\SFC_WEBPRINT', u'Spoelhof Fieldhouse'), (u'dash\\TE_PUB', u'Computer Lab (Theta Epsilon)'), (u'dash\\ZL_PUB', u'Computer Lab (Zeta Lamda)')]
        return uninstalledPrinters

    def getShowNotifications(self):
        return self.config.value('showNotifications', defaultValue=True,
                                 type=bool)

    def getRunOnStartup(self):
        return self.config.value('runOnStartup', defaultValue=True, type=bool)

    def getWebprintUrl(self):
        return str(self.config.value('webprintUrl',
                            defaultValue='https://webprint.calvin.edu:9192'))

    def getPassword(self):
        """Returns a stored password from the system password manager.
        Returns None if the password does not exist.
        """
        return keyring.get_password(_APP_NAME, self.getUsername())

    def setUsername(self, username):
        self.config.setValue('username', username)

    def setInstalledPrinters(self, installedPrinters):
        if installedPrinters == []:
            self.config.remove('installedPrinters')
            return
        self.config.beginWriteArray('installedPrinters')
        for i in xrange(len(installedPrinters)):
            name, location = installedPrinters[i]
            self.config.setArrayIndex(i)
            self.config.setValue('name', name)
            self.config.setValue('location', location)
        self.config.endArray() 

    def setUninstalledPrinters(self, uninstalledPrinters):
        if uninstalledPrinters == []:
            self.config.remove('uninstalledPrinters')
            return
        self.config.beginWriteArray('uninstalledPrinters')
        for i in xrange(len(uninstalledPrinters)):
            name, location = uninstalledPrinters[i]
            self.config.setArrayIndex(i)
            self.config.setValue('name', name)
            self.config.setValue('location', location)
        self.config.endArray()

    def setShowNotifications(self, showNotifications):
        self.config.setValue('showNotifications', showNotifications)

    def setRunOnStartup(self, runOnStartup):
        self.config.setValue('runOnStartup', runOnStartup)
        if not self._startupLinkExists() == runOnStartup:
            path = os.path.expanduser(
                    '~/Start Menu/Programs/Startup/Web Print Tray Icon.lnk')
            if runOnStartup:
                #Platform Dependant: Windows
                targetPath = os.path.abspath(sys.argv[0])
                workingDir = os.path.dirname(targetPath)
                self._createShortcut(path, target=targetPath, wDir=workingDir)
            else:
                if os.path.exists(path):
                    os.remove(path)

    def setWebprintUrl(self, webPrintUrl):
        self.config.setValue('webprintUrl', webprintUrl)

    def setWebprintFolder(self, webprintFolder):
        self.config.setValue('webprintFolder', webprintFolder)

    def setPassword(self, password):
        """Securely save password using the system password manager.

        Raises ValueError if the username is blank.
        """
        username = self.getUsername()
        if username == '':
            raise ValueError('Cannot save password for blank username.')
        password = unicode(password, 'utf-8')
        username = unicode(username, 'utf-8')
        return keyring.set_password(_APP_NAME, username, password)

    def deletePassword(self, username):
        return keyring.delete_password(_APP_NAME, username)

    def _startupLinkExists(self):
        """Check whether or not a link exists in the Startup folder.
        """
        #Platform Dependant: Windows
        return os.path.isfile(os.path.expanduser(
                            '~/Start Menu/Programs/Startup/Web Print Tray Icon.lnk'))

    def _createShortcut(self, path, target='', wDir='', icon=''):
        """Creates a shortcut in windows.

        Platform Dependant: Windows
        Code From: http://www.blog.pythonlibrary.org/2010/01/23/using-python-to-create-shortcuts/
        """
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(path)
        shortcut.Targetpath = target
        shortcut.WorkingDirectory = wDir
        if icon == '':
            pass
        else:
            shortcut.IconLocation = icon
        shortcut.save()

    def _getQSettings(self):
        return QSettings(QSettings.IniFormat, QSettings.UserScope,
                        _ORGANIZATION_NAME, _APP_NAME, parent=self)