#!/usr/bin/python
# -*- coding: utf-8 -*-
# Low level, stateless functions for communicating with a PaperCut front end.

import requests
import re
import json
import time
import random #for testing purposes
from bs4 import BeautifulSoup

URL = "https://webprint.calvin.edu:9192"
VALID_FILETYPES = ['xlam', 'xls', 'xlsb', 'xlsm', 'xlsx', 'xltm', 'xltx',
'pot', 'potm', 'potx', 'ppam', 'pps', 'ppsm', 'ppsx', 'ppt', 'pptm', 'pptx',
'doc', 'docm', 'docx', 'dot', 'dotm', 'dotx', 'xps', 'pdf']
JSON_TEST_DATA = [
{u'status': {u'text': u'Submitting', u'code': u'new', u'complete': False, u'formatted': u'Submitting'}, u'documentName': u'printme.doc', u'printer': u'dash\\KH_PUB'},
{u'status': {u'text': u'Rendering', u'code': u'rendering', u'messages': [{u'info': u'Queued in position 2.'}], u'complete': False, u'formatted': u'<span class="info">Queued in position 1.</span>'}, u'documentName': u'printme.doc', u'printer': u'dash\\KH_PUB'},
{u'status': {u'text': u'Finished: Queued for printing', u'code': u'queued', u'complete': True, u'formatted': u'Finished: Queued for printing'}, u'documentName': u'printme.doc', u'printer': u'dash\\KH_PUB', u'cost': u'$0.05', u'pages': 1}
]

def login(username, password, url=URL):
    """Returns a sessionId or None if login failed.
    """
    payload = {}
    payload['service'] = 'direct/0/Home/$Form'
    payload['sp'] = 'S0'
    payload['Form0'] = \
        '$Hidden,inputUsername,inputPassword,$PropertySelection,$Submit'
    payload['$Hidden'] = 'true'
    payload['$PropertySelection'] = 'en'
    payload['inputUsername'] = username
    payload['inputPassword'] = password
    response = requests.post(url + '/app', data=payload)
    # check for errors
    response.raise_for_status()
    soup = BeautifulSoup(response.text)
    title = soup.find('title')
    if title == None:
        raise UnexpectedResponse
    title = title.string
    if title == 'Login':
        raise AuthError('Invalid username or password')
    if title != 'PaperCut NG : Summary':
        raise UnexpectedResponse
    sessionId = response.cookies.get('JSESSIONID')
    if sessionId is None:
        raise UnexpectedResponse('No session cookie returned')
    return sessionId

def printFile(filename, printerName, sessionId, numCopies=1, url=URL,
            test=False):
    """Prints a file to the given printer.

    Returns the print job id, which can be used to retrieve status
    about the job with getPrintStatus()

    Throws OSError if the file cannot be opened.

    Args:
        PrinterName: Either the short or long name of a printer. (str)
    """
    with open(filename, 'rb') as file_handel:
        cookies = {'JSESSIONID': sessionId}
        if filename.rsplit('.', 1)[1] not in VALID_FILETYPES:
            raise FiletypeError('Invalid extension in file: %s' % filename)
        jobId = 0
        #setup - load form page:
        r = requests.get('https://webprint.calvin.edu:9192/app?service=action/1/UserWebPrint/0/$ActionLink',
                         cookies=cookies)
        soup = BeautifulSoup(r.text)
        title = soup.find('title')
        if title == None:
            raise UnexpectedResponse
        title = title.string
        if title == 'Login':
            raise LoginError('Not logged in')
        if title != 'PaperCut NG : Web Print - Step 1 - Printer Selection':
            raise UnexpectedResponse
        #find printer id
        printerId = _findPrinterId(soup, printerName)
        if printerId is None:
            raise InvalidPrinter('Printer %s not found' % printerName)
        #submit printer selection
        payload = {
        'service': 'direct/1/UserWebPrintSelectPrinter/$Form',
        'sp': 'S0',
        'Form0': '$Hidden,$Hidden$0,$TextField,$Submit,$RadioGroup,$Submit$0,$Submit$1',
        '$Hidden': '', '$Hidden$0': '', '$TextField': '',
        '$RadioGroup': str(printerId),
        '$Submit$1': '2. Print Options and Account Selection »'
        }
        response = requests.post(url + '/app', data=payload, cookies=cookies)
        response.raise_for_status()
        #submit number of copies (max = 999)
        payload = {
        'service': 'direct/1/UserWebPrintOptionsAndAccountSelection/$Form',
        'sp': 'S0', 'Form0': 'copies,$Submit,$Submit$0',
        'copies': str(numCopies),
        '$Submit': '3. Upload Document »'
        }
        response = requests.post(url + '/app', data=payload, cookies=cookies)
        response.raise_for_status()
        #upload file
        #get file upload url
        upload_path_match = re.search(
                    r'var uploadFormSubmitURL = \'(.*)\'', response.text)
        if not upload_path_match:
            raise UnexpectedResponse
        upload_path = upload_path_match.group(1)
        #upload file
        files = {'file': file_handel}
        if not test:
            response = requests.post(url + upload_path, files=files,
                                     cookies=cookies)
            response.raise_for_status()
            #upload other data
            payload = { 'service': 'direct/1/UserWebPrintUpload/$Form$0',
                        'sp': 'S1', 'Form1': ''}
            response = requests.post(url + '/app', data=payload, cookies=cookies)
            response.raise_for_status()
            match = re.search(r'var webPrintUID = \'(.*)\';', response.text)
            if match is None:
                raise UnexpectedResponse
            jobId = match.group(1)
    return jobId

def getPrintStatus(jobId, sessionId='', url=URL, test=False):
    """Returns a dictionary with status information about a print job.
    """
    cookies = {'JSESSIONID': sessionId}
    if test == False:
        response = requests.get('%s/rpc/web-print/job-status/%s.json'
                    % (url, str(jobId)), cookies=cookies)
        if response.status_code == '404':
            raise JobIdNotFound
        response.raise_for_status()
        json_data = json.loads(response.text)
    else:
        x = random.randrange(100)
        if x <= 65:
            json_data = JSON_TEST_DATA[random.randrange(0, len(JSON_TEST_DATA)-1)]
        else:
            json_data = JSON_TEST_DATA[-1]
    statusDict = {}
    statusDict['complete'] = json_data['status']['complete']
    statusDict['filename'] = json_data['documentName']
    statusDict['printer'] = json_data['printer']
    if json_data['status']['text'] == 'Submitting':
        statusDict['status'] = 'Submitting'
    elif json_data['status']['text'] == 'Rendering':
        statusDict['status'] = 'Rendering: %s' % (
                                json_data['status']['messages'][-1]['info'])
    else:
        statusDict['status'] = json_data['status']['text']
    if statusDict['complete']:
        statusDict['pages'] = int(json_data['pages'])
        statusDict['cost'] = json_data['cost'][1:]
    return statusDict

def listPrinters(sessionId, url=URL):
    """Returns a list of dictionaries of information about available printers.
    """
    printers = []
    cookies = {'JSESSIONID': sessionId}
    r = requests.get('https://webprint.calvin.edu:9192/app?service=action/1/UserWebPrint/0/$ActionLink',
                     cookies=cookies)
    r.raise_for_status()
    soup = BeautifulSoup(r.text)
    title = soup.find('title')
    if title == None:
        raise UnexpectedResponse
    title = title.string
    if title == 'Login':
        raise LoginError('Not logged in')
    if title != 'PaperCut NG : Web Print - Step 1 - Printer Selection':
        raise UnexpectedResponse
    input_tags = soup.findAll('input', 
                                attrs={'type':'radio', 'name':'$RadioGroup'})
    if input_tags is None:
        raise UnexpectedResponse
    for input_tag in input_tags:
        shortName = input_tag.parent.text.strip()
        name = input_tag.parent.parent.find_next_sibling('td',
                                class_='locationColumnValue').text.strip()
        printers.append((shortName, name))
    if printers == []:
        raise UnexpectedResponse
    return printers

def getBalance(sessionId, url=URL):
    """Returns the user's balance in cents. (int)
    """
    cookies = {'JSESSIONID': sessionId}
    response = requests.get(url + '/app?service=page/UserSummary',
                            cookies=cookies)
    response.raise_for_status()
    soup = BeautifulSoup(response.text)
    depth = soup.find('th', class_='desc', text='Balance')
    if depth is None:
        title = soup.find('title')
        if title is not None and title.text == 'Login':
            raise LoginError('Not logged in')
        raise UnexpectedResponse
    try:
        money_str = depth.find_next_sibling('td', class_='fields').text
        money_str = money_str.strip()[1:]
    except:
        raise UnexpectedResponse
    if money_str is None:
        raise UnexpectedResponse
    return int(float(money_str) * 100)

def _findPrinterId(soup, printerName):
    """Return the id of a printer (int), or None if not found.

    args:
        soup: Parsed /app?service=action/1/UserWebPrint/0/$ActionLink page.
            (BeautifulSoup)
        printerName: Either the short or long name of a printer. (str)
    raises:
        UnexpectedResponse: if data is not in the expected format.
    """
    input_tags = soup.findAll('input', 
                                attrs={'type':'radio', 'name':'$RadioGroup'})
    if input_tags is None:
        raise UnexpectedResponse
    for input_tag in input_tags:
        shortName = input_tag.parent.text.strip()
        name = input_tag.parent.parent.find_next_sibling('td',
                                class_='locationColumnValue').text.strip()
        if printerName == name or printerName == shortName:
            return int(input_tag['value'])
    return None


class LoginError(Exception):
    pass

class AuthError(Exception):
    pass

class UnexpectedResponse(Exception):
    pass

class FiletypeError(Exception):
    pass

class InvalidPrinter(Exception):
    pass

class JobIdNotFound(Exception):
    pass
