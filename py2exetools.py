#!/usr/bin/python
# -*- coding: utf-8 -*-
#py2exetools

import sys

def is_exe():
	return hasattr(sys, "frozen")

def log_without_popup():
    sys.stderr = _ErrorLog()
    #del _ErrorLog

class _ErrorLog(object):
    """Handled output from a stream (stderr or stdin)
    Logs output to app.exe.log, where app.exe is the filename of the
    executable being run.

    Modified from py2exe/boot_common.py to not display a
    popup window on exit.
    If the log file cannot be opened, nothing is logged.
    """
    softspace = 0
    _file = None
    _error = None
    def write(self, text, fname=sys.executable + '.log'):
        if self._file is None and self._error is None:
            try:
                self._file = open(fname, 'a')
            except Exception:
                #unable to open log file
                pass
        if self._file is not None:
            self._file.write(text)
            self._file.flush()
    def flush(self):
        if self._file is not None:
            self._file.flush()