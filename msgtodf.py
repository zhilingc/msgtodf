# -*- coding: utf-8 -*-
'''
Extracts .msg files to a pandas dataframe.

Modified from mattgwwalker's msg-extractor: https://github.com/mattgwwalker/msg-extractor
for use outside of commandline and to output to a dataframe.
'''


#%%
import os
import sys
from email.parser import Parser as EmailParser
import email.utils
import copy
import olefile as OleFile
import pandas as pd

class Message(OleFile.OleFileIO):
    def __init__(self,filename):
        OleFile.OleFileIO.__init__(self,filename)
        self.filename=filename
        
    def windowsUnicode(self,string):
        if string is None:
            return None
        if sys.version_info[0] >= 3:  # Python 3
            return str(string, 'utf_16_le')
        else:  # Python 2
            return unicode(string, 'utf_16_le')
    
    def _getStream(self, filename):
        if self.exists(filename):
            stream = self.openstream(filename)
            return stream.read()
        else:
            return None

    def _getStringStream(self, filename, prefer='unicode'):
        """Gets a string representation of the requested filename.
        Checks for both ASCII and Unicode representations and returns
        a value if possible.  If there are both ASCII and Unicode
        versions, then the parameter /prefer/ specifies which will be
        returned.
        """

        if isinstance(filename, list):
            # Join with slashes to make it easier to append the type
            filename = "/".join(filename)

        asciiVersion = self._getStream(filename + '001E')     
        unicodeVersion = self.windowsUnicode(self._getStream(filename + '001F'))
        if asciiVersion is None:
            return unicodeVersion
        elif unicodeVersion is None:
            return asciiVersion
        else:
            if prefer == 'unicode':
                return unicodeVersion
            else:
                return asciiVersion

    @property
    def subject(self):
        return self._getStringStream('__substg1.0_0037')

    @property
    def header(self):
        try:
            return self._header
        except Exception:
            headerText = self._getStringStream('__substg1.0_007D')
            if headerText is not None:
                self._header = EmailParser().parsestr(headerText)
            else:
                self._header = None
            return self._header

    @property
    def date(self):
        # Get the message's header and extract the date
        if self.header is None:
            return None
        else:
            return self.header['date']

    @property
    def parsedDate(self):
        return email.utils.parsedate(self.date)

    @property
    def sender(self):
        try:
            return self._sender
        except Exception:
            # Check header first
            if self.header is not None:
                headerResult = self.header["from"]
                if headerResult is not None:
                    self._sender = headerResult
                    return headerResult

            # Extract from other fields
            text = self._getStringStream('__substg1.0_0C1A')
            email = self._getStringStream('__substg1.0_0C1F')
            result = None
            if text is None:
                result = email
            else:
                result = text
                if email is not None:
                    result = result + " <" + email + ">"

            self._sender = result
            return result

    @property
    def to(self):
        try:
            return self._to
        except Exception:
            # Check header first
            if self.header is not None:
                headerResult = self.header["to"]
                if headerResult is not None:
                    self._to = headerResult
                    return headerResult

            # Extract from other fields
            # TODO: This should really extract data from the recip folders,
            # but how do you know which is to/cc/bcc?
            display = self._getStringStream('__substg1.0_0E04')
            self._to = display
            return display

    @property
    def cc(self):
        try:
            return self._cc
        except Exception:
            # Check header first
            if self.header is not None:
                headerResult = self.header["cc"]
                if headerResult is not None:
                    self._cc = headerResult
                    return headerResult

            # Extract from other fields
            # TODO: This should really extract data from the recip folders,
            # but how do you know which is to/cc/bcc?
            display = self._getStringStream('__substg1.0_0E03')
            self._cc = display
            return display

    @property
    def body(self):
        # Get the message body
        return self._getStringStream('__substg1.0_1000')

    @property
    def attachments(self):
        try:
            return self._attachments
        except Exception:
            # Get the attachments
            attachmentDirs = []

            for dir_ in self.listdir():
                if dir_[0].startswith('__attach') and dir_[0] not in attachmentDirs:
                    attachmentDirs.append(dir_[0])

            self._attachments = []

            for attachmentDir in attachmentDirs:
                self._attachments.append(Attachment(self, attachmentDir))

            return self._attachments
    

            
class Extract:
    def __init__(self,path,isdir=False,verbose=True):
        if isdir:
            self.files = [os.path.join(path,x) for x in os.listdir(path) if ".msg" in x]
        else: self.files = [path]
        self.result = None
           
    def xstr(self,string):
        return '' if string is None else str(string).encode("UTF-8")
    
    def makedf(self,verbose=True):
        print "Total number of files: %d" %len(self.files)
        for msg in self.files:
            if verbose: 
                sys.stdout.write('\rCurrent file: '+msg)
                sys.stdout.flush()
                
            mymsg = Message(msg)
            msgdf = pd.DataFrame(data = {'from': self.xstr(mymsg.sender),
                                        'to': self.xstr(mymsg.to),
                                        'cc': self.xstr(mymsg.cc),
                                        'subject': self.xstr(mymsg.subject),
                                        'date': self.xstr(mymsg.date),
                                        'body': self.xstr(mymsg.body)},
                                index = xrange(1))
            msgdf = msgdf[['from','to','cc','subject','date','body']]
            if self.result is None:
                self.result = copy.deepcopy(msgdf)
            else:
                self.result = pd.concat([self.result,msgdf])
        return self.result
                                        
                                        
