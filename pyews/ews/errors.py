##
## Created : Thu May 01 10:51:44 IST 2014
##
## Copyright (C) 2014 Sriram Karra <karra.etc@gmail.com>
##
## This file is part of pyews
##
## pyews is free software: you can redistribute it and/or modify it under
## the terms of the GNU Affero General Public License as published by the
## Free Software Foundation, version 3 of the License
##
## pyews is distributed in the hope that it will be useful, but WITHOUT
## ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
## FITNESS FOR A PARTICULAR PURPOSE.  See the GNU Affero General Public
##s License for more details.
##
## You should have a copy of the license in the doc/ directory of pyews.  If
## not, see <http://www.gnu.org/licenses/>.

from   pyews.soap import SoapClient, SoapMessageError, QName_M

class EWSMessageError(SoapMessageError):
    def __init__ (self, resp_code, xml_resp=None, node=None):
        SoapMessageError.__init__(self, resp_code, xml_resp, node)
        self.parse_error_msg()

    def parse_error_msg (self):
        if self.node is not None:
            self.node = SoapClient.parse_xml(self.xml_resp)

        if self.resp_code == 'ErrorSchemaValidation':
            for i in self.node.iter('faultstring'):
                if i is not None:
                    self.msg_text = i.text
        else:
            for i in self.node.iter(QName_M('MessageText')):
                if i is not None:
                    self.msg_text = i.text

        return None

    def __str__ (self):
        return '%s - %s' % (self.resp_code, self.msg_text)

class EWSCreateFolderError(Exception):
    pass

class EWSDeleteFolderError(Exception):
    pass
