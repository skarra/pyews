##
## Created : Sun Mar 16 21:07:35 IST 2014
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
## License for more details.
##
## You should have a copy of the license in the doc/ directory of pyews.  If
## not, see <http://www.gnu.org/licenses/>.

import logging, re, requests
from requests.auth import HTTPBasicAuth
import xml.etree.ElementTree as ET
import utils

M_NAMESPACE = 'http://schemas.microsoft.com/exchange/services/2006/messages'
T_NAMESPACE = 'http://schemas.microsoft.com/exchange/services/2006/types'

def unQName (name):
    return re.match('{.*)}(.*)', name).group(1)

def QName (namespace, name):
    return '{%s}%s' % (namespace, name)

def QName_M (name):
    return QName(M_NAMESPACE, name)

def QName_T (name):
    return QName(T_NAMESPACE, name)

class SoapClient(object):
    def __init__ (self, service_url, user, pwd):
        self.url = service_url
        self.user = user
        self.pwd = pwd
        self.auth = HTTPBasicAuth(user, pwd)

    def send (self, request):
        r = requests.post(self.url, auth=self.auth, data=request,
                          headers={'Content-Type':'text/xml; charset=utf-8',
                                   "Accept": "text/xml"})

        return r.text

    @staticmethod
    def parse_xml (soap_resp):
        return ET.fromstring(soap_resp)

    @staticmethod
    def get_response_code (soap_resp, root=None):
        """Return a (resp_code, root_node) tuple after parsing the soap
        response xml message. If root node is not present then the provided
        soap_response xml message is first parsed"""

        if not root:
            print utils.pretty_xml(soap_resp)
            root = SoapClient.parse_xml(soap_resp)

        for i in root.iter(QName_M('ResponseCode')):
            return (i.text, root) if i is not None else (None, root)

    @staticmethod
    def get_error_msg (soap_resp, root=None):
        if not root:
            root = SoapClient.parse_xml(soap_resp)

        for i in root.iter(QName_M('MessageText')):
            return (i.text, root) if i is not None else (None, root)
