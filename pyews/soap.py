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
from   requests.auth import HTTPBasicAuth
import xml.etree.ElementTree as ET
import utils
from   utils import pretty_xml

E_NAMESPACE = 'http://schemas.microsoft.com/exchange/services/2006/errors'
M_NAMESPACE = 'http://schemas.microsoft.com/exchange/services/2006/messages'
S_NAMESPACE = 'http://schemas.xmlsoap.org/soap/envelope/'
T_NAMESPACE = 'http://schemas.microsoft.com/exchange/services/2006/types'

def unQName (name):
    res = re.match('{.*}(.*)', name)
    return name if res is None else res.group(1)

def QName (namespace, name):
    return '{%s}%s' % (namespace, name)

def QName_E (name):
    return QName(E_NAMESPACE, name)

def QName_M (name):
    return QName(M_NAMESPACE, name)

def QName_S (name):
    return QName(S_NAMESPACE, name)

def QName_T (name):
    return QName(T_NAMESPACE, name)

class SoapMessageError(Exception):
    def __init__ (self, code, xml_resp=None, node=None):
        """xml_resp is the raw xml test
        node is a reference to the parsed xml root element.
        code is the response code"""

        self.xml_resp = xml_resp
        self.node = node
        self.resp_code = code

class SoapConnectionError(Exception):
    pass

class SoapClient(object):
    def __init__ (self, service_url, user, pwd):
        self.url = service_url
        self.user = user
        self.pwd = pwd
        self.auth = HTTPBasicAuth(user, pwd)

    def send (self, request, debug=False):
        """
        Send the given rquest to the server, and return the response text as
        well as a parsed node object as a (resp.text, node) tuple.

        The response text is the raw xml including the soap headers and stuff.
        """

        try:
            r = requests.post(self.url, auth=self.auth, data=request,
                              headers={'Content-Type':'text/xml; charset=utf-8',
                                       "Accept": "text/xml"})
        except requests.exceptions.ConnectionError as e:
            raise SoapConnectionError(e)

        if debug:
            logging.debug('%s', pretty_xml(r.text))

        return SoapClient.parse_xml(r.text)

    @staticmethod
    def parse_xml (soap_resp):
        return ET.fromstring(soap_resp)

    @staticmethod
    def get_node_attribute (root, node, att):
        """
        From given soap response xml find and return the attributes of the
        specified node (first occurrence)
        """

        for i in root.iter(QName_M(node)):
            return i.attrib[att] if i is not None else None

        return None

    @staticmethod
    def find_first_child (root, tag, ret='text'):
        """
        Look for the first child of root with specified tag and return it. If
        ret is 'text' then the value of the node is returned, else, the node
        is returned as an element.
        """

        for child in root.iter(tag):
            if ret == 'text':
                return child.text
            else:
                return child

        return []

    @staticmethod
    def get_node_detail (soap_resp, root, node):
        """
        From given soap response xml find the first occurrence of node and
        return a tuple (node.text, node.attr, root)

        node should be a string. root should be an Element object if not
        None. In the returned tuple node.text willbe a string, node.attrib
        will be a dictionary.
        """

        if root is not None:
            root = SoapClient.parse_xml(soap_resp)

        if root is None:
            return (None, None, None)

        for i in root.iter(node):
            if i is not None:
                return (i.text, i.attrib, root)

        return (None, None, root)
