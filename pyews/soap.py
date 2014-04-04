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
E_NAMESPACE = 'http://schemas.microsoft.com/exchange/services/2006/errors'
T_NAMESPACE = 'http://schemas.microsoft.com/exchange/services/2006/types'

def unQName (name):
    res = re.match('{.*}(.*)', name)
    return name if res is None else res.group(1)

def QName (namespace, name):
    return '{%s}%s' % (namespace, name)

def QName_M (name):
    return QName(M_NAMESPACE, name)

def QName_E (name):
    return QName(E_NAMESPACE, name)

def QName_T (name):
    return QName(T_NAMESPACE, name)

class SoapMessageError(Exception):
    pass

class SoapConnectionError(Exception):
    pass

class SoapClient(object):
    def __init__ (self, service_url, user, pwd):
        self.url = service_url
        self.user = user
        self.pwd = pwd
        self.auth = HTTPBasicAuth(user, pwd)

    def send (self, request):
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

        r_code, node = SoapClient.get_response_code(r.text)
        if r_code != 'NoError':
            r_msg, node = SoapClient.get_error_msg(r.text, node, code=r_code)
            raise SoapMessageError(r_msg)

        return r.text, node

    @staticmethod
    def parse_xml (soap_resp):
        return ET.fromstring(soap_resp)

    @staticmethod
    def get_node_attribute (soap_resp, root, node, att):
        """
        From given soap response xml find and return the attributes of the
        specified node (first occurrence)
        """

        if root is not None:
            root = SoapClient.parse_xml(soap_resp)

        for i in root.iter(QName_M(node)):
            return (i.attrib[att], root) if i is not None else (None, root)

        return (None, root)

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

    @staticmethod
    def get_response_code (soap_resp, root=None):
        """Return a (resp_code, root_node) tuple after parsing the soap
        response xml message. If root node is not present then the provided
        soap_response xml message is first parsed"""

        if not root:
            root = SoapClient.parse_xml(soap_resp)

        for i in root.iter(QName_M('ResponseCode')):
            if i is not None:
                return (i.text, root)

        for i in root.iter('faultcode'):
            if i is not None:
                ## hack alert: stripping out the namespace prefix...
                return (i.text[2:], root)

        for i in root.iter(QName_E('ResponseCode')):
            if i is not None:
                return (i.text, root)

        return (None, root)

    @staticmethod
    def get_error_msg (soap_resp, root=None, code=None):
        if root is not None:
            root = SoapClient.parse_xml(soap_resp)

        if code == 'ErrorSchemaValidation':
            for i in root.iter('faultstring'):
                if i is not None:
                    return (i.text, root)
        else:
            for i in root.iter(QName_M('MessageText')):
                if i is not None:
                    return (i.text, root)

        return (None, root)
