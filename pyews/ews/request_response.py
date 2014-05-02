##
## Created : Thu May 01 11:00:15 IST 2014
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

import logging, traceback
import pyews.utils as utils

from   abc            import ABCMeta, abstractmethod
from   pyews.soap     import QName_T, QName_M
from   pyews.utils    import pretty_xml, clean_xml
from   pyews.ews.contact    import Contact

##
## Base classes
##

class Request(object):
    __metaclass__ = ABCMeta

    def __init__ (self, ews, template=None):
        self.ews = ews
        self.template = template
        self.kwargs = None
        self.resp = None

    ##
    ## Abstract methods
    ##

    @abstractmethod
    def execute (self):
        pass

    ##
    ## Public methods
    ##

    def request_server (self, debug=False):
        r = self.ews.loader.load(self.template).generate(**self.kwargs)
        if debug:
            logging.debug('Request: %s', pretty_xml(r))
        return self.ews.send(r, debug)

class Response(object):
    def __init__ (self, req, node):
        self.req = req
        self.node = node
        self.err_cnt = 0
        self.suc_cnt = 0
        self.war_cnt = 0
        self.errors = {}

    def parse (self, tag):
        """
        Look in the present response node for all child nodes of given tag,
        while looking for errors and warnings as well.
        """

        assert self.node is not None

        i = 0
        for gfrm in self.node.iter(tag):
            resp_class = gfrm.attrib['ResponseClass']
            if resp_class == 'Error':
                self.err_cnt += 1
                self.errors.update({i: EWSErrorElement(gfrm)})
            elif resp_class == 'Warning':
                self.war_cnt += 1
                ## FIXME: Need to handle these
                logging.error('Ignoring a warning response class.')
            else:
                self.suc_cnt += 1

            i += 1

        ## Some nodes have a tag called faultcode. Here is some old code that
        ## might be handy at some point

        # for i in root.iter('faultcode'):
        #     if i is not None:
        #         ## hack alert: stripping out the namespace prefix...
        #         return i.text[2:]

        if self.has_errors():
            logging.error('Response.parse: Found %d errors', self.err_cnt)
            for ind, err in self.errors.iteritems():
                logging.error('  Item num %02d - %s', ind, str(err))

    def has_errors (self):
        return self.err_cnt > 0

class EWSErrorElement(object):
    """
    Wraps an XML response element that represents an erorr response from the
    server
    """

      # <m:GetItemResponseMessage ResponseClass="Error">
      #     <m:MessageText>Id is malformed.</m:MessageText>
      #     <m:ResponseCode>ErrorInvalidIdMalformed</m:ResponseCode>
      #     <m:DescriptiveLinkKey>0</m:DescriptiveLinkKey>
      #     <m:Items />
      # </m:GetItemResponseMessage>

    def __init__ (self, node):
        self.node = node
        
        t = node.find(QName_M('MessageText'))
        self.msg_text = t.text if t is not None else None
    
        t = node.find(QName_M('ResponseCode'))
        self.resp_code = t.text if t is not None else None

        t = node.find(QName_M('DescriptiveLinkKey'))
        self.des_link_key = t.text if t is not None else None

    def __str__ (self):
        return 'Code: %s; Text: %s' % (self.resp_code, self.msg_text)

##
## Bind
##

class GetFolderRequest(Request):
    def __init__ (self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_BIND_FOLDER)
        self.kwargs = kwargs

    ##
    ## Implement the abstract methods
    ##

    def execute (self):
        self.resp_node = self.request_server(debug=False)
        self.resp_obj = GetFolderResponse(self, self.resp_node)

        return self.resp_obj

class GetFolderResponse(Response):
    def __init__ (self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def init_from_node (self, node):
        """
        node is a parsed XML Element containing the response
        """

        self.parse(QName_M('GetFolderResponseMessage'))
        for child in self.node.iter(QName_T('Folder')):
            self.folder_node = child
            break

##
## FindFolders
##

class FindFoldersRequest(Request):
    def __init__ (self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_FIND_FOLDER_ID)
        self.kwargs = kwargs

    ##
    ## Implement the abstract methods
    ##

    def execute (self):
        traceback.print_stack()
        self.resp_node = self.request_server(debug=False)
        self.resp_obj = FindFoldersResponse(self, self.resp_node)

        return self.resp_obj

class FindFoldersResponse(Response):
    def __init__ (self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def init_from_node (self, node):
        """
        node is a parsed XML Element containing the response
        """

        from   pyews.ews.folder import Folder as F

        self.parse(QName_M('FindFolderResponseMessage'))

        self.folders = []
        for folders  in self.node.iter(QName_T('Folders')):
            for child in folders:
                self.folders.append(F(self.req.ews, None, node=child))
            break

##
## GetItems
##

class GetItemsRequest(Request):
    def __init__ (self, ews, **kwargs):
        Request.__init__(self, ews, template=utils.REQ_GET_ITEM)
        self.kwargs = kwargs

    ##
    ## Implement the abstract methods
    ##

    def execute (self):
        self.resp_node = self.request_server(debug=False)
        self.resp_obj = GetItemsResponse(self, self.resp_node)

        return self.resp_obj

class GetItemsResponse(Response):
    def __init__ (self, req, node=None):
        Response.__init__(self, req, node)

        if node is not None:
            self.init_from_node(node)

    def init_from_node (self, node):
        """
        node is a parsed XML Element containing the response
        """

        self.parse(QName_M('GetItemResponseMessage'))

        self.items = []
        ## FIXME: As we support additional item types we will add more such
        ## loops.
        for cxml in self.node.iter(QName_T('Contact')):
            self.items.append(Contact(self, resp_node=cxml))
