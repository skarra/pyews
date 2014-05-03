#!/usr/bin/python
##
## Created : Wed Mar 05 11:28:41 IST 2014
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

import logging, re

import utils
from   utils            import pretty_xml
from   ews.autodiscover import EWSAutoDiscover, ExchangeAutoDiscoverError
from   ews.data         import DistinguishedFolderId, WellKnownFolderName
from   ews.data         import FolderClass
from   ews.errors       import EWSMessageError, EWSCreateFolderError
from   ews.errors       import EWSDeleteFolderError
from   ews.folder       import Folder
from   ews.contact      import Contact

from ews.request_response import GetItemsRequest, GetItemsResponse
from ews.request_response import FindItemsRequest, FindItemsResponse
from ews.request_response import CreateItemsRequest, CreateItemsResponse
from ews.request_response import DeleteItemsRequest, DeleteItemsResponse
from ews.request_response import FindItemsLMTRequest, FindItemsLMTResponse
from ews.request_response import UpdateItemsRequest, UpdateItemsResponse
from ews.request_response import SyncFolderItemsRequest, SyncFolderItemsResponse

from   tornado import template
from   soap import SoapClient, SoapMessageError, QName_T

USER = u''
PWD  = u''
EWS_URL = u''

##
## Note: There is a feeeble attemp to mimick the names of classes and methods
## used in the EWS Managed Services API. However the similiarities are merely
## skin-deep, if anything at all.
##

class InvalidUserEmail(Exception):
    pass

class WebCredentials(object):
    def __init__ (self, user, pwd):
        self.user = user
        self.pwd  = pwd

class ExchangeService(object):
    def __init__ (self):
        self.ews_ad = None
        self.credentials = None
        self.root_folder = None
        self.loader = template.Loader(utils.REQUESTS_DIR)

    ##
    ## First the methods that are similar to the EWS Managed API. The names might
    ## be similar but please note that there is no effort made to really be a
    ## complete copy of the Managed API.
    ##
    ## FIXME: Each of these results in a EWS Request. We should just have a
    ## base Request class along with specific request types derived from the
    ## based request that does the needful. Hm. there is no end to this
    ## 'properisation'.
    ##

    def AutoDiscoverUrl (self):
        """blame the weird naming on the EWS MS APi."""

        creds = self.credentials
        self.ews_ad = EWSAutoDiscover(creds.user, creds.pwd)
        self.Url = self.ews_ad.discover()

    def CreateFolder (self, parent_id, info):
        """info should be an array of (name, class) tuples. class should be one
        of values in the ews.data.FolderClass enumeration.
        """

        logging.info('Sending folder create request to EWS...')
        req = self._render_template(utils.REQ_CREATE_FOLDER,
                                    parent_folder_id=parent_id,
                                    folders=info)
        try:
            resp, node = self.send(req)
        except SoapMessageError as e:
            raise EWSMessageError(e.resp_code, e.xml_resp, e.node)

        logging.info('Sending folder create request to EWS...done')
        return Folder(self, resp)

    def DeleteFolder (self, folder_ids, hard_delete=False):
        """Delete all specified folder ids. If hard_delete is True then the
        folders are completely nuked otherwise they are pushed to the Dumpster
        if that is enabed server side.
        """
        logging.info('pimdb_ex:DeleteFolder() - deleting folder_ids: %s',
                    folder_ids)

        dt = 'HardDelete'if hard_delete else 'SoftDelete'
        req = self._render_template(utils.REQ_DELETE_FOLDER,
                                    delete_type=dt, folder_ids=folder_ids)
        try:
            resp, node = self.send(req)
            logging.info('pimdb_ex:DeleteFolder() - successfully deleted.')
        except SoapMessageError as e:
            raise EWSMessageError(e.resp_code, e.xml_resp, e.node)

        return resp

    def FindItems (self, folder, eprops_xml=[], ids_only=False):
        """
        Fetch all the items in the given folder.  folder is an object of type
        ews.folder.Folder. This method will first find all the ItemIds of
        contained items, then go back and fetch all the details for each of
        the items in the folder. We return an array of Item objects of the
        right type.

        eprops_xml is an array of xml representation for additional extended
        properites that need to be fetched

        """

        logging.info('pimdb_ex:FindItems() - fetching items in folder %s...',
                     folder.DisplayName)

        i = 0
        ret = []
        while True:
            req = FindItemsRequest(self, batch_size=self.batch_size(),
                                   offset=i, folder_id=folder.Id)
            resp = req.execute()
            shells = resp.items
            if shells is not None and len(shells) > 0:
                ret += shells

            if resp.includes_last:
                break

            i += self.batch_size()
            ## just a safety net to avoid inifinite loops
            if i >= folder.TotalCount:
                logging.warning('pimdb_ex.FindItems(): Breaking strange loop')
                break

        logging.info('pimdb_ex:FindItems() - fetching items in folder %s...done',
                     folder.DisplayName)

        if len(ret) > 0 and ids_only == False:
            return self.GetItems([x.itemid for x in ret],
                                 eprops_xml=eprops_xml)
        else:
            return ret

    def FindItemsLMT (self, folder, lmt):
        """
        Fetch all the items in the given folder that were last modified at or
        after the provided timestamp (lmt). We return an array of Item objects
        of the right type. It will only contain the following fields:

        - ID
        - ChangeKey
        - DisplayName (if it is a contact)
        - Last Modified time

        """

        logging.info('pimdb_ex:FindItemsLMT() - fetching items in folder %s...',
                     folder.DisplayName)

        i = 0
        ret = []
        while True:
            req = FindItemsLMTRequest(self, batch_size=self.batch_size(),
                                      offset=i, folder_id=folder.Id, lmt=lmt)
            resp = req.execute()
            shells = resp.items
            if shells is not None and len(shells) > 0:
                ret += shells

            if resp.includes_last:
                break

            i += self.batch_size()
            ## just a safety net to avoid inifinite loops
            if i >= folder.TotalCount:
                logging.warning('pimdb_ex.FindItemsLMT(): Breaking strange loop')
                break

        logging.info('pimdb_ex:FindItemsLMT() - fetching items in folder %s...done',
                     folder.DisplayName)

        return ret

    def GetItems (self, itemids, eprops_xml=[]):
        """
        itemids is an array of itemids, and we will fetch that stuff and
        return an array of Item objects.

        FIXME: Need to make this work in batches to ensure data is not too
        much.
        """

        logging.info('pimdb_ex:GetItems() - fetching items....')
        req = GetItemsRequest(self, itemids=itemids,
                              custom_eprops_xml=eprops_xml)
        resp = req.execute()
        logging.info('pimdb_ex:GetItems() - fetching items...done')

        return resp.items

    def CreateItems (self, folder_id, items):
        """Create items in the exchange store."""

        logging.info('pimdb_ex:CreateItems() - creating items....')
        req = CreateItemsRequest(self, folder_id=folder_id, items=items)
        resp = req.execute()

        logging.info('pimdb_ex:CreateItems() - creating items....done')

    def DeleteItems (self, itemids):
        """Delete items in the exchange store."""

        logging.info('pimdb_ex:DeleteItems() - deleting items....')
        req = DeleteItemsRequest(self, itemids=itemids)
        logging.info('pimdb_ex:DeleteItems() - deleting items....done')

        return req.execute()

    def UpdateItems (self, items):
        """
        Fetch updates from the specified folder_id.  items in the exchange store.
        """

        logging.info('pimdb_ex:UpdateItems() - updating items....')

        req = UpdateItemsRequest(self, items=items)
        resp = req.execute()

        logging.info('pimdb_ex:UpdateItems() - updating items....done')
        return resp.items

    def SyncFolderItems (self, folder_id, sync_state):
        """
        Fetch updates from the specified folder_id. 
        """

        logging.info('pimdb_ex:SyncFolder() - fetching state...')

        req = SyncFolderItemsRequest(self, folder_id=folder_id,
                                     sync_state=sync_state,
                                     batch_size=self.batch_size())
        resp = req.execute()

        logging.info('pimdb_ex:SyncFolder() - fetching state...done')

        return resp

    ##
    ## Some internal messages
    ##

    ##
    ## Other external methods
    ##

    def init_soap_client (self):
        self.soap = SoapClient(self.Url, user=self.credentials.user,
                               pwd=self.credentials.pwd)

    def send (self, req, debug=False):
        """
        Will raise a SoapConnectionError if there is a connection problem.
        """

        return self.soap.send(req, debug)

    def get_distinguished_folder (self, name):
        elem = u'<t:DistinguishedFolderId Id="%s"/>' % name
        req  = self._render_template(utils.REQ_GET_FOLDER,
                                     folder_ids=elem)
        return self.soap.send(req)

    def get_root_folder (self):
        if not self.root_folder:
            self.root_folder = Folder.bind(self,
                                           WellKnownFolderName.MsgFolderRoot)
        return self.root_folder

    def batch_size (self):
        return 100

    ##
    ## Internal routines
    ##

    def _wsdl_url (self, url=None):
        if not url:
            url = self.Url

        res = re.match('(.*)Exchange.asmx$', url)
        return res.group(1) + 'Services.wsdl'

    ## FIXME: To be removed once all the requests become classes
    def _render_template (self, name, **kwargs):
        return self.loader.load(name).generate(**kwargs)

    ##
    ## Property getter/setter stuff
    ##

    @property
    def credentials (self):
        return self._credentials

    @credentials.setter
    def credentials (self, c):
        self._credentials = c

    @property
    def Url (self):
        return self._Url

    @Url.setter
    def Url (self, url):
        self._Url = url
        self.wsdl_url = self._wsdl_url()
