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
from   utils            import pretty_xml, clean_xml
from   ews.autodiscover import EWSAutoDiscover, ExchangeAutoDiscoverError
from   ews.data         import DistinguishedFolderId, WellKnownFolderName
from   ews.data         import FolderClass, EWSMessageError
from   ews.data         import EWSCreateFolderError, EWSDeleteFolderError
from   ews.folder       import Folder
from   ews.contact      import Contact

from   tornado import template
from   soap import SoapClient, SoapMessageError, QName_T

gna = SoapClient.get_node_attribute

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

    def FindItems (self, folder):
        """
        Fetch all the items in the given folder.  folder is an object of type
        ews.folder.Folder. This method will first find all the ItemIds of
        contained items, then go back and fetch all the details for each of
        the items in the folder. We return an array of Item objects of the
        right type.
        """

        logging.info('pimdb_ex:FindItems() - fetching items in folder %s...',
                     folder.DisplayName)

        i = 0
        ret = []
        while True:
            req = self._render_template(utils.REQ_FIND_ITEM,
                                        batch_size=self.batch_size(),
                                        offset=i,
                                        folder_id=folder.Id)
            try:
                resp, node = self.send(req)
                last, ign = gna(resp, node, 'RootFolder',
                                'IncludesLastItemInRange')
                shells = self._construct_items(resp, node)
                if shells is not None and len(shells) > 0:
                    ret += shells

                if last == "true":
                    break
            except SoapMessageError as e:
                raise EWSMessageError(e.resp_code, e.xml_resp, e.node)

            i += self.batch_size()
            ## just a safety net to avoid inifinite loops
            if i >= folder.TotalCount:
                logging.warning('pimdb_ex.FindItems(): Breaking strange loop')
                break

        logging.info('pimdb_ex:FindItems() - fetching items in folder %s...done',
                     folder.DisplayName)

        if len(ret) > 0:
            return self.GetItems([x.itemid for x in ret])
        else:
            return ret

    def GetItems (self, itemids):
        """
        itemids is an array of itemids, and we will fetch that stuff and
        return an array of Item objects.

        FIXME: Need to make this work in batches to ensure data is not too
        much.
        """

        logging.info('pimdb_ex:GetItems() - fetching items....')
        req = self._render_template(utils.REQ_GET_ITEM, itemids=itemids)
        try:
            resp, node = self.send(req)
            logging.debug('%s', pretty_xml(resp))
        except SoapMessageError as e:
            raise EWSMessageError(e.resp_code, e.xml_resp, e.node)

        logging.info('pimdb_ex:GetItems() - fetching items...done')
        return self._construct_items(resp, node)

    def CreateItems (self, folder_id, items):
        """Create items in the exchange store."""

        logging.info('pimdb_ex:CreateItems() - creating items....')
        req = self._render_template(utils.REQ_CREATE_ITEM,
                                    folder_id=folder_id, items=items)
        try:
            req = clean_xml(req)
            resp, node = self.send(req)
            logging.debug('%s', pretty_xml(resp))
        except SoapMessageError as e:
            raise EWSMessageError(e.resp_code, e.xml_resp, e.node)

        logging.info('pimdb_ex:CreateItems() - creating items....done')

    ##
    ## Some internal messages
    ##

    def _construct_items (self, resp, node=None):
        if node is not None:
            node = SoapClient.parse_xml(resp)

        ret = []
        ## As we support additional item types we will add more such loops.
        for cxml in node.iter(QName_T('Contact')):
            ret.append(Contact(self, resp_node=cxml))

        return ret

    ##
    ## Other external methods
    ##

    def init_soap_client (self):
        self.soap = SoapClient(self.Url, user=self.credentials.user,
                               pwd=self.credentials.pwd)

    def send (self, req):
        """
        Will raise a SoapConnectionError if there is a connection problem.
        """

        return self.soap.send(req)

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
