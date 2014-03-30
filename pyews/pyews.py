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
from   folder import Folder

from   tornado import template
from   soap import SoapClient

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
        self.loader = template.Loader(utils.REQUESTS_DIR)

    ##
    ## First the methods that are similar to the EWS Managed API. The names might
    ## be similar but please note that there is no effort made to really be a
    ## complete copy of the Managed API.
    ##

    def AutoDiscoverUrl (self):
        """blame the weird naming on the EWS MS APi."""

        creds = self.credentials
        self.ews_ad = EWSAutoDiscover(creds.user, creds.pwd)
        self.Url = self.ews_ad.discover()

    def CreateFolder (self, parent_id, info):
        """
        info should be an array of (name, class) tuples. class should be one
        of values in the ews.data.FolderClass enumeration.
        """

        req = self._render_template(utils.REQ_CREATE_FOLDER,
                                    parent_folder_id=parent_id,
                                    folders=info)
        return self.send(req)

    ##
    ## Other external methods
    ##

    def init_soap_client (self):
        self.soap = SoapClient(self.Url, user=self.credentials.user,
                               pwd=self.credentials.pwd)

    def send (self, req):
        return self.soap.send(req)

    def get_distinguished_folder (self, name):
        elem = u'<t:DistinguishedFolderId Id="%s"/>' % name
        req  = self._render_template(utils.REQ_GET_FOLDER,
                                     folder_ids=elem)
        return self.soap.send(req)

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
