##
## Created : Fri Mar 28 23:02:29 IST 2014
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

import soap, utils
from   soap import SoapClient, QName_M, QName_T, unQName
import xml.etree.ElementTree as ET

import logging

class NotImplementedError:
    pass

class EWSFolderError:
    pass

class Folder:
    """
    Represents a EWS Folder. Names match EWS Managed API as much as possible
    """

    def __init__ (self, service, name, resp=None):
        self.wkfn = name
        self.service = service

        self.ChildFolderCount = None
        self.DisplayName = None
        self.Id = None
        self.ChangeKey = None
        self.IsDirty = False
        self.ParentFolderId = None
        self.Service = service
        self.TotalCount = 0
        self.FolderClass = None

        ## The attributes that do not directly 
        self.bind_response = None

        if resp:
            self._init_fields(resp)

    ##
    ## Class methods
    ##

    @classmethod
    def bind (self, service, wkfn):
        """Bind to a given folder and returs a lt of awesome information about
        this folder. wkfn stands for well known folder name. This method
        returns a Folder object"""

        req = service._render_template(utils.REQ_BIND_FOLDER,
                                       folder_name=wkfn)
        resp, node = service.send(req)
        return Folder(service, wkfn, resp)

    ##
    ## Internal methods
    ##

    def _init_fields (self, resp):
        """resp is the XML response from the server in response to a Bind
        request."""

        resp, root = SoapClient.get_response_code(resp)
        if resp != 'NoError':
            raise EWSFoldeError('Could not bind folder (%s): %s',
                                self.wkfn, resp)
        else:
            self.bind_response = resp

        self.Id, self.ChangeKey = self._get_first_folder_id(root)
        assert(self.Id)

    def _get_first_folder_id (self, root):
        for i in root.iter(QName_T('FolderId')):
            if i is not None:
                return i.attrib['Id'], i.attrib['ChangeKey']

            return None, None

    def _parse_resp_for_folders (self, resp, root=None):
        """Return an array of Folder objects from stuff found in resp. Start
        with the parsed root object if it is not None. Return (ary, root)
        tuple."""

        if not root:
            root = SoapClient.parse_xml(resp)

        dn = QName_T('DisplayName')
        idn = QName_T('FolderId')
        cncn = QName_T('ChildFolderCount')
        
        ret = []
        for folders in root.iter(QName_T('Folders')):
            for folder_elem in folders:
                f = Folder(self.service, None)

                idelem = folder_elem.find(idn)
                f.Id = idelem.attrib['Id']
                f.ChangeKey = idelem.attrib['ChangeKey']
                f.DisplayName = folder_elem.find(dn).text
                f.ChildFolderCount = folder_elem.find(cncn).text
                f.FolderClass = folder_elem.find(QName_T('FolderClass')).text
                f.TotalCount  = folder_elem.find(QName_T('TotalCount')).text

                ret.append(f)
                
            break

        return ret, root

    ##
    ## External Methods
    ##

    def fetch_all_folders (self, root_id=None, ck=None, types=None):
        """Walk through the entire folder hierarchy of the message store and
        print one line per folder with some critical information. 

        This is a recursive function. If you want to start enumerating at the
        root folder of the current message store, invoke this routine without
        any arguments. The defaults will ensure the root folder if fetched and
        folders will be recursively enumerated.

        types can be an array of folder types enumerated in ews.data.FolderClass 
        """

        if not root_id:
            root_id = self.Id

        if not ck:
            ck = self.ChangeKey

        ret = []

        req = self.service._render_template(utils.REQ_FIND_FOLDER_ID,
                                       folder_ids=[(root_id, ck)])
        resp = self.service.send(req)
        folders, root = self._parse_resp_for_folders(resp, None)

        for f in folders:
            f.ParentFolderId = root_id

            if not types or f.FolderClass in types:
                ret.append(f)

            if f.ChildFolderCount > 0:
                logging.debug('Exploring deeper into: %s', f.DisplayName)
                ret += self.fetch_all_folders(f.Id, f.ChangeKey, types)

        return ret
        

    def __str__ (self):
        s = 'Name: %s' % slef.DisplayName
        s += '  ID                  : %s\n' % self.Id
        s += '  Parent Folder ID    : %s\n' % self.ParentFolderId
        s += '  Total Count         : %s\n' % self.TotalCount
        s += '  ChildFolderCount    : %s\n' % self.ChildFolderCount
        s += '  WellKnownFolderName : %s\n' % self.wkfn

        return s
