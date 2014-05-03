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

import pyews.soap as soap
import pyews.utils as utils
from   pyews.soap import SoapClient, QName_M, QName_T, unQName
from   pyews.ews.request_response import GetFolderRequest, GetFolderResponse
from   pyews.ews.request_response import FindFoldersRequest, FindFoldersResponse
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

    def __init__ (self, service, name, node=None):
        self.wkfn = name
        self.service = service

        ## FIXME: These properties are not consistent with our overall naming
        ## conventions. Needs to be Fixed.
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
        self.bind_response_code = None

        self.node = node
        if node is not None:
            self._init_fields(node)

    ##
    ## First the methods that are similar to the EWS Managed API. The names
    ## might be similar but please note that there is no effort made to really
    ## be a complete copy of the Managed API.
    ##

    def FindFolders (self, types=None, recursive=False):
        """Walk through the entire folder hierarchy of the message store and
        return an array of Folder objects of specified types.

        types can be an array of folder types enumerated in
        ews.data.FolderClass
        """

        root_id = self.Id
        ck = self.ChangeKey

        req = FindFoldersRequest(self.service, folder_ids=[(root_id, ck)],
                                 traversal='Deep' if recursive else 'Shallow')
        resp = req.execute()
        if types is not None:
            return [x for x in resp.folders if x.FolderClass in types]
        else:
            return resp.folders

    ##
    ## Class methods
    ##

    @classmethod
    def bind (self, service, wkfn):
        """Bind to a given folder and returs a lt of awesome information about
        this folder. wkfn stands for well known folder name. This method
        returns a Folder object"""

        req = GetFolderRequest(service, folder_name=wkfn)
        resp = req.execute()
        return Folder(service, wkfn, resp.folder_node)

    ##
    ## Internal methods
    ##

    def _init_fields (self, root):
        """root is the parsed XML response pointing to a Folder element"""

        # fid_elem = root.iter(QName_T('FolderId'))
        # if fid_elem is not None:
        #     self.Id = fid_elem.attrib['Id']
        #     self.ChangeKey = fid_elem.attrib['ChangeKey']

        # assert self.Id

        idelem = root.find(QName_T('FolderId'))
        self.Id = idelem.attrib['Id']
        self.ChangeKey = idelem.attrib['ChangeKey']

        self.DisplayName      = self._node_text(root, QName_T('DisplayName'))
        self.ChildFolderCount = self._node_text(root, QName_T('ChildFolderCount'))
        self.FolderClass      = self._node_text(root, QName_T('FolderClass'))
        self.TotalCount       = self._node_text(root, QName_T('TotalCount'))
        self.TotalCount       = int(self.TotalCount)

    def _node_text (self, node, tag):
        c = node.find(tag)
        return c.text if c is not None else None

    ##
    ## External Methods
    ##

    def get_updates (self, sync_state):
        """
        Given a sync_state ID, fetch all the changes since that time. This
        methods returns a tuple (new, mod, del, new_sync_state) -> where each
        element is an array of Contact objects
        """

        resp = self.service.SyncFolderItems(self.Id, sync_state)
        return resp.news, resp.mods, resp.dels, resp.sync_state

    def __str__ (self):
        s = 'Name: %s' % self.DisplayName
        s += '  ID                  : %s\n' % self.Id
        s += '  Parent Folder ID    : %s\n' % self.ParentFolderId
        s += '  Total Count         : %s\n' % self.TotalCount
        s += '  ChildFolderCount    : %s\n' % self.ChildFolderCount
        s += '  WellKnownFolderName : %s\n' % self.wkfn

        return s
