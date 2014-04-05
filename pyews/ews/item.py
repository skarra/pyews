##
## Created : Wed Apr 02 18:02:51 IST 2014
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

from         abc     import ABCMeta, abstractmethod
from         pyews.soap    import SoapClient, QName_M, QName_T, unQName
import       pyews.soap as soap
import       pyews.utils as utils
import       xml.etree.ElementTree as ET

gnd = SoapClient.get_node_detail

class Item:
    """
    Abstract wrapper class around an Exchange Item object. Frequently an
    object of this type is instantiated from a response.
    """

    __metaclass__ = ABCMeta

    def __init__ (self, service, parent=None, resp_node=None):
        self.parent = parent              # folder object
        self.service = service            # Exchange service object
        self.resp_node = resp_node

        if self.resp_node is not None:
            self._init_base_fields_from_resp(resp_node)

    def _find_text_safely (self, elem, node):
        r = elem.find(node)
        return r.text if r else None

    def _init_base_fields_from_resp (self, rnode):
        """Return a reference to the parsed Element object for response after
        snarfing all the common fields."""

        for child in rnode:
            tag = unQName(child.tag)

            if tag == 'ItemId':
                self.ItemId = child.attrib['Id']
                self.ChangeKey = child.attrib['ChangeKey']
            elif tag == 'ParentFolderId':
                self.ParentFolderId = child.attrib['Id']
                self.ParentFolderChangeKey = child.attrib['ChangeKey']
            elif tag == 'ItemClass':
                self.ItemClass = child.text
            elif tag == 'LastModifiedTime':
                self.LastModifiedTime = child.text
            elif tag == 'DateTimeCreated':
                self.DateTimeCreated = child.text
