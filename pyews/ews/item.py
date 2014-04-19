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

class Field:
    __metaclass__ = ABCMeta

    def __init__ (self, tag=None, text=None):
        self.tag = tag
        self.text = text
        self.attrib = {}
        self.children = []

    def add_attrib (self, key, val):
        self.attrib.update({key: val})

    def to_xml (self):
        self.children = self.get_children()

        if ((self.text is not None) or (len(self.children) > 0)):
            ats = ['%s="%s"' % (k, v) for k, v in self.attrib.iteritems() if v]
            xmls = [x.to_xml() for x in self.children]
            cs =  '\n'.join([y for y in xmls if y is not None])
            text = self.text if self.text is not None else ''

            ret =  '<t:%s %s>%s%s</t:%s>' % (self.tag, ' '.join(ats), text, cs,
                                             self.tag)
            return ret
        else:
            return ''

    def get_children (self):
        return self.children

    def __str__ (self):
        return self.text if self.text is not None else ""

class ItemId(Field):
    def __init__ (self, text=None):
        Field.__init__(self, 'ItemId', text)

class ChangeKey(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'ChangeKey'

class ParentFolderId(Field):
    def __init__ (self, text=None):
        Field.__init__(self, 'ParentFolderId', text)

class ParentFolderChangeKey(Field):
    def __init__ (self, text=None):
        Field.__init__(self, 'ParentFolderChangeKey', text)

class ItemClass(Field):
    def __init__ (self, text=None):
        Field.__init__(self, 'ItemClass', text)
 
## This might not be so easy...
class LastModifiedTime(Field):
    def __init__ (self, text=None):
        Field.__init__(self, 'LastModifiedTime', text)

class DateTimeCreated(Field):
    def __init__ (self, text=None):
        Field.__init__(self, 'DateTimeCreated', text)

class Item(Field):
    """
    Abstract wrapper class around an Exchange Item object. Frequently an
    object of this type is instantiated from a response.
    """

    __metaclass__ = ABCMeta

    def __init__ (self, service, parent_fid=None, resp_node=None, tag='Item'):
        Field.__init__(self, tag=tag)

        self.ParentFolderId = parent_fid             # folder object
        self.service = service                       # Exchange service object
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
                self.itemid = ItemId(child.attrib['Id'])
                self.change_key = ChangeKey(child.attrib['ChangeKey'])
            elif tag == 'ParentFolderId':
                self.parent_fid = ParentFolderId(child.attrib['Id'])
                self.parent_fck = ParentFolderChangeKey(child.attrib['ChangeKey'])
            elif tag == 'ItemClass':
                self.item_class = ItemClass(child.text)
            elif tag == 'LastModifiedTime':
                self.last_modified_time = LastModifiedTime(child.text)
            elif tag == 'DateTimeCreated':
                self.created_time = DateTimeCreated(child.text)
