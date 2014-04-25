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
from pyews.ews     import mapitags
import       xml.etree.ElementTree as ET
from pyews.ews.data import MapiPropertyTypeType, MapiPropertyTypeTypeInv

gnd = SoapClient.get_node_detail

class ReadOnly:
    """
    When applied as a Mixin, this class will ensure that no XML is generated
    when this field is written out as part of the containing Item. Usage for
    this:

    class LastModifiedTime(ReadOnly, ExtendedProperty):
        pass

    Note that for this to work in python (given the rules of method resolution
    order for multiple inheritance, the ReadOnly Mixin should come to the left
    of Field class or any its derived classes.
    """
    __metaclass__ = ABCMeta

    def write_to_xml (self):
        return ''

class Field:
    """
    Represents an XML Element
    """
    __metaclass__ = ABCMeta

    def __init__ (self, tag=None, text=None):
        self.tag = tag
        self.value = text
        self.attrib = {}
        self.children = []
        self.read_only = False

    @property
    def value (self):
        return self._value

    @value.setter
    def value (self, val):
        self._value = val
        
    def add_attrib (self, key, val):
        self.attrib.update({key: val})

    def atts_as_xml (self):
        ats = ['%s="%s"' % (k, v) for k, v in self.attrib.iteritems() if v]
        return ' '.join(ats)

    def value_as_xml (self):
        return self.value if self.value is not None else ''

    def children_as_xml (self):
        self.children = self.get_children()
        xmls = [x.write_to_xml() for x in self.children]
        return '\n'.join([y for y in xmls if y is not None])

    def write_to_xml (self):
        """
        Return an XML representation of this field.
        """

        if ((self.value is not None) or (len(self.attrib) > 0) or
            (len(self.children) > 0)):

            text = self.value_as_xml()
            ats = self.atts_as_xml()
            cs = self.children_as_xml()

            ret =  '<t:%s %s>%s%s</t:%s>' % (self.tag, ats, text, cs, self.tag)
            return ret
        else:
            return ''

    def write_to_xml_get (self):
        """
        Presently only makes sense for certain ExtendedProperties
        """

        ats = self.atts_as_xml()
        ret =  '<t:%s %s/>' % (self.tag, ats)
        return ret

    def get_children (self):
        return self.children

    def set (self, value):
        self.value = value

    def __str__ (self):
        return self.value if self.value is not None else ""

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
 
class DateTimeCreated(Field):
    def __init__ (self, text=None):
        Field.__init__(self, 'DateTimeCreated', text)

class PropVariant:
    UNKNOWN   = 0
    TAGGED    = 1
    NAMED_NUM = 2
    NAMED_STR = 3

class ExtendedProperty(Field):
    class ExtendedFieldURI(Field):
        def __init__ (self, node=None, dis_psetid=None, psetid=None,
                      ptag=None, pname=None, pid=None, ptype=None):
            Field.__init__(self, 'ExtendedFieldURI')

            ## See here for explanation of each of these and what the
            ## constraints are for mixing and matchig these fields.
            ## http://msdn.microsoft.com/en-us/library/office/aa564843.aspx

            ## Note - All of these values will be strings (if not None)
            self.attrib = {
                'DistinguishedPropertySetId' : dis_psetid,
                'PropertySetId'  : psetid,
                'PropertyTag'    : ptag,
                'PropertyName'   : pname,
                'PropertyId'     : pid,
                'PropertyType'   : ptype,
            }

            if node is not None:
                self.init_from_xml(node)

        def init_from_xml (self, node):
            """
            None is a parsed xml node (Element). Extract the data that we can
            from the node.
            """

            uri = node.find(QName_T('ExtendedFieldURI'))
            if uri is None:
                logging.debug('ExtendedProperty.init_from_xml(): no child node ' +
                              'ExtendedFieldURI in node: %s',
                              pretty_xml(node))
            else :
                self.attrib.update(uri.attrib)

    class Value(Field):
        def __init__ (self, text=None):
            Field.__init__(self, 'Value')

    def __init__ (self, node=None, dis_psetid=None, psetid=None,
                      ptag=None, pname=None, pid=None, ptype=None):
        """
        If node is not None, it should be a parsed XML element pointing to an
        Extended Property element
        """

        ## This guy needs to be here as we have overriden the .value()
        self.val = self.Value()

        Field.__init__(self, 'ExtendedProperty')
        self.efuri = self.ExtendedFieldURI(node, dis_psetid, psetid,
                                           ptag, pname, pid, ptype)

        ## FIXME: We can have a multi-valued property as well.
        if node is not None:
            self.value = node.find(QName_T('Value')).text

        self.children = [self.efuri, self.val]

    @property
    def value (self):
        return self.val.value

    @value.setter
    def value (self, text):
        self.val.value = text

    ##
    ## Overriding inherted methods
    ##

    def value_as_xml (self):
        ## Any attribute that has a 'Value' child should have its own value
        ## printed out as text in the XML represented. Note for extended
        ## properties self.value returns self.val.value for ease of access. So
        ## if we do not do this the value will go out twide

        return ''

    def write_to_xml_get (self):
        """
        Presently only makes sense for certain ExtendedProperties
        """

        return self.efuri.write_to_xml_get()

    def get_variant (self):
        """
        Outlook has three types of properties - Tagged Properties, Named
        Properites with Numeric Identifiers, and Named Properties with
        string identifiers. This method will identify which variant this
        property is. It will return one of the values defined in the
        PropVariant class above.
        """

        all_none = True
        for v in self.efuri.attrib.values():
            if v is not None:
                all_none = False
                break
        if all_none:
            return PropVariant.UNKNOWN

        if (self.efuri.attrib['PropertyTag'] is not None
            and self.efuri.attrib['PropertyType'] is not None
            and self.efuri.attrib['PropertySetId'] is None
            and self.efuri.attrib['DistinguishedPropertySetId'] is None
            and self.efuri.attrib['PropertyId'] is None):
            return PropVariant.TAGGED

        ## FIXME: Support checks for the other variants.
        return PropVariant.UNKNOWN

    def get_prop_tag (self):
        """Return a tag version of the current property. Note that the
        name tag is used in different ways by MS in different places. Here
        we are talking about the combined property_tag and property_type
        entity."""

        pt = self.efuri.attrib['PropertyTag']
        base = 16 if pt[0:2] == "0x" else 10
        pid = int(pt, base)
        ptype = int(MapiPropertyTypeTypeInv[self.efuri.attrib['PropertyType']])

        return mapitags.PROP_TAG(ptype, pid)

    def set (self, value):
        self.value.set(value)

    ##
    ## Some helper methods to easily "recognize" the extended properties
    ##

    @staticmethod
    def is_tagged_prop (xml_node):
        """
        If the element is a tagged property then returns (True, MAPITag)
        otherwise it will be (False, ignore)
        """
        tp = (len(xml_node.attrib) == 2 and
              'PropertyTag' in xml_node.attrib and
              'PropertyType' in xml_node.attrib)

        if tp:
            pid   = utils.safe_int(xml_node.attrib['PropertyTag'])
            ptype = xml_node.attrib['PropertyType']
            tag = mapitags.PROP_TAG(MapiPropertyTypeTypeInv[ptype], pid)

            return (True, tag)
        else:
            return (False, 0)

class LastModifiedTime(ReadOnly, ExtendedProperty):
    def __init__ (self, node=None, text=None):
        ptag  = mapitags.PROP_ID(mapitags.PR_LAST_MODIFICATION_TIME)
        ptype = mapitags.PROP_TYPE(mapitags.PR_LAST_MODIFICATION_TIME)
        ExtendedProperty.__init__(self, node=node, ptag=ptag, ptype=ptype)

class Item(Field):
    """
    Abstract wrapper class around an Exchange Item object. Frequently an
    object of this type is instantiated from a response.
    """

    __metaclass__ = ABCMeta

    def __init__ (self, service, parent_fid=None, resp_node=None, tag='Item'):
        Field.__init__(self, tag=tag)

        self.service = service                       # Exchange service object
        self.resp_node = resp_node

        self.parent_fid = ParentFolderId(parent_fid)
        self.parent_fck = None

        ## Now initialize some of the properties to default values
        self.itemid = ItemId()
        self.item_class = ItemClass()
        self.change_key = ChangeKey()
        self.created_time = DateTimeCreated()
        self.last_modified_time = None

        self.eprops = []
        self.eprops_tagged = {}

        if self.resp_node is not None:
            self._init_base_fields_from_resp(resp_node)

    ##
    ## First the abstract methods that will be implementd by sub classes
    ##

    @abstractmethod
    def add_extended_property (self, node):
        """
        Parse the XML element pointed to by node, figure out the type of
        property this is, initialize the property of the right type and then
        insert that into the self.eprops member variable
        """
        raise NotImplementedError

    @abstractmethod
    def add_tagged_property (self, node=None, tag=None, value=None):
        """
        Add a tagged property either from an XML node or from the individual
        components. If node is not None the other parameters are ignored.
        """
        raise NotImplementedError

    def add_named_str_property (self, node=None, psetid=None, pname=None,
                                ptype=None, value=None):
        raise NotImplementedError

    def add_named_int_property (self, node=None, psetid=None, pid=None, ptype=None,
                                value=None):
        raise NotImplementedError

    ##
    ## Next, the non-abstract external methods
    ##

    def get_extended_properties (self):
        return self.eprops

    ## properties with their getter and setters. This is just a placeholder
    ## for now. We will not implement all of them

    @property
    def parent_folder_id (self):
        return self.parent_folder_id

    @parent_folder_id.setter
    def parent_folder_id (self, val):
        self.parent_folder_id = val

    ##
    ## Finally, some internal methods and helper functions
    ##

    def _find_text_safely (self, elem, node):
        r = elem.find(node)
        return r.value if r else None

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
