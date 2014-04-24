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

from item    import Item, Field, ExtendedProperty, LastModificationTime
from item    import PropVariant
from pyews.soap    import SoapClient, unQName, QName_T
from pyews.utils   import pretty_xml
from pyews.ews     import mapitags
from pyews.ews.data import MapiPropertyTypeType, MapiPropertyTypeTypeInv
from pyews.ews.data import GenderType

gnd = SoapClient.get_node_detail

class FileAs(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'FileAs'

class Alias(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'Alias'

class DisplayName(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'DisplayName'

class CompleteName(Field):
    ##
    ## Exchnage handling of some of the name fields can, at best, be described
    ## as strange.... Some of these fields can be found outside of this
    ## complex data type. However we will try to make it simple and only use
    ## CompelteName field when possible
    ##
    class Title(Field):
        def __init__ (self, text=None):
            Field.__init__(self, 'Title', text)

    class FirstName(Field):
        def __init__ (self, text=None):
            Field.__init__(self, 'FirstName', text)

    class GivenName(Field):
        def __init__ (self, text=None):
            Field.__init__(self, 'GivenName', text)

    class MiddleName(Field):
        def __init__ (self, text=None):
            Field.__init__(self, 'MiddleName', text)

    class LastName(Field):
        def __init__ (self, text=None):
            Field.__init__(self, 'LastName', text)

    class Surname(Field):
        def __init__ (self, text=None):
            Field.__init__(self, 'Surname', text)

    class Suffix(Field):
        def __init__ (self, text=None):
            Field.__init__(self, 'Suffix', text)

    class Initials(Field):
        def __init__ (self, text=None):
            Field.__init__(self, 'Initials', text)

    class Nickname(Field):
        def __init__ (self, text=None):
            Field.__init__(self, 'Nickname', text)

    class FullName(Field):
        def __init__ (self, text=None):
            Field.__init__(self, 'FullName', text)

    def __init__ (self, text=None):
        Field.__init__(self, 'CompleteName', text)

        self.title        = self.Title()
        self.given_name   = self.GivenName()
        self.first_name   = self.FirstName()
        self.middle_name  = self.MiddleName()
        self.last_name    = self.LastName()
        self.surname      = self.Surname()
        self.suffix       = self.Suffix()
        self.initials     = self.Initials()
        self.nickname     = self.Nickname()
        self.full_name    = self.FullName()

        self.children = [self.title, self.first_name, self.middle_name,
                         self.last_name, self.suffix, self.initials,
                         self.nickname]

class SpouseName(Field):
    def __init__ (self, text=None):
        Field.__init__(self, 'SpouseName', text)

class JobTitle(Field):
    def __init__ (self, text=None):
        Field.__init__(self, 'JobTitle', text)

class CompanyName(Field):
    def __init__ (self, text=None):
        Field.__init__(self, 'CompanyName', text)

class Department(Field):
    def __init__ (self, text=None):
        Field.__init__(self, 'Department', text)

class Manager(Field):
    def __init__ (self, text=None):
        Field.__init__(self, 'Manager', text)

class AssistantName(Field):
    def __init__ (self, text=None):
        Field.__init__(self, 'AssistantName', text)

## FIXME: The values for these are strings, but might be better represented as
## DateTime objects... Hm
class Birthday(Field):
    def __init__ (self, text=None):
        Field.__init__(self, 'Birthday', text)

class WeddingAnniversary(Field):
    def __init__ (self, text=None):
        Field.__init__(self, 'WeddingAnniversary', text)

class Notes(Field):
    def __init__ (self, text=None):
        Field.__init__(self, 'Body', text)

class EmailAddresses(Field):
    class Email(Field):
        def __init__ (self, text=None):
            Field.__init__(self, 'Entry', text)
            self.attrib = {
                'Key' : None,
                'Name' : None,
                'RoutingType' : None,
                'MailBoxType' : None
                }

        def __str__ (self):
            return 'Key: %8s  Address: %s' % (self.attrib['Key'], self.value)

        def __repr__ (self):
            return self.__str__()

    def __init__ (self, node=None):
        Field.__init__(self, 'EmailAddresses')
        self.children = self.entries = []
        if node is not None:
            self.populate_from_node(node)

    def populate_from_node (self, node):
        for child in node:
            email = self.Email()
            for k, v in child.attrib.iteritems():
                email.add_attrib(k, v)

            email.value = child.text
            self.entries.append(email)

    def __str__ (self):
        s = '%s Numbers: ' % len(self.entries)
        s += '; '.join([str(x) for x in self.entries])
        return s

    def __repr__ (self):
        return self.__str__()

class PhoneNumbers(Field):
    class Phone(Field):
        def __init__ (self, text=None):
            Field.__init__(self, 'Entry', text)
            self.attrib = {
                'Key' : None,               # ews.data.PhoneKey
                }

        def __str__ (self):
            return 'Key: %8s  Number: %s' % (self.attrib['Key'], self.value)

        def __repr__ (self):
            return self.__str__()

    def __init__ (self, node=None):
        Field.__init__(self, 'PhoneNumbers')
        self.children = self.entries = []

        if node is not None:
            self.populate_from_node(node)

    def populate_from_node (self, node):
        for child in node:
            phone = self.Phone()
            for k, v in child.attrib.iteritems():
                phone.add_attrib(k, v)

            phone.value = child.text
            self.entries.append(phone)

    def __str__ (self):
        s = '%s Numbers: ' % len(self.entries)
        s += '; '.join([str(x) for x in self.entries])
        return s

    def __repr__ (self):
        return self.__str__()

##
## Extended Properties
##

class Gender(ExtendedProperty):
    def __init__ (self, node=None, text=GenderType.Unspecified):
        ptag  = mapitags.PROP_ID(mapitags.PR_GENDER)
        ptype = mapitags.PROP_TYPE(mapitags.PR_GENDER)
        ExtendedProperty.__init__(self, node=node, ptag=ptag,
                                  ptype=MapiPropertyTypeType[ptype])

    def __str__ (self):
        v = self.value.value
        if v is None:
            return 'Unspecified'
        elif int(v) == GenderType.Female:
            return 'Female'
        elif int(v) == GenderType.Male:
            return 'Male'
        else:
            return 'Unspecified'

class Contact(Item):
    """
    Abstract wrapper class around an Exchange Item object. Frequently an
    object of this type is instantiated from a response.
    """

    def __init__ (self, service, parent_fid=None, resp_node=None):
        Item.__init__(self, service, parent_fid, resp_node, tag='Contact')

        self.file_as      = FileAs()
        self.alias        = Alias()
        self.complete_name = CompleteName()
        self.display_name = DisplayName()
        self.spouse_name = SpouseName()

        self.job_title      = JobTitle()
        self.company_name   = CompanyName()
        self.department     = Department()
        self.manager        = Manager()
        self.assistant_name = AssistantName()

        self.birthday    = Birthday()
        self.anniversary = WeddingAnniversary()

        self.notes = Notes()
        self.emails = EmailAddresses()
        self.phones = PhoneNumbers()

        self.gender = Gender()

        self._init_from_resp()

    def _init_from_resp (self):
        if self.resp_node is None:
            return

        rnode = self.resp_node
        for child in rnode:
            tag = unQName(child.tag)

            if tag == 'FileAs':
                self.file_as.value = child.text
            elif tag == 'Alias':
                self.alias.value = child.text
            elif tag == 'SpouseName':
                self.spouse_nam.value = child.text
            elif tag == 'JobTitle':
                self.job_title.value = child.text
            elif tag == 'CompanyName':
                self.company_name.value = child.text
            elif tag == 'Department':
                self.department.value = child.text
            elif tag == 'Manager':
                self.manager.value = child.text
            elif tag == 'AssistantName':
                self.assistant_name.value = child.text
            elif tag == 'Birthday':
                self.birthday.value = child.text
            elif tag == 'WeddingAnniversary':
                self.anniversary.value = child.text
            elif tag == 'GivenName':
                self.complete_name.given_name.value = child.text
            elif tag == 'Surname':
                self.complete_name.surname.value = child.text
            elif tag == 'Initials':
                self.complete_name.initials.value = child.text
            elif tag == 'DisplayName':
                self.display_name.value = child.text
            elif tag == 'Body':
                ## FIXME: We are assuming a text body type, but they could
                ## contain html or other types as well... Oh, well.
                self.notes.value = child.text
            elif tag == 'EmailAddresses':
                self.emails.populate_from_node(child)
            elif tag == 'PhoneNumbers':
                self.phones.populate_from_node(child)
            elif tag == 'ExtendedProperty':
                self.add_extended_property(node=child)

        n = rnode.find('CompleteName')
        if n is not None:
            print 'Yo, ma =========================='
            rnode = n

        fts = self._find_text_safely
        self.complete_name.title.value = fts(rnode, 'Title')
        self.complete_name.first_name.value = fts(rnode, 'FirstName')
        self.complete_name.middle_name.value = fts(rnode, 'MiddleName')
        self.complete_name.last_name.value = fts(rnode, 'LastName')
        self.complete_name.suffix.value = fts(rnode, 'Suffix')
        self.complete_name.initials.value = fts(rnode, 'Initials')
        self.complete_name.full_name.value = fts(rnode, 'FullName')
        self.complete_name.nickname.value = fts(rnode, 'Nickname')

        ## It's a bit hard to understand why the hell they have so many
        ## variants for the same stupid information... Oh well, let's just
        ## have a few handy shortcuts for the information that matters
        if self.complete_name.first_name.value:
            self._firstname = self.complete_name.first_name.value
        else:
            self._firstname = self.complete_name.given_name.value

        if self.complete_name.last_name.value:
            self._lastname = self.complete_name.last_name.value
        else:
            self._lastname = self.complete_name.surname.value

        if self._firstname is None:
            self._firstname = ''

        if self._lastname is None:
            self._lastname = ''

        if self.display_name.value:
            self._displayname = self.display_name.value
        elif self.complete_name.full_name.value:
            self._displayname = self.complete_name.full_name.value
        else:
            self._displayname = self._firstname + ' ' + self._lastname

        ## yet to support following which are really multi-valued properties
        ## - Companies
        ## - PhysicalAddresses
        ## - ImAddresses

        ## yet to support following which are not returned normally and have
        ## to be dealt with as extended properties.
        ## - gender
        ## - LastModifiedTime

    def add_extended_property (self, node):
        uri = node.find(QName_T('ExtendedFieldURI'))
        if uri is None:
            logging.error('ExtendedProperty.init_from_xml(): no child node ' +
                          'ExtendedFieldURI in node: %s',
                          pretty_xml(node))
            return

        ## Look for known extended properties
        is_tp, tag = ExtendedProperty.is_tagged_prop(uri)
        if is_tp:
            self.add_tagged_property(tag, node)
        else:
            self.eprops.append(ExtendedProperty(node=node))

    def add_tagged_property (self, tag, node):
        eprop = None
        if tag == mapitags.PR_LAST_MODIFICATION_TIME:
            self.last_modified_time = LastModificationTime(node=node)
            eprop = self.last_modified_time
        elif tag == mapitags.PR_GENDER:
            self.gender = Gender(node=node)
            eprop = self.gender
        else:
            eprop = ExtendedProperty(node=node)
            self.eprops.append(eprop)
            self.eprops_tagged[tag] = eprop

    ## Override Field.get_chidren to refresh the children array every time
    def get_children (self):
        cn = self.complete_name
        ## Note that children is used for generating xml representation of
        ## this contact for CreateItem and update operations. The order of
        ## these fields is critical. I know, it's crazy.
        self.children = [self.notes] + self.eprops + [self.gender, self.file_as,
                         self.display_name, cn.given_name, cn.initials,
                         cn.middle_name, cn.nickname, self.company_name,
                         self.emails, self.phones, self.assistant_name,
                         self.birthday, self.department, self.job_title,
                         self.manager, self.spouse_name, cn.surname,
                         self.anniversary, self.alias, self.notes]

        return self.children

    def save (self):
        self.service.CreateItems(self.ParentFolderId, [self])

    def __str__ (self):
        lmt_tag = mapitags.PR_LAST_MODIFICATION_TIME

        s  = 'ItemId: %s' % self.itemid
        s += '\nCreated: %s' % self.created_time
        if lmt_tag in self.eprops_tagged:
            s += '\nLast Modified: %s' % self.eprops_tagged[lmt_tag].value.value
        s += '\nName: %s' % self._displayname
        s += '\nGEnder: %s' % self.gender
        s += '\nPhones: %s' % self.phones
        s += '\nEmails: %s' % self.emails
        s += '\nNotes: %s' % self.notes

        return s

## The XML Schema for a EWS Contact, taken from:
## http://msdn.microsoft.com/en-us/library/office/aa581315(v=exchg.150).aspx
##
## <Contact>
##    <MimeContent/>
##    <ItemId/>
##    <ParentFolderId/>
##    <ItemClass/>
##    <Subject/>
##    <Sensitivity/>
##    <Body/>
##    <Attachments/>
##    <DateTimeReceived/>
##    <Size/>
##    <Categories/>
##    <Importance/>
##    <InReplyTo/>
##    <IsSubmitted/>
##    <IsDraft/>
##    <IsFromMe/>
##    <IsResend/>
##    <IsUnmodified/>
##    <InternetMessageHeaders/>
##    <DateTimeSent/>
##    <DateTimeCreated/>
##    <ResponseObjects/>
##    <ReminderDueBy/>
##    <ReminderIsSet/>
##    <ReminderMinutesBeforeStart/>
##    <DisplayCc/>
##    <DisplayTo/>
##    <HasAttachments/>
##    <ExtendedProperty/>
##    <Culture/>
##    <EffectiveRights/>
##    <LastModifiedName/>
##    <LastModifiedTime/>
##    <IsAssociated/>
##    <WebClientReadFormQueryString/>
##    <WebClientEditFormQueryString/>
##    <ConversationId/>
##    <UniqueBody/>
##    <FileAs/>
##    <FileAsMapping/>
##    <DisplayName/>
##    <GivenName/>
##    <Initials/>
##    <MiddleName/>
##    <Nickname/>
##    <CompleteName>
##       <Title/>
##       <FirstName/>
##       <MiddleName/>
##       <LastName/>
##       <Suffix/>
##       <Initials/>
##       <FullName/>
##       <Nickname/>
##       <YomiFirstName/>
##       <YomiLastName/>
##    </CompleteName>
##    <CompanyName/>
##    <EmailAddresses/>
##    <PhysicalAddresses/>
##    <PhoneNumbers/>
##    <AssistantName/>
##    <Birthday/>
##    <BusinessHomePage/>
##    <Children/>
##    <Companies/>
##    <ContactSource/>
##    <Department/>
##    <Generation/>
##    <ImAddresses/>
##    <JobTitle/>
##    <Manager/>
##    <Mileage/>
##    <OfficeLocation/>
##    <PostalAddressIndex/>
##    <Profession/>
##    <SpouseName/>
##    <Surname/>
##    <WeddingAnniversary/>
##    <HasPicture/>
##    <PhoneticFullName/>
##    <PhoneticFirstName/>
##    <PhoneticLastName/>
##    <Alias/>
##    <Notes/>
##    <Photo/>
##    <UserSMIMECertificate/>
##    <MSExchangeCertificate/>
##    <DirectoryId/>
##    <ManagerMailbox/>
##    <DirectReports/>
## </Contact>
