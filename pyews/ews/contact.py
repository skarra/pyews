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

from item    import Item, Field
from pyews.soap    import SoapClient, unQName
from pyews.utils   import pretty_xml

gnd = SoapClient.get_node_detail

class FileAs(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'FileAs'

class Alias(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'Alias'

class Title(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'Title'

class FirstName(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'FirstName'

class GivenName(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'GivenName'

class MiddleName(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'MiddleName'

class LastName(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'LastName'

class Surname(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'Surname'

class Suffix(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'Suffix'

class Initials(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'Initials'

class Nickname(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'Nickname'

class DisplayName(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'DisplayName'

class FullName(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'FullName'

class SpouseName(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'FirstName'

class JobTitle(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'JobTitle'

class CompanyName(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'CompanyName'

class Department(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'Department'

class Manager(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'Manager'

class AssistantName(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'AssistantName'

## FIXME: The values for these are strings, but might be better represented as
## DateTime objects... Hm
class Birthday(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'Birthday'

class WeddingAnniversary(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'WeddingAnniversary'

class Notes(Field):
    def __init__ (self, text=None):
        Field.__init__(self, text)
        self.tag = 'Body'

class EmailAddresses(Field):
    class Email(Field):
        def __init__ (self, text=None):
            self.tag = 'Entry'
            self.attrib = {
                'Key' : None,
                'Name' : None,
                'RoutingType' : None,
                'MailBoxType' : None
                }

        def __str__ (self):
            return 'Key: %8s  Address: %s' % (self.attrib['Key'], self.text)

        def __repr__ (self):
            return self.__str__()

    def __init__ (self, node=None):
        self.tag = 'EmailAddresses'
        self.entries = []
        if node is not None:
            self.populate_from_node(node)

    def populate_from_node (self, node):
        for child in node:
            email = self.Email()
            for k, v in child.attrib.iteritems():
                email.add_attrib(k, v)

            email.text = child.text
            self.entries.append(email)

    def __str__ (self):
        s = '%s Numbers: ' % len(self.entries)
        s += '; '.join([str(x) for x in self.entries])
        return s

    def __repr__ (self):
        return self.__str__()

class PhoneNumbers(Field):
    class Phone(Field):
        def __init__ (self):
            self.tag = 'Entry'
            self.attrib = {
                'Key' : None,               # ews.data.PhoneKey
                }

        def __str__ (self):
            return 'Key: %8s  Number: %s' % (self.attrib['Key'], self.text)

        def __repr__ (self):
            return self.__str__()

    def __init__ (self, node=None):
        self.tag = 'PhoneNumbers'
        self.entries = []

        if node is not None:
            self.populate_from_node(node)

    def populate_from_node (self, node):
        for child in node:
            phone = self.Phone()
            for k, v in child.attrib.iteritems():
                phone.add_attrib(k, v)

            phone.text = child.text
            self.entries.append(phone)

    def to_xml (self):
        return None

    def __str__ (self):
        s = '%s Numbers: ' % len(self.entries)
        s += '; '.join([str(x) for x in self.entries])
        return s

    def __repr__ (self):
        return self.__str__()

class ExtendedProperty:
    def __init__ (self):
        ## if self.dis_guid is used guid and prop_tag cannot be used
        self.dis_guid = None

        self.guid = None
        self.prog_tag = None
        self.prop_id = None
        self.prop_type = None

        self.prop_name = None
        self.prop_value = None

    def init_from_xml (self, node):
        """
        None is a parsed xml node (Element). Extract the data that we can
        from the node.
        """

        uri = node.find(QName_T('ExtendedFieldURI'))
        if uri is None:
            logging.debug('ExtendedProperty.init_from_xml(): no child node ' +
                          'ExtendedFieldURI in node: %s', pretty_xml(node))

        self.dis_guid = uri.attrib['DistinguishedPropertySetId']
        self.guid = uri.attrib['PropertySetId']
        self.prop_tag = uri.attrib['PropertyTag']
        self.prop_id = uri.attrib['PropertyId']
        self.prop_name = uri.attrib['PropertyName']
        self.prop_type = uri.attrib['PropertyString']

        self.prop_value = node.find(QName_T('Value')).text

class Contact(Item):
    """
    Abstract wrapper class around an Exchange Item object. Frequently an
    object of this type is instantiated from a response.
    """

    def __init__ (self, service, parent_fid=None, resp_node=None):
        Item.__init__(self, service, parent_fid, resp_node)

        self.file_as      = FileAs()
        self.alias        = Alias()
        self.title        = Title()
        self.given_name   = GivenName()
        self.first_name   = FirstName()
        self.middle_name  = MiddleName()
        self.last_name    = LastName()
        self.surname      = Surname()
        self.suffix       = Suffix()
        self.initials     = Initials()
        self.nickname     = Nickname()
        self.display_name = DisplayName()
        self.full_name    = FullName()

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

        self._init_from_resp()

    def _init_from_resp (self):
        if self.resp_node is None:
            return

        rnode = self.resp_node
        for child in rnode:
            tag = unQName(child.tag)

            if tag == 'FileAs':
                self.file_as.text = child.text
            elif tag == 'Alias':
                self.alias.text = child.text
            elif tag == 'SpouseName':
                self.spouse_nam.text = child.text
            elif tag == 'JobTitle':
                self.job_title.text = child.text
            elif tag == 'CompanyName':
                self.company_name.text = child.text
            elif tag == 'Department':
                self.department.text = child.text
            elif tag == 'Manager':
                self.manager.text = child.text
            elif tag == 'AssistantName':
                self.assistant_name.text = child.text
            elif tag == 'Birthday':
                self.birthday.text = child.text
            elif tag == 'WeddingAnniversary':
                self.anniversary.text = child.text
            elif tag == 'GivenName':
                self.given_name.text = child.text
            elif tag == 'Surname':
                self.surname.text = child.text
            elif tag == 'DisplayName':
                self.display_name.text = child.text
            elif tag == 'Body':
                ## FIXME: We are assuming a text body type, but they could
                ## contain html or other types as well... Oh, well.
                self.notes.text = child.text
            elif tag == 'EmailAddresses':
                self.emails.populate_from_node(child)
            elif tag == 'PhoneNumbers':
                self.phones.populate_from_node(child)

        n = rnode.find('CompleteName')
        if n is not None:
            print 'Yo, ma =========================='
            rnode = n

        self.title.text = self._find_text_safely(rnode, 'Title')
        self.first_name.text = self._find_text_safely(rnode, 'FirstName')
        self.middle_name.text = self._find_text_safely(rnode, 'MiddleName')
        self.last_name.text = self._find_text_safely(rnode, 'LastName')
        self.suffix.text = self._find_text_safely(rnode, 'Suffix')
        self.initials.text = self._find_text_safely(rnode, 'Initials')
        self.full_name.text = self._find_text_safely(rnode, 'FullName')
        self.nickname.text = self._find_text_safely(rnode, 'Nickname')

        ## It's a bit hard to understand why the hell they have so many
        ## variants for the same stupid information... Oh well, let's just
        ## have a few handy shortcuts for the information that matters
        if self.first_name.text:
            self._firstname = self.first_name.text
        else:
            self._firstname = self.given_name.text

        if self.last_name.text:
            self._lastname = self.last_name.text
        else:
            self._lastname = self.surname.text

        if self.display_name.text:
            self._displayname = self.display_name.text
        elif self.full_name.text:
            self._displayname = self.full_name.text
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

    def __str__ (self):
        s  = 'ItemId: %s' % self.itemid
        s += '\nCreated: %s' % self.created_time
        s += '\nName: %s' % self._displayname
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
