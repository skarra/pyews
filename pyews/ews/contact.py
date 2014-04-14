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

from abc     import ABCMeta, abstractmethod
from item    import Item
from pyews.soap    import SoapClient, unQName
from pyews.utils   import pretty_xml

gnd = SoapClient.get_node_detail

class EmailAddresses:
    class Email:
        def __init__ (self):
            self.Key = None
            self.Name = None
            self.RoutingType = None
            self.MailboxType = None
            self.Address = None

        def __str__ (self):
            return 'Key: %8s  Address: %s' % (self.Key, self.Address)

        def __repr__ (self):
            return self.__str__()

    def __init__ (self, node):
        self.entries = []

        for child in node:
            email = self.Email()
            if 'Key' in child.attrib:
                email.Key = child.attrib['Key']
            if 'RoutingType' in child.attrib:
                email.RoutingType = child.attrib['RoutingType']
            if 'MailboxType' in child.attrib:
                email.MailboxType = child.attrib['MailboxType']

            email.Address = child.text
            self.entries.append(email)

    def __str__ (self):
        s = '%s Numbers: ' % len(self.entries)
        s += '; '.join([str(x) for x in self.entries])
        return s

    def __repr__ (self):
        return self.__str__()

class PhoneNumbers:
    class Phone:
        def __init__ (self):
            self.Key = None               # ews.data.PhoneKey
            self.Number = None

        def __str__ (self):
            return 'Key: %8s  Number: %s' % (self.Key, self.Number)

        def __repr__ (self):
            return self.__str__()

    def __init__ (self, node):
        self.entries = []

        for child in node:
            phone = self.Phone()
            if 'Key' in child.attrib:
                phone.Key = child.attrib['Key']

            phone.Number = child.text
            self.entries.append(phone)

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

    def __init__ (self, service, parent=None, resp_node=None):
        Item.__init__(self, service, parent, resp_node)

        self.FileAs = None
        self.Alias = None
        self.SpouseName = None
        self.JobTitle = None
        self.CompanyName = None
        self.Department = None
        self.Manager = None
        self.AssistantName = None
        self.Birthday = None
        self.WeddingAnniversary = None
        self.GivenName = None
        self.Surname = None
        self.DisplayName = None
        self.Emails = None
        self.Title = None
        self.FirstName = None
        self.MiddleName = None
        self.LastName = None
        self.Suffix = None
        self.Initials = None
        self.FullName = None
        self.Nickname = None
        self.Notes = None
        self.Emails = None
        self.Phones = None

        self._init_from_resp()

    def _init_from_resp (self):
        if self.resp_node is None:
            return

        rnode = self.resp_node
        for child in rnode:
            tag = unQName(child.tag)

            if tag == 'FileAs':
                self.FileAs = child.text
            elif tag == 'Alias':
                self.Alias = child.text
            elif tag == 'SpouseName':
                self.SpouseName = child.text
            elif tag == 'JobTitle':
                self.JobTitle = child.text
            elif tag == 'CompanyName':
                self.CompanyName = child.text
            elif tag == 'Department':
                self.Department = child.text
            elif tag == 'Manager':
                self.Manager = child.text
            elif tag == 'AssistantName':
                self.AssistantName = child.text
            elif tag == 'Birthday':
                self.Birthday = child.text
            elif tag == 'WeddingAnniversary':
                self.WeddingAnniversary = child.text
            elif tag == 'GivenName':
                self.GivenName = child.text
            elif tag == 'Surname':
                self.Surname = child.text
            elif tag == 'DisplayName':
                self.DisplayName = child.text
            elif tag == 'Body':
                ## FIXME: We are assuming a text body type, but they could
                ## contain html or other types as well... Oh, well.
                self.Notes = child.text
            elif tag == 'EmailAddresses':
                self.Emails = EmailAddresses(child)
            elif tag == 'PhoneNumbers':
                self.Phones = PhoneNumbers(child)

        n = rnode.find('CompleteName')
        if n is not None:
            rnode = n

        self.Title = self._find_text_safely(rnode, 'Title')
        self.FirstName = self._find_text_safely(rnode, 'FirstName')
        self.MiddleName = self._find_text_safely(rnode, 'MiddleName')
        self.LastName = self._find_text_safely(rnode, 'LastName')
        self.Suffix = self._find_text_safely(rnode, 'Suffix')
        self.Initials = self._find_text_safely(rnode, 'Initials')
        self.FullName = self._find_text_safely(rnode, 'FullName')
        self.Nickname = self._find_text_safely(rnode, 'Nickname')

        ## It's a bit hard to understand why the hell they have so many
        ## variants for the same stupid information... Oh well, let's just
        ## have a few handy shortcuts for the information that matters
        self._firstname = self.FirstName if self.FirstName else self.GivenName
        self._lastname = self.LastName if self.LastName else self.Surname
        if self.DisplayName:
            self._displayname = self.DisplayName
        elif self.FullName:
            self._displayname = self.FullName
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
        s  = 'ItemId: %s' % self.ItemId
        s += '\nCreated: %s' % self.DateTimeCreated
        s += '\nName: %s' % self._displayname
        s += '\nPhones: %s' % self.Phones
        s += '\nEmails: %s' % self.Emails
        s += '\nNotes: %s' % self.Notes

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
