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
from soap    import SoapClient, unQName

gnd = SoapClient.get_node_detail

class Contact(Item):
    """
    Abstract wrapper class around an Exchange Item object. Frequently an
    object of this type is instantiated from a response.
    """

    def __init__ (self, service, parent=None, resp_node=None):
        Item.__init__(self, service, parent, resp_node)

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

        n = rnode.find('CompleteName')
        if n is not None:
            rnode = n

        self.Title = self.._find_text_safely(rnode, 'Title')
        self.FirstName = self._find_text_safely(rnode, 'FirstName')
        self.MiddleName = self._find_text_safely(rnode, 'MiddleName')
        self.LastName = self._find_text_safely(rnode, 'LastName')
        self.Suffix = self.._find_text_safely(rnode, 'Suffix')
        self.Initials = self.._find_text_safely(rnode, 'Initials')
        self.FullName = self._find_text_safely(rnode, 'FullName')
        self.Nickname = self._find_text_safely(rnode, 'Nickname')

        ## yet to support following which are really multi-valued properties
        ## - Companies
        ## - EmailAddresses
        ## - PhysicalAddresses
        ## - PhoneNumbers
        ## - ImAddresses

    def __str__ (self):
        s  = 'ItemId: %s' % self.ItemId
        s += '\nName: %s' % self.GivenName
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
