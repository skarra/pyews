##
## Created : Fri Mar 28 22:47:40 IST 2014
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
##s License for more details.
##
## You should have a copy of the license in the doc/ directory of pyews.  If
## not, see <http://www.gnu.org/licenses/>.

##
#i# In this file we define some standard names, constants and values that will
 ## be recognizabe for EWS / EWS Managed API users
##

## perhaps this is not required?
class WellKnownFolderName:
    MsgFolderRoot = 'msgfolderroot'

class DistinguishedFolderId:
    """
    This list is largely taken from:
    http://msdn.microsoft.com/en-us/library/office/aa580808(v=exchg.150).aspx
    """
    calendar = 'calendar'
    contacts = 'contacts'
    deletedItems = 'deleteditems'
    drafts = 'drafts'
    inbox = 'inbox'
    journal = 'journal'
    notes = 'notes'
    outbox = 'outbox'
    sentItems = 'sentitems'
    tasks = 'tasks'
    msgFolderRoot = 'msgfolderroot'
    root = 'root'
    junkEmail = 'junkemail'
    searchFolders = 'searchfolders'
    voiceMail = 'voicemail'

    ##
    ## skipping some not-so-useful stuff
    ##

    myContacts = 'mycontacts'
    directory = 'directory'
    imContactList = 'imcontactlist'
    peopleConnect = 'peopleconnect'

class FolderClass:
    Contacts = 'IPF.Contact'
    Journals = 'IPF.Journal'
    Tasks    = 'IPF.Task'
    Calendars = 'IPF.Calendar'
    Notes = 'IPF.Note'

class ItemClass:
    Activity = 'IPM.Activity'
    Appointment = 'IPM.Appointment'
    Contact = 'IPM.Contact'
    DistList = 'IPM.DistList'
    Note = 'IPM.Note'
    Task = 'IPM.Task'


class PhoneKey:
    AssistantPhone   = 'AssistantPhone'
    BusinessFax      = 'BusinessFax'
    BusinessPhone    = 'BusinessPhone'
    BusinessPhone2   = 'BusinessPhone2'
    Callback         = 'Callback'
    CarPhone         = 'CarPhone'
    CompanyMainPhone = 'CompanyMainPhone'
    HomeFax          = 'HomeFax'
    HomePhone        = 'HomePhone'
    HomePhone2       = 'HomePhone2'
    Isdn             = 'Isdn'
    MobilePhone      = 'MobilePhone'
    OtherFax         = 'OtherFax'
    OtherTelephone   = 'OtherTelephone'
    Pager            = 'Pager'
    PrimaryPhone     = 'PrimaryPhone'
    RadioPhone       = 'RadioPhone'
    Telex            = 'Telex'
    TtyTddPhone      = 'TtyTddPhone'

class EWSMessageError(Exception):
    pass

class EWSCreateFolderError(Exception):
    pass

class EWSDeleteFolderError(Exception):
    pass
