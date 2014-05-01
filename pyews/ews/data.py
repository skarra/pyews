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
## In this file we define some standard names, constants and values that will
## be recognizabe for EWS / EWS Managed API users
##

import mapitags
from   pyews.soap import SoapClient, SoapMessageError, QName_M

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

class GenderType:
    Unspecified = 0x0001
    Female      = 0x0002
    Male        = 0x0003

MapiPropertyTypeType = {
    mapitags.PT_UNSPECIFIED : "Unspecified",
    mapitags.PT_NULL        : "Null",
    mapitags.PT_I2          : "Short",
    mapitags.PT_LONG        : "Integer",
    mapitags.PT_R4          : "Float",
    mapitags.PT_DOUBLE      : "Double",
    mapitags.PT_CURRENCY    : "Currency",
    mapitags.PT_APPTIME     : "ApplicationTime",
    mapitags.PT_ERROR       : "Error",
    mapitags.PT_BOOLEAN     : "Boolean",
    mapitags.PT_OBJECT      : "Object",
    mapitags.PT_I8          : "Long",
    mapitags.PT_STRING8     : "String",   # Hack. This really is 8-bit single char
    mapitags.PT_UNICODE     : "String",
    mapitags.PT_SYSTIME     : "SystemTime",
    mapitags.PT_CLSID       : "CLSID",
    mapitags.PT_BINARY      : "Binary",

    ## Need to support the Array types
    }

MapiPropertyTypeTypeInv = {}
for k, v in MapiPropertyTypeType.iteritems():
    MapiPropertyTypeTypeInv[v] = k

class ConflictResoltion:
    NeverOverwrite  = 'NeverOverwrite'
    AutoResolve     = 'AutoResolve'
    AlwaysOverwrite = 'AlwaysOverwrite'

def ews_pt (tag):
    """Return the EWS Property Type enumeration for the specific MAPI Property
    Tag. """

    return MapiPropertyTypeType[mapitags.PROP_TYPE(tag)]

def ews_pid (tag):
    """Return the EWS Property ID for the specific MAPI Property Tag. """

    return mapitags.PROP_ID(tag)
