
import logging, sys

sys.path.append("../")
sys.path.append("../pyews")

from   pyews.pyews            import WebCredentials, ExchangeService
from   pyews.ews.autodiscover import EWSAutoDiscover, ExchangeAutoDiscoverError
from   pyews.ews.data         import DistinguishedFolderId, WellKnownFolderName
from   pyews.ews.data         import FolderClass
from   pyews                  import utils
from   pyews.ews.folder       import Folder

def main ():
    logging.getLogger().setLevel(logging.DEBUG)

    global USER, PWD, EWS_URL, ews

    with open('auth.txt', 'r') as inf:
        USER    = inf.readline().strip()
        PWD     = inf.readline().strip()
        EWS_URL = inf.readline().strip()

        logging.debug('Username: %s; Url: %s', USER, EWS_URL)

    creds = WebCredentials(USER, PWD)
    ews = ExchangeService()
    ews.credentials = creds

    try:
        ews.AutoDiscoverUrl()
    except ExchangeAutoDiscoverError as e:
        logging.info('ExchangeAutoDiscoverError: %s', e)
        logging.info('Falling back on manual url setting.')
        ews.Url = EWS_URL

    ews.init_soap_client()

    root = bind()
    # test_create_folder(root)
    cons = test_list_items(root)
    test_find_item(cons[0].itemid.text)

def bind ():
    return Folder.bind(ews, WellKnownFolderName.MsgFolderRoot)    

def test_fetch_contact_folder ():
    contacts = root.fetch_all_folders(types=FolderClass.Contacts)
    for f in contacts:
        print 'DisplayName: %s; Id: %s' % (f.DisplayName, f.Id)    

def test_create_folder (parent):
    resp = ews.CreateFolder(parent.Id,
                            [('Test Contacts', FolderClass.Contacts)])
    print utils.pretty_xml(resp)

def test_list_items (root):
    cfs = root.FindFolders(types=FolderClass.Contacts)
    contacts = ews.FindItems(cfs[0])
    for con in contacts:
        n = con.display_name.text
        print 'Name: %-10s; itemid: %s' % (n, con.itemid)

    return contacts

def test_find_item (itemid):
    cons = ews.GetItems([itemid])
    if cons is None or len(cons) <= 0:
        print 'WTF. Could not find itemid ', itemid
    else:
        print cons[0].to_xml()

if __name__ == "__main__":
    main()
