
import md5, os, re, urllib2, xml.dom.minidom

CUR_DIR = os.path.dirname(os.path.realpath(__file__))
REQUESTS_DIR = os.path.abspath(os.path.join(CUR_DIR, "templates"))

def template_fn (fn):
    # return os.path.abspath(os.path.join(REQUESTS_DIR, fn))
    return fn

REQ_AUTODIS_EPS = template_fn("autodiscover_endpoints.xml")
REQ_GET_FOLDER  = template_fn("get_folder.xml")
REQ_CREATE_FOLDER = template_fn("create_folder.xml")
REQ_DELETE_FOLDER = template_fn("delete_folder.xml")
REQ_BIND_FOLDER = template_fn("bind.xml")
REQ_FIND_FOLDER_ID = template_fn("find_folder_id.xml")
REQ_FIND_FOLDER_DI = template_fn("find_folder_distinguishd.xml")
REQ_FIND_ITEM = template_fn("find_item.xml")
REQ_FIND_ITEM_LMT = template_fn("find_item_lmt.xml")
REQ_GET_ITEM = template_fn("get_item.xml")
REQ_CREATE_ITEM = template_fn("create_item.xml")
REQ_DELETE_ITEM = template_fn("delete_item.xml")
REQ_UPDATE_ITEM = template_fn("update_item.xml")
REQ_SYNC_FOLDER = template_fn("sync_folder.xml")

def pretty_xml (x):
    x = xml.dom.minidom.parseString(x).toprettyxml()
    lines = x.splitlines()
    lines = [s for s in lines if not re.match(s.strip(), '^\s*$')]
    return os.linesep.join(lines)

def pretty_eid (x):
    """
    Returns a 32-digit md5 digest of the input. This is useful for printing
    really long and butt-ugly Base64-encoded Entry IDs used by
    Exchange/Outlook
    """

    return md5.new(x).hexdigest()

def safe_int (s):
    """Convert string s into an integer taking into account if s is a hex
    representaiton with a leading 0x."""

    if s[0:2] == '0x':
        return int(s, 16)
    elif s[0:1] == '0':
        return int(s, 8)
    else:
        return int(s, 10)
