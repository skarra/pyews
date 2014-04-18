
import os, urllib2, xml.dom.minidom

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
REQ_GET_ITEM = template_fn("get_item.xml")
REQ_CREATE_ITEM = template_fn("create_item.xml")

def pretty_xml (x):
    return xml.dom.minidom.parseString(x).toprettyxml()

def clean_xml (x):
    return x
    strs = x.split('\n')
    resp = []
    for line in strs:
        if not line.isspace() and line != '':
            print "Processing '%s'" % line
            resp.append(line)

    resp = ''.join(resp)
    return pretty_xml(resp)
