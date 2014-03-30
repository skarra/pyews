
import os, urllib2, xml.dom.minidom

CUR_DIR = os.path.dirname(os.path.realpath(__file__))
REQUESTS_DIR = os.path.abspath(os.path.join(CUR_DIR, "templates"))

def template_fn (fn):
    # return os.path.abspath(os.path.join(REQUESTS_DIR, fn))
    return fn

REQ_AUTODIS_EPS = template_fn("autodiscover_endpoints.xml")
REQ_GET_FOLDER  = template_fn("get_folder.xml")
REQ_CREATE_FOLDER = template_fn("create_folder.xml")
REQ_BIND_FOLDER = template_fn("bind.xml")
REQ_FIND_FOLDER_ID = template_fn("find_folder_id.xml")
REQ_FIND_FOLDER_DI = template_fn("find_folder_distinguishd.xml")

def pretty_xml (x):
    return xml.dom.minidom.parseString(x).toprettyxml()
