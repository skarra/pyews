
import os, urllib2

REQUESTS_DIR = "templates"

def template_fn (fn):
    # return os.path.abspath(os.path.join(REQUESTS_DIR, fn))
    return fn

REQ_AUTODIS_EPS = template_fn("autodiscover_endpoints.xml")
REQ_GET_FOLDER  = template_fn("get_folder.xml")
REQ_CREATE_FOLDER = template_fn("create_folder.xml")
