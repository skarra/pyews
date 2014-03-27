
import os, urllib2

REQUESTS_DIR = "requests"

def template_fn (fn):
    return os.path.abspath(os.path.join(REQUESTS_DIR, fn))

REQ_AUTODIS_EPS = template_fn("autodiscover_endpoints.mustache")
REQ_GET_FOLDER  = template_fn("get_folder.mustache")
