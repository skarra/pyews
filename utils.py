
import os, urllib2

REQUESTS_DIR = "requests"

def template_fn (fn):
    return os.path.abspath(os.path.join(REQUESTS_DIR, fn))

AUTODIS_EPS_REQUEST_FN = template_fn("autodiscover_endpoints.mustache")
