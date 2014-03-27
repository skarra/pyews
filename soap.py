##
## Created : Sun Mar 16 21:07:35 IST 2014
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

import logging, requests
from requests.auth import HTTPBasicAuth

class SoapClient(object):
    def __init__ (self, service_url, user, pwd):
        self.url = service_url
        self.user = user
        self.pwd = pwd
        self.auth = HTTPBasicAuth(user, pwd)

    def send (self, request):
        r = requests.post(self.url, auth=self.auth, data=request,
                          headers={'Content-Type':'text/xml; charset=utf-8',
                                   "Accept": "text/xml"})

        return r.text
