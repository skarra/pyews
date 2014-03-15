#!/usr/bin/python
##
## Created : Wed Mar 05 11:28:41 IST 2014
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

import logging
import utils
from   autodiscover import EWSAutoDiscover, ExchangeAutoDiscoverError

#from   httplib import HTTPException
#from   ntlm    import HTTPNtlmAuthHandler

USER = u'skarra@asynk.onmicrosoft.com'
PWD  = u'tsYWpw8m'

class InvalidUserEmail(Exception):
    pass

class EWS:
    def __init__ (self, user, pwd):
        self.user = user
        self.pwd  = pwd
        self.ews_ad = EWSAutoDiscover(user, pwd)
        self.ews_url = self.ews_ad.discover()

def main ():
    logging.getLogger().setLevel(logging.DEBUG)
    ews = EWS(USER, PWD)

if __name__ == "__main__":
    main()
