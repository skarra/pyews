#!/usr/bin/python
##
## Created : Sat Mar 15 17:30:44 IST 2014
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

import logging, re, urllib2

class ExchangeAutoDiscoverError(Exception):
    pass

class EWSAutoDiscover:
    def __init__ (self, user, pwd):
        self.user = user
        self.pwd  = pwd
        self.url  = ""

        res = re.match("(.*)@(.*)$", self.user)
        if not res:
            raise InvalidUserEmail("Could not get domain from user email id")
        self.domain = res.group(2)

    def discover (self):
        """Based on a username and password try to autodiscover the EWS
        endpoint by following the steps listed here:

        Doc 1: http://msdn.microsoft.com/en-us/library/ee332364.aspx
        Doc 2: http://msdn.microsoft.com/en-us/library/office/jj900169(v=exchg.150).aspx

        In case there is a failure in performing the autodiscovery for any
        reason an exception ExchangeAutoDiscoverError is raised."""

        ## FIXME: Till we implement something ...
        raise ExchangeAutoDiscoverError('Not Implemented')

        self.url = ""

        ## Step 1 in Doc 1 above
        logging.debug("Trying Autodiscover through SCP...")
        ret = self.discover_through_scp()
        if ret:
            return ret

        ## Steps 2, 3 in Doc 1 above
        logging.debug("Trying Autodiscover through email domain...")
        ret = self.discover_through_email_domain()
        if ret:
            return ret

        ## Step 4 in Doc 1 above
        logging.debug("Trying Autodiscover through unauth get...")
        ret = self.discover_through_unauth_get()
        if ret:
            return ret

        # ep_req = self.pysren.render_path(utils.REQ_AUTODIS_EPS,
        #                                  {'mailbox' : self.user})
        # print ep_req

    def discover_through_scp (self):
        """Hard to test this as this requires the client to be on a computer
        that is attached to the domain"""

        return None

    def discover_through_email_domain (self):
        endpoints = self.email_domain_endpoints()

        ## There is no way for us to test this at the moment. So we will wing
        ## it... FIXME. For now we could at least try to post something and
        ## see if we get a response back...
        return None

    def email_domain_endpoints (self):
        servers = ["/".join(["https:/", self.domain, "autodiscover",
                             "autodiscover.svc"]),
                   "/".join(["https:/", "autodiscover.%s" % self.domain,
                             "autodiscover", "autodiscover.svc"])]

        logging.debug(servers)
        return servers

    def discover_through_unauth_get (self):
        top_level_url = "http://autodiscover.%s" % self.domain
        url = "/".join([top_level_url, "autodiscover", "autodiscover.xml"])

        logging.debug('  trying url: %s', url)

        password_mgr = urllib2.HTTPPasswordMgrWithDefaultRealm()
        password_mgr.add_password(None, top_level_url, self.user, self.pwd)
        handler = urllib2.HTTPBasicAuthHandler(password_mgr)

        # Open the URL and hope for the best.
        try:
            urllib2.build_opener(handler).open(url)
        except IOError, e:
            if hasattr(e, 'code'):
                if e.code != 401:
                    print 'We got another error'
                    print e.code
                else:
                    print
                    print '** Headers: **'
                    print e.headers

    class HTTPRedirectHandlerNo302(urllib2.HTTPRedirectHandler):
        def http_error_302 (self, req, fp, code, msg, headers):
            print '** Inside 302 handler **'
            print headers
            return urllib2.HTTPRedirectHandler.http_error_302(self, req, fp,
                                                              code, msg, headers)
