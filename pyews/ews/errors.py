##
## Created : Thu May 01 10:51:44 IST 2014
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
##s License for more details.
##
## You should have a copy of the license in the doc/ directory of pyews.  If
## not, see <http://www.gnu.org/licenses/>.

from   pyews.soap import SoapClient, SoapMessageError, QName_M

class EWSMessageError(Exception):
    def __init__ (self, resp_obj):
        self.resp_obj = resp_obj

    def __str__ (self):
        return '%s - %s' % (self.resp_obj.fault_code, self.resp_obj.fault_str)

class EWSResponseError(Exception):
    def __init__ (self, resp_obj):
        self.resp_obj = resp_obj

    def __str__ (self):
        s = 'Errors: '
        for i, err in self.resp_obj.errors.iteritems():
            s += '\n  %02d - %s' % (i, str(err))

        return s

class EWSCreateFolderError(Exception):
    pass

class EWSDeleteFolderError(Exception):
    pass
