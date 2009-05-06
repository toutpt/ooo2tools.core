# -*- coding: utf-8 -*-
#
# File: able_toconnect.py
# This script check the Python (module UNO) 
# and the connection to OpenOffice.org2
#
# Copyright (c) 2006 by ['Jean-Michel FRANCOIS']
#
# GNU General Public License (GPL)
#
# This program is free software; you can redistribute it and/or
# modify it under the terms of the GNU General Public License
# as published by the Free Software Foundation; either version 2
# of the License, or (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
# 02110-1301, USA.
#

from optparse import OptionParser
parser = OptionParser()
parser.add_option("", "--port",
                type="int", dest="port", default=2002,
                help="specify the port")
(options, args) = parser.parse_args()
port = options.port
print "port: %s"%port
try:
    import uno
    print "import uno: OK"
except:
    raise "import uno: ERROR"
try:
    localContext = uno.getComponentContext()
    print "local context: OK"
except:
    raise "local context: ERROR"
try:
    resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext )
    print "resolver: OK"
except:
    raise "resolver: ERROR"
try:
    ctx = resolver.resolve("uno:socket,host=localhost,port=%s;urp;StarOffice.ComponentContext"%port)
    print "connexion to OpenOffice.org2: OK"
except:
    raise "connexion to OpenOffice.org2: ERROR"