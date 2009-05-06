# -*- coding: utf-8 -*-
# File: footer.py
#
# Copyright (c) 2006 by ['Jean-Michel FRANCOIS']
#
# GNU General Public License (GPL)
#
# This program is free software; you can redistribute it and/or
# modify it under the terms of the GNU General Public License
# as published by the Free Software Foundation; either version 2
# of the License, or (at your option) any later version.

# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.

# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.

"""For merging purpose with standard.py"""

from standard import *
import logging
logger = logging.getLogger('ooo2tools.core.footer')
parser = OptionParser()
parser.add_option("", "--port",
                type="int", dest="port", default=2002,
                help="specify the port")
parser.add_option("", "--sequence",
                type="string", dest="seq", default=2002,
                help="specify the file that contains command sequences")

(options, args) = parser.parse_args()

def execute_cmd(ooo_info,cmd_args,b_toc=False):
    """Do it"""
    cmd = 'cmd_'
    cmd += cmd_args.split(':')[0]
    args = cmd_args.split(':')[1]
    logger.info("DBUG: cmd %s"%(cmd))
    if os.path.isfile(args):
        logger.info("execute cmd %s %s file in argument"%(cmd,args))
    if not ooo_info['globals'].has_key(cmd):
        raise KeyError("Can't find cmd %s %s"%(cmd,args))
    ooo_info['globals'].get(cmd)(ooo_info,args)

list_cmd = parse_seq(options.seq)
ooo_info = ooo_connect()
for cmd in list_cmd:
    execute_cmd(ooo_info,cmd)
if ooo_info:
    ooo_info['doc'].close(False)
else:
    raise Exception("Can't connect to openoffice.org server")
