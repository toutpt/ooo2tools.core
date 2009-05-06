# -*- coding: utf-8 -*-
# File: standard.py
# Give common OpenOffice.org task under simple command
# 
# Copyright (c) 2006 by ['Jean-Michel FRANCOIS']
# Generator: ArchGenXML Version 1.5.0
#            http://plone.org/products/archgenxml
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

import os
from optparse import OptionParser


#UNO import
import uno
from unohelper import systemPathToFileUrl, absolutize
from com.sun.star.beans import PropertyValue
from com.sun.star.style.BreakType import PAGE_BEFORE, PAGE_AFTER, PAGE_BOTH, NONE
from com.sun.star.uno import Exception as UnoException, RuntimeException
from com.sun.star.connection import NoConnectException
from com.sun.star.lang import IllegalArgumentException
from com.sun.star.io import IOException
from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK 


#to parse argument from the cmd: 
"""
parser = OptionParser()
parser.add_option("", "--port",
                type="int", dest="port", default=2002,
                help="specify the port")
parser.add_option("", "--sequence",
                type="string", dest="seq", default=2002,
                help="specify the file that contains command sequences")

(options, args) = parser.parse_args()
"""

#this is a global dictionnary to handle the current document
ooo_info = dict()

def parse_seq(file_seq):
    """
    Parse le fichier de sequence
    -> dictionnary
    {cmd:args}
    """
    l = list()
    f = open(file_seq,'r')
    lines = str(f.read()).splitlines()
    for line in lines:
        if line is None:
            continue
        cmd = line.split(':')
        if len(cmd) == 2:
            l.append(line.decode("UTF-8"))
        else:
            print "cmd invalide :%s"%cmd
    return l

def cmd_open(ooo_info,syspath_doc):
    """
    ooo_info:
    {ctx:l objet, desktop: , ....}
    """
    url = systemPathToFileUrl(syspath_doc)
    #load as template:
    #pp_template = PropertyValue()
    #pp_template.Name  = "AsTemplate"
    #pp_template.Value = True
    pp_hidden = PropertyValue()
    pp_hidden.Name = "Hidden"
    pp_hidden.Value = True
    load_pp = (pp_hidden,)
    ooo_info['doc'] = ooo_info['desktop'].loadComponentFromURL(url,"_blank",0,load_pp)
    ooo_info['cursor'] = ooo_info['doc'].Text.createTextCursor()
    return ooo_info

def cmd_set_title(ooo_info,args):
    """
    Set le title of the current document
    """
    title = args
    ooo_info['doc'].DocumentInfo.Title = '%s'%title

def cmd_save_to(ooo_info,args):
    """
    #<toutpt> seems storeAsURL doesn t work with the filter writer_pdf_Export
    #<toutpt> on 2.0.4
    #<toutpt> working with storeToURL
    #<lgodard> toutpt : yes
    #<lgodard> toutpt: it is a feature - StoreASUrl is for filter that OOo can render
    writer_pdf_Export
    """
    dest_file = args.split(',')[0]
    format = args.split(',')[1]
    doc = ooo_info['doc']
    dest_url = systemPathToFileUrl(dest_file)
    print "save to %s from %s"%(format,dest_file)
    pp_filter = PropertyValue()
    pp_url = PropertyValue()
    pp_overwrite = PropertyValue()
    pp_filter.Name = "FilterName"
    pp_filter.Value = format
    pp_url.Name = "URL"
    pp_url.Value = dest_url
    pp_overwrite.Name = "Overwrite"
    pp_overwrite.Value = True
    outprop = (pp_filter, pp_url, pp_overwrite)
    doc.storeToURL(dest_url, outprop)

def cmd_save_as(ooo_info,args):
    """
    The main diff is save as render the result
    save current document.
    TEXT
    HTML (StarWriter)
    MS Word 97
    """
    dest_file = args.split(',')[0]
    format = args.split(',')[1]
    doc = ooo_info['doc']
    dest_url = systemPathToFileUrl(dest_file)
    print "save as %s from %s"%(format,dest_file)
    pp_filter = PropertyValue()
    pp_url = PropertyValue()
    pp_overwrite = PropertyValue()
    pp_selection = PropertyValue()
    pp_filter.Name = "FilterName"
    pp_filter.Value = format
    pp_url.Name = "URL"
    pp_url.Value = dest_url
    #pp_overwrite.Name = "Overwrite"
    #pp_overwrite.Value = True
    pp_selection.Name = "SelectionOnly"
    pp_selection.Value = True
    #outprop = (pp_filter, pp_url, pp_overwrite)
    outprop = (pp_filter, pp_url, pp_selection)
    doc.storeAsURL(dest_url, outprop)
    return ooo_info

def cmd_insert_breakpage(ooo_info,args=''):
    ooo_info['cursor'].gotoEnd(False)
    ooo_info['cursor'].gotoEndOfSentence(False)
    ooo_info['cursor'].BreakType = PAGE_AFTER
    ooo_info['doc'].Text.insertControlCharacter(ooo_info['cursor'], PARAGRAPH_BREAK, False)
    ooo_info['cursor'].PageDescName="Standard"

def cmd_refresh_toc(ooo_info):
    """
    Si une TOC est presente, elle est rafraichie
    """
    #list_names = ooo_info['doc'].DocumentIndexes.getElementNames()
    #co = ooo_info['doc'].DocumentIndexes.getCount()
    index = ooo_info['doc'].DocumentIndexes.getByIndex(0)
    index.update()

def cmd_set_pagestyle(ooo_info,args):
    pagestyle = args
    ooo_info['cursor'].PageDescName = pagestyle

def cmd_insert_file(ooo_info,args):
    file = args
    if os.path.isfile(file):
        cursor = ooo_info['cursor']
        try:
            fileUrl = systemPathToFileUrl(file)

            print "Appending %s" % fileUrl
            try:
                cursor.gotoEnd(False)
                cursor.insertDocumentFromURL(fileUrl, ())
            except IOException, e:
                sys.stderr.write("Error during opening ( " + e.Message + ")\n")
            except IllegalArgumentException, e:
                sys.stderr.write("The url is invalid ( " + e.Message + ")\n")

        except IOException, e:
            sys.stderr.write("Error during opening: " + e.Message + "\n")
        except UnoException, e:
            sys.stderr.write( "Error ("+repr(e.__class__)+") during conversion:" + 
                    e.Message + "\n" )
    else:
        raise IOException
    return ooo_info

def cmd_add_tag(ooo_info,args):
    """
    add a tag to current document to be able to switch on later.
    """
    tag = args
    if not ooo_info.has_key(tag):
        ooo_info[tag] = ooo_info['doc']

def cmd_switch_document(ooo_info,args):
    """
    """
    tag = args
    if not ooo_info.has_key(tag):
        raise "no document declared under the tag %s"%tag
    ooo_info['doc'] = ooo_info[tag]

def ooo_connect():
    """
    Connection to open office server.
    Fill a dictionnary and return it
    """
    # get the uno component context from the PyUNO runtime
    localContext = uno.getComponentContext()

    # create the UnoUrlResolver
    resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext )

    # connect to the running office
    try:
        ctx = resolver.resolve("uno:socket,host=Huntingdon,port=%s;urp;StarOffice.ComponentContext"%(2002))
    except:
        raise "impossible de se connecter au serveur openoffice"


    smgr = ctx.ServiceManager

    # get the central desktop object
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
    d = dict()
    d['ctx'] = ctx
    d['smgr'] = smgr
    d['doc'] = None
    d['desktop'] = desktop
    d['globals'] = globals()
    return d
