# -*- coding: utf-8 *- 
import os
from optparse import OptionParser

parser = OptionParser()
parser.add_option("", "--port",
                type="int", dest="port", default=2002,
                help="specify the port")
parser.add_option("", "--file",
                type="string", dest="file", default='error',
                help="specify the file to convert")
parser.add_option("", "--destfile",
                type="string", dest="destfile", default='error',
                help="specify the file where to save the pdf")

(options, args) = parser.parse_args()

ooo_info = dict()

def cmd_open(ooo_info,syspath_doc):
    """
    ooo_info:
    {ctx:l objet, desktop: , ....}
    """
    url = systemPathToFileUrl(syspath_doc)
    pp_hidden = PropertyValue()
    pp_hidden.Name = "Hidden"
    pp_hidden.Value = True
    load_pp = () #la pp hidden empeche la conversion PDF sur certaine version de OOo2
    ooo_info['doc'] = ooo_info['desktop'].loadComponentFromURL(url,"_blank",0,load_pp)
    ooo_info['cursor'] = ooo_info['doc'].Text.createTextCursor()

def export2pdf(ooo_info,dest_file):
    """
    """
    from unohelper import systemPathToFileUrl
    dest_url = systemPathToFileUrl(dest_file)
    print "save as pdf from %s"%(dest_file)
    pp_filter = PropertyValue()
    pp_filter.Name = "FilterName"
    pp_filter.Value = "writer_pdf_Export"
    outprop = (pp_filter, PropertyValue( "Overwrite" , 0, True , 0 ),)
    ooo_info['doc'].storeToURL(dest_url, outprop)

import uno
from unohelper import systemPathToFileUrl, absolutize
from com.sun.star.beans import PropertyValue

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
        ctx = resolver.resolve("uno:socket,host=localhost,port=%s;urp;StarOffice.ComponentContext"%(options.port))
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
    return d

ooo_info = ooo_connect()
if ooo_info:
    cmd_open(ooo_info,options.file)
    export2pdf(ooo_info,options.destfile)
    ooo_info['doc'].close(False)
else:
    raise 'problem de connexion au serveur OOo2'