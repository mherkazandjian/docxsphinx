#!/usr/bin/env python3.6
"""
Open and modify Microsoft Word 2007 docx files (called 'OpenXML' and 'Office OpenXML' by Microsoft)

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
"""

from lxml import etree
from PIL import Image
import zipfile
import shutil
import re
import time
import os
from os.path import join

# Record template directory's location which is just 'template' for a docx
# developer or 'site-packages/docx-template' if you have installed docx
TEMPLATE_DIR = join(os.path.dirname(__file__), 'docx-template')  # installed
if not os.path.isdir(TEMPLATE_DIR):
    TEMPLATE_DIR = join(os.path.dirname(__file__), 'template')  # dev

# FIXME: QUICK-HACK to prevent picture() from staining template directory.
# temporary directory will create per module import.
template_dir = TEMPLATE_DIR
def set_template(template_path):
    global template_dir
    template_dir = template_path
    update_stylenames(join(template_dir, 'word', 'styles.xml'))

# END of QUICK-HACK


# All Word prefixes / namespace matches used in document.xml & core.xml.
# LXML doesn't actually use prefixes (just the real namespace) , but these
# make it easier to copy Word output more easily. 
nsprefixes = {
    # Text Content
    'mv':'urn:schemas-microsoft-com:mac:vml',
    'mo':'http://schemas.microsoft.com/office/mac/office/2008/main',
    've':'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'o':'urn:schemas-microsoft-com:office:office',
    'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm':'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'v':'urn:schemas-microsoft-com:vml',
    'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w10':'urn:schemas-microsoft-com:office:word',
    'wne':'http://schemas.microsoft.com/office/word/2006/wordml',
    # Drawing
    'wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic':'http://schemas.openxmlformats.org/drawingml/2006/picture',
    # Properties (core and extended)
    'cp':"http://schemas.openxmlformats.org/package/2006/metadata/core-properties", 
    'dc':"http://purl.org/dc/elements/1.1/", 
    'dcterms':"http://purl.org/dc/terms/",
    'dcmitype':"http://purl.org/dc/dcmitype/",
    'xsi':"http://www.w3.org/2001/XMLSchema-instance",
    'ep':'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    # Content Types (we're just making up our own namespaces here to save time)
    'ct':'http://schemas.openxmlformats.org/package/2006/content-types',
    # Package Relationships (we're just making up our own namespaces here to save time)
    'pr':'http://schemas.openxmlformats.org/package/2006/relationships'
    }


# QUICK-HACK: style name mapping
# ex. Name of <... w:styleId='heading'> is not static name.
# Use <w:aliases w:val='Heading'> or <w:name w:val='Heading'>) static name.
stylenames = {
    'Normal': 'Normal',
    'Heading1': 'Heading1',
    'Heading2': 'Heading2',
    'Heading3': 'Heading3',
    'Heading4': 'Heading4',
    'Heading5': 'Heading5',
    'TableNormal': 'TableNormal',
    'NoList': 'NoList',
    'Title': 'Title',
    'SubTitle': 'SubTitle',
    'Strong': 'Bold',
    'Emphasis': 'Italic',
    'NoSpacing': 'NoSpacing',
    'BlockQuote': 'BlockQuote',
    'LiteralBlock': 'LiteralBlock',
    'BookTitle': 'BookTitle',
    'ListBullet': 'ListBullet',
    'ListNumber': 'ListNumber',
}


def norm_name(name, namespaces):
    ns, name = name.split(':', 1)
    ns = namespaces[ns]
    return "{%s}%s" % (ns, name)


def update_stylenames(style_file):
    xmlcontent = open(style_file).read().encode()
    xml = etree.fromstring(xmlcontent)
    style_elems = xml.xpath('w:style', namespaces=nsprefixes)
    for style_elem in style_elems:
        aliases_elems = style_elem.xpath('w:aliases', namespaces=nsprefixes)
        if aliases_elems:
            name = aliases_elems[0].attrib[norm_name('w:val', nsprefixes)]
        else:
            name_elem = style_elem.xpath('w:name', namespaces=nsprefixes)[0]
            name = name_elem.attrib[norm_name('w:val', nsprefixes)]
        value = style_elem.attrib[norm_name('w:styleId', nsprefixes)]
        stylenames[name] = value
        print("### '%s' = '%s'" % (name, value))


import tempfile
temp_dir = tempfile.mkdtemp(prefix='docx-')
os.rmdir(temp_dir)
shutil.copytree(TEMPLATE_DIR, temp_dir)
set_template(temp_dir)
# END of QUICK-HACK


def opendocx(file):
    """
    Open a docx file, return a document XML tree
    .. todo:: add doc
    :param file:
    :return:
    """
    mydoc = zipfile.ZipFile(file)
    xmlcontent = mydoc.read('word/document.xml').encode()
    document = etree.fromstring(xmlcontent)
    return document


def newdocument():
    """
    .. todo:: add doc
    :return:
    """
    document = makeelement('document')
    document.append(makeelement('body'))
    return document


def makeelement(tagname,
                tagtext=None,
                nsprefix='w',
                attributes=None,
                attrnsprefix=None):
    """
    Create an element & return it
    .. todo:: add doc
    :param tagname:
    :param tagtext:
    :param nsprefix:
    :param attributes:
    :param attrnsprefix:
    :return:
    """
    # Deal with list of nsprefix by making namespacemap
    namespacemap = None
    if type(nsprefix) == list:
        namespacemap = {}
        for prefix in nsprefix:
            namespacemap[prefix] = nsprefixes[prefix]
        nsprefix = nsprefix[0] # FIXME: rest of code below expects a single prefix
    if nsprefix:
        namespace = '{'+nsprefixes[nsprefix]+'}'
    else:
        # For when namespace = None
        namespace = ''
    newelement = etree.Element(namespace+tagname, nsmap=namespacemap)
    # Add attributes with namespaces
    if attributes:
        # If they haven't bothered setting attribute namespace, use an empty string
        # (equivalent of no namespace)
        if not attrnsprefix:
            # Quick hack: it seems every element that has a 'w' nsprefix for its tag uses the same prefix for it's attributes  
            if nsprefix == 'w':
                attributenamespace = namespace
            else:
                attributenamespace = ''
        else:
            attributenamespace = '{'+nsprefixes[attrnsprefix]+'}'
                    
        for tagattribute in attributes:
            newelement.set(attributenamespace+tagattribute, attributes[tagattribute])
    if tagtext:
        newelement.text = tagtext    
    return newelement


def pagebreak(type='page', orient='portrait'):
    """
    Insert a break, default 'page'.
    See http://openxmldeveloper.org/forums/thread/4075.aspx
    Return our page break element."""
    # Need to enumerate different types of page breaks.
    validtypes = ['page', 'section']
    if type not in validtypes:
        raise ValueError('Page break style "%s" not implemented. Valid styles: %s.' % (type, validtypes))
    pagebreak = makeelement('p')
    if type == 'page':
        run = makeelement('r')
        br = makeelement('br',attributes={'type':type})
        run.append(br)
        pagebreak.append(run)
    elif type == 'section':
        pPr = makeelement('pPr')
        sectPr = makeelement('sectPr')
        if orient == 'portrait':
            pgSz = makeelement('pgSz',attributes={'w':'12240','h':'15840'})
        elif orient == 'landscape':
            pgSz = makeelement('pgSz',attributes={'h':'12240','w':'15840', 'orient':'landscape'})
        sectPr.append(pgSz)
        pPr.append(sectPr)
        pagebreak.append(pPr)
    return pagebreak    


def paragraph(paratext, style='BodyText', breakbefore=False):
    """
    Make a new paragraph element, containing a run, and some text.
    Return the paragraph element.
    .. todo:: add doc
    :param paratext:
    :param style:
    :param breakbefore:
    :return:
    """
    # Make our elements
    paragraph = makeelement('p')
    run = makeelement('r')    
    
    # Insert lastRenderedPageBreak for assistive technologies like
    # document narrators to know when a page break occurred.
    if breakbefore:
        lastRenderedPageBreak = makeelement('lastRenderedPageBreak')
        run.append(lastRenderedPageBreak)    
    text = makeelement('t',tagtext=paratext)
    pPr = makeelement('pPr')
    style = stylenames.get(style, 'BodyText')
    pStyle = makeelement('pStyle',attributes={'val':style})
    pPr.append(pStyle)
                
    # Add the text the run, and the run to the paragraph
    run.append(text)    
    paragraph.append(pPr)    
    paragraph.append(run)    
    # Return the combined paragraph
    return paragraph


def contenttypes():
    """
    .. todo:: add doc
    :return:
    """
    prev_dir = os.getcwd() # save previous working dir
    os.chdir(template_dir)

    filename = '[Content_Types].xml'
    if not os.path.exists(filename):
        raise RuntimeError('You need %r file in template' % filename)

    parts = dict([
        (x.attrib['PartName'], x.attrib['ContentType'])
        for x in etree.fromstring(open(filename).read().encode()).xpath('*')
        if 'PartName' in x.attrib
    ])

    # FIXME - doesn't quite work...read from string as temp hack...
    #types = makeelement('Types',nsprefix='ct')
    types = etree.fromstring('''<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>''')
    for part in parts:
        types.append(makeelement('Override',nsprefix=None,attributes={'PartName':part,'ContentType':parts[part]}))
    # Add support for filetypes
    filetypes = {'rels':'application/vnd.openxmlformats-package.relationships+xml','xml':'application/xml','jpeg':'image/jpeg','gif':'image/gif','png':'image/png'}
    for extension in filetypes:
        types.append(makeelement('Default',nsprefix=None,attributes={'Extension':extension,'ContentType':filetypes[extension]}))

    os.chdir(prev_dir)
    return types


def heading(headingtext, headinglevel):
    """
    Make a new heading, return the heading element
    .. todo:: add doc
    :param headingtext:
    :param headinglevel:
    :return:
    """
    # Make our elements
    paragraph = makeelement('p')
    pr = makeelement('pPr')
    style = stylenames.get('Heading' + str(headinglevel), 'Normal')
    pStyle = makeelement('pStyle',attributes={'val': style})
    run = makeelement('r')
    text = makeelement('t',tagtext=headingtext)
    # Add the text the run, and the run to the paragraph
    pr.append(pStyle)
    run.append(text)
    paragraph.append(pr)   
    paragraph.append(run)    
    # Return the combined paragraph
    return paragraph   


def table(contents):
    """Get a list of lists, return a table
    .. todo:: add doc
    :param contents:
    :return:
    """
    table = makeelement('tbl')
    print('WARNING'*100)
    print('tables are temporarily disabled due to bad formatting')
    print('END WARNING'*100)
    return table
    columns = len(contents[0][0])    
    # Table properties
    tableprops = makeelement('tblPr')
    tablestyle = makeelement('tblStyle',attributes={'val':'ColorfulGrid-Accent1'})
    tablewidth = makeelement('tblW',attributes={'w':'0','type':'auto'})
    tablelook = makeelement('tblLook',attributes={'val':'0400'})
    for tableproperty in [tablestyle,tablewidth,tablelook]:
        tableprops.append(tableproperty)
    table.append(tableprops)    
    # Table Grid    
    tablegrid = makeelement('tblGrid')
    for _ in range(columns):
        tablegrid.append(makeelement('gridCol',attributes={'w':'2390'}))
    table.append(tablegrid)     
    # Heading Row    
    row = makeelement('tr')
    rowprops = makeelement('trPr')
    cnfStyle = makeelement('cnfStyle',attributes={'val':'000000100000'})
    rowprops.append(cnfStyle)
    row.append(rowprops)
    for heading in contents[0]:
        cell = makeelement('tc')  
        # Cell properties  
        cellprops = makeelement('tcPr')
        cellwidth = makeelement('tcW',attributes={'w':'2390','type':'dxa'})
        cellstyle = makeelement('shd',attributes={'val':'clear','color':'auto','fill':'548DD4','themeFill':'text2','themeFillTint':'99'})
        cellprops.append(cellwidth)
        cellprops.append(cellstyle)
        cell.append(cellprops)        
        # Paragraph (Content)
        cell.append(paragraph(heading))
        row.append(cell)
    table.append(row)            
    # Contents Rows   
    for contentrow in contents[1:]:
        row = makeelement('tr')     
        for content in contentrow:   
            cell = makeelement('tc')
            # Properties
            cellprops = makeelement('tcPr')
            cellwidth = makeelement('tcW',attributes={'type':'dxa'})
            cellprops.append(cellwidth)
            cell.append(cellprops)
            # Paragraph (Content)
            cell.append(paragraph(content))
            row.append(cell)    
        table.append(row)   
    return table                 


def picture(relationshiplist,
            picname,
            picdescription,
            pixelwidth=None,
            pixelheight=None,
            nochangeaspect=True,
            nochangearrowheads=True):
    """
    Take a relationshiplist, picture file name, and return a paragraph
     containing the image and an updated relationshiplist
    .. todo:: add doc
    :param relationshiplist:
    :param picname:
    :param picdescription:
    :param pixelwidth:
    :param pixelheight:
    :param nochangeaspect:
    :param nochangearrowheads:
    :return: """
    # http://openxmldeveloper.org/articles/462.aspx
    # Create an image. Size may be specified, otherwise it will based on the
    # pixel size of image. Return a paragraph containing the picture'''  
    # Copy the file into the media dir
    media_dir = join(template_dir,'word','media')
    if not os.path.isdir(media_dir):
        os.mkdir(media_dir)
    picpath, picname = os.path.abspath(picname), os.path.basename(picname)
    shutil.copyfile(picpath, join(media_dir,picname))

    # Check if the user has specified a size
    if not pixelwidth or not pixelheight:
        # If not, get info from the picture itself
        pixelwidth, pixelheight = Image.open(picpath).size[0:2]

    # OpenXML measures on-screen objects in English Metric Units
    # 1cm = 36000 EMUs            
    emuperpixel = 12667
    width = str(pixelwidth * emuperpixel)
    height = str(pixelheight * emuperpixel)   
    
    # Set relationship ID to the first available  
    picid = '2'    
    picrelid = 'rId'+str(len(relationshiplist)+1)
    relationshiplist.append([
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        'media/'+picname])
    
    # There are 3 main elements inside a picture
    # 1. The Blipfill - specifies how the image fills the picture area (stretch, tile, etc.)
    blipfill = makeelement('blipFill',nsprefix='pic')
    blipfill.append(makeelement('blip',nsprefix='a',attrnsprefix='r',attributes={'embed':picrelid}))
    stretch = makeelement('stretch',nsprefix='a')
    stretch.append(makeelement('fillRect',nsprefix='a'))
    blipfill.append(makeelement('srcRect',nsprefix='a'))
    blipfill.append(stretch)
    
    # 2. The non visual picture properties 
    nvpicpr = makeelement('nvPicPr',nsprefix='pic')
    cnvpr = makeelement('cNvPr',nsprefix='pic',
                        attributes={'id':'0','name':'Picture 1','descr':picname}) 
    nvpicpr.append(cnvpr) 
    cnvpicpr = makeelement('cNvPicPr',nsprefix='pic')                           
    cnvpicpr.append(makeelement('picLocks', nsprefix='a', 
                    attributes={'noChangeAspect':str(int(nochangeaspect)),
                    'noChangeArrowheads':str(int(nochangearrowheads))}))
    nvpicpr.append(cnvpicpr)
        
    # 3. The Shape properties
    sppr = makeelement('spPr',nsprefix='pic',attributes={'bwMode':'auto'})
    xfrm = makeelement('xfrm',nsprefix='a')
    xfrm.append(makeelement('off',nsprefix='a',attributes={'x':'0','y':'0'}))
    xfrm.append(makeelement('ext',nsprefix='a',attributes={'cx':width,'cy':height}))
    prstgeom = makeelement('prstGeom',nsprefix='a',attributes={'prst':'rect'})
    prstgeom.append(makeelement('avLst',nsprefix='a'))
    sppr.append(xfrm)
    sppr.append(prstgeom)
    
    # Add our 3 parts to the picture element
    pic = makeelement('pic',nsprefix='pic')    
    pic.append(nvpicpr)
    pic.append(blipfill)
    pic.append(sppr)
    
    # Now make the supporting elements
    # The following sequence is just: make element, then add its children
    graphicdata = makeelement('graphicData',nsprefix='a',
        attributes={'uri':'http://schemas.openxmlformats.org/drawingml/2006/picture'})
    graphicdata.append(pic)
    graphic = makeelement('graphic',nsprefix='a')
    graphic.append(graphicdata)

    framelocks = makeelement('graphicFrameLocks',nsprefix='a',attributes={'noChangeAspect':'1'})    
    framepr = makeelement('cNvGraphicFramePr',nsprefix='wp')
    framepr.append(framelocks)
    docpr = makeelement('docPr',nsprefix='wp',
        attributes={'id':picid,'name':'Picture 1','descr':picdescription})
    effectextent = makeelement('effectExtent',nsprefix='wp',
        attributes={'l':'25400','t':'0','r':'0','b':'0'})
    extent = makeelement('extent',nsprefix='wp',attributes={'cx':width,'cy':height})
    inline = makeelement('inline',
        attributes={'distT':"0",'distB':"0",'distL':"0",'distR':"0"},nsprefix='wp')
    inline.append(extent)
    inline.append(effectextent)
    inline.append(docpr)
    inline.append(framepr)
    inline.append(graphic)
    drawing = makeelement('drawing')
    drawing.append(inline)
    run = makeelement('r')
    run.append(drawing)
    paragraph = makeelement('p')
    paragraph.append(run)
    return relationshiplist,paragraph


def search(document, _search):
    """
    Search a document for a regex, return success / fail result
    .. todo:: add doc
    :param document: 
    :param _search: 
    :return: 
    """
    result = False
    searchre = re.compile(_search)
    for element in document.iter():
        if element.tag == '{%s}t' % nsprefixes['w']: # t (text) elements
            if element.text:
                if searchre.search(element.text):
                    result = True
    return result


def replace(document, search, replace):
    """
    Replace all occurences of string with a different string, return updated document
    .. todo:: add doc
    :param document:
    :param search:
    :param replace:
    :return:
    """
    newdocument = document
    searchre = re.compile(search)
    for element in newdocument.iter():
        if element.tag == '{%s}t' % nsprefixes['w']: # t (text) elements
            if element.text:
                if searchre.search(element.text):
                    element.text = re.sub(search,replace,element.text)
    return newdocument


def getdocumenttext(document):
    """
    Return the raw text of a document, as a list of paragraphs.
    .. todo:: add doc
    :param document:
    :return:
    """
    paratextlist=[]   
    # Compile a list of all paragraph (p) elements
    paralist = []
    for element in document.iter():
        # Find p (paragraph) elements
        if element.tag == '{'+nsprefixes['w']+'}p':
            paralist.append(element)    
    # Since a single sentence might be spread over multiple text elements, iterate through each 
    # paragraph, appending all text (t) children to that paragraphs text.     
    for para in paralist:      
        paratext=u''  
        # Loop through each paragraph
        for element in para.iter():
            # Find t (text) elements
            if element.tag == '{'+nsprefixes['w']+'}t':
                if element.text:
                    paratext = paratext+element.text
        # Add our completed paragraph text to the list of paragraph text    
        if not len(paratext) == 0:
            paratextlist.append(paratext)                    
    return paratextlist        


def coreproperties(title, subject, creator, keywords, lastmodifiedby=None):
    """Create core properties (common document properties referred to in the
     'Dublin Core' specification). See appproperties() for other stuff.
    .. todo:: add doc
    :param title:
    :param subject:
    :param creator:
    :param keywords:
    :param lastmodifiedby:
    :return:
    """
    coreprops = makeelement('coreProperties',nsprefix='cp')    
    coreprops.append(makeelement('title',tagtext=title,nsprefix='dc'))
    coreprops.append(makeelement('subject',tagtext=subject,nsprefix='dc'))
    coreprops.append(makeelement('creator',tagtext=creator,nsprefix='dc'))
    coreprops.append(makeelement('keywords',tagtext=','.join(keywords),nsprefix='cp'))    
    if not lastmodifiedby:
        lastmodifiedby = creator
    coreprops.append(makeelement('lastModifiedBy',tagtext=lastmodifiedby,nsprefix='cp'))
    coreprops.append(makeelement('revision',tagtext='1',nsprefix='cp'))
    coreprops.append(makeelement('category',tagtext='Examples',nsprefix='cp'))
    coreprops.append(makeelement('description',tagtext='Examples',nsprefix='dc'))
    currenttime = time.strftime('%Y-%m-%dT%H:%M:%SZ')
    # Document creation and modify times
    # Prob here: we have an attribute who name uses one namespace, and that 
    # attribute's value uses another namespace.
    # We're creating the lement from a string as a workaround...
    for doctime in ['created','modified']:
        coreprops.append(etree.fromstring('''<dcterms:'''+doctime+''' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:dcterms="http://purl.org/dc/terms/" xsi:type="dcterms:W3CDTF">'''+currenttime+'''</dcterms:'''+doctime+'''>'''))
        pass
    return coreprops


def appproperties():
    """
    Create app-specific properties. See docproperties() for more common document
    properties.
    .. todo:: add doc
    :return:
    """
    appprops = makeelement('Properties',nsprefix='ep')
    appprops = etree.fromstring(
    '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"></Properties>'''.encode())
    props = {
            'Template':'Normal.dotm',
            'TotalTime':'6',
            'Pages':'1',  
            'Words':'83',   
            'Characters':'475', 
            'Application':'Microsoft Word 12.0.0',
            'DocSecurity':'0',
            'Lines':'12', 
            'Paragraphs':'8',
            'ScaleCrop':'false', 
            'LinksUpToDate':'false', 
            'CharactersWithSpaces':'583',  
            'SharedDoc':'false',
            'HyperlinksChanged':'false',
            'AppVersion':'12.0000',    
            }
    for prop in props:
        appprops.append(makeelement(prop,tagtext=props[prop],nsprefix=None))
    return appprops


def websettings():
    """
    Generate websettings
    .. todo:: add doc
    """
    web = makeelement('webSettings')
    web.append(makeelement('allowPNG'))
    web.append(makeelement('doNotSaveAsSingleFile'))
    return web


def relationshiplist():
    """
    .. todo:: add doc
    :return: .. todo:: add doc
    """
    prev_dir = os.getcwd() # save previous working dir
    os.chdir(template_dir)

    filename = 'word/_rels/document.xml.rels'
    if not os.path.exists(filename):
        raise RuntimeError('You need %r file in template' % filename)

    relationships = etree.fromstring(open(filename).read().encode())
    relationshiplist = [
            [x.attrib['Type'], x.attrib['Target']]
            for x in relationships.xpath('*')
    ]

    os.chdir(prev_dir)
    return relationshiplist


def wordrelationships(relationshiplist):
    """
    Generate a Word relationships file
    .. todo:: add doc
    :param relationshiplist:
    :return:
    """
    # Default list of relationships
    # FIXME: using string hack instead of making element
    #relationships = makeelement('Relationships',nsprefix='pr')    
    relationships = etree.fromstring(
    '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">      	
        </Relationships>'''    
    )
    count = 0
    for relationship in relationshiplist:
        # Relationship IDs (rId) start at 1.
        relationships.append(makeelement('Relationship',attributes={'Id':'rId'+str(count+1),
        'Type':relationship[0],'Target':relationship[1]},nsprefix=None))
        count += 1
    return relationships    


def savedocx(document,
             coreprops,
             appprops,
             contenttypes,
             websettings,
             wordrelationships,
             docxfilename):
    """
    Save a modified document

    .. todo:: add doc

    :param document:
    :param coreprops:
    :param appprops:
    :param contenttypes:
    :param websettings:
    :param wordrelationships:
    :param docxfilename:
    :return:
    """
    assert os.path.isdir(template_dir)
    docxfile = zipfile.ZipFile(docxfilename,mode='w',compression=zipfile.ZIP_DEFLATED)
    
    # Move to the template data path
    prev_dir = os.path.abspath('.') # save previous working dir
    os.chdir(template_dir)
    
    # Serialize our trees into out zip file
    treesandfiles = {document:'word/document.xml',
                     coreprops:'docProps/core.xml',
                     appprops:'docProps/app.xml',
                     contenttypes:'[Content_Types].xml',
                     websettings:'word/webSettings.xml',
                     wordrelationships:'word/_rels/document.xml.rels'}
    for tree in treesandfiles:
        print('Saving: '+treesandfiles[tree])
        treestring = etree.tostring(tree, pretty_print=True)
        docxfile.writestr(treesandfiles[tree],treestring)

    # Add & compress support files
    files_to_ignore = ['.DS_Store'] # nuisance from some os's
    files_to_skip = treesandfiles.values()
    for dirpath,dirnames,filenames in os.walk('.'):
        for filename in filenames:
            if filename in files_to_ignore:
                continue
            templatefile = join(dirpath,filename)
            archivename = os.path.normpath(templatefile)
            archivename = '/'.join(archivename.split(os.sep))  # multibyte ok?
            if archivename in files_to_skip:
                continue
            print('Saving: ' + archivename)
            docxfile.write(templatefile, archivename)
    print('Saved new file to: ' + docxfilename)
    os.chdir(prev_dir)  # restore previous working dir
    return
