__author__ = 'david_zemens'

"""
Version 2
*****************************************************************
This script contains methods to build a PPAM file from module code under source control repository
        1. adds all of the VBComponents to the working PPTM file
        2. adds required project references
        3. converts the PPTM to a .ZIP
        4. Adds the CustomUI XML, jdplogo.jpg, etc. to the .ZIP directory
        5. converts the .ZIP to a PPTM

*****************************************************************

:param vba_source_control_path: specify the local folder which contains the modules to be imported
:param output_path: specify the destination path & filename (.pptm) for the build file
;param copy_path: a temporary copy of the .ZIP archive must be created for read/write
:param ribbon_xml_path: specify the destination path & filename for the Ribbon XML
:param ribbon_logo_path: specify the path & name of the logo file "jdplogo.jpg"
:param CustomUI: boolean, define this as False if you do NOT want to add the CustomUI xml/etc.

"""
import win32com.client
import os
import zipfile
import uuid
import os_version
import xml.etree.ElementTree as ET


vba_source_control_path = r"C:\Repos\CB\ChartBuilder\VBA\ChartBuilder_PPT\Modules"
ribbon_xml_path = r"C:\Repos\CB\ChartBuilder\VBA\ChartBuilder_PPT\Ribbon XML\ribbon_xml.xml"
ribbon_logo_path = r"C:\Repos\CB\ChartBuilder\VBA\ChartBuilder_PPT\Ribbon XML\jdplogo.jpg"
output_path = r"C:\debug\output.pptm"
copy_path=r'C:\debug\copy.zip'
customUI = True

def build_addin(pres, path):
    """
    :param pres : PowerPoint Presentation
    :param path : path where the file will be output
    :rtype boolean
    This procedure does the following:
        1. adds all of the VBComponents to the working PPTM file
        2. adds required project references
    The .PPTM file is used for local development & debugging and
    is only usually packaged as a PPAM for Testing and Distribution
    """
    refs = {}
    try:
        # import the VB Components
        for fn in [fn for fn in os.listdir(path) if not(fn[-4:]=='.frx')]:
            pres.VBProject.VBComponents.Import(path + "\\" + fn)

        refs = ref_dict(pres)

        for k, v in refs.iteritems():
            try:
                pres.VBProject.References.AddFromFile(v)

            except Exception:
                # non-critical errors can be passed.
                print 'failed to add reference to: ' + k + ' from ' + v
                pass

        # Clean up old files, if any
        if os.path.isfile(output_path):
            os.remove(output_path)
        if os.path.isfile(output_path.replace(".pptm", ".zip")):
            os.remove(output_path.replace(".pptm", ".zip"))

        return True

    except Exception:
        raise Exception



def build_ribbon_zip():

    """
        build_ribbon_zip handles manipulation of the .ZIP contents and places the
        necessary components within the PPTM ZIP archive structure
        3. converts the PPTM to a .ZIP
        4. Adds the CustomUI XML to the .ZIP directory
        5. converts the .ZIP to a PPTM

    """
    bom = u'\ufeff'
    _path=output_path.replace('.pptm', '.zip')


    # Convert to ZIP archive
    os.rename(output_path, _path)
    z=zipfile.ZipFile(_path, 'a', zipfile.ZIP_DEFLATED)
    copy=zipfile.ZipFile(copy_path, 'w', zipfile.ZIP_DEFLATED)

    guid=str(uuid.uuid4()).replace('-', '')[:16]

    """
        the .rels files are written directly from XML string built in procedure
        the [Content_Types].xml file needs to include additional parameter for the 'jpg' extension
    """
    for itm in [itm for itm in z.infolist() if itm.filename != r'_rels/.rels']:
        buffer = z.read(itm.filename)
        if itm.filename == "[Content_Types].xml":
            # Modify the [Content_Types].xml file to include the jpg reference
            # <Default Extension="jpg" ContentType="image/.jpg" />
            # copy the XML from the original zip archive, this file has not been copied in the above loop
            root = ET.fromstring(buffer)

            ET.SubElement(root, '{http://schemas.openxmlformats.org/package/2006/content-types}Default', {'Extension': 'jpg', 'ContentType': 'image/.jpg'})

            copy.writestr(itm, ET.tostring(root).encode('utf-8'))

            # Append the Logo file to the .zip and create the archive
            copy.write(ribbon_logo_path, r'\customUI\images\jdplogo.jpg')

        else:
            copy.writestr(itm, buffer)

    # append the CustomUI xml part to the .zip and create the archive
    copy.write(ribbon_xml_path, r'\customUI\customUI14.xml')

    # create the string & append the .rels to CustomUI\_rels
    rels_xml = """<?xml version="1.0" encoding="utf-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/jdplogo.jpg" Id="jdplogo" />
    </Relationships>"""

    copy.writestr(r'customUI\_rels\customUI14.xml.rels', rels_xml.encode('utf-8'))

    # get the existing _rels/.rels XML content and copy to the copied archiveI:

    rels_xml = r'<?xml version="1.0" encoding="utf-8" ?>'
    rels_xml += r'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    rels_xml += r'<Relationship Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/'
    rels_xml += r'core-properties" '
    rels_xml += r'Target="docProps/core.xml" Id="rId3" />'
    rels_xml += r'<Relationship Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" '
    rels_xml += r'Target="docProps/thumbnail.jpeg" Id="rId2" />'
    rels_xml += r'<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    rels_xml += r'Target="ppt/presentation.xml" Id="rId1" />'
    rels_xml += r'<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
    rels_xml += r'Target="docProps/app.xml" Id="rId4" /><Relationship '
    rels_xml += r'Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" '
    rels_xml += r'Target="/customUI/customUI14.xml" Id="R'+guid+'" /></Relationships>'

    copy.writestr(r'_rels\.rels', rels_xml.encode('utf-8'))

    z.close()
    copy.close()

    os.remove(_path)
    os.rename(copy_path, output_path)

def ref_dict(pres):
    """
    builds a dictionary of reference string paths to add to the pres.VBProject
    :param pres: a PowerPoint Presentation instance
    :return: refs{} dictionary
    """

    d = {}
    version = str(int(float(pres.Application.version)))

    d["Microsoft ActiveX Data Objects 6.1 Libarary"] = r'C:\Program Files (x86)\Common Files\System\ado\msado15.dll'
    #d["VBA"] = r'C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA7\VBE7.DLL'
    # .AddFromGuid('{000204EF-0000-0000-C000-000000000046}')
    d["VBIDE"] = r'C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB'
    d["Microsoft XML, v6.0"] = r'C:\Windows\System32\msxml6.dll'
    d["Microsoft Excel"] = r'C:\Program Files (x86)\Microsoft Office\Office'+version+'\EXCEL.EXE'
    # Note: stdole and MSForms are unnecessary
    if os_version.Is64Windows():
        #64-b location
        d["Microsoft Windows Common Controls 6.0 (SP6)"] = r'C:\Windows\SysWOW64\MSCOMCTL.OCX'
        #d["stdole2"] = r'C:\Windows\SysWOW64\stdole2.tlb'
        #d["MSForms"] = r'C:\Windows\SysWOW64\FM20.DLL'
    else:
        d["Microsoft Windows Common Controls 6.0 (SP6)"] = r'C:\Windows\System32\MSCOMCTL.OCX'
        #d["MSForms"] = r'C:\Windows\System32\FM20.DLL'
        #d["stdole2"] = r'C:\Windows\System32\stdole2.tlb'

    return d


if __name__ == "__main__":
    """
    Procedure to create a new PowerPoint Presentation and insert the Code Modules from source control
    optionally this will also build the ribbon components (currently a work-in-progress)

    """
    ppApp = win32com.client.Dispatch("PowerPoint.Application")

    pres = ppApp.Presentations.Add(False)

    pres.Slides.AddSlide(1, pres.SlideMaster.CustomLayouts(1))

    if build_addin(pres, vba_source_control_path)==True:

        # Save the new file with VBProject components
        pres.SaveAs(output_path)

        pres.Close()

        ppApp.Quit()

        if customUI:
            build_ribbon_zip()
