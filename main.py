__author__ = 'david_zemens'

"""
Version 1
*****************************************************************
This script contains methods to build a PPAM file from module code under source control repository
        1. adds all of the VBComponents to the working PPTM file
        2. adds required project references
        3. converts the PPTM to a .ZIP
        4. Adds the CustomUI XML, jdplogo.jpg, etc. to the .ZIP directory
        5. converts the .ZIP to a PPTM

Known Issues: PowerPoint needs to "repair" this file the first time a user opens it.
User should open the PPTM output file, and then "repair" the file, then close PowerPoint and re-open.
Ensure that the code compiles and no references are missing/broken (if so, they can be added manually)

*****************************************************************

:param vba_source_control_path: specify the local folder which contains the modules to be imported
:param output_path: specify the destination path & filename (.pptm) for the build file
:param ribbon_xml_path: specify the destination path & filename for the Ribbon XML
:param ribbon_logo_path: specify the path & name of the logo file "jdplogo.jpg"
:param CustomUI: boolean, define this as False if you do NOT want to add the CustomUI xml/etc.

"""
import win32com.client
import os
import zipfile
import uuid

vba_source_control_path = r"C:\Repos\CB\ChartBuilder\VBA\ChartBuilder_PPT\Modules"
ribbon_xml_path = r"C:\Repos\CB\ChartBuilder\VBA\ChartBuilder_PPT\Ribbon XML\ribbon_xml.xml"
ribbon_logo_path = r"C:\Repos\CB\ChartBuilder\VBA\ChartBuilder_PPT\Ribbon XML\jdplogo.jpg"
output_path = r"C:\debug\output.pptm"
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

    _path=output_path.replace('.pptm', '.zip')
    copy_path=r'C:\debug\copy.zip'

    # Convert to ZIP archive
    os.rename(output_path, _path)
    z=zipfile.ZipFile(_path, 'a', zipfile.ZIP_DEFLATED)
    copy=zipfile.ZipFile(copy_path, 'w', zipfile.ZIP_DEFLATED)

    guid=str(uuid.uuid4()).replace('-', '')[:16]

    for itm in [itm for itm in z.infolist() if itm.filename != r'_rels/.rels']:
        buffer = z.read(itm.filename)
        copy.writestr(itm, buffer)

    # Append the Logo file to the .zip and create the archive
    copy.write(ribbon_logo_path, r'\CustomUI\images\jdplogo.jpg')

    # append the CustomUI xml part to the .zip and create the archive
    copy.write(ribbon_xml_path, r'\CustomUI\customUI14.xml')

    # create the string & append the .rels to CustomUI\_rels
    rels_xml=r'<?xml version="1.0" encoding="utf-8" ?>'
    rels_xml += r'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    rels_xml += r'<Relationship Id="jdplogo" Type=http://schemas.openxmlformats.org/officeDocument/2006/'
    rels_xml += r'relationships/image" Target="images/jdplogo.jpg"/>'
    rels_xml += r'</Relationships>'

    copy.writestr(r'CustomUI\_rels\customUI14.xml.rels', rels_xml.encode('utf-8'))

    # get the existing _rels/.rels XML content and copy to the copied archiveI:

    rels_xml = r'<?xml version="1.0" encoding="utf-8"?>'
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

    copy.writestr(r'_rels/.rels', rels_xml.encode('utf-8'))

    z.close()
    copy.close()

    os.remove(_path)
    #os.remove(output_path)
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
    d["Microsoft XML, v6.0"] = r'C:\Windows\System32\msxml6.dll'
    d["Microsoft Excel"] = r'C:\Program Files (x86)\Microsoft Office\Office'+version+'\EXCEL.EXE'
    d["Visual Basic for Applications Extensibility"] = r'C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA7\VBE7.DLL'
    #pres.VBProject.References.AddFromFile(r'C:\PROGRA~2\COMMON~1\MICROS~1\VBA\VBA7\VBE7.DLL')
    if version == '14':
        d["Microsoft Windows Common Controls 6.0 (SP6)"] = r'C:\Windows\SysWOW64\MSCOMCTL.OCX'
    else:
        d["Microsoft Windows Common Controls 6.0 (SP6)"] = r'C:\Windows\System32\MSCOMCTL.OCX'

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
