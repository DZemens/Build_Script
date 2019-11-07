__author__ = 'dz'

import os
import zipfile
import uuid
import xml.etree.ElementTree as ET

from win32com import client

from ppambuilder import os_version


class PPAMFactory:
    """
    The PPAMFactory instantiates a PowerPoint.Application COMObject.
    Use the `create` method to build new PPAM from source control.
    """
    def __init__(self):
        self._app = None
        self._is64bwin = os_version.Is64Windows()
        self._dispatch()

    def __del__(self):
        """
        this will close the PowerPoint.Application automatically.

        .. todo: should we keep this, or handle it more explicitly?
        """
        self._app.Quit()

    def _dispatch(self):
        if not self._app:
            self._app = client.Dispatch("POWERPOINT.APPLICATION")

    @property
    def is64bwin(self):
        return self._is64bwin

    @property
    def app(self):
        """
        read-only property
        :return: PowerPoint.Application COMObject
        """
        return self._app

    def create(self, vba_src_path:str, ribbon_path:str, logo_path:str, output_path:str, copy_path:str, custom_ui:bool=True):
        """

        """
        self._dispatch()
        pres = self.app.Presentations.Add(False)
        pres.Slides.AddSlide(1, pres.SlideMaster.CustomLayouts(1))
        if build_addin(pres, self.is64bwin, vba_src_path, output_path):
            # Save the new file with VBProject components
            pres.SaveAs(output_path)
            pres.Close()
            if custom_ui:
                build_ribbon_zip(output_path, copy_path, logo_path, ribbon_path)


def build_addin(pres, is64bwin, vba_src_path, output_path):
    """
    :param pres : PowerPoint Presentation
    :param is64bwin: bool, OS is 64-bit?
    :param vba_src_path : vba_src_path where the file will be output
    :param output_path: path to the output file
    :rtype boolean
    This procedure does the following:
        1. adds all of the VBComponents to the working PPTM file
        2. adds required project references
    The .PPTM file is used for local development & debugging and
    is only usually packaged as a PPAM for Testing and Distribution
    """
    valid_extensions = ['.bas', '.cls']
    try:
        # import the VB Components
        for fn in [fn for fn in os.listdir(vba_src_path) if fn[-4:] in valid_extensions]:
            pres.VBProject.VBComponents.Import(os.path.join(vba_src_path, fn))

        refs = ref_dict(pres, is64bwin)
        for k, v in refs.items():
            try:
                pres.VBProject.References.AddFromFile(v)
            except Exception:
                # non-critical errors can be passed.
                print(f'failed to add reference to: {k} from {v}')
                pass

        # Clean up old files, if any
        if os.path.isfile(output_path):
            os.remove(output_path)
        if os.path.isfile(output_path.replace(".pptm", ".zip")):
            os.remove(output_path.replace(".pptm", ".zip"))

        return True

    except Exception:
        raise Exception


def build_ribbon_zip(output_path:str, copy_path:str, ribbon_logo_path:str, ribbon_xml_path:str):

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
        if itm.filename != "[Content_Types].xml":
            copy.writestr(itm, buffer)
        else:
            # Modify the [Content_Types].xml file to include the jpg reference
            # <Default Extension="jpg" ContentType="image/.jpg" />
            # copy the XML from the original zip archive, this file has not been copied in the above loop
            root = ET.fromstring(buffer)
            ET.SubElement(root, '{http://schemas.openxmlformats.org/package/2006/content-types}Default',
                          {'Extension': 'jpg', 'ContentType': 'image/.jpg'})
            copy.writestr(itm, ET.tostring(root).encode('utf-8'))
            # Append the Logo file to the .zip and create the archive
            copy.write(ribbon_logo_path, r'\customUI\images\jdplogo.jpg')

    # append the CustomUI xml part to the .zip and create the archive
    copy.write(ribbon_xml_path, r'\customUI\customUI14.xml')

    # create the string & append the .rels to CustomUI\_rels
    rels_xml = """<?xml version="1.0" encoding="utf-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/jdplogo.jpg" Id="jdplogo" />
    </Relationships>"""

    #copy.writestr(r'customUI\_rels\customUI14.xml.rels', rels_xml.encode('utf-8'))
    copy.writestr(r'customUI\_rels\customUI14.xml.rels', rels_xml)

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

    copy.writestr(r'_rels\.rels', rels_xml)  # rels_xml.encode('utf-8')

    z.close()
    copy.close()

    os.remove(_path)
    os.rename(copy_path, output_path)

def ref_dict(pres, is64bwin:bool):
    """
    builds a dictionary of reference string paths to add to the pres.VBProject
    :param pres: a PowerPoint Presentation instance
    :return: refs{} dictionary
    """
    version = str(int(float(pres.Application.version)))
    comctl = 'SysWow64' if is64bwin else 'System32'
    excel_ref = f'\Microsoft Office 15\Root\Office{version}' if is64bwin else ' (x86)\Microsoft Office\Office{version}'
    d = {}

    d["Microsoft ActiveX Data Objects 6.1 Library"] = r'C:\Program Files (x86)\Common Files\System\ado\msado15.dll'
    #d["VBA"] = r'C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA7\VBE7.DLL'
    # .AddFromGuid('{000204EF-0000-0000-C000-000000000046}')
    d["VBIDE"] = r'C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB'
    d["Microsoft XML, v6.0"] = r'C:\Windows\System32\msxml6.dll'
    """
        this path structure seems to be inconsistent versus previous versions of Office, it used to be like:
        d["Microsoft Excel"] = f'C:\Program Files{excel_ref}\EXCEL.EXE'
    """
    d["Microsoft Excel"] = f'C:\Program Files{excel_ref}\EXCEL.EXE'
    d["Microsoft Windows Common Controls 6.0 (SP6)"] = f'C:\Windows\{comctl}\MSCOMCTL.OCX'

    return d