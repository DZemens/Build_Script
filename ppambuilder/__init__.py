__author__ = 'dz'
__all__ = ['ppamfactory', 'os_version']


"""
example:

from ppambuilder.ppamfactory import *
factory = PPAMFactory()
print(f'Is 64b windows? \t{factory.is64bwin}')

# creates without ribbon components
factory.create(r'c:\debug\test', r'c:\debug\test\output.pptm', r'c:\debug\test\copy.pptm')

# creates with ribbon + JPEG



factory.create(r'c:\debug\test\modules', r'c:\debug\test\output.pptm', r'c:\debug\test\copy.zip', r'c:\debug\test\ribbonxml\ribbon_xml.xml', r'c:\debug\test\ribbonxml\jdplogo.jpg')
"""
