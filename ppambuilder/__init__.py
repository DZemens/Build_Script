__author__ = 'dz'
__all__ = ['ppamfactory', 'os_version']


"""
example:

from ppambuilder.ppamfactory import *
factory = PPAMFactory()
print(f'Is 64b windows? \t{factory.is64bwin}')

# creates without ribbon components
factory.create(r'c:\debug\test', r'c:\debug\test\output.ppam', r'c:\debug\test\copy.ppam')

# creates with ribbon + JPEG
factory.create(r'c:\debug\test', r'c:\debug\test\output.ppam', r'c:\debug\test\copy.ppam', r'c:\debug\test\ribbon.xml', r'c:\debug\test\image.jpg')
"""
