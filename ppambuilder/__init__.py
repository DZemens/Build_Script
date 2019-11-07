__author__ = 'dz'
__all__ = ['ppamfactory', 'os_version']


"""
example:

from ppambuilder.ppamfactory import *
factory = PPAMFactory()
print(f'Is 64b windows? \t{factory.is64bwin}')
factory.create(r'c:\debug\test', '', '', r'c:\debug\test\output.ppam', r'c:\debug\test\copy.ppam', False)
"""
