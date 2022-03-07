#!/usr/bin/env python

import os
import sys

from setuptools import setup, find_packages

if sys.argv[-1] == 'publish':
    os.system('python setup.py sdist upload')
    sys.exit()

with open('README') as readmeFile:
    long_desc = readmeFile.read()

setup(
    name='miette',
    version='1.5',
    description='Miette is a light-weight Microsoft Office documents reader',
    long_description=long_desc,
    url='https://github.com/rembish/Miette',
    author='Alex Rembish',
    author_email='alex@rembish.org',
    packages=find_packages(),
    install_requires=['cfb'],
    license='BSD 2-Clause license',
    classifiers=(
        'Development Status :: 5 - Production/Stable',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: BSD License',
        'Natural Language :: English',
        'Programming Language :: Python',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Topic :: Software Development :: Libraries',
        'Topic :: Software Development :: Libraries :: Python Modules',
        'Topic :: Text Processing'
    ),
)
