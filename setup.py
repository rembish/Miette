#!/usr/bin/env python
import os
import sys
import miette as miette

try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

if sys.argv[-1] == 'publish':
    os.system('python setup.py sdist upload')
    sys.exit()

packages = [
    'miette',
    'miette.doc',
    'miette.cfb',
    'miette.tools'
]

requires = []

setup(
    name='miette',
    version=miette.__version__,
    description='Miette is a light-weight Microsoft Office documents reader',
    author='Alex Rembish',
    author_email='alex@rembish.ru',
    packages=packages,
    #package_data={'': ['LICENSE'],},
    package_dir={'miette': 'miette'},
    include_package_data=True,
    install_requires=requires,
    #license=open('LICENSE').read(),
    zip_safe=False,
    classifiers=(
        'Development Status :: 5 - Production/Stable',
        'Intended Audience :: Developers',
        'Natural Language :: English',
        'Programming Language :: Python',
        'Programming Language :: Python :: 2.6',
        'Programming Language :: Python :: 2.7',
    ),
)