import os
from setuptools import setup

README = open(os.path.join(os.path.dirname(__file__), 'README.md')).read()
LICENSE = open(os.path.join(os.path.dirname(__file__), 'LICENSE.txt')).read()

# allow setup.py to be run from any path
os.chdir(os.path.normpath(os.path.join(os.path.abspath(__file__), os.pardir)))

setup(
    name='django-spreadsheetresponsemixin',
    version='0.1.6',
    packages=['spreadsheetresponsemixin'],
    include_package_data=True,
    license=LICENSE,
    description='An mixin for views with a queryset that provides a CSV/Excel export.',
    long_description=README,
    url='https://github.com/birdsarah/django-spreadsheetresponsemixin',
    author='Sarah Bird',
    author_email='sarah@bonvaya.com',
    install_requires=['django>=1.5', 'openpyxl>=1.6.2'],
    classifiers=[
                'Development Status :: 3 - Alpha',
                'Environment :: Web Environment',
                'Framework :: Django',
                'Intended Audience :: Developers',
                'License :: OSI Approved :: GNU General Public License v3 (GPLv3)',
                'Operating System :: OS Independent',
                'Programming Language :: Python',
                'Programming Language :: Python :: 2.7',
                'Topic :: Internet :: WWW/HTTP',
                'Topic :: Internet :: WWW/HTTP :: Dynamic Content'],
)
