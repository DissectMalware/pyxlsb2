import os.path
from setuptools import setup
from pyxlsb2 import __version__

# Get a handy base dir
project_dir = os.path.abspath(os.path.dirname(__file__))

with open(os.path.join(project_dir, 'README.rst')) as f:
    README = f.read()

setup(
    name='pyxlsb2',
    version=__version__,

    description='Excel 2007+ Binary Workbook (xlsb) parser',
    long_description=README,

    author='DissectMalware',
    author_email='amir@inquest.net',

    url='https://github.com/DissectMalware/pyxlsb2',

    license='MIT',

    classifiers=[
        'Development Status :: 5 - Production/Stable',

        'License :: OSI Approved :: MIT License',

        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
    ],

    packages=['pyxlsb2'],


    zip_safe=False
)
