##  setup.py
##  A script to install dependencies for HemoDownloader 1.2

##  Copyright © 2019 Martin Rune Hassan Hansen <martinrunehassanhansen@ph.au.dk>

##  This program is free software: you can redistribute it and/or modify
##  it under the terms of the GNU General Public License as published by
##  the Free Software Foundation, either version 3 of the License, or
##  (at your option) any later version.

##  This program is distributed in the hope that it will be useful,
##  but WITHOUT ANY WARRANTY; without even the implied warranty of
##  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
##  GNU General Public License for more details.

##  You should have received a copy of the GNU General Public License
##  along with this program.  If not, see <https://www.gnu.org/licenses/>.

try:
    from setuptools import setup
except:
    from distutils.core import setup

setup(name='HemoDownloader.pyw',
      version='1.2',
      description='A GUI utility for downloading data from HemoCue® HbA1c devices',
      author='Martin Rune Hassan Hansen',
      author_email='martinrunehassanhansen@ph.au.dk',
      install_requires=['pyserial>=3.4','xlsxwriter>=1.1.8','xlwt>=1.3.0'],
     )
