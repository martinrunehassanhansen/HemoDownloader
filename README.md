# HemoDownloader
HemoDownloader is a third-party utility for downloading biochemical results from HemoCue HbA1c 501 devices (see https://www.hemocue.com/en/solutions/diabetes/hemocue-hba1c-501-system) over a RS232 null-modem cable.

If you use HemoDownloader in your academic work, please cite both this GitHub repository (https://github.com/martinrunehassanhansen/HemoDownloader) and the following paper:

Hansen MRH, Schlünssen V, and Sandbæk A. "HemoDownloader: Open source software utility to extract data from HemoCue HbA1c 501 devices in epidemiological studies of diabetes mellitus." Plos one 15.11 (2020): e0242087. https://doi.org/10.1371/journal.pone.0242087

# Disclaimer
PLEASE NOTE THAT HEMODOWNLOADER IS INTENDED FOR USE IN EPIDEMIOLOGICAL STUDIES ONLY. THE SOFTWARE IS *NOT* APPROVED AS A MEDICAL DEVICE AND MUST NOT BE USED AS SUCH. THAT MEANS THE SOFTWARE MUST NOT BE USED ON HUMAN BEINGS FOR DIAGNOSIS, PREVENTION, MONITORING, PREDICTION, PROGNOSIS, TREATMENT OR ALLEVIATION OF DISEASE OR ANY OTHER MEDICAL PURPOSES.

HemoCue® is a registered trademark of HemoCue AB (Ängelholm, Sweden). This software is not produced nor endorsed by HemoCue AB.

# Author information
HemoDownloader was created by Martin Rune Hassan Hansen, Aarhus University, Denmark. For any questions, contact martinrunehassanhansen@ph.au.dk

# License
This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program.  If not, see <https://www.gnu.org/licenses/>.

# How to run HemoDownloader from Python source
HemoDownloader is written for Python 3 that is freely available from www.python.org/downloads/

During Python installation, make sure you include the optional feature 'pip'.

Once you have installed Python 3 with pip, you have to install the following three third-party Python modules using pip (see https://packaging.python.org/tutorials/installing-packages/#use-pip-for-installing):
* pyserial (https://pypi.org/project/pyserial/)
* XlsxWriter (https://pypi.org/project/XlsxWriter/)
* xlwt (https://pypi.org/project/xlwt/)

When all dependencies are installed, you can start the graphical user interface by exceuting the file HemoDownloader.pyw

Futher instructions can be found in the menu 'Help' → 'Instructions for use' in the graphical user interface.
