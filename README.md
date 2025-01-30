# Analyzer MEA SpAnner
Analysis of a Synopsis output from SpAnner

## Description
This python script analyses relevant output data from all synopsis files generated
with SpAnner placed in the input directory. The user has to enter some further
information. Then the analysis takes place and an output xlsx file is saved in the
output directory.

## Usage
1. Clone the repository
2. Install all Prerequisites
   - [Python 3](https://www.python.org/downloads/) (Python 3.6+)
   - [pandas](https://pandas.pydata.org/docs/getting_started/install.html)
   - [openpyxl](https://openpyxl.readthedocs.io/en/stable/tutorial.html)
   - [scipy](https://scipy.org/install/)
5. Place all xlsx symposis files you wish to analyse with the Analyzer MEA SpAnner into the folder _input_.
6. Start the python script (if you use windows, simply open the file "run.bat"). The script will now analyze all .xslx files located in the _input_ folder.
8. The analysis output files will be saved into the folder _output_.

## Licence
This script can be used following the MIT Licence.

## Acknowledgments
This script uses the publically available python modules pandas and scipy.
