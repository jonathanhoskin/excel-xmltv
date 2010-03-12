## Description

This is a parser to convert program guide data from MS Excel spreadsheets to XMLTV compliant xml.

Given the schema followed, this will only ever really work with the spreadsheets created by Triangle TV in New Zealand.

The hope of this project is to provide a simple script to do one-line conversion. e.g;

ruby excel-xmltv.rb triangle.xls > triangle.xml


## Dependecies

Requires Ruby, RebyGems. Also requires the 'builder' and 'roo' gems.


## See also

XMLTV Wiki: http://wiki.xmltv.org 

Triangle TV: http://www.tritv.co.nz/
