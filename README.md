# tm35wgs84
Convert TM35 coordinates to WGS84, for more info about coordinate systems see https://epsg.io/3067

# Perl script
This simple script was developed to use with an excel file with TM35 coordinates N and E in two columns. By default the script assumes the file is named 'R1_byggnad.xls', just rename it according to your needs. Also make sure you have free columns to the right of the ones with TM35-coordinates. The script will overwrite values in colums name n and n+1, where n is a number you can configure in the script. Sorry for being lazy but this can be polished.

# Dependencies
Spreadsheet::ParseExcel
Geo::Coordinates::ETRSTM35FIN
Data::Dumper

# Run
Run with command "perl tm35wgs84.pl"
