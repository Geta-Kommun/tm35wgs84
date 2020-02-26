#!/usr/bin/perl -w
 
use strict;
use warnings;
#use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::SaveParser;
use Geo::Coordinates::ETRSTM35FIN;
use Data::Dumper;
		 
my $gce = new Geo::Coordinates::ETRSTM35FIN;

# Open an existing file with SaveParser
my $parser   = Spreadsheet::ParseExcel::SaveParser->new();
my $workbook = $parser->Parse('R1_byggnad.xls');

#print Dumper(ref($workbook));

if ( !defined $workbook ) {
    die $parser->error(), ".\n";
}

my $NCol = 1; 		# Column: Easting is assumed to be Northing + 1
my $LATCOL = 4; 	# Column: Longitude is assumed to be latitude + 1 


my $worksheet = $workbook->worksheet(0); 



my ( $row_min, $row_max ) = $worksheet->row_range();
my ( $col_min, $col_max ) = $worksheet->col_range();

for my $row ( $row_min .. $row_max ) {
      
      my $north = $worksheet->get_cell( $row, $NCol );
      my $east = $worksheet->get_cell( $row, $NCol+1 );
      next unless ($north && $east);
      
      # ETRS-TM35FIN coordinates to WGS84 coordinates
      my ( $lat, $lon ) = $gce->ETRSTM35FINxy_to_WGS84lalo($north->value(), $east->value());
      $worksheet->AddCell( $row, $LATCOL, $lat);
      $worksheet->AddCell( $row, $LATCOL+1, $lon);
            
      print "Row $row: (", $north->value(),", ", $east->value(), "), ==> (", $lat,", ",$lon, ") \n";
      
}
# Write over the existing file or write a new file.
$workbook->SaveAs('R1_byggnad.xls');