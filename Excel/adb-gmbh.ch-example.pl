#
#   Example from http://www.adp-gmbh.ch/perl/excel.html
#
use warnings;
use strict;

use Win32::Ole;

use File::Spec qw(tmpdir);

my $excel = CreateObject Win32::OLE 'Excel.Application' or die $!;
$excel->{'Visible'} = 1;

my $workbook = $excel -> Workbooks -> Add;

$workbook -> ActiveSheet -> Range('A1')->{'Value'} = "Hello";

$workbook -> ActiveSheet -> Range('C2:D3')->{'Value'} = [ 
  ['one',   'two' ],
  ['three', 'four'],
];

my $save_as_name = File::Spec->tmpdir . '\\perl.xlsx';;
$workbook -> SaveAs ($save_as_name) or die "Could not save Excelfile to $save_as_name";

$excel -> Quit;

print "Excelfile written to $save_as_name\n";
