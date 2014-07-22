use warnings;
use strict;

use Win32::Ole;

my $excel = CreateObject Win32::OLE 'Excel.Application' or die;
$excel->{'Visible'} = 1;

my $workbook = $excel -> Workbooks -> Add(1);

my $sheet = $excel -> Sheets(1);

my $row = 1;

fill_cell($row++, "#.00", 4242.4242);
fill_cell($row++, "#"   , 4242.4242);
fill_cell($row++, "@"   , 4242.4242);


sub fill_cell { # {{{
    my $row    = shift;
    my $format = shift;
    my $value  = shift;

    my $cell;

    $cell = $sheet -> Cells ($row, 1);
    $cell -> {NumberFormat} = '@';
    $cell -> {Value       } =  $format;

    $cell = $sheet -> Cells ($row, 2); 
    $cell -> {NumberFormat} = $format; 
    $cell -> {Value}        = $value;
} # }}}

$workbook -> {Saved} = 1;
