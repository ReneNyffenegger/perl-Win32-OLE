use warnings;
use strict;

use Win32::Ole;

my $excel = CreateObject Win32::OLE 'Excel.Application' or die;
$excel->{'Visible'} = 1;

my $workbook = $excel -> Workbooks -> Add(5);

$workbook -> Sheets(1) -> Cells(5,2)  -> {Value} = "2nd column of 5th row on 1st sheet";
$workbook -> Sheets(2) -> Cells(2,1)  -> {Value} = "2nd row, first column";
$workbook -> Sheets(3) -> Cells(4,2)  -> {Value} = "4th row, 2nd column";

$workbook -> {Saved} = 1;
