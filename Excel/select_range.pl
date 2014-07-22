use warnings;
use strict;

use Win32::Ole;

use File::Spec qw(tmpdir);

my $excel = CreateObject Win32::OLE 'Excel.Application' or die;
$excel->{'Visible'} = 1;

my $workbook = $excel -> Workbooks -> Add(1);

my $sheet = $excel -> Sheets(1);

my $row_start =  5;
my $col_start =  3;

my $row_end   = 11;
my $col_end   =  6;

my $top_left_cell     = $sheet -> Cells ($row_start, $col_start);
my $bottom_right_cell = $sheet -> Cells ($row_end  , $col_end  );

$workbook -> Sheets(1) -> Range ($top_left_cell, $bottom_right_cell) -> Select;

$workbook -> {Saved} = 1;
