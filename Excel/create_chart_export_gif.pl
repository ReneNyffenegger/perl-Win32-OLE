use warnings;
use strict;

use Win32::Ole;
use Win32::OLE::Const 'Microsoft.Excel';

use Cwd qw(getcwd);
use File::Spec;

$Win32::OLE::Warn = 3;

my $excel = CreateObject Win32::OLE 'Excel.Application' or die;
$excel->{'Visible'} = 1;

my $workbook = $excel -> workbooks -> add(1);
my $sheet    = $workbook -> sheets(1);
my $shape    = $sheet    -> shapes -> addChart;
my $chart    = $shape    -> chart;

$chart -> {chartType} = xlXYScatterSmoothNoMarkers;

$sheet -> cells(1,1) -> {value} = "x";
$sheet -> cells(1,2) -> {value} = "sin(x)";

my $row = 1;
for my $x (map {$_ / 10} (0..100)) {
  $row ++;
  $sheet -> cells($row, 1) ->{value} =     $x;
  $sheet -> cells($row, 2) ->{value} = sin($x);
}

$chart -> setSourceData($sheet->range($sheet->cells(1,2), $sheet->cells(100,2)));

$chart -> SeriesCollection(1) -> {XValues} = '=Sheet1!$A$2:$A$102';
#
# Following doesn't work, fails with 
#   Win32::OLE(0.1709) error 0x80020003:
#     in PROPERTYPUTREF "XValues"
#
# $chart -> SeriesCollection(1) -> {XValues} = $sheet -> Range($sheet->cells(2,1), $sheet->cells(2, $row));
#
#  -> http://stackoverflow.com/questions/25280606
#

$chart -> Export(File::Spec->canonpath(getcwd) . "\\sin.gif");  # Export must not be spelled «export»!

$workbook -> {saved} = 1;
