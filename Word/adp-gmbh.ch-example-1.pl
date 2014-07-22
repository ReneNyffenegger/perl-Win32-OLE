#
#   First example from http://www.adp-gmbh.ch/perl/word.html
#
use warnings;
use strict;

use Win32::OLE;

my $word = CreateObject Win32::OLE 'Word.Application' or die $!;
$word->{'Visible'} = 1;

my $document = $word->Documents->Add;

my $selection = $word->Selection;

$selection -> TypeText("Hello World");
$selection -> TypeParagraph;
$selection -> TypeText("How do you feel today");
$selection -> TypeParagraph;

$selection -> TypeText("Some header");
$selection -> {'Style'} = "Heading 1";

$selection -> TypeParagraph;

my $heading_1 = $document->Styles("Heading 1");
my $heading_1_font = $heading_1 -> Font;

$heading_1_font -> {Name} = "Bookmann";
$heading_1_font -> {Size} = 20;
$heading_1_font -> {Bold} =  1;
