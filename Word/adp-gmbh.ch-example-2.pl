#   Second example from http://www.adp-gmbh.ch/perl/word.html
#
#   The following functions are defined but not called:
#      style_keep_with_next
#      style_keep_together
#      style_border>
#      style_tab_at_position
#      style_space_before
#      style_space_after
#      style_alignment
#      goto_end_of_document 


use warnings;
use strict;

use Win32::OLE;
use Win32::OLE::Const 'Microsoft.Word';    # wd  constants
use Win32::OLE::Const 'Microsoft Office 14.0 Object Library';  # mso constants

my $cur_style    = 'a';
my $cur_bookmark = 'a';

my $word = CreateObject Win32::OLE 'Word.Application' or die $!;
$word->{'Visible'} = 1;

my $document = $word->Documents->Add;

# selection is the insertion point.
my $selection = $word->Selection;

my $st_bold_10_arial   = create_style($document, "Arial", 10, 1, 0);
my $st_italic_10_arial = create_style($document, "Arial", 10, 0, 1);
my $st_italic_20_arial = create_style($document, "Times new Roman", 20, 0, 1);

set_style($document, $st_bold_10_arial);
text     ($document, "This text is bold, 10 points and arial");
enter    ($document);
set_style($document, $st_italic_10_arial);
text     ($document, "This text is italic, 10 points and arial");
enter    ($document);
set_style($document, $st_italic_20_arial);
text     ($document, "This text is italic, 20 points and Times");

# insert something into the header
switch_view($document, wdSeekCurrentPageHeader);
my $st_header = create_style($document, "Verdana", 8, 0, 0);
set_style($document, $st_header);
text     ($document, "this is the document's header, it will be repeated on all pages");

# insert something into the footer
switch_view($document, wdSeekCurrentPageFooter);
my $st_footer = create_style($document, "Verdana", 8, 0, 0);
set_style($document, $st_footer);
text     ($document, "this is the document's footer, it will be repeated on all pages");

# Inserting current page, total page
text     ($document, " Page ");
$selection->Fields->Add ($selection->{Range}, wdFieldPage    );
text     ($document, " of ");
$selection->Fields->Add ($selection->{Range}, wdFieldNumPages);

# go back to body of document
switch_view($document, wdSeekMainDocument);

my $picture = insert_picture($document, 'c:\some\path\to\the\file.eps', 100, 100, 100, 30);

my $st_12_arial = create_style($document, "Arial", 12, 0, 0);



#############################   Demonstrating Bookmarks #####
enter    ($document);                                       #
set_style($document, $st_12_arial);                         #
text     ($document, "Not bold");                           #
bold     ($document, 1);                                    #
text     ($document, " but this IS bold ");                 #
bold     ($document, 0);                                    #
text     ($document, "this is not bold anymore");           #
enter    ($document);                                       #
                                                            #
my $first_bookmark = insert_bookmark($document);            #
text     ($document, "here's a bookmark");                  #
enter    ($document);                                       #
text     ($document, "Writing some lines");                 #
enter    ($document);                                       #
text     ($document, "Before going to the bookmark");       #
enter    ($document);                                       #
                                                            #
my $second_bookmark = insert_bookmark($document);           #
goto_bookmark($document, $first_bookmark);                  #
text     ($document, "Went back to first bookmark");        #
goto_bookmark($document, $second_bookmark);                 #
text     ($document, "Went back to second bookmark");       #
#############################################################


############### indenting ##############

my $st_indent = create_style($document, "Arial", 10, 0, 0);
style_indents($document, $st_indent, 100, 200);
enter        ($document);
set_style    ($document, $st_indent);
text         ($document, join " ", 1 .. 100); 
enter        ($document);
text         ($document, join " ", 'a' .. 'cz'); 

items        ($document, "Fruits", $st_indent, qw(Apple Pear Peach Cherry Nectarine Orange)); 
items        ($document, "Shapes", $st_indent, qw(Triangle Rhombus Trapezoid Pentagon Circle)); 

## inserting a box
my $st_box  = create_style($document, "Arial", 8, 0, 1);
style_indents($document, $st_box, 0, 0);
my $box = insert_box($document, 350, 550, 140, 40);
set_style($document, $st_box);
text         ($document, "hello out of the box");

save_doc_as($document, 'c:\generated.doc');

## uncomment the following two if word should shut down
# close_doc($document);
# $word->Quit;

sub text {
  my $document = shift;
  my $text      = shift;

  $document->ActiveWindow->Selection -> TypeText($text);
}

# aka new line, newline or NL
sub enter {
  my $document = shift;

  $document->ActiveWindow->Selection -> TypeParagraph;
}

sub set_style {
  my $document = shift;
  my $style_arg = shift;

  $document->ActiveWindow->Selection -> {Style} = $style_arg -> {name};
}

sub create_style {
  my $document  = shift;
  my $fontname  = shift;
  my $font_size = shift;
  my $bold      = shift;
  my $italic    = shift;

  my $style = $document->Styles->Add($cur_style);

  my $style_font = $style->{Font};

  $style_font -> {Name  } = $fontname;
  $style_font -> {Size  } = $font_size;
  $style_font -> {Bold  } = $bold;
  $style_font -> {Italic} = $italic;

  my %style;

  $style{name} = $cur_style++;
  return \%style;
}

# use switch_view to change to header, footer, main document and so on...
# possible constants for view are: wdSeekCurrentPageFooter 
#
#   o  wdSeekCurrentPageHeader 
#   o  wdSeekEndnotes 
#   o  wdSeekEvenPagesFooter 
#   o  wdSeekEvenPagesHeader 
#   o  wdSeekFirstPageFooter 
#   o  wdSeekFirstPageHeader 
#   o  wdSeekFootnotes 
#   o  wdSeekMainDocument 
#   o  wdSeekPrimaryFooter 
#   o  wdSeekPrimaryHeader 
#
sub switch_view {
  my $document = shift;
  my $view     = shift;
  $document -> ActiveWindow -> ActivePane -> View -> {SeekView} = $view;
}

sub insert_picture {
  my $document = shift;
  my $file     = shift;
  my $left     = shift;
  my $top      = shift;
  my $width    = shift;
  my $height   = shift;

  my $picture = 
    $document-> Shapes -> AddPicture (
      $file, 
      msoFalse, # link to file
      msoTrue,  # save with document
      $left, $top, $width, $height, 
      $document->ActiveWindow->Selection->{Range}
    );

  return $picture;
}

sub bold {
  my $document = shift;
  my $bold     = shift;

  $document->ActiveWindow->Selection->{Font}->{Bold} = $bold ? msoTrue : msoFalse;
}

sub goto_bookmark {
  my $document = shift;
  my $bookmark = shift;

  $document->ActiveWindow->Selection -> GoTo(wdGoToBookmark, 0, 0, $bookmark->{Name});
}

sub insert_bookmark {
  my $document = shift;

  my $bookmark = $document -> Bookmarks -> Add ($cur_bookmark++, $document->ActiveWindow->Selection->Range);
  return $bookmark;
}                      

sub style_indents {
  my $document          = shift;
  my $style_arg         = shift;
  my $first_line_indent = shift;
  my $other_line_indent = shift;

  my $style = $document->Styles($style_arg->{name});

  $style->ParagraphFormat->{LeftIndent     } =  $other_line_indent;
  $style->ParagraphFormat->{FirstLineIndent} = -$other_line_indent + $first_line_indent;
}

sub items {
  my $document = shift;
  my $title    = shift;
  my $style    = shift;
  my @array    = @_;

  enter($document);
  set_style($document, $style);

  bold($document, 1);
  text($document, $title);
  bold($document, 0);
  text($document, "\x09");

  foreach my $a (@array) {
    text($document, $a);
    text($document, "\x0b");
  }
}

sub insert_box {
  my $document = shift;
  my $left     = shift;
  my $top      = shift;
  my $width    = shift;
  my $height   = shift;

  my $shape = $document->Shapes->AddTextbox(msoTextOrientationHorizontal, $left, $top, $width, $height);
  $shape -> Select;
  my $selection = $word->Selection;
  $selection -> ShapeRange -> Line -> {DashStyle} = msoLineRoundDot;

  return $shape
}

sub close_doc {
  my $document = shift;
  $document -> Close;
}

sub save_doc_as {
  my $document = shift;
  my $filename = shift;

  $document->SaveAs($filename);
}

sub style_keep_with_next {
  my $document     = shift;
  my $style_arg    = shift;

  my $style = $document->Styles($style_arg->{name});

  $style->{ParagraphFormat}->{KeepWithNext} = msoTrue;
}

sub style_keep_together {
  my $document     = shift;
  my $style_arg    = shift;

  my $style = $document->Styles($style_arg->{name});

  $style->{ParagraphFormat}->{KeepTogether} = msoTrue;
}

sub style_border {
  my $document     = shift;
  my $style_arg    = shift;
  my $border       = shift; 
  my $border_style = shift;
  my $border_width = shift;
  my $border_color = shift;

  my $style = $document->Styles($style_arg->{name});

  $style->Borders($border) -> {LineStyle} = $border_style;
  $style->Borders($border) -> {LineWidth} = $border_width;
  $style->Borders($border) -> {Color    } = $border_color;
}

sub style_tab_at_position {
  my $document     = shift;
  my $style_arg    = shift;
  my $position     = shift;
  my $left_or_right= shift;

  my $style = $document->Styles($style_arg->{name});

  $style->ParagraphFormat->{TabStops}->Add($word->InchesToPoints($position), $left_or_right);
}

sub style_space_before {
  my $document     = shift;
  my $style_arg    = shift;
  my $space        = shift; 

  my $style = $document->Styles($style_arg->{name});

  $style->ParagraphFormat->{SpaceBefore} = $space;
}

sub style_space_after {
  my $document     = shift;
  my $style_arg    = shift;
  my $space        = shift; 

  my $style = $document->Styles($style_arg->{name});

  $style->ParagraphFormat->{SpaceAfter} = $space;
}

sub style_alignment {
  my $document     = shift;
  my $style_arg    = shift;
  my $alignment    = shift;

  my $style = $document->Styles($style_arg->{name});

  $style->ParagraphFormat->{Alignment} = $alignment;
}

sub goto_end_of_document {
  my $document  = shift;

  $document->ActiveWindow->Selection->{Range} -> EndKey(wdStory);

  #my $selection = $word->Selection;

  #$selection -> EndKey (wdStory);
}

sub insert_page_break {
  my $document  = shift;

  #my $selection = $word->Selection;

  #$selection -> InsertBreak(wdPageBreak);
  $document->ActiveWindow->Selection->{Range} -> InsertBreak(wdPageBreak);
}

sub landscape {
  my $document = shift;

  $document->PageSetup->{Orientation} = wdOrientLandscape;
}
