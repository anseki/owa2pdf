#!/usr/bin/perl

# cleanpdf.pl
# owa2pdf
# https://github.com/anseki/owa2pdf
#
# Copyright (c) 2014 anseki
# Licensed under the MIT license.

use strict;
use warnings;

use PDF::API2;
use CAM::PDF;

my $path_src = $ARGV[0];
_error("Specify PDF file. $path_src") unless -f $path_src;
my $path_dest = "$path_src.dst.pdf";
my $path_temp = "$path_src.tmp.pdf";

my $pdf_src = eval { PDF::API2->open($path_src); };
unless ($pdf_src) { # Try to convert version via CAM::PDF
    my $campdf = eval { CAM::PDF->new($path_src); } or _error("<CAM::PDF> Can't open. $path_src");
    # toPDF() and openScalar() can't pass to PDF::API2.
    $campdf->cleanoutput($path_temp);
    $pdf_src = eval { PDF::API2->open($path_temp); } or _error("<PDF::API2> Can't open. $path_temp");
}

my $pages_count = $pdf_src->pages() or _error('<PDF::API2> No page.');
my $pdf_dest = PDF::API2->new(-onecolumn => 1);
delete $pdf_dest->{pdf}->{Info};
delete $pdf_dest->{catalog}->{PageLayout};
foreach my $page_num (1 .. $pages_count) {
    $pdf_dest->importpage($pdf_src, $page_num);
}

$pdf_dest->saveas($path_dest);
$pdf_src->end();
$pdf_dest->end();

unlink $path_src;
rename $path_dest, $path_src or _error("rename ERROR: $!");

exit;

sub _error {
    my $msg = shift;
    print STDERR "$msg\n";
    unlink grep $_ && -f $_, ($path_dest, $path_temp);
    exit 1;
}

