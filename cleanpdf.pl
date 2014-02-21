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
unless (-f $path_src) {
    print STDERR "Specify PDF file. $path_src\n";
    exit 1;
}
my $path_dest = "$path_src.dst.pdf";
my $path_temp = "$path_src.tmp.pdf";

my $pdf_src = eval { PDF::API2->open($path_src); };
unless ($pdf_src) { # Try to convert version via CAM::PDF
    my $campdf = eval { CAM::PDF->new($path_src); };
    unless ($campdf) {
        print STDERR "<CAM::PDF> Can't open. $path_src\n";
        exit 1;
    }
    # toPDF() and openScalar() can't pass to PDF::API2.
    $campdf->cleanoutput($path_temp);
    $pdf_src = eval { PDF::API2->open($path_temp); };
    unless ($pdf_src) {
        print STDERR "<PDF::API2> Can't open. $path_temp\n";
        unlink $path_dest, $path_temp;
        exit 1;
    }
}

my $pages_count = $pdf_src->pages();
unless ($pages_count) {
    print STDERR "<PDF::API2> No page.\n";
    unlink $path_dest, $path_temp;
    exit 1;
}

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
unless (rename $path_dest, $path_src) {
    print STDERR "rename ERROR: $!\n";
    unlink $path_dest, $path_temp;
    exit 1;
}

exit;
