package Foswiki::Plugins::ExportExcelPlugin;

use strict;
use warnings;

use Foswiki::Func ();
use Foswiki::Plugins ();

use Excel::Writer::XLSX;
use File::Temp;
use URL::Encode;

use Data::Dumper;

use utf8;

use version;
our $VERSION = version->declare( '1.0.0' );
our $RELEASE = '1.0.0';
our $NO_PREFS_IN_TOPIC = 1;
our $SHORTDESCRIPTION = 'Provides table exports to Excel 2007+ format.';

sub initPlugin {
  my ( $topic, $web, $user, $installWeb ) = @_;

  if ( $Foswiki::Plugins::VERSION < 2.0 ) {
      Foswiki::Func::writeWarning( 'Version mismatch between ',
          __PACKAGE__, ' and Plugins.pm' );
      return 0;
  }

  my $context = Foswiki::Func::getContext();
  if ( $context->{'edit'} ) {
    return 1;
  }

  my $path = '%PUBURLPATH%/%SYSTEMWEB%/ExportExcelPlugin';
  my $script = "<script type=\"text/javascript\" src=\"$path/scripts/export_excel.js\"></script>";
  Foswiki::Func::addToZone( 'script', 'EXPORT::EXCEL::SCRIPTS', $script, 'JQUERYPLUGIN::FOSWIKI' );

  my $style = "<link rel=\"stylesheet\" type=\"text/css\" media=\"all\" href=\"$path/styles/export_excel.css\" />";
  Foswiki::Func::addToZone( 'head', 'EXPORT::EXCEL::STYLES', $style );

  Foswiki::Func::registerRESTHandler( 'convert', \&_restConvert, authenticate => 0, http_allow => 'POST' );
  Foswiki::Func::registerRESTHandler( 'get', \&_restGet, authenticate => 0, http_allow => 'GET' );
  return 1;
}

sub _restConvert {
  my ( $session, $subject, $verb, $response ) = @_;
  my $query = $session->{request};

  my $param = $query->{param}->{table}[0];

  my $tmpDir = Foswiki::Func::getWorkArea( 'ExportExcelPlugin' );
  my $xlsxFile = new File::Temp( DIR => $tmpDir, SUFFIX => '.xlsx', UNLINK => 0 );

  $xlsxFile =~ m/$tmpDir\/(.+)/;
  my $attachment = $1;
  $response->header( -status  => 200 );

  my $fh;
  open $fh, "> $xlsxFile";
  binmode $fh;

  my $workbook  = Excel::Writer::XLSX->new( $fh );
  my $worksheet = $workbook->add_worksheet();

  my $header = $workbook->add_format();
  $header->set_format_properties(
    bold => 1,
    size => 12,
    bg_color => '#cccccc',
    color => 'black' );

  if ( $param ) {
    my @rows = split( "\n", $param );
    my ( $i, $j ) = ( 0, 0 );
    for my $row (@rows) {
      my @cols = split( ";", $row );
      for my $col (@cols) {
        my $value = URL::Encode::url_decode_utf8( $col );
        if ( $value =~ /TH:(.+)/ ) {
          $worksheet->write( $i, $j, $1, $header );
        } else {
          $worksheet->write( $i, $j, $value );
        }

        $j = $j + 1;
      }

      $i = $i + 1;
      $j = 0;
    }
  }
  else {
    $worksheet->write( 0, 0, "ToDo -> Fehler bearbeiten" );
  }
  $workbook->close();
  close $fh;

  return $attachment;
}

sub _restGet {
  my ( $session, $subject, $verb, $response ) = @_;
  my $query = $session->{request};
  my $filename = $query->{param}->{filename}[0];

  unless ( $filename ) {
    Foswiki::Func::writeWarning( "Invalid file: $filename." );
    return;
  }

  my $tmpDir = Foswiki::Func::getWorkArea( 'ExportExcelPlugin' );
  my $attachment = "$tmpDir/$filename";
  $response->header(
    -status  => 200,
    -type    => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  );

  my $name = $session->{webName} . '.' . $session->{topicName} . '.xlsx';
  $response->pushHeader( "Content-Disposition", "inline; filename=\"$name\"" );

  my $file;
  open $file, "< $attachment";
  while( <$file> ) {
    $response->print( $_ );
  }

  unless ( unlink $attachment ) {
    Foswiki::Func::writeWarning( "Unable to delete temporary Excel file: $attachment." );
  }

  return;
}

1;

__END__
Foswiki - The Free and Open Source Wiki, http://foswiki.org/

Author: Sven Meyer <meyer@modell-aachen.de>

Copyright (C) 2008-2013 Foswiki Contributors. Foswiki Contributors
are listed in the AUTHORS file in the root of this distribution.
NOTE: Please extend that file, not this notice.

This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version. For
more details read LICENSE in the root of this distribution.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.

As per the GPL, removal of this notice is prohibited.
