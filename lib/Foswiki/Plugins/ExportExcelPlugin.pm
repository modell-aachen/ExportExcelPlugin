package Foswiki::Plugins::ExportExcelPlugin;

use strict;
use warnings;

use Foswiki::Func ();
use Foswiki::Plugins ();

use Excel::Writer::XLSX;
use File::Temp;
use JSON;

use utf8;

use version;
our $VERSION = version->declare( '1.0.0' );
our $RELEASE = "1.0";
our $NO_PREFS_IN_TOPIC = 1;
our $SHORTDESCRIPTION = "Provides table exports to Excel .xlsx format.";

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
  my $script = "<script type=\"text/javascript\" src=\"$path/scripts/export_excel.js?v=$RELEASE\"></script>";
  Foswiki::Func::addToZone( 'script', 'EXPORT::EXCEL::SCRIPTS', $script, 'JQUERYPLUGIN::FOSWIKI' );

  my $style = "<link rel=\"stylesheet\" type=\"text/css\" media=\"all\" href=\"$path/styles/export_excel.css?v=$RELEASE\" />";
  Foswiki::Func::addToZone( 'head', 'EXPORT::EXCEL::STYLES', $style );

  Foswiki::Func::registerRESTHandler( 'export', \&_restExport, authenticate => 1, validate => 0, http_allow => 'POST' );
  Foswiki::Func::registerRESTHandler( 'get', \&_restGet, authenticate => 1, validate => 0, http_allow => 'GET' );
  Foswiki::Func::registerTagHandler( 'XLSXLINK', \&_tagExport );

  my $classes = $Foswiki::cfg{Plugins}{ExportExcelPlugin}{Classes} || '';
  my $stats = $Foswiki::cfg{Plugins}{ExportExcelPlugin}{AllowExportWebStatistics} || 0;
  my $hover = $Foswiki::cfg{Plugins}{ExportExcelPlugin}{HoverMode};
  $hover = 1 unless defined $hover;
  my $useHover = $hover ? 'true' : 'false';
  if ( $classes ) {
    Foswiki::Func::addToZone(
    "script",
    "EXPORTEXCELPLUGIN::EXTENSIONS",
    "<script type='text/javascript'>jQuery.extend( foswiki.preferences, { \"excelExport\": { \"classes\": \"$classes\", \"webstatistics\": \"$stats\", \"useHover\": $useHover } } );</script>",
    "EXPORT::EXCEL::SCRIPTS" );
  }

  return 1;
}

sub _tagExport {
  my( $session, $params, $topic, $web, $topicObject ) = @_;

  my $selector = $params->{_DEFAULT} || $params->{selector} || '';
  my $text = $params->{text} || '%MAKETEXT{"Export to Excel"}%';
  my $columns = $params->{columns} || '';
  my $cls = $params->{extraclasses} || '';
  $cls =~ s/,/ /g;

  return '' unless $selector;
  return <<LINK;
<a class="xlsxlink $cls" href="#" data-selector="$selector" data-columns="$columns">
$text
</a>
LINK
}

sub _restExport {
  my ( $session, $subject, $verb, $response ) = @_;
  my $q = $session->{request};
  my $payload = $q->param('payload') || '{}';
  my $r = from_json($payload);
  my $web = $r->{web};
  my $topic = $r->{topic};

  unless (scalar(@{$r->{data}})) {
    $response->header( -status  => 400 );
    return to_json({
        status => 'error',
        msg => "Received invalid table data!"
    });
  }

  my $file = new File::Temp(
    DIR => Foswiki::Func::getWorkArea('ExportExcelPlugin'),
    SUFFIX => '.xlsx',
    UNLINK => 0
  );

  my $xlsx  = Excel::Writer::XLSX->new($file);
  my $sheet = $xlsx->add_worksheet();
  $sheet->add_write_handler(qr/^.*$/, \&store_string_widths);

  my $format = $xlsx->add_format();
  $format->set_text_wrap();
  $format->set_shrink(1);

  my $header = $xlsx->add_format();
  $header->set_format_properties(
    bold => 1,
    size => 12,
    bg_color => '#cccccc',
    color => 'black'
  );

  my @spans = @{$r->{spans}};
  foreach my $span (@spans) {
    my $lastRow = $span->{row} + ($span->{rowspan} -1);
    my $lastCol = $span->{col} + ($span->{colspan} -1);
    $sheet->merge_range( $span->{row}, $span->{col}, $lastRow, $lastCol, '', $format );
  }
  my @rows = @{$r->{data}};
  for (my $i = 0; $i < scalar(@rows); $i++) {
    my @row = $rows[$i];
    for (my $j = 0; $j < scalar(@row); $j++) {
      $sheet->write( $i, $j, $row[$j], $i == 0 && $r->{header} ? $header : $format );
    }
  }

  autofit_columns($sheet);
  $xlsx->close();

  my $filename = $1 if $file =~ /.+\/(.+)$/;
  my $suffix = $Foswiki::cfg{ScriptSuffix} || '';
  my $path = $Foswiki::cfg{ScriptUrlPath} || '/bin';
  $response->header( -status  => 200 );
  return "$path/restauth$suffix/ExportExcelPlugin/get?w=$web&t=$topic&filename=$filename";
}

sub autofit_columns {
  my $worksheet = shift;
  my $col = 0;

  for my $width (@{$worksheet->{__col_widths}}) {
    $worksheet->set_column($col, $col, $width) if $width;
    $col++;
  }
}

sub store_string_widths {
  my $worksheet = shift;
  my $col       = $_[1];
  my $token     = $_[2];
  # Ignore some tokens that we aren't interested in.
  return if not defined $token;       # Ignore undefs.
  return if $token eq '';             # Ignore blank cells.
  return if ref $token eq 'ARRAY';    # Ignore array refs.
  return if $token =~ /^=/;           # Ignore formula

  # Ignore numbers
  return if $token =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/;

  # Ignore various internal and external hyperlinks. In a real scenario
  # you may wish to track the length of the optional strings used with
  # urls.
  return if $token =~ m{^[fh]tt?ps?://};
  return if $token =~ m{^mailto:};
  return if $token =~ m{^(?:in|ex)ternal:};

  # We store the string width as data in the Worksheet object. We use
  # a double underscore key name to avoid conflicts with future names.
  #
  my $old_width    = $worksheet->{__col_widths}->[$col];
  my $string_width = string_width($token);

  if (not defined $old_width or $string_width > $old_width) {
    $worksheet->{__col_widths}->[$col] = $string_width < 15 ? 15 : $string_width;
  }

  # Return control to write();
  return undef;
}

sub string_width {
  return 1.0 * length $_[0];
}

sub _restGet {
  my ( $session, $subject, $verb, $response ) = @_;
  my $query = $session->{request};
  my $filename = $query->{param}->{filename}[0];
  my $web = $query->{param}->{w}[0];
  my $topic = $query->{param}->{t}[0];

  unless ( $filename ) {
    Foswiki::Func::writeWarning( "Invalid file: $filename." );
    return;
  }

  my $name = '';
  if ( $web && $topic ) {
    $name = "$web.$topic.xlsx";
  } else {
    $name = 'export.xlsx';
  }

  my $tmpDir = Foswiki::Func::getWorkArea('ExportExcelPlugin');
  my $attachment = "$tmpDir/$filename";

  my $file;
  open $file, "< $attachment";
  binmode($file, ":raw");

  local $/;
  my $xls = <$file>;

  close $file;
  unlink $attachment;

  $response->headers({
    "Content-Disposition" => "attachment; filename=$name",
    "Content-Type" => "application/vnd.ms-excel",
  });
  $response->body( $xls );
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
