(function($) {
  $.fn.extend( {
    convertTable: function( table ) {
      var binUrl = foswiki.getPreference( 'SCRIPTURL' );
      var suffix = foswiki.getPreference( 'SCRIPTSUFFIX' );
      var url = binUrl + '/rest' + suffix + '/ExportExcelPlugin/convert';
      $.ajax({
        url: url,
        type: 'POST',
        data: { table: table },
        success: function( data, status, xhr ) { $(this).download( data ); },
        error: function( xhr, status, error ) { alert( 'Oops, something went wrong.\nPlease try again.' ); }
      });
    },

    download: function( file ) {
      var binUrl = foswiki.getPreference( 'SCRIPTURL' );
      var suffix = foswiki.getPreference( 'SCRIPTSUFFIX' );
      var url = binUrl + '/rest' + suffix + '/ExportExcelPlugin/get?filename=' + file;
      window.location = url;
    }
  });

  $(document).ready( function() {
    var tables = $('table.exportable,table.atpSearch');
    $.each( tables, function( index, table ) {
      var tableWrapper = '<div class="excel-wrapper"></div>';
      $(table).wrap( $(tableWrapper) );

      var exportText = '';
      var lang = navigator.language || navigator.userLanguage;
      if ( /de/.test( lang ) ) {
        exportText = 'Excel-Export';
      } else {
        exportText = 'Export to Excel';
      }

      var link = $('<div class="excel-export">' + exportText + '</div>');
      $(link).appendTo( $(table).parent() );

      var lw = $(link).width();
      var tw = $(table).width();
      $(link).css( 'left', tw - (lw+25) );



      $(link).on( 'click', function() {
        var arg = "";
        var rows = $(table).find('tr');
        $.each( rows, function( i, row ) {
          var line = "";

          // ToDo: geht auch schöner!!
          var cols = $(row).find('th');
          $.each( cols, function( j, col ) {
            var c = encodeURIComponent( 'TH:' + $(col).text() );
            line += c + ";";
          });

          cols = $(row).find('td');
          $.each( cols, function( j, col ) {
            var c = encodeURIComponent( $(col).text() );
            line += c + ";";
          });

          arg += line + "\n";
        });

        $(this).convertTable( arg );
      });

      $(table).parent().on( 'mouseenter', function() {
        $(link).css( 'display', 'inline-block' );
      });

      $(table).parent().on( 'mouseleave', function() {
        $(link).css( 'display', 'none' );
      });
    });
  });
})(jQuery);
