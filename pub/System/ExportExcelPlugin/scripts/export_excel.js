(function($) {
  $.fn.extend( {
    convertTable: function( table ) {
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

      var binUrl = foswiki.getPreference( 'SCRIPTURL' );
      var suffix = foswiki.getPreference( 'SCRIPTSUFFIX' );
      var url = binUrl + '/rest' + suffix + '/ExportExcelPlugin/convert';
      $.ajax({
        url: url,
        type: 'POST',
        data: { table: arg },
        success: function( data, status, xhr ) { $(this).download( data ); },
        error: function( xhr, status, error ) {
          console.log( 'ExportExcelPlugin: ' + error );
        }
      });
    },

    download: function( file ) {
      var web = foswiki.getPreference( 'WEB' );
      var topic = foswiki.getPreference( 'TOPIC' );
      var binUrl = foswiki.getPreference( 'SCRIPTURL' );
      var suffix = foswiki.getPreference( 'SCRIPTSUFFIX' );
      var url = binUrl + '/rest' + suffix + '/ExportExcelPlugin/get?filename=' + file + '&w=' + web + '&t=' + topic;
      window.location = url;
    }
  });

  $(document).ready( function() {
    if ( !foswiki.preferences.excelExport ) return;

    var selector = foswiki.preferences.excelExport.classes;
    var classes = selector.split( ',' );
    for( var i = 0; i < classes.length; i++ ) {
      selector = selector.replace( classes[i], 'table.' + classes[i] );
    }

    var tables = $(selector);
    $.each( tables, function( index, table ) {
      var tableWrapper = '<div class="excel-wrapper"></div>';
      $(table).wrap( $(tableWrapper) );

      var exportText = '';
      var lang = navigator.language || navigator.userLanguage;
      if ( /de/.test( lang ) ) {
        exportText = 'Nach Excel exportieren';
      } else {
        exportText = 'Export to Excel';
      }

      var link = $('<div class="excel-export"><img src="/pub/System/ExportExcelPlugin/images/excel-logo.png" title="' + exportText + '" /></div>');
      $(link).appendTo( $(table).parent() );

      if ( $('div.excel-wrapper').parent().attr('class') == 'foswikiTopic' ) {
        var padTop = $('div.foswikiTopic').css( 'padding-top' ).replace( 'px', '' );
        $(link).css( 'top', parseFloat( padTop ) );
      }

      $(link).on( 'click', function() {
        $(this).convertTable( table );
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
