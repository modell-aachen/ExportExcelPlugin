var ExcelExporter = function() {};
ExcelExporter.prototype.export = function(table, filterCols) {
  var payload = this.serializeTable(table, filterCols);
  if (payload === null) return null;

  var fw = foswiki;
  var target = [
    fw.getPreference('SCRIPTURL'),
    '/restauth',
    fw.getPreference('SCRIPTSUFFIX'),
    '/ExportExcelPlugin/export'
  ].join('');

  payload.web = fw.getPreference('WEB');
  payload.topic = fw.getPreference('TOPIC');

  return $.ajax({
    url: target,
    method: 'POST',
    data: {payload: JSON.stringify(payload)},
  });
};

ExcelExporter.prototype.serializeTable = function(table, filterCols) {
  if (!table || !$(table).is('table')) {
    if (window.console && console.error) {
      console.error('Invalid table data!');
    }

    return null;
  }

  var $table = $(table);
  filterCols = filterCols || [];
  var retval = {data: [], header: $table.find('tr th').length};
  $table.find('tr').each(function() {
    var arr = [];
    var cols = $(this).find('th,td');
    for( var i = 0; i < cols.length; ++i) {
      if (filterCols.length && filterCols.indexOf(i) == -1) continue;
      var txt = $(cols[i]).text();
      txt = txt.replace(/^[\n\r\s\t]+/mg, '');
      txt = txt.replace(/[\n\r\s\t]+$/mg, '');
      arr.push(txt);
    }

    retval.data.push(arr);
  });

  return retval;
};

(function($) {
  var exporter = new ExcelExporter();
  var $table = null;
  var timer = null;
  var $img = $('<img class="xslxhint" src="/pub/System/ExportExcelPlugin/images/excel-32.png" title="" />');

  var handleXLSXLINK = function() {
    var $link = $(this);
    var $table = $($link.data('selector'));
    var filter = $link.data('columns').split(',').map(function(c) {
      return parseInt(c);
    });

    $.blockUI();
    exporter.export($table, filter).done(function(xlsx) {
      window.location = xlsx;
    }).fail(function(err) {
      if (window.console && console.error) {
        console.error(err);
      }
    }).always($.unblockUI);
    return false;
  };

  var keepHint = function() {
    if (timer !== null) {
      clearTimeout(timer);
    }
  };

  var removeHint = function() {
    $img.detach();
    $table = null;
    timer = null;
  };

  var onMouseEnter = function() {
    $table = $(this);

      var offsetY = $table.css('margin-top');
      offsetY = parseFloat(offsetY.replace(/px/, ''));

      var pos = $table.position();
      $img.css('top', pos.top + offsetY);
      $img.css('left', pos.left - 40);
      $img.appendTo('.foswikiTopic');
  };

  var onMouseLeave = function() {
    timer = setTimeout(removeHint, 300);
  };


  $(document).ready( function() {
    $('body').on('click', '.xlsxlink', handleXLSXLINK);

    if ( !foswiki.preferences.excelExport ) return;
    var selector = foswiki.preferences.excelExport.classes;
    var classes = selector.split( ',' );
    for( var i = 0; i < classes.length; i++ ) {
      selector = selector.replace( classes[i], 'table.' + classes[i] );
    }

    if ( foswiki.preferences.excelExport.webstatistics && foswiki.preferences.TOPIC === 'WebStatistics' ) {
      selector += ',#modacContents table.foswikiTable';
    }

    $img.on('mouseenter', keepHint);
    $img.on('mouseleave', removeHint);
    $img.on('click', function() {
      $.blockUI();
      exporter.export($table).done(function(xlsx) {
        window.location = xlsx;
      }).always($.unblockUI);
    });

    $('.foswikiTopic').on('mouseenter', selector, onMouseEnter);
    $('.foswikiTopic').on('mouseleave', selector, onMouseLeave);
  });
})(jQuery);
