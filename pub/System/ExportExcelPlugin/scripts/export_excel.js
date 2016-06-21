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
  var $hoveredTable = null;
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
    $hoveredTable = null;
    timer = null;
  };

  var handleStaticHint = function() {
    var $exportable = $(this);
    var $div = $('<div class="xlsxcontainer"></div>');
    var $icon = $img.clone().addClass('static');
    $icon.attr('src', $icon.attr('src').replace(/excel-32/, 'excel-24'));

    $exportable.wrap($div);
    $exportable.parent().prepend($icon);
  };

  var onMouseEnter = function() {
    $hoveredTable = $(this);

    var offsetY = $hoveredTable.css('margin-top');
    offsetY = parseFloat(offsetY.replace(/px/, ''));

    var pos = $hoveredTable.position();
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
    var cfg = foswiki.preferences.excelExport;
    var selector = cfg.classes;
    var classes = selector.split( ',' );
    for( var i = 0; i < classes.length; i++ ) {
      selector = selector.replace( classes[i], 'table.' + classes[i] );
    }

    if ( cfg.webstatistics && foswiki.preferences.TOPIC === 'WebStatistics' ) {
      selector += ',#modacContents table.foswikiTable';
    }

    $img.on('mouseenter', keepHint);
    $img.on('mouseleave', removeHint);

    $('body').on('click', '.xslxhint', function() {
      $.blockUI();
      var $exportable = $hoveredTable || $(this).next();
      exporter.export($exportable).done(function(xlsx) {
        window.location = xlsx;
      }).always($.unblockUI);
    });

    if (cfg.useHover) {
      $('.foswikiTopic').on('mouseenter', selector, onMouseEnter);
      $('.foswikiTopic').on('mouseleave', selector, onMouseLeave);
    } else {
      $(selector).livequery(handleStaticHint);
    }
  });
})(jQuery);
