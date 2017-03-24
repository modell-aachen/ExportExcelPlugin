#---+ Extensions
#---++ ExportExcelPlugin

# **STRING**
# Comma separated list of CSS classes to which the export functionality shall be applied.
$Foswiki::cfg{Plugins}{ExportExcelPlugin}{Classes} = 'export_excel,atpSearc';

# **BOOLEAN**
# Show export icon only on hover. Turn this of to render an icon at each exportable table
$Foswiki::cfg{Plugins}{ExportExcelPlugin}{HoverMode} = 1;

# **BOOLEAN**
# Make WebStatistics exportable
$Foswiki::cfg{Plugins}{ExportExcelPlugin}{AllowExportWebStatistics} = 0;
