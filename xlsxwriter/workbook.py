###############################################################################
#
# Workbook - A class for writing the Excel XLSX Workbook file.
#
# Copyright 2013-2015, John McNamara, jmcnamara@cpan.org
#

# Standard packages.
import sys
import re
import os
import operator
from warnings import warn
from datetime import datetime
from zipfile import ZipFile, ZIP_DEFLATED
from struct import unpack

from .compatibility import str_types

# Package imports.
from . import xmlwriter
from xlsxwriter.worksheet import Worksheet
from xlsxwriter.chartsheet import Chartsheet
from xlsxwriter.sharedstrings import SharedStringTable
from xlsxwriter.format import Format
from xlsxwriter.packager import Packager
from .utility import xl_cell_to_rowcol
from xlsxwriter.chart_area import ChartArea
from xlsxwriter.chart_bar import ChartBar
from xlsxwriter.chart_column import ChartColumn
from xlsxwriter.chart_doughnut import ChartDoughnut
from xlsxwriter.chart_line import ChartLine
from xlsxwriter.chart_pie import ChartPie
from xlsxwriter.chart_radar import ChartRadar
from xlsxwriter.chart_scatter import ChartScatter
from xlsxwriter.chart_stock import ChartStock


class Workbook(xmlwriter.XMLwriter):
    """
    A class for writing the Excel XLSX Workbook file.


    """

    ###########################################################################
    #
    # Public API.
    #
    ###########################################################################

    def __init__(self, filename=None, options={}):
        """
        Constructor.

        """

        super(Workbook, self).__init__()

        self.filename = filename

        self.tmpdir = options.get('tmpdir', None)
        self.date_1904 = options.get('date_1904', False)
        self.strings_to_numbers = options.get('strings_to_numbers', False)
        self.strings_to_formulas = options.get('strings_to_formulas', True)
        self.strings_to_urls = options.get('strings_to_urls', True)
        self.default_date_format = options.get('default_date_format', None)
        self.optimization = options.get('constant_memory', False)
        self.in_memory = options.get('in_memory', False)
        self.excel2003_style = options.get('excel2003_style', False)
        self.default_format_properties = \
            options.get('default_format_properties', {})

        self.worksheet_meta = WorksheetMeta()
        self.selected = 0
        self.fileclosed = 0
        self.filehandle = None
        self.internal_fh = 0
        self.sheet_name = 'Sheet'
        self.chart_name = 'Chart'
        self.sheetname_count = 0
        self.chartname_count = 0
        self.worksheets_objs = []
        self.charts = []
        self.drawings = []
        self.sheetnames = []
        self.formats = []
        self.xf_formats = []
        self.xf_format_indices = {}
        self.dxf_formats = []
        self.dxf_format_indices = {}
        self.palette = []
        self.font_count = 0
        self.num_format_count = 0
        self.defined_names = []
        self.named_ranges = []
        self.custom_colors = []
        self.doc_properties = {}
        self.localtime = datetime.now()
        self.num_vml_files = 0
        self.num_comment_files = 0
        self.x_window = 240
        self.y_window = 15
        self.window_width = 16095
        self.window_height = 9660
        self.tab_ratio = 500
        self.str_table = SharedStringTable()
        self.vba_project = None
        self.vba_is_stream = False
        self.vba_codename = None
        self.image_types = {}
        self.images = []
        self.border_count = 0
        self.fill_count = 0
        self.drawing_count = 0
        self.calc_mode = "auto"
        self.calc_on_load = True
        self.allow_zip64 = False
        self.calc_id = 124519

        # We can't do 'constant_memory' mode while doing 'in_memory' mode.
        if self.in_memory:
            self.optimization = False

        # Add the default cell format.
        if self.excel2003_style:
            self.add_format({'xf_index': 0, 'font_family': 0})
        else:
            self.add_format({'xf_index': 0})

        # Add a default URL format.
        self.default_url_format = self.add_format({'color': 'blue',
                                                   'underline': 1})

        # Add the default date format.
        if self.default_date_format is not None:
            self.default_date_format = \
                self.add_format({'num_format': self.default_date_format})

    def __del__(self):
        """Close file in destructor if it hasn't been closed explicitly."""
        try:
            if not self.fileclosed:
                self.close()
        except:
            raise Exception("Exception caught in workbook destructor. "
                            "Explicit close() may be required for workbook.")

    def add_worksheet(self, name=None):
        """
        Add a new worksheet to the Excel workbook.

        Args:
            name: The worksheet name. Defaults to 'Sheet1', etc.

        Returns:
            Reference to a worksheet object.

        """
        return self._add_sheet(name, is_chartsheet=False)

    def add_chartsheet(self, name=None):
        """
        Add a new chartsheet to the Excel workbook.

        Args:
            name: The chartsheet name. Defaults to 'Sheet1', etc.

        Returns:
            Reference to a chartsheet object.

        """
        return self._add_sheet(name, is_chartsheet=True)

    def add_format(self, properties={}):
        """
        Add a new Format to the Excel Workbook.

        Args:
            properties: The format properties.

        Returns:
            Reference to a Format object.

        """
        format_properties = self.default_format_properties.copy()

        if self.excel2003_style:
            format_properties = {'font_name': 'Arial', 'font_size': 10,
                                 'theme': 1 * -1}

        format_properties.update(properties)

        xf_format = Format(format_properties,
                           self.xf_format_indices,
                           self.dxf_format_indices)

        # Store the format reference.
        self.formats.append(xf_format)

        return xf_format

    def add_chart(self, options):
        """
        Create a chart object.

        Args:
            options: The chart type and subtype options.

        Returns:
            Reference to a Chart object.

        """

        # Type must be specified so we can create the required chart instance.
        chart_type = options.get('type')
        if chart_type is None:
            warn("Chart type must be defined in add_chart()")
            return

        if chart_type == 'area':
            chart = ChartArea(options)
        elif chart_type == 'bar':
            chart = ChartBar(options)
        elif chart_type == 'column':
            chart = ChartColumn(options)
        elif chart_type == 'doughnut':
            chart = ChartDoughnut(options)
        elif chart_type == 'line':
            chart = ChartLine(options)
        elif chart_type == 'pie':
            chart = ChartPie(options)
        elif chart_type == 'radar':
            chart = ChartRadar(options)
        elif chart_type == 'scatter':
            chart = ChartScatter(options)
        elif chart_type == 'stock':
            chart = ChartStock(options)
        else:
            warn("Unknown chart type '%s' in add_chart()" % chart_type)
            return

        # Set the embedded chart name if present.
        if 'name' in options:
            chart.chart_name = options['name']

        chart.embedded = True
        chart.date_1904 = self.date_1904

        self.charts.append(chart)

        return chart

    def add_vba_project(self, vba_project, is_stream=False):
        """
        Add a vbaProject binary to the Excel workbook.

        Args:
            vba_project: The vbaProject binary file name.
            is_stream:   vba_project is an in memory byte stream.

        Returns:
            Nothing.

        """
        if not is_stream and not os.path.exists(vba_project):
            warn("VBA project binary file '%s' not found." % vba_project)
            return -1

        self.vba_project = vba_project
        self.vba_is_stream = is_stream

    def close(self):
        """
        Call finalisation code and close file.

        Args:
            None.

        Returns:
            Nothing.

        """
        if not self.fileclosed:
            self.fileclosed = 1
            self._store_workbook()

    def set_properties(self, properties):
        """
        Set the document properties such as Title, Author etc.

        Args:
            properties: Dictionary of document properties.

        Returns:
            Nothing.

        """
        self.doc_properties = properties

    def set_calc_mode(self, mode, calc_id=None):
        """
        Set the Excel caclcuation mode for the workbook.

        Args:
            mode: String containing one of:
                * manual
                * auto_except_tables
                * auto

        Returns:
            Nothing.

        """
        self.calc_mode = mode

        if mode == 'manual':
            self.calc_on_load = False
        elif mode == 'auto_except_tables':
            self.calc_mode = 'autoNoTable'

        # Leave undocumented for now. Rarely required.
        if calc_id:
            self.calc_id = calc_id

    def define_name(self, name, formula):
        # Create a defined name in Excel. We handle global/workbook level
        # names and local/worksheet names.
        """
        Create a defined name in the workbook.

        Args:
            name:    The defined name.
            formula: The cell or range that the defined name refers to.

        Returns:
            Nothing.

        """
        sheet_index = None
        sheetname = ''

        # Remove the = sign from the formula if it exists.
        if formula.startswith('='):
            formula = formula.lstrip('=')

        # Local defined names are formatted like "Sheet1!name".
        sheet_parts = re.compile(r'^(.*)!(.*)$')
        match = sheet_parts.match(name)

        if match:
            sheetname = match.group(1)
            name = match.group(2)
            sheet_index = self._get_sheet_index(sheetname)

            # Warn if the sheet index wasn't found.
            if sheet_index is None:
                warn("Unknown sheet name '%s' in defined_name()" % sheetname)
                return -1
        else:
            # Use -1 to indicate global names.
            sheet_index = -1

        # Warn if the defined name contains invalid chars as defined by Excel.
        if (not re.match(r'^[\w\\][\w.]*$', name, re.UNICODE)
                or re.match(r'^\d', name)):
            warn("Invalid Excel characters in defined_name(): '%s'" % name)
            return -1

        # Warn if the defined name looks like a cell name.
        if re.match(r'^[a-zA-Z][a-zA-Z]?[a-dA-D]?[0-9]+$', name):
            warn("Name looks like a cell name in defined_name(): '%s'" % name)
            return -1

        # Warn if the name looks like a R1C1 cell reference.
        if (re.match(r'^[rcRC]$', name)
                or re.match(r'^[rcRC]\d+[rcRC]\d+$', name)):
            warn("Invalid name '%s' like a RC cell ref in defined_name()"
                 % name)
            return -1

        self.defined_names.append([name, sheet_index, formula, False])

    def worksheets(self):
        """
        Return a list of the worksheet objects in the workbook.

        Args:
            None.

        Returns:
            A list of worksheet objects.

        """
        return self.worksheets_objs

    def use_zip64(self):
        """
        Allow ZIP64 extensions when writing xlsx file zip container.

        Args:
            None.

        Returns:
            Nothing.

        """
        self.allow_zip64 = True

    def set_vba_name(self, name=None):
        """
        Set the VBA name for the workbook. By default the workbook is referred
        to as ThisWorkbook in VBA.

        Args:
            name: The VBA name for the workbook.

        Returns:
            Nothing.

        """
        if name is not None:
            self.vba_codename = name
        else:
            self.vba_codename = 'ThisWorkbook'

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _assemble_xml_file(self):
        # Assemble and write the XML file.

        # Prepare format object for passing to Style.pm.
        self._prepare_format_properties()

        # Write the XML declaration.
        self._xml_declaration()

        # Write the workbook element.
        self._write_workbook()

        # Write the fileVersion element.
        self._write_file_version()

        # Write the workbookPr element.
        self._write_workbook_pr()

        # Write the bookViews element.
        self._write_book_views()

        # Write the sheets element.
        self._write_sheets()

        # Write the workbook defined names.
        self._write_defined_names()

        # Write the calcPr element.
        self._write_calc_pr()

        # Close the workbook tag.
        self._xml_end_tag('workbook')

        # Close the file.
        self._xml_close()

    def _store_workbook(self):
        # Assemble worksheets into a workbook.
        packager = Packager()

        # Add a default worksheet if non have been added.
        if not self.worksheets():
            self.add_worksheet()

        # Ensure that at least one worksheet has been selected.
        if self.worksheet_meta.activesheet == 0:
            self.worksheets_objs[0].selected = 1
            self.worksheets_objs[0].hidden = 0

        # Set the active sheet.
        for sheet in self.worksheets():
            if sheet.index == self.worksheet_meta.activesheet:
                sheet.active = 1

        # Convert the SST strings data structure.
        self._prepare_sst_string_data()

        # Prepare the worksheet VML elements such as comments and buttons.
        self._prepare_vml()

        # Set the defined names for the worksheets such as Print Titles.
        self._prepare_defined_names()

        # Prepare the drawings, charts and images.
        self._prepare_drawings()

        # Add cached data to charts.
        self._add_chart_data()

        # Prepare the worksheet tables.
        self._prepare_tables()

        # Package the workbook.
        packager._add_workbook(self)
        packager._set_tmpdir(self.tmpdir)
        packager._set_in_memory(self.in_memory)
        xml_files = packager._create_package()

        # Free up the Packager object.
        packager = None

        xlsx_file = ZipFile(self.filename, "w", compression=ZIP_DEFLATED,
                            allowZip64=self.allow_zip64)

        # Add XML sub-files to the Zip file with their Excel filename.
        for os_filename, xml_filename, is_binary in xml_files:
            if self.in_memory:
                # The files are in-memory StringIOs.
                if is_binary:
                    xlsx_file.writestr(xml_filename, os_filename.getvalue())
                else:
                    xlsx_file.writestr(xml_filename,
                                       os_filename.getvalue().encode('utf-8'))
            else:
                # The files are tempfiles.
                xlsx_file.write(os_filename, xml_filename)
                os.remove(os_filename)

        xlsx_file.close()

    def _add_sheet(self, name, is_chartsheet):
        # Utility for shared code in add_worksheet() and add_chartsheet().

        sheet_index = len(self.worksheets_objs)
        name = self._check_sheetname(name, is_chartsheet)

        # Initialisation data to pass to the worksheet.
        init_data = {
            'name': name,
            'index': sheet_index,
            'str_table': self.str_table,
            'worksheet_meta': self.worksheet_meta,
            'optimization': self.optimization,
            'tmpdir': self.tmpdir,
            'date_1904': self.date_1904,
            'strings_to_numbers': self.strings_to_numbers,
            'strings_to_formulas': self.strings_to_formulas,
            'strings_to_urls': self.strings_to_urls,
            'default_date_format': self.default_date_format,
            'default_url_format': self.default_url_format,
            'excel2003_style': self.excel2003_style,
        }

        if is_chartsheet:
            worksheet = Chartsheet()
        else:
            worksheet = Worksheet()

        worksheet._initialize(init_data)

        self.worksheets_objs.append(worksheet)
        self.sheetnames.append(name)

        return worksheet

    def _check_sheetname(self, sheetname, is_chartsheet=False):
        # Check for valid worksheet names. We check the length, if it contains
        # any invalid chars and if the sheetname is unique in the workbook.
        invalid_char = re.compile(r'[\[\]:*?/\\]')

        # Increment the Sheet/Chart number used for default sheet names below.
        if is_chartsheet:
            self.chartname_count += 1
        else:
            self.sheetname_count += 1

        # Supply default Sheet/Chart sheetname if none has been defined.
        if sheetname is None:
            if is_chartsheet:
                sheetname = self.chart_name + str(self.chartname_count)
            else:
                sheetname = self.sheet_name + str(self.sheetname_count)

        # Check that sheet sheetname is <= 31. Excel limit.
        if len(sheetname) > 31:
            raise Exception("Excel worksheet name '%s' must be <= 31 chars." %
                            sheetname)

        # Check that sheetname doesn't contain any invalid characters
        if invalid_char.search(sheetname):
            raise Exception(
                "Invalid Excel character '[]:*?/\\' in sheetname '%s'" %
                sheetname)

        # Check that the worksheet name doesn't already exist since this is a
        # fatal Excel error. The check must be case insensitive like Excel.
        for worksheet in self.worksheets():
            if sheetname.lower() == worksheet.name.lower():
                raise Exception(
                    "Sheetname '%s', with case ignored, is already in use." %
                    sheetname)

        return sheetname

    def _prepare_format_properties(self):
        # Prepare all Format properties prior to passing them to styles.py.

        # Separate format objects into XF and DXF formats.
        self._prepare_formats()

        # Set the font index for the format objects.
        self._prepare_fonts()

        # Set the number format index for the format objects.
        self._prepare_num_formats()

        # Set the border index for the format objects.
        self._prepare_borders()

        # Set the fill index for the format objects.
        self._prepare_fills()

    def _prepare_formats(self):
        # Iterate through the XF Format objects and separate them into
        # XF and DXF formats. The XF and DF formats then need to be sorted
        # back into index order rather than creation order.
        xf_formats = []
        dxf_formats = []

        # Sort into XF and DXF formats.
        for xf_format in self.formats:
            if xf_format.xf_index is not None:
                xf_formats.append(xf_format)

            if xf_format.dxf_index is not None:
                dxf_formats.append(xf_format)

        # Pre-extend the format lists.
        self.xf_formats = [None] * len(xf_formats)
        self.dxf_formats = [None] * len(dxf_formats)

        # Rearrange formats into index order.
        for xf_format in xf_formats:
            index = xf_format.xf_index
            self.xf_formats[index] = xf_format

        for dxf_format in dxf_formats:
            index = dxf_format.dxf_index
            self.dxf_formats[index] = dxf_format

    def _set_default_xf_indices(self):
        # Set the default index for each format. Only used for testing.

        formats = list(self.formats)

        # Delete the default url format.
        del formats[1]

        # Skip the default date format if set.
        if self.default_date_format is not None:
            del formats[1]

        # Set the remaining formats.
        for xf_format in formats:
            xf_format._get_xf_index()

    def _prepare_fonts(self):
        # Iterate through the XF Format objects and give them an index to
        # non-default font elements.
        fonts = {}
        index = 0

        for xf_format in self.xf_formats:
            key = xf_format._get_font_key()
            if key in fonts:
                # Font has already been used.
                xf_format.font_index = fonts[key]
                xf_format.has_font = 0
            else:
                # This is a new font.
                fonts[key] = index
                xf_format.font_index = index
                xf_format.has_font = 1
                index += 1

        self.font_count = index

        # For DXF formats we only need to check if the properties have changed.
        for xf_format in self.dxf_formats:
            # The only font properties that can change for a DXF format are:
            # color, bold, italic, underline and strikethrough.
            if (xf_format.font_color or xf_format.bold or xf_format.italic
                    or xf_format.underline or xf_format.font_strikeout):
                xf_format.has_dxf_font = 1

    def _prepare_num_formats(self):
        # User defined records in Excel start from index 0xA4.
        num_formats = {}
        index = 164
        num_format_count = 0

        for xf_format in (self.xf_formats + self.dxf_formats):
            num_format = xf_format.num_format

            # Check if num_format is an index to a built-in number format.
            if not isinstance(num_format, str_types):
                xf_format.num_format_index = int(num_format)
                continue

            if num_format in num_formats:
                # Number xf_format has already been used.
                xf_format.num_format_index = num_formats[num_format]
            else:
                # Add a new number xf_format.
                num_formats[num_format] = index
                xf_format.num_format_index = index
                index += 1

                # Only increase font count for XF formats (not DXF formats).
                if xf_format.xf_index:
                    num_format_count += 1

        self.num_format_count = num_format_count

    def _prepare_borders(self):
        # Iterate through the XF Format objects and give them an index to
        # non-default border elements.
        borders = {}
        index = 0

        for xf_format in self.xf_formats:
            key = xf_format._get_border_key()

            if key in borders:
                # Border has already been used.
                xf_format.border_index = borders[key]
                xf_format.has_border = 0
            else:
                # This is a new border.
                borders[key] = index
                xf_format.border_index = index
                xf_format.has_border = 1
                index += 1

        self.border_count = index

        # For DXF formats we only need to check if the properties have changed.
        has_border = re.compile(r'[^0:]')

        for xf_format in self.dxf_formats:
            key = xf_format._get_border_key()

            if has_border.search(key):
                xf_format.has_dxf_border = 1

    def _prepare_fills(self):
        # Iterate through the XF Format objects and give them an index to
        # non-default fill elements.
        # The user defined fill properties start from 2 since there are 2
        # default fills: patternType="none" and patternType="gray125".
        fills = {}
        index = 2  # Start from 2. See above.

        # Add the default fills.
        fills['0:0:0'] = 0
        fills['17:0:0'] = 1

        # Store the DXF colours separately since them may be reversed below.
        for xf_format in self.dxf_formats:
            if xf_format.pattern or xf_format.bg_color or xf_format.fg_color:
                xf_format.has_dxf_fill = 1
                xf_format.dxf_bg_color = xf_format.bg_color
                xf_format.dxf_fg_color = xf_format.fg_color

        for xf_format in self.xf_formats:
            # The following logical statements jointly take care of special
            # cases in relation to cell colours and patterns:
            # 1. For a solid fill (_pattern == 1) Excel reverses the role of
            # foreground and background colours, and
            # 2. If the user specifies a foreground or background colour
            # without a pattern they probably wanted a solid fill, so we fill
            # in the defaults.
            if (xf_format.pattern == 1 and xf_format.bg_color != 0
                    and xf_format.fg_color != 0):
                tmp = xf_format.fg_color
                xf_format.fg_color = xf_format.bg_color
                xf_format.bg_color = tmp

            if (xf_format.pattern <= 1 and xf_format.bg_color != 0
                    and xf_format.fg_color == 0):
                xf_format.fg_color = xf_format.bg_color
                xf_format.bg_color = 0
                xf_format.pattern = 1

            if (xf_format.pattern <= 1 and xf_format.bg_color == 0
                    and xf_format.fg_color != 0):
                xf_format.bg_color = 0
                xf_format.pattern = 1

            key = xf_format._get_fill_key()

            if key in fills:
                # Fill has already been used.
                xf_format.fill_index = fills[key]
                xf_format.has_fill = 0
            else:
                # This is a new fill.
                fills[key] = index
                xf_format.fill_index = index
                xf_format.has_fill = 1
                index += 1

        self.fill_count = index

    def _prepare_defined_names(self):
        # Iterate through the worksheets and store any defined names in
        # addition to any user defined names. Stores the defined names
        # for the Workbook.xml and the named ranges for App.xml.
        defined_names = self.defined_names

        for sheet in self.worksheets():
            # Check for Print Area settings.
            if sheet.autofilter_area:
                hidden = 1
                sheet_range = sheet.autofilter_area
                # Store the defined names.
                defined_names.append(['_xlnm._FilterDatabase',
                                      sheet.index, sheet_range, hidden])

            # Check for Print Area settings.
            if sheet.print_area_range:
                hidden = 0
                sheet_range = sheet.print_area_range
                # Store the defined names.
                defined_names.append(['_xlnm.Print_Area',
                                      sheet.index, sheet_range, hidden])

            # Check for repeat rows/cols referred to as Print Titles.
            if sheet.repeat_col_range or sheet.repeat_row_range:
                hidden = 0
                sheet_range = ''
                if sheet.repeat_col_range and sheet.repeat_row_range:
                    sheet_range = (sheet.repeat_col_range + ',' +
                                   sheet.repeat_row_range)
                else:
                    sheet_range = (sheet.repeat_col_range +
                                   sheet.repeat_row_range)
                # Store the defined names.
                defined_names.append(['_xlnm.Print_Titles',
                                      sheet.index, sheet_range, hidden])

        defined_names = self._sort_defined_names(defined_names)
        self.defined_names = defined_names
        self.named_ranges = self._extract_named_ranges(defined_names)

    def _sort_defined_names(self, names):
        # Sort the list of list of internal and user defined names in
        # the same order as used by Excel.

        # Add a normalise name string to each list for sorting.
        for name_list in names:
            (defined_name, _, sheet_name, _) = name_list

            # Normalise the defined name by removing any leading '_xmln.'
            # from internal names and lowercasing the string.
            defined_name = defined_name.replace('_xlnm.', '').lower()

            # Normalise the sheetname by removing the leading quote and
            # lowercasing the string.
            sheet_name = sheet_name.lstrip("'").lower()

            name_list.append(defined_name + "::" + sheet_name)

        # Sort based on the normalised key.
        names.sort(key=operator.itemgetter(4))

        # Remove the extra key used for sorting.
        for name_list in names:
            name_list.pop()

        return names

    def _prepare_drawings(self):
        # Iterate through the worksheets and set up chart and image drawings.
        chart_ref_id = 0
        image_ref_id = 0
        drawing_id = 0
        x_dpi = 96
        y_dpi = 96

        for sheet in self.worksheets():
            chart_count = len(sheet.charts)
            image_count = len(sheet.images)
            shape_count = len(sheet.shapes)

            header_image_count = len(sheet.header_images)
            footer_image_count = len(sheet.footer_images)
            has_drawing = False

            if not (chart_count or image_count or shape_count
                    or header_image_count or footer_image_count):
                continue

            # Don't increase the drawing_id header/footer images.
            if chart_count or image_count or shape_count:
                drawing_id += 1
                has_drawing = True

            # Prepare the worksheet charts.
            for index in range(chart_count):
                chart_ref_id += 1
                sheet._prepare_chart(index, chart_ref_id, drawing_id)

            # Prepare the worksheet images.
            for index in range(image_count):
                filename = sheet.images[index][2]
                image_data = sheet.images[index][10]
                (image_type, width, height, name, x_dpi, y_dpi) = \
                    self._get_image_properties(filename, image_data)
                image_ref_id += 1

                sheet._prepare_image(index, image_ref_id, drawing_id, width,
                                     height, name, image_type, x_dpi, y_dpi)

            # Prepare the header images.
            for index in range(header_image_count):

                filename = sheet.header_images[index][0]
                image_data = sheet.header_images[index][1]
                position = sheet.header_images[index][2]

                (image_type, width, height, name, x_dpi, y_dpi) = \
                    self._get_image_properties(filename, image_data)

                image_ref_id += 1

                sheet._prepare_header_image(image_ref_id, width, height,
                                            name, image_type, position,
                                            x_dpi, y_dpi)

            # Prepare the footer images.
            for index in range(footer_image_count):

                filename = sheet.footer_images[index][0]
                image_data = sheet.footer_images[index][1]
                position = sheet.footer_images[index][2]

                (image_type, width, height, name, x_dpi, y_dpi) = \
                    self._get_image_properties(filename, image_data)

                image_ref_id += 1

                sheet._prepare_header_image(image_ref_id, width, height,
                                            name, image_type, position,
                                            x_dpi, y_dpi)

            if has_drawing:
                drawing = sheet.drawing
                self.drawings.append(drawing)

        # Remove charts that were created but not inserted into worksheets.
        for chart in self.charts[:]:
            if chart.id == -1:
                self.charts.remove(chart)

        # Sort the workbook charts references into the order that the were
        # written to the worksheets above.
        self.charts = sorted(self.charts, key=lambda chart: chart.id)

        self.drawing_count = drawing_id

    def _get_image_properties(self, filename, image_data):
        # Extract dimension information from the image file.
        height = 0
        width = 0
        x_dpi = 96
        y_dpi = 96

        if not image_data:
            # Open the image file and read in the data.
            fh = open(filename, "rb")
            data = fh.read()
        else:
            # Read the image data from the user supplied byte stream.
            data = image_data.getvalue()

        # Get the image filename without the path.
        image_name = os.path.basename(filename)

        # Look for some common image file markers.
        marker1 = (unpack('3s', data[1:4]))[0]
        marker2 = (unpack('>H', data[:2]))[0]
        marker3 = (unpack('2s', data[:2]))[0]

        if sys.version_info < (2, 6, 0):
            # Python 2.5/Jython.
            png_marker = 'PNG'
            bmp_marker = 'BM'
        else:
            # Eval the binary literals for Python 2.5/Jython compatibility.
            png_marker = eval("b'PNG'")
            bmp_marker = eval("b'BM'")

        if marker1 == png_marker:
            self.image_types['png'] = 1
            (image_type, width, height, x_dpi, y_dpi) = self._process_png(data)

        elif marker2 == 0xFFD8:
            self.image_types['jpeg'] = 1
            (image_type, width, height, x_dpi, y_dpi) = self._process_jpg(data)

        elif marker3 == bmp_marker:
            self.image_types['bmp'] = 1
            (image_type, width, height) = self._process_bmp(data)

        else:
            raise Exception("%s: Unknown or unsupported image file format."
                            % filename)

        # Check that we found the required data.
        if not height or not width:
            raise Exception("%s: no size data found in image file." % filename)

        # Store image data to copy it into file container.
        self.images.append([filename, image_type, image_data])

        if not image_data:
            fh.close()

        return image_type, width, height, image_name, x_dpi, y_dpi

    def _process_png(self, data):
        # Extract width and height information from a PNG file.
        offset = 8
        data_length = len(data)
        end_marker = False
        width = 0
        height = 0
        x_dpi = 96
        y_dpi = 96

        # Look for numbers rather than strings for Python 2.6/3 compatibility.
        marker_ihdr = 0x49484452  # IHDR
        marker_phys = 0x70485973  # pHYs
        marker_iend = 0X49454E44  # IEND

        # Search through the image data to read the height and width in the
        # IHDR element. Also read the DPI in the pHYs element.
        while not end_marker and offset < data_length:

            length = (unpack('>I', data[offset + 0:offset + 4]))[0]
            marker = (unpack('>I', data[offset + 4:offset + 8]))[0]

            # Read the image dimensions.
            if marker == marker_ihdr:
                width = (unpack('>I', data[offset + 8:offset + 12]))[0]
                height = (unpack('>I', data[offset + 12:offset + 16]))[0]

            # Read the image DPI.
            if marker == marker_phys:
                x_density = (unpack('>I', data[offset + 8:offset + 12]))[0]
                y_density = (unpack('>I', data[offset + 12:offset + 16]))[0]
                units = (unpack('b', data[offset + 16:offset + 17]))[0]

                if units == 1:
                    x_dpi = x_density * 0.0254
                    y_dpi = y_density * 0.0254

            if marker == marker_iend:
                end_marker = True
                continue

            offset = offset + length + 12

        return 'png', width, height, x_dpi, y_dpi

    def _process_jpg(self, data):
        # Extract width and height information from a JPEG file.
        offset = 2
        data_length = len(data)
        end_marker = False
        width = 0
        height = 0
        x_dpi = 96
        y_dpi = 96

        # Search through the image data to read the height and width in the
        # 0xFFC0/C2 element. Also read the DPI in the 0xFFE0 element.
        while not end_marker and offset < data_length:

            marker = (unpack('>H', data[offset + 0:offset + 2]))[0]
            length = (unpack('>H', data[offset + 2:offset + 4]))[0]

            # Read the image dimensions.
            if marker == 0xFFC0 or marker == 0xFFC2:
                height = (unpack('>H', data[offset + 5:offset + 7]))[0]
                width = (unpack('>H', data[offset + 7:offset + 9]))[0]

            # Read the image DPI.
            if marker == 0xFFE0:
                units = (unpack('b', data[offset + 11:offset + 12]))[0]
                x_density = (unpack('>H', data[offset + 12:offset + 14]))[0]
                y_density = (unpack('>H', data[offset + 14:offset + 16]))[0]

                if units == 1:
                    x_dpi = x_density
                    y_dpi = y_density

                if units == 2:
                    x_dpi = x_density * 2.54
                    y_dpi = y_density * 2.54

            if marker == 0xFFDA:
                end_marker = True
                continue

            offset = offset + length + 2

        return 'jpeg', width, height, x_dpi, y_dpi

    def _process_bmp(self, data):
        # Extract width and height information from a BMP file.
        width = (unpack('<L', data[18:22]))[0]
        height = (unpack('<L', data[22:26]))[0]
        return 'bmp', width, height

    def _extract_named_ranges(self, defined_names):
        # Extract the named ranges from the sorted list of defined names.
        # These are used in the App.xml file.
        named_ranges = []

        for defined_name in defined_names:

            name = defined_name[0]
            index = defined_name[1]
            sheet_range = defined_name[2]

            # Skip autoFilter ranges.
            if name == '_xlnm._FilterDatabase':
                continue

            # We are only interested in defined names with ranges.
            if '!' in sheet_range:
                sheet_name, _ = sheet_range.split('!', 1)

                # Match Print_Area and Print_Titles xlnm types.
                if name.startswith('_xlnm.'):
                    xlnm_type = name.replace('_xlnm.', '')
                    name = sheet_name + '!' + xlnm_type
                elif index != -1:
                    name = sheet_name + '!' + name

                named_ranges.append(name)

        return named_ranges

    def _get_sheet_index(self, sheetname):
        # Convert a sheet name to its index. Return None otherwise.
        sheetname = sheetname.strip("'")

        if sheetname in self.sheetnames:
            return self.sheetnames.index(sheetname)
        else:
            return None

    def _prepare_vml(self):
        # Iterate through the worksheets and set up the VML objects.
        comment_id = 0
        vml_drawing_id = 0
        vml_data_id = 1
        vml_header_id = 0
        vml_shape_id = 1024
        vml_files = 0
        comment_files = 0
        has_button = False

        for sheet in self.worksheets():
            if not sheet.has_vml and not sheet.has_header_vml:
                continue

            vml_files += 1

            if sheet.has_vml:
                if sheet.has_comments:
                    comment_files += 1
                    comment_id += 1

                vml_drawing_id += 1

                count = sheet._prepare_vml_objects(vml_data_id,
                                                   vml_shape_id,
                                                   vml_drawing_id,
                                                   comment_id)

                # Each VML should start with a shape id incremented by 1024.
                vml_data_id += 1 * int((1024 + count) / 1024)
                vml_shape_id += 1024 * int((1024 + count) / 1024)

            if sheet.has_header_vml:
                vml_header_id += 1
                vml_drawing_id += 1
                sheet._prepare_header_vml_objects(vml_header_id,
                                                  vml_drawing_id)

            self.num_vml_files = vml_files
            self.num_comment_files = comment_files

            if len(sheet.buttons_list):
                has_button = True

                # Set the sheet vba_codename if it has a button and the
                # workbook has a vbaProject binary.
                if self.vba_project and sheet.vba_codename is None:
                    sheet.set_vba_name()

        # Add a font format for cell comments.
        if comment_files > 0:
            xf = self.add_format({'font_name': 'Tahoma', 'font_size': 8,
                                  'color_indexed': 81, 'font_only': True})
            xf._get_xf_index()

        # Set the workbook vba_codename if one of the sheets has a button and
        # the workbook has a vbaProject binary.
        if has_button and self.vba_project and self.vba_codename is None:
            self.set_vba_name()

    def _prepare_tables(self):
        # Set the table ids for the worksheet tables.
        table_id = 0

        for sheet in self.worksheets():
            table_count = len(sheet.tables)

            if not table_count:
                continue

            sheet._prepare_tables(table_id + 1)
            table_id += table_count

    def _add_chart_data(self):
        # Add "cached" data to charts to provide the numCache and strCacher
        # data for series and title/axis ranges.
        worksheets = {}
        seen_ranges = {}

        # Map worksheet names to worksheet objects.
        for worksheet in self.worksheets():
            worksheets[worksheet.name] = worksheet

        for chart in self.charts:

            for c_range in chart.formula_ids.keys():
                r_id = chart.formula_ids[c_range]

                # Skip if the series has user defined data.
                if chart.formula_data[r_id] is not None:
                    if (c_range not in seen_ranges
                            or seen_ranges[c_range] is None):
                        data = chart.formula_data[r_id]
                        seen_ranges[c_range] = data
                    continue

                # Check to see if the data is already cached locally.
                if c_range in seen_ranges:
                    chart.formula_data[r_id] = seen_ranges[c_range]
                    continue

                # Convert the range formula to a sheet name and cell range.
                (sheetname, cells) = self._get_chart_range(c_range)

                # Skip if we couldn't parse the formula.
                if sheetname is None:
                    continue

                # Handle non-contiguous ranges like:
                #     (Sheet1!$A$1:$A$2,Sheet1!$A$4:$A$5).
                # We don't try to parse them. We just return an empty list.
                if sheetname.startswith('('):
                    chart.formula_data[r_id] = []
                    seen_ranges[c_range] = []
                    continue

                # Warn if the name is unknown since it indicates a user error
                # in a chart series formula.
                if sheetname not in worksheets:
                    warn("Unknown worksheet reference '%s' in range "
                         "'%s' passed to add_series()" % (sheetname, c_range))
                    chart.formula_data[r_id] = []
                    seen_ranges[c_range] = []
                    continue

                # Find the worksheet object based on the sheet name.
                worksheet = worksheets[sheetname]

                # Get the data from the worksheet table.
                data = worksheet._get_range_data(*cells)

                # TODO. Handle SST string ids if required.

                # Add the data to the chart.
                chart.formula_data[r_id] = data

                # Store range data locally to avoid lookup if seen again.
                seen_ranges[c_range] = data

    def _get_chart_range(self, c_range):
        # Convert a range formula such as Sheet1!$B$1:$B$5 into a sheet name
        # and cell range such as ( 'Sheet1', 0, 1, 4, 1 ).

        # Split the range formula into sheetname and cells at the last '!'.
        pos = c_range.rfind('!')
        if pos > 0:
            sheetname = c_range[:pos]
            cells = c_range[pos + 1:]
        else:
            return None

        # Split the cell range into 2 cells or else use single cell for both.
        if cells.find(':') > 0:
            (cell_1, cell_2) = cells.split(':', 1)
        else:
            (cell_1, cell_2) = (cells, cells)

        # Remove leading/trailing quotes and convert escaped quotes to single.
        sheetname = sheetname.strip("'")
        sheetname = sheetname.replace("''", "'")

        (row_start, col_start) = xl_cell_to_rowcol(cell_1)
        (row_end, col_end) = xl_cell_to_rowcol(cell_2)

        # Check that we have a 1D range only.
        if row_start != row_end and col_start != col_end:
            return None

        return sheetname, [row_start, col_start, row_end, col_end]

    def _prepare_sst_string_data(self):
        # Convert the SST string data from a dict to a list.
        self.str_table._sort_string_data()

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    def _write_workbook(self):
        # Write <workbook> element.

        schema = 'http://schemas.openxmlformats.org'
        xmlns = schema + '/spreadsheetml/2006/main'
        xmlns_r = schema + '/officeDocument/2006/relationships'

        attributes = [
            ('xmlns', xmlns),
            ('xmlns:r', xmlns_r),
        ]

        self._xml_start_tag('workbook', attributes)

    def _write_file_version(self):
        # Write the <fileVersion> element.

        app_name = 'xl'
        last_edited = 4
        lowest_edited = 4
        rup_build = 4505

        attributes = [
            ('appName', app_name),
            ('lastEdited', last_edited),
            ('lowestEdited', lowest_edited),
            ('rupBuild', rup_build),
        ]

        if self.vba_project:
            attributes.append(
                ('codeName', '{37E998C4-C9E5-D4B9-71C8-EB1FF731991C}'))

        self._xml_empty_tag('fileVersion', attributes)

    def _write_workbook_pr(self):
        # Write <workbookPr> element.
        default_theme_version = 124226
        attributes = []

        if self.vba_codename:
            attributes.append(('codeName', self.vba_codename))
        if self.date_1904:
            attributes.append(('date1904', 1))

        attributes.append(('defaultThemeVersion', default_theme_version))

        self._xml_empty_tag('workbookPr', attributes)

    def _write_book_views(self):
        # Write <bookViews> element.
        self._xml_start_tag('bookViews')
        self._write_workbook_view()
        self._xml_end_tag('bookViews')

    def _write_workbook_view(self):
        # Write <workbookView> element.
        attributes = [
            ('xWindow', self.x_window),
            ('yWindow', self.y_window),
            ('windowWidth', self.window_width),
            ('windowHeight', self.window_height),
        ]

        # Store the tabRatio attribute when it isn't the default.
        if self.tab_ratio != 500:
            attributes.append(('tabRatio', self.tab_ratio))

        # Store the firstSheet attribute when it isn't the default.
        if self.worksheet_meta.firstsheet > 0:
            firstsheet = self.worksheet_meta.firstsheet + 1
            attributes.append(('firstSheet', firstsheet))

        # Store the activeTab attribute when it isn't the first sheet.
        if self.worksheet_meta.activesheet > 0:
            attributes.append(('activeTab', self.worksheet_meta.activesheet))

        self._xml_empty_tag('workbookView', attributes)

    def _write_sheets(self):
        # Write <sheets> element.
        self._xml_start_tag('sheets')

        id_num = 1
        for worksheet in self.worksheets():
            self._write_sheet(worksheet.name, id_num, worksheet.hidden)
            id_num += 1

        self._xml_end_tag('sheets')

    def _write_sheet(self, name, sheet_id, hidden):
        # Write <sheet> element.
        attributes = [
            ('name', name),
            ('sheetId', sheet_id),
        ]

        if hidden:
            attributes.append(('state', 'hidden'))

        attributes.append(('r:id', 'rId' + str(sheet_id)))

        self._xml_empty_tag('sheet', attributes)

    def _write_calc_pr(self):
        # Write the <calcPr> element.
        attributes = [('calcId', self.calc_id)]

        if self.calc_mode == 'manual':
            attributes.append(('calcMode', self.calc_mode))
            attributes.append(('calcOnSave', "0"))
        elif self.calc_mode == 'autoNoTable':
            attributes.append(('calcMode', self.calc_mode))

        if self.calc_on_load:
            attributes.append(('fullCalcOnLoad', '1'))

        self._xml_empty_tag('calcPr', attributes)

    def _write_defined_names(self):
        # Write the <definedNames> element.
        if not self.defined_names:
            return

        self._xml_start_tag('definedNames')

        for defined_name in self.defined_names:
            self._write_defined_name(defined_name)

        self._xml_end_tag('definedNames')

    def _write_defined_name(self, defined_name):
        # Write the <definedName> element.
        name = defined_name[0]
        sheet_id = defined_name[1]
        sheet_range = defined_name[2]
        hidden = defined_name[3]

        attributes = [('name', name)]

        if sheet_id != -1:
            attributes.append(('localSheetId', sheet_id))
        if hidden:
            attributes.append(('hidden', 1))

        self._xml_data_element('definedName', sheet_range, attributes)


# A metadata class to share data between worksheets.
class WorksheetMeta(object):
    """
    A class to track worksheets data such as the active sheet and the
    first sheet..

    """

    def __init__(self):
        self.activesheet = 0
        self.firstsheet = 0
