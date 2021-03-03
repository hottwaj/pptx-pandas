from __future__ import annotations
from pptx import Presentation
from pptx.util import Inches, Cm, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
from pptx.exc import PackageNotFoundError
import pptx
import numbers
import pandas
from prettypandas import PrettyPandas
DIST_METRIC = Inches
from six import text_type
from itertools import zip_longest

from shutil import copyfile

class NoExistingSlideFoundError(RuntimeError): pass

BASE_POSITIONS = {
    'content': [5.2, 4.1, 5, 2.7],
    'index_width': 0.7, 
    'col_width': 0.6, 
    'single_table_margin_top': 4.3,
    'single_table_margin_left': None, # set to value other than None for different table vs chart left position
    'single_item_margin_top': 1.7, 
    'single_item_margin_left': 2.2,
    'multi_item_margin_top': 1.4, 
    'multi_item_margin_left': 1.5,
    'charts_horizontal_gap': 0.7, 
    'charts_vertical_gap': 0.5,
    'table_row_height': 0.2
}
    
class PresentationWriter():
    def __init__(self, pptx_file: str, pptx_title: str, pptx_template: str = None, 
                 default_overwrite_if_present: bool = False,
                 default_overwrite_only: bool = False,
                 default_position_overrides: dict = {},
                 default_table_font_attrs: dict = {'size': Pt(7), 'name': 'Calibri'},
                 default_include_internal_cell_borders = True,
                 default_border_kwargs = {}):
        self.pptx_file = pptx_file
        self.pptx_title = pptx_title
        
        if pptx_template is not None:
            copyfile(pptx_template, pptx_file)
        
        try:
            self.presentation = Presentation(self.pptx_file) #open existing file if present
        except PackageNotFoundError as ex:
            if pptx_template is None:
                self.presentation = Presentation() #open a blank presentation
            else:
                raise PackageNotFoundError(f'Could not open file file "{pptx_file}" after copying template file "{pptx_template}" onto it') from ex # re-raise the exception
        core_props = self.presentation.core_properties
        core_props.title = pptx_title
        core_props.keywords = ''
        
        self.default_overwrite_if_present = default_overwrite_if_present
        self.default_overwrite_only = default_overwrite_only
        self.default_positions = dict(**BASE_POSITIONS)
        self.default_positions.update(default_position_overrides)
        self.default_table_font_attrs = default_table_font_attrs
        self.default_include_internal_cell_borders = default_include_internal_cell_borders
        self.default_border_kwargs = default_border_kwargs
        
    def save_presentation(self):
        self.presentation.save(self.pptx_file)
            
    def write_slide(self, title: str, charts = [], tables: list[PrettyPandas] = [], 
                    charts_and_tables = [],
                    overwrite_if_present: Optional[bool] = None,
                    overwrite_only: Optional[bool] = None,
                    charts_per_row: int = 2, tables_per_row: int = 1,
                    position_overrides: dict = {},
                    auto_position_charts_and_tables: bool = False,
                    table_font_attrs: Optional[dict] = None,
                    caption_font_attrs: Optional[dict] = None,
                    include_internal_cell_borders: Optional[bool] = None,
                    border_kwargs: Optional[dict] = None) -> pptx.slide.Slide:
        
        if charts_and_tables:
            if charts or tables:
                raise ValueError('Either "charts_and_tables" or a combination of "charts" and "tables" args should be used')
            charts = []
            tables = []
            for item in charts_and_tables:
                if isinstance(item, PrettyPandas):
                    tables.append(item)
                else:
                    charts.append(item)
            auto_position_charts_and_tables = True

        positions = dict(**self.default_positions)
        positions.update(position_overrides)
        
        if not isinstance(charts, (list, tuple)):
            charts = [charts]
        if not isinstance(tables, (list, tuple)):
            tables = [tables]
            
        overwrite_if_present = overwrite_if_present if overwrite_if_present is not None else self.default_overwrite_if_present
        overwrite_only = overwrite_only if overwrite_only is not None else self.default_overwrite_only
        if overwrite_if_present or overwrite_only:
            try:
                self.overwrite_pptx(title, charts, tables)
            except NoExistingSlideFoundError:
                if overwrite_only:
                    raise
            else:
                return
            
        include_internal_cell_borders = include_internal_cell_borders if include_internal_cell_borders is not None else self.default_include_internal_cell_borders
        border_kwargs = border_kwargs if border_kwargs is not None else self.default_border_kwargs
        prs = self.presentation
        title_only_slide_layout = prs.slide_layouts[3]
        slide = prs.slides.add_slide(title_only_slide_layout)
        shapes = slide.shapes

        content, subtitle, title_shape, *others = [s for s in slide.shapes if s.has_text_frame]
        title_shape.text = self.pptx_title
        subtitle.text = title

        content.left, content.top, content.width, content.height = (Inches(x) for x in positions['content'])

        current_row = []
        shapes_added = [current_row]
        chart_shapes = []
        if charts:
            total_chart_width = positions['multi_item_margin_left'] if len(charts) > 1 else positions['single_item_margin_left']
            initial_width = total_chart_width
            total_chart_height = positions['multi_item_margin_top'] if len(charts) > 2 else positions['single_item_margin_top']
            for i, chrt in enumerate(charts):
                if hasattr(chrt, 'width'):
                    chart_width = chrt.width / 800.0 * 5.0
                    chart_height = chrt.height / 800.0 * 5.0
                else:
                    chart_width = chrt.figure.get_figwidth() #already in inches
                    chart_height = chrt.figure.get_figheight()

                img_file = 'test%d.png' % i
                chart_to_file(chrt, img_file)

                pic = slide.shapes.add_picture(img_file, 
                                               left = Inches(total_chart_width), 
                                               top = Inches(total_chart_height), 
                                               width = Inches(chart_width), 
                                               height = Inches(chart_height))
                total_chart_width += chart_width + positions['charts_horizontal_gap']
                current_row.append(pic)
                chart_shapes.append(pic)
                if (i % charts_per_row) == (charts_per_row-1):
                    total_chart_height += chart_height + positions['charts_vertical_gap']
                    total_chart_width = initial_width
                    current_row = []
                    shapes_added.append(current_row)
        
        if current_row:
            current_row = []
            shapes_added.append(current_row)
        table_shapes = []
        if tables:
            total_width = (positions['multi_item_margin_left'] 
                           if len(tables) > 1 
                           else positions['single_table_margin_left'] or positions['single_item_margin_left'])
            initial_width = total_width
            total_height = positions['multi_item_margin_top'] if len(charts) == 0 else positions['single_table_margin_top']
            font_attrs = table_font_attrs or self.default_table_font_attrs
            if caption_font_attrs is None:
                caption_font_attrs = {'bold': True, **font_attrs}

            for i, table in enumerate(tables):
                col_widths = [positions['index_width']] + [positions['col_width']]*len(table.columns)
                table_width = sum(col_widths)
                table_df = prettypandas_to_formatted_df(table)
                table_height = positions['table_row_height']*(len(table_df)+get_index_numlevels(table_df.columns))
                table_shape = create_pptx_table(slide, table_df, left = total_width, top = total_height, 
                                                col_width = col_widths, row_height = positions['table_row_height'],
                                                font_attrs = font_attrs, 
                                                border_kwargs = border_kwargs, include_internal_cell_borders = include_internal_cell_borders)
                table_caption = getattr(table, 'caption')
                if table_caption:
                    caption_box = slide.shapes.add_textbox(left = DIST_METRIC(total_width), 
                                                           top = DIST_METRIC(total_height) - caption_font_attrs['size']*2, 
                                                           width = DIST_METRIC(table_width), 
                                                           height = caption_font_attrs['size'])
                    set_cell_text(caption_box, table_caption)
                    set_cell_font_attrs(caption_box, **caption_font_attrs)
                    table_shape._pptx_pandas_caption = caption_box
                total_width += table_width + positions['charts_horizontal_gap']
                current_row.append(table_shape)
                table_shapes.append(table_shape)
                if (i % tables_per_row) == (tables_per_row-1):
                    total_height += table_height + positions['charts_vertical_gap']
                    total_width = initial_width
                    current_row = []
                    shapes_added.append(current_row)

        if not current_row:
            shapes_added = shapes_added[:-1]
            
        if auto_position_charts_and_tables:
            if charts_and_tables:
                # create a new "shapes_added" list of items by row, in the order of charts/tables spcified by "charts_and_tables"
                chart_shapes_iter = iter(chart_shapes)
                table_shapes_iter = iter(table_shapes)
                shapes_added = []
                current_row = []
                for i, item in enumerate(charts_and_tables):
                    if isinstance(item, PrettyPandas):
                        current_row.append(next(table_shapes_iter))
                    else:
                        current_row.append(next(chart_shapes_iter))
                    if (i+1) % charts_per_row == 0:
                        shapes_added.append(current_row)
                        current_row = []
                if current_row:
                    shapes_added.append(current_row)

            content_area_top = subtitle.top + subtitle.height
            content_area_height = self.presentation.slide_height - content_area_top
            if len(others) == 1:
                footnotes = others[0] #assume remaining text shape is cell footnotes
                content_area_height -= self.presentation.slide_height - footnotes.top 
            # calculate shift from slide centre needed to put shapes in centre between bottom of subtitle and top of footnotes (if present)
            centre_y_shift = content_area_height/2 + content_area_top - self.presentation.slide_height/2
            self.auto_position_shapes(shapes_added,
                                      Inches(positions['charts_horizontal_gap']),
                                      Inches(positions['charts_vertical_gap']),
                                      centre_y_shift)
            for table_shape in table_shapes:
                caption_box = getattr(table_shape, '_pptx_pandas_caption', None)
                if caption_box:
                    caption_box.left = table_shape.left
                    caption_box.top = table_shape.top - caption_box.height*2
        self.save_presentation()
        return slide
    
    def auto_position_shapes(self, shapes_by_row: list[list[pptx.shapes.base.BaseShape]], 
                             horizontal_margin: int, vertical_margin: int, centre_y_shift: int = 0):
        shapes_by_col = list(zip_longest(*shapes_by_row))
        col_widths = [max([s.width for s in col if s is not None]) for col in shapes_by_col]
        row_heights = [max([s.height for s in row]) for row in shapes_by_row]
        total_width = sum(col_widths)
        total_height = sum(row_heights)

        total_height_inc_margins = total_height + (len(row_heights)-1)*vertical_margin
        top_pos = (self.presentation.slide_height - total_height_inc_margins) / 2
        for cell_height, row in zip(row_heights, shapes_by_row):
            if len(row) < len(col_widths):
                # row that has fewer items than other rows
                # with current layout logic this can happen if we have a final ragged row e.g. rows of 3, 3, 3, 1)
                # can also happen if different charts_per_row and tables_per_row values are used.  might get odd results in that situation
                cell_widths = [s.width for s in row]
                row_width = sum(cell_widths)
            else:
                cell_widths = col_widths
                row_width = total_width
            row_width += (len(row)-1) * horizontal_margin

            # calculate left position to centre the cell on the page, given row width
            left_pos = (self.presentation.slide_width - row_width) / 2
            for shape, width in zip(row, cell_widths):
                #now position the shape in the centre of containing cell
                shape.top = int(top_pos + (cell_height - shape.height) / 2 + centre_y_shift)
                shape.left = int(left_pos + (width - shape.width) / 2)
                left_pos += width + horizontal_margin

            top_pos += cell_height
    
    def overwrite_pptx(self, slide_title, charts = None, tables = None):
        prs = self.presentation

        for i, slide in enumerate(prs.slides):
            text_boxes = [s for s in slide.shapes if s.has_text_frame]

            matched_title = False
            for tb in text_boxes:
                if tb.text == slide_title:
                    matched_title = True
                    break

            if matched_title:
                break

        if matched_title:
            found_charts = []
            found_tables = []
            for s in slide.shapes:
                if isinstance(s, pptx.shapes.picture.Picture):
                    found_charts.append(s)
                elif isinstance(s, pptx.shapes.graphfrm.GraphicFrame) and s.has_table:
                    found_tables.append(s)

            if charts:
                if len(charts) != len(found_charts):
                    raise RuntimeError('Need %d Picture shapes but %d available on slide "%s"'
                                       % (len(strats_chart), len(found_charts), slide_title))

                for i, strats_chart in enumerate(charts):
                    img_file = 'test%d.png' % i
                    chart_to_file(strats_chart, img_file)

                    # Replace image:
                    picture = found_charts[i]
                    with open(img_file, 'rb') as f:
                        imgBlob = f.read()
                    imgRID = picture._pic.xpath('./p:blipFill/a:blip/@r:embed')[0]
                    imgPart = slide.part.related_parts[imgRID]
                    imgPart._blob = imgBlob

            if tables:
                if len(tables) != len(found_tables):
                    raise RuntimeError('Need %d Table shapes but %d available on slide "%s"'
                                       % (tables, len(found_tables), slide_title))
                    
                for i, strats_table in enumerate(tables):
                    table_df = prettypandas_to_formatted_df(table)
                    write_pptx_dataframe(table_df, found_tables[i].table, overwrite_formatting = False)

            self.save_presentation()
        else:
            raise NoExistingSlideFoundError('No existing slide titled "%s"' % slide_title)

def prettypandas_to_formatted_df(pp_instance):
    table_df = pp_instance.get_formatted_df()
    table_df.columns = [s.replace('<br>', '\n')
                        for s in table_df.columns
                        if isinstance(s, str)]
    return table_df

empty_text_set = {'', None}
def set_cell_text(cell, text, overwrite_formatting = True):
    if text in empty_text_set:
        text = "\u00A0" # unicode nbsp - needed to fill empty cells as otherwise formatting is not applied by PPT
    if overwrite_formatting:
        p = cell.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = text
    else:
        p = cell.text_frame.paragraphs[0]
        if p.runs:
            p.runs[0].text = text
        else:
            r = p.add_run()
            r.text = text
        
def set_cell_font_attrs(cell, **kwargs):
    for p in cell.text_frame.paragraphs:
        for r in p.runs:
            for k, v in kwargs.items():
                if k == 'color_rgb':
                    r.font.color.rgb = v
                else:
                    setattr(r.font, k, v)
                
def format_cell_text(val, float_format = '{:.0f}', int_format = '{:d}'):
    if isinstance(val, numbers.Integral):
        return int_format.format(val)
    elif isinstance(val, numbers.Real):
        return float_format.format(val)
    else:
        return text_type(val)
    
def set_cell_borders(cell, border_color="4f81bd", border_width='6350', border_scheme_color = 'accent1', borders = 'LRTB'):
    """ Hack function to enable the setting of border width and border color"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    for border in borders:
        lnL = SubElement(tcPr, 'a:ln' + border, w=border_width, cap='flat', cmpd='sng', algn='ctr')
        lnL_solidFill = SubElement(lnL, 'a:solidFill')
        if border_scheme_color is not None:
            lnL_srgbClr = SubElement(lnL_solidFill, 'a:schemeClr', val=border_scheme_color)
        else:
            lnL_srgbClr = SubElement(lnL_solidFill, 'a:srgbClr', val=border_color)
        lnL_prstDash = SubElement(lnL, 'a:prstDash', val='solid')
        lnL_round_ = SubElement(lnL, 'a:round')
        lnL_headEnd = SubElement(lnL, 'a:headEnd', type='none', w='med', len='med')
        lnL_tailEnd = SubElement(lnL, 'a:tailEnd', type='none', w='med', len='med')
    cell.fill.background()
    
def get_index_numlevels(pd_index):
    if isinstance(pd_index, pandas.MultiIndex):
        return len(pd_index.levels)
    else:
        return 1
        
def write_pptx_dataframe(dataframe, pptx_table, col_width = 1.0, format_opts = {}, font_attrs = {'size': Pt(8), 'name': 'Calibri'},
                         header_font_attrs = {'bold': True, 'color_rgb': RGBColor(34, 64, 97)},
                         border_kwargs = {},
                         include_internal_cell_borders = True,
                         overwrite_formatting = True):
    num_rows, num_cols = dataframe.shape
    if isinstance(dataframe.index, pandas.MultiIndex):
        raise RuntimeError('Cannot yet cope with MultiIndex in rows')
    num_indexes = 1
    
    num_header_rows = get_index_numlevels(dataframe.columns)
        
    if (num_rows + num_header_rows) != len(pptx_table.rows):
        raise RuntimeError('Need %d rows but PPTX table has %d'
                           % (num_rows + num_header_rows, len(pptx_table.rows)))
    if (num_cols + num_indexes) != len(pptx_table.columns):
        raise RuntimeError('Need %d columns but PPTX table has %d'
                           % (num_cols + num_indexes, len(pptx_table.columns)))
        
    header_font_attrs = dict(header_font_attrs, **font_attrs)
    last_col = len(dataframe.columns.values)-1
    for i in range(num_header_rows):
        #headers
        prev_header = no_prev_header = '##special missing value'
        first_merged_cell = 0
        mergeable_cell_count = 0
        for c, header_name in enumerate(dataframe.columns.values):
            col_name = header_name[i] if num_header_rows > 1 else header_name
            if prev_header == no_prev_header or prev_header != col_name:
                if mergeable_cell_count > 0:
                    pptx_table.cell(i, first_merged_cell + num_indexes).merge(pptx_table.cell(i, first_merged_cell + mergeable_cell_count + num_indexes))
                prev_header = col_name
                first_merged_cell = c
                mergeable_cell_count = 0
                
                cell = pptx_table.cell(i, c + num_indexes)
                set_cell_text(cell, 
                              format_cell_text(col_name, **format_opts), 
                              overwrite_formatting = overwrite_formatting)
                if overwrite_formatting:
                    set_cell_font_attrs(cell, **header_font_attrs)
                    if include_internal_cell_borders:
                        borders = 'LRTB'
                    else:
                        borders = '' + ('L' if c == 0 else '') + ('R' if c == last_col else '') + ('T' if i == 0 else '')
                    set_cell_borders(cell, borders = borders, **border_kwargs)
            else:
                mergeable_cell_count += 1
        if mergeable_cell_count > 0:
            pptx_table.cell(i, first_merged_cell + num_indexes).merge(pptx_table.cell(i, first_merged_cell + mergeable_cell_count + num_indexes))
    
    last_row = num_rows-1
    last_col = num_cols-1
    for c in range(num_cols):
        #set column widths
        if isinstance(col_width, numbers.Number):
            w = DIST_METRIC(col_width)
        else:
            w = DIST_METRIC(col_width[c + 1])
        if overwrite_formatting:
            pptx_table.columns[c + num_indexes].width = w

        #body cells
        for r in range(num_rows):
            cell = pptx_table.cell(r + num_header_rows, c + num_indexes)
            set_cell_text(cell, 
                          format_cell_text(dataframe.iloc[r, c], **format_opts), 
                          overwrite_formatting = overwrite_formatting)
            if overwrite_formatting:
                set_cell_font_attrs(cell, **font_attrs)
                if include_internal_cell_borders:
                    borders = 'LRTB'
                else:
                    borders = '' + ('R' if c == last_col else '') + ('B' if r == last_row else '')
                set_cell_borders(cell, borders = borders, **border_kwargs)

    #index
    for r in range(num_rows):
        cell = pptx_table.cell(r + num_header_rows, 0)
        set_cell_text(cell, 
                      format_cell_text(dataframe.index[r], **format_opts), 
                      overwrite_formatting = overwrite_formatting)

        if overwrite_formatting:
            set_cell_font_attrs(cell, **header_font_attrs)
            if include_internal_cell_borders:
                borders = 'LRTB'
            else:
                borders = 'L' + ('B' if r == last_row else '')
            set_cell_borders(cell, borders = borders, **border_kwargs)
    
    if overwrite_formatting:
        pptx_table.columns[0].width = DIST_METRIC(col_width if isinstance(col_width, numbers.Number) else col_width[0])
    
    #index name
    for i in range(num_header_rows):
        cell = pptx_table.cell(i, 0)
        set_cell_text(cell, 
                      format_cell_text(dataframe.index.name if dataframe.index.name is not None else '', **format_opts), 
                      overwrite_formatting = overwrite_formatting)
        if overwrite_formatting:
            set_cell_font_attrs(cell, **header_font_attrs)
            if include_internal_cell_borders:
                borders = 'LRTB'
            else:
                borders = 'L' + ('T' if i == 0 else '')
            set_cell_borders(cell, borders = borders, **border_kwargs)
        
def create_pptx_table(pptx_slide, dataframe, left, top, col_width, row_height, include_internal_cell_borders, **write_kwargs):
    num_rows, num_cols = dataframe.shape
    if isinstance(dataframe.index, pandas.MultiIndex):
        raise RuntimeError('Cannot yet cope with MultiIndex rows')
    num_indexes = 1

    num_header_rows = get_index_numlevels(dataframe.columns)

    width = DIST_METRIC(col_width * num_cols if isinstance(col_width, numbers.Number) else sum(col_width))
    height = DIST_METRIC(row_height * num_rows)

    table_shape = pptx_slide.shapes.add_table(num_rows + num_header_rows, num_cols + num_indexes, 
                             DIST_METRIC(left), DIST_METRIC(top), 
                             width, height)
    
    table = table_shape.table
    write_pptx_dataframe(dataframe, table, col_width = col_width, include_internal_cell_borders = include_internal_cell_borders, **write_kwargs)
    return table_shape
    
def chart_to_file(chart_obj, img_file: str):
    if hasattr(chart_obj, 'write_image'):
        #ply_pd plots
        chart_obj.write_image(img_file, scale=2, width = chart_obj.width, height = chart_obj.height)
    else:
        #matplotlib plots...
        chart_obj.figure.savefig(img_file, dpi = 300, bbox_inches = 'tight')

def SubElement(parent, tagname, **kwargs):
        element = OxmlElement(tagname)
        element.attrib.update(kwargs)
        parent.append(element)
        return element
