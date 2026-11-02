"""Worksheet copying that preserves editable content and presentation."""
from copy import copy, deepcopy
import io
import os
import re

import openpyxl
from openpyxl.formula.translate import Translator
from engine.safety import atomic_save


def copy_sheet(source, target, rows=None):
    rows = list(rows if rows is not None else range(1, source.max_row + 1))
    row_map = {old: new for new, old in enumerate(rows, 1)}
    for old, new in row_map.items():
        for cell in source[old]:
            dst = target.cell(new, cell.column)
            value = cell.value
            if cell.data_type == 'f' and isinstance(value, str) and new != old:
                value = Translator(value, origin=cell.coordinate).translate_formula(dst.coordinate)
            dst.value = value
            if cell.has_style:
                for attr in ('font', 'fill', 'border', 'alignment', 'protection'):
                    setattr(dst, attr, copy(getattr(cell, attr)))
                dst.number_format = cell.number_format
            if cell.hyperlink: dst._hyperlink = copy(cell.hyperlink)
            if cell.comment: dst.comment = copy(cell.comment)
        dim = copy(source.row_dimensions[old]); dim.index = new
        target.row_dimensions[new] = dim
    for key, dim in source.column_dimensions.items():
        target.column_dimensions[key] = copy(dim)
    for merged in source.merged_cells.ranges:
        if all(r in row_map for r in range(merged.min_row, merged.max_row+1)):
            target.merge_cells(start_row=row_map[merged.min_row], end_row=row_map[merged.max_row],
                               start_column=merged.min_col, end_column=merged.max_col)
    target.sheet_format = copy(source.sheet_format)
    target.sheet_properties = deepcopy(source.sheet_properties)
    target.page_setup = copy(source.page_setup)
    target.page_margins = copy(source.page_margins)
    target.print_options = copy(source.print_options)
    target.HeaderFooter = deepcopy(source.HeaderFooter)
    target.freeze_panes = source.freeze_panes
    target.sheet_view.showGridLines = source.sheet_view.showGridLines
    if len(rows) == source.max_row:
        target.auto_filter = deepcopy(source.auto_filter)
        target.data_validations = deepcopy(source.data_validations)
        target.conditional_formatting = deepcopy(source.conditional_formatting)
        for chart in source._charts: target.add_chart(deepcopy(chart))
        for image in source._images:
            # openpyxl closes an image's original stream during save; retain a new stream.
            data = image._data()
            image.ref = io.BytesIO(data)
            clone = openpyxl.drawing.image.Image(io.BytesIO(data))
            clone.anchor = deepcopy(image.anchor)
            target.add_image(clone)


def merge_workbooks(paths, out_path):
    result = openpyxl.Workbook(); result.remove(result.active)
    for path in paths:
        wb = openpyxl.load_workbook(path, data_only=False)
        try:
            mapping = {}
            for source in wb:
                title = re.sub(r'[\\/:*?"<>|\[\]]','_', os.path.splitext(os.path.basename(path))[0]+'_'+source.title)[:31]
                original = title; i = 1
                while title.casefold() in {n.casefold() for n in result.sheetnames}:
                    suffix = f'_{i}'; title = original[:31-len(suffix)]+suffix; i += 1
                mapping[source.title] = title
                copy_sheet(source, result.create_sheet(title))
            from openpyxl.formula import Tokenizer
            for name in mapping.values():
                for row in result[name]:
                    for cell in row:
                        if cell.data_type != 'f' or not isinstance(cell.value,str): continue
                        tokens = Tokenizer(cell.value).items
                        for token in tokens:
                            if token.type == 'OPERAND' and token.subtype == 'RANGE' and '!' in token.value:
                                prefix, reference = token.value.rsplit('!',1)
                                source_name = prefix.strip("'").replace("''", "'")
                                if source_name in mapping:
                                    token.value = "'"+mapping[source_name].replace("'","''")+"'!"+reference
                        cell.value = '=' + ''.join(t.value for t in tokens)
        finally:
            wb.close()
    if not result.sheetnames:
        raise ValueError('没有可合并的工作表')
    atomic_save(result.save, out_path)
    count = len(result.sheetnames); result.close()
    return count
