"""
국내 터미널 재고표 취합 (클라우드 호환 버전)
Usage: python merge_stock_cloud.py <YYYYMMDD> <input_dir> <output_path>

로컬 merge_stock.py와의 차이 (클라우드/리눅스 샌드박스용):
  - win32com(Excel COM) 제거. HJNC Report 수식은 openpyxl data_only(발신처가
    저장 시 캐시해둔 값)로 해석한다. Excel 설치가 없는 리눅스에서 동작.
  - .xls 처리는 xlrd==1.2.0(formatting_info 지원)에 의존 → requirements.txt 핀.
  - Gmail/Drive 로직 없음. 이메일 검색/다운로드/발송은 routine 에이전트가
    Gmail 커넥터로 수행하고, 이 스크립트는 '취합'만 담당한다.
"""
import sys
import os
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy
import xlrd

# (filename, [(src_sheet, dst_sheet_name), ...])
TERMINALS = [
    ("HBGT", [(0, "main"), ("Summary", "detail")]),
    ("BPTS", [("Sheet1", "main"), ("Sheet2", "detail")]),
    ("BPTG", [("KMTC", "main"), ("LIST_M", "detail")]),
    ("HJIT", [(0, None)]),
    ("ICT",  [(0, None)]),
    ("SNCT", [(0, None)]),
    ("HJNC", [("Report", "main"), ("List", "detail")]),
    ("PNCT", [("Move", "main"), ("List", "detail")]),
    ("DGT",  [("Summary", "main"), ("M-List", "detail")]),
]


# ============================================================
# .xls -> openpyxl converter with formatting (xlrd 1.2.0)
# ============================================================
def xls_to_xlsx(xls_path):
    old_wb = xlrd.open_workbook(xls_path, formatting_info=True)
    new_wb = openpyxl.Workbook()
    new_wb.remove(new_wb.active)
    xf_list, font_list, format_map = old_wb.xf_list, old_wb.font_list, old_wb.format_map

    for si in range(old_wb.nsheets):
        old_ws = old_wb.sheet_by_index(si)
        new_ws = new_wb.create_sheet(title=old_ws.name)
        if hasattr(old_ws, 'colinfo_map'):
            for ci, info in old_ws.colinfo_map.items():
                new_ws.column_dimensions[get_column_letter(ci + 1)].width = info.width / 256.0
        if hasattr(old_ws, 'rowinfo_map'):
            for ri, info in old_ws.rowinfo_map.items():
                if info.height: new_ws.row_dimensions[ri + 1].height = info.height / 20.0

        for ri in range(old_ws.nrows):
            for ci in range(old_ws.ncols):
                cell = old_ws.cell(ri, ci)
                nc = new_ws.cell(row=ri + 1, column=ci + 1)
                if cell.ctype == xlrd.XL_CELL_EMPTY:
                    nc.value = None
                elif cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        from datetime import datetime
                        nc.value = datetime(*xlrd.xldate_as_tuple(cell.value, old_wb.datemode))
                    except Exception:
                        nc.value = cell.value
                else:
                    nc.value = cell.value
                try:
                    xi = cell.xf_index
                    if xi is not None and xi < len(xf_list):
                        xf = xf_list[xi]
                        if xf.font_index < len(font_list):
                            f = font_list[xf.font_index]
                            nc.font = openpyxl.styles.Font(
                                name=f.name, size=f.height/20.0 if f.height else None,
                                bold=f.bold, italic=f.italic,
                                underline='single' if f.underline_type else None,
                                strike=f.struck_out, color=_xc(f.colour_index, old_wb))
                        nc.alignment = openpyxl.styles.Alignment(
                            horizontal=_ha(xf.alignment.hor_align),
                            vertical=_va(xf.alignment.vert_align),
                            wrap_text=bool(xf.alignment.text_wrapped))
                        fk = xf.format_key
                        if fk in format_map: nc.number_format = format_map[fk].format_str
                        bg = _bg(xf, old_wb)
                        if bg: nc.fill = openpyxl.styles.PatternFill(start_color=bg, end_color=bg, fill_type='solid')
                        bd = _bd(xf, old_wb)
                        if bd: nc.border = bd
                except Exception:
                    pass

        for cr in old_ws.merged_cells:
            rlo, rhi, clo, chi = cr
            new_ws.merge_cells(start_row=rlo+1, start_column=clo+1, end_row=rhi, end_column=chi)
    return new_wb


def _xc(ci, wb):
    if ci is None or ci in (0x7FFF, 64): return None
    try:
        cm = wb.colour_map
        if ci in cm and cm[ci]:
            r, g, b = cm[ci]
            return openpyxl.styles.Color(rgb=f"FF{r:02X}{g:02X}{b:02X}")
    except Exception:
        pass
    return None

def _ha(v): return {0:'general',1:'left',2:'center',3:'right',4:'fill',5:'justify',6:'centerContinuous',7:'distributed'}.get(v,'general')
def _va(v): return {0:'top',1:'center',2:'bottom',3:'justify',4:'distributed'}.get(v,'bottom')

def _bg(xf, wb):
    try:
        bi = xf.background.pattern_colour_index
        if bi is None or bi in (64, 0x7FFF): return None
        cm = wb.colour_map
        if bi in cm and cm[bi]:
            r, g, b = cm[bi]
            if (r, g, b) == (255, 255, 255): return None
            return f"FF{r:02X}{g:02X}{b:02X}"
    except Exception:
        pass
    return None

def _bs(bi, ci, wb):
    if bi == 0: return openpyxl.styles.Side()
    sm = {1:'thin',2:'medium',3:'dashed',4:'dotted',5:'thick',6:'double',7:'hair'}
    return openpyxl.styles.Side(style=sm.get(bi,'thin'), color=_xc(ci, wb))

def _bd(xf, wb):
    try:
        b = xf.border
        if any([b.top_line_style, b.bottom_line_style, b.left_line_style, b.right_line_style]):
            return openpyxl.styles.Border(
                left=_bs(b.left_line_style, b.left_colour_index, wb),
                right=_bs(b.right_line_style, b.right_colour_index, wb),
                top=_bs(b.top_line_style, b.top_colour_index, wb),
                bottom=_bs(b.bottom_line_style, b.bottom_colour_index, wb))
    except Exception:
        pass
    return None


# ============================================================
# Sheet copy with styles
# ============================================================
def copy_sheet(src_ws, dst_ws):
    for cl, cd in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[cl].width = cd.width
        dst_ws.column_dimensions[cl].hidden = cd.hidden
    for rn, rd in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[rn].height = rd.height
        dst_ws.row_dimensions[rn].hidden = rd.hidden
    for row in src_ws.iter_rows():
        for cell in row:
            if isinstance(cell, openpyxl.cell.cell.MergedCell): continue
            nc = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                nc.font = copy(cell.font); nc.border = copy(cell.border)
                nc.fill = copy(cell.fill); nc.number_format = cell.number_format
                nc.protection = copy(cell.protection); nc.alignment = copy(cell.alignment)
    for mr in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(mr))


def copy_sheet_values_only(src_style_ws, src_data_ws, dst_ws):
    """Copy styles from src_style_ws but values from src_data_ws (formula-resolved/cached)."""
    for cl, cd in src_style_ws.column_dimensions.items():
        dst_ws.column_dimensions[cl].width = cd.width
        dst_ws.column_dimensions[cl].hidden = cd.hidden
    for rn, rd in src_style_ws.row_dimensions.items():
        dst_ws.row_dimensions[rn].height = rd.height
        dst_ws.row_dimensions[rn].hidden = rd.hidden
    for row in src_style_ws.iter_rows():
        for cell in row:
            if isinstance(cell, openpyxl.cell.cell.MergedCell): continue
            dc = src_data_ws.cell(row=cell.row, column=cell.column)
            nc = dst_ws.cell(row=cell.row, column=cell.column, value=dc.value)
            if cell.has_style:
                nc.font = copy(cell.font); nc.border = copy(cell.border)
                nc.fill = copy(cell.fill); nc.number_format = cell.number_format
                nc.protection = copy(cell.protection); nc.alignment = copy(cell.alignment)
    for mr in src_style_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(mr))


# ============================================================
# Main
# ============================================================
def main():
    if len(sys.argv) < 4:
        print("Usage: python merge_stock_cloud.py <YYYYMMDD> <input_dir> <output_path>")
        sys.exit(1)

    date_str = sys.argv[1]
    input_dir = sys.argv[2]
    output_path = sys.argv[3]

    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    failed = []

    for terminal, sheet_specs in TERMINALS:
        fpath = None
        for ext in ['.xlsx', '.xls']:
            candidate = os.path.join(input_dir, f"{terminal}{ext}")
            if os.path.exists(candidate):
                fpath = candidate
                break
        if not fpath:
            print(f"  SKIP {terminal}: file not found")
            continue

        print(f"  Processing {terminal}: {os.path.basename(fpath)}")

        try:
            if fpath.endswith('.xls'):
                wb_src = xls_to_xlsx(fpath)
            else:
                wb_src = openpyxl.load_workbook(fpath)

            # HJNC Report: 수식값을 data_only(캐시값)로 해석 (win32com 대체)
            hjnc_values_wb = None
            if terminal == "HJNC" and fpath.endswith('.xlsx'):
                try:
                    hjnc_values_wb = openpyxl.load_workbook(fpath, data_only=True)
                except Exception as e:
                    print(f"  HJNC data_only load error: {e}")

            staged = []
            for src_sheet, suffix in sheet_specs:
                if isinstance(src_sheet, int):
                    src_name = wb_src.sheetnames[src_sheet]
                else:
                    src_name = src_sheet
                dst_name = f"{terminal}_{suffix}" if suffix else terminal

                if terminal == "HJNC" and src_name == "Report" and hjnc_values_wb:
                    src_style = wb_src[src_name]
                    src_data = hjnc_values_wb[src_name]
                    staged.append((dst_name, "values", src_style, src_data))
                else:
                    staged.append((dst_name, "copy", wb_src[src_name], None))

            for dst_name, mode, a, b in staged:
                dst_ws = wb_out.create_sheet(title=dst_name[:31])
                if mode == "values":
                    copy_sheet_values_only(a, b, dst_ws)
                    print(f"    {a.title} -> {dst_name} (values-only / cached)")
                else:
                    copy_sheet(a, dst_ws)
                    print(f"    {a.title} -> {dst_name}")

        except Exception as e:
            failed.append(terminal)
            print(f"  FAIL {terminal}: {e}")

    wb_out.save(output_path)
    print(f"\nSaved: {output_path}")
    print(f"Sheets: {wb_out.sheetnames}")
    if failed:
        print(f"FAILED:{','.join(failed)}")


if __name__ == "__main__":
    main()
