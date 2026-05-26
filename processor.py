import os
import tempfile
import xlrd
from openpyxl import Workbook
from openpyxl.styles import Font


# -------------------------- 工具函数 --------------------------
def read_xls_with_xlrd(file_path):
    """用 xlrd 读取 .xls 文件，返回二维列表（行、列，0-based）"""
    wb = xlrd.open_workbook(file_path)
    ws = wb.sheet_by_index(0)
    data = []
    for r in range(ws.nrows):
        row = []
        for c in range(ws.ncols):
            cell = ws.cell(r, c)
            if cell.ctype == xlrd.XL_CELL_EMPTY:
                row.append(None)
            elif cell.ctype == xlrd.XL_CELL_NUMBER:
                row.append(cell.value)
            else:
                row.append(str(cell.value))
        data.append(row)
    return data


def read_xlsx_with_openpyxl(file_path):
    """用 openpyxl 读取 .xlsx 文件，返回二维列表（0-based）"""
    from openpyxl import load_workbook
    wb = load_workbook(file_path, data_only=True)
    ws = wb.active
    data = []
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column, values_only=True):
        data.append(list(row))
    wb.close()
    return data


# -------------------------- Web版核心处理函数 --------------------------
def process_excel(file_bytes: bytes, file_name: str = "input.xlsx") -> bytes:
    """
    Web版入口：接收文件bytes，处理，返回结果bytes
    """
    # 根据后缀创建临时文件
    suffix = ".xls" if file_name.lower().endswith(".xls") else ".xlsx"
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    
    try:
        result_path = _process_single_file(tmp_path)
        
        # 读取结果文件返回bytes
        with open(result_path, "rb") as f:
            result_bytes = f.read()
        
        # 清理临时文件
        os.unlink(tmp_path)
        os.unlink(result_path)
        
        return result_bytes
        
    except Exception:
        # 出错也要清理
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        raise


def _process_single_file(file_path):
    """
    原process_excel逻辑，改为返回输出路径
    """
    base, ext = os.path.splitext(file_path)
    output_path = base + "_处理后.xlsx"

    # 1. 读取文件数据
    if ext.lower() == '.xls':
        raw_data = read_xls_with_xlrd(file_path)
    elif ext.lower() == '.xlsx':
        raw_data = read_xlsx_with_openpyxl(file_path)
    else:
        raise ValueError(f"不支持的文件格式: {ext}")

    if not raw_data:
        raise ValueError("文件为空")

    # 2. 检查 A1 是否以"实时数据"结尾
    a1_value = raw_data[0][0] if len(raw_data) > 0 and len(raw_data[0]) > 0 else None
    if a1_value and str(a1_value).strip().endswith("实时数据"):
        raise ValueError("文件是实时数据，已跳过处理。")

    # 3. 确定数据起始行
    if ext.lower() == '.xls':
        start_row = 4  # 1-based
    else:
        row6 = raw_data[5] if len(raw_data) >= 6 else []
        row7 = raw_data[6] if len(raw_data) >= 7 else []
        if any(v is not None for v in row6):
            start_row = 6
        elif any(v is not None for v in row7):
            start_row = 7
        else:
            raise ValueError("无法确定 .xlsx 文件的数据起始行")

    start_row_idx = start_row - 1
    max_row_idx = len(raw_data) - 1
    max_col_idx = max(len(row) for row in raw_data) - 1

    # 4. 构建列名映射
    def get_row_data(row_idx):
        return raw_data[row_idx] if row_idx < len(raw_data) else []

    row2 = get_row_data(1)
    row3 = get_row_data(2)

    col_map2 = {}
    col_map3 = {}
    for idx, val in enumerate(row2):
        if val:
            col_map2[str(val).strip()] = idx
    for idx, val in enumerate(row3):
        if val:
            col_map3[str(val).strip()] = idx

    # 5. 排放量列
    emission_cols = [idx for idx, val in enumerate(row3) if val and "排放量" in str(val)]
    if not emission_cols:
        raise ValueError("未找到排放量列")

    # 6. 体积列及参数
    volume_col = None
    concentration_offset = None
    divisor = None
    for col_name, col_idx in col_map2.items():
        if "废气排放量" in col_name and ("标立方米" in col_name or "Nm³" in col_name):
            volume_col = col_idx
            concentration_offset = 3
            divisor = 1000000
            break
        elif "废水排放量" in col_name and ("t" in col_name or "吨" in col_name):
            volume_col = col_idx
            concentration_offset = 2
            divisor = 1000
            break
    if volume_col is None:
        raise ValueError("未找到'废气排放量'或'废水排放量'列")

    # 7. 人工、自动、工况标记列
    manual_col = None
    auto_col = None
    condition_col = None
    for name, idx in col_map3.items():
        if "人工" in name:
            manual_col = idx
        if "自动" in name:
            auto_col = idx
    for name, idx in col_map2.items():
        if "工况标记" in name:
            condition_col = idx
            break

    # 8. 获取列数据的辅助函数
    def get_column_data(col_idx):
        if col_idx is None:
            return [None] * (max_row_idx - start_row_idx + 1)
        col_data = []
        for r in range(start_row_idx, max_row_idx + 1):
            if r < len(raw_data) and col_idx < len(raw_data[r]):
                col_data.append(raw_data[r][col_idx])
            else:
                col_data.append(None)
        return col_data

    vol_data = get_column_data(volume_col)
    manual_data = get_column_data(manual_col)
    auto_data = get_column_data(auto_col)
    cond_data = get_column_data(condition_col)

    # 9. 处理每个排放量列
    for emission_col_idx in emission_cols:
        conc_col_idx = emission_col_idx - concentration_offset
        if conc_col_idx < 0:
            continue

        conc_data = get_column_data(conc_col_idx)

        # 计算辅助列
        aux_vals = []
        for c, v in zip(conc_data, vol_data):
            try:
                c_num = float(c) if c not in (None, "") else None
                v_num = float(v) if v not in (None, "") else None
            except (ValueError, TypeError):
                c_num, v_num = None, None
            if c_num is None or v_num is None:
                aux = None
            else:
                raw = c_num * v_num / divisor
                aux = raw if raw >= 0 else 0
            aux_vals.append(aux)

        # 处理特殊标记
        final_vals = []
        any_aux_valid = any(v is not None for v in aux_vals)
        max_aux_filtered = None
        if manual_col is not None:
            valid_aux = []
            for i, aux in enumerate(aux_vals):
                man_text = str(manual_data[i]) if manual_data[i] is not None else ""
                if man_text not in ["校准", "日常维护", "调试", "故障"] and aux is not None:
                    valid_aux.append(aux)
            if valid_aux:
                max_aux_filtered = max(valid_aux)

        for i in range(len(aux_vals)):
            man_text = str(manual_data[i]) if manual_data[i] is not None else ""
            auto_text = str(auto_data[i]) if auto_data[i] is not None else ""
            cond_text = str(cond_data[i]) if cond_data[i] is not None else ""

            if cond_text and any(kw in cond_text for kw in ["停炉", "停运"]):
                final_vals.append(0 if any_aux_valid else "")
            elif (man_text and any(kw in man_text for kw in ["校准", "日常维护", "故障"])) or \
                 (auto_text and "核查比对" in auto_text):
                final_vals.append(max_aux_filtered if max_aux_filtered is not None else "")
            else:
                final_vals.append(aux_vals[i] if aux_vals[i] is not None else "")

        # 写回 raw_data
        for i, val in enumerate(final_vals):
            r = start_row_idx + i
            while len(raw_data) <= r:
                raw_data.append([])
            while len(raw_data[r]) <= emission_col_idx:
                raw_data[r].append(None)
            raw_data[r][emission_col_idx] = val

    # ========== 重新确定数据范围 ==========
    current_max_row_idx = len(raw_data) - 1
    avg_row_idx = current_max_row_idx + 1
    max_row_idx_final = current_max_row_idx + 2

    while len(raw_data) <= max_row_idx_final:
        raw_data.append([])

    # 构建筛选条件数组
    is_valid = [True] * (current_max_row_idx - start_row_idx + 1)
    for i in range(len(is_valid)):
        if auto_col is not None:
            val = auto_data[i] if i < len(auto_data) else None
            if str(val).strip() != "正常":
                is_valid[i] = False
        if manual_col is not None:
            val = manual_data[i] if i < len(manual_data) else None
            if val not in (None, "", " "):
                is_valid[i] = False
        if condition_col is not None:
            val = cond_data[i] if i < len(cond_data) else None
            if val not in (None, "", " "):
                is_valid[i] = False

    # 从第2列开始计算平均值和最大值
    for col_idx in range(1, max_col_idx + 1):
        col_vals = []
        for i, r in enumerate(range(start_row_idx, current_max_row_idx + 1)):
            if r < len(raw_data) and col_idx < len(raw_data[r]):
                val = raw_data[r][col_idx]
                if is_valid[i]:
                    try:
                        num_val = float(val)
                        if num_val != 0:
                            col_vals.append(num_val)
                    except (ValueError, TypeError):
                        pass
        avg = sum(col_vals) / len(col_vals) if col_vals else ""
        mx = max(col_vals) if col_vals else ""

        while len(raw_data[avg_row_idx]) <= col_idx:
            raw_data[avg_row_idx].append(None)
        raw_data[avg_row_idx][col_idx] = avg

        while len(raw_data[max_row_idx_final]) <= col_idx:
            raw_data[max_row_idx_final].append(None)
        raw_data[max_row_idx_final][col_idx] = mx

    # 第一列添加标签
    while len(raw_data[avg_row_idx]) < 1:
        raw_data[avg_row_idx].insert(0, None)
    raw_data[avg_row_idx][0] = "平均值"
    while len(raw_data[max_row_idx_final]) < 1:
        raw_data[max_row_idx_final].insert(0, None)
    raw_data[max_row_idx_final][0] = "最大值"

    # ========== 写入新文件，并设置字体颜色 ==========
    wb = Workbook()
    ws = wb.active

    red_font = Font(color="0000FF")
    blue_font = Font(color="FF0000")

    def safe_get_cell(row_idx, col_idx):
        if col_idx is None:
            return None
        if row_idx < len(raw_data) and col_idx < len(raw_data[row_idx]):
            return raw_data[row_idx][col_idx]
        return None

    for r, row in enumerate(raw_data, start=1):
        for c, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=val)
            if r >= start_row and (c - 1) in emission_cols and r <= current_max_row_idx + 1:
                row_idx = r - 1
                man_val = safe_get_cell(row_idx, manual_col)
                auto_val = safe_get_cell(row_idx, auto_col)
                cond_val = safe_get_cell(row_idx, condition_col)

                man_text = str(man_val) if man_val is not None else ""
                auto_text = str(auto_val) if auto_val is not None else ""
                cond_text = str(cond_val) if cond_val is not None else ""

                if cond_text and any(kw in cond_text for kw in ["停炉", "停运"]):
                    cell.font = red_font
                elif (man_text and any(kw in man_text for kw in ["校准", "日常维护", "故障"])) or \
                     (auto_text and "核查比对" in auto_text):
                    cell.font = blue_font

    # 冻结窗格
    ws.freeze_panes = f'A{start_row}'
    wb.save(output_path)
    wb.close()

    return output_path