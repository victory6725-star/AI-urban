import sys
import json
import os
import win32com.client

def generate_hwp(data_path, hwp_path):
    with open(data_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
        
    layers = data.get('layers', {})
    boundary_area = data.get('boundary_area', 1.0)
    boundary_layer = data.get('boundary_layer', None)
    
    # Format helpers
    def fmt_m2(val):
        return f"{val:,.0f}" if val > 0 else "-"
        
    def fmt_pct(val):
        if boundary_area > 0 and val > 0:
            return f"{(val / boundary_area * 100):.1f}"
        return "-"

    # Initialize HWP (Run in background quietly)
    hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
    try:
        hwp.XHwpWindows.Item(0).Visible = False
    except Exception:
        pass
        
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    hwp.Clear(1)
    
    # Helper to insert text using OLE Action
    def insert_text_helper(text):
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = text
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    
    # 1. Insert Title
    insert_text_helper("토지이용계획 면적산출표\r\n\r\n")
    
    # 2. Build Table Data (Flat list layout)
    num_cols = 4
    table_data = [
        ["구분 (레이어명)", "면적 (㎡)", "구성비 (%)", "비고"],
        ["합계", fmt_m2(boundary_area), "100.0", ""]
    ]
    
    for name, val in sorted(layers.items()):
        # Skip boundary layer itself in rows
        if boundary_layer and name == boundary_layer:
            continue
        if val <= 0:
            continue
        table_data.append([name, fmt_m2(val), fmt_pct(val), ""])

    num_rows = len(table_data)
    
    # 3. Create Table using standard CreateAction
    act = hwp.CreateAction("TableCreate")
    pset = act.CreateSet()
    act.GetDefault(pset)
    pset.SetItem("Rows", num_rows)
    pset.SetItem("Cols", num_cols)
    pset.SetItem("WidthType", 1)  # 단너비 맞춤
    pset.SetItem("HeightType", 0) # 자동 높이
    
    act.Execute(pset)
    
    # 4. Fill cells
    # The caret starts at the first cell
    for row in table_data:
        for val in row:
            insert_text_helper(val)
            hwp.HAction.Run("TableRightCell")
            
    # 5. Clean up selection, Save, and Close background HWP process
    hwp.HAction.Run("Cancel")
    hwp.SaveAs(hwp_path)
    hwp.Quit()

if __name__ == '__main__':
    if len(sys.argv) < 3:
        sys.exit(1)
    generate_hwp(sys.argv[1], sys.argv[2])
