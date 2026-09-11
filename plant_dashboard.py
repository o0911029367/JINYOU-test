import os
import re
import pandas as pd
import openpyxl
from datetime import datetime

def main():
    base_dir = r"C:\Users\JIN YOU\Desktop\AI賦能智造製造業數據驅動與人機協作實戰\GITHUB TEST\資料來源"
    
    if not os.path.exists(base_dir):
        print(f"Directory not found: {base_dir}")
        return

    # Dynamic file scanning
    excel_path = None
    for fname in os.listdir(base_dir):
        if '訂單' in fname and fname.endswith('.xlsx') and not fname.startswith('~$'):
            excel_path = os.path.join(base_dir, fname)
            break

    if not excel_path:
        print(f"Error:找不到訂單 Excel 檔案於 {base_dir}")
        return

    print(f"正在讀取檔案與解析製程色塊：{excel_path}")
    
    orders_df = pd.read_excel(excel_path, sheet_name="訂單")

    # Read rich text / cell font colors for column T
    wb = openpyxl.load_workbook(excel_path, data_only=True, rich_text=True)
    sheet = wb['訂單']
    
    uncompleted_col = []
    
    for r in range(2, len(orders_df) + 2):
        cell = sheet.cell(row=r, column=20)
        val = cell.value
        cell_color = str(cell.font.color.rgb) if cell.font and cell.font.color else 'None'
        
        uncompleted_parts = []
        
        # If the entire cell font color is green, it's fully completed -> uncompleted is empty
        if '00B050' in cell_color or '008000' in cell_color or 'FF00B050' in cell_color:
            uncompleted_parts = []
        elif hasattr(val, '__iter__') and not isinstance(val, str):
            for block in val:
                if isinstance(block, str):
                    uncompleted_parts.append(block)
                else:
                    text = getattr(block, 'text', str(block))
                    color = block.font.color.rgb if hasattr(block, 'font') and block.font and block.font.color else 'None'
                    # Keep uncompleted if NOT green
                    if not ('00B050' in str(color) or '008000' in str(color) or 'FF00B050' in str(color)):
                        uncompleted_parts.append(text)
        else:
            # If entire cell is red or plain text without green parts
            if 'FF0000' in cell_color or 'FFFF0000' in cell_color or cell_color == 'None':
                uncompleted_parts.append(str(val) if pd.notna(val) else '')
            else:
                uncompleted_parts = []
            
        uncompleted_col.append(''.join(uncompleted_parts).strip())

    orders_df['未完成製程'] = uncompleted_col

    # Clean numeric columns
    orders_df['order_qty_num'] = pd.to_numeric(orders_df['訂單數量'], errors='coerce').fillna(0)
    orders_df['unshipped_num'] = pd.to_numeric(orders_df['未交'], errors='coerce').fillna(0)

    # Parse production qty from column R (生產數量) summing numbers before hyphen
    def parse_production_qty(val):
        if pd.isna(val):
            return 0.0
        text = str(val)
        total = 0.0
        matches = re.findall(r'(\d+)\s*-[^\d\n]*', text)
        if not matches:
            matches = re.findall(r'(\d+)\s*-', text)
        for m in matches:
            total += float(m)
        return total

    orders_df['produced_num'] = orders_df['生產數量'].apply(parse_production_qty)

    # Filter 1: Only keep rows where unshipped > 0
    active_orders = orders_df[orders_df['unshipped_num'] > 0].copy()

    # Filter 2: Exclude non-production / test items
    exclude_keywords = ['暫停', '樣品', '待', '未排', '停', '測試', 'test']
    
    def is_valid_delivery_date(val):
        if pd.isna(val):
            return False
        s = str(val)
        for kw in exclude_keywords:
            if kw in s:
                return False
        return True

    active_orders = active_orders[active_orders['交貨日'].apply(is_valid_delivery_date)].copy()

    # Mask customer names for presentation privacy
    unique_customers = active_orders['客戶'].dropna().unique()
    customer_mask_map = {cust: f"客戶 {chr(65 + i)}" for i, cust in enumerate(unique_customers)}
    active_orders['masked_customer'] = active_orders['客戶'].map(customer_mask_map).fillna("客戶 X")

    ref_date = pd.Timestamp.today().normalize()
    
    def parse_and_calculate_risk(date_val):
        if pd.isna(date_val):
            return "PENDING", "", "其他月份"
        try:
            date_str = str(date_val).split('\n')[0].strip()
            d = pd.to_datetime(date_str, errors='coerce')
            if pd.isna(d):
                return "PENDING", str(date_val), "其他月份"
            
            delta_days = (d - ref_date).days
            if delta_days < 0:
                risk = "RED-OVERDUE"
            elif delta_days <= 7:
                risk = "YELLOW-7DAYS"
            else:
                risk = "GREEN-NORMAL"
                
            month_label = d.strftime('%Y年 %m月')
            return risk, d.strftime('%Y-%m-%d'), month_label
        except:
            return "PENDING", str(date_val), "其他月份"

    risk_results = active_orders['交貨日'].apply(parse_and_calculate_risk)
    active_orders['risk_level'] = [r[0] for r in risk_results]
    active_orders['clean_date'] = [r[1] for r in risk_results]
    active_orders['delivery_month'] = [r[2] for r in risk_results]

    # Map risk level to pure emoji lights
    risk_symbol_map = {
        "RED-OVERDUE": "🔴",
        "YELLOW-7DAYS": "🟡",
        "GREEN-NORMAL": "🟢",
        "PENDING": "⚪"
    }
    active_orders['risk_symbol'] = active_orders['risk_level'].map(risk_symbol_map)

    # Columns: 交期風險, 預交日, 客戶, 產品圖號, 訂單號碼, 未交數, 生產數, 未完成製程, delivery_month, risk_code
    summary_df = active_orders[['risk_symbol', 'clean_date', 'masked_customer', '圖號', '訂單號碼', 'unshipped_num', 'produced_num', '未完成製程', 'delivery_month', 'risk_level']].copy()
    summary_df.columns = ['交期風險', '預定交貨日', '客戶', '產品圖號', '訂單編號', '未交數量', '已生產數量', '未完成製程', 'delivery_month', 'risk_code']

    risk_order = {"RED-OVERDUE": 1, "YELLOW-7DAYS": 2, "GREEN-NORMAL": 3, "PENDING": 4}
    summary_df['risk_sort'] = summary_df['risk_code'].map(risk_order)
    
    # Sort by delivery month then risk then unshipped qty
    summary_df = summary_df.sort_values(by=['delivery_month', 'risk_sort', '未交數量'], ascending=[True, True, False])

    # Export columns for CSV
    export_df = summary_df[['交期風險', '預定交貨日', '客戶', '產品圖號', '訂單編號', '未交數量', '已生產數量', '未完成製程', 'delivery_month']].copy()

    # Save CSV report
    output_csv_path = os.path.join(os.path.dirname(__file__), "plant_summary_dashboard.csv")
    export_df.to_csv(output_csv_path, index=False, encoding='utf-8-sig')

    # Save styled HTML report grouped by month with exact column widths
    output_html_path = os.path.join(os.path.dirname(__file__), "plant_summary_dashboard.html")
    
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>廠務未交訂單交期與製程進度看板</title>
        <style>
            body {{ font-family: "Microsoft JhengHei", Arial, sans-serif; margin: 20px; background-color: #f8f9fa; }}
            h2 {{ color: #333; margin-bottom: 5px; }}
            h3 {{ color: #495057; margin-top: 30px; border-bottom: 2px solid #6c757d; padding-bottom: 5px; }}
            .controls {{ margin: 15px 0; display: flex; justify-content: space-between; align-items: center; }}
            .search-box {{ padding: 8px 12px; width: 300px; border: 1px solid #ced4da; border-radius: 4px; font-size: 14px; }}
            table {{ border-collapse: collapse; width: 100%; background: white; box-shadow: 0 2px 5px rgba(0,0,0,0.1); margin-bottom: 20px; table-layout: fixed; }}
            th, td {{ border: 1px solid #dee2e6; padding: 10px 8px; text-align: left; font-size: 13px; overflow: hidden; text-overflow: ellipsis; word-wrap: break-word; }}
            th {{ background-color: #343a40; color: white; cursor: pointer; user-select: none; position: relative; }}
            th:hover {{ background-color: #495057; }}
            tr.red {{ background-color: #f8d7da; color: #721c24; font-weight: bold; }}
            tr.yellow {{ background-color: #fff3cd; color: #856404; }}
            tr.green {{ background-color: #d4edda; color: #155724; }}
            tr.pending {{ background-color: #e2e3e5; color: #383d41; }}
            .light-cell {{ text-align: center; font-size: 16px; }}
            .uncompleted-text {{ color: #721c24; font-weight: bold; }}
            
            /* Column Widths */
            .col-risk {{ width: 35px; text-align: center; }}
            .col-date {{ width: 80px; }}
            .col-cust {{ width: 50px; }}
            .col-part {{ width: 150px; }}
            .col-order {{ width: 120px; }}
            .col-unshipped {{ width: 50px; text-align: right; }}
            .col-produced {{ width: 50px; text-align: right; }}
            .col-uncompleted {{ width: 500px; }}
        </style>
        <script>
            function filterTable() {{
                var input = document.getElementById("searchInput");
                var filter = input.value.toLowerCase();
                var tables = document.getElementsByClassName("month-table");

                for (var t = 0; t < tables.length; t++) {{
                    var tr = tables[t].getElementsByTagName("tr");
                    for (var i = 1; i < tr.length; i++) {{
                        var td = tr[i].getElementsByTagName("td");
                        var match = false;
                        for (var j = 0; j < td.length; j++) {{
                            if (td[j]) {{
                                if (td[j].innerHTML.toLowerCase().indexOf(filter) > -1) {{
                                    match = true;
                                    break;
                                }}
                            }}
                        }}
                        tr[i].style.display = match ? "" : "none";
                    }}
                }}
            }}
        </script>
    </head>
    <body>
        <h2>廠務生產現場未交訂單交期看板 (依月份彙整・發表演示版)</h2>
        <div class="controls">
            <p>評估基準日（今日）：<strong>{ref_date.strftime('%Y-%m-%d')}</strong> | 有效未交訂單總計：<strong>{len(summary_df)}</strong> 筆（客戶已遮罩、已排除測試/樣品）</p>
            <input type="text" id="searchInput" class="search-box" onkeyup="filterTable()" placeholder="🔍 快速搜尋任何關鍵字（客戶、圖號、狀態...）">
        </div>
    """

    # Group by month
    grouped = summary_df.groupby('delivery_month')
    for month_name, group in grouped:
        html_content += f"""
        <h3>📅 預定交貨月份：{month_name} （共 {len(group)} 筆）</h3>
        <table class="month-table">
            <thead>
                <tr>
                    <th class="col-risk">交期風險</th>
                    <th class="col-date">預交日</th>
                    <th class="col-cust">客戶</th>
                    <th class="col-part">產品圖號</th>
                    <th class="col-order">訂單號碼</th>
                    <th class="col-unshipped">未交數</th>
                    <th class="col-produced">生產數</th>
                    <th class="col-uncompleted">未完成製程</th>
                </tr>
            </thead>
            <tbody>
        """
        for _, row in group.iterrows():
            symbol = row['交期風險']
            risk_code = row['risk_code']
            tr_class = "pending"
            if risk_code == "RED-OVERDUE":
                tr_class = "red"
            elif risk_code == "YELLOW-7DAYS":
                tr_class = "yellow"
            elif risk_code == "GREEN-NORMAL":
                tr_class = "green"

            html_content += f"""
                <tr class="{tr_class}">
                    <td class="col-risk light-cell">{symbol}</td>
                    <td class="col-date">{row['預定交貨日']}</td>
                    <td class="col-cust">{row['客戶']}</td>
                    <td class="col-part">{row['產品圖號']}</td>
                    <td class="col-order">{row['訂單編號'] if pd.notna(row['訂單編號']) else ''}</td>
                    <td class="col-unshipped">{row['未交數量']:,.0f}</td>
                    <td class="col-produced">{row['已生產數量']:,.0f}</td>
                    <td class="col-uncompleted uncompleted-text">{row['未完成製程'] if pd.notna(row['未完成製程']) else ''}</td>
                </tr>
            """
        html_content += """
            </tbody>
        </table>
        """

    html_content += """
    </body>
    </html>
    """

    with open(output_html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"\nSuccessfully generated HTML colored dashboard: {output_html_path}")
    print(f"Successfully updated CSV report: {output_csv_path}")

if __name__ == "__main__":
    main()
