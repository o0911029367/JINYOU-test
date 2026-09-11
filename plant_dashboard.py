import os
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

    print(f"正在讀取檔案與整理資料：{excel_path}")
    
    orders_df = pd.read_excel(excel_path, sheet_name="訂單")

    # Read rich text colors for column T
    wb = openpyxl.load_workbook(excel_path, data_only=True, rich_text=True)
    sheet = wb['訂單']
    
    completed_col = []
    uncompleted_col = []
    
    for r in range(2, len(orders_df) + 2):
        cell = sheet.cell(row=r, column=20)
        val = cell.value
        completed_parts = []
        uncompleted_parts = []
        
        if hasattr(val, '__iter__') and not isinstance(val, str):
            for block in val:
                if isinstance(block, str):
                    uncompleted_parts.append(block)
                else:
                    text = getattr(block, 'text', str(block))
                    color = block.font.color.rgb if hasattr(block, 'font') and block.font and block.font.color else 'None'
                    if '00B050' in str(color) or '008000' in str(color):
                        completed_parts.append(text)
                    else:
                        uncompleted_parts.append(text)
        else:
            uncompleted_parts.append(str(val) if pd.notna(val) else '')
            
        completed_col.append(''.join(completed_parts).strip())
        uncompleted_col.append(''.join(uncompleted_parts).strip())

    orders_df['已完成製程'] = completed_col
    orders_df['未完成製程'] = uncompleted_col

    # Clean numeric columns
    orders_df['order_qty_num'] = pd.to_numeric(orders_df['訂單數量'], errors='coerce').fillna(0)
    orders_df['unshipped_num'] = pd.to_numeric(orders_df['未交'], errors='coerce').fillna(0)
    orders_df['produced_num'] = pd.to_numeric(orders_df['生產數量'], errors='coerce').fillna(0)

    # Filter 1: Only keep rows where unshipped > 0
    active_orders = orders_df[orders_df['unshipped_num'] > 0].copy()

    # Filter 2: Exclude non-production items
    exclude_keywords = ['暫停', '樣品', '待', '未排', '停']
    
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
            return "PENDING", ""
        try:
            date_str = str(date_val).split('\n')[0].strip()
            d = pd.to_datetime(date_str, errors='coerce')
            if pd.isna(d):
                return "PENDING", str(date_val)
            
            delta_days = (d - ref_date).days
            if delta_days < 0:
                return "RED-OVERDUE", d.strftime('%Y-%m-%d')
            elif delta_days <= 7:
                return "YELLOW-7DAYS", d.strftime('%Y-%m-%d')
            else:
                return "GREEN-NORMAL", d.strftime('%Y-%m-%d')
        except:
            return "PENDING", str(date_val)

    risk_results = active_orders['交貨日'].apply(parse_and_calculate_risk)
    active_orders['risk_level'] = [r[0] for r in risk_results]
    active_orders['clean_date'] = [r[1] for r in risk_results]

    # Map risk level to pure emoji lights for the first column
    risk_symbol_map = {
        "RED-OVERDUE": "🔴",
        "YELLOW-7DAYS": "🟡",
        "GREEN-NORMAL": "🟢",
        "PENDING": "⚪"
    }
    active_orders['risk_symbol'] = active_orders['risk_level'].map(risk_symbol_map)

    # Reorder columns: 交期風險(燈號)、預交日、客戶、產品圖號、訂單號碼、未交數、生產數、未完成製程
    summary_df = active_orders[['risk_symbol', 'clean_date', 'masked_customer', '圖號', '訂單號碼', 'unshipped_num', 'produced_num', '未完成製程', '已完成製程', '備註', 'risk_level']].copy()
    summary_df.columns = ['交期風險', '預定交貨日', '客戶', '產品圖號', '訂單編號', '未交數量', '已生產數量', '未完成製程', '已完成製程', '備註', 'risk_code']

    risk_order = {"RED-OVERDUE": 1, "YELLOW-7DAYS": 2, "GREEN-NORMAL": 3, "PENDING": 4}
    summary_df['risk_sort'] = summary_df['risk_code'].map(risk_order)
    summary_df = summary_df.sort_values(by=['risk_sort', '未交數量'], ascending=[True, False]).drop(columns=['risk_sort', 'risk_code'])

    # Save CSV report
    output_csv_path = os.path.join(os.path.dirname(__file__), "plant_summary_dashboard.csv")
    summary_df.to_csv(output_csv_path, index=False, encoding='utf-8-sig')

    # Save styled HTML report
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
            .controls {{ margin: 15px 0; display: flex; justify-content: space-between; align-items: center; }}
            .search-box {{ padding: 8px 12px; width: 300px; border: 1px solid #ced4da; border-radius: 4px; font-size: 14px; }}
            table {{ border-collapse: collapse; width: 100%; background: white; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }}
            th, td {{ border: 1px solid #dee2e6; padding: 10px 12px; text-align: left; font-size: 13px; }}
            th {{ background-color: #343a40; color: white; cursor: pointer; user-select: none; position: relative; }}
            th:hover {{ background-color: #495057; }}
            th::after {{ content: " ↕"; font-size: 11px; opacity: 0.6; }}
            tr.red {{ background-color: #f8d7da; color: #721c24; font-weight: bold; }}
            tr.yellow {{ background-color: #fff3cd; color: #856404; }}
            tr.green {{ background-color: #d4edda; color: #155724; }}
            tr.pending {{ background-color: #e2e3e5; color: #383d41; }}
            .light-cell {{ text-align: center; font-size: 16px; }}
            .completed-text {{ color: #155724; font-weight: 500; }}
            .uncompleted-text {{ color: #721c24; font-weight: bold; }}
        </style>
        <script>
            function filterTable() {{
                var input = document.getElementById("searchInput");
                var filter = input.value.toLowerCase();
                var table = document.getElementById("dashboardTable");
                var tr = table.getElementsByTagName("tr");

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

            function sortTable(n) {{
                var table, rows, switching, i, x, y, shouldSwitch, dir, switchcount = 0;
                table = document.getElementById("dashboardTable");
                switching = true;
                dir = "asc";
                while (switching) {{
                    switching = false;
                    rows = table.rows;
                    for (i = 1; i < (rows.length - 1); i++) {{
                        shouldSwitch = false;
                        x = rows[i].getElementsByTagName("TD")[n];
                        y = rows[i + 1].getElementsByTagName("TD")[n];
                        var xVal = x.textContent || x.innerText;
                        var yVal = y.textContent || y.innerText;
                        if (dir == "asc") {{
                            if (xVal > yVal) {{ shouldSwitch = true; break; }}
                        }} else if (dir == "desc") {{
                            if (xVal < yVal) {{ shouldSwitch = true; break; }}
                        }}
                    }}
                    if (shouldSwitch) {{
                        rows[i].parentNode.insertBefore(rows[i + 1], rows[i]);
                        switching = true;
                        switchcount++;
                    }} else {{
                        switchcount++;
                        if (switchcount == 1 && dir == "asc") {{
                            dir = "desc";
                            switching = true;
                        }}
                    }}
                }}
            }}
        </script>
    </head>
    <body>
        <h2>廠務生產現場未交訂單與製程進度看板 (發表演示版)</h2>
        <div class="controls">
            <p>評估基準日（今日）：<strong>{ref_date.strftime('%Y-%m-%d')}</strong> | 未交訂單總計：<strong>{len(summary_df)}</strong> 筆（客戶欄位已遮罩）</p>
            <input type="text" id="searchInput" class="search-box" onkeyup="filterTable()" placeholder="🔍 快速搜尋任何關鍵字（客戶、圖號、狀態...）">
        </div>
        <table id="dashboardTable">
            <thead>
                <tr>
                    <th onclick="sortTable(0)" style="width: 70px; text-align: center;">交期風險</th>
                    <th onclick="sortTable(1)">預交日</th>
                    <th onclick="sortTable(2)">客戶</th>
                    <th onclick="sortTable(3)">產品圖號</th>
                    <th onclick="sortTable(4)">訂單號碼</th>
                    <th onclick="sortTable(5)">未交數</th>
                    <th onclick="sortTable(6)">生產數</th>
                    <th onclick="sortTable(7)">未完成製程</th>
                    <th onclick="sortTable(8)">已完成製程</th>
                    <th onclick="sortTable(9)">備註</th>
                </tr>
            </thead>
            <tbody>
    """

    for _, row in summary_df.iterrows():
        symbol = row['交期風險']
        tr_class = "pending"
        if symbol == "🔴":
            tr_class = "red"
        elif symbol == "🟡":
            tr_class = "yellow"
        elif symbol == "🟢":
            tr_class = "green"

        html_content += f"""
            <tr class="{tr_class}">
                <td class="light-cell">{symbol}</td>
                <td>{row['預定交貨日']}</td>
                <td>{row['客戶']}</td>
                <td>{row['產品圖號']}</td>
                <td>{row['訂單編號'] if pd.notna(row['訂單編號']) else ''}</td>
                <td>{row['未交數量']:,.0f}</td>
                <td>{row['已生產數量']:,.0f}</td>
                <td class="uncompleted-text">{row['未完成製程'] if pd.notna(row['未完成製程']) else ''}</td>
                <td class="completed-text">{row['已完成製程'] if pd.notna(row['已完成製程']) else ''}</td>
                <td>{row['備註'] if pd.notna(row['備註']) else ''}</td>
            </tr>
        """

    html_content += """
            </tbody>
        </table>
    </body>
    </html>
    """

    with open(output_html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"\nSuccessfully generated HTML colored dashboard: {output_html_path}")
    print(f"Successfully updated CSV report: {output_csv_path}")

if __name__ == "__main__":
    main()
