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

    print(f"正在讀取檔案與解析 Rich Text 色彩分拆製程：{excel_path}")
    
    # 2. Read with pandas first to get exact row count
    orders_df = pd.read_excel(excel_path, sheet_name="訂單")

    # 1. Read with openpyxl (rich_text=True) to extract green (completed) vs red (uncompleted) parts of Column T (col 20)
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
            # Fallback if plain string
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

    # Filter 2: Exclude non-production items (e.g. contain '暫停', '樣品', '待', '未排', '停' in 交貨日)
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

    summary_df = active_orders[['訂單號碼', 'masked_customer', '圖號', 'order_qty_num', 'unshipped_num', 'produced_num', 'clean_date', 'risk_level', '已完成製程', '未完成製程', '備註']].copy()
    summary_df.columns = ['訂單編號', '客戶', '產品圖號', '訂單數量', '未交數量', '已生產數量', '預定交貨日', '交期風險燈號', '已完成製程', '未完成製程', '備註']

    risk_order = {"RED-OVERDUE": 1, "YELLOW-7DAYS": 2, "GREEN-NORMAL": 3, "PENDING": 4}
    summary_df['risk_sort'] = summary_df['交期風險燈號'].map(risk_order)
    summary_df = summary_df.sort_values(by=['risk_sort', '未交數量'], ascending=[True, False]).drop(columns=['risk_sort'])

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
            h2 {{ color: #333; }}
            table {{ border-collapse: collapse; width: 100%; background: white; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }}
            th, td {{ border: 1px solid #dee2e6; padding: 10px 12px; text-align: left; font-size: 13px; }}
            th {{ background-color: #343a40; color: white; }}
            tr.red {{ background-color: #f8d7da; color: #721c24; font-weight: bold; }}
            tr.yellow {{ background-color: #fff3cd; color: #856404; }}
            tr.green {{ background-color: #d4edda; color: #155724; }}
            tr.pending {{ background-color: #e2e3e5; color: #383d41; }}
            .badge {{ padding: 4px 8px; border-radius: 4px; font-size: 12px; }}
            .badge-red {{ background: #dc3545; color: white; }}
            .badge-yellow {{ background: #ffc107; color: black; }}
            .badge-green {{ background: #28a745; color: white; }}
            .badge-gray {{ background: #6c757d; color: white; }}
            .completed-text {{ color: #155724; font-weight: 500; }}
            .uncompleted-text {{ color: #721c24; font-weight: bold; }}
        </style>
    </head>
    <body>
        <h2>廠務生產現場未交訂單與製程進度看板 (發表演示版 - 客戶已遮罩)</h2>
        <p>評估基準日（今日）：<strong>{ref_date.strftime('%Y-%m-%d')}</strong> | 現場執行未交訂單總計：<strong>{len(summary_df)}</strong> 筆</p>
        <table>
            <tr>
                <th>訂單編號</th>
                <th>客戶</th>
                <th>產品圖號</th>
                <th>訂單數量</th>
                <th>未交數量</th>
                <th>已生產數量</th>
                <th>預定交貨日</th>
                <th>交期風險燈號</th>
                <th>已完成製程 (綠字)</th>
                <th>未完成製程 (紅字)</th>
                <th>備註</th>
            </tr>
    """

    for _, row in summary_df.iterrows():
        risk = row['交期風險燈號']
        tr_class = "pending"
        badge_class = "badge-gray"
        if risk == "RED-OVERDUE":
            tr_class = "red"
            badge_class = "badge-red"
            risk_text = "🔴 紅燈 (已逾期)"
        elif risk == "YELLOW-7DAYS":
            tr_class = "yellow"
            badge_class = "badge-yellow"
            risk_text = "🟡 黃燈 (7天內到期)"
        elif risk == "GREEN-NORMAL":
            tr_class = "green"
            badge_class = "badge-green"
            risk_text = "🟢 綠燈 (正常)"
        else:
            risk_text = "⚪ 待確認"

        html_content += f"""
            <tr class="{tr_class}">
                <td>{row['訂單編號'] if pd.notna(row['訂單編號']) else ''}</td>
                <td>{row['客戶'] if pd.notna(row['客戶']) else ''}</td>
                <td>{row['產品圖號'] if pd.notna(row['產品圖號']) else ''}</td>
                <td>{row['訂單數量']:,.0f}</td>
                <td>{row['未交數量']:,.0f}</td>
                <td>{row['已生產數量']:,.0f}</td>
                <td>{row['預定交貨日']}</td>
                <td><span class="badge {badge_class}">{risk_text}</span></td>
                <td class="completed-text">{row['已完成製程'] if pd.notna(row['已完成製程']) else ''}</td>
                <td class="uncompleted-text">{row['未完成製程'] if pd.notna(row['未完成製程']) else ''}</td>
                <td>{row['備註'] if pd.notna(row['備註']) else ''}</td>
            </tr>
        """

    html_content += """
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
