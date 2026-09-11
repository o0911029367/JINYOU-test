import os
import pandas as pd
from datetime import datetime

def main():
    base_dir = r"C:\Users\JIN YOU\Desktop\AI賦能智造製造業數據驅動與人機協作實戰\GITHUB TEST\資料來源"
    db_excel = os.path.join(base_dir, "生管", "資料庫遷移結果", "生管資料庫遷移結果.xlsx")
    
    if not os.path.exists(db_excel):
        print(f"File not found: {db_excel}")
        return

    orders_df = pd.read_excel(db_excel, sheet_name="orders")
    
    # Clean numeric columns
    orders_df['order_qty_num'] = pd.to_numeric(orders_df['order_qty'], errors='coerce').fillna(0)
    orders_df['unshipped_qty_num'] = pd.to_numeric(orders_df['unshipped_qty'], errors='coerce').fillna(0)
    orders_df['produced_qty_num'] = pd.to_numeric(orders_df['produced_qty'], errors='coerce').fillna(0)

    # Filter: Exclude shipped orders (keep only active un-shipped orders where status != '已出貨')
    active_orders = orders_df[orders_df['status'] != '已出貨'].copy()

    ref_date = pd.to_datetime('2026-06-15')
    
    def calculate_risk(date_val):
        if pd.isna(date_val):
            return "PENDING"
        try:
            d = pd.to_datetime(date_val)
            delta_days = (d - ref_date).days
            if delta_days < 0:
                return "RED-OVERDUE"
            elif delta_days <= 7:
                return "YELLOW-7DAYS"
            else:
                return "GREEN-NORMAL"
        except:
            return "PENDING"

    active_orders['risk_level'] = active_orders['promised_date'].apply(calculate_risk)

    summary_df = active_orders[['order_no', 'customer_id', 'product_id', 'order_qty_num', 'unshipped_qty_num', 'produced_qty_num', 'promised_date', 'risk_level', 'status']].copy()
    summary_df.columns = ['訂單編號', '客戶', '產品圖號', '訂單數量', '未交數量', '已生產數量', '預定交貨日', '交期風險燈號', '訂單狀態']

    risk_order = {"RED-OVERDUE": 1, "YELLOW-7DAYS": 2, "GREEN-NORMAL": 3, "PENDING": 4}
    summary_df['risk_sort'] = summary_df['交期風險燈號'].map(risk_order)
    summary_df = summary_df.sort_values(by=['risk_sort', '未交數量'], ascending=[True, False]).drop(columns=['risk_sort'])

    # Save CSV report
    output_csv_path = os.path.join(os.path.dirname(__file__), "plant_summary_dashboard.csv")
    summary_df.to_csv(output_csv_path, index=False, encoding='utf-8-sig')

    # Save styled HTML report with color highlighting
    output_html_path = os.path.join(os.path.dirname(__file__), "plant_summary_dashboard.html")
    
    html_content = """
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>廠務未交訂單交期風險看板</title>
        <style>
            body { font-family: "Microsoft JhengHei", Arial, sans-serif; margin: 20px; background-color: #f8f9fa; }
            h2 { color: #333; }
            table { border-collapse: collapse; width: 100%; background: white; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
            th, td { border: 1px solid #dee2e6; padding: 10px 12px; text-align: left; }
            th { background-color: #343a40; color: white; }
            tr.red { background-color: #f8d7da; color: #721c24; font-weight: bold; }
            tr.yellow { background-color: #fff3cd; color: #856404; }
            tr.green { background-color: #d4edda; color: #155724; }
            tr.pending { background-color: #e2e3e5; color: #383d41; }
            .badge { padding: 4px 8px; border-radius: 4px; font-size: 12px; }
            .badge-red { background: #dc3545; color: white; }
            .badge-yellow { background: #ffc107; color: black; }
            .badge-green { background: #28a745; color: white; }
            .badge-gray { background: #6c757d; color: white; }
        </style>
    </head>
    <body>
        <h2>廠務未交訂單交期風險即時看板</h2>
        <p>基準日：2026-06-15 | 進行中未交訂單總計：<strong>""" + str(len(summary_df)) + """</strong> 筆</p>
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
                <th>訂單狀態</th>
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
                <td>{row['訂單編號']}</td>
                <td>{row['客戶']}</td>
                <td>{row['產品圖號']}</td>
                <td>{row['訂單數量']:,.0f}</td>
                <td>{row['未交數量']:,.0f}</td>
                <td>{row['已生產數量']:,.0f}</td>
                <td>{str(row['預定交貨日']).split(' ')[0]}</td>
                <td><span class="badge {badge_class}">{risk_text}</span></td>
                <td>{row['訂單狀態']}</td>
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
