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

    # Filter: Only keep orders with unshipped quantity > 0
    active_orders = orders_df[orders_df['unshipped_qty_num'] > 0].copy()

    ref_date = pd.to_datetime('2026-06-15')
    
    def calculate_risk(date_val):
        if pd.isna(date_val):
            return "[PENDING]"
        try:
            d = pd.to_datetime(date_val)
            delta_days = (d - ref_date).days
            if delta_days < 0:
                return "[RED-OVERDUE]"
            elif delta_days <= 7:
                return "[YELLOW-7DAYS]"
            else:
                return "[GREEN-NORMAL]"
        except:
            return "[PENDING]"

    active_orders['risk_level'] = active_orders['promised_date'].apply(calculate_risk)

    summary_df = active_orders[['order_no', 'customer_id', 'product_id', 'order_qty_num', 'unshipped_qty_num', 'produced_qty_num', 'promised_date', 'risk_level', 'status']].copy()
    summary_df.columns = ['訂單編號', '客戶', '產品圖號', '訂單數量', '未交數量', '已生產數量', '預定交貨日', '交期風險燈號', '訂單狀態']

    risk_order = {"[RED-OVERDUE]": 1, "[YELLOW-7DAYS]": 2, "[GREEN-NORMAL]": 3, "[PENDING]": 4}
    summary_df['risk_sort'] = summary_df['交期風險燈號'].map(risk_order)
    summary_df = summary_df.sort_values(by=['risk_sort', '未交數量'], ascending=[True, False]).drop(columns=['risk_sort'])

    print(f"\n--- Plant Unshipped Orders Risk Dashboard ({len(summary_df)} records) ---")
    print(summary_df[['訂單編號', '客戶', '未交數量', '預定交貨日', '交期風險燈號']].head(15).to_string(index=False))

    output_report_path = os.path.join(os.path.dirname(__file__), "plant_summary_dashboard.csv")
    summary_df.to_csv(output_report_path, index=False, encoding='utf-8-sig')
    print(f"\nSuccessfully updated risk summary report: {output_report_path}")

if __name__ == "__main__":
    main()
