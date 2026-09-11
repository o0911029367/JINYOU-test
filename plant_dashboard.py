import os
import pandas as pd

def main():
    base_dir = r"C:\Users\JIN YOU\Desktop\AI賦能智造製造業數據驅動與人機協作實戰\GITHUB TEST\資料來源"
    db_excel = os.path.join(base_dir, "生管", "資料庫遷移結果", "生管資料庫遷移結果.xlsx")
    
    if not os.path.exists(db_excel):
        print(f"File not found: {db_excel}")
        return

    print(f"正在讀取生管資料庫：{db_excel}")
    orders_df = pd.read_excel(db_excel, sheet_name="orders")
    
    # Clean numeric columns
    orders_df['order_qty_num'] = pd.to_numeric(orders_df['order_qty'], errors='coerce').fillna(0)
    orders_df['unshipped_qty_num'] = pd.to_numeric(orders_df['unshipped_qty'], errors='coerce').fillna(0)
    orders_df['produced_qty_num'] = pd.to_numeric(orders_df['produced_qty'], errors='coerce').fillna(0)

    # Select and rename key columns for plant management
    summary_df = orders_df[['order_no', 'customer_id', 'product_id', 'order_qty_num', 'unshipped_qty_num', 'produced_qty_num', 'promised_date', 'status']].copy()
    summary_df.columns = ['訂單編號', '客戶', '產品圖號', '訂單數量', '未交數量', '已生產數量', '預定交貨日', '訂單狀態']

    # Sort by unshipped quantity descending
    summary_df = summary_df.sort_values(by='未交數量', ascending=False)

    print(f"\n--- 廠務每日訂單與未交進度彙整表 (共 {len(summary_df)} 筆) ---")
    print(summary_df.head(10).to_string(index=False))

    # Save summary report
    output_report_path = os.path.join(os.path.dirname(__file__), "plant_summary_dashboard.csv")
    summary_df.to_csv(output_report_path, index=False, encoding='utf-8-sig')
    print(f"\n已成功更新廠務彙整報表至：{output_report_path}")

if __name__ == "__main__":
    main()
