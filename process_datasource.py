import os
import pandas as pd

def main():
    base_dir = r"C:\Users\JIN YOU\Desktop\AI賦能智造製造業數據驅動與人機協作實戰\GITHUB TEST\資料來源"
    db_excel = os.path.join(base_dir, "生管", "資料庫遷移結果", "生管資料庫遷移結果.xlsx")
    
    if not os.path.exists(db_excel):
        print(f"File not found: {db_excel}")
        return

    orders_df = pd.read_excel(db_excel, sheet_name="orders")
    work_orders_df = pd.read_excel(db_excel, sheet_name="work_orders")
    
    orders_df['order_qty_num'] = pd.to_numeric(orders_df['order_qty'], errors='coerce')
    orders_df['unshipped_qty_num'] = pd.to_numeric(orders_df['unshipped_qty'], errors='coerce')
    
    total_order_qty = orders_df['order_qty_num'].sum()
    total_unshipped = orders_df['unshipped_qty_num'].sum()
    
    print(f"=== 廠務生管資料庫分析報告 ===")
    print(f"- 總訂單筆數: {len(orders_df)}")
    print(f"- 總生產工單筆數: {len(work_orders_df)}")
    print(f"- 總訂單數量: {total_order_qty:,.0f}")
    print(f"- 總未交數量 (Unshipped): {total_unshipped:,.0f}")
    
    # Show top unshipped orders
    print(f"\n--- 前 5 筆未交數量最多的訂單 ---")
    top_unshipped = orders_df.sort_values(by='unshipped_qty_num', ascending=False).head(5)
    for idx, row in top_unshipped.iterrows():
        print(f"單號: {row.get('order_no')} | 客戶: {row.get('customer_id')} | 未交數: {row.get('unshipped_qty_num')} | 交貨日: {row.get('promised_date')}")

if __name__ == "__main__":
    main()
