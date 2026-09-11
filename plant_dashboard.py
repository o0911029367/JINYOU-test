import os
import csv

def main():
    base_dir = r"C:\Users\JIN YOU\Desktop\AI賦能智造製造業數據驅動與人機協作實戰"
    
    # Find student materials directory
    edu_dir = None
    for d in os.listdir(base_dir):
        p = os.path.join(base_dir, d)
        if os.path.isdir(p) and any("生產日報" in f for f in os.listdir(p) if os.path.isfile(os.path.join(p, f))):
            edu_dir = p
            break
            
    if not edu_dir:
        # Fallback to checking subdirectories
        for d in os.listdir(base_dir):
            p = os.path.join(base_dir, d)
            if os.path.isdir(p):
                for sub in os.listdir(p):
                    sub_p = os.path.join(p, sub)
                    if os.path.isdir(sub_p) and any("生產日報" in f for f in os.listdir(sub_p)):
                        edu_dir = sub_p
                        break

    print(f"Materials directory located: {edu_dir}")

    # 1. Parse Customer Files (C01 - C20)
    customer_dir = None
    for root, dirs, files in os.walk(base_dir):
        if "C01.txt" in files:
            customer_dir = root
            break

    customer_records = []
    if customer_dir:
        for fname in sorted(os.listdir(customer_dir)):
            if fname.startswith("C") and fname.endswith(".txt"):
                fpath = os.path.join(customer_dir, fname)
                try:
                    with open(fpath, 'r', encoding='utf-8-sig', errors='ignore') as f:
                        lines = [line.strip() for line in f if line.strip()]
                except:
                    with open(fpath, 'r', encoding='cp950', errors='ignore') as f:
                        lines = [line.strip() for line in f if line.strip()]
                
                sender = "未知"
                subject = "未知"
                for line in lines[:3]:
                    if "寄件者" in line or "From" in line:
                        sender = line.split("：")[-1].strip() if "：" in line else line.split(":")[-1].strip()
                    elif "主旨" in line or "Subject" in line:
                        subject = line.split("：")[-1].strip() if "：" in line else line.split(":")[-1].strip()
                
                customer_records.append({
                    "id": fname.replace(".txt", ""),
                    "sender": sender,
                    "subject": subject
                })

    print(f"\n[1] 已成功解析客戶檔案共 {len(customer_records)} 筆。")
    print(f"{'編號':<6} | {'客戶/寄件者':<25} | {'主旨/事項':<30}")
    print("-" * 68)
    for r in customer_records[:10]:
        print(f"{r['id']:<6} | {r['sender']:<25} | {r['subject']:<30}")

    # 2. Parse Production Daily Report
    prod_records = []
    prod_csv_path = None
    for root, dirs, files in os.walk(base_dir):
        for f in files:
            if "生產日報" in f and f.endswith(".csv"):
                prod_csv_path = os.path.join(root, f)
                break
        if prod_csv_path:
            break

    if prod_csv_path:
        try:
            with open(prod_csv_path, 'r', encoding='utf-8-sig', errors='ignore') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    prod_records.append(row)
        except Exception as e:
            print(f"Error reading production csv: {e}")

    print(f"\n[2] 已成功解析生產日報共 {len(prod_records)} 筆記錄。")
    
    # 3. Generate Summary Report
    output_report_path = os.path.join(os.path.dirname(__file__), "plant_summary_dashboard.csv")
    with open(output_report_path, 'w', encoding='utf-8-sig', newline='') as csvfile:
        fieldnames = ['id', 'sender', 'subject']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for r in customer_records:
            writer.writerow(r)

    print(f"\n廠務數據彙整報表已成功產出：{output_report_path}")

if __name__ == "__main__":
    main()
