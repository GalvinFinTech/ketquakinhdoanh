import os
import json
import re
import subprocess
from datetime import datetime
import pandas as pd

def extract_and_generate():
    """
    Kịch bản tự động hóa 100% cho Yuanta Việt Nam:
    1. Đọc dữ liệu từ Dashboard_KQKD.xlsx (Sheet: Dashboard)
    2. Lọc danh sách DN đã công bố KQKD có TĂNG TRƯỜNG LNST Q2 DƯƠNG (TT_LNST_Q_YoY > 0)
    3. Tiêu chí phân tầng 2-Tier:
       - Tầng 1: Lấy các doanh nghiệp Big Cap (Vốn hóa >= 10.000 tỷ) có TT_LNST_Q_YoY > 0, xếp từ cao xuống thấp.
       - Tầng 2: Nếu chưa đủ 30 mã, bù tiếp bằng Mid Cap (1.000 tỷ <= Vốn hóa < 10.000 tỷ) có TT_LNST_Q_YoY > 0 cho đủ đúng 30 mã.
    4. Xuất Báo cáo Tóm tắt (JSON, TXT) đầy đủ 11 chỉ tiêu tài chính cho PowerPoint.
    5. Cập nhật dữ liệu vào index.html & Tự động Push Git (Git Sync).
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(script_dir, "Dashboard_KQKD.xlsx")
    html_path = os.path.join(script_dir, "index.html")
    json_report_path = os.path.join(script_dir, "summary_report.json")
    txt_report_path = os.path.join(script_dir, "summary_report.txt")

    if not os.path.exists(excel_path):
        print(f"❌ LỖI: Không tìm thấy file 'Dashboard_KQKD.xlsx' tại {script_dir}")
        return

    print("📊 Đang đọc dữ liệu KQKD thực tế từ Excel...")

    try:
        df_raw = pd.read_excel(excel_path, sheet_name="Dashboard")
        header_idx = None
        for idx, row in df_raw.iterrows():
            row_str = " ".join([str(val) for val in row.values if pd.notna(val)])
            if "Mã CP" in row_str or "MaCP" in row_str or "Ticker" in row_str:
                header_idx = idx
                break
        
        if header_idx is not None and header_idx > 0:
            df = pd.read_excel(excel_path, sheet_name="Dashboard", skiprows=header_idx + 1)
        else:
            df = df_raw

        df.columns = [str(c).strip() for c in df.columns]

        column_mapping = {
            'MaCP': ['Mã CP', 'MaCP', 'Ticker'],
            'TenCongTy': ['Tên công ty', 'Company Name'],
            'San': ['Sàn', 'Exchange'],
            'VonHoa': ['Vốn hóa thị trường (tỷ đồng)', 'Vốn hóa', 'Market Cap'],
            'NganhL1': ['Ngành L1'],
            'NganhL2': ['Ngành L2'],
            'NganhL4': ['Ngành L4'],
            'DTT_Q2': ['DTT Q2/26 (tỷ đồng)', 'DTT Q2/26', 'DTT Q2'],
            'DTT_6M': ['DTT 6T/26 (tỷ đồng)', 'DTT 6T/26', 'DTT 6M'],
            'TT_DTT_Q_YoY': ['TT DTT Q YoY', 'DTT Q YoY'],
            'TT_DTT_6M_YoY': ['TT DTT 6M YoY', 'DTT 6M YoY'],
            'LNST_Q2': ['LNST Q2/26 (tỷ đồng)', 'LNST Q2/26', 'LNST Q2'],
            'LNST_6M': ['LNST 6T/26 (tỷ đồng)', 'LNST 6T/26', 'LNST 6M'],
            'TT_LNST_Q_YoY': ['TT LNST Q YoY', 'LNST Q YoY'],
            'TT_LNST_6M_YoY': ['TT LNST 6M YoY', 'LNST 6M YoY'],
            'TT_DTT_QoQ': ['TT DTT QoQ', 'DTT QoQ'],
            'TT_LNST_QoQ': ['TT LNST QoQ', 'LNST QoQ']
        }

        matched_cols = {}
        for target, aliases in column_mapping.items():
            for col in df.columns:
                if any(alias.lower() == col.lower() or alias.lower() in col.lower() for alias in aliases):
                    matched_cols[col] = target
                    break

        df_renamed = df.rename(columns=matched_cols)
        df_filtered = df_renamed.dropna(subset=['MaCP']).copy()
        df_filtered['MaCP'] = df_filtered['MaCP'].astype(str).str.strip()
        df_filtered = df_filtered[df_filtered['MaCP'] != 'nan']

        clean_records = []
        for _, row in df_filtered.iterrows():
            rec = {}
            for target_key in column_mapping.keys():
                val = row.get(target_key, None)
                if pd.isna(val) or val == 'nan' or val is None:
                    rec[target_key] = None
                elif isinstance(val, (int, float)):
                    rec[target_key] = round(float(val), 2)
                else:
                    rec[target_key] = str(val).strip()
            clean_records.append(rec)

        total_companies = len(clean_records)
        total_market_cap = sum(r['VonHoa'] for r in clean_records if r.get('VonHoa'))
        
        # Published records (DTT_Q2 or LNST_Q2 is not null)
        published_records = [r for r in clean_records if r.get('DTT_Q2') is not None or r.get('LNST_Q2') is not None]
        published_count = len(published_records)
        published_market_cap = sum(r['VonHoa'] for r in published_records if r.get('VonHoa'))
        
        market_cap_percentage = (published_market_cap / total_market_cap * 100) if total_market_cap > 0 else 0
        
        profitable_count = len([r for r in published_records if r.get('LNST_Q2') is not None and r['LNST_Q2'] > 0])
        loss_count = len([r for r in published_records if r.get('LNST_Q2') is not None and r['LNST_Q2'] < 0])
        
        dn_co_lnst_count = profitable_count + loss_count
        profit_ratio = (profitable_count / dn_co_lnst_count * 100) if dn_co_lnst_count > 0 else 0

        # Tier 1: Big Cap (VonHoa >= 10,000 tỷ) WITH STRICT POSITIVE YoY GROWTH (> 0%)
        bigcap_positive = [
            r for r in published_records 
            if r.get('VonHoa') is not None and r['VonHoa'] >= 10000 
            and r.get('TT_LNST_Q_YoY') is not None and r['TT_LNST_Q_YoY'] > 0
        ]
        bigcap_sorted = sorted(bigcap_positive, key=lambda x: x['TT_LNST_Q_YoY'], reverse=True)

        top30_selected = list(bigcap_sorted)

        # Tier 2: Mid Cap (1,000 tỷ <= VonHoa < 10,000 tỷ) WITH STRICT POSITIVE YoY GROWTH (> 0%)
        if len(top30_selected) < 30:
            needed = 30 - len(top30_selected)
            midcap_positive = [
                r for r in published_records 
                if r.get('VonHoa') is not None and 1000 <= r['VonHoa'] < 10000 
                and r.get('TT_LNST_Q_YoY') is not None and r['TT_LNST_Q_YoY'] > 0
            ]
            midcap_sorted = sorted(midcap_positive, key=lambda x: x['TT_LNST_Q_YoY'], reverse=True)
            top30_selected.extend(midcap_sorted[:needed])

        report_data = {
            "executive_summary": {
                "total_companies": total_companies,
                "published_count": published_count,
                "total_market_cap_bil_vnd": round(total_market_cap, 2),
                "published_market_cap_bil_vnd": round(published_market_cap, 2),
                "published_market_cap_percentage": round(market_cap_percentage, 2),
                "profitable_companies": profitable_count,
                "loss_companies": loss_count,
                "profit_ratio_percentage": round(profit_ratio, 2)
            },
            "top_30_large_mid_growth": top30_selected
        }

        with open(json_report_path, 'w', encoding='utf-8') as f:
            json.dump(report_data, f, ensure_ascii=False, indent=4)
            
        with open(txt_report_path, 'w', encoding='utf-8') as f:
            f.write("BÁO CÁO NHANH KẾT QUẢ KINH DOANH Q2\n")
            f.write("="*50 + "\n")
            f.write(f"- Thống kê Q2: Có {published_count}/{total_companies} DN đã công bố.\n")
            f.write(f"- Nhóm này chiếm {market_cap_percentage:.2f}% tổng vốn hóa toàn thị trường.\n")
            f.write(f"- Tỷ lệ có lãi: {profit_ratio:.2f}% ({profitable_count} DN Lãi / {loss_count} DN Lỗ).\n\n")
            f.write(f"TOP {len(top30_selected)} DOANH NGHIỆP VỐN HÓA LỚN & VỪA TĂNG TRƯỜNG LNST Q2 DƯƠNG MẠNH NHẤT:\n")
            f.write("-" * 80 + "\n")
            for r in top30_selected:
                f.write(f"{r['MaCP']:<8} | Vốn hóa: {r['VonHoa']:>10,.2f} tỷ | LNST Q2: {r['LNST_Q2']:>10,.2f} tỷ | YoY: +{r['TT_LNST_Q_YoY']:>6.2f}%\n")

        print(f"✅ Đã trích xuất thành công Top {len(top30_selected)} DN tăng trưởng LNST Q2 dương: {txt_report_path}")

        if os.path.exists(html_path):
            with open(html_path, 'r', encoding='utf-8') as f:
                html_content = f.read()

            pattern = r"let\s+rawStockData\s*=\s*\[.*?\];"
            replacement = f"let rawStockData = {json.dumps(clean_records, ensure_ascii=False)};"
            
            updated_html = re.sub(pattern, replacement, html_content, flags=re.DOTALL)
            
            with open(html_path, 'w', encoding='utf-8') as f:
                f.write(updated_html)
                
            print("🎉 Cập nhật thành công dữ liệu thực tế vào Web Dashboard chính (index.html)!")

        sync_git_repository(script_dir)

    except Exception as e:
        print(f"❌ Xảy ra lỗi trong quá trình xử lý: {e}")
        import traceback
        traceback.print_exc()

def sync_git_repository(repo_dir):
    """
    Tự động hóa đồng bộ Git repository sau khi cập nhật dữ liệu:
    1. git add .
    2. git commit -m "Auto-update KQKD data & summary reports [Timestamp]"
    3. git push origin main/master
    """
    print("🔄 Đang kiểm tra và đồng bộ Git repository...")
    try:
        # Check if git repository exists
        git_check = subprocess.run(["git", "rev-parse", "--is-inside-work-tree"], cwd=repo_dir, capture_output=True, text=True)
        if git_check.returncode != 0:
            print("ℹ️ Thư mục không thuộc Git repository, bỏ qua bước Git Sync.")
            return

        # Stage updated files
        subprocess.run(["git", "add", "."], cwd=repo_dir, check=True)

        # Create commit message with current timestamp
        commit_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        commit_msg = f"Auto-update KQKD data & summary reports ({commit_time})"
        
        # Commit changes
        commit_res = subprocess.run(["git", "commit", "-m", commit_msg], cwd=repo_dir, capture_output=True, text=True)
        if "nothing to commit" in commit_res.stdout.lower() or "nothing to commit" in commit_res.stderr.lower():
            print("ℹ️ Dữ liệu Git đã đồng bộ mới nhất (Không có thay đổi mới để commit).")
            return

        # Push to remote repository
        push_res = subprocess.run(["git", "push"], cwd=repo_dir, capture_output=True, text=True)
        if push_res.returncode == 0:
            print("🚀 Đã push thành công dữ liệu mới lên GitHub repository!")
        else:
            print(f"⚠️ Git push cảnh báo: {push_res.stderr.strip()}")

    except Exception as git_err:
        print(f"⚠️ Không thể tự động Sync Git: {git_err}")

if __name__ == "__main__":
    print("=" * 70)
    print("🚀 HỆ THỐNG TRÍCH XUẤT KQKD & ĐỒNG BỘ GIT - YUANTA VIỆT NAM")
    print("=" * 70)
    extract_and_generate()