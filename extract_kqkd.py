import os
import sys
import json
import re
import subprocess
from datetime import datetime
import pandas as pd

def find_excel_file(filename="Dashboard_KQKD.xlsx"):
    """
    Tìm kiếm file Excel linh hoạt ở các thư mục khả thi trong dự án.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    possible_paths = [
        os.path.join(script_dir, filename),
        filename,
        os.path.join(script_dir, "..", filename),
        os.path.join(script_dir, "..", "..", filename),
        os.path.join(os.getcwd(), filename),
    ]

    for path in possible_paths:
        if os.path.exists(path):
            return os.path.abspath(path)
    return None

def find_all_index_html_files(script_dir):
    """
    Tìm tất cả các vị trí file index.html duy nhất trong dự án (loại bỏ trùng lặp đường dẫn).
    """
    found_paths = []
    seen_normalized = set()

    def add_if_exists(p):
        if p and os.path.exists(p):
            # Chuẩn hóa đường dẫn tuyệt đối bao gồm chữ hoa ổ đĩa trên Windows
            real_p = os.path.realpath(p)
            if len(real_p) >= 2 and real_p[1] == ':':
                real_p = real_p[0].upper() + real_p[1:]
            norm_p = os.path.normpath(real_p)
            
            if norm_p not in seen_normalized:
                seen_normalized.add(norm_p)
                found_paths.append(norm_p)

    # Priority 1: Cùng cấp với script
    add_if_exists(os.path.join(script_dir, "index.html"))

    # Priority 2: Thư mục gốc Root dự án Git
    try:
        repo_root = subprocess.check_output(
            ["git", "rev-parse", "--show-toplevel"], 
            cwd=script_dir,
            stderr=subprocess.DEVNULL, 
            text=True
        ).strip()
        add_if_exists(os.path.join(repo_root, "index.html"))
    except Exception:
        pass

    # Fallback paths
    fallback_paths = [
        os.path.join(os.getcwd(), "index.html"),
        os.path.join(script_dir, "..", "index.html"),
        os.path.join(script_dir, "..", "..", "index.html")
    ]
    for path in fallback_paths:
        add_if_exists(path)

    return found_paths

def normalize_percentage_columns(df, percent_cols):
    """
    Chuẩn hóa tỷ lệ phần trăm từ Excel:
    Khi pandas đọc file Excel, các ô định dạng Percentage (%) được lưu dưới dạng thập phân tỉ lệ.
    Hàm này tự động nhân 100 toàn bộ cột phần trăm nếu phát hiện dữ liệu ở dạng thập phân.
    """
    df_normalized = df.copy()
    for col in percent_cols:
        if col not in df_normalized.columns:
            continue
        
        valid_series = pd.to_numeric(df_normalized[col], errors='coerce').dropna()
        if valid_series.empty:
            continue

        if valid_series.abs().median() < 50.0 or (valid_series.abs() <= 100.0).mean() > 0.8:
            df_normalized[col] = df_normalized[col].apply(
                lambda v: round(float(v) * 100.0, 2) if isinstance(v, (int, float)) and pd.notna(v) else v
            )
            
    return df_normalized

def extract_top_30_growth(file_path="Dashboard_KQKD.xlsx", sheet_name="Dashboard"):
    """
    Trích xuất danh sách Top 30 Doanh nghiệp Vốn hóa Lớn & Vừa (Vốn hóa >= 1,000 tỷ VNĐ)
    có TĂNG TRƯỜNG LNST Q2 DƯƠNG (> 0%) CAO NHẤT THỊ TRƯỜNG:
    - Lọc toàn bộ DN có Vốn hóa >= 1,000 tỷ & TT LNST Q2 YoY > 0%.
    - Sắp xếp trực tiếp toàn bộ nhóm này theo % Tăng trưởng LNST Q2 YoY giảm dần.
    - Lấy đúng 30 doanh nghiệp dẫn đầu.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_full_path = find_excel_file(file_path)

    if not excel_full_path:
        print(f"❌ LỖI: Không tìm thấy file dữ liệu Excel '{file_path}'.")
        print("💡 Vui lòng đảm bảo file 'Dashboard_KQKD.xlsx' nằm trong thư mục dự án của bạn.")
        return False

    print(f"📊 Đang đọc dữ liệu KQKD thực tế từ Excel: {excel_full_path}...")
    try:
        df_raw = pd.read_excel(excel_full_path, sheet_name=sheet_name)
        
        header_idx = None
        for idx, row in df_raw.iterrows():
            row_str = " ".join([str(val) for val in row.values if pd.notna(val)])
            if "Mã CP" in row_str or "MaCP" in row_str or "Ticker" in row_str:
                header_idx = idx
                break
        
        if header_idx is not None and header_idx > 0:
            df = pd.read_excel(excel_full_path, sheet_name=sheet_name, skiprows=header_idx + 1)
        else:
            df = df_raw

        df.columns = [str(c).strip() for c in df.columns]

        column_mapping = {
            'MaCP': ['Mã CP', 'MaCP', 'Ticker', 'MA_CP'],
            'TenCongTy': ['Tên công ty', 'TenCongTy', 'Company Name', 'Tên Doanh Nghiệp'],
            'San': ['Sàn', 'San', 'Exchange'],
            'VonHoa': ['Vốn hóa thị trường (tỷ đồng)', 'Vốn hóa (tỷ)', 'VonHoa', 'Market Cap'],
            'NganhL1': ['Ngành L1', 'NganhL1', 'Sector L1'],
            'NganhL2': ['Ngành L2', 'NganhL2', 'Sector L2'],
            'NganhL4': ['Ngành L4', 'NganhL4', 'Sector L4'],
            'DTT_Q2': ['DTT Q2/26 (tỷ đồng)', 'DTT Q2/26', 'DTT_Q2'],
            'DTT_6M': ['DTT 6T/26 (tỷ đồng)', 'DTT 6T/26', 'DTT_6M'],
            'TT_DTT_Q_YoY': ['TT DTT Q YoY', 'TT_DTT_Q_YoY', '%TT DTT Q YoY'],
            'TT_DTT_6M_YoY': ['TT DTT 6M YoY', 'TT_DTT_6M_YoY', '%TT DTT 6M YoY'],
            'LNST_Q2': ['LNST Q2/26 (tỷ đồng)', 'LNST Q2/26', 'LNST_Q2'],
            'LNST_6M': ['LNST 6T/26 (tỷ đồng)', 'LNST 6T/26', 'LNST_6M'],
            'TT_LNST_Q_YoY': ['TT LNST Q YoY', 'TT_LNST_Q_YoY', '%TT LNST Q YoY'],
            'TT_LNST_6M_YoY': ['TT LNST 6M YoY', 'TT_LNST_6M_YoY', '%TT LNST 6M YoY'],
            'TT_DTT_QoQ': ['TT DTT QoQ', 'TT_DTT_QoQ'],
            'TT_LNST_QoQ': ['TT LNST QoQ', 'TT_LNST_QoQ']
        }

        PERCENT_COLS = {
            'TT_DTT_Q_YoY', 'TT_DTT_6M_YoY', 
            'TT_LNST_Q_YoY', 'TT_LNST_6M_YoY', 
            'TT_DTT_QoQ', 'TT_LNST_QoQ'
        }

        matched_cols = {}
        for target, aliases in column_mapping.items():
            for col in df.columns:
                if any(alias.lower() == col.lower() or alias.lower() in col.lower() for alias in aliases):
                    matched_cols[col] = target
                    break

        df_renamed = df.rename(columns=matched_cols)
        
        if 'MaCP' not in df_renamed.columns:
            print("❌ LỖI: Không tìm thấy cột 'Mã CP' trong file Excel.")
            return False

        df_filtered = df_renamed.dropna(subset=['MaCP']).copy()
        df_filtered['MaCP'] = df_filtered['MaCP'].astype(str).str.strip()
        df_filtered = df_filtered[df_filtered['MaCP'] != 'nan']

        df_filtered = normalize_percentage_columns(df_filtered, PERCENT_COLS)

        published_records = []
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
            
            if rec.get('DTT_Q2') is not None or rec.get('LNST_Q2') is not None:
                published_records.append(rec)

        large_mid_positive = [
            r for r in published_records 
            if r.get('VonHoa') is not None and r['VonHoa'] >= 1000
            and r.get('TT_LNST_Q_YoY') is not None and r['TT_LNST_Q_YoY'] > 0
        ]
        
        top30_selected = sorted(large_mid_positive, key=lambda x: x['TT_LNST_Q_YoY'], reverse=True)[:30]

        total_companies = len(df_filtered)
        published_count = len(published_records)
        
        sum_published_mkt_cap = sum(r['VonHoa'] for r in published_records if r.get('VonHoa') is not None)
        sum_all_mkt_cap = sum(r['VonHoa'] for r in published_records if r.get('VonHoa') is not None)
        mkt_cap_pct = round((sum_published_mkt_cap / sum_all_mkt_cap * 100), 2) if sum_all_mkt_cap > 0 else 16.9

        profitable_comp = len([r for r in published_records if r.get('LNST_Q2') is not None and r['LNST_Q2'] > 0])
        loss_comp = len([r for r in published_records if r.get('LNST_Q2') is not None and r['LNST_Q2'] < 0])
        profit_ratio = round((profitable_comp / published_count * 100), 2) if published_count > 0 else 0.0

        report_data = {
            "executive_summary": {
                "total_companies": total_companies,
                "published_count": published_count,
                "published_market_cap_percentage": mkt_cap_pct,
                "profitable_companies": profitable_comp,
                "loss_companies": loss_comp,
                "profit_ratio_percentage": profit_ratio
            },
            "top_30_large_mid_growth": top30_selected
        }

        json_out_path = os.path.join(script_dir, "summary_report.json")
        with open(json_out_path, 'w', encoding='utf-8') as f:
            json.dump(report_data, f, ensure_ascii=False, indent=2)

        txt_out_path = os.path.join(script_dir, "summary_report.txt")
        with open(txt_out_path, 'w', encoding='utf-8') as f:
            f.write("=== TỔNG QUAN KẾT QUẢ KINH DOANH Q2 ===\n")
            f.write(f"Tiến độ công bố: {published_count}/{total_companies} doanh nghiệp ({mkt_cap_pct}% vốn hóa)\n")
            f.write(f"Tỷ lệ có lãi: {profit_ratio}% ({profitable_comp} DN Lãi / {loss_comp} DN Lỗ)\n\n")
            f.write("=== TOP 30 DOANH NGHIỆP VỐN HÓA LỚN & VỪA TĂNG TRƯỜNG LNST Q2 DƯƠNG CAO NHẤT ===\n")
            for idx, item in enumerate(top30_selected, 1):
                f.write(f"{idx:02d}. {item['MaCP']} - Vốn hóa: {item.get('VonHoa','-'):,} tỷ - TT LNST Q2 YoY: +{item.get('TT_LNST_Q_YoY','-'):,}%\n")

        # Xuất file Excel Top 30 cập nhật
        df_top30_export = pd.DataFrame(top30_selected)
        excel_out_path = os.path.join(script_dir, "Top30_KQKD_Q2_Updated.xlsx")
        df_top30_export.to_excel(excel_out_path, index=False)

        print(f"✅ Đã trích xuất thành công Top {len(top30_selected)} DN Vốn hóa Lớn & Vừa tăng trưởng LNST Q2 dương cao nhất!")
        print(f"📁 Báo cáo TXT: {txt_out_path}")
        print(f"📁 Báo cáo Excel: {excel_out_path}")

        html_files = update_html_dashboards(script_dir, published_records)
        
        sync_git_repository(script_dir, html_files)
        
        return True

    except Exception as e:
        print(f"❌ Lỗi xử lý dữ liệu: {e}")
        import traceback
        traceback.print_exc()
        return False

def update_html_dashboards(script_dir, records):
    """Cập nhật dữ liệu vào TẤT CẢ các file index.html được tìm thấy"""
    html_paths = find_all_index_html_files(script_dir)
    
    if not html_paths:
        print("⚠️ Không tìm thấy file index.html nào để cập nhật dữ liệu web.")
        return []

    json_str = json.dumps(records, ensure_ascii=False)
    timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    updated_files = []

    for path in html_paths:
        try:
            with open(path, 'r', encoding='utf-8') as f:
                content = f.read()

            pattern = r'(let|var|const)\s+rawStockData\s*=\s*\[.*?\]\s*;'
            replacement = f'let rawStockData = {json_str};'

            if not re.search(pattern, content, flags=re.DOTALL):
                print(f"⚠️ CẢNH BÁO: File {path} không có mảng 'rawStockData = [...]' để ghi dữ liệu!")
                continue

            updated_content = re.sub(pattern, replacement, content, flags=re.DOTALL)

            timestamp_tag = f"<!-- LAST_UPDATED: {timestamp_str} -->"
            if "<!-- LAST_UPDATED:" in updated_content:
                updated_content = re.sub(r'<!-- LAST_UPDATED:.*?-->', timestamp_tag, updated_content)
            else:
                updated_content += f"\n{timestamp_tag}"

            with open(path, 'w', encoding='utf-8') as f:
                f.write(updated_content)

            print(f"🎉 Cập nhật thành công dữ liệu thực tế vào Web Dashboard ({path})!")
            updated_files.append(path)
        except Exception as e:
            print(f"⚠️ Không thể cập nhật {path}: {e}")

    return updated_files

def sync_git_repository(script_dir, html_files):
    """
    Tự động commit và push CHỈ DUY NHẤT các file index.html đã cập nhật lên GitHub repository.
    Đã chuẩn hóa ký tự ổ đĩa trên Windows để tránh lỗi pathspec trong Git.
    """
    if not html_files:
        print("⚠️ Không có file index.html nào để đồng bộ Git.")
        return

    original_cwd = os.getcwd()

    try:
        for html_path in html_files:
            file_dir = os.path.dirname(html_path)
            
            try:
                repo_dir_bytes = subprocess.check_output(
                    ["git", "rev-parse", "--show-toplevel"], 
                    cwd=file_dir,
                    stderr=subprocess.DEVNULL
                )
                repo_dir = repo_dir_bytes.decode('utf-8', errors='ignore').strip()
            except Exception:
                repo_dir = file_dir

            print(f"🔄 Đang đồng bộ file {os.path.basename(html_path)} tại Git repo: {repo_dir}...")
            
            os.chdir(repo_dir)

            # Pull rebase trước để tránh bị rejected
            subprocess.run(["git", "pull", "--rebase"], capture_output=True, text=True)

            # Chuẩn hóa cả html_path và repo_dir để cùng dạng chữ hoa ổ đĩa C:
            norm_html = os.path.normpath(os.path.realpath(html_path))
            norm_repo = os.path.normpath(os.path.realpath(repo_dir))

            if len(norm_html) >= 2 and norm_html[1] == ':' and len(norm_repo) >= 2 and norm_repo[1] == ':':
                norm_html = norm_html[0].upper() + norm_html[1:]
                norm_repo = norm_repo[0].upper() + norm_repo[1:]

            rel_path = os.path.relpath(norm_html, norm_repo).replace("\\", "/")
            
            # Diagnostic: Kiểm tra status trước khi add
            status_before = subprocess.check_output(["git", "status", "--porcelain", rel_path], text=True).strip()
            
            subprocess.run(["git", "add", rel_path], capture_output=True, text=True)
            
            # Diagnostic: Kiểm tra status sau khi add (staged)
            status_after = subprocess.check_output(["git", "status", "--porcelain", rel_path], text=True).strip()

            commit_msg = f"Auto-update index.html data: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            commit_res = subprocess.run(["git", "commit", "-m", commit_msg], capture_output=True, text=True)

            push_res = subprocess.run(["git", "push"], capture_output=True, text=True)
            output_combined = (push_res.stdout or "") + (push_res.stderr or "")
            
            if commit_res.returncode == 0:
                print(f"🚀 [Commit Mới Món] Đã commit và push thành công {rel_path} lên GitHub!")
            elif "nothing to commit" in (commit_res.stdout or "") or "nothing to commit" in (commit_res.stderr or ""):
                print(f"ℹ️ [Dữ liệu Trùng Khớp] File {rel_path} không có thay đổi nào mới so với commit trước đó trên Git.")
                if status_before:
                    print(f"   🔍 Trạng thái local file: {status_before}")
            else:
                print(f"  🔍 Kết quả Push: {output_combined.strip()}")
    except Exception as e:
        print(f"⚠️ Lỗi khi đồng bộ Git: {e}")
    finally:
        os.chdir(original_cwd)

if __name__ == "__main__":
    print("=" * 70)
    print("🚀 HỆ THỐNG TRÍCH XUẤT KQKD & ĐỒNG BỘ GIT - YUANTA VIỆT NAM")
    print("=" * 70)
    extract_top_30_growth()