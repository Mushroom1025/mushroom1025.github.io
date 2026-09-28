import json
import re
import pandas as pd

def parse_json_to_excel(json_filename="export.json", excel_filename="回覆資料庫.xlsx"):
    try:
        with open(json_filename, "r", encoding="utf-8") as f:
            data = json.load(f)
    except FileNotFoundError:
        print(f"❌ 找不到檔案：{json_filename}，請確認檔名是否正確！")
        return

    records = []
    messages = data.get("messages", [])

    for msg in messages:
        embeds = msg.get("embeds", [])
        for embed in embeds:
            # 確保這是我們要的問卷回覆 Embed（標題含有「作答者：」）
            title = embed.get("title", "")
            if not title.startswith("作答者："):
                continue

            # 抓取作答者姓名
            author_name = title.replace("作答者：", "").strip()
            
            row = {
                "作答者": author_name,
                "聯絡方式": "",
                "留言備註": ""
            }

            # 解析 Fields (聯絡方式、留言備註、作答明細)
            fields = embed.get("fields", [])
            for field in fields:
                name = field.get("name", "").strip()
                val = field.get("value", "").strip()

                if "聯絡方式" in name:
                    row["聯絡方式"] = val
                elif "留言備註" in name:
                    row["留言備註"] = val
                elif "作答明細" in name:
                    # 使用 Regex 切出 Q1, Q2... 及其回答
                    # 匹配格式如: **Q1. 題目**\n：答案
                    qa_matches = re.findall(r"\*\*(Q\d+\..*?)\*\*\n\s*[:：]\s*(.*?)(?=\n\n\*\*Q|\Z)", val, re.DOTALL)
                    for q_title, q_ans in qa_matches:
                        clean_q = q_title.strip()
                        clean_a = q_ans.strip()
                        row[clean_q] = clean_a

            if row["聯絡方式"]:
                records.append(row)

    if not records:
        print("⚠️ 沒有找到任何有效的問卷回覆資料。")
        return

    # 轉為 Pandas DataFrame
    df = pd.DataFrame(records)

    # 重新排列欄位順序：作答者、聯絡方式、留言備註排在最前面，後面接著 Q1, Q2...
    base_cols = ["作答者", "聯絡方式", "留言備註"]
    q_cols = [c for c in df.columns if c not in base_cols]
    
    # 依題號 Q1, Q2, Q3 排序
    q_cols_sorted = sorted(q_cols, key=lambda x: int(re.search(r"Q(\d+)", x).group(1)) if re.search(r"Q(\d+)", x) else 999)
    
    df = df[base_cols + q_cols_sorted]

    # 匯出至 Excel
    df.to_excel(excel_filename, index=False)
    print(f"✅ 成功解析 {len(records)} 筆回覆！已儲存至：{excel_filename}")

if __name__ == "__main__":
    parse_json_to_excel()