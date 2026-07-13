import pandas as pd

# 女性特有の疾患のキーワード
FEMALE_DISEASES = [
    "卵巣癌", "子宮癌", "子宮体癌", "子宮頸癌", "卵巣嚢腫", "子宮筋腫", 
    "子宮内膜癌", "子宮内膜増殖症", "子宮肉腫", "卵巣境界悪性腫瘍", 
    "卵巣のう腫", "子宮がん", "卵巣がん", "子宮内膜"
]

# 男性特有の疾患のキーワード
MALE_DISEASES = [
    "前立腺癌", "前立腺肥大症", "精巣癌", "前立腺がん", "精巣がん"
]

# 女性特有の検査・治療のキーワード
FEMALE_PROCEDURES = [
    "子宮内膜組織診", "経膣超音波検査", "子宮全摘術", "両側付属器切除術", "付属器切除術"
]

# 男性関係者キーワード
MALE_RELATIONS = ["父", "祖父", "父方祖父", "母方祖父", "伯父", "叔父", "兄弟", "息子", "夫", "叔父・伯父"]
# 女性関係者キーワード
FEMALE_RELATIONS = ["母", "祖母", "父方祖母", "母方祖母", "伯母", "叔母", "姉妹", "娘", "妻", "叔母・伯母"]

def check_crf_errors(raw_results):
    """
    パースされた生のCRFデータリストを精査し、不整合（エラー）を抽出する。
    """
    errors = []

    for parsed in raw_results:
        filename = parsed["登録票"].get("ファイル名", "Unknown")
        karte_no = parsed["登録票"].get("カルテ番号", "Unknown")
        
        # 性別の抽出
        # 登録票から性別を取得（「女」または「男」）
        gender_raw = parsed["登録票"].get("患者情報_性別", parsed["登録票"].get("性別"))
        if gender_raw is None:
            # 登録票の別キーを探す
            for k, v in parsed["登録票"].items():
                if "性別" in k:
                    gender_raw = v
                    break
        
        gender = str(gender_raw).strip() if gender_raw is not None else ""
        
        # --- 1. 本人の性別と疾患の矛盾のチェック (登録時情報_診断治療歴) ---
        diags = parsed.get("登録時情報_診断治療歴", [])
        for diag in diags:
            disease = diag.get("病名", "")
            
            if gender == "男":
                # 男性の女性特有疾患チェック
                matched_female_dis = [fd for fd in FEMALE_DISEASES if fd in disease]
                if matched_female_dis:
                    errors.append({
                        "ファイル名": filename,
                        "カルテ番号": karte_no,
                        "エラーの分類": "診断・治療歴",
                        "項目/対象": "性別と病名の矛盾",
                        "検出された値": f"性別: 男 / 病名: {disease}",
                        "エラーメッセージ": f"男性患者に女性特有の疾患「{disease}」が登録されています。"
                    })
                    
            elif gender == "女":
                # 女性の男性特有疾患チェック
                matched_male_dis = [md for md in MALE_DISEASES if md in disease]
                if matched_male_dis:
                    errors.append({
                        "ファイル名": filename,
                        "カルテ番号": karte_no,
                        "エラーの分類": "診断・治療歴",
                        "項目/対象": "性別と病名の矛盾",
                        "検出された値": f"性別: 女 / 病名: {disease}",
                        "エラーメッセージ": f"女性患者に男性特有の疾患「{disease}」が登録されています。"
                    })




        # --- 3. 本人の性別とサーベイランス（検査）の矛盾 (サーベイランス) ---
        surveillances = parsed.get("サーベイランス", [])
        for surv in surveillances:
            exam_type = surv.get("検査種別", "")
            if gender == "男":
                matched_female_proc = [fp for fp in FEMALE_PROCEDURES if fp in exam_type]
                if matched_female_proc:
                    errors.append({
                        "ファイル名": filename,
                        "カルテ番号": karte_no,
                        "エラーの分類": "サーベイランス",
                        "項目/対象": "性別と検査の矛盾",
                        "検出された値": f"性別: 男 / 検査種別: {exam_type}",
                        "エラーメッセージ": f"男性患者に女性特有の検査「{exam_type}」が登録されています。"
                    })

        # --- 4. 登録票の子宮全摘理由などの矛盾 (登録票) ---
        if gender == "男":
            for k, v in parsed["登録票"].items():
                if "子宮全摘" in k and v is not None and str(v).strip() != "" and pd.notna(v):
                    errors.append({
                        "ファイル名": filename,
                        "カルテ番号": karte_no,
                        "エラーの分類": "登録票",
                        "項目/対象": "性別と子宮全摘情報の矛盾",
                        "検出された値": f"性別: 男 / {k}: {v}",
                        "エラーメッセージ": f"男性患者に子宮全摘関連情報（{k}: {v}）が登録されています。"
                    })

        # --- 5. 家系情報の矛盾 (家系情報) ---
        family = parsed.get("家系情報", [])
        for fam in family:
            relation = fam.get("関係", "")
            f_gender = fam.get("性別", "")
            diseases = fam.get("疾患一覧", [])
            
            # 関係と性別の矛盾
            if relation in MALE_RELATIONS and f_gender in ["F", "女"]:
                errors.append({
                    "ファイル名": filename,
                    "カルテ番号": karte_no,
                    "エラーの分類": "家系情報",
                    "項目/対象": "関係と性別の矛盾",
                    "検出された値": f"関係: {relation} / 性別: {f_gender}",
                    "エラーメッセージ": f"家系情報において、男性の続柄「{relation}」の性別が「女性({f_gender})」として登録されています。"
                })
            elif relation in FEMALE_RELATIONS and f_gender in ["M", "男"]:
                errors.append({
                    "ファイル名": filename,
                    "カルテ番号": karte_no,
                    "エラーの分類": "家系情報",
                    "項目/対象": "関係と性別の矛盾",
                    "検出された値": f"関係: {relation} / 性別: {f_gender}",
                    "エラーメッセージ": f"家系情報において、女性の続柄「{relation}」の性別が「男性({f_gender})」として登録されています。"
                })
                
            # 血縁者の性別と疾患の不整合
            # 性別がM（男）または続柄が男性で、性別が空/男性の場合
            is_male_member = f_gender in ["M", "男"] or (f_gender == "" and relation in MALE_RELATIONS)
            is_female_member = f_gender in ["F", "女"] or (f_gender == "" and relation in FEMALE_RELATIONS)
            
            for dis in diseases:
                dis_name = dis.get("病名", "")
                
                if is_male_member:
                    matched_female_dis = [fd for fd in FEMALE_DISEASES if fd in dis_name]
                    if matched_female_dis:
                        errors.append({
                            "ファイル名": filename,
                            "カルテ番号": karte_no,
                            "エラーの分類": "家系情報",
                            "項目/対象": "血縁者の性別と病名の矛盾",
                            "検出された値": f"関係: {relation} (性別: {f_gender or '男'}) / 病名: {dis_name}",
                            "エラーメッセージ": f"家系情報において、男性の血縁者「{relation}」に女性特有の疾患「{dis_name}」が登録されています。"
                        })
                elif is_female_member:
                    matched_male_dis = [md for md in MALE_DISEASES if md in dis_name]
                    if matched_male_dis:
                        errors.append({
                            "ファイル名": filename,
                            "カルテ番号": karte_no,
                            "エラーの分類": "家系情報",
                            "項目/対象": "血縁者の性別と病名の矛盾",
                            "検出された値": f"関係: {relation} (性別: {f_gender or '女'}) / 病名: {dis_name}",
                            "エラーメッセージ": f"家系情報において、女性の血縁者「{relation}」に男性特有の疾患「{dis_name}」が登録されています。"
                        })

    return pd.DataFrame(errors) if errors else pd.DataFrame(columns=["ファイル名", "カルテ番号", "エラーの分類", "項目/対象", "検出された値", "エラーメッセージ"])
