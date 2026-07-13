import streamlit as st
import pandas as pd

# stlite (Pyodide Wasm) 環境での pyarrow 互換性パッチ
# plotly.expressはDataFrameを渡すとnarwhals経由でpyarrow.ChunkedArrayを要求するため
# 全チャートはgo.Figureで生データ(list)を使用し、narwhalsを完全回避する
try:
    import pyarrow as pa
    if not hasattr(pa, 'ChunkedArray'):
        class _DummyChunkedArray: pass
        pa.ChunkedArray = _DummyChunkedArray
    if not hasattr(pa, 'Table'):
        class _DummyTable: pass
        pa.Table = _DummyTable
except Exception:
    pass

import plotly.graph_objects as go
import io
import os
from parser import combine_multiple_files
from validator import check_crf_errors

def clean_df_for_streamlit(df):
    if df is None or df.empty:
        return df
    df_clean = df.copy()
    for col in df_clean.columns:
        if pd.api.types.is_string_dtype(df_clean[col]):
            df_clean[col] = df_clean[col].astype(object)
    return df_clean


# --- ページ設定 ---
st.set_page_config(
    page_title="CRF Data Integrator & Validator",
    page_icon="🧬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- カスタムCSSによるリッチデザインの適用 ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&display=swap');
    
    /* フォント適用 */
    html, body, [class*="css"], .stMarkdown {
        font-family: 'Outfit', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }
    
    /* メイン背景とサイドバー */
    .stApp {
        background: linear-gradient(135deg, #0e1117 0%, #161a24 100%);
    }
    
    /* カードデザイン (Glassmorphism) */
    div[data-testid="stMetricValue"] {
        font-size: 2.2rem !important;
        font-weight: 700 !important;
        color: #00f2fe !important;
    }
    div[data-testid="stMetricLabel"] {
        font-size: 1rem !important;
        font-weight: 500 !important;
        color: #a3b8cc !important;
    }
    
    /* カスタムカードコンテナ */
    .custom-card {
        background: rgba(255, 255, 255, 0.03);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 16px;
        padding: 20px;
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.2);
        margin-bottom: 20px;
        transition: all 0.3s ease;
    }
    
    .custom-card:hover {
        transform: translateY(-4px);
        border-color: rgba(0, 242, 254, 0.3);
        box-shadow: 0 12px 40px 0 rgba(0, 242, 254, 0.1);
    }
    
    /* タイトルグラデーション */
    .title-gradient {
        background: linear-gradient(90deg, #00f2fe 0%, #4facfe 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
        font-size: 3rem;
        margin-bottom: 5px;
    }
    
    .subtitle {
        color: #8b9bb4;
        font-size: 1.1rem;
        margin-bottom: 30px;
    }
    
    /* タブスタイル */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        background-color: rgba(255, 255, 255, 0.02);
        padding: 8px;
        border-radius: 12px;
        border: 1px solid rgba(255, 255, 255, 0.05);
    }
    
    .stTabs [data-baseweb="tab"] {
        height: 45px;
        border-radius: 8px;
        background-color: transparent;
        border: none;
        color: #8b9bb4;
        font-weight: 600;
        font-size: 0.95rem;
        transition: all 0.2s ease;
    }
    
    .stTabs [data-baseweb="tab"]:hover {
        color: #00f2fe;
        background-color: rgba(255, 255, 255, 0.05);
    }
    
    .stTabs [aria-selected="true"] {
        background-color: rgba(0, 242, 254, 0.15) !important;
        color: #00f2fe !important;
        border: 1px solid rgba(0, 242, 254, 0.3) !important;
    }
    
    /* テーブルヘッダーと全体のフォントサイズ微調整 */
    .dataframe {
        font-size: 0.9rem;
    }
    
    /* エラーアラートのスタイル */
    .error-count-badge {
        background: linear-gradient(135deg, #ff416c 0%, #ff4b2b 100%);
        color: white;
        padding: 4px 12px;
        border-radius: 20px;
        font-weight: 700;
        font-size: 0.9rem;
        display: inline-block;
        margin-left: 10px;
    }
    
    .success-badge {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
        color: white;
        padding: 4px 12px;
        border-radius: 20px;
        font-weight: 700;
        font-size: 0.9rem;
        display: inline-block;
        margin-left: 10px;
    }
</style>
""", unsafe_allow_html=True)

# --- タイトルヘッダー ---
st.markdown('<div class="title-gradient">🧬 LSRegistry CRF Tool</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">複数の症例報告書（CRF）の自動結合・可視化・矛盾データの検出システム</div>', unsafe_allow_html=True)

# --- セッションステート初期化 ---
if "uploaded_files_cache" not in st.session_state:
    st.session_state.uploaded_files_cache = {}
if "dfs" not in st.session_state:
    st.session_state.dfs = None
if "raw_results" not in st.session_state:
    st.session_state.raw_results = None
if "errors_df" not in st.session_state:
    st.session_state.errors_df = None

# --- サイドバー：ファイルアップローダー ---
with st.sidebar:
    st.markdown("### 📁 CRFファイル アップロード")
    st.markdown("多数の患者のCRF Excelファイル（`.xlsx`）を選択してください。")
    
    uploaded_files = st.file_uploader(
        "Excelファイルを選択してください",
        type=["xlsx"],
        accept_multiple_files=True,
        key="uploader"
    )
    
    # サンプルデータのロードボタン（デモ・テスト用）
    st.markdown("---")
    st.markdown("### 🧪 テスト・検証用")
    load_samples = st.button("デモ用サンプルデータをロード", use_container_width=True)
    
    if load_samples:
        # ローカルとWasm環境両方のファイルパスを考慮
        sample_paths = {
            "★CRF20260210_normal.xlsx": "★CRF20260210_normal.xlsx" if os.path.exists("★CRF20260210_normal.xlsx") else "sample_normal.xlsx",
            "★CRF20260210_error_male.xlsx": "★CRF20260210_error_male.xlsx" if os.path.exists("★CRF20260210_error_male.xlsx") else "sample_error_male.xlsx",
            "★CRF20260210_error_female.xlsx": "★CRF20260210_error_female.xlsx" if os.path.exists("★CRF20260210_error_female.xlsx") else "sample_error_female.xlsx"
        }
        
        # ファイルが存在するか確認
        samples_exist = True
        for path in sample_paths.values():
            if not os.path.exists(path):
                samples_exist = False
                break
                
        if samples_exist:
            st.session_state.uploaded_files_cache = sample_paths
            with st.spinner("サンプルデータを解析中..."):
                dfs, raw_results = combine_multiple_files(sample_paths)
                st.session_state.dfs = dfs
                st.session_state.raw_results = raw_results
                st.session_state.errors_df = check_crf_errors(raw_results)
            st.success("デモ用データをロードしました！")
            st.rerun()
        else:
            st.error("テスト用ファイルが見つかりません。プロジェクトのルートディレクトリにファイルを生成してください。")

# ファイルが手動でアップロードされた場合の処理
if uploaded_files:
    # 新しいアップロードファイルを辞書に変換
    new_files = {}
    for f in uploaded_files:
        new_files[f.name] = f
        
    # セッションステートに保存
    st.session_state.uploaded_files_cache = new_files
    
    # データの結合処理
    with st.spinner("アップロードされたCRFファイルを解析・結合中..."):
        dfs, raw_results = combine_multiple_files(new_files)
        st.session_state.dfs = dfs
        st.session_state.raw_results = raw_results
        st.session_state.errors_df = check_crf_errors(raw_results)

# --- メインコンテンツ領域 ---
if not st.session_state.uploaded_files_cache:
    st.info("👈 左側のサイドバーからCRF Excelファイルをアップロードするか、デモ用サンプルデータをロードしてください。")
    
    # プレースホルダーデザインを表示
    st.markdown("""
    <div style="background: rgba(255, 255, 255, 0.02); border: 1px dashed rgba(255, 255, 255, 0.1); border-radius: 16px; padding: 60px; text-align: center; margin-top: 40px;">
        <h3 style="color: #8b9bb4; font-weight: 500;">ファイルをロードすると、ここにダッシュボードと結合結果が表示されます</h3>
        <p style="color: #5f6e85;">Excelファイルの各シート（登録票、登録時情報、家系情報、サーベイランス等）を自動的にパースし、エラーの抽出を行います。</p>
    </div>
    """, unsafe_allow_html=True)
else:
    dfs = st.session_state.dfs
    errors_df = st.session_state.errors_df
    num_files = len(st.session_state.uploaded_files_cache)
    num_errors = len(errors_df)
    
    # タブの作成
    tab_dashboard, tab_combine, tab_errors = st.tabs([
        "📊 可視化ダッシュボード",
        "📁 結合データの表示・ダウンロード",
        f"⚠️ エラー・矛盾データの検証 ({num_errors})"
    ])
    
    # ==========================================
    # タブ1: 可視化ダッシュボード
    # ==========================================
    with tab_dashboard:
        # 1. メトリクスカード
        st.markdown("### 📊 主要インジケータ")
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        
        # 患者数計算
        num_patients = 0
        if dfs and "登録票" in dfs and not dfs["登録票"].empty:
            num_patients = dfs["登録票"]["カルテ番号"].nunique()
            
        # 性別比計算
        male_pct, female_pct = 0.0, 0.0
        if dfs and "登録票" in dfs and not dfs["登録票"].empty and "性別" in dfs["登録票"].columns:
            gender_counts = dfs["登録票"]["性別"].value_counts()
            total_gender = gender_counts.sum()
            if total_gender > 0:
                male_pct = (gender_counts.get("男", 0) / total_gender) * 100
                female_pct = (gender_counts.get("女", 0) / total_gender) * 100
                
        # 平均年齢
        avg_age_str = "N/A"
        if dfs and "登録時情報_基本" in dfs and not dfs["登録時情報_基本"].empty:
            # 登録票から登録時年齢（「50歳4ヶ月」など）を取得するか、計算する
            # ここでは登録票に「登録時年齢（16歳以上）」がある。
            # 「50歳4ヶ月」などの文字列から数値にする簡易パース
            ages = []
            age_col = [col for col in dfs["登録票"].columns if "年齢" in col]
            if age_col:
                for a_val in dfs["登録票"][age_col[0]]:
                    if pd.notna(a_val):
                        try:
                            # 「50歳4ヶ月」などの最初の数値を抽出
                            import re
                            match = re.search(r'(\d+)', str(a_val))
                            if match:
                                ages.append(int(match.group(1)))
                        except:
                            pass
            if ages:
                avg_age_str = f"{sum(ages)/len(ages):.1f} 歳"

        with col_m1:
            st.markdown(f"""
            <div class="custom-card">
                <div style="font-size: 0.9rem; color: #a3b8cc; font-weight: 500;">解析したCRFファイル数</div>
                <div style="font-size: 2.2rem; font-weight: 700; color: #00f2fe; margin-top: 10px;">{num_files}</div>
            </div>
            """, unsafe_allow_html=True)
            
        with col_m2:
            st.markdown(f"""
            <div class="custom-card">
                <div style="font-size: 0.9rem; color: #a3b8cc; font-weight: 500;">登録患者数</div>
                <div style="font-size: 2.2rem; font-weight: 700; color: #38ef7d; margin-top: 10px;">{num_patients}</div>
            </div>
            """, unsafe_allow_html=True)
            
        with col_m3:
            st.markdown(f"""
            <div class="custom-card">
                <div style="font-size: 0.9rem; color: #a3b8cc; font-weight: 500;">平均年齢</div>
                <div style="font-size: 2.2rem; font-weight: 700; color: #fccb90; margin-top: 10px;">{avg_age_str}</div>
            </div>
            """, unsafe_allow_html=True)
            
        with col_m4:
            color = "#ff4b2b" if num_errors > 0 else "#38ef7d"
            st.markdown(f"""
            <div class="custom-card">
                <div style="font-size: 0.9rem; color: #a3b8cc; font-weight: 500;">検出されたエラー数</div>
                <div style="font-size: 2.2rem; font-weight: 700; color: {color}; margin-top: 10px;">{num_errors}</div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("### 📈 チャート分析")
        
        # グラフ用のカラム
        col_g1, col_g2 = st.columns(2)
        
        # グラフ1: 性別の割合
        with col_g1:
            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.markdown("#### 🧬 性別割合")
            if dfs and "登録票" in dfs and not dfs["登録票"].empty and "性別" in dfs["登録票"].columns:
                gender_counts = dfs["登録票"]["性別"].value_counts()
                labels = [str(x) for x in gender_counts.index.tolist()]
                values = [int(x) for x in gender_counts.values.tolist()]
                colors = ["#e0c3fc", "#8ec5fc", "#ffd1ff", "#ff9a9e"]
                fig_gender = go.Figure(go.Pie(
                    labels=labels,
                    values=values,
                    hole=0.4,
                    marker=dict(colors=colors[:len(labels)])
                ))
                fig_gender.update_layout(
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    margin=dict(t=10, b=10, l=10, r=10),
                    height=300,
                    template="plotly_dark"
                )
                st.plotly_chart(fig_gender, use_container_width=True)
            else:
                st.write("性別データがありません")
            st.markdown('</div>', unsafe_allow_html=True)
            
        # グラフ2: 遺伝子の分布
        with col_g2:
            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.markdown("#### 🧬 遺伝子バリアント陽性頻度")
            gene_list = []
            if dfs and "登録票" in dfs and not dfs["登録票"].empty:
                cols = [c for c in dfs["登録票"].columns if "遺伝子" in c]
                for col in cols:
                    gene_list.extend(dfs["登録票"][col].dropna().tolist())
                    
            if gene_list:
                gene_counts = pd.Series(gene_list).value_counts()
                x_vals = [str(x) for x in gene_counts.index.tolist()]
                y_vals = [int(x) for x in gene_counts.values.tolist()]
                bar_colors = ["#8ec5fc", "#e0c3fc", "#ffd1ff", "#ff9a9e", "#a1c4fd",
                              "#fddb92", "#d1c4e9", "#b2ebf2"]
                fig_genes = go.Figure(go.Bar(
                    x=x_vals,
                    y=y_vals,
                    marker=dict(
                        color=bar_colors[:len(x_vals)],
                        line=dict(color='rgba(255,255,255,0.1)', width=1)
                    )
                ))
                fig_genes.update_layout(
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    margin=dict(t=10, b=10, l=10, r=10),
                    height=300,
                    showlegend=False,
                    template="plotly_dark"
                )
                st.plotly_chart(fig_genes, use_container_width=True)
            else:
                st.write("遺伝子データがありません")
            st.markdown('</div>', unsafe_allow_html=True)
            
        col_g3, col_g4 = st.columns(2)
        
        # グラフ3: 最も多い病名
        with col_g3:
            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.markdown("#### 腫瘍・病名頻度（上位10項目）")
            if dfs and "登録時情報_診断治療歴" in dfs and not dfs["登録時情報_診断治療歴"].empty:
                diag_counts = dfs["登録時情報_診断治療歴"]["病名"].value_counts().head(10)
                y_vals = [str(x) for x in diag_counts.index.tolist()]
                x_vals = [int(x) for x in diag_counts.values.tolist()]
                # 数値に応じたグラデーションカラー
                max_v = max(x_vals) if x_vals else 1
                grad_colors = [f"rgba({int(160 + 80*v/max_v)}, {int(100 + 80*(1-v/max_v))}, 220, 0.85)"
                               for v in x_vals]
                fig_diag = go.Figure(go.Bar(
                    y=y_vals,
                    x=x_vals,
                    orientation='h',
                    marker=dict(color=grad_colors)
                ))
                fig_diag.update_layout(
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    margin=dict(t=10, b=10, l=10, r=10),
                    height=320,
                    template="plotly_dark"
                )
                st.plotly_chart(fig_diag, use_container_width=True)
            else:
                st.write("診断歴データがありません")
            st.markdown('</div>', unsafe_allow_html=True)

        # グラフ4: 医療機関ごとの登録数
        with col_g4:
            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.markdown("#### 病院・医療機関別 登録割合")
            hosp_col = []
            if dfs and "登録票" in dfs and not dfs["登録票"].empty:
                hosp_col = [c for c in dfs["登録票"].columns if "登録医療機関" in c]
                
            if hosp_col:
                hosp_counts = dfs["登録票"][hosp_col[0]].value_counts()
                hosp_labels = [str(x) for x in hosp_counts.index.tolist()]
                hosp_values = [int(x) for x in hosp_counts.values.tolist()]
                hosp_colors = ["#f7971e", "#ffd200", "#e96c5e", "#a18cd1",
                               "#fbc2eb", "#84fab0", "#8fd3f4", "#d4fc79"]
                fig_hosp = go.Figure(go.Pie(
                    labels=hosp_labels,
                    values=hosp_values,
                    marker=dict(colors=hosp_colors[:len(hosp_labels)])
                ))
                fig_hosp.update_layout(
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    margin=dict(t=10, b=10, l=10, r=10),
                    height=320,
                    template="plotly_dark"
                )
                st.plotly_chart(fig_hosp, use_container_width=True)
            else:
                st.write("医療機関データがありません")
            st.markdown('</div>', unsafe_allow_html=True)

    # ==========================================
    # タブ2: 結合データの表示・ダウンロード
    # ==========================================
    with tab_combine:
        st.markdown("### 📂 結合データ プレビュー")
        st.write("シートを選択して、結合された全患者のデータを確認できます。")
        
        # プレビュー用のシート選択
        sheets = list(dfs.keys())
        selected_sheet = st.selectbox("結合シートの切り替え", sheets)
        
        if selected_sheet and not dfs[selected_sheet].empty:
            st.dataframe(clean_df_for_streamlit(dfs[selected_sheet]), use_container_width=True)
        else:
            st.warning(f"シート「{selected_sheet}」にはデータが存在しないか、空です。")
            
        st.markdown("---")
        st.markdown("### 📥 全結合データの一括ダウンロード")
        st.write("アップロードされたすべてのCRF Excelファイルをパースし、シートごとに結合した単一のExcelワークブックをダウンロードできます。")
        
        # Excelファイルのバイトデータをメモリ上で生成
        output_excel = io.BytesIO()
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            for s_name, s_df in dfs.items():
                if not s_df.empty:
                    s_df.to_excel(writer, sheet_name=s_name, index=False)
        excel_data = output_excel.getvalue()
        
        # ダウンロードボタン
        st.download_button(
            label="📥 結合データをExcelでダウンロード",
            data=excel_data,
            file_name="combined_crf_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # ==========================================
    # タブ3: エラー・矛盾データの検証
    # ==========================================
    with tab_errors:
        st.markdown("### ⚠️ 辻褄の合わないデータの検出")
        st.write("性別と疾患の矛盾（男性の卵巣がん・子宮がん、女性の前立腺がん等）や、家系情報、検査等の矛盾データを自動抽出しています。")
        
        if num_errors > 0:
            st.markdown(f'<div class="error-count-badge">⚠️ {num_errors} 件の矛盾を検出</div>', unsafe_allow_html=True)
            st.markdown("<br><br>", unsafe_allow_html=True)
            
            # エラーのフィルタリングオプション
            categories = ["すべて"] + list(errors_df["エラーの分類"].unique())
            selected_cat = st.selectbox("エラー分類でフィルタ", categories)
            
            filtered_errors = errors_df
            if selected_cat != "すべて":
                filtered_errors = errors_df[errors_df["エラーの分類"] == selected_cat]
                
            # インタラクティブテーブルでのエラー表示
            st.dataframe(
                clean_df_for_streamlit(filtered_errors),
                column_config={
                    "ファイル名": st.column_config.TextColumn("ファイル名", width="medium"),
                    "カルテ番号": st.column_config.TextColumn("カルテ番号", width="small"),
                    "エラーの分類": st.column_config.TextColumn("分類", width="small"),
                    "項目/対象": st.column_config.TextColumn("対象", width="medium"),
                    "検出された値": st.column_config.TextColumn("検出値", width="medium"),
                    "エラーメッセージ": st.column_config.TextColumn("説明", width="large"),
                },
                use_container_width=True,
                hide_index=True
            )
            
            # エラーデータのダウンロード
            st.markdown("#### 📥 エラーリストのダウンロード")
            
            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                # CSV形式
                csv_errors = filtered_errors.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📥 エラーリストをCSVでダウンロード",
                    data=csv_errors,
                    file_name="crf_validation_errors.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            with col_dl2:
                # Excel形式
                err_output = io.BytesIO()
                with pd.ExcelWriter(err_output, engine='openpyxl') as writer:
                    filtered_errors.to_excel(writer, sheet_name="Validation Errors", index=False)
                err_excel_data = err_output.getvalue()
                
                st.download_button(
                    label="📥 エラーリストをExcelでダウンロード",
                    data=err_excel_data,
                    file_name="crf_validation_errors.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
        else:
            st.markdown('<div class="success-badge">✅ 不整合なし</div>', unsafe_allow_html=True)
            st.markdown("<br><br>", unsafe_allow_html=True)
            st.success("🎉 素晴らしい！矛盾データや不整合は一切検出されませんでした。すべてのデータの辻褄が合っています。")
            st.markdown("""
            <div style="background: rgba(17, 153, 142, 0.05); border: 1px solid rgba(56, 239, 125, 0.2); border-radius: 16px; padding: 40px; text-align: center; margin-top: 20px;">
                <h3 style="color: #38ef7d; font-weight: 600;">Data is Clean!</h3>
                <p style="color: #8b9bb4;">アップロードされたすべてのデータは整合性が確認され、臨床研究への使用に適しています。</p>
            </div>
            """, unsafe_allow_html=True)
