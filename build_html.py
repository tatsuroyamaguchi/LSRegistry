import base64
import os
import json

def build_index_html():
    # 1. ソースコードの読み込み
    with open("app.py", "r", encoding="utf-8") as f:
        app_py = f.read()
    
    with open("parser.py", "r", encoding="utf-8") as f:
        parser_py = f.read()
        
    with open("validator.py", "r", encoding="utf-8") as f:
        validator_py = f.read()

    # 安全にJavaScript文字列に変換するためにjson.dumpsを使用
    app_py_json = json.dumps(app_py)
    parser_py_json = json.dumps(parser_py)
    validator_py_json = json.dumps(validator_py)

    # 2. ExcelファイルのBase64エンコード
    def get_base64(filepath):
        if os.path.exists(filepath):
            with open(filepath, "rb") as f:
                return base64.b64encode(f.read()).decode("utf-8")
        return ""

    normal_xlsx_b64 = get_base64("★CRF20260210_normal.xlsx")
    error_male_xlsx_b64 = get_base64("★CRF20260210_error_male.xlsx")
    error_female_xlsx_b64 = get_base64("★CRF20260210_error_female.xlsx")

    # 3. HTMLテンプレートの作成
    html_template = f"""<!DOCTYPE html>
<html>
  <head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no" />
    <title>LSRegistry CRF Tool</title>
    <link
      rel="stylesheet"
      href="https://cdn.jsdelivr.net/npm/@stlite/mountable@0.58.0/build/stlite.css"
    />
    <style>
      /* 初期ローディング画面のスタイル */
      #loading-screen {{
        position: fixed;
        top: 0;
        left: 0;
        width: 100vw;
        height: 100vh;
        background: linear-gradient(135deg, #0e1117 0%, #161a24 100%);
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        z-index: 9999;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
        color: white;
      }}
      .spinner {{
        width: 50px;
        height: 50px;
        border: 5px solid rgba(255, 255, 255, 0.1);
        border-radius: 50%;
        border-top-color: #00f2fe;
        animation: spin 1s ease-in-out infinite;
      }}
      @keyframes spin {{
        to {{ transform: rotate(360deg); }}
      }}
      .loading-title {{
        font-size: 24px;
        font-weight: 700;
        margin-top: 20px;
        background: linear-gradient(90deg, #00f2fe 0%, #4facfe 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
      }}
      .loading-desc {{
        font-size: 14px;
        color: #8b9bb4;
        margin-top: 10px;
        max-width: 400px;
        text-align: center;
        line-height: 1.5;
      }}
    </style>
  </head>
  <body>
    <!-- Streamlitのマウント先 -->
    <div id="root">
      <!-- ローディング画面 (マウント時に自動的に上書き消去されます) -->
      <div id="loading-screen">
        <div class="spinner"></div>
        <div class="loading-title">LSRegistry CRF Tool</div>
        <div id="loading-desc" class="loading-desc">
          ブラウザ上でPython環境を起動しています...<br>
          (初回起動時はライブラリのダウンロードに10〜30秒程度かかります)
        </div>
        <!-- エラー詳細表示エリア -->
        <div id="error-display" style="display:none; margin-top:20px; padding:15px; background:rgba(255, 75, 75, 0.15); border:1px solid #ff4b2b; border-radius:8px; max-width:600px; color:#ff4b2b; font-family:monospace; font-size:12px; white-space:pre-wrap; text-align:left;"></div>
      </div>
    </div>

    <script>
      // グローバルJSエラーハンドラ
      window.onerror = function(message, source, lineno, colno, error) {{
        const errDiv = document.getElementById("error-display");
        const descDiv = document.getElementById("loading-desc");
        if (errDiv) {{
          errDiv.style.display = "block";
          errDiv.innerText = "JavaScriptエラーが発生しました:\\n" + message + "\\n\\nSource: " + source + ":" + lineno + ":" + colno + (error ? "\\nStack: " + error.stack : "");
        }}
        if (descDiv) {{
          descDiv.innerHTML = "<span style='color:#ff4b2b; font-weight:bold;'>起動エラー</span>";
        }}
        return false;
      }};
    </script>
    <script src="https://cdn.jsdelivr.net/npm/@stlite/mountable@0.58.0/build/stlite.js" crossorigin></script>
    <script>
      // Base64からUint8Arrayに変換するヘルパー
      function base64ToUint8Array(base64) {{
        var binaryString = atob(base64);
        var bytes = new Uint8Array(binaryString.length);
        for (var i = 0; i < binaryString.length; i++) {{
          bytes[i] = binaryString.charCodeAt(i);
        }}
        return bytes;
      }}

      // サンプルExcelデータ
      const normalXlsx = base64ToUint8Array("{normal_xlsx_b64}");
      const errorMaleXlsx = base64ToUint8Array("{error_male_xlsx_b64}");
      const errorFemaleXlsx = base64ToUint8Array("{error_female_xlsx_b64}");

      stlite.mount({{
        requirements: ["pandas", "plotly", "openpyxl"],
        entrypoint: "app.py",
        files: {{
          "app.py": {app_py_json},
          "parser.py": {parser_py_json},
          "validator.py": {validator_py_json},
          "sample_normal.xlsx": normalXlsx,
          "sample_error_male.xlsx": errorMaleXlsx,
          "sample_error_female.xlsx": errorFemaleXlsx
        }}
      }}, document.getElementById("root")).catch(function(err) {{
        const errDiv = document.getElementById("error-display");
        const descDiv = document.getElementById("loading-desc");
        if (errDiv) {{
          errDiv.style.display = "block";
          errDiv.innerText = "stliteマウントエラー:\\n" + err.message + "\\n\\nStack: " + err.stack;
        }}
        if (descDiv) {{
          descDiv.innerHTML = "<span style='color:#ff4b2b; font-weight:bold;'>起動エラー</span>";
        }}
      }});
    </script>
  </body>
</html>
"""

    # 4. index.htmlの書き出し
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html_template)
    
    print("Successfully built index.html for GitHub Pages (stlite)")

if __name__ == "__main__":
    build_index_html()
