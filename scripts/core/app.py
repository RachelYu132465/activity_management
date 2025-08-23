from flask import Flask, render_template
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[2]  # 專案根目錄
app = Flask(__name__, template_folder=BASE_DIR / "templates")

app = Flask(__name__)

@app.route("/")
def index():
    data = {"title": "測試頁面", "message": "Hello, world!"}
    return render_template("template.html", **data)

if __name__ == "__main__":
    app.run(debug=True)  # debug=True 啟用自動重新載入
