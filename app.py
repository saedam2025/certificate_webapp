from flask import Flask, render_template, request, redirect, url_for, flash, session
import pandas as pd
from jinja2 import Template
import smtplib, os

from weasyprint import HTML

from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from datetime import datetime, timedelta

app = Flask(__name__)
app.secret_key = "saedam-super-secret"

DATA_PATH = "pending_submissions.xlsx"
TEMPLATE_PATH = "certificate_template.html"
OUTPUT_DIR = "output_pdfs"
SEAL_IMAGE = "seal.gif"

EMAIL_ADDRESS = "lunch9797@gmail.com"
APP_PASSWORD = "txnb ofpi jgys jpfq"

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

if not os.path.exists(DATA_PATH):
    pd.DataFrame(columns=[
    "신청일", "증명서종류", "성명", "주민번호", "자택주소", "근무시작일", "근무종료일",
    "근무장소", "강의과목", "이메일주소", "상태", "발급일", "발급번호"
    ]).to_excel(DATA_PATH, index=False)


def format_korean_date(date_str):
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    return dt.strftime("%Y년 %#m월 %#d일")  # 윈도우용

def excel_date_to_str(value):
    if isinstance(value, (int, float)):
        base_date = datetime(1899, 12, 30)
        date = base_date + timedelta(days=int(value))
    elif isinstance(value, datetime):
        date = value
    else:
        return str(value)
    return date.strftime("%Y년 %m월 %d일")

def get_issue_number():
    year = datetime.today().year
    counter_file = f"counter_{year}.txt"
    if not os.path.exists(counter_file):
        with open(counter_file, "w") as f:
            f.write("1")
        return 1
    with open(counter_file, "r") as f:
        num = int(f.read().strip())
    with open(counter_file, "w") as f:
        f.write(str(num + 1))
    return num

def send_email(to_email, name, pdf_path):
    msg = MIMEMultipart()
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = to_email
    msg["Subject"] = f"[경력증명서] {name} 강사님 문서입니다"
    msg.attach(MIMEText(f"{name} 강사님, 안녕하세요.\n요청하신 경력증명서를 첨부드립니다.\n\n- 새담청소년교육문화원", "plain"))
    with open(pdf_path, "rb") as f:
        part = MIMEApplication(f.read(), _subtype="pdf")
        part.add_header("Content-Disposition", "attachment", filename=os.path.basename(pdf_path))
        msg.attach(part)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(EMAIL_ADDRESS, APP_PASSWORD)
        server.send_message(msg)

from weasyprint import HTML  # 상단에 추가

def generate_pdf(row, 발급번호):
    with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
        template = Template(f.read())

    시작일 = format_korean_date(row["근무시작일"])
    종료일 = "현재까지" if row["근무종료일"] == "현재까지" else format_korean_date(row["근무종료일"])

    html = template.render(
        증명서종류 = row["증명서종류"],
        성명=row["성명"],
        주민번호=row["주민번호"],
        주소=row["자택주소"],
        과목=row["강의과목"],
        장소=row["근무장소"],
        시작=시작일,
        종료=종료일,
        발급일자=datetime.today().strftime("%Y년 %m월 %d일"),
        발급번호=발급번호
    )

    output_path = os.path.join(OUTPUT_DIR, f"{row['성명']}_경력증명서.pdf")
    HTML(string=html, base_url='.').write_pdf(output_path)  # ✅ 여기만 바뀜
    return output_path


@app.route("/")
def index():
    return render_template("form.html")

@app.route("/submit", methods=["POST"])
def submit():
    df = pd.read_excel(DATA_PATH)
    form_data = dict(request.form)

    # ✅ 근무종료일 처리
    if form_data.get("종료일선택") == "현재까지":
        form_data["근무종료일"] = "현재까지"
    else:
        form_data["근무종료일"] = form_data.get("근무종료일", "")

    form_data["신청일"] = datetime.today().strftime("%Y-%m-%d %H:%M")
    form_data["상태"] = "대기"
    form_data["발급일"] = ""
    form_data["증명서종류"] = form_data.get("증명서종류", "")

    # ✅ 저장 전에 불필요한 select 항목 제거
    if "종료일선택" in form_data:
        del form_data["종료일선택"]

    df.loc[len(df)] = form_data
    df.to_excel(DATA_PATH, index=False)
    return render_template("success.html")

@app.route("/admin")
@app.route("/admin/<int:page>")
def admin(page=1):
    df = pd.read_excel(DATA_PATH)
    df = df.iloc[::-1].reset_index(drop=True)

    per_page = 10
    total_pages = (len(df) + per_page - 1) // per_page
    total_count = len(df)
    page_data = df.iloc[(page - 1) * per_page : page * per_page].copy()
    page_data = page_data.fillna("")
    page_data = page_data.to_dict(orient="records")

    return render_template("admin.html",
        submissions=page_data,
        page=page,
        total_pages=total_pages,
        total_count=total_count
    )

@app.route("/generate/<int:idx>")
def generate(idx):
    df = pd.read_excel(DATA_PATH)
    df = df.iloc[::-1].reset_index(drop=True)
    if idx >= len(df):
        return "잘못된 요청입니다.", 400

    row = df.iloc[idx]
    발급번호 = f"제{datetime.today().strftime('%y')}-{get_issue_number()}호"
    pdf = generate_pdf(row, 발급번호)
    send_email(row["이메일주소"], row["성명"], pdf)

    original_df = pd.read_excel(DATA_PATH)
    match = (
        (original_df["성명"] == row["성명"]) &
        (original_df["이메일주소"] == row["이메일주소"]) &
        (original_df["신청일"] == row["신청일"])
    )
    original_df.loc[match, "상태"] = "발급완료"
    original_df.loc[match, "발급일"] = datetime.today().strftime("%Y-%m-%d %H:%M")
    original_df.loc[match, "발급번호"] = 발급번호

    original_df.to_excel(DATA_PATH, index=False)
    return redirect(url_for("admin"))

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0")
