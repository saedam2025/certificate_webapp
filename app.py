from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash, session
import pandas as pd
import pdfkit
from jinja2 import Template
import smtplib, os
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo # ✅ 미국 서버를 한국 시간으로 조정
from flask import render_template_string
import shutil

# ✅ 한국 시간 반환 함수
def now_kst():
    return datetime.now(ZoneInfo("Asia/Seoul"))

# ✅ 발급번호 생성============================
def get_year_prefix():
    return now_kst().strftime('%y')

def get_next_issue_number():
    year_prefix = get_year_prefix()
    file_name = os.path.join("/mnt/data", f"last_number_{year_prefix}.txt")

    # 파일이 없으면 0부터 시작
    if not os.path.exists(file_name):
        last = 0
    else:
        with open(file_name, 'r') as f:
            try:
                last = int(f.read().strip())
            except ValueError:
                last = 0  # 혹시 파일 내용이 비어있거나 이상할 경우 대비

    next_number = last + 1

    # 새로운 번호 저장
    with open(file_name, 'w') as f:
        f.write(str(next_number))

    return f"제{year_prefix}-{next_number:04d}호"
# ✅ 발급번호 생성============================



WKHTMLTOPDF_PATH = shutil.which("wkhtmltopdf") or "/usr/bin/wkhtmltopdf"
config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)

app = Flask(__name__, template_folder=".")
app.secret_key = "saedam-super-secret"

# pdf저장디렉토리
base_dir = "/mnt/data" if os.path.exists("/mnt/data") else "."

pdf_folder1 = os.path.join(base_dir, "output_pdfs01")
pdf_folder2 = os.path.join(base_dir, "output_pdfs02")

os.makedirs(pdf_folder1, exist_ok=True)
os.makedirs(pdf_folder2, exist_ok=True)

# 시스템별 비밀번호
USER_PASSWORDS = {
    "system01": "0070",
    "system02": "0070"
}

ADMIN_PASSWORDS = {
    "system01": "1900",
    "system02": "8016"
}

 # system02 담당자 이메일
ADMIN_EMAILS = {
    "system01": "lunch97@naver.com",
    "system02": "windows7@hanmail.net" 
}

SEAL_IMAGE = "seal.gif"
EMAIL_ADDRESS = os.environ.get("EMAIL_ADDRESS")
APP_PASSWORD = os.environ.get("APP_PASSWORD")

 # 신청오면 메일보내주기 시작----------
def send_admin_notification(system, name, cert_type):
    to_email = ADMIN_EMAILS.get(system)
    if not to_email:
        print(f"❌ 시스템에 맞는 이메일 없음: {system}")
        return

    msg = MIMEText(f"새담 홈페이지를 통해 새로운 강사 경력증명발급 신청이 접수되었습니다.\n\n시스템: {system}\n\n신청자: {name}\n\n증명서 종류: {cert_type}")
    msg['Subject'] = f'[{system.upper()}] 새담 강사경력증명서 신청 알림 (신청자: {name})'
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = to_email

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(EMAIL_ADDRESS, APP_PASSWORD)
            smtp.send_message(msg)
            print(f"✅ 신청 알림 메일 전송됨: {to_email}")
    except Exception as e:
        print(f"❌ 메일 전송 실패: {e}")
 # 신청오면 메일보내주기 끝----------


def ensure_data_file(data_path):
    if not os.path.exists(data_path):
        pd.DataFrame(columns=[
            "신청일", "증명서종류", "성명", "주민번호", "자택주소",
            "근무시작일", "근무종료일", "근무장소", "강의과목", "용도", "직책",
            "이메일주소", "상태", "발급일", "발급번호", "종료사유"
        ]).to_excel(data_path, index=False)

def format_korean_date(date_str):
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    return dt.strftime("%Y년 %#m월 %#d일")

def send_email(to_email, name, pdf_path, certificate_type):
    msg = MIMEMultipart()
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = to_email
    msg["Subject"] = f"[{certificate_type}] {name} 강사님 문서입니다"
    body = f"{name} 강사님, 안녕하세요.\n\n요청하신 {certificate_type}를 첨부드립니다.\n\n(사)새담청소년교육문화원"
    msg.attach(MIMEText(body, "plain"))
    with open(pdf_path, "rb") as f:
        part = MIMEApplication(f.read(), _subtype="pdf")
        part.add_header("Content-Disposition", "attachment", filename=os.path.basename(pdf_path))
        msg.attach(part)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(EMAIL_ADDRESS, APP_PASSWORD)
        server.send_message(msg)

def generate_pdf(row, 발급번호, system):
    template_path = "certificate_template.html"
    with open(template_path, "r", encoding="utf-8") as f:
        template = Template(f.read())
    시작일 = format_korean_date(row["근무시작일"])
    종료일 = "현재까지" if row["근무종료일"] == "현재까지" else format_korean_date(row["근무종료일"])
    주민번호_원본 = row["주민번호"]
    if "-" in 주민번호_원본:
        앞, 뒤 = 주민번호_원본.split("-")
        마스킹주민번호 = 앞 + "-" + 뒤[0] + "******"
    else:
        마스킹주민번호 = 주민번호_원본
    html = template.render(
        증명서종류=row.get("증명서종류", ""),
        성명=row["성명"],
        주민번호=마스킹주민번호,
        주소=row["자택주소"],
        과목=row["강의과목"],
        용도=row.get("용도", ""),
        직책=row.get("직책", ""),
        장소=row["근무장소"],
        시작=시작일,
        종료=종료일,
        종료사유=row.get("종료사유", ""),
        발급일자=now_kst().strftime("%Y년 %m월 %d일"),
        발급번호=발급번호
    )
    seal_path = os.path.abspath(SEAL_IMAGE)
    html = html.replace('src="seal.gif"', f'src="file:///{seal_path}"')
    output_dir = os.path.join("/mnt/data", f"output_pdfs{system[-2:]}")
    os.makedirs(output_dir, exist_ok=True)
    cert_type = row.get("증명서종류", "증명서").replace(" ", "")
    output_path = os.path.join(output_dir, f"{발급번호}_{row['성명']}_{cert_type}.pdf")
    options = {'enable-local-file-access': ''}
    pdfkit.from_string(html, output_path, configuration=config, options=options)
    return output_path

@app.route("/system01/")
def redirect_system01():
    return redirect(url_for("form_login", system="system01"))

@app.route("/system02/")
def redirect_system02():
    return redirect(url_for("form_login", system="system02"))

@app.route('/<system>/update/<int:idx>', methods=['POST'])
def update(system, idx):
    data_path = os.path.join(base_dir, f"pending_submissions_{system[-2:]}.xlsx")
    page = int(request.form.get("page", 1))  # 🔹 page 값 받기
    df = pd.read_excel(data_path)
    df = df.iloc[::-1].reset_index(drop=True)
    form_data = dict(request.form)

    # 수정 반영
    for key in form_data:
        df.at[idx, key] = form_data[key]

    # 역순 저장
    original_df = pd.read_excel(data_path)
    original_index = len(original_df) - 1 - idx
    for key in form_data:
        original_df.at[original_index, key] = form_data[key]
    original_df.to_excel(data_path, index=False)
    flash('수정이 완료되었습니다')  # ✅ 메시지 추가
    return redirect(url_for('admin', system=system, page=page))  # 🔹 해당 페이지로 이동


@app.route('/<system>/delete/<int:idx>')
def delete(system, idx):
    data_path = os.path.join(base_dir, f"pending_submissions_{system[-2:]}.xlsx")
    page = int(request.args.get("page", 1))  # 🔹 쿼리스트링에서 page 받기
    df = pd.read_excel(data_path)
    df = df.iloc[::-1].reset_index(drop=True)
    df = df.drop(index=idx).reset_index(drop=True)
    # 역순 저장
    final_df = df.iloc[::-1].reset_index(drop=True)
    final_df.to_excel(data_path, index=False)
    return redirect(url_for('admin', system=system, page=page))  # 🔹 해당 페이지로 이동


@app.route('/<system>/submit', methods=['POST'])
def submit(system):
    data_path = os.path.join(base_dir, f"pending_submissions_{system[-2:]}.xlsx")
    ensure_data_file(data_path)
    df = pd.read_excel(data_path)

    form_data = dict(request.form)
    form_data["근무종료일"] = "현재까지" if form_data.get("종료일선택") == "현재까지" else form_data.get("근무종료일", "")
    form_data["신청일"] = now_kst().strftime("%Y-%m-%d")
    form_data["상태"] = "대기"
    form_data["발급일"] = ""
    if "종료일선택" in form_data:
        del form_data["종료일선택"]
    종료사유 = form_data.get("종료사유") if form_data.get("증명서종류") == "강사 해촉증명서" else ""

    ordered_fields = [
        "신청일", "증명서종류", "성명", "주민번호", "자택주소",
        "근무시작일", "근무종료일", "근무장소", "강의과목", "용도", "직책",
        "이메일주소", "상태", "발급일", "발급번호", "종료사유"
    ]

    if "종료사유" not in df.columns:
        df["종료사유"] = ""

    row_data = {col: form_data.get(col, "") for col in ordered_fields}
    row_data["종료사유"] = 종료사유

    df.loc[len(df)] = row_data
    df.to_excel(data_path, index=False)

    # ✅ 인증 세션 유지
    session[f'user_authenticated_{system}'] = True

    # ✅ 알림 메일 전송
    send_admin_notification(system, row_data["성명"], row_data["증명서종류"])

    return render_template(f"{system}/success.html", system=system, **row_data)


    # 패스워드 걸기 시작===================

@app.route('/<system>/form', methods=['GET', 'POST'])
def form_login(system):
    if request.method == 'POST':
        pw = request.form.get('password')
        if pw == USER_PASSWORDS.get(system):
            session[f'user_authenticated_{system}'] = True
            return redirect(url_for('show_form', system=system))
        else:
            flash("비밀번호가 틀렸습니다.")
            return redirect(url_for('form_login', system=system))
    
    # ✅ 들여쓰기 수정됨
    return render_template(f"{system}/form_login.html", system=system, title="경력증명서 신청")


@app.route('/<system>/form_page', methods=['GET', 'POST'])
def show_form(system):
    if not session.get(f'user_authenticated_{system}'):
        flash("접근 권한이 없습니다.")
        return redirect(url_for('form_login', system=system))
    return render_template(f"{system}/form.html", system=system)



@app.route("/<system>/admin", defaults={'page': 1}, methods=["GET", "POST"])
@app.route("/<system>/admin/<int:page>", methods=["GET", "POST"])
def admin(system, page):
    if request.method == "POST":
        input_pw = request.form.get("password")
        correct_pw = ADMIN_PASSWORDS.get(system)
        if input_pw == correct_pw:
            session[f"{system}_authenticated"] = True
            return redirect(url_for("admin", system=system, page=page))
        else:
            flash("비밀번호가 틀렸습니다.")
            return render_template(f"{system}/admin_login.html", system=system)

    # GET 요청 시: 인증되어 있지 않으면 로그인 폼 보여줌
    if not session.get(f"{system}_authenticated"):
        return render_template(f"{system}/admin_login.html", system=system)

    # 패스워드 걸기 끝===================

    data_path = os.path.join(base_dir, f"pending_submissions_{system[-2:]}.xlsx")
    ensure_data_file(data_path)
    df = pd.read_excel(data_path)
    df = df.iloc[::-1].reset_index(drop=True)  # 최신순 정렬

    total_count = len(df)
    issued_count = len(df[df["상태"] == "발급완료"])
    pending_count = len(df[df["상태"] == "대기"])

    total_pages = (total_count - 1) // 10 + 1
    start = (page - 1) * 10
    end = start + 10
    submissions = df.iloc[start:end].fillna("").to_dict(orient="records")

    return render_template(
        f"{system}/admin.html",
        submissions=submissions,
        df=df,  # 
        total_count=total_count,
        issued_count=issued_count,
        pending_count=pending_count,
        total_pages=total_pages,
        page=page,
        system=system
    )


@app.route("/<system>/logout")
def logout(system):
    session.pop(f"{system}_authenticated", None)
    return redirect(url_for("admin", system=system))

@app.route('/<system>/pdf/<filename>')
def download_pdf(system, filename):
    pdf_dir = f"/mnt/data/output_pdfs{system[-2:]}"  # 예: output_pdfs01
    return send_from_directory(pdf_dir, filename)

# ✅ 발급번호 생성 방식 변경 적용된 generate 함수
@app.route("/<system>/generate/<int:idx>")
def generate(system, idx):
    data_path = os.path.join(base_dir, f"pending_submissions_{system[-2:]}.xlsx")
    page = int(request.args.get("page", 1))
    ensure_data_file(data_path)
    df = pd.read_excel(data_path)
    df = df.iloc[::-1].reset_index(drop=True)
    row = df.iloc[idx]

    # ✅ 발급번호 생성 방식 변경
    발급번호 = get_next_issue_number()

    pdf = generate_pdf(row, 발급번호, system)
    send_email(row["이메일주소"], row["성명"], pdf, row["증명서종류"])

    original_df = pd.read_excel(data_path)
    original_index = len(original_df) - 1 - idx
    original_df.at[original_index, "상태"] = "발급완료"
    original_df.at[original_index, "발급일"] = now_kst().strftime("%Y-%m-%d")
    original_df.at[original_index, "발급번호"] = 발급번호
    original_df.to_excel(data_path, index=False)

    return redirect(url_for("admin", system=system, page=page))

# ✅ 삭제 라우트: 엑셀 행 + PDF 함께 삭제
@app.route("/<system>/delete/<int:idx>", methods=["POST"])
def delete_submission(system, idx):
    data_path = os.path.join(base_dir, f"pending_submissions_{system[-2:]}.xlsx")
    ensure_data_file(data_path)
    df = pd.read_excel(data_path)
    df = df.iloc[::-1].reset_index(drop=True)
    row = df.iloc[idx]

    # PDF 파일 삭제 시도
    발급번호 = str(row.get("발급번호", "")).strip()
    성명 = str(row.get("성명", "")).strip()
    cert_type = str(row.get("증명서종류", "증명서")).replace(" ", "")
    pdf_folder = os.path.join(base_dir, f"output_pdfs{system[-2:]}")
    pdf_filename = f"{발급번호}_{성명}_{cert_type}.pdf"
    pdf_path = os.path.join(pdf_folder, pdf_filename)

    if os.path.exists(pdf_path):
        os.remove(pdf_path)

    # 행 삭제 후 역순 저장
    df = df.drop(index=idx).reset_index(drop=True)
    final_df = df.iloc[::-1].reset_index(drop=True)
    final_df.to_excel(data_path, index=False)

    flash("삭제가 완료되었습니다.")
    return redirect(url_for("admin", system=system))

# ✅ 이메일 수정창 부분=============================
@app.route('/<system>/update_email', methods=["POST"])
def update_email(system):
    index = int(request.form.get("index"))
    page = int(request.form.get("page", 1))
    new_email = request.form.get("이메일주소")

    file_path = f"pending_submissions_{system[-2:]}.xlsx"
    df = pd.read_excel(file_path)

    if 0 <= index < len(df):
        df.at[index, "이메일주소"] = new_email
        df.to_excel(file_path, index=False)
        flash("이메일이 수정되었습니다.")
    else:
        flash("⚠️ 유효하지 않은 인덱스입니다.")

    return redirect(url_for("admin", system=system, page=page))



# ✅  PDF 생성
@app.route("/<system>/pdf/<filename>")
def serve_pdf(system, filename):
    return send_from_directory(f"output_pdfs{system[-2:]}", filename)


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0")
