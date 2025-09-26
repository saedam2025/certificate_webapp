# ===== Imports (deduped) =====
# 표준 라이브러리
import os
import re
import shutil
import uuid
import threading
import smtplib
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

# 서드파티
import pandas as pd
import pdfkit
from jinja2 import Template
from flask import (
    Flask, request, jsonify, render_template, render_template_string,
    redirect, url_for, send_from_directory, flash, session
)
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
# ===== end =====

# 렌더서버는 미국서버이므로 한국시간으로 변경-------
def now_kst():
    return datetime.now(ZoneInfo("Asia/Seoul"))
#--------

#---지급날짜 선택기능을 위한 코드--------------
def resolve_send_date(form) -> datetime.date:
    """
    <input name="send_date" type="date"> (yyyy-mm-dd)을 받음.
    비어있거나 잘못되면 KST 오늘 날짜로 처리.
    """
    raw = (form.get('send_date') or '').strip()
    if raw:
        try:
            return datetime.strptime(raw, '%Y-%m-%d').date()
        except ValueError:
            pass
    return now_kst().date()
#--------------------------------------------

"""
Unified Flask app for Render
- Part A: Payroll slip email sender (original app.py)
- Part B: Instructor certificate system (original appf.py)

Notes:
- Consolidated into a single Flask app instance.
- Removed duplicate/conflicting routes and function names.
- Switched credentials to environment variables with safe fallbacks.
- Verified paths for Render (/mnt/data) and static/templates usage.
"""

# =============================
# Common App Setup
# =============================
app = Flask(__name__, template_folder=".")
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "saedam-super-secret")

# Render's ephemeral disk safe base dir
BASE_DIR = "/mnt/data" if os.path.exists("/mnt/data") else "."

# 지속 저장 디렉터리(급여명세서 광고 이미지 ad01.jpg등등 저장용)
AD_DIR = os.path.join(BASE_DIR, "ad_images")
os.makedirs(AD_DIR, exist_ok=True)


# =============================
# Email Credentials (ENV first)
# =============================

EMAIL_ADDRESS_01 = os.environ.get("EMAIL_ADDRESS_01")
APP_PASSWORD_01  = os.environ.get("APP_PASSWORD_01")
EMAIL_ADDRESS_02 = os.environ.get("EMAIL_ADDRESS_02")
APP_PASSWORD_02  = os.environ.get("APP_PASSWORD_02")

def _system_email_login_params(system: str):
    # system은 "system01" 또는 "system02"
    if str(system).endswith("01"):
        return (EMAIL_ADDRESS_01 or os.environ.get("EMAIL_ADDRESS"),
                APP_PASSWORD_01  or os.environ.get("APP_PASSWORD"))
    else:  # system02 기본
        return (EMAIL_ADDRESS_02 or os.environ.get("EMAIL_ADDRESS"),
                APP_PASSWORD_02  or os.environ.get("APP_PASSWORD"))


# ===== 저장용: 구글메일주소와 앱비밀번호는 렌더서버 환경셋팅에 셋팅함. 아래를 써도 됨.
#EMAIL_ADDRESS = os.environ.get("EMAIL_ADDRESS") or 'lunch9797@gmail.com'
#APP_PASSWORD = os.environ.get("APP_PASSWORD") or 'txnb ofpi jgys jpfq'
#EMAIL_ADDRESS = os.environ.get("EMAIL_ADDRESS") or 'saedam2025@gmail.com'
#APP_PASSWORD = os.environ.get("APP_PASSWORD") or 'wjuy bedx stdm szdt'
#EMAIL_ADDRESS = os.environ.get("EMAIL_ADDRESS") or 'daompluse@gmail.com'
#APP_PASSWORD = os.environ.get("APP_PASSWORD") or 'qore sqyq ogvb edwo'
# ========================================================




# =========================================================
# Part A — PAYROLL SENDER (from original app.py, two separate operators: send01, send02)
# =========================================================

# ===== Part A config =====
UPLOAD_FOLDER_BASE = os.path.join(BASE_DIR, "uploads")

SENDER_KEYS = ("send01", "send02")

def _get_env(key, fallback=None):
    return os.environ.get(key) or fallback

SENDER_CONF = {
    "send01": {
        "upload_dir": os.path.join(UPLOAD_FOLDER_BASE, "send01"),  # <-- 문자열 경로
        "template_base": "send01",                                  # templates/send01/
        "email": _get_env("EMAIL_ADDRESS_01", _get_env("EMAIL_ADDRESS")),
        "app_pw": _get_env("APP_PASSWORD_01",  _get_env("APP_PASSWORD")),
    },
    "send02": {
        "upload_dir": os.path.join(UPLOAD_FOLDER_BASE, "send02"),  # <-- 문자열 경로
        "template_base": "send02",                                  # templates/send02/
        "email": _get_env("EMAIL_ADDRESS_02", _get_env("EMAIL_ADDRESS")),
        "app_pw": _get_env("APP_PASSWORD_02",  _get_env("APP_PASSWORD")),
    },
}

# 디렉터리 생성
for key in SENDER_KEYS:
    os.makedirs(SENDER_CONF[key]["upload_dir"], exist_ok=True)
# ===== end =====


# ---- operator-scoped runtime states ----
runtime = {
    "send01": {
        "sent_count": 0,
        "sent_names": [],
        "sent_count_lock": threading.Lock(),
        "stop_requested": False,
        "stop_lock": threading.Lock(),
    },
    "send02": {
        "sent_count": 0,
        "sent_names": [],
        "sent_count_lock": threading.Lock(),
        "stop_requested": False,
        "stop_lock": threading.Lock(),
    },
}

# 이미지 캐시: 공용 + send01 + send02 모두 로드====
image_cache = {}

def load_images():
    """
    우선순위: AD_DIR/(send01|send02)/파일 -> AD_DIR/공용파일 -> static/(send..)/파일 -> static/공용파일
    """
    variants = ["", "send01", "send02"]  # ""=공용
    files = ["logo01.jpg", "ad1.jpg", "ad2.jpg", "ad3.jpg"]

    def read_bytes(rel):
        # 1) persistent: 버킷별
        p1 = os.path.join(AD_DIR, rel)
        if os.path.exists(p1):
            with open(p1, "rb") as f:
                return f.read()

        # 2) persistent: 공용
        fname = rel.split("/", 1)[-1]
        p2 = os.path.join(AD_DIR, fname)
        if os.path.exists(p2):
            with open(p2, "rb") as f:
                return f.read()

        # 3) static: 버킷별
        p3 = os.path.join("static", rel)
        if os.path.exists(p3):
            with open(p3, "rb") as f:
                return f.read()

        # 4) static: 공용
        p4 = os.path.join("static", fname)
        if os.path.exists(p4):
            with open(p4, "rb") as f:
                return f.read()

        return None

    for v in variants:
        for fname in files:
            rel = f"{v}/{fname}" if v else fname
            data = read_bytes(rel)
            if data:
                image_cache[rel] = data

# 앱 시작 시 로고 및 광고이미지 1회 로드===
load_images()
#============================

def render_email_template(template_base, template_name, context):
    # templates/<template_base>/<template_name>
    with open(os.path.join('templates', template_base, template_name), 'r', encoding='utf-8') as f:
        template_str = f.read()
    return render_template_string(template_str, **context)

def _email_login_params(sender_key):
    return SENDER_CONF[sender_key]["email"], SENDER_CONF[sender_key]["app_pw"]

# ---- Upload UI (per-operator) ----
@app.route('/send01', methods=['GET', 'POST'])
@app.route('/send02', methods=['GET', 'POST'])
def payroll_upload_file_multi():
    sender_key = request.path.strip('/')

    if request.method == 'POST':
        # reset stop flag
        with runtime[sender_key]["stop_lock"]:
            runtime[sender_key]["stop_requested"] = False

        chosen_date = min(resolve_send_date(request.form), now_kst().date())
        runtime[sender_key]["send_date"] = chosen_date
        runtime[sender_key]["send_date_str"] = chosen_date.strftime('%Y년 %m월 %d일')
        runtime[sender_key]["send_date_iso"] = chosen_date.strftime('%Y-%m-%d')

        file = request.files.get('excel')
        if file and file.filename.lower().endswith('.xlsx'):
            safe_filename = f"{uuid.uuid4()}.xlsx"
            save_dir = SENDER_CONF[sender_key]["upload_dir"]
            path = os.path.join(save_dir, safe_filename)
            file.save(path)
            try:
                result_html = process_excel_multi(sender_key, path)
            except Exception as e:
                return f"처리 중 오류 발생: {e}"
            finally:
                try:
                    os.remove(path)
                except Exception:
                    pass
            return (result_html or "") + f'</div><br><a href="/{sender_key}" style="display:block; margin:16px auto 24px; width:fit-content; padding:8px 16px; background:#1f3c88; color:#fff; text-decoration:none; border-radius:5px;">발송 페이지로 가기</a>'
        else:
            return "엑셀 파일(.xlsx)만 업로드 가능합니다."

    # 각 담당자 폴더의 업로드 폼 사용: templates/send01/upload_form.html, templates/send02/upload_form.html
    upload_form_path = os.path.join("templates", SENDER_CONF[sender_key]["template_base"], "upload_form.html")
    return render_template_string(
        open(upload_form_path, encoding="utf-8").read(),
        uuid1=str(uuid.uuid4()), uuid2=str(uuid.uuid4()), uuid3=str(uuid.uuid4())
    )

# ---- Stop & Status (per-operator) ----
@app.post('/send01/stop')
@app.post('/send02/stop')
def stop_sending_multi():
    sender_key = request.path.split('/')[1]
    with runtime[sender_key]["stop_lock"]:
        runtime[sender_key]["stop_requested"] = True
    return f'''
    <script>
        alert("({sender_key}) 발송이 중단되었습니다.");
        location.href = "/{sender_key}";
    </script>
    '''

@app.get('/send01/status')
@app.get('/send02/status')
def status_multi():
    sender_key = request.path.split('/')[1]
    return jsonify({
        "sent_count": runtime[sender_key]["sent_count"],
        "sent_names": list(reversed(runtime[sender_key]["sent_names"]))
    })

# ---- 광고 이미지 교체 (담당자별 분리 + 공용 폴백) ----
@app.post('/send/upload_ad_image')
@app.post('/send01/upload_ad_image')
@app.post('/send02/upload_ad_image')
def upload_ad_image_multi():
    file   = request.files.get('ad_file')
    target = request.form.get('target')              # 'ad1.jpg' | 'ad2.jpg' | 'ad3.jpg' | 'logo01.jpg'
    bucket = request.form.get('bucket', '')          # '', 'send01', 'send02'

    valid = {'logo01.jpg', 'ad1.jpg', 'ad2.jpg', 'ad3.jpg'}
    if not file or target not in valid:
        return "잘못된 요청입니다.", 400

    # ✅ 지속 저장소(/mnt/data/ad_images ...)에 저장
    folder = os.path.join(AD_DIR, bucket) if bucket in ('send01', 'send02') else AD_DIR
    os.makedirs(folder, exist_ok=True)
    file.save(os.path.join(folder, target))

    # 캐시 갱신
    load_images()

    # 캐시 무력화 리다이렉트
    return '''
    <script>
      alert("이미지 교체 완료");
      const url = new URL(document.referrer || '/', window.location.origin);
      url.searchParams.set('cb', Date.now().toString());
      window.location.replace(url.toString());
    </script>
    '''

# ---- 광고 이미지 교체 mnt/data/ad_images로  ----
from flask import send_file, abort

def _safe_join(base, rel):
    rel = rel.replace("\\", "/")
    if ".." in rel:
        return None
    return os.path.join(base, rel)

@app.get("/ad/<path:rel>")
def serve_ad(rel):
    # 1) persistent 우선
    p = _safe_join(AD_DIR, rel)
    if p and os.path.exists(p):
        return send_file(p)
    # 2) static 폴백
    p2 = _safe_join("static", rel)
    if p2 and os.path.exists(p2):
        return send_file(p2)
    abort(404)
# ---- 광고 이미지 교체 mnt/data/ad_images로 끝  ----


# ---- Core processor (per-operator) ----
def process_excel_multi(sender_key, filepath):
    # init runtime
    runtime[sender_key]["sent_count"] = 0
    runtime[sender_key]["sent_names"] = []

    summary_by_sheet = {}

    # header row at index 2 (3rd Excel row)
    excel_data = pd.read_excel(filepath, sheet_name=None, header=2)

    def format_account_number(account_number):
        s = str(account_number or "").strip()
        digits = ''.join(ch for ch in s if ch.isdigit())
        if not digits:
            return ""
        # 4자리 그룹 하이픈
        return '-'.join([digits[i:i+4] for i in range(0, len(digits), 4)])

    def process_row(row, template_name, sheet_summary):
        # stop check
        with runtime[sender_key]["stop_lock"]:
            if runtime[sender_key]["stop_requested"]:
                return

        EMAIL_ADDRESS, APP_PASSWORD = _email_login_params(sender_key)

        try:
            name_raw = row.get('강사명') or row.get('직원명')
            name = str(name_raw).strip() if pd.notna(name_raw) else ''
            receiver_raw = row.get('이메일')
            receiver = str(receiver_raw).strip() if pd.notna(receiver_raw) else ''

            has_name = bool(name) and name.lower() not in ('nan', 'none', 'non')
            has_email = bool(receiver) and receiver.lower() not in ('nan', 'none', 'non')

            if not has_name and not has_email:
                return

            if not (has_name and has_email):
                display_name = name if has_name else '이름 없음'
                display_email = receiver if has_email else '이메일 없음'
                msg = f"<span style='color:red;'>{display_name} - 이메일: {display_email}</span>"
                with runtime[sender_key]["sent_count_lock"]:
                    runtime[sender_key]["sent_names"].append(msg)
                    sheet_summary.append(msg)
                return

            job = str(row.get('학교명', '')).strip()
            subject = str(row.get('과목', '')).strip()
            bank = str(row.get('은행', '')).strip()
            account_src = str(row.get('계좌번호', '')).strip()
            account = format_account_number(account_src)

            # 업로드 폼에서 선택한 날짜(없으면 오늘)
            send_date_display = runtime[sender_key].get("send_date_str") or now_kst().strftime('%Y년 %m월 %d일')
            send_date_iso = runtime[sender_key].get("send_date_iso") or now_kst().strftime('%Y-%m-%d')


            def safe_amount(_row, key):
                try:
                    val = _row.get(key)
                    if val is None or (isinstance(val, float) and pd.isna(val)):
                        return "0"
                    s = str(val).replace(',', '').strip()
                    return f"{int(float(s)):,}"
                except:
                    return "0"

            def safe_text(_row, key):
                try:
                    val = _row.get(key)
                    if val is None or (isinstance(val, float) and pd.isna(val)):
                        return ''
                    return str(val).strip()
                except:
                    return ''

            def to_int(v):
                try:
                    return int(float(str(v).replace(',', '').strip()))
                except:
                    return 0

            real_amount = to_int(row.get('지급총액')) - to_int(row.get('공제총액'))
            income_tax_total = to_int(row.get('근로소득세')) + to_int(row.get('지방소득세'))

            remark_val = row.get('강사전달비고') or row.get('직원전달비고') or row.get('전달비고') or ''
            remark = (str(remark_val).strip() if pd.notna(remark_val) else '') or '&nbsp;'

            context = {
                'name': name,
                'job': job,
                'subject': subject,
                'bank': bank,
                'account': account,
                'remark': remark,
       
                'today': send_date_display,        # 날짜 선택용 코드
                'send_date': send_date_display,      # 표시용(한글 형식)
                'send_date_iso': send_date_iso,      # YYYY-MM-DD

                'real_amount': real_amount,
                'row': row,
                'safe_amount': safe_amount,
                'income_tax_total': income_tax_total,
                'safe_text': safe_text
            }

            template_base = SENDER_CONF[sender_key]["template_base"]

            with app.app_context():
                html = render_email_template(template_base, template_name, context)

                from email.mime.multipart import MIMEMultipart
                from email.mime.text import MIMEText
                from email.mime.image import MIMEImage

                msg = MIMEMultipart('related')
                msg['Subject'] = f'[새담 지급명세서] {name}님 - {send_date_display}'
                msg['From'] = EMAIL_ADDRESS
                msg['To'] = receiver

                html_part = MIMEMultipart('alternative')
                html_part.attach(MIMEText(html, 'html'))
                msg.attach(html_part)

                # teacher vs others ad rule

                # 담당자별 이미지 첨부 + 공용 폴백
                base = sender_key  # 'send01' 또는 'send02'
                image_list = [
                    ('logo_image', f'{base}/logo01.jpg'),
                    ('ad1_image',   f'{base}/ad1.jpg'),
                    ('ad2_image',   f'{base}/ad2.jpg' if template_name == 'teacher.html' else f'{base}/ad3.jpg'),
                ]
                for cid, rel in image_list:
                    # 1순위: 담당자 폴더 이미지, 2순위: 공용('logo01.jpg' 등)으로 폴백
                    data = image_cache.get(rel)
                    if data is None:
                        fallback = rel.split('/', 1)[-1]  # 'logo01.jpg' 등
                        data = image_cache.get(fallback)
                    if data:
                        mime_img = MIMEImage(data, _subtype='jpeg')
                        mime_img.add_header('Content-ID', f'<{cid}>')
                        msg.attach(mime_img)

                with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                    smtp.login(EMAIL_ADDRESS, APP_PASSWORD)
                    smtp.send_message(msg)

                with runtime[sender_key]["sent_count_lock"]:
                    runtime[sender_key]["sent_names"].append(f"{job} - {name}")
                    sheet_summary.append(f"{job} - {name}")
                    runtime[sender_key]["sent_count"] += 1

        except Exception as e:
            print(f"❌ [{sender_key}] {row.get('강사명', row.get('직원명', '이름없음'))} 실패: {e}")

    # template rules
    template_rules = [
        (['강사', '선택형', '맞춤형'], 'teacher.html'),
        (['직원근로자','임직원/센터장'], 'employee_worker.html'),
        (['직원사업자','센터장/안전코디'], 'employee_business.html'),
        (['퇴직자'], 'retired.html'),
    ]
    DEFAULT_TEMPLATE = 'teacher.html'

    def pick_template(payroll_type_raw: str) -> str:
        s = (payroll_type_raw or '').strip()
        s_lower = str(s).lower()
        for keywords, tpl in template_rules:
            for kw in keywords:
                if kw.lower() in s_lower:
                    return tpl
        return DEFAULT_TEMPLATE

    # --- 순차 발송 (병렬 제거 및 너무빠른 메일발송을 막기위한 딜레이 초 설정) ------------------------------------------------------
    import time
    SEND_DELAY_SEC = float(os.environ.get("SEND_DELAY_SEC", "0.5"))  # 0이면 지연 없음

    for sheet_name, df in excel_data.items():
        df.columns = df.columns.str.strip()
        sheet_summary = []

        # try to infer template from first raw row
        try:
            raw_df = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
            first_row = raw_df.iloc[0].astype(str).str.strip().tolist()
            all_keywords = [kw for kws, _ in template_rules for kw in kws]
            payroll_type = next(
                (v for v in first_row if any(kw.lower() in v.lower() for kw in all_keywords)),
                ''
            )
        except Exception:
            payroll_type = ''

        template_name = pick_template(payroll_type)

        # ✅ 한 명씩 순차 처리
        for _, row in df.iterrows():
            # 중단 요청 체크
            with runtime[sender_key]["stop_lock"]:
                if runtime[sender_key]["stop_requested"]:
                    break

            process_row(row, template_name, sheet_summary)

            # Gmail 블록 방지용 속도 제한
            if SEND_DELAY_SEC > 0:
                try:
                    time.sleep(SEND_DELAY_SEC)
                except Exception:
                    pass

        summary_by_sheet[sheet_name] = sheet_summary

    # === 결과 HTML 생성 ===
    result_html = f"""
    <html>
    <head>
      <meta charset='UTF-8'>
      <link href="https://fonts.googleapis.com/css2?family=Nanum+Gothic&display=swap" rel="stylesheet">
      <style>
        :root {{
          --bg: #f7f8fb;
          --card: #ffffff;
          --border: #e5e7eb;
          --text: #1f2937;
          --muted: #6b7280;
          --brand: #1f3c88;
          --shadow: 0 8px 24px rgba(15, 23, 42, 0.08);
        }}
        * {{ box-sizing: border-box; }}
        body {{
          margin: 0; padding: 48px 32px;
          background: var(--bg);
          font-family: 'Nanum Gothic', sans-serif; color: var(--text);
        }}
        .page {{ max-width: 1440px; margin: 0 auto; }}
        .header {{ display:flex; align-items:center; gap:12px; margin-bottom:18px; }}
        .title {{ font-size: 22px; font-weight: 800; color: var(--brand); }}
        .badge {{
          display:inline-block; background:#eef2ff; color:#3730a3;
          border:1px solid #c7d2fe; padding:4px 10px; border-radius:999px;
          font-size:13px; font-weight:700;
        }}
        .card {{
          background: var(--card); border: 1px solid var(--border);
          border-radius: 12px; box-shadow: var(--shadow);
          padding: 18px;
        }}
        .sheet {{
          border: 1px solid var(--border);
          border-radius: 10px; padding: 16px; margin: 14px 0; background:#fff;
        }}
        .sheet-title {{ font-size:16px; font-weight:700; margin-bottom:12px; }}
        .names {{
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(220px, 1fr));
          gap: 6px 18px;
          align-items: start; justify-items: start;
        }}
        .name-item {{
          font-size:14px; line-height:1.5; white-space: nowrap;
          overflow: hidden; text-overflow: ellipsis;
        }}
        .actions {{ margin-top:22px; text-align:center; }}
        .btn {{
          display:inline-block; background: var(--brand); color:#fff;
          padding:10px 16px; border-radius:8px; text-decoration:none;
          border:1px solid #1f3c88;
        }}
        .btn:hover {{ opacity:.92; }}
      </style>
    </head>
    <body>
      <div class="page">
        <div class="header">
          <div class="title">[지급명세서] 메일 발송 결과</div>
          <span class="badge">총 {runtime[sender_key]["sent_count"]}명</span>
        </div>

        <div class="card">
    """
    for sheet, names in summary_by_sheet.items():
        result_html += f"""
          <div class="sheet">
            <div class="sheet-title">시트명: {sheet} (총 {len(names)}명)</div>
            <div class="names">
        """
        for entry in names:
            result_html += f"<div class='name-item'>• {entry}</div>"
        result_html += "</div></div>"

    result_html += f"""
        </div>
      </div>
      <br><img src="https://www.saedam.org/img_sub/logo.gif" alt="로고" style="display:block; margin:0 auto; height:48px; width:auto;"><br><br>
    </body>
    </html>
    """
    return result_html


# =============================
# Part B — CERTIFICATE SYSTEM (from original appf.py)
# =============================

# wkhtmltopdf configuration (Render compatible)
WKHTMLTOPDF_PATH = shutil.which("wkhtmltopdf") or "/usr/bin/wkhtmltopdf"
config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)

# PDF output folders
pdf_folder1 = os.path.join(BASE_DIR, "output_pdfs01")
pdf_folder2 = os.path.join(BASE_DIR, "output_pdfs02")
os.makedirs(pdf_folder1, exist_ok=True)
os.makedirs(pdf_folder2, exist_ok=True)

# System passwords
USER_PASSWORDS = {
    "system01": os.environ.get("USER_PW_SYS01", "0070"),
    "system02": os.environ.get("USER_PW_SYS02", "0070"),
}

ADMIN_PASSWORDS = {
    "system01": os.environ.get("ADMIN_PW_SYS01", "1900"),
    "system02": os.environ.get("ADMIN_PW_SYS02", "8016"),
}

# Admin notification targets
ADMIN_EMAILS = {
    "system01": os.environ.get("ADMIN_EMAIL_SYS01", "lunch97@naver.com"),
    "system02": os.environ.get("ADMIN_EMAIL_SYS02", "comedu74@nate.com"),
}

SEAL_IMAGE = "seal.gif"

# Issue number helpers

def get_year_prefix():
    return now_kst().strftime('%y')

def get_next_issue_number():
    year_prefix = get_year_prefix()
    file_name = os.path.join(BASE_DIR, f"last_number_{year_prefix}.txt")

    if not os.path.exists(file_name):
        last = 0
    else:
        with open(file_name, 'r') as f:
            try:
                last = int(f.read().strip())
            except ValueError:
                last = 0

    next_number = last + 1
    with open(file_name, 'w') as f:
        f.write(str(next_number))

    return f"제{year_prefix}-{next_number:04d}호"


def send_admin_notification(system, name, cert_type):
    to_email = ADMIN_EMAILS.get(system)
    if not to_email:
        print(f"❌ 시스템에 맞는 이메일 없음: {system}")
        return

    from_addr, from_pw = _system_email_login_params(system)

    msg = MIMEText(
        f"새담 홈페이지를 통해 새로운 강사 경력증명발급 신청이 접수되었습니다.\n\n시스템: {system}\n\n신청자: {name}\n\n증명서 종류: {cert_type}"
    )
    msg['Subject'] = f'[{system.upper()}] 새담 강사경력증명서 신청 알림 (신청자: {name})'
    msg['From'] = from_addr
    msg['To'] = to_email

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(from_addr, from_pw)
            smtp.send_message(msg)
            print(f"✅ 신청 알림 메일 전송됨: {to_email}")
    except Exception as e:
        print(f"❌ 메일 전송 실패: {e}")


def ensure_data_file(data_path):
    if not os.path.exists(data_path):
        pd.DataFrame(columns=[
            "신청일", "증명서종류", "성명", "주민번호", "자택주소",
            "근무시작일", "근무종료일", "근무장소", "강의과목", "용도", "직책",
            "이메일주소", "상태", "발급일", "발급번호", "종료사유"
        ]).to_excel(data_path, index=False)


def format_korean_date(date_str):
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    # On Linux, %-m/%-d would avoid leading zeros; Windows used %#m. Keep generic.
    return dt.strftime("%Y년 %m월 %d일")


def send_certificate_email(system, to_email, name, pdf_path, certificate_type):
    from_addr, from_pw = _system_email_login_params(system)

    msg = MIMEMultipart()
    msg["From"] = from_addr
    msg["To"] = to_email
    msg["Subject"] = f"[{certificate_type}] {name} 강사님 문서입니다"
    body = f"{name} 강사님, 안녕하세요.\n\n요청하신 {certificate_type}를 첨부드립니다.\n\n(사)새담청소년교육문화원"
    msg.attach(MIMEText(body, "plain"))
    with open(pdf_path, "rb") as f:
        part = MIMEApplication(f.read(), _subtype="pdf")
        part.add_header("Content-Disposition", "attachment", filename=os.path.basename(pdf_path))
        msg.attach(part)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(from_addr, from_pw)
        server.send_message(msg)


def generate_pdf(row, issue_no, system):
    template_path = "certificate_template.html"
    with open(template_path, "r", encoding="utf-8") as f:
        template = Template(f.read())

    def fmt(date_val):
        return "현재까지" if date_val == "현재까지" else format_korean_date(date_val)

    resident_raw = row["주민번호"]
    if "-" in resident_raw:
        앞, 뒤 = resident_raw.split("-")
        masked_resident = 앞 + "-" + 뒤[0] + "******"
    else:
        masked_resident = resident_raw

    html = template.render(
        증명서종류=row.get("증명서종류", ""),
        성명=row["성명"],
        주민번호=masked_resident,
        주소=row["자택주소"],
        과목=row["강의과목"],
        용도=row.get("용도", ""),
        직책=row.get("직책", ""),
        장소=row["근무장소"],
        시작=fmt(row["근무시작일"]),
        종료=fmt(row["근무종료일"]),
        종료사유=row.get("종료사유", ""),
        발급일자=now_kst().strftime("%Y년 %m월 %d일"),
        발급번호=issue_no
    )

    # Seal absolute path for wkhtmltopdf
    seal_path = os.path.abspath(SEAL_IMAGE)
    html = html.replace('src="seal.gif"', f'src="file:///{seal_path}"')

    output_dir = os.path.join(BASE_DIR, f"output_pdfs{system[-2:]}")
    os.makedirs(output_dir, exist_ok=True)
    cert_type = row.get("증명서종류", "증명서").replace(" ", "")
    output_path = os.path.join(output_dir, f"{issue_no}_{row['성명']}_{cert_type}.pdf")
    options = {'enable-local-file-access': ''}
    pdfkit.from_string(html, output_path, configuration=config, options=options)
    return output_path


# ---- Convenience redirects for system roots ----
@app.route("/system01/")
def redirect_system01():
    return redirect(url_for("form_login", system="system01"))

@app.route("/system02/")
def redirect_system02():
    return redirect(url_for("form_login", system="system02"))


# ---- CRUD & Workflows ----
@app.route('/<system>/update/<int:idx>', methods=['POST'])
def update_submission(system, idx):
    data_path = os.path.join(BASE_DIR, f"pending_submissions_{system[-2:]}.xlsx")
    page = int(request.form.get("page", 1))
    df = pd.read_excel(data_path)
    df = df.iloc[::-1].reset_index(drop=True)
    form_data = dict(request.form)

    # apply edits
    for key in form_data:
        df.at[idx, key] = form_data[key]

    # save in original order
    original_df = pd.read_excel(data_path)
    original_index = len(original_df) - 1 - idx
    for key in form_data:
        original_df.at[original_index, key] = form_data[key]
    original_df.to_excel(data_path, index=False)
    flash('수정이 완료되었습니다')
    return redirect(url_for('admin', system=system, page=page))


@app.route('/<system>/delete/<int:idx>')
def delete_submission_simple(system, idx):
    """Delete row AND corresponding PDF if exists (merged behavior)."""
    data_path = os.path.join(BASE_DIR, f"pending_submissions_{system[-2:]}.xlsx")
    page = int(request.args.get("page", 1))
    df = pd.read_excel(data_path)
    df = df.iloc[::-1].reset_index(drop=True)

    # remove PDF if present
    row = df.iloc[idx]
    issue_no = str(row.get("발급번호", "")).strip()
    name = str(row.get("성명", "")).strip()
    cert_type = str(row.get("증명서종류", "증명서")).replace(" ", "")
    pdf_dir = os.path.join(BASE_DIR, f"output_pdfs{system[-2:]}")
    pdf_filename = f"{issue_no}_{name}_{cert_type}.pdf"
    pdf_path = os.path.join(pdf_dir, pdf_filename)
    if os.path.exists(pdf_path):
        os.remove(pdf_path)

    # drop row and save back in original order
    df = df.drop(index=idx).reset_index(drop=True)
    final_df = df.iloc[::-1].reset_index(drop=True)
    final_df.to_excel(data_path, index=False)
    return redirect(url_for('admin', system=system, page=page))


@app.route('/<system>/submit', methods=['POST'])
def submit(system):
    data_path = os.path.join(BASE_DIR, f"pending_submissions_{system[-2:]}.xlsx")
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

    # keep user session authenticated
    session[f'user_authenticated_{system}'] = True

    # notify admins
    send_admin_notification(system, row_data["성명"], row_data["증명서종류"])

    return render_template(f"{system}/success.html", system=system, **row_data)


# ---- Auth gates ----
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

    if not session.get(f"{system}_authenticated"):
        return render_template(f"{system}/admin_login.html", system=system)

    data_path = os.path.join(BASE_DIR, f"pending_submissions_{system[-2:]}.xlsx")
    ensure_data_file(data_path)
    df = pd.read_excel(data_path)
    df = df.iloc[::-1].reset_index(drop=True)

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
        df=df,
        total_count=total_count,
        issued_count=issued_count,
        pending_count=pending_count,
        total_pages=total_pages,
        page=page,
        system=system
    )


@app.route("/<system>/bulk_delete", methods=["POST"])
def bulk_delete(system):
    ids_str = request.form.get("selected_ids", "")
    page = request.form.get("page", "1")

    if not ids_str:
        flash("삭제할 항목이 선택되지 않았습니다.")
        return redirect(url_for('admin', system=system, page=page))

    selected_indices = [int(i) for i in ids_str.split(',') if i.isdigit()]
    data_path = os.path.join(BASE_DIR, f"pending_submissions_{system[-2:]}.xlsx")
    pdf_folder = os.path.join(BASE_DIR, f"output_pdfs{system[-2:]}")

    original_df = pd.read_excel(data_path)
    total_len = len(original_df)

    # map visible indices to original order
    original_indices = [total_len - 1 - i for i in selected_indices]

    for idx in sorted(original_indices, reverse=True):
        row = original_df.iloc[idx]
        pdf_filename = f"{row['발급번호']}_{row['성명']}_{row['증명서종류'].replace(' ', '')}.pdf"
        pdf_path = os.path.join(pdf_folder, pdf_filename)
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
            print(f"✅ 삭제됨: {pdf_filename}")
        else:
            print(f"❌ PDF 없음: {pdf_filename}")
        original_df.drop(index=idx, inplace=True)

    original_df.reset_index(drop=True, inplace=True)
    original_df.to_excel(data_path, index=False)

    flash(f"{len(selected_indices)}건이 삭제되었습니다.")
    return redirect(url_for('admin', system=system, page=page))


@app.route("/<system>/logout")
def logout(system):
    session.pop(f"{system}_authenticated", None)
    return redirect(url_for("admin", system=system))


@app.route('/<system>/pdf/<filename>')
def download_pdf(system, filename):
    pdf_dir = os.path.join(BASE_DIR, f"output_pdfs{system[-2:]}")
    return send_from_directory(pdf_dir, filename)


@app.route("/<system>/generate/<int:idx>")
def generate(system, idx):
    data_path = os.path.join(BASE_DIR, f"pending_submissions_{system[-2:]}.xlsx")
    page = int(request.args.get("page", 1))
    ensure_data_file(data_path)
    df = pd.read_excel(data_path)
    df = df.iloc[::-1].reset_index(drop=True)
    row = df.iloc[idx]

    issue_no = get_next_issue_number()
    pdf_path = generate_pdf(row, issue_no, system)
    send_certificate_email(system, row["이메일주소"], row["성명"], pdf_path, row["증명서종류"])

    original_df = pd.read_excel(data_path)
    original_index = len(original_df) - 1 - idx
    original_df.at[original_index, "상태"] = "발급완료"
    original_df.at[original_index, "발급일"] = now_kst().strftime("%Y-%m-%d")
    original_df.at[original_index, "발급번호"] = issue_no
    original_df.to_excel(data_path, index=False)

    return redirect(url_for("admin", system=system, page=page))


# ===== (입금용 엑셀 생성기) Blueprint — 기존 코드 수정 없이 추가 =====

from flask import Blueprint, request, send_file, render_template_string, abort, current_app
import io, os, re, zipfile
import pandas as pd
from datetime import datetime

trweb_bp = Blueprint("trweb", __name__)
DEPOSIT_CHUNK_SIZE = int(os.environ.get("CHUNK_SIZE", "200"))  # 시트 누적 인원 기준 (기본 200)

def deposit_clean_account(acc):
    if pd.isna(acc):
        return ""
    s = str(acc).strip()
    if re.fullmatch(r"\d+(\.\d+)?[eE][+-]?\d+", s):
        try:
            from decimal import Decimal
            s = format(Decimal(s), "f").rstrip("0").rstrip(".")
        except Exception:
            pass
    if s.endswith(".0"):
        s = s[:-2]
    s = s.replace(" ", "")
    # s = s.replace("-", "")  # 필요 시 하이픈 제거
    return s

DEPOSIT_INDEX_HTML = """
<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <title>입금용 엑셀 생성기</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <style>
    body {font-family: system-ui, -apple-system, Segoe UI, Roboto, 'Noto Sans KR', sans-serif; background:#f7f7fb; margin:0; padding:40px;}
    .wrap {max-width:720px; margin:0 auto; background:#fff; border:1px solid #e5e7eb; border-radius:16px; padding:28px;}
    h1 {margin:0 0 8px; font-size:22px;}
    p.desc {margin:0 0 20px; color:#666; line-height:1.6;}
    .box {border:1px dashed #cbd5e1; border-radius:12px; padding:16px; background:#fafafa; margin-bottom:16px;}
    label {display:block; margin:10px 0 6px; font-weight:600;}
    input[type=file] {display:block; width:100%; padding:10px; border:1px solid #e5e7eb; border-radius:10px; background:#fff;}
    .btn {display:inline-block; margin-top:16px; background:#2563eb; color:#fff; border:none; border-radius:10px; padding:12px 18px; font-weight:700; cursor:pointer;}
    .btn:disabled {opacity:.6; cursor:not-allowed;}
    .hint {font-size:12px; color:#777; margin-top:4px;}
    .footer {margin-top:18px; font-size:12px; color:#888;}
    ul {margin:8px 0 0 18px; color:#555;}
    code {background:#f1f5f9; padding:2px 6px; border-radius:6px;}
  </style>
</head>
<body>
  <div class="wrap">
    <h1>입금용 엑셀 생성기</h1>
    <p class="desc">
      모든 시트에서 <code>은행</code>, <code>계좌번호</code>, <code>예금주</code>, <code>입금액</code>만 추출합니다. (헤더는 <b>3번째 줄</b>)<br/>
      <b>분할 규칙:</b> 시트 순서대로 누적 인원을 합산하다가 <b>{{ chunk_size }}</b>명을 초과하면 <u>직전 시트까지만</u> 한 파일로,
      나머지는 다음 파일로 저장합니다. (시트는 절대 쪼개지지 않음)
    </p>

    <div class="box">
      <form method="post" action="{{ url_for('trweb.deposit_process') }}" enctype="multipart/form-data">
        <label for="file">엑셀 파일 업로드 (.xlsx)</label>
        <input id="file" name="file" type="file" accept=".xlsx" required />
        <div class="hint">예: sss.xlsx</div>
        <button class="btn" type="submit">변환 & 다운로드</button>
      </form>
    </div>

    <div>
      <strong>처리 규칙</strong>
      <ul>
        <li>각 시트에서 <code>은행/계좌번호/예금주/입금액</code>이 모두 있으면 추출</li>
        <li><code>예금주</code>가 있는 행만 사용</li>
        <li><code>계좌번호</code>는 <b>문자</b>로 강제(선행 0 보존), 지수표기/공백/".0" 정리</li>
        <li>각 결과 파일에서 시트 구간 마지막 행 아래에 <b>시트별 합계(H/I)</b></li>
        <li>전체 총합은 <b>H2/I2</b>, 모든 데이터 행의 <b>F열</b>에 "새담청소년교육"</li>
      </ul>
    </div>

    <div class="footer">문제가 있으면 헤더 위치(3행), 열 이름 철자를 확인하세요.</div>
  </div>
  <script>
    const form = document.querySelector("form");
    const file = document.getElementById("file");
    form.addEventListener("submit", () => {
      if (!file.files.length) return;
      const btn = form.querySelector(".btn");
      btn.textContent = "처리 중...";
      btn.disabled = true;
    });
  </script>
</body>
</html>
"""

@trweb_bp.get("/")
def deposit_index():
    return render_template_string(DEPOSIT_INDEX_HTML, chunk_size=DEPOSIT_CHUNK_SIZE)

@trweb_bp.post("/process")
def deposit_process():
    up = request.files.get("file")
    if not up:
        abort(400, "파일이 없습니다.")
    if not up.filename.lower().endswith(".xlsx"):
        abort(400, "xlsx 형식만 지원합니다.")

    # 기존 코드의 BASE_DIR을 사용 (같은 파일 내 전역이라 접근 가능)
    upload_root = os.path.join(BASE_DIR, "uploads_deposit")
    os.makedirs(upload_root, exist_ok=True)

    input_path = os.path.join(upload_root, f"_upload_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    up.save(input_path)

    try:
        sheets = deposit_read_sheets_as_list(input_path)
        if not sheets:
            abort(400, "추출할 데이터가 없습니다. (열 이름/헤더 3행 확인)")

        parts = deposit_split_by_sheet_boundary(sheets, DEPOSIT_CHUNK_SIZE)
        today = datetime.today().strftime("%Y-%m-%d")

        # 파트 1개면 단일 파일
        if len(parts) == 1:
            merged = pd.concat([df for _, df in parts[0]], ignore_index=True)
            xlsx_bytes = deposit_build_excel_bytes(merged, file_label=today)
            return send_file(
                io.BytesIO(xlsx_bytes),
                as_attachment=True,
                download_name=f"입금내역_{today}.xlsx",
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # 여러 파트면 ZIP
        buff = io.BytesIO()
        with zipfile.ZipFile(buff, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for i, part in enumerate(parts, start=1):
                merged = pd.concat([df for _, df in part], ignore_index=True)
                part_no = f"{i:02d}"
                fname = f"입금내역_{today}_part{part_no}.xlsx"
                xlsx_bytes = deposit_build_excel_bytes(merged, file_label=f"{today} part {part_no}")
                zf.writestr(fname, xlsx_bytes)

        buff.seek(0)
        return send_file(buff, as_attachment=True, download_name=f"입금내역_{today}_split.zip", mimetype="application/zip")

    finally:
        try:
            os.remove(input_path)
        except Exception:
            pass

def deposit_read_sheets_as_list(input_path: str):
    """[(sheet_name, df_with__시트), ...]"""
    target = ["은행", "계좌번호", "예금주", "입금액"]
    out = []
    xlsx = pd.ExcelFile(input_path)
    for sheet_name in xlsx.sheet_names:
        df = pd.read_excel(xlsx, sheet_name=sheet_name, header=2, dtype={"계좌번호": str})
        if all(col in df.columns for col in target):
            filtered = df[target].dropna(subset=["예금주"]).copy()
            if filtered.empty:
                continue
            filtered["계좌번호"] = filtered["계좌번호"].apply(deposit_clean_account)
            filtered["입금액"] = pd.to_numeric(filtered["입금액"], errors="coerce")
            filtered["_시트"] = sheet_name
            out.append((sheet_name, filtered))
    return out

def deposit_split_by_sheet_boundary(sheet_dfs, chunk_size: int):
    """시트 경계 기준 누적 인원 chunk_size 초과 시 파트 분할 (시트는 절대 쪼개지지 않음)."""
    parts, current, count = [], [], 0
    for sheet_name, df in sheet_dfs:
        rows = len(df)
        if rows == 0:
            current.append((sheet_name, df))
            continue
        if count > 0 and (count + rows) > chunk_size:
            parts.append(current)
            current, count = [], 0
        current.append((sheet_name, df))
        count += rows
        if rows > chunk_size:
            parts.append(current)
            current, count = [], 0
    if current:
        parts.append(current)
    return parts

def deposit_build_excel_bytes(df_with_sheet: pd.DataFrame, file_label: str) -> bytes:
    """'_시트'는 계산용(출력 제외). 시트 구간 마지막 행 아래 H/I에 합계, H2/I2에 총합, F열 '새담청소년교육'."""
    visible_cols = [c for c in df_with_sheet.columns if c != "_시트"]
    result_df = df_with_sheet[visible_cols].copy()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        result_df.to_excel(writer, index=False)
        workbook = writer.book
        ws = writer.sheets["Sheet1"]

        for idx, col in enumerate(result_df.columns):
            if col == "계좌번호":
                ws.set_column(idx, idx, 25, workbook.add_format({"num_format": "@"}))
            elif col == "입금액":
                ws.set_column(idx, idx, 15, workbook.add_format({"num_format": "#,##0"}))
            else:
                ws.set_column(idx, idx, 20)

        # F열 "새담청소년교육" (헤더 제외)
        tf = workbook.add_format({"num_format": "@"})
        for r in range(1, len(result_df) + 1):
            ws.write(r, 5, "새담청소년교육", tf)
        ws.set_column(5, 5, 20, tf)

        # 계좌번호 문자열 강제
        if "계좌번호" in result_df.columns:
            acc_idx = list(result_df.columns).index("계좌번호")
            for r, v in enumerate(result_df["계좌번호"].fillna(""), start=1):
                ws.write_string(r, acc_idx, str(v))

        # 시트별 합계 (H/I)
        amt_fmt = workbook.add_format({"num_format": "#,##0"})
        if "_시트" in df_with_sheet.columns and "입금액" in df_with_sheet.columns and len(df_with_sheet) > 0:
            ss = df_with_sheet["_시트"].astype(str).tolist()
            blocks = [0]
            for i in range(1, len(ss)):
                if ss[i] != ss[i-1]:
                    blocks.append(i)
            blocks.append(len(ss))
            for b in range(len(blocks)-1):
                start, end = blocks[b], blocks[b+1]
                sheet_name = ss[start]
                target_row = end  # 헤더 다음부터 데이터 1행이므로 end가 맞음
                sub_total = pd.to_numeric(df_with_sheet.iloc[start:end]["입금액"], errors="coerce").sum(skipna=True)
                ws.write(target_row, 7, f"{sheet_name} 합계:")
                ws.write_number(target_row, 8, sub_total if pd.notna(sub_total) else 0, amt_fmt)

        grand_total = pd.to_numeric(result_df.get("입금액", []), errors="coerce").sum(skipna=True) if len(result_df) else 0
        ws.write(1, 7, "총입금액:")
        ws.write_number(1, 8, grand_total if pd.notna(grand_total) else 0, amt_fmt)

    output.seek(0)
    return output.getvalue()

# === 블루프린트 등록 (모든 라우트 정의 끝난 뒤, Entry Point 위 한 줄) ===
app.register_blueprint(trweb_bp, url_prefix="/trweb")



# =============================
# Entry Point
# =============================
if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0")
