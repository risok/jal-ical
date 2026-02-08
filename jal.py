# app.py
import re
import uuid
import io
import zipfile
from dataclasses import dataclass
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import streamlit as st

JST = ZoneInfo("Asia/Tokyo")

MAX_INPUT_CHARS = 20000
MAX_LINES = 2000
MAX_FIELD_LEN = 200
MAX_UPLOAD_BYTES = 200000
CONTROL_CHAR_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]")

AIRPORT_CODE_MAP = {
    "東京(羽田)": "HND", "東京（羽田）": "HND", "羽田": "HND",
    "札幌(新千歳)": "CTS", "札幌（新千歳）": "CTS", "新千歳": "CTS",
    "大阪(伊丹)": "ITM", "大阪（伊丹）": "ITM", "伊丹": "ITM",
    "福岡": "FUK",
    "那覇": "OKA", "沖縄(那覇)": "OKA", "沖縄（那覇）": "OKA",
    "名古屋(中部)": "NGO", "名古屋（中部）": "NGO", "中部": "NGO"
}

def guess_airport_code(name: str) -> str:
    if name in AIRPORT_CODE_MAP:
        return AIRPORT_CODE_MAP[name]
    paren = re.search(r'[（(]([^（）()]+)[)）]', name)
    if paren:
        inside = paren.group(1)
        m = re.search(r'\b([A-Z]{3})\b', inside)
        if m:
            return m.group(1)
    m2 = re.search(r'\b([A-Z]{3})\b', name)
    if m2:
        return m2.group(1)
    letters = ''.join(ch for ch in name if ch.isalpha() and 'A' <= ch.upper() <= 'Z')
    if len(letters) >= 3:
        return letters[:3].upper()
    return name

@dataclass
class Flight:
    year: int
    month: int
    day: int
    flight_no: str
    dep_name: str
    dep_time: str
    arr_name: str
    arr_time: str
    seat_class: str | None = None
    seat_no: str | None = None

    @property
    def dep_dt(self):
        return datetime(self.year, self.month, self.day,
                        int(self.dep_time[:2]), int(self.dep_time[3:]), tzinfo=JST)

    @property
    def arr_dt(self):
        arr = datetime(self.year, self.month, self.day,
                       int(self.arr_time[:2]), int(self.arr_time[3:]), tzinfo=JST)
        if arr < self.dep_dt:
            arr += timedelta(days=1)
        return arr

    def dep_code(self):
        return guess_airport_code(self.dep_name)

    def arr_code(self):
        return guess_airport_code(self.arr_name)

def normalize(text: str) -> str:
    return text.replace("\u3000", " ").replace("\r", "").strip()

def sanitize_user_text(text: str) -> str:
    if not text:
        return ""
    text = CONTROL_CHAR_RE.sub("", text)
    if len(text) > MAX_INPUT_CHARS:
        text = text[:MAX_INPUT_CHARS]
    lines = text.split("\n")
    if len(lines) > MAX_LINES:
        lines = lines[:MAX_LINES]
    return "\n".join(lines)

def trim_field(value: str) -> str:
    value = CONTROL_CHAR_RE.sub("", value).strip()
    if len(value) > MAX_FIELD_LEN:
        value = value[:MAX_FIELD_LEN]
    return value

def escape_ics_text(value: str) -> str:
    value = trim_field(value)
    return (
        value.replace("\\", "\\\\")
             .replace("\n", "\\n")
             .replace(";", "\\;")
             .replace(",", "\\,")
    )

def extract_location_and_time(segment: str):
    m = re.match(r'(.+?)(\d{1,2}:\d{2})発', segment.strip())
    if not m:
        return None, None
    name = m.group(1).strip()
    time = m.group(2)
    return name, time

def is_valid_time(value: str) -> bool:
    if not re.match(r"^\d{1,2}:\d{2}$", value or ""):
        return False
    h, m = value.split(":")
    return 0 <= int(h) <= 23 and 0 <= int(m) <= 59

def parse_flights_email(raw: str):
    """メールフォーマットのパーサー"""
    text = normalize(raw)
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    flights = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if "JAL" in line and "便" in line and "年" in line and "月" in line and "日" in line:
            md = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日.*?JAL(\d+)便', line)
            if not md:
                i += 1
                continue
            year, month, day, flight_no = map(int, md.group(1,2,3,4))
            route_line = lines[i+1] if i + 1 < len(lines) else ""
            dep_name = dep_time = arr_name = arr_time = None
            if route_line and ("発" in route_line and "着" in route_line):
                parts = re.split(r'\s{2,}|\t+', route_line)
                parts = [p for p in parts if p]
                if len(parts) < 2:
                    parts = [route_line]
                if len(parts) >= 2:
                    dep_name, dep_time = extract_location_and_time(parts[0])
                    arr_m = re.match(r'(.+?)(\d{1,2}:\d{2})着', parts[1])
                    if arr_m:
                        arr_name = arr_m.group(1).strip()
                        arr_time = arr_m.group(2)
            seat_class = seat_no = None
            if i + 2 < len(lines) and lines[i+2].startswith("座席"):
                seat_line = lines[i+2]
                sc = re.search(r'クラス\s*([A-Zぁ-んァ-ヶ一-龠A-Za-z]+)', seat_line)
                if sc:
                    seat_class = sc.group(1).strip()
                sn = re.search(r'座席番号：?([0-9A-Z]{1,4})', seat_line)
                if sn:
                    seat_no = sn.group(1).strip()
            if all([dep_name, dep_time, arr_name, arr_time]) and is_valid_time(dep_time) and is_valid_time(arr_time):
                flights.append(Flight(
                    year, month, day, trim_field(str(flight_no)),
                    trim_field(dep_name), dep_time, trim_field(arr_name), arr_time,
                    trim_field(seat_class) if seat_class else None,
                    trim_field(seat_no) if seat_no else None
                ))
            i += 1
        else:
            i += 1
    return flights

def parse_flights_homepage(raw: str):
    """JALホームページフォーマットのパーサー"""
    text = normalize(raw)
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    flights = []
    i = 0
    
    while i < len(lines):
        line = lines[i]
        # 日付行を探す: 2026年2月10日（火）
        date_match = re.match(r'(\d{4})年(\d{1,2})月(\d{1,2})日', line)
        if date_match:
            year, month, day = map(int, date_match.groups())
            
            # 次の行から便情報を探す
            dep_time = dep_name = arr_time = arr_name = None
            flight_no = None
            seat_class = None
            found_dep = False
            
            j = i + 1
            while j < len(lines) and j < i + 15:  # 15行先まで探索
                curr_line = lines[j]
                
                # 出発時刻と空港: 11:55東京 (羽田)
                if not found_dep and re.match(r'^\d{1,2}:\d{2}', curr_line):
                    dep_match = re.match(r'^(\d{1,2}:\d{2})(.+)', curr_line)
                    if dep_match:
                        dep_time = dep_match.group(1)
                        dep_name = dep_match.group(2).strip()
                        found_dep = True
                        j += 1
                        continue
                
                # 到着時刻と空港: 14:50 沖縄 (那覇) または 14:50沖縄 (那覇)
                if found_dep and not arr_time and re.match(r'^\d{1,2}:\d{2}', curr_line):
                    arr_match = re.match(r'^(\d{1,2}:\d{2})\s*(.+)', curr_line)
                    if arr_match:
                        arr_time = arr_match.group(1)
                        arr_name = arr_match.group(2).strip()
                
                # クラス: クラス： クラス J
                if "クラス：" in curr_line or "クラス:" in curr_line:
                    class_match = re.search(r'クラス[：:]\s*(.+)', curr_line)
                    if class_match:
                        seat_class = class_match.group(1).strip()
                
                # 便名：JAL915
                if "便名：" in curr_line or "便名:" in curr_line:
                    flight_match = re.search(r'便名[：:]\s*JAL(\d+)', curr_line)
                    if flight_match:
                        flight_no = flight_match.group(1)
                
                # 次の日付行が来たら終了
                if j > i + 1 and re.match(r'\d{4}年\d{1,2}月\d{1,2}日', curr_line):
                    break
                
                j += 1
            
            # フライト情報が揃っていれば追加
            if all([dep_time, dep_name, arr_time, arr_name, flight_no]) and is_valid_time(dep_time) and is_valid_time(arr_time):
                flights.append(Flight(
                    year, month, day, trim_field(flight_no),
                    trim_field(dep_name), dep_time, trim_field(arr_name), arr_time,
                    trim_field(seat_class) if seat_class else None, None
                ))
                i = j - 1
            else:
                i += 1
        else:
            i += 1
    
    return flights

def parse_flights(raw: str):
    """両方のフォーマットを試してパース"""
    raw = sanitize_user_text(raw)
    # メールフォーマットを試す
    flights = parse_flights_email(raw)
    
    # 見つからなければホームページフォーマットを試す
    if not flights:
        flights = parse_flights_homepage(raw)
    
    return flights

def to_ics(flight: Flight) -> str:
    dep_dt = flight.dep_dt
    arr_dt = flight.arr_dt
    dtstamp = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    def fmt(dt): return dt.strftime("%Y%m%dT%H%M%S")
    dep_code = flight.dep_code()
    arr_code = flight.arr_code()
    def short(s): return s if len(s) <= 8 else s[:8]
    summary = escape_ics_text(f"JAL{flight.flight_no} {short(dep_code)}->{short(arr_code)}")
    location = escape_ics_text(f"{flight.dep_name} -> {flight.arr_name}")
    desc_parts = [
        f"Flight: JAL{flight.flight_no}",
        f"From: {flight.dep_name} ({dep_code}) {flight.dep_time}",
        f"To: {flight.arr_name} ({arr_code}) {flight.arr_time}",
    ]
    if flight.seat_class or flight.seat_no:
        desc_parts.append(f"Seat: {flight.seat_class or ''} {flight.seat_no or ''}".strip())
    description = escape_ics_text("\\n".join(desc_parts))
    uid = f"{uuid.uuid4()}@jal-parser"
    return (
        "BEGIN:VEVENT\n"
        f"UID:{uid}\n"
        f"DTSTAMP:{dtstamp}\n"
        f"DTSTART;TZID=Asia/Tokyo:{fmt(dep_dt)}\n"
        f"DTEND;TZID=Asia/Tokyo:{fmt(arr_dt)}\n"
        f"SUMMARY:{summary}\n"
        f"LOCATION:{location}\n"
        f"DESCRIPTION:{description}\n"
        "END:VEVENT\n"
    )

def flights_to_ics(flights):
    events = "".join(to_ics(f) for f in flights)
    return (
        "BEGIN:VCALENDAR\n"
        "PRODID:-//JAL Flight Parser//JP\n"
        "VERSION:2.0\n"
        "CALSCALE:GREGORIAN\n"
        "METHOD:PUBLISH\n"
        f"{events}"
        "END:VCALENDAR\n"
    )

SAMPLE = """旅程1
2025年9月20日（土）　JAL511便
東京(羽田)10:30発        札幌(新千歳)12:05着
座席：クラス J 座席番号：15H

旅程2
2025年9月23日（火）　JAL528便
札幌(新千歳)21:15発        東京(羽田)22:55着
座席：クラス J 座席番号：8D

旅程3
2025年10月1日（水）　JAL999便
架空空港A(テスト)07:00発        架空空港B(AAA)09:10着
"""

SAMPLE_HOMEPAGE = """予約番号：
FNMGMS
購入期限：
-
購入済み
(JALオンライン)
2026年2月10日（火）
運賃：
ビジネスフレックス
11:55東京 (羽田)
14:50 沖縄 (那覇)
クラス： クラス J
便名：JAL915
座席： 指定済み
2026年2月13日（金）
運賃：
ビジネスフレックス
18:20沖縄 (那覇)
20:30 東京 (羽田)
クラス： クラス J
便名：JAL916
座席： 指定済み
"""

# session_stateの初期化
if "text_input" not in st.session_state:
    st.session_state["text_input"] = ""

def set_text_input(value: str):
    st.session_state["text_input"] = sanitize_user_text(value or "")

def load_sample_email():
    set_text_input(SAMPLE)

def load_sample_hp():
    set_text_input(SAMPLE_HOMEPAGE)

def reset_text_input():
    set_text_input("")

def handle_upload():
    uploaded_file = st.session_state.get("uploader")
    if not uploaded_file:
        return
    if uploaded_file.size and uploaded_file.size > MAX_UPLOAD_BYTES:
        st.error("ファイルサイズが大きすぎます。")
        return
    text_input = uploaded_file.read().decode("utf-8", errors="ignore")
    set_text_input(text_input)

st.set_page_config(page_title="JAL フライト → ICS", page_icon="✈")

st.title("JAL フライト本文 → ICS 生成")
with st.expander("サンプル表示"):
    st.text("メールフォーマット:")
    st.code(SAMPLE, language="text")
    st.text("ホームページフォーマット:")
    st.code(SAMPLE_HOMEPAGE, language="text")

col_in, col_opts = st.columns([3,1])
with col_in:
    text_input = st.text_area(
        "フライト情報本文を貼り付け", 
        height=300, 
        placeholder="メール本文やホームページからコピーした内容を貼り付け...",
        key="text_input"
    )
    
with col_opts:
    st.button("サンプル読込(メール)", use_container_width=True, on_click=load_sample_email)
    st.button("サンプル読込(HP)", use_container_width=True, on_click=load_sample_hp)
    if st.session_state["text_input"]:
        st.button("リセット", use_container_width=True, type="secondary", on_click=reset_text_input)

st.file_uploader(
    "テキストファイル読込",
    type=["txt"],
    key="uploader",
    on_change=handle_upload
)

run = st.button("解析する")
if run:
    flights = parse_flights(st.session_state["text_input"] or "")
    if not flights:
        st.error("フライト情報を検出できませんでした。")
    else:
        st.success(f"検出フライト数: {len(flights)}")
        rows = []
        for f in flights:
            rows.append({
                "Date": f"{f.year:04d}-{f.month:02d}-{f.day:02d}",
                "Flight": f"JAL{f.flight_no}",
                "From": f"{f.dep_name} ({f.dep_code()})",
                "Dep": f.dep_time,
                "To": f"{f.arr_name} ({f.arr_code()})",
                "Arr": f.arr_time,
                "SeatClass": f.seat_class or "",
                "SeatNo": f.seat_no or ""
            })
        st.dataframe(rows, use_container_width=True)

        ics_all = flights_to_ics(flights)
        st.download_button(
            "まとめICSダウンロード",
            data=ics_all.encode("utf-8"),
            file_name="jal_flights_all.ics",
            mime="text/calendar"
        )

        # ZIP 個別
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in flights:
                fname = f"JAL{f.flight_no}_{f.year:04d}{f.month:02d}{f.day:02d}.ics"
                zf.writestr(fname, flights_to_ics([f]))
        st.download_button(
            "個別ICS ZIPダウンロード",
            data=buf.getvalue(),
            file_name="jal_flights.zip",
            mime="application/zip"
        )

        # 個別も表示
        with st.expander("生成ICS(まとめ)を表示"):
            st.code(ics_all, language="text")

st.caption("Paste JAL itinerary text and export to calendar.")