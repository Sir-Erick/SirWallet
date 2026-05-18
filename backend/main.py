import calendar
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from fastapi.responses import FileResponse
from datetime import datetime, timedelta
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from sqlalchemy import extract
from sqlalchemy.orm import Session
from database import SessionLocal, engine
from models import Base, Transaction, User

Base.metadata.create_all(bind=engine)

app = FastAPI()

pending_delete = {}

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


class Message(BaseModel):
    text: str
class RegisterRequest(BaseModel):
    username: str
    email: str
    password: str
class LoginRequest(BaseModel):
    email: str
    password: str

@app.post("/login")
def login(data: LoginRequest):

    db: Session = SessionLocal()

    user = db.query(User).filter(
        User.email == data.email
    ).first()

    if not user:
        return {
            "message": "Email tidak ditemukan."
        }

    if user.password != data.password:
        return {
            "message": "Password salah."
        }

    return {
        "message": "Login berhasil.",
        "user_id": user.id,
        "username": user.username
    }

@app.get("/")
def root():
    return {"message": "Finance AI Backend Running"}


# =========================
# DETECT CATEGORY
# =========================
def detect_category(text):

    if (
        "makan" in text
        or "jajan" in text
        or "kopi" in text
        or "ayam" in text
        or "nasgor" in text
    ):
        return "Makanan"

    elif (
        "bensin" in text
        or "transport" in text
        or "gojek" in text
    ):
        return "Transport"

    elif (
        "skin care" in text
        or "skincare" in text
        or "sabun" in text
        or "parfum" in text
    ):
        return "Perawatan"
    
    elif (
        "service" in text
        or "alat" in text
        or "listrik" in text
        or "air" in text
        or "peralatan" in text
    ):
        return "kebutuhan"

    elif (
        "game" in text
        or "steam" in text
        or "nongkrong" in text
        or "top up" in text
    ):
        return "Hiburan"

    else:
        return "Lainnya"


# =========================
# EXTRACT AMOUNT
# =========================
def extract_amount(text):

    text = text.lower().replace(".", "")

    match = re.search(
        r'(\d+(?:\.\d+)?)\s*(rb|ribu|jt|juta|k)?',
        text
    )

    if not match:
        return 0

    number = float(match.group(1))
    unit = match.group(2)

    if unit in ["rb", "ribu", "k"]:
        number *= 1000

    elif unit in ["jt", "juta"]:
        number *= 1000000

    return int(number)

# =========================
# EXTRACT MONTH & YEAR
# =========================
def extract_month_year(text):

    bulan_map = {
        "januari": 1,
        "februari": 2,
        "maret": 3,
        "april": 4,
        "mei": 5,
        "juni": 6,
        "juli": 7,
        "agustus": 8,
        "september": 9,
        "oktober": 10,
        "november": 11,
        "desember": 12
    }

    text = text.lower()

    selected_month = None
    selected_year = datetime.utcnow().year

    for nama_bulan, nomor_bulan in bulan_map.items():

        if nama_bulan in text:
            selected_month = nomor_bulan
            break

    tahun_match = re.search(r'(20\d{2})', text)

    if tahun_match:
        selected_year = int(tahun_match.group(1))

    return selected_month, selected_year

# ==================================
# REGISTER
# ==================================
@app.post("/register")
def register(data: RegisterRequest):

    db: Session = SessionLocal()

    # CEK USERNAME
    existing_username = db.query(User).filter(
        User.username == data.username
    ).first()

    if existing_username:

        return {
            "message": "Username sudah digunakan."
        }

    # CEK EMAIL
    existing_email = db.query(User).filter(
        User.email == data.email
    ).first()

    if existing_email:

        return {
            "message": "Email sudah digunakan."
        }

    # SIMPAN USER
    new_user = User(
        username=data.username,
        email=data.email,
        password=data.password
    )

    db.add(new_user)

    db.commit()

    return {
        "message": "Register berhasil."
    }

# ==================================
# HELPER: BUAT LAPORAN EXCEL 3 SHEET
# ==================================
def buat_laporan_excel(transactions, nama_bulan: str, tahun: int):

    wb = Workbook()

    # --- STYLE HELPERS ---
    def header_style(ws, row, col, value, bg_hex, font_color="FFFFFF"):
        cell = ws.cell(row=row, column=col, value=value)
        cell.font = Font(bold=True, color=font_color, size=11)
        cell.fill = PatternFill("solid", fgColor=bg_hex)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )
        return cell

    def data_style(ws, row, col, value, number_format=None):
        cell = ws.cell(row=row, column=col, value=value)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )
        if number_format:
            cell.number_format = number_format
        return cell

    def set_col_widths(ws, widths):
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    pemasukan_list = [t for t in transactions if t.type == "pemasukan"]
    pengeluaran_list = [t for t in transactions if t.type == "pengeluaran"]

    total_masuk = sum(t.amount for t in pemasukan_list)
    total_keluar = sum(t.amount for t in pengeluaran_list)
    saldo = total_masuk - total_keluar

    # =====================================
    # SHEET 1: RINGKASAN
    # =====================================
    ws1 = wb.active
    ws1.title = "Ringkasan"
    ws1.row_dimensions[1].height = 32
    ws1.row_dimensions[2].height = 20

    # Judul
    ws1.merge_cells("A1:C1")
    title_cell = ws1["A1"]
    title_cell.value = f"💼 SirWallet — Laporan {nama_bulan} {tahun}"
    title_cell.font = Font(bold=True, size=14, color="111827")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.fill = PatternFill("solid", fgColor="F3F4F6")

    # Header tabel ringkasan
    header_style(ws1, 3, 1, "Keterangan",   "111827")
    header_style(ws1, 3, 2, "Jumlah (Rp)",  "111827")
    header_style(ws1, 3, 3, "Transaksi",    "111827")

    rows_ringkasan = [
        ("💰 Total Pemasukan",  total_masuk,  len(pemasukan_list)),
        ("💸 Total Pengeluaran", total_keluar, len(pengeluaran_list)),
        ("💳 Saldo Akhir",       saldo,        len(transactions)),
    ]

    colors = ["D1FAE5", "FEE2E2", "DBEAFE"]
    font_colors = ["065F46", "991B1B", "1E3A5F"]

    for i, (ket, jml, jlh) in enumerate(rows_ringkasan, 4):
        ws1.row_dimensions[i].height = 22
        c1 = ws1.cell(row=i, column=1, value=ket)
        c1.font = Font(bold=True, color=font_colors[i-4])
        c1.fill = PatternFill("solid", fgColor=colors[i-4])
        c1.alignment = Alignment(vertical="center")
        c1.border = Border(left=Side(style="thin"), right=Side(style="thin"),
                           top=Side(style="thin"), bottom=Side(style="thin"))

        c2 = ws1.cell(row=i, column=2, value=jml)
        c2.font = Font(bold=True, color=font_colors[i-4])
        c2.fill = PatternFill("solid", fgColor=colors[i-4])
        c2.number_format = '#,##0'
        c2.alignment = Alignment(horizontal="right", vertical="center")
        c2.border = Border(left=Side(style="thin"), right=Side(style="thin"),
                           top=Side(style="thin"), bottom=Side(style="thin"))

        c3 = ws1.cell(row=i, column=3, value=jlh)
        c3.font = Font(color=font_colors[i-4])
        c3.fill = PatternFill("solid", fgColor=colors[i-4])
        c3.alignment = Alignment(horizontal="center", vertical="center")
        c3.border = Border(left=Side(style="thin"), right=Side(style="thin"),
                           top=Side(style="thin"), bottom=Side(style="thin"))

    # Ringkasan per kategori pengeluaran
    ws1.cell(row=8, column=1, value="📊 Ringkasan Pengeluaran per Kategori").font = Font(bold=True, size=11)
    ws1.merge_cells("A8:C8")
    ws1["A8"].alignment = Alignment(horizontal="left")

    header_style(ws1, 9, 1, "Kategori",     "374151")
    header_style(ws1, 9, 2, "Total (Rp)",   "374151")
    header_style(ws1, 9, 3, "Jumlah Transaksi", "374151")

    kategori_map = {}
    for t in pengeluaran_list:
        k = t.category or "Lainnya"
        kategori_map[k] = kategori_map.get(k, {"total": 0, "count": 0})
        kategori_map[k]["total"] += t.amount
        kategori_map[k]["count"] += 1

    row_idx = 10
    for kat, val in sorted(kategori_map.items(), key=lambda x: -x[1]["total"]):
        ws1.row_dimensions[row_idx].height = 20
        data_style(ws1, row_idx, 1, kat)
        c = data_style(ws1, row_idx, 2, val["total"], '#,##0')
        c.alignment = Alignment(horizontal="right")
        data_style(ws1, row_idx, 3, val["count"]).alignment = Alignment(horizontal="center")
        row_idx += 1

    set_col_widths(ws1, [28, 20, 18])

    # =====================================
    # SHEET 2: PEMASUKAN
    # =====================================
    ws2 = wb.create_sheet("Pemasukan")
    ws2.row_dimensions[1].height = 28

    ws2.merge_cells("A1:E1")
    t = ws2["A1"]
    t.value = f"💰 Pemasukan — {nama_bulan} {tahun}"
    t.font = Font(bold=True, size=13, color="065F46")
    t.fill = PatternFill("solid", fgColor="D1FAE5")
    t.alignment = Alignment(horizontal="center", vertical="center")

    headers = ["No", "Tanggal", "Kategori", "Catatan", "Jumlah (Rp)"]
    for col, h in enumerate(headers, 1):
        header_style(ws2, 2, col, h, "059669")

    for i, item in enumerate(sorted(pemasukan_list, key=lambda x: x.created_at), 1):
        r = i + 2
        ws2.row_dimensions[r].height = 20
        data_style(ws2, r, 1, i).alignment = Alignment(horizontal="center")
        data_style(ws2, r, 2, item.created_at.strftime("%d-%m-%Y %H:%M"))
        data_style(ws2, r, 3, item.category or "-")
        data_style(ws2, r, 4, item.note or "-")
        c = data_style(ws2, r, 5, item.amount, '#,##0')
        c.alignment = Alignment(horizontal="right")

    # Baris total
    total_row = len(pemasukan_list) + 3
    ws2.cell(row=total_row, column=4, value="TOTAL").font = Font(bold=True)
    ws2.cell(row=total_row, column=4).alignment = Alignment(horizontal="right")
    tc = ws2.cell(row=total_row, column=5, value=total_masuk)
    tc.font = Font(bold=True, color="065F46")
    tc.number_format = '#,##0'
    tc.alignment = Alignment(horizontal="right")
    tc.fill = PatternFill("solid", fgColor="D1FAE5")

    set_col_widths(ws2, [5, 18, 16, 40, 18])

    # =====================================
    # SHEET 3: PENGELUARAN
    # =====================================
    ws3 = wb.create_sheet("Pengeluaran")
    ws3.row_dimensions[1].height = 28

    ws3.merge_cells("A1:E1")
    t3 = ws3["A1"]
    t3.value = f"💸 Pengeluaran — {nama_bulan} {tahun}"
    t3.font = Font(bold=True, size=13, color="991B1B")
    t3.fill = PatternFill("solid", fgColor="FEE2E2")
    t3.alignment = Alignment(horizontal="center", vertical="center")

    for col, h in enumerate(headers, 1):
        header_style(ws3, 2, col, h, "DC2626")

    for i, item in enumerate(sorted(pengeluaran_list, key=lambda x: x.created_at), 1):
        r = i + 2
        ws3.row_dimensions[r].height = 20
        data_style(ws3, r, 1, i).alignment = Alignment(horizontal="center")
        data_style(ws3, r, 2, item.created_at.strftime("%d-%m-%Y %H:%M"))
        data_style(ws3, r, 3, item.category or "-")
        data_style(ws3, r, 4, item.note or "-")
        c = data_style(ws3, r, 5, item.amount, '#,##0')
        c.alignment = Alignment(horizontal="right")

    total_row3 = len(pengeluaran_list) + 3
    ws3.cell(row=total_row3, column=4, value="TOTAL").font = Font(bold=True)
    ws3.cell(row=total_row3, column=4).alignment = Alignment(horizontal="right")
    tc3 = ws3.cell(row=total_row3, column=5, value=total_keluar)
    tc3.font = Font(bold=True, color="991B1B")
    tc3.number_format = '#,##0'
    tc3.alignment = Alignment(horizontal="right")
    tc3.fill = PatternFill("solid", fgColor="FEE2E2")

    set_col_widths(ws3, [5, 18, 16, 40, 18])

    return wb


# =========================
# CHAT AI
# =========================
@app.post("/chat")
def chat(message: Message):

    db: Session = SessionLocal()

    text = message.text.lower()

    # ==================================
    # PENGELUARAN
    # ==================================
    if (
        "keluar" in text
        or "beli" in text
        or "bayar" in text
        or "jajan" in text
    ):

        amount = extract_amount(text)

        category = detect_category(text)

        transaction = Transaction(
            type="pengeluaran",
            amount=amount,
            note=text,
            category=category
        )

        db.add(transaction)
        db.commit()

        pemasukan = db.query(Transaction).filter(
            Transaction.type == "pemasukan"
        ).all()

        pengeluaran = db.query(Transaction).filter(
            Transaction.type == "pengeluaran"
        ).all()

        total_masuk = sum(item.amount for item in pemasukan)
        total_keluar = sum(item.amount for item in pengeluaran)

        saldo = total_masuk - total_keluar

        response = (
            f"💸 Pengeluaran berhasil dicatat\n\n"
            f"Kategori: {category}\n"
            f"Jumlah: Rp{amount:,}"
        )

        # =========================
        # WARNING UMUM
        # =========================

                # WARNING TRANSAKSI BESAR
        if amount >= 500000:

            response += (
                "\n\n🚨 Pengeluaran transaksi ini sangat besar."
            )

        elif amount >= 100000:

            response += (
                "\n\n⚠️ Pengeluaran transaksi ini cukup besar."
            )

        # =========================
        # WARNING SALDO
        # =========================

        if saldo <= 0:

            response += (
                "\n\n🚨 Saldo kamu sudah minus!"
                "\nSegera kurangi pengeluaran."
            )

        elif saldo < 50000:

            response += (
                "\n\n⚠️ Saldo kamu mulai menipis."
            )

        elif saldo < 100000:

            response += (
                "\n\n⚠️ Saldo kamu tinggal sedikit."
            )

        # =========================
        # TOTAL KATEGORI BULAN INI
        # =========================

        bulan_lalu = datetime.utcnow() - timedelta(days=30)

        makanan = 0
        transport = 0
        hiburan = 0
        perawatan = 0
        lainnya = 0

        for item in pengeluaran:

            if item.created_at >= bulan_lalu:

                if item.category == "Makanan":
                    makanan += item.amount

                elif item.category == "Transport":
                    transport += item.amount

                elif item.category == "Hiburan":
                    hiburan += item.amount

                elif item.category == "Perawatan":
                    perawatan += item.amount

                elif item.category == "Lainnya":
                    lainnya += item.amount

        # =========================
        # WARNING KATEGORI BULANAN
        # =========================

        # MAKANAN
        if (
            category == "Makanan"
            and makanan >= 500000
        ):

            response += (
                f"\n\n🍔 Pengeluaran makanan bulan ini "
                f"sudah Rp{makanan:,}."
            )

        # TRANSPORT
        if (
            category == "Transport"
            and transport >= 400000
        ):

            response += (
                f"\n\n🛵 Pengeluaran transport bulan ini "
                f"sudah Rp{transport:,}."
            )

        # HIBURAN
        if (
            category == "Hiburan"
            and hiburan >= 400000
        ):

            response += (
                f"\n\n🎮 Pengeluaran hiburan bulan ini "
                f"sudah Rp{hiburan:,}."
            )

        # PERAWATAN
        if (
            category == "Perawatan"
            and perawatan >= 500000
        ):

            response += (
                f"\n\n🧴 Pengeluaran perawatan bulan ini "
                f"sudah Rp{perawatan:,}."
            )

        # LAINNYA
        if (
            category == "Lainnya"
            and lainnya >= 700000
        ):

            response += (
                f"\n\n📦 Pengeluaran lainnya bulan ini "
                f"sudah Rp{lainnya:,}."
            )

    # ==================================
    # PEMASUKAN
    # ==================================
    elif (
        "masuk" in text
        or "gaji" in text
        or "gajian" in text
        or "transfer dari" in text
    ):

        amount = extract_amount(text)

        transaction = Transaction(
            type="pemasukan",
            amount=amount,
            note=text,
            category="Pemasukan"
        )

        db.add(transaction)
        db.commit()

        response = (
            f"💰 Pemasukan berhasil dicatat\n\n"
            f"Jumlah: Rp{amount:,}"
        )

    # ==================================
    # LAPORAN HARIAN
    # ==================================
    elif "laporan harian" in text:

        today = datetime.utcnow().date()

        transactions = db.query(Transaction).all()

        pemasukan_total = 0
        pengeluaran_total = 0

        kategori_pemasukan = {}
        kategori_pengeluaran = {}

        for item in transactions:

            if item.created_at.date() == today:

                if item.type == "pemasukan":

                    pemasukan_total += item.amount

                    kategori_pemasukan[item.category] = (
                        kategori_pemasukan.get(item.category, 0)
                        + item.amount
                    )

                else:

                    pengeluaran_total += item.amount

                    kategori_pengeluaran[item.category] = (
                        kategori_pengeluaran.get(item.category, 0)
                        + item.amount
                    )

        saldo = pemasukan_total - pengeluaran_total

        response = "📊 LAPORAN HARIAN\n\n"

        response += "💰 PEMASUKAN\n"

        if kategori_pemasukan:

            for kategori, jumlah in kategori_pemasukan.items():
                response += f"• {kategori}: Rp{jumlah:,}\n"

        else:
            response += "Tidak ada pemasukan\n"

        response += f"\nTotal pemasukan: Rp{pemasukan_total:,}\n\n"

        response += "💸 PENGELUARAN\n"

        if kategori_pengeluaran:

            for kategori, jumlah in kategori_pengeluaran.items():
                response += f"• {kategori}: Rp{jumlah:,}\n"

        else:
            response += "Tidak ada pengeluaran\n"

        response += f"\nTotal pengeluaran: Rp{pengeluaran_total:,}\n\n"

        response += f"💳 Sisa saldo: Rp{saldo:,}"

    # ==================================
    # LAPORAN MINGGUAN
    # ==================================
    elif "laporan mingguan" in text:

        minggu_lalu = datetime.utcnow() - timedelta(days=7)

        transactions = db.query(Transaction).all()

        pemasukan_total = 0
        pengeluaran_total = 0

        kategori_pemasukan = {}
        kategori_pengeluaran = {}

        for item in transactions:

            if item.created_at >= minggu_lalu:

                if item.type == "pemasukan":

                    pemasukan_total += item.amount

                    kategori_pemasukan[item.category] = (
                        kategori_pemasukan.get(item.category, 0)
                        + item.amount
                    )

                else:

                    pengeluaran_total += item.amount

                    kategori_pengeluaran[item.category] = (
                        kategori_pengeluaran.get(item.category, 0)
                        + item.amount
                    )

        saldo = pemasukan_total - pengeluaran_total

        response = "📊 LAPORAN MINGGUAN\n\n"

        response += "💰 PEMASUKAN\n"

        if kategori_pemasukan:

            for kategori, jumlah in kategori_pemasukan.items():
                response += f"• {kategori}: Rp{jumlah:,}\n"

        else:
            response += "Tidak ada pemasukan\n"

        response += f"\nTotal pemasukan: Rp{pemasukan_total:,}\n\n"

        response += "💸 PENGELUARAN\n"

        if kategori_pengeluaran:

            for kategori, jumlah in kategori_pengeluaran.items():
                response += f"• {kategori}: Rp{jumlah:,}\n"

        else:
            response += "Tidak ada pengeluaran\n"

        response += f"\nTotal pengeluaran: Rp{pengeluaran_total:,}\n\n"

        response += f"💳 Sisa saldo: Rp{saldo:,}"

    # ==================================
    # LAPORAN BULANAN
    # ==================================
    elif "laporan bulanan" in text:

        bulan_lalu = datetime.utcnow() - timedelta(days=30)

        transactions = db.query(Transaction).all()

        pemasukan_total = 0
        pengeluaran_total = 0

        kategori_pemasukan = {}
        kategori_pengeluaran = {}

        for item in transactions:

            if item.created_at >= bulan_lalu:

                if item.type == "pemasukan":

                    pemasukan_total += item.amount

                    kategori_pemasukan[item.category] = (
                        kategori_pemasukan.get(item.category, 0)
                        + item.amount
                    )

                else:

                    pengeluaran_total += item.amount

                    kategori_pengeluaran[item.category] = (
                        kategori_pengeluaran.get(item.category, 0)
                        + item.amount
                    )

        saldo = pemasukan_total - pengeluaran_total

        response = "📊 LAPORAN BULANAN\n\n"

        response += "💰 PEMASUKAN\n"

        if kategori_pemasukan:

            for kategori, jumlah in kategori_pemasukan.items():
                response += f"• {kategori}: Rp{jumlah:,}\n"

        else:
            response += "Tidak ada pemasukan\n"

        response += f"\nTotal pemasukan: Rp{pemasukan_total:,}\n\n"

        response += "💸 PENGELUARAN\n"

        if kategori_pengeluaran:

            for kategori, jumlah in kategori_pengeluaran.items():
                response += f"• {kategori}: Rp{jumlah:,}\n"

        else:
            response += "Tidak ada pengeluaran\n"

        response += f"\nTotal pengeluaran: Rp{pengeluaran_total:,}\n\n"

        response += f"💳 Sisa saldo: Rp{saldo:,}"

    # ==================================
    # SALDO
    # ==================================
    elif "saldo" in text:

        pemasukan = db.query(Transaction).filter(
            Transaction.type == "pemasukan"
        ).all()

        pengeluaran = db.query(Transaction).filter(
            Transaction.type == "pengeluaran"
        ).all()

        total_masuk = sum(item.amount for item in pemasukan)
        total_keluar = sum(item.amount for item in pengeluaran)

        saldo = total_masuk - total_keluar

        response = f"💳 Saldo kamu sekarang Rp{saldo:,}"

          # ==================================
    # HAPUS RIWAYAT BULANAN
    # ==================================
    elif (
        "hapus laporan" in text
        or "hapus riwayat" in text
    ):

        bulan_map = {
            "januari": 1,
            "februari": 2,
            "maret": 3,
            "april": 4,
            "mei": 5,
            "juni": 6,
            "juli": 7,
            "agustus": 8,
            "september": 9,
            "oktober": 10,
            "november": 11,
            "desember": 12
        }

        now = datetime.utcnow()

        bulan = now.month
        tahun = now.year

        # DETEKSI BULAN
        for nama_bulan, nomor_bulan in bulan_map.items():

            if nama_bulan in text:
                bulan = nomor_bulan

        # DETEKSI TAHUN
        tahun_match = re.search(r'20\d{2}', text)

        if tahun_match:
            tahun = int(tahun_match.group())

        # SIMPAN PENDING DELETE
        pending_delete["bulan"] = bulan
        pending_delete["tahun"] = tahun
        pending_delete["waktu"] = datetime.utcnow()

        nama_bulan = list(bulan_map.keys())[bulan - 1]

        response = (
            f"⚠️ Kamu yakin ingin menghapus "
            f"seluruh transaksi bulan "
            f"{nama_bulan.capitalize()} {tahun}?\n\n"
            f"Ketik:\n"
            f"- ya hapus\n"
            f"- batal"
        )

    # ==================================
    # KONFIRMASI HAPUS
    # ==================================
    elif text == "ya hapus":

        if not pending_delete:

            response = (
                "Tidak ada data yang menunggu dihapus."
            )

        else:

            waktu_request = pending_delete["waktu"]

            selisih = datetime.utcnow() - waktu_request

            # TIMEOUT 30 DETIK
            if selisih.total_seconds() > 30:

                pending_delete.clear()

                response = (
                    "⚠️ Konfirmasi hapus sudah kadaluarsa."
                )

            else:

                bulan = pending_delete["bulan"]
                tahun = pending_delete["tahun"]

                transaksi_hapus = db.query(Transaction).filter(
                    extract("month", Transaction.created_at) == bulan,
                    extract("year", Transaction.created_at) == tahun
                ).all()

                jumlah_data = len(transaksi_hapus)

                for item in transaksi_hapus:
                    db.delete(item)

                db.commit()

                bulan_map = {
                    1: "Januari",
                    2: "Februari",
                    3: "Maret",
                    4: "April",
                    5: "Mei",
                    6: "Juni",
                    7: "Juli",
                    8: "Agustus",
                    9: "September",
                    10: "Oktober",
                    11: "November",
                    12: "Desember"
                }

                response = (
                    f"🗑️ {jumlah_data} transaksi "
                    f"bulan {bulan_map[bulan]} {tahun} "
                    f"berhasil dihapus."
                )

                pending_delete.clear()

    # ==================================
    # BATAL HAPUS
    # ==================================
    elif text == "batal":

        pending_delete.clear()

        response = "❌ Penghapusan dibatalkan."
        
    # ==================================
    # DOWNLOAD EXCEL PER BULAN
    # ==================================
    elif "download laporan" in text:

        bulan_map = {
            "januari": 1,
            "februari": 2,
            "maret": 3,
            "april": 4,
            "mei": 5,
            "juni": 6,
            "juli": 7,
            "agustus": 8,
            "september": 9,
            "oktober": 10,
            "november": 11,
            "desember": 12
        }

        now = datetime.utcnow()

        bulan = now.month
        tahun = now.year

        for nama_bulan, nomor_bulan in bulan_map.items():
            if nama_bulan in text:
                bulan = nomor_bulan

        tahun_match = re.search(r'20\d{2}', text)
        if tahun_match:
            tahun = int(tahun_match.group())

        transactions = db.query(Transaction).all()

        filtered = [
            item for item in transactions
            if item.created_at.month == bulan and item.created_at.year == tahun
        ]

        if not filtered:
            return {"reply": "Tidak ada data di bulan tersebut."}

        nama_bulan_display = list(bulan_map.keys())[bulan - 1].capitalize()
        nama_bulan_file = list(bulan_map.keys())[bulan - 1]
        file_name = f"laporan_{nama_bulan_file}_{tahun}.xlsx"

        wb = buat_laporan_excel(filtered, nama_bulan_display, tahun)
        wb.save(file_name)

        return FileResponse(
            path=file_name,
            filename=file_name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    # ==================================
    # DEFAULT
    # ==================================
    else:
        response = "Aku belum memahami transaksi itu."

    return {
        "reply": response
    }


# ==================================
# DOWNLOAD ENDPOINT
# ==================================
@app.get("/chat-download")
def download_laporan(text: str):

    db: Session = SessionLocal()

    bulan_map = {
        "januari": 1,
        "februari": 2,
        "maret": 3,
        "april": 4,
        "mei": 5,
        "juni": 6,
        "juli": 7,
        "agustus": 8,
        "september": 9,
        "oktober": 10,
        "november": 11,
        "desember": 12
    }

    now = datetime.utcnow()

    bulan = now.month
    tahun = now.year

    for nama_bulan, nomor_bulan in bulan_map.items():
        if nama_bulan in text.lower():
            bulan = nomor_bulan

    tahun_match = re.search(r'20\d{2}', text)
    if tahun_match:
        tahun = int(tahun_match.group())

    transactions = db.query(Transaction).all()

    filtered = [
        item for item in transactions
        if item.created_at.month == bulan and item.created_at.year == tahun
    ]

    if not filtered:
        return {"reply": "Tidak ada data di bulan tersebut."}

    nama_bulan_display = list(bulan_map.keys())[bulan - 1].capitalize()
    nama_bulan_file = list(bulan_map.keys())[bulan - 1]
    file_name = f"laporan_{nama_bulan_file}_{tahun}.xlsx"

    wb = buat_laporan_excel(filtered, nama_bulan_display, tahun)
    wb.save(file_name)

    return FileResponse(
        path=file_name,
        filename=file_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
