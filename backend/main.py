import calendar
import re
from openpyxl import Workbook
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

        # DETEKSI BULAN DARI CHAT
        for nama_bulan, nomor_bulan in bulan_map.items():

            if nama_bulan in text:
                bulan = nomor_bulan

        # DETEKSI TAHUN
        tahun_match = re.search(r'20\d{2}', text)

        if tahun_match:
            tahun = int(tahun_match.group())

        transactions = db.query(Transaction).all()

        wb = Workbook()

        ws = wb.active
        ws.title = "Laporan Keuangan"

        ws.append([
            "Tipe",
            "Kategori",
            "Jumlah",
            "Catatan",
            "Tanggal"
        ])

        total_data = 0

        for item in transactions:

            if (
                item.created_at.month == bulan
                and item.created_at.year == tahun
            ):

                ws.append([
                    item.type,
                    item.category,
                    item.amount,
                    item.note,
                    item.created_at.strftime("%d-%m-%Y %H:%M")
                ])

                total_data += 1

        if total_data == 0:

            return {
                "reply": "Tidak ada data di bulan tersebut."
            }

        nama_bulan_file = list(bulan_map.keys())[bulan - 1]

        file_name = f"laporan_{nama_bulan_file}_{tahun}.xlsx"

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

    wb = Workbook()

    ws = wb.active
    ws.title = "Laporan Keuangan"

    ws.append([
        "Tipe",
        "Kategori",
        "Jumlah",
        "Catatan",
        "Tanggal"
    ])

    total_data = 0

    for item in transactions:

        if (
            item.created_at.month == bulan
            and item.created_at.year == tahun
        ):

            ws.append([
                item.type,
                item.category,
                item.amount,
                item.note,
                item.created_at.strftime("%d-%m-%Y %H:%M")
            ])

            total_data += 1

    if total_data == 0:
        return {
            "reply": "Tidak ada data di bulan tersebut."
        }

    nama_bulan_file = list(bulan_map.keys())[bulan - 1]

    file_name = f"laporan_{nama_bulan_file}_{tahun}.xlsx"

    wb.save(file_name)

    return FileResponse(
        path=file_name,
        filename=file_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )