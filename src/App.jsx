import { Filesystem, Directory } from "@capacitor/filesystem";
import { useState, useEffect, useRef } from "react";
import "./App.css";

const BASE_URL = "https://sirwallet-production-b6cc.up.railway.app";

function App() {
  const [message, setMessage] = useState("");
  const [chat, setChat] = useState([]);
  const [username, setUsername] = useState("");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [user, setUser] = useState(null);
  const [isRegister, setIsRegister] = useState(false);
  const [showCommands, setShowCommands] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const chatEndRef = useRef(null);

  // ─── AUTH ────────────────────────────────────────────────
  async function login() {
    const response = await fetch(`${BASE_URL}/login`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email, password }),
    });
    const data = await response.json();
    if (data.user_id) {
      localStorage.setItem("sirwallet_user", JSON.stringify(data));
      setUser(data);
    } else {
      alert(data.message);
    }
  }

  function logout() {
    localStorage.removeItem("sirwallet_user");
    setUser(null);
    setChat([]);
    setMessage("");
    setEmail("");
    setPassword("");
  }

  async function register() {
    const response = await fetch(`${BASE_URL}/register`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ username, email, password }),
    });
    const data = await response.json();
    if (data.message === "Register berhasil.") {
      alert("Register berhasil 🎉");
      setIsRegister(false);
      setUsername("");
      setEmail("");
      setPassword("");
    } else {
      alert(data.message);
    }
  }

  // ─── SEND MESSAGE ────────────────────────────────────────
  async function sendMessage() {
    if (!message.trim() || isLoading) return;

    const trimmed = message.trim();
    setMessage("");
    setShowCommands(false);
    setIsLoading(true);

    // Tambah pesan user ke chat
    setChat((prev) => [...prev, { sender: "user", text: trimmed }]);

    // ── DOWNLOAD LAPORAN ──────────────────────────────────
    if (trimmed.toLowerCase().includes("download laporan")) {
      await handleDownloadLaporan(trimmed);
      setIsLoading(false);
      return;
    }

    // ── CHAT BIASA ────────────────────────────────────────
    try {
      const response = await fetch(`${BASE_URL}/chat`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ text: trimmed, user_id: user?.user_id }),
      });
      const data = await response.json();
      setChat((prev) => [...prev, { sender: "bot", text: data.reply }]);
    } catch {
      setChat((prev) => [
        ...prev,
        { sender: "bot", text: "❌ Gagal terhubung ke server. Coba lagi." },
      ]);
    } finally {
      setIsLoading(false);
    }
  }

  // ─── HANDLER DOWNLOAD LAPORAN ────────────────────────────
  async function handleDownloadLaporan(text) {
    // Deteksi nama bulan dari teks untuk nama file
    const bulanMap = {
      januari: "januari",
      februari: "februari",
      maret: "maret",
      april: "april",
      mei: "mei",
      juni: "juni",
      juli: "juli",
      agustus: "agustus",
      september: "september",
      oktober: "oktober",
      november: "november",
      desember: "desember",
    };

    const textLower = text.toLowerCase();
    let namaBulanFile = "";
    for (const [nama] of Object.entries(bulanMap)) {
      if (textLower.includes(nama)) {
        namaBulanFile = nama;
        break;
      }
    }

    const tahunMatch = text.match(/20\d{2}/);
    const tahun = tahunMatch
      ? tahunMatch[0]
      : new Date().getFullYear().toString();

    if (!namaBulanFile) {
      // Bulan sekarang jika tidak disebutkan
      const bln = [
        "januari",
        "februari",
        "maret",
        "april",
        "mei",
        "juni",
        "juli",
        "agustus",
        "september",
        "oktober",
        "november",
        "desember",
      ];
      namaBulanFile = bln[new Date().getMonth()];
    }

    const fileName = `laporan_${namaBulanFile}_${tahun}.xlsx`;

    try {
      const response = await fetch(
        `${BASE_URL}/chat-download?text=${encodeURIComponent(text)}`,
      );

      // Cek apakah respons adalah JSON (berarti tidak ada data)
      const contentType = response.headers.get("content-type") || "";
      if (contentType.includes("application/json")) {
        const data = await response.json();
        setChat((prev) => [
          ...prev,
          {
            sender: "bot",
            text: data.reply || "Tidak ada data di bulan tersebut.",
          },
        ]);
        return;
      }

      // Respons adalah file Excel
      const blob = await response.blob();
      const saved = await simpanFileDiHP(blob, fileName);

      if (saved.success) {
        setChat((prev) => [
          ...prev,
          {
            sender: "bot",
            text:
              `📊 Laporan keuangan berhasil dibuat.\n\n` +
              `📄 File: ${fileName}\n` +
              `📁 Lokasi: ${saved.path}\n\n` +
              `Silakan buka File Manager untuk melihat atau membagikan laporan.`,
          },
        ]);
      } else {
        // Fallback: download via browser (untuk web/PWA)
        downloadViaBrowser(blob, fileName);
        setChat((prev) => [
          ...prev,
          {
            sender: "bot",
            text:
              `📊 Laporan keuangan berhasil dibuat.\n\n` +
              `📄 File: ${fileName}\n` +
              `📁 Lokasi: ${saved.path}\n\n` +
              `Silakan buka File Manager untuk melihat atau membagikan laporan.`,
          },
        ]);
      }
    } catch (err) {
      setChat((prev) => [
        ...prev,
        {
          sender: "bot",
          text: `❌ Gagal mengunduh laporan. Coba lagi.\n(${err.message})`,
        },
      ]);
    }
  }

  // ─── SIMPAN FILE DI HP (Capacitor / Android) ─────────────
  async function simpanFileDiHP(blob, fileName) {
    try {
      // Konversi blob ke base64
      const base64data = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(blob);
        reader.onloadend = () => resolve(reader.result.split(",")[1]);
        reader.onerror = () => reject(new Error("FileReader gagal"));
      });

      // Coba simpan ke Documents dulu, fallback ke Downloads
      const directories = [
        { dir: Directory.External, label: "Download/SirWallet" },
        { dir: Directory.Documents, label: "Documents/SirWallet" },
        { dir: Directory.Data, label: "Internal Storage/SirWallet" },
      ];

      for (const { dir, label } of directories) {
        try {
          const result = await Filesystem.writeFile({
            path: `SirWallet/${fileName}`,
            data: base64data,
            directory: dir,
            recursive: true, // otomatis buat folder SirWallet
          });
          return {
            success: true,
            path: `${label}/${fileName}`,
            uri: result.uri,
          };
        } catch {
          // Coba direktori berikutnya
          continue;
        }
      }

      // Semua direktori gagal
      return { success: false };
    } catch {
      return { success: false };
    }
  }

  // ─── FALLBACK: DOWNLOAD VIA BROWSER (PWA / Web) ──────────
  function downloadViaBrowser(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  // ─── KEYBOARD ────────────────────────────────────────────
  function handleKeyDown(e) {
    if (e.key === "Enter") sendMessage();
  }

  function handleInputChange(e) {
    const val = e.target.value;
    setMessage(val);
    setShowCommands(val.startsWith("/"));
  }

  // ─── EFFECTS ─────────────────────────────────────────────
  useEffect(() => {
    const savedUser = localStorage.getItem("sirwallet_user");
    if (savedUser) setUser(JSON.parse(savedUser));
  }, []);

  useEffect(() => {
    if (user && chat.length === 0) {
      setChat([
        {
          sender: "bot",
          text:
            `👋 Halo ${user.username}!\n\n` +
            `Selamat datang di SirWallet 💼\n\n` +
            `Saya bisa membantu mencatat dan mengelola keuanganmu.\n\n` +
            `📌 Contoh transaksi:\n` +
            `• pemasukan 500rb gaji\n` +
            `• pengeluaran 25000 makan\n\n` +
            `📌 Informasi:\n` +
            `• saldo saya\n\n` +
            `📌 Download laporan Excel:\n` +
            `• download laporan\n` +
            `• download laporan april\n` +
            `• download laporan mei 2025\n\n` +
            `File tersimpan otomatis di folder SirWallet 📁`,
        },
      ]);
    }
  }, [user]);

  useEffect(() => {
    chatEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [chat]);

  // ─── LOGIN / REGISTER SCREEN ─────────────────────────────
  if (!user) {
    return (
      <div className="login-container">
        <div className="login-box">
          <div className="login-brand">
            <h1>💼 SirWallet</h1>
            <p>
              {isRegister
                ? "Create your SirWallet account"
                : "Personal Finance Assistant"}
            </p>
          </div>

          {isRegister && (
            <input
              type="text"
              placeholder="Username"
              value={username}
              onChange={(e) => setUsername(e.target.value)}
            />
          )}

          <input
            type="email"
            placeholder="Email"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
          />

          <input
            type="password"
            placeholder="Password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
          />

          <button onClick={isRegister ? register : login}>
            {isRegister ? "Register" : "Login"}
          </button>

          <div className="auth-switch">
            {isRegister ? (
              <p>
                Sudah punya akun?{" "}
                <span onClick={() => setIsRegister(false)}>Login</span>
              </p>
            ) : (
              <p>
                Belum punya akun?{" "}
                <span onClick={() => setIsRegister(true)}>Register</span>
              </p>
            )}
          </div>
        </div>
      </div>
    );
  }

  // ─── MAIN CHAT SCREEN ────────────────────────────────────
  return (
    <div className="container">
      <div className="header">
        <div>💼 SirWallet</div>
        <button className="logout-btn" onClick={logout}>
          Logout
        </button>
      </div>

      <div className="chat-box">
        {chat.map((item, index) => (
          <div
            key={index}
            className={item.sender === "user" ? "message user" : "message bot"}
          >
            {item.text}
          </div>
        ))}

        {isLoading && (
          <div className="message bot loading-indicator">
            <span>.</span>
            <span>.</span>
            <span>.</span>
          </div>
        )}

        <div ref={chatEndRef} />
      </div>

      {/* Command suggestions */}
      {showCommands && (
        <div className="command-menu">
          {commands
            .filter((cmd) => cmd.startsWith(message))
            .map((cmd, index) => (
              <div
                key={index}
                className="command-item"
                onClick={() => {
                  setMessage(cmd.replace("/", ""));
                  setShowCommands(false);
                }}
              >
                {cmd}
              </div>
            ))}
        </div>
      )}

      <div className="input-box">
        <input
          type="text"
          placeholder="Tulis transaksi atau tanya laporan keuangan..."
          value={message}
          onChange={handleInputChange}
          onKeyDown={handleKeyDown}
          disabled={isLoading}
        />
        <button onClick={sendMessage} disabled={isLoading}>
          {isLoading ? "⏳" : "➤"}
        </button>
      </div>
    </div>
  );
}

const commands = [
  "/saldo",
  "/laporan harian",
  "/laporan mingguan",
  "/laporan bulanan",
  "/download laporan",
  "/download laporan april",
  "/download laporan mei",
  "/pemasukan",
  "/pengeluaran",
  "/hapus riwayat",
];

export default App;
