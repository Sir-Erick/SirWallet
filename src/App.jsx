import { Filesystem, Directory } from "@capacitor/filesystem";
import { useState, useEffect, useRef } from "react";
import "./App.css";

function App() {
  const [message, setMessage] = useState("");
  const [chat, setChat] = useState([]);

  const [username, setUsername] = useState("");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");

  const [user, setUser] = useState(null);
  const [isRegister, setIsRegister] = useState(false);

  const chatEndRef = useRef(null);

  async function login() {
    const response = await fetch(
      "https://sirwallet-production-b6cc.up.railway.app/login",
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          email,
          password,
        }),
      },
    );

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
    const response = await fetch(
      "https://sirwallet-production-b6cc.up.railway.app/register",
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          username,
          email,
          password,
        }),
      },
    );

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

  async function sendMessage() {
    if (!message.trim()) return;

    const userMessage = {
      sender: "user",
      text: message,
    };

    setChat((prev) => [...prev, userMessage]);

    // DOWNLOAD LAPORAN
    if (message.toLowerCase().includes("download laporan")) {
      const response = await fetch(
        `https://sirwallet-production-b6cc.up.railway.app/chat-download?text=${encodeURIComponent(message)}`,
      );

      // KALAU TIDAK ADA DATA
      if (response.headers.get("content-type")?.includes("application/json")) {
        const data = await response.json();

        const botMessage = {
          sender: "bot",
          text: data.reply,
        };

        setChat((prev) => [...prev, botMessage]);

        setMessage("");

        return;
      }

      // KALAU FILE ADA
      const blob = await response.blob();

      const reader = new FileReader();

      reader.readAsDataURL(blob);

      reader.onloadend = async () => {
        const base64data = reader.result.split(",")[1];

        try {
          await Filesystem.writeFile({
            path: "laporan_keuangan.xlsx",
            data: base64data,
            directory: Directory.Documents,
          });

          setChat((prev) => [
            ...prev,
            {
              sender: "bot",
              text: "📥 Laporan berhasil didownload dan disimpan di folder Documents.",
            },
          ]);
        } catch {
          setChat((prev) => [
            ...prev,
            {
              sender: "bot",
              text: "📥 Laporan berhasil didownload.",
            },
          ]);
        }
      };

      setMessage("");

      return;
    }

    const response = await fetch(
      "https://sirwallet-production-b6cc.up.railway.app/chat",
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          text: message,
          user_id: user.user_id,
        }),
      },
    );

    const data = await response.json();

    const botMessage = {
      sender: "bot",
      text: data.reply,
    };

    setChat((prev) => [...prev, botMessage]);

    setMessage("");
  }

  function handleKeyDown(e) {
    if (e.key === "Enter") {
      sendMessage();
    }
  }

  useEffect(() => {
    const savedUser = localStorage.getItem("sirwallet_user");

    if (savedUser) {
      setUser(JSON.parse(savedUser));
    }
  }, []);

  useEffect(() => {
    if (user && chat.length === 0) {
      setChat([
        {
          sender: "bot",
          text: `👋 Halo ${user.username}!

Selamat datang di SirWallet 💼

Saya bisa membantu mencatat dan mengelola keuanganmu.

📌 Contoh transaksi:
• pemasukan 500rb gaji
• pengeluaran 25000 makan

📌 Informasi:
• saldo saya

📌 Download laporan:
• download laporan
• download laporan april
• download laporan mei

Silakan mulai mencatat 😊`,
        },
      ]);
    }
  }, [user]);

  useEffect(() => {
    chatEndRef.current?.scrollIntoView({
      behavior: "smooth",
    });
  }, [chat]);

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

        <div ref={chatEndRef}></div>
      </div>

      <div className="input-box">
        <input
          type="text"
          placeholder="Tulis transaksi atau tanya laporan keuangan..."
          value={message}
          onChange={(e) => setMessage(e.target.value)}
          onKeyDown={handleKeyDown}
        />

        <button onClick={sendMessage}>➤</button>
      </div>
    </div>
  );
}

export default App;
