import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";
import "./index.css";

// Reemplaza el window.storage propio de los artifacts de Claude por uno que
// habla con nuestra propia API (/api/storage), respaldada por Vercel KV.
// El resto de App.jsx no necesita cambios: usa la misma firma get/set/delete.
window.storage = {
  async get(key) {
    const res = await fetch(`/api/storage?key=${encodeURIComponent(key)}`);
    if (res.status === 404) {
      throw new Error("not found");
    }
    if (!res.ok) throw new Error("storage error");
    return res.json();
  },
  async set(key, value) {
    const res = await fetch("/api/storage", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ key, value }),
    });
    if (!res.ok) return null;
    return res.json();
  },
  async delete(key) {
    const res = await fetch(`/api/storage?key=${encodeURIComponent(key)}`, {
      method: "DELETE",
    });
    if (!res.ok) return null;
    return res.json();
  },
};

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
