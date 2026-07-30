import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  Plus, Search, Download, Upload, X, Pencil, Trash2, Car,
  ArrowDownCircle, ArrowUpCircle, Check, Loader2, LogOut, Users, ChevronDown, Key, Camera, Image as ImageIcon
} from "lucide-react";

// El logo ahora es un archivo estático en /public en vez de un string base64
// gigante embebido en el código (más liviano, más rápido de cargar y de editar).
// Subí tu logo como /public/logo.jpg (o cambiá esta ruta) para reemplazar el placeholder.
const LOGO_MDM = "/logo.jpg";

// ---------- Paleta y tokens ----------
const C = {
  bg: "#12161D",
  surface: "#1B212B",
  surface2: "#232B37",
  border: "#2A3341",
  borderLight: "#38424F",
  text: "#EDEFF3",
  textMuted: "#8B94A5",
  compra: "#4EA87C",
  compraBg: "rgba(78,168,124,0.12)",
  venta: "#E2924F",
  ventaBg: "rgba(226,146,79,0.12)",
  danger: "#D96666",
  dangerBg: "rgba(217,102,102,0.12)",
  accent: "#5B8DEF",
};

const ESTADOS_NEG = ["Contacto inicial", "En negociación", "Acordado", "Cerrado", "Cancelado"];
const ESTADOS_DOC = ["Pendiente", "En trámite", "Completo"];
const FORMAS_PAGO = ["Efectivo", "Transferencia", "Financiación", "Permuta", "Cheque", "Mixto"];

const ESTADO_NEG_COLOR = {
  "Contacto inicial": C.textMuted,
  "En negociación": C.accent,
  "Acordado": C.venta,
  "Cerrado": C.compra,
  "Cancelado": C.danger,
};

const EMPTY_CLIENTE = {
  operacion: "venta",
  nombre: "", dni_cuit: "", telefono: "", email: "", direccion: "",
  localidad: "", provincia: "", codigo_postal: "",
  marca: "", modelo: "", anio: "", patente: "", chasis: "", km: "", color: "",
  precio: "", forma_pago: FORMAS_PAGO[0], plan_ahorro: false,
  estado_negociacion: ESTADOS_NEG[0], estado_documentacion: ESTADOS_DOC[0],
  observaciones: "",
  entrega: false,
  entrega_marca: "", entrega_modelo: "", entrega_anio: "", entrega_patente: "",
  entrega_chasis: "", entrega_km: "", entrega_color: "", entrega_estado: "",
  historial: [],
};

const EXPORT_HEADERS = [
  "Operación", "Nombre", "DNI/CUIT", "Teléfono", "Email", "Dirección",
  "Localidad", "Provincia", "Código postal",
  "Marca", "Modelo", "Año", "Patente", "Chasis/VIN", "Kilometraje", "Color",
  "Precio", "Forma de pago", "Plan de ahorro", "Estado negociación", "Estado documentación",
  "Entrega vehículo", "Marca entrega", "Modelo entrega", "Año entrega",
  "Patente entrega", "Chasis/VIN entrega", "Km entrega", "Color entrega", "Estado entrega",
  "Observaciones", "Cargado por", "Fecha de carga",
];

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

function resizeImageFile(file, maxDim = 1280, quality = 0.72) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new window.Image();
      img.onload = () => {
        let { width, height } = img;
        if (width > maxDim || height > maxDim) {
          if (width >= height) { height = Math.round(height * (maxDim / width)); width = maxDim; }
          else { width = Math.round(width * (maxDim / height)); height = maxDim; }
        }
        const canvas = document.createElement("canvas");
        canvas.width = width; canvas.height = height;
        canvas.getContext("2d").drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.onerror = reject;
      img.src = e.target.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function formatARS(v) {
  const n = Number(v);
  if (!v || isNaN(n)) return "-";
  return new Intl.NumberFormat("es-AR", { style: "currency", currency: "ARS", maximumFractionDigits: 0 }).format(n);
}

function clienteToRow(c) {
  return {
    "Operación": c.operacion === "compra" ? "Compra" : "Venta",
    "Nombre": c.nombre, "DNI/CUIT": c.dni_cuit, "Teléfono": c.telefono, "Email": c.email, "Dirección": c.direccion,
    "Localidad": c.localidad, "Provincia": c.provincia, "Código postal": c.codigo_postal,
    "Marca": c.marca, "Modelo": c.modelo, "Año": c.anio, "Patente": c.patente, "Chasis/VIN": c.chasis,
    "Kilometraje": c.km, "Color": c.color, "Precio": c.precio, "Forma de pago": c.forma_pago,
    "Plan de ahorro": c.plan_ahorro ? "Sí" : "No",
    "Estado negociación": c.estado_negociacion, "Estado documentación": c.estado_documentacion,
    "Entrega vehículo": c.entrega ? "Sí" : "No",
    "Marca entrega": c.entrega_marca, "Modelo entrega": c.entrega_modelo, "Año entrega": c.entrega_anio,
    "Patente entrega": c.entrega_patente, "Chasis/VIN entrega": c.entrega_chasis,
    "Km entrega": c.entrega_km, "Color entrega": c.entrega_color, "Estado entrega": c.entrega_estado,
    "Observaciones": c.observaciones, "Cargado por": c.cargado_por || "", "Fecha de carga": c.fecha_carga || "",
  };
}

function rowToCliente(r) {
  const get = (k) => (r[k] !== undefined && r[k] !== null ? String(r[k]) : "");
  const opRaw = get("Operación").toLowerCase();
  return {
    id: uid(),
    operacion: opRaw.startsWith("compra") ? "compra" : "venta",
    nombre: get("Nombre"), dni_cuit: get("DNI/CUIT"), telefono: get("Teléfono"),
    email: get("Email"), direccion: get("Dirección"),
    localidad: get("Localidad"), provincia: get("Provincia"), codigo_postal: get("Código postal"),
    marca: get("Marca"), modelo: get("Modelo"), anio: get("Año"), patente: get("Patente"),
    chasis: get("Chasis/VIN"), km: get("Kilometraje"), color: get("Color"),
    precio: get("Precio"), forma_pago: get("Forma de pago") || FORMAS_PAGO[0],
    plan_ahorro: get("Plan de ahorro").toLowerCase().startsWith("s"),
    estado_negociacion: get("Estado negociación") || ESTADOS_NEG[0],
    estado_documentacion: get("Estado documentación") || ESTADOS_DOC[0],
    entrega: get("Entrega vehículo").toLowerCase().startsWith("s"),
    entrega_marca: get("Marca entrega"), entrega_modelo: get("Modelo entrega"), entrega_anio: get("Año entrega"),
    entrega_patente: get("Patente entrega"), entrega_chasis: get("Chasis/VIN entrega"),
    entrega_km: get("Km entrega"), entrega_color: get("Color entrega"), entrega_estado: get("Estado entrega"),
    observaciones: get("Observaciones"),
    cargado_por: get("Cargado por") || "Importado",
    fecha_carga: get("Fecha de carga") || new Date().toLocaleDateString("es-AR"),
    historial: [{ id: uid(), fecha: new Date().toLocaleDateString("es-AR"), autor: get("Cargado por") || "Importado", texto: "Cliente importado desde Excel" }],
  };
}

// ---------- Fotos (cámara / galería) ----------
// ARREGLO: antes había un solo botón "Agregar foto" que abría el selector
// genérico de archivos; en muchos celulares eso no ofrece la opción de
// cámara. Ahora hay dos botones separados, cada uno con su propio input y
// el atributo capture="environment" en el de cámara (fuerza la cámara
// trasera en dispositivos móviles).
function FotosSection({ label, fotos, onAdd, onRemove, max = 8 }) {
  const camaraRef = useRef(null);
  const galeriaRef = useRef(null);
  const [subiendo, setSubiendo] = useState(false);

  const procesar = async (fileList) => {
    const files = Array.from(fileList || []).slice(0, Math.max(0, max - fotos.length));
    if (files.length === 0) return;
    setSubiendo(true);
    for (const file of files) {
      try {
        const dataUrl = await resizeImageFile(file);
        onAdd(dataUrl);
      } catch (e) {
        console.error("Error procesando foto", e);
      }
    }
    setSubiendo(false);
  };

  return (
    <div>
      <p className="text-xs font-semibold mb-2 uppercase tracking-wide" style={{ color: C.textMuted }}>{label}</p>
      <div className="flex flex-wrap gap-2">
        {fotos.map((src, i) => (
          <div key={i} className="relative h-16 w-16 rounded-lg overflow-hidden shrink-0" style={{ border: `1px solid ${C.border}` }}>
            <img src={src} className="h-full w-full object-cover" alt="" />
            <button
              onClick={() => onRemove(i)}
              className="absolute top-0 right-0 p-0.5 rounded-bl-lg"
              style={{ backgroundColor: "rgba(0,0,0,0.65)" }}
            >
              <X size={11} color="#fff" />
            </button>
          </div>
        ))}
        {fotos.length < max && (
          <>
            <button
              type="button"
              onClick={() => camaraRef.current?.click()}
              disabled={subiendo}
              className="h-16 w-16 rounded-lg flex flex-col items-center justify-center gap-0.5 shrink-0"
              style={{ backgroundColor: C.surface2, border: `1px dashed ${C.borderLight}` }}
            >
              <Camera size={16} color={C.textMuted} />
              <span className="text-[9px] text-center leading-tight" style={{ color: C.textMuted }}>Cámara</span>
            </button>
            <button
              type="button"
              onClick={() => galeriaRef.current?.click()}
              disabled={subiendo}
              className="h-16 w-16 rounded-lg flex flex-col items-center justify-center gap-0.5 shrink-0"
              style={{ backgroundColor: C.surface2, border: `1px dashed ${C.borderLight}` }}
            >
              <ImageIcon size={16} color={C.textMuted} />
              <span className="text-[9px] text-center leading-tight" style={{ color: C.textMuted }}>Galería</span>
            </button>
          </>
        )}
      </div>
      {subiendo && <p className="text-[10px] mt-1.5" style={{ color: C.textMuted }}>Procesando foto...</p>}
      <input
        ref={camaraRef}
        type="file"
        accept="image/*"
        capture="environment"
        className="hidden"
        onChange={(e) => { procesar(e.target.files); e.target.value = ""; }}
      />
      <input
        ref={galeriaRef}
        type="file"
        accept="image/*"
        multiple
        className="hidden"
        onChange={(e) => { procesar(e.target.files); e.target.value = ""; }}
      />
    </div>
  );
}

// ---------- Toast ----------
function Toast({ msg, onDone }) {
  useEffect(() => {
    const t = setTimeout(onDone, 2400);
    return () => clearTimeout(t);
  }, [msg]);
  if (!msg) return null;
  return (
    <div
      className="fixed bottom-5 left-1/2 -translate-x-1/2 z-50 px-4 py-2.5 rounded-lg text-sm font-medium shadow-lg"
      style={{ backgroundColor: C.surface2, color: C.text, border: `1px solid ${C.borderLight}` }}
    >
      {msg}
    </div>
  );
}

// ---------- Pantalla de ingreso de usuario ----------
function PinDots({ value, length = 4 }) {
  return (
    <div className="flex gap-2 justify-center my-4">
      {Array.from({ length }).map((_, i) => (
        <div
          key={i}
          className="h-3.5 w-3.5 rounded-full"
          style={{ backgroundColor: i < value.length ? C.accent : C.surface2, border: `1px solid ${C.borderLight}` }}
        />
      ))}
    </div>
  );
}

function PinPad({ value, onChange, maxLength = 4 }) {
  const press = (d) => { if (value.length < maxLength) onChange(value + d); };
  const del = () => onChange(value.slice(0, -1));
  return (
    <div className="grid grid-cols-3 gap-2.5 max-w-[240px] mx-auto">
      {["1","2","3","4","5","6","7","8","9"].map((d) => (
        <button key={d} onClick={() => press(d)} className="py-3.5 rounded-xl text-lg font-semibold" style={{ backgroundColor: C.surface, border: `1px solid ${C.border}`, color: C.text }}>
          {d}
        </button>
      ))}
      <div />
      <button onClick={() => press("0")} className="py-3.5 rounded-xl text-lg font-semibold" style={{ backgroundColor: C.surface, border: `1px solid ${C.border}`, color: C.text }}>
        0
      </button>
      <button onClick={del} className="py-3.5 rounded-xl text-sm font-semibold" style={{ backgroundColor: C.surface, border: `1px solid ${C.border}`, color: C.textMuted }}>
        Borrar
      </button>
    </div>
  );
}

function Brand() {
  return (
    <div className="flex items-center gap-2 mb-1">
      <div className="h-11 w-11 rounded-lg flex items-center justify-center overflow-hidden" style={{ backgroundColor: "#F4F5F6" }}>
        <img src={LOGO_MDM} alt="MDM Automotores" className="h-full w-full object-contain p-2" />
      </div>
      <span className="text-sm font-semibold tracking-wide" style={{ color: C.textMuted }}>MDM AUTOMOTORES</span>
    </div>
  );
}

// usuarios: [{ id, nombre, pin, esAdmin }]
function LoginScreen({ usuarios, onCrearAdmin, onLogin }) {
  const [nombreAdmin, setNombreAdmin] = useState("");
  const [pinAdmin, setPinAdmin] = useState("");
  const [pinAdmin2, setPinAdmin2] = useState("");
  const [pasoAdmin, setPasoAdmin] = useState("nombre");

  const [seleccionado, setSeleccionado] = useState(null);
  const [pin, setPin] = useState("");
  const [error, setError] = useState("");

  const intentarLogin = (p) => {
    if (p === seleccionado.pin) {
      onLogin(seleccionado);
    } else {
      setError("PIN incorrecto");
      setTimeout(() => { setPin(""); setError(""); }, 700);
    }
  };

  useEffect(() => {
    if (seleccionado && pin.length === 4) intentarLogin(pin);
  }, [pin]);

  useEffect(() => {
    if (pasoAdmin === "pin" && pinAdmin.length === 4) setPasoAdmin("confirmar");
    if (pasoAdmin === "confirmar" && pinAdmin2.length === 4) {
      if (pinAdmin2 === pinAdmin) {
        onCrearAdmin(nombreAdmin.trim(), pinAdmin);
      } else {
        setError("Los PIN no coinciden");
        setTimeout(() => { setPinAdmin(""); setPinAdmin2(""); setPasoAdmin("pin"); setError(""); }, 800);
      }
    }
  }, [pinAdmin, pinAdmin2]);

  if (usuarios.length === 0) {
    return (
      <div className="min-h-screen flex items-center justify-center px-5" style={{ backgroundColor: C.bg }}>
        <div className="w-full max-w-sm text-center">
          <div className="flex justify-center"><Brand /></div>
          <h1 className="text-xl font-bold mb-1" style={{ color: C.text }}>Configurar administrador</h1>
          <p className="text-xs mb-6" style={{ color: C.textMuted }}>Esta es la primera vez que se abre la app. Vas a ser el administrador.</p>

          {pasoAdmin === "nombre" && (
            <div className="space-y-3">
              <input
                autoFocus
                value={nombreAdmin}
                onChange={(e) => setNombreAdmin(e.target.value)}
                placeholder="Tu nombre"
                className="w-full px-4 py-3 rounded-xl outline-none text-sm text-center"
                style={{ backgroundColor: C.surface, border: `1px solid ${C.borderLight}`, color: C.text }}
                onKeyDown={(e) => e.key === "Enter" && nombreAdmin.trim() && setPasoAdmin("pin")}
              />
              <button
                disabled={!nombreAdmin.trim()}
                onClick={() => setPasoAdmin("pin")}
                className="w-full px-4 py-3 rounded-xl text-sm font-semibold disabled:opacity-40"
                style={{ backgroundColor: C.accent, color: "#0B0E13" }}
              >
                Continuar
              </button>
            </div>
          )}

          {pasoAdmin === "pin" && (
            <div>
              <p className="text-sm font-medium mb-1" style={{ color: C.text }}>Elegí un PIN de 4 dígitos</p>
              <PinDots value={pinAdmin} />
              <PinPad value={pinAdmin} onChange={setPinAdmin} />
            </div>
          )}

          {pasoAdmin === "confirmar" && (
            <div>
              <p className="text-sm font-medium mb-1" style={{ color: C.text }}>Confirmá el PIN</p>
              <PinDots value={pinAdmin2} />
              {error && <p className="text-xs font-medium mb-1" style={{ color: C.danger }}>{error}</p>}
              <PinPad value={pinAdmin2} onChange={setPinAdmin2} />
            </div>
          )}
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen flex items-center justify-center px-5" style={{ backgroundColor: C.bg }}>
      <div className="w-full max-w-sm text-center">
        <div className="flex justify-center"><Brand /></div>

        {!seleccionado ? (
          <>
            <h1 className="text-2xl font-bold mb-6" style={{ color: C.text }}>¿Quién sos?</h1>
            <div className="space-y-2">
              {usuarios.map((u) => (
                <button
                  key={u.id}
                  onClick={() => { setSeleccionado(u); setPin(""); }}
                  className="w-full flex items-center justify-between px-4 py-3 rounded-xl text-left transition-colors"
                  style={{ backgroundColor: C.surface, border: `1px solid ${C.border}`, color: C.text }}
                >
                  <span className="font-medium">{u.nombre}{u.esAdmin ? " · admin" : ""}</span>
                  <ChevronDown size={16} className="-rotate-90" color={C.textMuted} />
                </button>
              ))}
            </div>
            <p className="text-[11px] mt-4" style={{ color: C.textMuted }}>
              ¿Sos nuevo? Pedile al administrador que te cree un usuario con un PIN.
            </p>
          </>
        ) : (
          <div>
            <p className="text-sm font-medium mb-1" style={{ color: C.text }}>Hola, {seleccionado.nombre}</p>
            <p className="text-xs mb-1" style={{ color: C.textMuted }}>Ingresá tu PIN</p>
            <PinDots value={pin} />
            {error && <p className="text-xs font-medium mb-1" style={{ color: C.danger }}>{error}</p>}
            <PinPad value={pin} onChange={setPin} />
            <button onClick={() => { setSeleccionado(null); setPin(""); }} className="text-xs font-medium mt-4" style={{ color: C.textMuted }}>
              Volver a la lista
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

// ---------- Cambiar PIN propio ----------
function CambiarPin({ usuarioActual, onGuardar, onClose }) {
  const [paso, setPaso] = useState("actual");
  const [pinActual, setPinActual] = useState("");
  const [pinNuevo, setPinNuevo] = useState("");
  const [pinNuevo2, setPinNuevo2] = useState("");
  const [error, setError] = useState("");

  useEffect(() => {
    if (paso === "actual" && pinActual.length === 4) {
      if (pinActual === usuarioActual.pin) { setPaso("nuevo"); setError(""); }
      else { setError("PIN actual incorrecto"); setTimeout(() => { setPinActual(""); setError(""); }, 700); }
    }
  }, [pinActual]);

  useEffect(() => {
    if (paso === "nuevo" && pinNuevo.length === 4) setPaso("confirmar");
  }, [pinNuevo]);

  useEffect(() => {
    if (paso === "confirmar" && pinNuevo2.length === 4) {
      if (pinNuevo2 === pinNuevo) onGuardar(pinNuevo);
      else { setError("No coincide"); setTimeout(() => { setPinNuevo(""); setPinNuevo2(""); setPaso("nuevo"); setError(""); }, 700); }
    }
  }, [pinNuevo2]);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center px-6" style={{ backgroundColor: "rgba(0,0,0,0.6)" }}>
      <div className="w-full max-w-xs rounded-2xl p-5 text-center" style={{ backgroundColor: C.surface, border: `1px solid ${C.border}` }}>
        <div className="flex justify-between items-center mb-3">
          <p className="text-sm font-semibold" style={{ color: C.text }}>Cambiar mi PIN</p>
          <button onClick={onClose}><X size={18} color={C.textMuted} /></button>
        </div>
        {paso === "actual" && <p className="text-xs mb-1" style={{ color: C.textMuted }}>Ingresá tu PIN actual</p>}
        {paso === "nuevo" && <p className="text-xs mb-1" style={{ color: C.textMuted }}>Elegí tu nuevo PIN</p>}
        {paso === "confirmar" && <p className="text-xs mb-1" style={{ color: C.textMuted }}>Confirmá el nuevo PIN</p>}
        <PinDots value={paso === "actual" ? pinActual : paso === "nuevo" ? pinNuevo : pinNuevo2} />
        {error && <p className="text-xs font-medium mb-1" style={{ color: C.danger }}>{error}</p>}
        <PinPad
          value={paso === "actual" ? pinActual : paso === "nuevo" ? pinNuevo : pinNuevo2}
          onChange={paso === "actual" ? setPinActual : paso === "nuevo" ? setPinNuevo : setPinNuevo2}
        />
      </div>
    </div>
  );
}

// ---------- Panel del administrador: gestionar usuarios ----------
// ARREGLO: se agregó onEliminar + usuarioActual como props, botón de
// eliminar por usuario (oculto para admins y para el usuario logueado) y
// confirmación inline antes de borrar.
function PanelUsuarios({ usuarios, onCrear, onEliminar, usuarioActual, onClose }) {
  const [nombre, setNombre] = useState("");
  const [pin, setPin] = useState("");
  const [creando, setCreando] = useState(false);
  const [msg, setMsg] = useState("");
  const [confirmarId, setConfirmarId] = useState(null);

  const crear = () => {
    if (!nombre.trim() || pin.length !== 4) return;
    onCrear(nombre.trim(), pin);
    setMsg(`Usuario "${nombre.trim()}" creado con PIN ${pin}. Compartíselo para que ingrese.`);
    setNombre(""); setPin(""); setCreando(false);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center px-6" style={{ backgroundColor: "rgba(0,0,0,0.6)" }}>
      <div className="w-full max-w-sm rounded-2xl p-5" style={{ backgroundColor: C.surface, border: `1px solid ${C.border}` }}>
        <div className="flex justify-between items-center mb-3">
          <p className="text-sm font-semibold" style={{ color: C.text }}>Usuarios</p>
          <button onClick={onClose}><X size={18} color={C.textMuted} /></button>
        </div>

        <div className="space-y-1.5 mb-4 max-h-48 overflow-y-auto">
          {usuarios.map((u) => (
            <div key={u.id} className="flex items-center justify-between px-3 py-2 rounded-lg" style={{ backgroundColor: C.surface2 }}>
              <span className="text-xs font-medium" style={{ color: C.text }}>{u.nombre}</span>
              <div className="flex items-center gap-2">
                {u.esAdmin && <span className="text-[10px] font-semibold" style={{ color: C.accent }}>admin</span>}
                {!u.esAdmin && u.id !== usuarioActual?.id && (
                  <button onClick={() => setConfirmarId(u.id)}>
                    <Trash2 size={14} color={C.danger} />
                  </button>
                )}
              </div>
            </div>
          ))}
        </div>

        {confirmarId && (
          <div className="mb-3 p-3 rounded-lg" style={{ backgroundColor: C.surface2, border: `1px solid ${C.border}` }}>
            <p className="text-xs mb-2" style={{ color: C.text }}>¿Eliminar este usuario?</p>
            <div className="flex gap-2">
              <button onClick={() => setConfirmarId(null)} className="flex-1 py-1.5 rounded-lg text-xs font-semibold" style={{ backgroundColor: C.surface, color: C.text }}>
                Cancelar
              </button>
              <button
                onClick={() => { onEliminar(confirmarId); setConfirmarId(null); }}
                className="flex-1 py-1.5 rounded-lg text-xs font-semibold"
                style={{ backgroundColor: C.danger, color: "#fff" }}
              >
                Eliminar
              </button>
            </div>
          </div>
        )}

        {msg && <p className="text-xs mb-3 p-2 rounded-lg" style={{ backgroundColor: C.compraBg, color: C.compra }}>{msg}</p>}

        {!creando ? (
          <button
            onClick={() => { setCreando(true); setMsg(""); }}
            className="w-full flex items-center justify-center gap-1.5 py-2.5 rounded-lg text-xs font-semibold"
            style={{ backgroundColor: C.accent, color: "#0B0E13" }}
          >
            <Plus size={14} /> Crear usuario
          </button>
        ) : (
          <div className="space-y-2.5">
            <input
              autoFocus
              value={nombre}
              onChange={(e) => setNombre(e.target.value)}
              placeholder="Nombre del nuevo usuario"
              className="w-full px-3 py-2.5 rounded-lg outline-none text-sm"
              style={{ backgroundColor: C.surface2, border: `1px solid ${C.border}`, color: C.text }}
            />
            <input
              value={pin}
              onChange={(e) => setPin(e.target.value.replace(/\D/g, "").slice(0, 4))}
              placeholder="PIN inicial de 4 dígitos"
              inputMode="numeric"
              className="w-full px-3 py-2.5 rounded-lg outline-none text-sm"
              style={{ backgroundColor: C.surface2, border: `1px solid ${C.border}`, color: C.text }}
            />
            <p className="text-[10px]" style={{ color: C.textMuted }}>El usuario podrá cambiar este PIN después de ingresar.</p>
            <div className="flex gap-2">
              <button onClick={() => setCreando(false)} className="flex-1 py-2.5 rounded-lg text-xs font-semibold" style={{ backgroundColor: C.surface2, color: C.text }}>
                Cancelar
              </button>
              <button
                onClick={crear}
                disabled={!nombre.trim() || pin.length !== 4}
                className="flex-1 py-2.5 rounded-lg text-xs font-semibold disabled:opacity-40"
                style={{ backgroundColor: C.accent, color: "#0B0E13" }}
              >
                Crear
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ---------- Formulario de cliente ----------
function ClienteForm({ initial, usuarioActual, onSave, onClose }) {
  const [f, setF] = useState(initial || EMPTY_CLIENTE);
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const clienteIdRef = useRef(f.id || uid());
  const [fotosVehiculo, setFotosVehiculo] = useState([]);
  const [fotosEntrega, setFotosEntrega] = useState([]);

  useEffect(() => {
    if (!f.id) return;
    (async () => {
      try {
        const r = await window.storage.get(`fotos:${f.id}`, true);
        const data = r ? JSON.parse(r.value) : { vehiculo: [], entrega: [] };
        setFotosVehiculo(data.vehiculo || []);
        setFotosEntrega(data.entrega || []);
      } catch {
        setFotosVehiculo([]); setFotosEntrega([]);
      }
    })();
  }, []);

  const Field = ({ label, k, type = "text", full }) => (
    <div className={full ? "sm:col-span-2" : ""}>
      <label className="block text-xs font-medium mb-1" style={{ color: C.textMuted }}>{label}</label>
      <input
        type={type}
        value={f[k]}
        onChange={set(k)}
        className="w-full px-3 py-2.5 rounded-lg outline-none text-sm"
        style={{ backgroundColor: C.surface2, border: `1px solid ${C.border}`, color: C.text }}
      />
    </div>
  );

  const Select = ({ label, k, options }) => (
    <div>
      <label className="block text-xs font-medium mb-1" style={{ color: C.textMuted }}>{label}</label>
      <select
        value={f[k]}
        onChange={set(k)}
        className="w-full px-3 py-2.5 rounded-lg outline-none text-sm"
        style={{ backgroundColor: C.surface2, border: `1px solid ${C.border}`, color: C.text }}
      >
        {options.map((o) => <option key={o} value={o}>{o}</option>)}
      </select>
    </div>
  );

  const submit = async () => {
    if (!f.nombre.trim()) return;
    const now = new Date().toLocaleDateString("es-AR");
    const esNuevo = !f.id;
    const idFinal = clienteIdRef.current;
    const nuevoEvento = {
      id: uid(),
      fecha: now,
      autor: usuarioActual.nombre,
      texto: esNuevo ? "Cliente creado" : "Datos actualizados",
    };
    try {
      await window.storage.set(`fotos:${idFinal}`, JSON.stringify({ vehiculo: fotosVehiculo, entrega: fotosEntrega }), true);
    } catch (e) {
      console.error("No se pudieron guardar las fotos", e);
    }
    onSave({
      ...f,
      id: idFinal,
      cargado_por: f.cargado_por || usuarioActual.nombre,
      fecha_carga: f.fecha_carga || now,
      historial: [...(f.historial || []), nuevoEvento],
    });
  };

  return (
    <div className="fixed inset-0 z-40 flex sm:items-center sm:justify-center" style={{ backgroundColor: "rgba(0,0,0,0.6)" }}>
      <div
        className="w-full sm:max-w-2xl sm:rounded-2xl sm:max-h-[88vh] h-full sm:h-auto overflow-y-auto"
        style={{ backgroundColor: C.bg, border: `1px solid ${C.border}` }}
      >
        <div className="sticky top-0 flex items-center justify-between px-5 py-4 z-10" style={{ backgroundColor: C.bg, borderBottom: `1px solid ${C.border}` }}>
          <h2 className="text-base font-bold" style={{ color: C.text }}>
            {f.id ? "Editar cliente" : "Nuevo cliente"}
          </h2>
          <button onClick={onClose}><X size={20} color={C.textMuted} /></button>
        </div>

        <div className="p-5 space-y-5">
          <div className="flex rounded-lg overflow-hidden" style={{ border: `1px solid ${C.border}` }}>
            {["venta", "compra"].map((op) => (
              <button
                key={op}
                onClick={() => setF({ ...f, operacion: op })}
                className="flex-1 py-2.5 text-sm font-semibold flex items-center justify-center gap-1.5"
                style={{
                  backgroundColor: f.operacion === op ? (op === "compra" ? C.compraBg : C.ventaBg) : "transparent",
                  color: f.operacion === op ? (op === "compra" ? C.compra : C.venta) : C.textMuted,
                }}
              >
                {op === "compra" ? <ArrowDownCircle size={15} /> : <ArrowUpCircle size={15} />}
                {op === "compra" ? "Compra (nos vende)" : "Venta (nos compra)"}
              </button>
            ))}
          </div>

          <div>
            <p className="text-xs font-semibold mb-2 uppercase tracking-wide" style={{ color: C.textMuted }}>Datos del cliente</p>
            <div className="grid sm:grid-cols-2 gap-3">
              <Field label="Nombre y apellido" k="nombre" full />
              <Field label="DNI / CUIT" k="dni_cuit" />
              <Field label="Teléfono" k="telefono" />
              <Field label="Email" k="email" />
              <Field label="Dirección" k="direccion" full />
              <Field label="Localidad" k="localidad" />
              <Field label="Provincia" k="provincia" />
              <Field label="Código postal" k="codigo_postal" />
            </div>
          </div>

          <div>
            <p className="text-xs font-semibold mb-2 uppercase tracking-wide" style={{ color: C.textMuted }}>Vehículo</p>
            <div className="grid sm:grid-cols-2 gap-3">
              <Field label="Marca" k="marca" />
              <Field label="Modelo" k="modelo" />
              <Field label="Año" k="anio" />
              <Field label="Patente" k="patente" />
              <Field label="Chasis / VIN" k="chasis" />
              <Field label="Kilometraje" k="km" />
              <Field label="Color" k="color" />
            </div>
            <div className="mt-3">
              <FotosSection
                label="Fotos del vehículo"
                fotos={fotosVehiculo}
                onAdd={(src) => setFotosVehiculo((p) => [...p, src])}
                onRemove={(i) => setFotosVehiculo((p) => p.filter((_, idx) => idx !== i))}
              />
            </div>
          </div>

          <div>
            <p className="text-xs font-semibold mb-2 uppercase tracking-wide" style={{ color: C.textMuted }}>Comercial</p>
            <div className="grid sm:grid-cols-2 gap-3">
              <Field label="Precio (ARS)" k="precio" type="number" />
              <Select label="Forma de pago" k="forma_pago" options={FORMAS_PAGO} />
              <Select label="Estado de negociación" k="estado_negociacion" options={ESTADOS_NEG} />
              <Select label="Estado de documentación" k="estado_documentacion" options={ESTADOS_DOC} />
            </div>
            <label className="flex items-center gap-2.5 cursor-pointer select-none mt-3">
              <input
                type="checkbox"
                checked={f.plan_ahorro}
                onChange={(e) => setF({ ...f, plan_ahorro: e.target.checked })}
                className="h-4 w-4"
              />
              <span className="text-xs font-medium" style={{ color: C.text }}>Ingresa a plan de ahorro</span>
            </label>
          </div>

          <div>
            <label className="flex items-center gap-2.5 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={f.entrega}
                onChange={(e) => setF({ ...f, entrega: e.target.checked })}
                className="h-4 w-4"
              />
              <span className="text-xs font-semibold uppercase tracking-wide" style={{ color: C.textMuted }}>
                El cliente entrega un vehículo (parte de pago)
              </span>
            </label>
            {f.entrega && (
              <div className="grid sm:grid-cols-2 gap-3 mt-3 p-3 rounded-lg" style={{ backgroundColor: C.surface2, border: `1px solid ${C.border}` }}>
                <Field label="Marca" k="entrega_marca" />
                <Field label="Modelo" k="entrega_modelo" />
                <Field label="Año" k="entrega_anio" />
                <Field label="Patente" k="entrega_patente" />
                <Field label="Chasis / VIN" k="entrega_chasis" />
                <Field label="Kilometraje" k="entrega_km" />
                <Field label="Color" k="entrega_color" />
                <Field label="Estado del vehículo" k="entrega_estado" full />
                <div className="sm:col-span-2">
                  <FotosSection
                    label="Fotos del vehículo que entrega"
                    fotos={fotosEntrega}
                    onAdd={(src) => setFotosEntrega((p) => [...p, src])}
                    onRemove={(i) => setFotosEntrega((p) => p.filter((_, idx) => idx !== i))}
                  />
                </div>
              </div>
            )}
          </div>

          <div>
            <label className="block text-xs font-medium mb-1" style={{ color: C.textMuted }}>Observaciones</label>
            <textarea
              value={f.observaciones}
              onChange={set("observaciones")}
              rows={3}
              className="w-full px-3 py-2.5 rounded-lg outline-none text-sm resize-none"
              style={{ backgroundColor: C.surface2, border: `1px solid ${C.border}`, color: C.text }}
            />
          </div>
        </div>

        <div className="sticky bottom-0 flex gap-3 px-5 py-4" style={{ backgroundColor: C.bg, borderTop: `1px solid ${C.border}` }}>
          <button onClick={onClose} className="flex-1 py-3 rounded-xl text-sm font-semibold" style={{ backgroundColor: C.surface2, color: C.text }}>
            Cancelar
          </button>
          <button
            onClick={submit}
            disabled={!f.nombre.trim()}
            className="flex-1 py-3 rounded-xl text-sm font-semibold disabled:opacity-40"
            style={{ backgroundColor: C.accent, color: "#0B0E13" }}
          >
            Guardar
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------- Tarjeta de cliente ----------
function ClienteCard({ c, onEdit, onDelete, onOpen }) {
  const isCompra = c.operacion === "compra";
  return (
    <div
      onClick={() => onOpen(c)}
      className="rounded-xl p-4 relative cursor-pointer active:opacity-80"
      style={{ backgroundColor: C.surface, border: `1px solid ${C.border}`, borderLeft: `3px solid ${isCompra ? C.compra : C.venta}` }}
    >
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <div className="flex items-center gap-2 mb-1">
            <span
              className="text-[10px] font-bold px-1.5 py-0.5 rounded uppercase tracking-wide"
              style={{ backgroundColor: isCompra ? C.compraBg : C.ventaBg, color: isCompra ? C.compra : C.venta }}
            >
              {isCompra ? "Compra" : "Venta"}
            </span>
            <span className="text-[11px] font-medium" style={{ color: ESTADO_NEG_COLOR[c.estado_negociacion] }}>
              {c.estado_negociacion}
            </span>
            {c.entrega && (
              <span className="text-[10px] font-semibold px-1.5 py-0.5 rounded" style={{ backgroundColor: C.surface2, color: C.textMuted }}>
                + entrega
              </span>
            )}
          </div>
          <p className="font-semibold text-sm truncate" style={{ color: C.text }}>{c.nombre || "Sin nombre"}</p>
          <p className="text-xs truncate" style={{ color: C.textMuted }}>
            {[c.marca, c.modelo, c.anio].filter(Boolean).join(" ") || "Vehículo sin datos"}
            {c.patente && ` · ${c.patente}`}
          </p>
        </div>
        <div className="flex gap-1 shrink-0">
          <button onClick={(e) => { e.stopPropagation(); onEdit(c); }} className="p-1.5 rounded-lg" style={{ backgroundColor: C.surface2 }}>
            <Pencil size={14} color={C.textMuted} />
          </button>
          <button onClick={(e) => { e.stopPropagation(); onDelete(c.id); }} className="p-1.5 rounded-lg" style={{ backgroundColor: C.surface2 }}>
            <Trash2 size={14} color={C.danger} />
          </button>
        </div>
      </div>
      <div className="flex items-center justify-between mt-3 pt-3" style={{ borderTop: `1px solid ${C.border}` }}>
        <span className="text-sm font-bold" style={{ color: C.text }}>{formatARS(c.precio)}</span>
        <span className="text-[11px]" style={{ color: C.textMuted }}>{c.telefono || c.email || "sin contacto"}</span>
      </div>
    </div>
  );
}

// ---------- Detalle e historial del cliente ----------
function ClienteDetalle({ cliente, usuarioActual, onClose, onAddNota, onEdit }) {
  const [nota, setNota] = useState("");
  const [fotos, setFotos] = useState({ vehiculo: [], entrega: [] });
  const isCompra = cliente.operacion === "compra";

  useEffect(() => {
    (async () => {
      try {
        const r = await window.storage.get(`fotos:${cliente.id}`, true);
        setFotos(r ? JSON.parse(r.value) : { vehiculo: [], entrega: [] });
      } catch {
        setFotos({ vehiculo: [], entrega: [] });
      }
    })();
  }, [cliente.id]);

  const Row = ({ label, value }) => (
    !value ? null : (
      <div className="flex justify-between gap-3 py-1.5" style={{ borderBottom: `1px solid ${C.border}` }}>
        <span className="text-xs" style={{ color: C.textMuted }}>{label}</span>
        <span className="text-xs font-medium text-right" style={{ color: C.text }}>{value}</span>
      </div>
    )
  );

  const Galeria = ({ label, srcs }) => (
    !srcs || srcs.length === 0 ? null : (
      <div>
        <p className="text-xs font-semibold mb-1.5 uppercase tracking-wide" style={{ color: C.textMuted }}>{label}</p>
        <div className="flex flex-wrap gap-2">
          {srcs.map((src, i) => (
            <a key={i} href={src} target="_blank" rel="noreferrer" className="h-20 w-20 rounded-lg overflow-hidden block" style={{ border: `1px solid ${C.border}` }}>
              <img src={src} className="h-full w-full object-cover" alt="" />
            </a>
          ))}
        </div>
      </div>
    )
  );

  const enviarNota = () => {
    if (!nota.trim()) return;
    onAddNota(cliente.id, nota.trim());
    setNota("");
  };

  const historialOrdenado = [...(cliente.historial || [])].reverse();

  return (
    <div className="fixed inset-0 z-40 flex sm:items-center sm:justify-center" style={{ backgroundColor: "rgba(0,0,0,0.6)" }}>
      <div
        className="w-full sm:max-w-lg sm:rounded-2xl sm:max-h-[88vh] h-full sm:h-auto overflow-y-auto"
        style={{ backgroundColor: C.bg, border: `1px solid ${C.border}` }}
      >
        <div className="sticky top-0 flex items-center justify-between px-5 py-4 z-10" style={{ backgroundColor: C.bg, borderBottom: `1px solid ${C.border}` }}>
          <div>
            <span
              className="text-[10px] font-bold px-1.5 py-0.5 rounded uppercase tracking-wide"
              style={{ backgroundColor: isCompra ? C.compraBg : C.ventaBg, color: isCompra ? C.compra : C.venta }}
            >
              {isCompra ? "Compra" : "Venta"}
            </span>
            <h2 className="text-base font-bold mt-1" style={{ color: C.text }}>{cliente.nombre}</h2>
          </div>
          <div className="flex items-center gap-2">
            <button onClick={() => onEdit(cliente)} className="p-2 rounded-lg" style={{ backgroundColor: C.surface2 }}>
              <Pencil size={15} color={C.textMuted} />
            </button>
            <button onClick={onClose}><X size={20} color={C.textMuted} /></button>
          </div>
        </div>

        <div className="p-5 space-y-5">
          <div>
            <p className="text-xs font-semibold mb-1.5 uppercase tracking-wide" style={{ color: C.textMuted }}>Cliente</p>
            <Row label="DNI / CUIT" value={cliente.dni_cuit} />
            <Row label="Teléfono" value={cliente.telefono} />
            <Row label="Email" value={cliente.email} />
            <Row label="Dirección" value={cliente.direccion} />
            <Row label="Localidad" value={cliente.localidad} />
            <Row label="Provincia" value={cliente.provincia} />
            <Row label="Código postal" value={cliente.codigo_postal} />
          </div>

          <div>
            <p className="text-xs font-semibold mb-1.5 uppercase tracking-wide" style={{ color: C.textMuted }}>Vehículo</p>
            <Row label="Marca / Modelo" value={[cliente.marca, cliente.modelo].filter(Boolean).join(" ")} />
            <Row label="Año" value={cliente.anio} />
            <Row label="Patente" value={cliente.patente} />
            <Row label="Chasis / VIN" value={cliente.chasis} />
            <Row label="Kilometraje" value={cliente.km} />
            <Row label="Color" value={cliente.color} />
            <div className="mt-2"><Galeria label="Fotos del vehículo" srcs={fotos.vehiculo} /></div>
          </div>

          <div>
            <p className="text-xs font-semibold mb-1.5 uppercase tracking-wide" style={{ color: C.textMuted }}>Comercial</p>
            <Row label="Precio" value={formatARS(cliente.precio)} />
            <Row label="Forma de pago" value={cliente.forma_pago} />
            <Row label="Plan de ahorro" value={cliente.plan_ahorro ? "Sí" : "No"} />
            <Row label="Estado de negociación" value={cliente.estado_negociacion} />
            <Row label="Estado de documentación" value={cliente.estado_documentacion} />
          </div>

          {cliente.entrega && (
            <div>
              <p className="text-xs font-semibold mb-1.5 uppercase tracking-wide" style={{ color: C.textMuted }}>Vehículo que entrega</p>
              <Row label="Marca / Modelo" value={[cliente.entrega_marca, cliente.entrega_modelo].filter(Boolean).join(" ")} />
              <Row label="Año" value={cliente.entrega_anio} />
              <Row label="Patente" value={cliente.entrega_patente} />
              <Row label="Chasis / VIN" value={cliente.entrega_chasis} />
              <Row label="Kilometraje" value={cliente.entrega_km} />
              <Row label="Color" value={cliente.entrega_color} />
              <Row label="Estado" value={cliente.entrega_estado} />
              <div className="mt-2"><Galeria label="Fotos del vehículo que entrega" srcs={fotos.entrega} /></div>
            </div>
          )}

          {cliente.observaciones && (
            <div>
              <p className="text-xs font-semibold mb-1.5 uppercase tracking-wide" style={{ color: C.textMuted }}>Observaciones</p>
              <p className="text-xs" style={{ color: C.text }}>{cliente.observaciones}</p>
            </div>
          )}

          <div>
            <p className="text-xs font-semibold mb-2 uppercase tracking-wide" style={{ color: C.textMuted }}>Historial</p>
            <div className="flex gap-2 mb-3">
              <input
                value={nota}
                onChange={(e) => setNota(e.target.value)}
                placeholder="Agregar una nota al historial..."
                className="flex-1 px-3 py-2 rounded-lg outline-none text-xs"
                style={{ backgroundColor: C.surface2, border: `1px solid ${C.border}`, color: C.text }}
                onKeyDown={(e) => e.key === "Enter" && enviarNota()}
              />
              <button onClick={enviarNota} className="px-3 rounded-lg text-xs font-semibold" style={{ backgroundColor: C.accent, color: "#0B0E13" }}>
                Agregar
              </button>
            </div>
            <div className="space-y-2.5">
              {historialOrdenado.length === 0 && (
                <p className="text-xs" style={{ color: C.textMuted }}>Sin eventos todavía.</p>
              )}
              {historialOrdenado.map((h) => (
                <div key={h.id} className="pl-3 relative" style={{ borderLeft: `2px solid ${C.border}` }}>
                  <p className="text-xs" style={{ color: C.text }}>{h.texto}</p>
                  <p className="text-[10px]" style={{ color: C.textMuted }}>{h.autor} · {h.fecha}</p>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

// ---------- App principal ----------
export default function App() {
  const [loading, setLoading] = useState(true);
  const [usuarios, setUsuarios] = useState([]);
  const [usuarioActual, setUsuarioActual] = useState(null);
  const [clientes, setClientes] = useState([]);
  const [filtro, setFiltro] = useState("todas");
  const [busqueda, setBusqueda] = useState("");
  const [modalOpen, setModalOpen] = useState(false);
  const [detalleCliente, setDetalleCliente] = useState(null);
  const [editando, setEditando] = useState(null);
  const [confirmarBorrar, setConfirmarBorrar] = useState(null);
  const [toast, setToast] = useState("");
  const [panelUsuariosOpen, setPanelUsuariosOpen] = useState(false);
  const [cambiarPinOpen, setCambiarPinOpen] = useState(false);
  const fileInputRef = useRef(null);

  useEffect(() => {
    (async () => {
      try {
        let u = [];
        let c = [];
        try {
          const ru = await window.storage.get("usuarios", true);
          u = ru ? JSON.parse(ru.value) : [];
        } catch { u = []; }
        try {
          const rc = await window.storage.get("clientes", true);
          c = rc ? JSON.parse(rc.value) : [];
        } catch { c = []; }

        const formatoValido = u.every(
          (x) => x && typeof x === "object" && typeof x.nombre === "string" && typeof x.pin === "string"
        );
        if (!formatoValido) {
          u = [];
          try { await window.storage.set("usuarios", JSON.stringify(u), true); } catch {}
        }

        setUsuarios(u);
        setClientes(c);
      } catch (e) {
        console.error(e);
      } finally {
        setLoading(false);
      }
    })();
  }, []);

  const persistClientes = async (next) => {
    setClientes(next);
    try {
      await window.storage.set("clientes", JSON.stringify(next), true);
    } catch (e) {
      setToast("No se pudo guardar. Reintentá.");
    }
  };

  const persistUsuarios = async (next) => {
    setUsuarios(next);
    try {
      await window.storage.set("usuarios", JSON.stringify(next), true);
    } catch (e) {
      setToast("No se pudo guardar el usuario.");
    }
  };

  const handleCrearAdmin = async (nombre, pin) => {
    const admin = { id: uid(), nombre, pin, esAdmin: true };
    await persistUsuarios([admin]);
    setUsuarioActual(admin);
  };

  const handleCrearUsuarioAdmin = async (nombre, pin) => {
    const nuevo = { id: uid(), nombre, pin, esAdmin: false };
    await persistUsuarios([...usuarios, nuevo]);
  };

  // ARREGLO: función nueva, antes no existía forma de eliminar usuarios.
  const handleEliminarUsuarioAdmin = async (id) => {
    const next = usuarios.filter((u) => u.id !== id);
    await persistUsuarios(next);
    setToast("Usuario eliminado");
  };

  const handleLogin = (usuario) => {
    setUsuarioActual(usuario);
  };

  const handleCambiarPin = async (nuevoPin) => {
    const next = usuarios.map((u) => (u.id === usuarioActual.id ? { ...u, pin: nuevoPin } : u));
    await persistUsuarios(next);
    setUsuarioActual({ ...usuarioActual, pin: nuevoPin });
    setCambiarPinOpen(false);
    setToast("PIN actualizado");
  };

  const handleGuardarCliente = async (c) => {
    const exists = clientes.some((x) => x.id === c.id);
    const next = exists ? clientes.map((x) => (x.id === c.id ? c : x)) : [...clientes, c];
    await persistClientes(next);
    setModalOpen(false);
    setEditando(null);
    setToast(exists ? "Cliente actualizado" : "Cliente agregado");
  };

  const handleBorrar = async (id) => {
    await persistClientes(clientes.filter((c) => c.id !== id));
    try {
      const r = await window.storage.get(`fotos:${id}`, true);
      if (r) await window.storage.delete(`fotos:${id}`, true);
    } catch {}
    setConfirmarBorrar(null);
    setToast("Cliente eliminado");
  };

  const handleAddNota = async (clienteId, texto) => {
    const evento = { id: uid(), fecha: new Date().toLocaleDateString("es-AR"), autor: usuarioActual.nombre, texto };
    const next = clientes.map((c) =>
      c.id === clienteId ? { ...c, historial: [...(c.historial || []), evento] } : c
    );
    await persistClientes(next);
    const actualizado = next.find((c) => c.id === clienteId);
    setDetalleCliente(actualizado);
  };

  // ARREGLO: la generación se envuelve en try/catch completo (antes solo
  // protegía el click de descarga) y usa un data: URL en base64 en vez de
  // Blob + URL.createObjectURL, mucho más confiable en navegadores móviles.
  const exportarExcel = () => {
    if (clientes.length === 0) {
      setToast("No hay clientes para exportar");
      return;
    }
    try {
      const rows = clientes.map(clienteToRow);
      const ws = XLSX.utils.json_to_sheet(rows, { header: EXPORT_HEADERS });
      ws["!autofilter"] = { ref: ws["!ref"] };
      ws["!cols"] = EXPORT_HEADERS.map(() => ({ wch: 16 }));

      const historialRows = [];
      clientes.forEach((c) => {
        (c.historial || []).forEach((h) => {
          historialRows.push({
            "Cliente": c.nombre,
            "Operación": c.operacion === "compra" ? "Compra" : "Venta",
            "Vehículo": [c.marca, c.modelo, c.anio].filter(Boolean).join(" "),
            "Evento": h.texto,
            "Usuario": h.autor,
            "Fecha": h.fecha,
          });
        });
      });
      const wsHist = XLSX.utils.json_to_sheet(historialRows, {
        header: ["Cliente", "Operación", "Vehículo", "Evento", "Usuario", "Fecha"],
      });
      if (wsHist["!ref"]) wsHist["!autofilter"] = { ref: wsHist["!ref"] };
      wsHist["!cols"] = [{ wch: 22 }, { wch: 10 }, { wch: 22 }, { wch: 30 }, { wch: 14 }, { wch: 12 }];

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Clientes");
      XLSX.utils.book_append_sheet(wb, wsHist, "Historial");

      const nombreArchivo = `clientes_vehiculos_${new Date().toISOString().slice(0, 10)}.xlsx`;
      const out = XLSX.write(wb, { type: "array", bookType: "xlsx" });
      const file = new File([out], nombreArchivo, {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      // En celulares, navigator.share (con archivos) suele funcionar mejor
      // que el <a download> tradicional: abre el panel nativo de "compartir"
      // del sistema operativo, desde donde se puede guardar en Archivos,
      // Drive, WhatsApp, etc. Si no está disponible, cae al método clásico.
      if (navigator.canShare && navigator.canShare({ files: [file] })) {
        navigator.share({
          files: [file],
          title: "Clientes MDM Automotores",
        }).then(() => {
          setToast("Excel compartido");
        }).catch((err) => {
          if (err?.name !== "AbortError") {
            setToast("No se pudo compartir el Excel");
          }
        });
        return;
      }

      const blob = new Blob([out], { type: file.type });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = nombreArchivo;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 10000);
      setToast("Excel descargado");
    } catch (e) {
      console.error(e);
      setToast("No se pudo generar el Excel: " + (e?.message || "error desconocido"));
    }
  };

  const importarExcel = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
        const nuevos = rows.filter((r) => (r["Nombre"] || "").toString().trim()).map(rowToCliente);
        if (nuevos.length === 0) {
          setToast("El archivo no tiene filas válidas");
          return;
        }
        await persistClientes([...clientes, ...nuevos]);
        setToast(`${nuevos.length} clientes importados`);
      } catch (err) {
        setToast("No se pudo leer el archivo");
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const filtrados = useMemo(() => {
    return clientes
      .filter((c) => filtro === "todas" || c.operacion === filtro)
      .filter((c) => {
        if (!busqueda.trim()) return true;
        const q = busqueda.toLowerCase();
        return [c.nombre, c.patente, c.marca, c.modelo, c.telefono, c.dni_cuit]
          .some((v) => (v || "").toLowerCase().includes(q));
      })
      .sort((a, b) => (b.fecha_carga || "").localeCompare(a.fecha_carga || ""));
  }, [clientes, filtro, busqueda]);

  const counts = useMemo(() => ({
    compra: clientes.filter((c) => c.operacion === "compra" && c.estado_negociacion !== "Cerrado" && c.estado_negociacion !== "Cancelado").length,
    venta: clientes.filter((c) => c.operacion === "venta" && c.estado_negociacion !== "Cerrado" && c.estado_negociacion !== "Cancelado").length,
    cerradas: clientes.filter((c) => c.estado_negociacion === "Cerrado").length,
  }), [clientes]);

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center" style={{ backgroundColor: C.bg }}>
        <Loader2 className="animate-spin" color={C.accent} size={28} />
      </div>
    );
  }

  if (!usuarioActual) {
    return <LoginScreen usuarios={usuarios} onCrearAdmin={handleCrearAdmin} onLogin={handleLogin} />;
  }

  return (
    <div className="min-h-screen pb-24" style={{ backgroundColor: C.bg }}>
      <div className="sticky top-0 z-20 px-4 pt-4 pb-3" style={{ backgroundColor: C.bg, borderBottom: `1px solid ${C.border}` }}>
        <div className="flex items-center justify-between mb-3">
          <div className="flex items-center gap-2">
            <div className="h-9 w-9 rounded-lg flex items-center justify-center overflow-hidden" style={{ backgroundColor: "#F4F5F6" }}>
              <img src={LOGO_MDM} alt="MDM Automotores" className="h-full w-full object-contain p-2" />
            </div>
            <div>
              <p className="text-sm font-bold leading-tight" style={{ color: C.text }}>MDM Automotores</p>
              <p className="text-[11px] leading-tight" style={{ color: C.textMuted }}>
                {usuarioActual.nombre}{usuarioActual.esAdmin ? " · admin" : ""}
              </p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            {usuarioActual.esAdmin && (
              <button onClick={() => setPanelUsuariosOpen(true)} className="p-2 rounded-lg" style={{ backgroundColor: C.surface2 }}>
                <Users size={15} color={C.textMuted} />
              </button>
            )}
            <button onClick={() => setCambiarPinOpen(true)} className="p-2 rounded-lg" style={{ backgroundColor: C.surface2 }}>
              <Key size={15} color={C.textMuted} />
            </button>
            <button onClick={() => setUsuarioActual(null)} className="p-2 rounded-lg" style={{ backgroundColor: C.surface2 }}>
              <LogOut size={15} color={C.textMuted} />
            </button>
          </div>
        </div>

        <div className="grid grid-cols-3 gap-2 mb-3">
          <div className="rounded-lg px-3 py-2" style={{ backgroundColor: C.compraBg }}>
            <p className="text-lg font-bold" style={{ color: C.compra }}>{counts.compra}</p>
            <p className="text-[10px] font-medium" style={{ color: C.compra }}>Compras activas</p>
          </div>
          <div className="rounded-lg px-3 py-2" style={{ backgroundColor: C.ventaBg }}>
            <p className="text-lg font-bold" style={{ color: C.venta }}>{counts.venta}</p>
            <p className="text-[10px] font-medium" style={{ color: C.venta }}>Ventas activas</p>
          </div>
          <div className="rounded-lg px-3 py-2" style={{ backgroundColor: C.surface2 }}>
            <p className="text-lg font-bold" style={{ color: C.text }}>{counts.cerradas}</p>
            <p className="text-[10px] font-medium" style={{ color: C.textMuted }}>Cerradas</p>
          </div>
        </div>

        <div className="relative mb-3">
          <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2" color={C.textMuted} />
          <input
            value={busqueda}
            onChange={(e) => setBusqueda(e.target.value)}
            placeholder="Buscar por nombre, patente, marca..."
            className="w-full pl-9 pr-3 py-2.5 rounded-lg outline-none text-sm"
            style={{ backgroundColor: C.surface, border: `1px solid ${C.border}`, color: C.text }}
          />
        </div>

        <div className="flex gap-2">
          {[
            { k: "todas", label: "Todas" },
            { k: "venta", label: "Venta" },
            { k: "compra", label: "Compra" },
          ].map((op) => (
            <button
              key={op.k}
              onClick={() => setFiltro(op.k)}
              className="px-3 py-1.5 rounded-full text-xs font-semibold"
              style={{
                backgroundColor: filtro === op.k ? C.accent : C.surface2,
                color: filtro === op.k ? "#0B0E13" : C.textMuted,
              }}
            >
              {op.label}
            </button>
          ))}
        </div>
      </div>

      <div className="px-4 pt-3">
        <button
          onClick={() => { setEditando(null); setModalOpen(true); }}
          className="w-full flex items-center justify-center gap-2 py-3 rounded-xl text-sm font-bold"
          style={{ backgroundColor: C.accent, color: "#0B0E13" }}
        >
          <Plus size={17} /> Cargar cliente
        </button>
      </div>

      <div className="px-4 pt-3 flex gap-2">
        <button
          onClick={exportarExcel}
          className="flex-1 flex items-center justify-center gap-1.5 py-2 rounded-lg text-xs font-semibold"
          style={{ backgroundColor: C.surface, border: `1px solid ${C.border}`, color: C.text }}
        >
          <Download size={14} /> Exportar Excel
        </button>
        <button
          onClick={() => fileInputRef.current?.click()}
          className="flex-1 flex items-center justify-center gap-1.5 py-2 rounded-lg text-xs font-semibold"
          style={{ backgroundColor: C.surface, border: `1px solid ${C.border}`, color: C.text }}
        >
          <Upload size={14} /> Importar Excel
        </button>
        <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={importarExcel} className="hidden" />
      </div>

      <div className="px-4 pt-4 space-y-2.5">
        {filtrados.length === 0 && (
          <div className="text-center py-16">
            <Users size={28} color={C.textMuted} className="mx-auto mb-2" />
            <p className="text-sm font-medium" style={{ color: C.textMuted }}>
              {clientes.length === 0 ? "Todavía no cargaste clientes" : "No hay resultados"}
            </p>
          </div>
        )}
        {filtrados.map((c) => (
          <ClienteCard
            key={c.id}
            c={c}
            onEdit={(cl) => { setEditando(cl); setModalOpen(true); }}
            onDelete={(id) => setConfirmarBorrar(id)}
            onOpen={(cl) => setDetalleCliente(cl)}
          />
        ))}
      </div>

      <button
        onClick={() => { setEditando(null); setModalOpen(true); }}
        className="fixed bottom-6 right-5 h-14 w-14 rounded-full flex items-center justify-center shadow-lg z-30"
        style={{ backgroundColor: C.accent }}
      >
        <Plus size={24} color="#0B0E13" />
      </button>

      {modalOpen && (
        <ClienteForm
          initial={editando}
          usuarioActual={usuarioActual}
          onSave={handleGuardarCliente}
          onClose={() => { setModalOpen(false); setEditando(null); }}
        />
      )}

      {detalleCliente && (
        <ClienteDetalle
          cliente={detalleCliente}
          usuarioActual={usuarioActual}
          onClose={() => setDetalleCliente(null)}
          onAddNota={handleAddNota}
          onEdit={(cl) => { setDetalleCliente(null); setEditando(cl); setModalOpen(true); }}
        />
      )}

      {confirmarBorrar && (
        <div className="fixed inset-0 z-50 flex items-center justify-center px-6" style={{ backgroundColor: "rgba(0,0,0,0.6)" }}>
          <div className="w-full max-w-xs rounded-2xl p-5" style={{ backgroundColor: C.surface, border: `1px solid ${C.border}` }}>
            <p className="text-sm font-semibold mb-1" style={{ color: C.text }}>¿Eliminar cliente?</p>
            <p className="text-xs mb-4" style={{ color: C.textMuted }}>Esta acción no se puede deshacer.</p>
            <div className="flex gap-2">
              <button onClick={() => setConfirmarBorrar(null)} className="flex-1 py-2.5 rounded-lg text-xs font-semibold" style={{ backgroundColor: C.surface2, color: C.text }}>
                Cancelar
              </button>
              <button onClick={() => handleBorrar(confirmarBorrar)} className="flex-1 py-2.5 rounded-lg text-xs font-semibold" style={{ backgroundColor: C.danger, color: "#fff" }}>
                Eliminar
              </button>
            </div>
          </div>
        </div>
      )}

      {panelUsuariosOpen && (
        <PanelUsuarios
          usuarios={usuarios}
          usuarioActual={usuarioActual}
          onCrear={handleCrearUsuarioAdmin}
          onEliminar={handleEliminarUsuarioAdmin}
          onClose={() => setPanelUsuariosOpen(false)}
        />
      )}

      {cambiarPinOpen && (
        <CambiarPin
          usuarioActual={usuarioActual}
          onGuardar={handleCambiarPin}
          onClose={() => setCambiarPinOpen(false)}
        />
      )}

      <Toast msg={toast} onDone={() => setToast("")} />
    </div>
  );
}
