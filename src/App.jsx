import React, { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import { Upload, Image as ImageIcon, Trash2, Plus, Minus, ShoppingBag, X, Mail, Package, Check, AlertCircle, Search, Lock, RotateCcw, CreditCard, ListOrdered, Eye, EyeOff, MessageCircle, FileText, Tag } from "lucide-react";

const uid = () => Math.random().toString(36).slice(2, 10);
const emptyProduct = () => ({ id: uid(), sku: "", stock: 0, price: 0, description: "", marca: "", rubro: "", status: "activo", image: null });

function normalizeHeader(h) {
  return String(h || "").trim().toLowerCase().replace(/[áàä]/g, "a").replace(/[éèë]/g, "e").replace(/[íìï]/g, "i").replace(/[óòö]/g, "o").replace(/[úùü]/g, "u");
}

function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        const parsed = rows.map((row) => {
          const map = {};
          Object.keys(row).forEach((k) => { map[normalizeHeader(k)] = row[k]; });
          return {
            id: uid(),
            sku: String(map["sku"] ?? map["codigo"] ?? map["cod"] ?? "").trim(),
            stock: Number(map["stock"] ?? 0) || 0,
            price: Number(map["precio"] ?? map["lista1"] ?? map["preciolista1"] ?? map["preciodeventa"] ?? 0) || 0,
            description: String(map["descripcion"] ?? map["producto"] ?? "").trim(),
            marca: String(map["marca"] ?? "").trim(),
            rubro: String(map["rubro"] ?? "").trim(),
            image: null,
          };
        }).filter((p) => p.sku);
        resolve(parsed);
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function fileToDataURL(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function money(n) {
  return new Intl.NumberFormat("es-AR", { style: "currency", currency: "ARS", maximumFractionDigits: 0 }).format(Number(n) || 0);
}

async function api(path, options = {}) {
  const res = await fetch(path, options);
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || "Error de red");
  return data;
}

export default function App() {
  const [tab, setTab] = useState("tienda");
  const [products, setProducts] = useState([]);
  const [email, setEmail] = useState("");
  const [whatsappNumber, setWhatsappNumber] = useState("");
  const [whatsappMessage, setWhatsappMessage] = useState("");
  const [loading, setLoading] = useState(true);
  const [toast, setToast] = useState(null);
  const [adminSecret, setAdminSecret] = useState("");
  const [authed, setAuthed] = useState(false);

  useEffect(() => {
    (async () => {
      try {
        const p = await api("/api/products");
        setProducts(p);
      } catch {}
      try {
        const c = await api("/api/config");
        setEmail(c.email || "");
        setWhatsappNumber(c.whatsappNumber || "");
        setWhatsappMessage(c.whatsappMessage || "");
      } catch {}
      setLoading(false);
    })();
  }, []);

  function flash(msg) {
    setToast(msg);
    setTimeout(() => setToast(null), 2600);
  }

  async function refreshPublicProducts() {
    try {
      const p = await api("/api/products");
      setProducts(p);
    } catch {}
  }

  async function persistConfig(fields) {
    if ("email" in fields) setEmail(fields.email);
    if ("whatsappNumber" in fields) setWhatsappNumber(fields.whatsappNumber);
    if ("whatsappMessage" in fields) setWhatsappMessage(fields.whatsappMessage);
    try {
      await api("/api/admin/config", {
        method: "PUT",
        headers: { "Content-Type": "application/json", "x-admin-secret": adminSecret },
        body: JSON.stringify({ email, whatsappNumber, whatsappMessage, ...fields }),
      });
    } catch {}
  }

  return (
    <div style={{ minHeight: "100vh", background: "#FBF8F3", color: "#20191c" }}>
      <style>{`
        * { box-sizing: border-box; }
        .btn { border: none; cursor: pointer; font-family: 'Inter', sans-serif; font-weight: 600; transition: transform .12s ease, opacity .12s ease; }
        .btn:active { transform: scale(0.97); }
        .btn:focus-visible, input:focus-visible, textarea:focus-visible { outline: 2px solid #8A6D3B; outline-offset: 2px; }
        input, textarea, select { font-family: 'Inter', sans-serif; }
      `}</style>

      <Header tab={tab} setTab={setTab} />

      {loading ? (
        <div style={{ padding: 60, textAlign: "center", color: "#8a7f74" }}>Cargando catálogo…</div>
      ) : tab === "tienda" ? (
        <Store products={products} email={email} flash={flash} />
      ) : tab === "ofertas" ? (
        <Offers />
      ) : authed ? (
        <Admin
          email={email} whatsappNumber={whatsappNumber} whatsappMessage={whatsappMessage} setConfig={persistConfig}
          adminSecret={adminSecret} onCatalogChanged={refreshPublicProducts} flash={flash}
        />
      ) : (
        <AdminGate onSuccess={(secret) => { setAdminSecret(secret); setAuthed(true); }} />
      )}

      {tab !== "admin" && <Footer whatsappNumber={whatsappNumber} whatsappMessage={whatsappMessage} />}

      {toast && (
        <div style={{ position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)", background: "#20191c", color: "#FBF8F3", padding: "12px 22px", borderRadius: 8, fontSize: 14, fontWeight: 500, boxShadow: "0 8px 24px rgba(0,0,0,0.25)", zIndex: 200 }}>
          {toast}
        </div>
      )}
    </div>
  );
}

function Header({ tab, setTab }) {
  return (
    <div style={{ borderBottom: "1px solid #E7DFD3", background: "#FBF8F3", position: "sticky", top: 0, zIndex: 50 }}>
      <div style={{ maxWidth: 1100, margin: "0 auto", padding: "14px 24px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <img src="/logo.png" alt="PC. Perfumería Boutique" style={{ height: 52, objectFit: "contain" }} />
        <div style={{ display: "flex", gap: 4, background: "#F0E9DC", padding: 4, borderRadius: 10 }}>
          {[["tienda", "Tienda"], ["ofertas", "Ofertas"], ["admin", "Admin"]].map(([key, label]) => (
            <button key={key} className="btn" onClick={() => setTab(key)}
              style={{ padding: "8px 18px", borderRadius: 7, background: tab === key ? "#20191c" : "transparent", color: tab === key ? "#FBF8F3" : "#5c5148", fontSize: 13.5 }}>
              {label}
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

const HORARIOS = [
  ["Lunes", "9 a.m.–8 p.m."],
  ["Martes", "9 a.m.–8 p.m."],
  ["Miércoles", "9 a.m.–8 p.m."],
  ["Jueves", "9 a.m.–8 p.m."],
  ["Viernes", "9 a.m.–8 p.m."],
  ["Sábado", "9 a.m.–8 p.m."],
  ["Domingo", "9 a.m.–2 p.m."],
];

function Footer({ whatsappNumber, whatsappMessage }) {
  const waHref = whatsappNumber
    ? `https://wa.me/${whatsappNumber.replace(/[^\d]/g, "")}${whatsappMessage ? `?text=${encodeURIComponent(whatsappMessage)}` : ""}`
    : null;

  return (
    <div style={{ borderTop: "1px solid #E7DFD3", marginTop: 40, background: "#F0E9DC" }}>
      <div style={{ maxWidth: 1100, margin: "0 auto", padding: "32px 24px", display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 24, fontSize: 13.5, color: "#3a322c" }}>
        <div>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 16, marginBottom: 8 }}>Dirección</div>
          <div style={{ color: "#5c5148" }}>1804, Canning, Provincia de Buenos Aires</div>
        </div>
        <div>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 16, marginBottom: 8 }}>Contacto</div>
          {waHref ? (
            <a href={waHref} target="_blank" rel="noopener noreferrer" style={{ color: "#5c5148", textDecoration: "none", display: "flex", alignItems: "center", gap: 6 }}>
              <MessageCircle size={15} color="#7C8A63" /> {whatsappNumber}
            </a>
          ) : (
            <span style={{ color: "#8a7f74" }}>Todavía no configurado</span>
          )}
        </div>
        <div>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 16, marginBottom: 8 }}>Horario</div>
          {HORARIOS.map(([dia, horas]) => (
            <div key={dia} style={{ display: "flex", justifyContent: "space-between", color: "#5c5148", maxWidth: 220 }}>
              <span>{dia}</span><span>{horas}</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}



function Store({ products, email, flash }) {
  const [cart, setCart] = useState({});
  const [showCheckout, setShowCheckout] = useState(false);
  const [query, setQuery] = useState("");
  const [sortBy, setSortBy] = useState("relevancia");
  const [onlyInStock, setOnlyInStock] = useState(false);
  const [marcaFilter, setMarcaFilter] = useState("");
  const [rubroFilter, setRubroFilter] = useState("");

  const marcas = useMemo(() => {
    const set = new Set(products.map((p) => p.marca).filter(Boolean));
    return Array.from(set).sort();
  }, [products]);

  const rubros = useMemo(() => {
    const set = new Set(products.map((p) => p.rubro).filter(Boolean));
    return Array.from(set).sort();
  }, [products]);

  const cartItems = useMemo(() => Object.entries(cart)
    .map(([id, qty]) => ({ product: products.find((p) => p.id === id), qty }))
    .filter((c) => c.product && c.qty > 0), [cart, products]);

  const total = cartItems.reduce((s, c) => s + c.product.price * c.qty, 0);
  const cartCount = cartItems.reduce((s, c) => s + c.qty, 0);

  const visibleProducts = useMemo(() => {
    const q = query.trim().toLowerCase();
    let list = products.filter((p) => {
      if (onlyInStock && p.stock <= 0) return false;
      if (marcaFilter && p.marca !== marcaFilter) return false;
      if (rubroFilter && p.rubro !== rubroFilter) return false;
      if (!q) return true;
      return p.sku.toLowerCase().includes(q) || p.description.toLowerCase().includes(q);
    });
    if (sortBy === "precio-asc") list = [...list].sort((a, b) => a.price - b.price);
    if (sortBy === "precio-desc") list = [...list].sort((a, b) => b.price - a.price);
    if (sortBy === "nombre-asc") list = [...list].sort((a, b) => a.description.localeCompare(b.description));
    if (sortBy === "nombre-desc") list = [...list].sort((a, b) => b.description.localeCompare(a.description));
    return list;
  }, [products, query, sortBy, onlyInStock, marcaFilter, rubroFilter]);

  function setQty(id, qty, stock) {
    const clamped = Math.max(0, Math.min(qty, stock));
    setCart((c) => ({ ...c, [id]: clamped }));
  }

  return (
    <div style={{ maxWidth: 1100, margin: "0 auto", padding: "36px 24px 120px" }}>
      <div style={{ marginBottom: 32 }}>
        <h1 style={{ fontFamily: "'Fraunces', serif", fontWeight: 400, fontSize: 34, margin: 0 }}>Catálogo</h1>
        <p style={{ color: "#8a7f74", marginTop: 6, fontSize: 14.5 }}>Elegí tus fragancias y armá el pedido. Te confirmamos por mail.</p>
      </div>

      {rubros.length > 0 && (
        <div style={{ display: "flex", gap: 8, overflowX: "auto", paddingBottom: 6, marginBottom: 18 }}>
          <button className="btn" onClick={() => setRubroFilter("")}
            style={{ padding: "8px 16px", borderRadius: 20, background: rubroFilter === "" ? "#20191c" : "#F0E9DC", color: rubroFilter === "" ? "#fff" : "#5c5148", fontSize: 13, whiteSpace: "nowrap", flexShrink: 0 }}>
            Todos los rubros
          </button>
          {rubros.map((r) => (
            <button key={r} className="btn" onClick={() => setRubroFilter(r)}
              style={{ padding: "8px 16px", borderRadius: 20, background: rubroFilter === r ? "#20191c" : "#F0E9DC", color: rubroFilter === r ? "#fff" : "#5c5148", fontSize: 13, whiteSpace: "nowrap", flexShrink: 0 }}>
              {r}
            </button>
          ))}
        </div>
      )}

      <div style={{ display: "flex", flexWrap: "wrap", gap: 10, marginBottom: 26, alignItems: "center" }}>
        <div style={{ position: "relative", flex: "1 1 240px", minWidth: 220 }}>
          <Search size={16} style={{ position: "absolute", left: 12, top: "50%", transform: "translateY(-50%)", color: "#8a7f74" }} />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Buscar por SKU o nombre…"
            style={{ width: "100%", padding: "11px 12px 11px 36px", borderRadius: 9, border: "1px solid #E7DFD3", fontSize: 14, background: "#fff" }} />
        </div>
        <select value={sortBy} onChange={(e) => setSortBy(e.target.value)}
          style={{ padding: "11px 12px", borderRadius: 9, border: "1px solid #E7DFD3", fontSize: 13.5, background: "#fff", color: "#3a322c" }}>
          <option value="relevancia">Orden: relevancia</option>
          <option value="precio-asc">Precio: menor a mayor</option>
          <option value="precio-desc">Precio: mayor a menor</option>
          <option value="nombre-asc">Nombre: A-Z</option>
          <option value="nombre-desc">Nombre: Z-A</option>
        </select>
        {marcas.length > 0 && (
          <select value={marcaFilter} onChange={(e) => setMarcaFilter(e.target.value)}
            style={{ padding: "11px 12px", borderRadius: 9, border: "1px solid #E7DFD3", fontSize: 13.5, background: "#fff", color: "#3a322c" }}>
            <option value="">Todas las marcas</option>
            {marcas.map((m) => <option key={m} value={m}>{m}</option>)}
          </select>
        )}
        <button className="btn" onClick={() => setOnlyInStock((v) => !v)}
          style={{ padding: "11px 16px", borderRadius: 9, background: onlyInStock ? "#20191c" : "#F0E9DC", color: onlyInStock ? "#fff" : "#5c5148", fontSize: 13.5, whiteSpace: "nowrap" }}>
          Solo con stock
        </button>
      </div>

      {products.length === 0 ? (
        <EmptyState text="Todavía no hay productos cargados. Pedile al admin que suba el Excel del catálogo." />
      ) : visibleProducts.length === 0 ? (
        <EmptyState text="No encontramos productos que coincidan con la búsqueda o los filtros." />
      ) : (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(380px, 1fr))", gap: 18 }}>
          {visibleProducts.map((p) => (
            <ProductCard key={p.id} product={p} qty={cart[p.id] || 0} onQty={(q) => setQty(p.id, q, p.stock)} />
          ))}
        </div>
      )}

      {cartCount > 0 && (
        <button className="btn" onClick={() => setShowCheckout(true)}
          style={{ position: "fixed", bottom: 24, right: 24, background: "#20191c", color: "#FBF8F3", borderRadius: 50, padding: "14px 24px", display: "flex", alignItems: "center", gap: 10, boxShadow: "0 10px 30px rgba(0,0,0,0.3)", fontSize: 14.5, zIndex: 90 }}>
          <ShoppingBag size={18} /> {cartCount} · {money(total)}
        </button>
      )}

      {showCheckout && (
        <Checkout cartItems={cartItems} total={total} email={email}
          onQty={setQty}
          onRemove={(id) => setQty(id, 0, 0)}
          onReset={() => { setCart({}); setShowCheckout(false); }}
          onClose={() => setShowCheckout(false)}
          onDone={() => { setCart({}); setShowCheckout(false); flash("Pedido listo. Revisá tu cliente de mail para enviarlo."); }} />
      )}
    </div>
  );
}

function ProductCard({ product, qty, onQty }) {
  const outOfStock = product.stock <= 0;
  return (
    <div style={{ background: "#fff", border: "1px solid #EEE6D9", borderRadius: 12, overflow: "hidden", display: "flex", height: 148 }}>
      <div style={{ position: "relative", width: 148, flexShrink: 0, background: "#F5F0E6" }}>
        {product.image ? (
          <img src={product.image} alt={product.sku} style={{ width: "100%", height: "100%", objectFit: "cover" }} />
        ) : (
          <div style={{ width: "100%", height: "100%", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <ImageIcon size={26} color="#C9BC9F" />
          </div>
        )}
        <span style={{ position: "absolute", bottom: 0, left: 0, right: 0, fontFamily: "'IBM Plex Mono', monospace", fontSize: 10, letterSpacing: 1.5, color: "#FBF8F3", background: "rgba(32,25,28,0.65)", padding: "3px 8px" }}>
          {product.sku}
        </span>
      </div>
      <div style={{ padding: "12px 14px", display: "flex", flexDirection: "column", gap: 6, flex: 1, minWidth: 0 }}>
        {product.marca && (
          <span style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: 1, textTransform: "uppercase", color: "#8A6D3B" }}>{product.marca}</span>
        )}
        <div style={{ fontSize: 13.5, lineHeight: 1.4, color: "#3a322c", overflow: "hidden", display: "-webkit-box", WebkitLineClamp: 2, WebkitBoxOrient: "vertical" }}>
          {product.description || "Sin descripción"}
        </div>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <span style={{ fontFamily: "'Fraunces', serif", fontSize: 19 }}>{money(product.price)}</span>
          <span style={{ fontSize: 11.5, color: outOfStock ? "#B4483C" : "#7C8A63", fontWeight: 600 }}>
            {outOfStock ? "Sin stock" : `Stock: ${product.stock}`}
          </span>
        </div>
        <div style={{ marginTop: "auto" }}>
          {outOfStock ? (
            <div style={{ textAlign: "center", fontSize: 12.5, color: "#8a7f74", padding: "8px 0" }}>No disponible</div>
          ) : (
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", background: "#F5F0E6", borderRadius: 8, padding: 4 }}>
              <button className="btn" onClick={() => onQty(qty - 1)} disabled={qty === 0} style={{ background: "transparent", width: 28, height: 28, borderRadius: 6, display: "flex", alignItems: "center", justifyContent: "center", opacity: qty === 0 ? 0.35 : 1 }}>
                <Minus size={13} />
              </button>
              <span style={{ fontWeight: 600, fontSize: 13.5 }}>{qty}</span>
              <button className="btn" onClick={() => onQty(qty + 1)} style={{ background: "transparent", width: 28, height: 28, borderRadius: 6, display: "flex", alignItems: "center", justifyContent: "center" }}>
                <Plus size={13} />
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function EmptyState({ text }) {
  return (
    <div style={{ border: "1px dashed #D9CDB8", borderRadius: 12, padding: "48px 24px", textAlign: "center", color: "#8a7f74" }}>
      <Package size={28} style={{ marginBottom: 10, opacity: 0.6 }} />
      <div style={{ fontSize: 14.5 }}>{text}</div>
    </div>
  );
}

function Checkout({ cartItems, total, email, onQty, onRemove, onReset, onClose, onDone }) {
  const [form, setForm] = useState({ nombre: "", apellido: "", direccion: "", tipoDoc: "DNI", documento: "", razonSocial: "", whatsapp: "", pago: "efectivo" });
  const [errors, setErrors] = useState({});
  const [mpStatus, setMpStatus] = useState("idle");

  function update(field, value) { setForm((f) => ({ ...f, [field]: value })); }

  function validate() {
    const e = {};
    if (!form.nombre.trim()) e.nombre = true;
    if (!form.apellido.trim()) e.apellido = true;
    if (!form.direccion.trim()) e.direccion = true;
    if (!form.documento.trim()) e.documento = true;
    if (form.tipoDoc === "CUIT" && !form.razonSocial.trim()) e.razonSocial = true;
    if (!form.whatsapp.trim()) e.whatsapp = true;
    setErrors(e);
    return Object.keys(e).length === 0;
  }

  async function saveOrderRecord(orderId, pago) {
    try {
      await api("/api/orders", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ...form, pago, orderId, items: cartItems.map((c) => ({ sku: c.product.sku, qty: c.qty, price: c.product.price })), total }),
      });
    } catch {}
  }

  async function payWithMercadoPago() {
    if (!validate()) return;
    setMpStatus("loading");
    const orderId = uid();
    try {
      const data = await api("/api/create-preference", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          items: cartItems.map((c) => ({ title: `${c.product.sku} - ${c.product.description}`, quantity: c.qty, unit_price: c.product.price })),
          payer: { name: `${form.nombre} ${form.apellido}` },
          externalReference: orderId,
        }),
      });
      await saveOrderRecord(orderId, "mercadopago");
      window.location.href = data.init_point;
    } catch {
      setMpStatus("error");
    }
  }

  async function submit() {
    if (!validate()) return;
    const orderId = uid();
    const lines = cartItems.map((c) => `${c.product.sku} x${c.qty} - ${c.product.description} - ${money(c.product.price)} c/u - subtotal ${money(c.product.price * c.qty)}`);
    const bodyLines = [
      `Cliente: ${form.nombre} ${form.apellido}`,
      `Dirección: ${form.direccion}`,
      `${form.tipoDoc}: ${form.documento}`,
      ...(form.tipoDoc === "CUIT" ? [`Razón social: ${form.razonSocial}`] : []),
      `WhatsApp para avisar retiro: ${form.whatsapp}`,
      `Forma de pago: Efectivo en local`,
      "",
      "Pedido:",
      ...lines,
      "",
      `Total: ${money(total)}`,
    ];
    const subject = encodeURIComponent(`Pedido web - ${form.nombre} ${form.apellido}`);
    const body = encodeURIComponent(bodyLines.join("\n"));
    const to = email && email.includes("@") ? email : "";

    await saveOrderRecord(orderId, "efectivo");

    const a = document.createElement("a");
    a.href = `mailto:${to}?subject=${subject}&body=${body}`;
    a.click();
    onDone();
  }

  const inputStyle = (field) => ({ width: "100%", padding: "11px 12px", borderRadius: 8, border: `1px solid ${errors[field] ? "#B4483C" : "#E7DFD3"}`, fontSize: 14, background: "#fff" });
  const label = { fontSize: 12.5, fontWeight: 600, color: "#5c5148", marginBottom: 5, display: "block" };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(32,25,28,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 150, padding: 20 }} onClick={onClose}>
      <div style={{ background: "#FBF8F3", borderRadius: 14, width: "100%", maxWidth: 480, maxHeight: "88vh", overflowY: "auto", padding: 26 }} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
          <h2 style={{ fontFamily: "'Fraunces', serif", fontWeight: 400, fontSize: 22, margin: 0 }}>Datos del pedido</h2>
          <button className="btn" onClick={onClose} style={{ background: "transparent", padding: 6 }}><X size={20} /></button>
        </div>

        <div style={{ background: "#F0E9DC", borderRadius: 10, padding: 14, marginBottom: 18 }}>
          {cartItems.length === 0 ? (
            <div style={{ fontSize: 13, color: "#8a7f74", textAlign: "center", padding: "10px 0" }}>El carrito quedó vacío.</div>
          ) : (
            cartItems.map((c) => (
              <div key={c.product.id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 0", borderBottom: "1px solid #E0D3B8" }}>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 13, fontWeight: 600, fontFamily: "'IBM Plex Mono', monospace" }}>{c.product.sku}</div>
                  <div style={{ fontSize: 11.5, color: "#8a7f74" }}>{money(c.product.price)} c/u</div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 4, background: "#fff", borderRadius: 7, padding: 3 }}>
                  <button className="btn" onClick={() => onQty(c.product.id, c.qty - 1, c.product.stock)} style={{ background: "transparent", width: 24, height: 24, borderRadius: 5, display: "flex", alignItems: "center", justifyContent: "center" }}>
                    <Minus size={12} />
                  </button>
                  <span style={{ fontSize: 12.5, fontWeight: 600, minWidth: 16, textAlign: "center" }}>{c.qty}</span>
                  <button className="btn" onClick={() => onQty(c.product.id, c.qty + 1, c.product.stock)} style={{ background: "transparent", width: 24, height: 24, borderRadius: 5, display: "flex", alignItems: "center", justifyContent: "center" }}>
                    <Plus size={12} />
                  </button>
                </div>
                <span style={{ fontSize: 13, fontWeight: 600, minWidth: 70, textAlign: "right" }}>{money(c.product.price * c.qty)}</span>
                <button className="btn" onClick={() => onRemove(c.product.id)} title="Quitar producto" style={{ background: "transparent", color: "#B4483C", width: 24, height: 24, borderRadius: 5, display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <X size={15} />
                </button>
              </div>
            ))
          )}
          <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 700, marginTop: 10, paddingTop: 8, borderTop: "1px solid #E0D3B8" }}>
            <span>Total</span><span>{money(total)}</span>
          </div>
        </div>

        <button className="btn" onClick={onReset} style={{ width: "100%", background: "transparent", color: "#8a7f74", border: "1px solid #E7DFD3", padding: "9px 0", borderRadius: 8, fontSize: 12.5, marginBottom: 18, display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}>
          <RotateCcw size={13} /> Empezar de nuevo
        </button>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 12 }}>
          <div><label style={label}>Nombre</label><input style={inputStyle("nombre")} value={form.nombre} onChange={(e) => update("nombre", e.target.value)} /></div>
          <div><label style={label}>Apellido</label><input style={inputStyle("apellido")} value={form.apellido} onChange={(e) => update("apellido", e.target.value)} /></div>
        </div>

        <div style={{ marginBottom: 12 }}><label style={label}>Dirección</label><input style={inputStyle("direccion")} value={form.direccion} onChange={(e) => update("direccion", e.target.value)} /></div>

        <div style={{ marginBottom: 12 }}>
          <label style={label}>Tipo de documento</label>
          <div style={{ display: "flex", gap: 8 }}>
            {["DNI", "CUIT"].map((t) => (
              <button key={t} className="btn" onClick={() => update("tipoDoc", t)}
                style={{ flex: 1, padding: "10px 0", borderRadius: 8, background: form.tipoDoc === t ? "#20191c" : "#F0E9DC", color: form.tipoDoc === t ? "#fff" : "#5c5148", fontSize: 13.5 }}>
                {t}
              </button>
            ))}
          </div>
        </div>

        <div style={{ marginBottom: 12 }}><label style={label}>{form.tipoDoc}</label><input style={inputStyle("documento")} value={form.documento} onChange={(e) => update("documento", e.target.value)} /></div>

        {form.tipoDoc === "CUIT" && (
          <div style={{ marginBottom: 12 }}><label style={label}>Razón social</label><input style={inputStyle("razonSocial")} value={form.razonSocial} onChange={(e) => update("razonSocial", e.target.value)} /></div>
        )}

        <div style={{ marginBottom: 12 }}>
          <label style={label}>WhatsApp para avisarte cuando esté listo</label>
          <input style={inputStyle("whatsapp")} value={form.whatsapp} onChange={(e) => update("whatsapp", e.target.value)} placeholder="Ej: 11 2345 6789" />
        </div>

        <div style={{ marginBottom: 18 }}>
          <label style={label}>Forma de pago</label>
          <div style={{ display: "flex", gap: 8 }}>
            {[["efectivo", "Efectivo en local"], ["mercadopago", "Mercado Pago"]].map(([val, lbl]) => (
              <button key={val} className="btn" onClick={() => { update("pago", val); setMpStatus("idle"); }}
                style={{ flex: 1, padding: "10px 8px", borderRadius: 8, background: form.pago === val ? "#20191c" : "#F0E9DC", color: form.pago === val ? "#fff" : "#5c5148", fontSize: 13 }}>
                {lbl}
              </button>
            ))}
          </div>
          {form.pago === "mercadopago" && mpStatus === "error" && (
            <div style={{ display: "flex", gap: 6, alignItems: "center", fontSize: 12, color: "#B4483C", marginTop: 8 }}>
              <AlertCircle size={13} /> No pudimos abrir Mercado Pago. Puede que la perfumería no lo haya conectado todavía.
            </div>
          )}
        </div>

        {Object.keys(errors).length > 0 && (
          <div style={{ display: "flex", gap: 6, alignItems: "center", color: "#B4483C", fontSize: 12.5, marginBottom: 10 }}>
            <AlertCircle size={14} /> Completá los campos obligatorios.
          </div>
        )}

        {form.pago === "mercadopago" ? (
          <button className="btn" onClick={payWithMercadoPago} disabled={cartItems.length === 0 || mpStatus === "loading"}
            style={{ width: "100%", background: "#009EE3", color: "#fff", padding: "13px 0", borderRadius: 9, fontSize: 14.5, display: "flex", alignItems: "center", justifyContent: "center", gap: 8, opacity: cartItems.length === 0 ? 0.4 : 1 }}>
            <CreditCard size={16} /> {mpStatus === "loading" ? "Abriendo Mercado Pago…" : "Pagar con Mercado Pago"}
          </button>
        ) : (
          <button className="btn" onClick={submit} disabled={cartItems.length === 0}
            style={{ width: "100%", background: "#20191c", color: "#fff", padding: "13px 0", borderRadius: 9, fontSize: 14.5, display: "flex", alignItems: "center", justifyContent: "center", gap: 8, opacity: cartItems.length === 0 ? 0.4 : 1 }}>
            <Mail size={16} /> Enviar pedido por mail
          </button>
        )}
      </div>
    </div>
  );
}

/* ---------------- ADMIN ---------------- */

function Offers() {
  const [status, setStatus] = useState("loading"); // loading | available | empty

  useEffect(() => {
    api("/api/offers").then((d) => setStatus(d.available ? "available" : "empty")).catch(() => setStatus("empty"));
  }, []);

  return (
    <div style={{ maxWidth: 1100, margin: "0 auto", padding: "36px 24px 80px" }}>
      <div style={{ marginBottom: 24 }}>
        <h1 style={{ fontFamily: "'Fraunces', serif", fontWeight: 400, fontSize: 34, margin: 0 }}>Ofertas</h1>
        <p style={{ color: "#8a7f74", marginTop: 6, fontSize: 14.5 }}>Nuestro folleto de promociones vigentes.</p>
      </div>

      {status === "loading" ? (
        <div style={{ padding: 40, textAlign: "center", color: "#8a7f74" }}>Cargando…</div>
      ) : status === "empty" ? (
        <EmptyState text="Todavía no hay ofertas cargadas." />
      ) : (
        <div style={{ border: "1px solid #EEE6D9", borderRadius: 12, overflow: "hidden", background: "#fff" }}>
          <iframe src="/api/offers.pdf" title="Ofertas" style={{ width: "100%", height: "80vh", border: "none", display: "block" }} />
          <div style={{ padding: 14, borderTop: "1px solid #EEE6D9", textAlign: "center" }}>
            <a href="/api/offers.pdf" target="_blank" rel="noopener noreferrer" style={{ color: "#8A6D3B", fontSize: 13.5, fontWeight: 600, textDecoration: "none" }}>
              Abrir en una pestaña nueva / descargar
            </a>
          </div>
        </div>
      )}
    </div>
  );
}

function AdminGate({ onSuccess }) {
  const [secret, setSecret] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  async function handleLogin() {
    setError("");
    if (!secret.trim()) { setError("Ingresá la clave."); return; }
    setBusy(true);
    try {
      await api("/api/admin/verify", { method: "POST", headers: { "x-admin-secret": secret.trim() } });
      onSuccess(secret.trim());
    } catch (err) {
      setError("Clave incorrecta.");
    }
    setBusy(false);
  }

  return (
    <div style={{ maxWidth: 380, margin: "60px auto", padding: "0 24px" }}>
      <div style={{ background: "#fff", border: "1px solid #EEE6D9", borderRadius: 14, padding: 28 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6, color: "#8A6D3B" }}>
          <Lock size={16} />
          <span style={{ fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: 1 }}>Acceso admin</span>
        </div>
        <h2 style={{ fontFamily: "'Fraunces', serif", fontWeight: 400, fontSize: 22, margin: "0 0 18px" }}>Ingresá tu clave</h2>
        <div style={{ marginBottom: 16 }}>
          <input type="password" value={secret} onChange={(e) => setSecret(e.target.value)} onKeyDown={(e) => e.key === "Enter" && handleLogin()}
            placeholder="ADMIN_SECRET"
            style={{ width: "100%", padding: "11px 12px", borderRadius: 8, border: "1px solid #E7DFD3", fontSize: 14 }} />
        </div>
        {error && (
          <div style={{ display: "flex", gap: 6, alignItems: "center", color: "#B4483C", fontSize: 12.5, marginBottom: 14 }}>
            <AlertCircle size={14} /> {error}
          </div>
        )}
        <button className="btn" onClick={handleLogin} disabled={busy}
          style={{ width: "100%", background: "#20191c", color: "#fff", padding: "12px 0", borderRadius: 9, fontSize: 14 }}>
          {busy ? "Verificando…" : "Entrar"}
        </button>
        <p style={{ fontSize: 11.5, color: "#8a7f74", marginTop: 14, lineHeight: 1.5 }}>
          Es la misma clave <code>ADMIN_SECRET</code> que configuraste al desplegar el servidor. Sesión válida hasta que recargues la página.
        </p>
      </div>
    </div>
  );
}

function Admin({ email, whatsappNumber, whatsappMessage, setConfig, adminSecret, onCatalogChanged, flash }) {
  const [products, setProductsState] = useState([]);
  const [productsLoaded, setProductsLoaded] = useState(false);
  const [preview, setPreview] = useState(null);
  const [mpToken, setMpToken] = useState("");
  const [mpStatus, setMpStatus] = useState("checking");
  const [orders, setOrders] = useState(null);
  const [offersStatus, setOffersStatus] = useState("checking");
  const [offersBusy, setOffersBusy] = useState(false);
  const fileRef = useRef();
  const offersFileRef = useRef();

  useEffect(() => {
    api("/api/admin/products", { headers: { "x-admin-secret": adminSecret } })
      .then((p) => { setProductsState(p); setProductsLoaded(true); })
      .catch(() => flash("No se pudo cargar el catálogo."));
    api("/api/mp-status").then((d) => setMpStatus(d.connected ? "connected" : "disconnected")).catch(() => setMpStatus("unknown"));
    api("/api/offers").then((d) => setOffersStatus(d.available ? "available" : "empty")).catch(() => setOffersStatus("empty"));
  }, []);

  async function uploadOffers(file) {
    setOffersBusy(true);
    try {
      const bytes = await file.arrayBuffer();
      await api("/api/admin/upload-offers", {
        method: "POST",
        headers: { "Content-Type": "application/pdf", "x-admin-secret": adminSecret },
        body: bytes,
      });
      setOffersStatus("available");
      flash("PDF de ofertas actualizado.");
    } catch {
      flash("No se pudo subir el PDF.");
    }
    setOffersBusy(false);
  }

  async function deleteOffers() {
    setOffersBusy(true);
    try {
      await api("/api/admin/delete-offers", { method: "POST", headers: { "x-admin-secret": adminSecret } });
      setOffersStatus("empty");
      flash("Ofertas eliminadas.");
    } catch {
      flash("No se pudo borrar el PDF.");
    }
    setOffersBusy(false);
  }

  async function setProducts(next) {
    setProductsState(next);
    try {
      await api("/api/admin/products", {
        method: "PUT",
        headers: { "Content-Type": "application/json", "x-admin-secret": adminSecret },
        body: JSON.stringify(next),
      });
      onCatalogChanged();
    } catch {
      flash("No se pudo guardar el catálogo. Reintentá.");
    }
  }

  async function loadOrders() {
    try {
      const data = await api("/api/admin/orders", { headers: { "x-admin-secret": adminSecret } });
      setOrders(data);
    } catch {
      flash("No se pudieron cargar los pedidos.");
    }
  }

  async function connectMp() {
    if (!mpToken.trim()) { flash("Pegá tu Access Token."); return; }
    try {
      await api("/api/admin/connect-mercadopago", {
        method: "POST",
        headers: { "Content-Type": "application/json", "x-admin-secret": adminSecret },
        body: JSON.stringify({ accessToken: mpToken.trim() }),
      });
      setMpStatus("connected");
      setMpToken("");
      flash("Mercado Pago conectado.");
    } catch {
      flash("No se pudo conectar. Revisá el Access Token.");
    }
  }

  async function disconnectMp() {
    try {
      await api("/api/admin/disconnect-mercadopago", { method: "POST", headers: { "x-admin-secret": adminSecret } });
      setMpStatus("disconnected");
      flash("Mercado Pago desconectado.");
    } catch {
      flash("No se pudo desconectar.");
    }
  }

  async function handleExcel(e) {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const parsed = await parseExcel(file);
      if (parsed.length === 0) { flash("No se encontraron filas válidas (revisá columnas SKU, Stock, Precio, Descripcion)."); return; }
      setPreview(parsed);
    } catch {
      flash("No pude leer el archivo. ¿Es un .xlsx válido?");
    }
    e.target.value = "";
  }

  async function attachImage(idx, file) {
    const dataUrl = await fileToDataURL(file);
    setPreview((prev) => prev.map((p, i) => (i === idx ? { ...p, image: dataUrl } : p)));
  }

  function confirmImport() {
    const importedBySku = new Map(preview.map((p) => [p.sku, p]));
    const seenSkus = new Set();

    const updatedExisting = products.map((p) => {
      const imported = importedBySku.get(p.sku);
      if (imported) {
        seenSkus.add(p.sku);
        return { ...p, ...imported, image: imported.image || p.image || null, status: "activo" };
      }
      return { ...p, status: "suspendido" };
    });

    const newOnes = preview.filter((p) => !seenSkus.has(p.sku)).map((p) => ({ ...p, status: "activo" }));

    const next = [...updatedExisting, ...newOnes];
    setProducts(next);
    setPreview(null);
    const suspendidos = next.filter((p) => p.status === "suspendido").length;
    flash(`Catálogo actualizado. ${suspendidos} producto(s) quedaron suspendidos por no estar en este Excel.`);
  }

  function updateProduct(id, field, value) {
    setProducts(products.map((p) => (p.id === id ? { ...p, [field]: field === "stock" || field === "price" ? Number(value) || 0 : value } : p)));
  }

  async function updateImage(id, file) {
    const dataUrl = await fileToDataURL(file);
    setProducts(products.map((p) => (p.id === id ? { ...p, image: dataUrl } : p)));
  }

  function removeProduct(id) { setProducts(products.filter((p) => p.id !== id)); }
  function addBlank() { setProducts([...products, emptyProduct()]); }

  return (
    <div style={{ maxWidth: 1000, margin: "0 auto", padding: "36px 24px 80px" }}>
      <h1 style={{ fontFamily: "'Fraunces', serif", fontWeight: 400, fontSize: 30, marginBottom: 6 }}>Panel admin</h1>
      <p style={{ color: "#8a7f74", fontSize: 14, marginBottom: 28 }}>Cargá el catálogo, gestioná pedidos y conectá Mercado Pago.</p>

      <Section icon={<Mail size={16} />} title="Mail de destino de los pedidos">
        <input value={email} onChange={(e) => setConfig({ email: e.target.value })} placeholder="pedidos@perfumeria.com"
          style={{ width: "100%", maxWidth: 380, padding: "11px 12px", borderRadius: 8, border: "1px solid #E7DFD3", fontSize: 14 }} />
      </Section>

      <Section icon={<MessageCircle size={16} />} title="WhatsApp de contacto">
        <p style={{ fontSize: 12.5, color: "#8a7f74", marginTop: -6, marginBottom: 12 }}>
          Este número aparece en el pie de la tienda como link directo a WhatsApp. Usá el formato internacional sin espacios ni signos (ej: 5491168114857) para que el link abra correctamente.
        </p>
        <div style={{ display: "flex", flexDirection: "column", gap: 10, maxWidth: 420 }}>
          <input value={whatsappNumber} onChange={(e) => setConfig({ whatsappNumber: e.target.value })} placeholder="5491168114857"
            style={{ padding: "11px 12px", borderRadius: 8, border: "1px solid #E7DFD3", fontSize: 14 }} />
          <textarea value={whatsappMessage} onChange={(e) => setConfig({ whatsappMessage: e.target.value })} placeholder="Mensaje pre-cargado, ej: Hola! Quería consultar por..." rows={2}
            style={{ padding: "11px 12px", borderRadius: 8, border: "1px solid #E7DFD3", fontSize: 14, resize: "vertical" }} />
        </div>
      </Section>

      <Section icon={<FileText size={16} />} title="Ofertas (PDF)">
        <p style={{ fontSize: 12.5, color: "#8a7f74", marginTop: -6, marginBottom: 12 }}>
          Se muestra en la pestaña "Ofertas" de la tienda. Subir un PDF nuevo reemplaza al anterior.
        </p>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <input ref={offersFileRef} type="file" accept="application/pdf" style={{ display: "none" }}
            onChange={(e) => e.target.files[0] && uploadOffers(e.target.files[0])} />
          <button className="btn" onClick={() => offersFileRef.current.click()} disabled={offersBusy}
            style={{ background: "#20191c", color: "#fff", padding: "10px 18px", borderRadius: 8, fontSize: 13.5, display: "flex", alignItems: "center", gap: 8 }}>
            <Upload size={15} /> {offersStatus === "available" ? "Reemplazar PDF" : "Subir PDF"}
          </button>
          {offersStatus === "available" && (
            <button className="btn" onClick={deleteOffers} disabled={offersBusy}
              style={{ background: "transparent", color: "#B4483C", border: "1px solid #E7DFD3", padding: "10px 18px", borderRadius: 8, fontSize: 13.5 }}>
              Quitar ofertas
            </button>
          )}
          {offersStatus === "available" && <Check size={16} color="#7C8A63" />}
        </div>
      </Section>

      <Section icon={<CreditCard size={16} />} title="Mercado Pago">
        <div style={{ marginBottom: 14 }}>
          {mpStatus === "connected" && <span style={{ display: "flex", alignItems: "center", gap: 6, color: "#0F6E56", fontSize: 13, fontWeight: 600 }}><Check size={15} /> Conectado</span>}
          {mpStatus === "disconnected" && <span style={{ display: "flex", alignItems: "center", gap: 6, color: "#8A6D3B", fontSize: 13, fontWeight: 600 }}><AlertCircle size={15} /> Todavía no conectado</span>}
          {mpStatus === "checking" && <span style={{ fontSize: 13, color: "#8a7f74" }}>Verificando…</span>}
        </div>
        <div style={{ marginBottom: 12 }}>
          <input type="password" value={mpToken} onChange={(e) => setMpToken(e.target.value)} placeholder="Access Token de producción (APP_USR-...)"
            style={{ width: "100%", maxWidth: 420, padding: "11px 12px", borderRadius: 8, border: "1px solid #E7DFD3", fontSize: 14 }} />
        </div>
        <div style={{ display: "flex", gap: 10 }}>
          <button className="btn" onClick={connectMp} style={{ background: "#009EE3", color: "#fff", padding: "10px 18px", borderRadius: 8, fontSize: 13.5 }}>Conectar</button>
          {mpStatus === "connected" && (
            <button className="btn" onClick={disconnectMp} style={{ background: "transparent", color: "#B4483C", border: "1px solid #E7DFD3", padding: "10px 18px", borderRadius: 8, fontSize: 13.5 }}>Desconectar</button>
          )}
        </div>
      </Section>

      <Section icon={<Upload size={16} />} title="Cargar catálogo desde Excel">
        <p style={{ fontSize: 12.5, color: "#8a7f74", marginTop: -6, marginBottom: 12 }}>
          Columnas reconocidas: <strong>Codigo</strong> (o SKU), <strong>Stock</strong>, <strong>Lista1</strong> (o Precio), <strong>Descripcion</strong>, <strong>Marca</strong>, <strong>Rubro</strong>. Al confirmar: los productos del Excel se actualizan (o se crean si son nuevos), y los que ya estaban cargados pero <strong>no</strong> aparecen en este Excel pasan a <strong>suspendidos</strong> — no se borran, solo dejan de verse en la tienda hasta que vuelvan a aparecer en un Excel futuro.
        </p>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleExcel} style={{ display: "none" }} />
        <button className="btn" onClick={() => fileRef.current.click()} style={{ background: "#20191c", color: "#fff", padding: "10px 18px", borderRadius: 8, fontSize: 13.5, display: "flex", alignItems: "center", gap: 8 }}>
          <Upload size={15} /> Elegir archivo .xlsx
        </button>
      </Section>

      {preview && (
        <Section icon={<Check size={16} />} title={`Previsualización (${preview.length} filas)`}>
          <div style={{ display: "flex", gap: 6, alignItems: "center", color: "#8A6D3B", fontSize: 12.5, marginTop: -4, marginBottom: 12 }}>
            <AlertCircle size={14} /> Los productos que no estén en este Excel van a quedar suspendidos (no se borran, ni pierden su foto).
          </div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead>
                <tr style={{ textAlign: "left", color: "#8a7f74", fontSize: 11.5, textTransform: "uppercase" }}>
                  <th style={{ padding: "6px 8px" }}>Foto</th><th style={{ padding: "6px 8px" }}>SKU</th><th style={{ padding: "6px 8px" }}>Marca</th><th style={{ padding: "6px 8px" }}>Rubro</th><th style={{ padding: "6px 8px" }}>Stock</th><th style={{ padding: "6px 8px" }}>Precio</th><th style={{ padding: "6px 8px" }}>Descripción</th>
                </tr>
              </thead>
              <tbody>
                {preview.map((p, idx) => (
                  <tr key={p.id} style={{ borderTop: "1px solid #EEE6D9" }}>
                    <td style={{ padding: "8px" }}><ImagePicker image={p.image} onPick={(f) => attachImage(idx, f)} /></td>
                    <td style={{ padding: "8px", fontFamily: "'IBM Plex Mono', monospace" }}>{p.sku}</td>
                    <td style={{ padding: "8px" }}>{p.marca}</td>
                    <td style={{ padding: "8px", color: "#8a7f74" }}>{p.rubro}</td>
                    <td style={{ padding: "8px" }}>{p.stock}</td>
                    <td style={{ padding: "8px" }}>{money(p.price)}</td>
                    <td style={{ padding: "8px", maxWidth: 260 }}>{p.description}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
            <button className="btn" onClick={confirmImport} style={{ background: "#20191c", color: "#fff", padding: "10px 18px", borderRadius: 8, fontSize: 13.5 }}>Actualizar catálogo</button>
            <button className="btn" onClick={() => setPreview(null)} style={{ background: "#F0E9DC", color: "#5c5148", padding: "10px 18px", borderRadius: 8, fontSize: 13.5 }}>Cancelar</button>
          </div>
        </Section>
      )}

      <Section icon={<Package size={16} />} title={`Catálogo actual (${products.length})`}>
        {!productsLoaded ? (
          <div style={{ padding: 20, textAlign: "center", color: "#8a7f74", fontSize: 13 }}>Cargando…</div>
        ) : products.length === 0 ? <EmptyState text="Sin productos todavía." /> : (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            <div style={{ fontSize: 11.5, color: "#8a7f74", marginBottom: -2 }}>Marca y Rubro se usan como filtros de categoría en la tienda. Un producto sin precio, o marcado como suspendido, no aparece.</div>
            {products.map((p) => {
              const suspendido = p.status === "suspendido";
              const sinPrecio = !(Number(p.price) > 0);
              return (
              <div key={p.id} style={{ display: "grid", gridTemplateColumns: "56px 100px 100px 100px 80px 100px 1fr 32px 32px", gap: 10, alignItems: "center", background: "#fff", border: `1px solid ${suspendido ? "#D9CDB8" : "#EEE6D9"}`, borderRadius: 10, padding: 10, opacity: suspendido ? 0.55 : 1 }}>
                <ImagePicker image={p.image} onPick={(f) => updateImage(p.id, f)} small />
                <input value={p.sku} onChange={(e) => updateProduct(p.id, "sku", e.target.value)} placeholder="SKU" style={{ padding: "7px 8px", borderRadius: 6, border: "1px solid #E7DFD3", fontSize: 12.5, fontFamily: "'IBM Plex Mono', monospace" }} />
                <input value={p.marca || ""} onChange={(e) => updateProduct(p.id, "marca", e.target.value)} placeholder="Marca" style={{ padding: "7px 8px", borderRadius: 6, border: "1px solid #E7DFD3", fontSize: 12.5 }} />
                <input value={p.rubro || ""} onChange={(e) => updateProduct(p.id, "rubro", e.target.value)} placeholder="Rubro" style={{ padding: "7px 8px", borderRadius: 6, border: "1px solid #E7DFD3", fontSize: 12.5 }} />
                <input type="number" value={p.stock} onChange={(e) => updateProduct(p.id, "stock", e.target.value)} placeholder="Stock" style={{ padding: "7px 8px", borderRadius: 6, border: "1px solid #E7DFD3", fontSize: 12.5 }} />
                <input type="number" value={p.price} onChange={(e) => updateProduct(p.id, "price", e.target.value)} placeholder="Precio" style={{ padding: "7px 8px", borderRadius: 6, border: `1px solid ${sinPrecio ? "#B4483C" : "#E7DFD3"}`, fontSize: 12.5 }} />
                <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
                  <input value={p.description} onChange={(e) => updateProduct(p.id, "description", e.target.value)} placeholder="Descripción" style={{ padding: "7px 8px", borderRadius: 6, border: "1px solid #E7DFD3", fontSize: 12.5 }} />
                  {(suspendido || sinPrecio) && (
                    <span style={{ fontSize: 10.5, color: "#B4483C" }}>
                      {suspendido ? "Suspendido — no visible en la tienda" : "Sin precio — no visible en la tienda"}
                    </span>
                  )}
                </div>
                <button className="btn" onClick={() => updateProduct(p.id, "status", suspendido ? "activo" : "suspendido")} title={suspendido ? "Activar" : "Suspender"}
                  style={{ background: "transparent", color: suspendido ? "#7C8A63" : "#5c5148", padding: 6, display: "flex", justifyContent: "center" }}>
                  {suspendido ? <Eye size={16} /> : <EyeOff size={16} />}
                </button>
                <button className="btn" onClick={() => removeProduct(p.id)} title="Borrar definitivamente" style={{ background: "transparent", color: "#B4483C", padding: 6, display: "flex", justifyContent: "center" }}><Trash2 size={16} /></button>
              </div>
              );
            })}
          </div>
        )}
        <button className="btn" onClick={addBlank} style={{ marginTop: 14, background: "#F0E9DC", color: "#5c5148", padding: "9px 16px", borderRadius: 8, fontSize: 13, display: "flex", alignItems: "center", gap: 6 }}>
          <Plus size={14} /> Agregar producto manual
        </button>
      </Section>

      <Section icon={<ListOrdered size={16} />} title="Pedidos recibidos">
        {orders === null ? (
          <button className="btn" onClick={loadOrders} style={{ background: "#F0E9DC", color: "#5c5148", padding: "9px 16px", borderRadius: 8, fontSize: 13 }}>Cargar pedidos</button>
        ) : orders.length === 0 ? (
          <EmptyState text="Todavía no llegó ningún pedido." />
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {orders.map((o) => (
              <div key={o.id} style={{ background: "#fff", border: "1px solid #EEE6D9", borderRadius: 10, padding: 12, fontSize: 13 }}>
                <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 600 }}>
                  <span>{o.nombre} {o.apellido}</span>
                  <span>{money(o.total)}</span>
                </div>
                <div style={{ color: "#8a7f74", fontSize: 12, marginTop: 2 }}>
                  {o.tipoDoc}: {o.documento} · WhatsApp: {o.whatsapp} · Pago: {o.pago === "mercadopago" ? "Mercado Pago" : "Efectivo"} · {new Date(o.createdAt).toLocaleString("es-AR")}
                </div>
              </div>
            ))}
          </div>
        )}
      </Section>
    </div>
  );
}

function Section({ icon, title, children }) {
  return (
    <div style={{ marginBottom: 26 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12, color: "#5c5148" }}>{icon}<h3 style={{ fontSize: 14.5, margin: 0, fontWeight: 600 }}>{title}</h3></div>
      <div style={{ background: "#FBF8F3", border: "1px solid #EEE6D9", borderRadius: 12, padding: 18 }}>{children}</div>
    </div>
  );
}

function ImagePicker({ image, onPick, small }) {
  const ref = useRef();
  const size = small ? 40 : 44;
  return (
    <>
      <input ref={ref} type="file" accept="image/*" style={{ display: "none" }} onChange={(e) => e.target.files[0] && onPick(e.target.files[0])} />
      <button className="btn" onClick={() => ref.current.click()}
        style={{ width: size, height: size, borderRadius: 8, background: "#F0E9DC", border: "1px solid #E7DFD3", display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden", padding: 0 }}>
        {image ? <img src={image} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : <ImageIcon size={16} color="#8A6D3B" />}
      </button>
    </>
  );
}
