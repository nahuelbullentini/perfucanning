// Backend único: sirve la tienda (build de Vite) y la API.
// Todo vive en esta misma URL de Render — no hace falta configurar
// ninguna otra URL desde el Admin.
//
// DEPLOY EN RENDER:
// Build Command: npm install && npm run build
// Start Command: npm start
// Variable de entorno: ADMIN_SECRET = una clave que inventes vos (tu llave maestra
// para entrar al Admin y conectar Mercado Pago).
//
// Nota: en el plan free, el disco de Render no está garantizado entre redeploys.
// Si redeployás, puede que tengas que volver a cargar el catálogo y reconectar MP.

import express from "express";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import * as XLSX from "xlsx";

const app = express();
app.use(express.json({ limit: "20mb" })); // las fotos van como base64, pueden pesar
app.use("/api/admin/import-excel", express.raw({ type: "*/*", limit: "20mb" }));

const PORT = process.env.PORT || 3000;
const ADMIN_SECRET = process.env.ADMIN_SECRET;

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DATA_DIR = path.join(__dirname, "data");
const CONFIG_PATH = path.join(DATA_DIR, "config.json");
const PRODUCTS_PATH = path.join(DATA_DIR, "products.json");
const ORDERS_PATH = path.join(DATA_DIR, "orders.json");

function readJSON(p, fallback) {
  try {
    return JSON.parse(fs.readFileSync(p, "utf8"));
  } catch {
    return fallback;
  }
}
function writeJSON(p, data) {
  fs.mkdirSync(path.dirname(p), { recursive: true });
  fs.writeFileSync(p, JSON.stringify(data, null, 2));
}

const SAMPLE_PRODUCTS = [
  { id: "pf-001", sku: "PF-001", stock: 12, price: 24500, description: "Eau de Parfum floral amaderado, 100ml", image: "https://images.unsplash.com/photo-1541643600914-78b084683601?w=500&q=80" },
  { id: "pf-002", sku: "PF-002", stock: 5, price: 31900, description: "Colonia cítrica fresca, 90ml", image: "https://images.unsplash.com/photo-1592945403244-b3fbafd7f539?w=500&q=80" },
  { id: "pf-003", sku: "PF-003", stock: 0, price: 18700, description: "Eau de Toilette oriental especiado, 75ml", image: "https://images.unsplash.com/photo-1523293182086-7651a899d37f?w=500&q=80" },
  { id: "pf-004", sku: "PF-004", stock: 20, price: 15300, description: "Body splash vainilla y almizcle, 200ml", image: "https://images.unsplash.com/photo-1615634260167-c8cdede054de?w=500&q=80" },
];

let config = readJSON(CONFIG_PATH, {});
let products = readJSON(PRODUCTS_PATH, null);
let orders = readJSON(ORDERS_PATH, []);

if (!products) {
  products = SAMPLE_PRODUCTS;
  writeJSON(PRODUCTS_PATH, products);
}

function requireAdmin(req, res, next) {
  if (!ADMIN_SECRET) {
    return res.status(500).json({ error: "Falta configurar ADMIN_SECRET en el servidor" });
  }
  if (req.headers["x-admin-secret"] !== ADMIN_SECRET) {
    return res.status(401).json({ error: "Clave incorrecta" });
  }
  next();
}

/* ---------- Público ---------- */

app.get("/api/products", (req, res) => res.json(products));

app.get("/api/config", (req, res) => res.json({ email: config.email || "" }));

app.post("/api/orders", (req, res) => {
  const order = {
    ...req.body,
    id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
    createdAt: new Date().toISOString(),
  };
  orders.unshift(order);
  orders = orders.slice(0, 500);
  writeJSON(ORDERS_PATH, orders);
  res.json({ ok: true });
});

app.post("/api/create-preference", async (req, res) => {
  const ACCESS_TOKEN = config.mpAccessToken;
  if (!ACCESS_TOKEN) {
    return res.status(500).json({ error: "Todavía no conectaste Mercado Pago desde el Admin" });
  }
  try {
    const { items, payer, externalReference } = req.body;
    if (!Array.isArray(items) || items.length === 0) {
      return res.status(400).json({ error: "El pedido no tiene items" });
    }
    const origin = req.headers.origin || `${req.protocol}://${req.get("host")}`;
    const preference = {
      items: items.map((it) => ({
        title: String(it.title).slice(0, 200),
        quantity: Number(it.quantity) || 1,
        unit_price: Number(it.unit_price) || 0,
        currency_id: "ARS",
      })),
      payer: payer ? { name: payer.name } : undefined,
      external_reference: externalReference || undefined,
      back_urls: { success: origin, pending: origin, failure: origin },
      auto_return: "approved",
    };
    const mpRes = await fetch("https://api.mercadopago.com/checkout/preferences", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${ACCESS_TOKEN}` },
      body: JSON.stringify(preference),
    });
    const data = await mpRes.json();
    if (!mpRes.ok) {
      return res.status(mpRes.status).json({ error: data.message || "Error creando la preferencia" });
    }
    res.json({ init_point: data.init_point, id: data.id });
  } catch {
    res.status(500).json({ error: "Error inesperado creando la preferencia" });
  }
});

app.get("/api/mp-status", (req, res) => res.json({ connected: !!config.mpAccessToken }));

function normalizeHeader(h) {
  return String(h || "").trim().toLowerCase()
    .replace(/[áàä]/g, "a").replace(/[éèë]/g, "e").replace(/[íìï]/g, "i").replace(/[óòö]/g, "o").replace(/[úùü]/g, "u");
}

app.post("/api/admin/import-excel", requireAdmin, (req, res) => {
  try {
    const wb = XLSX.read(req.body, { type: "buffer" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    const prevBySku = new Map(products.map((p) => [p.sku, p]));

    const parsed = rows.map((row) => {
      const map = {};
      Object.keys(row).forEach((k) => { map[normalizeHeader(k)] = row[k]; });
      const sku = String(map["sku"] ?? map["codigo"] ?? map["cod"] ?? "").trim();
      return {
        id: prevBySku.get(sku)?.id || (Date.now().toString(36) + Math.random().toString(36).slice(2, 6)),
        sku,
        stock: Number(map["stock"] ?? 0) || 0,
        price: Number(map["precio"] ?? map["lista1"] ?? map["preciolista1"] ?? 0) || 0,
        description: String(map["descripcion"] ?? "").trim(),
        image: prevBySku.get(sku)?.image || null,
      };
    }).filter((p) => p.sku);

    if (parsed.length === 0) {
      return res.status(400).json({ error: "No se encontraron filas válidas en el Excel" });
    }

    products = parsed;
    writeJSON(PRODUCTS_PATH, products);
    res.json({ ok: true, count: products.length });
  } catch (err) {
    res.status(500).json({ error: "No se pudo leer el Excel" });
  }
});

app.post("/api/admin/verify", requireAdmin, (req, res) => res.json({ ok: true }));

app.put("/api/admin/products", requireAdmin, (req, res) => {
  if (!Array.isArray(req.body)) return res.status(400).json({ error: "Formato inválido" });
  products = req.body;
  writeJSON(PRODUCTS_PATH, products);
  res.json({ ok: true });
});

app.put("/api/admin/config", requireAdmin, (req, res) => {
  config.email = req.body.email || "";
  writeJSON(CONFIG_PATH, config);
  res.json({ ok: true });
});

app.get("/api/admin/orders", requireAdmin, (req, res) => res.json(orders));

app.post("/api/admin/connect-mercadopago", requireAdmin, (req, res) => {
  const { accessToken } = req.body;
  if (!accessToken || !accessToken.trim()) {
    return res.status(400).json({ error: "Falta el Access Token" });
  }
  config.mpAccessToken = accessToken.trim();
  writeJSON(CONFIG_PATH, config);
  res.json({ connected: true });
});

app.post("/api/admin/disconnect-mercadopago", requireAdmin, (req, res) => {
  delete config.mpAccessToken;
  writeJSON(CONFIG_PATH, config);
  res.json({ connected: false });
});

/* ---------- Servir el frontend ---------- */

const DIST_DIR = path.join(__dirname, "dist");
app.use(express.static(DIST_DIR));
app.get("*", (req, res) => res.sendFile(path.join(DIST_DIR, "index.html")));

app.listen(PORT, () => console.log(`Perfumería web escuchando en el puerto ${PORT}`));
