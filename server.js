// Servidor Express para Render.
// Crea una preferencia de pago en TU cuenta de Mercado Pago y devuelve
// la URL de checkout real (init_point) para redirigir al cliente.
//
// El Access Token de Mercado Pago NO se pega en Render ni en el código:
// se conecta una sola vez desde el panel Admin de la tienda, y este
// servidor lo guarda de forma privada en su propio disco (data/config.json).
//
// DEPLOY (una sola vez):
// 1) Subí esta carpeta a un repo de GitHub.
// 2) En render.com: New > Web Service > conectá el repo.
//    - Build command: npm install
//    - Start command: npm start
// 3) En Render, pestaña Environment, agregá SOLO esta variable:
//      ADMIN_SECRET = una clave que inventes vos (tipo una contraseña larga)
//    Esta es tu "llave maestra": la vas a usar una sola vez desde el Admin
//    de la tienda para autorizar la conexión con Mercado Pago.
// 4) Render te da una URL, por ejemplo: https://tu-servicio.onrender.com
//    Pegá esa URL (sin nada más al final) en el panel Admin de la tienda,
//    en "Checkout de Mercado Pago".
// 5) En el Admin de la tienda: pegá esa misma ADMIN_SECRET y tu Access Token
//    de Mercado Pago (Mercado Pago > Tu negocio > Credenciales de producción),
//    y tocá "Conectar". De ahí en más queda guardado en el servidor.
//
// Nota: en el plan free de Render el disco no está garantizado entre
// redeploys (si volvés a desplegar, puede que tengas que reconectar).
// Para persistencia sólida a largo plazo conviene el plan con disco persistente.
// Nota 2: en el plan free el servicio "duerme" tras un rato sin uso,
// así que la primera petición después de estar inactivo puede tardar ~30s.

import express from "express";
import cors from "cors";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const app = express();
app.use(cors());
app.use(express.json());

const PORT = process.env.PORT || 3000;
const ADMIN_SECRET = process.env.ADMIN_SECRET; // vos la elegís al desplegar, es tu "llave maestra"

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CONFIG_PATH = path.join(__dirname, "data", "config.json");

function loadConfig() {
  try {
    return JSON.parse(fs.readFileSync(CONFIG_PATH, "utf8"));
  } catch {
    return {};
  }
}

function saveConfig(cfg) {
  fs.mkdirSync(path.dirname(CONFIG_PATH), { recursive: true });
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(cfg, null, 2));
}

let config = loadConfig();
// Compatibilidad: si venís de la versión anterior con MP_ACCESS_TOKEN en variables de entorno.
if (!config.mpAccessToken && process.env.MP_ACCESS_TOKEN) {
  config.mpAccessToken = process.env.MP_ACCESS_TOKEN;
}

app.get("/", (req, res) => {
  res.send("Backend de Mercado Pago funcionando. Endpoint: POST /api/create-preference");
});

app.get("/api/mp-status", (req, res) => {
  res.json({ connected: !!config.mpAccessToken });
});

app.post("/api/connect-mercadopago", (req, res) => {
  if (!ADMIN_SECRET) {
    return res.status(500).json({ error: "El servidor todavía no tiene configurado ADMIN_SECRET" });
  }
  const { secret, accessToken } = req.body;
  if (secret !== ADMIN_SECRET) {
    return res.status(401).json({ error: "Clave incorrecta" });
  }
  if (!accessToken || !accessToken.trim()) {
    return res.status(400).json({ error: "Falta el Access Token" });
  }
  config.mpAccessToken = accessToken.trim();
  saveConfig(config);
  res.json({ connected: true });
});

app.post("/api/disconnect-mercadopago", (req, res) => {
  const { secret } = req.body;
  if (secret !== ADMIN_SECRET) {
    return res.status(401).json({ error: "Clave incorrecta" });
  }
  delete config.mpAccessToken;
  saveConfig(config);
  res.json({ connected: false });
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

    const origin = req.headers.origin || "https://www.mercadopago.com.ar";

    const preference = {
      items: items.map((it) => ({
        title: String(it.title).slice(0, 200),
        quantity: Number(it.quantity) || 1,
        unit_price: Number(it.unit_price) || 0,
        currency_id: "ARS",
      })),
      payer: payer ? { name: payer.name, email: payer.email || undefined } : undefined,
      external_reference: externalReference || undefined,
      back_urls: { success: origin, pending: origin, failure: origin },
      auto_return: "approved",
    };

    const mpRes = await fetch("https://api.mercadopago.com/checkout/preferences", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${ACCESS_TOKEN}`,
      },
      body: JSON.stringify(preference),
    });

    const data = await mpRes.json();

    if (!mpRes.ok) {
      return res.status(mpRes.status).json({ error: data.message || "Error creando la preferencia" });
    }

    // init_point = URL real de checkout de Mercado Pago para redirigir al cliente.
    return res.status(200).json({ init_point: data.init_point, id: data.id });
  } catch (err) {
    return res.status(500).json({ error: "Error inesperado creando la preferencia" });
  }
});

app.listen(PORT, () => {
  console.log(`Backend de Mercado Pago escuchando en el puerto ${PORT}`);
});
