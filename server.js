// ============================================
// BOT WHATSAPP - VERSÃO FINAL OTIMIZADA PARA RAILWAY
// PrintQRInTerminal REMOVIDO + Rota /qr
// ============================================

const express = require("express");
const path = require("path");
const qrcode = require("qrcode");
const pino = require("pino");
const fs = require("fs");
const multer = require("multer");
const xlsx = require("xlsx");
const csv = require("csv-parser");
const fetch = require("node-fetch");

const app = express();
const PORT = process.env.PORT || 8080;

app.use(express.json());
app.use(express.static("public"));
app.use(express.urlencoded({ extended: true }));
app.use('/uploads', express.static('uploads'));

// Health Check
app.get("/health", (req, res) => {
    res.status(200).json({ status: "healthy", timestamp: new Date().toISOString() });
});

app.get("/ping", (req, res) => res.send("pong"));

// Baileys
const {
    default: makeWASocket,
    useMultiFileAuthState,
    DisconnectReason,
    fetchLatestBaileysVersion,
    makeCacheableSignalKeyStore
} = require("@whiskeysockets/baileys");

// ====================== WHATSAPP MANAGER ======================
class WhatsAppManager {
    constructor() {
        this.connections = new Map();
        this.nextConnectionId = 1;
    }

    generateConnectionId() { return `conn_${this.nextConnectionId++}`; }

    async createConnection(connectionId = null) {
        const id = connectionId || this.generateConnectionId();
        const baseDir = process.env.RAILWAY_VOLUME_MOUNT_PATH || '/tmp';
        const authDir = path.join(baseDir, `auth_info_${id}`);

        console.log(`🔄 [${id}] Criando conexão - Pasta: ${authDir}`);

        if (!fs.existsSync(authDir)) fs.mkdirSync(authDir, { recursive: true });

        const { state, saveCreds } = await useMultiFileAuthState(authDir);
        const { version } = await fetchLatestBaileysVersion();

        const sock = makeWASocket({
            version,
            logger: pino({ level: "silent" }),
            // printQRInTerminal REMOVIDO aqui!
            auth: {
                creds: state.creds,
                keys: makeCacheableSignalKeyStore(state.keys, pino({ level: "silent" })),
            },
            browser: ["Windows 10", "Chrome", "130.0.0.0"],
            keepAliveIntervalMs: 30000,
        });

        const connData = { id, sock, authDir, qrCode: null, isConnected: false, status: "disconnected", saveCreds };
        this.setupConnectionEvents(connData);
        this.connections.set(id, connData);
        return connData;
    }

    setupConnectionEvents(connData) {
        const { sock, id } = connData;

        sock.ev.on("connection.update", async (update) => {
            const { connection, lastDisconnect, qr } = update;

            if (qr) {
                try {
                    connData.qrCode = await qrcode.toDataURL(qr);
                    connData.status = "awaiting_qr";
                    console.log(`✅ [${id}] QR Code gerado! Acesse: /qr`);
                } catch (e) { console.error("Erro QR:", e.message); }
            }

            if (connection === "open") {
                console.log(`✅ [${id}] CONECTADO COM SUCESSO!`);
                connData.isConnected = true;
                connData.status = "connected";
                connData.qrCode = null;
            }

            if (connection === "close") {
                console.log(`❌ [${id}] Conexão fechada`);
                if (lastDisconnect?.error?.output?.statusCode !== DisconnectReason.loggedOut) {
                    setTimeout(() => this.createConnection(id), 8000);
                }
            }
        });

        sock.ev.on("creds.update", connData.saveCreds);
    }

    getAllConnections() { return Array.from(this.connections.values()); }
}

// Instâncias
const whatsappManager = new WhatsAppManager();

// ====================== ROTAS ======================
app.get("/qr", (req, res) => {
    const conn = whatsappManager.getAllConnections()[0];
    if (conn?.qrCode) {
        res.send(`<h2>Escaneie o QR Code</h2><img src="${conn.qrCode}" width="400"><br><a href="/qr">Atualizar</a>`);
    } else {
        res.send(`<h2>Aguardando QR Code...</h2><p>Recarregue em alguns segundos.</p><a href="/qr">Atualizar</a>`);
    }
});

app.get("/", (req, res) => {
    res.send(`<h1>Bot WhatsApp no Railway</h1><p><a href="/qr">Ver QR Code</a> | <a href="/health">Health</a></p>`);
});

// Iniciar servidor
app.listen(PORT, '0.0.0.0', () => {
    console.log(`🚀 Servidor rodando na porta ${PORT}`);
    console.log(`📱 Acesse /qr para escanear o código`);

    setTimeout(() => whatsappManager.createConnection(), 6000);
});
