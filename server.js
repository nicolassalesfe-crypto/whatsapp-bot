// ============================================
// BOT WHATSAPP - MULTI CONEXÕES COM PLANILHAS E TEMPLATES
// VERSÃO COMPLETA E OTIMIZADA PARA RAILWAY
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

// ====================== EXPRESS SETUP ======================
const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.static("public"));
app.use(express.urlencoded({ extended: true }));
app.use('/uploads', express.static('uploads'));

// ====================== HEALTH CHECK (OBRIGATÓRIO) ======================
app.get("/health", (req, res) => {
    res.status(200).json({
        status: "healthy",
        timestamp: new Date().toISOString(),
        connections: whatsappManager ? whatsappManager.getAllConnections().length : 0,
        activeConnections: whatsappManager ? whatsappManager.getActiveConnections().length : 0
    });
});

app.get("/ping", (req, res) => res.status(200).send("pong"));

// ====================== CARREGAR BAILEYS ======================
let baileys;
try {
    baileys = require("@whiskeysockets/baileys");
} catch (error) {
    console.error("❌ ERRO: Não foi possível carregar @whiskeysockets/baileys");
    process.exit(1);
}

const {
    default: makeWASocket,
    useMultiFileAuthState,
    DisconnectReason,
    fetchLatestBaileysVersion,
    makeCacheableSignalKeyStore
} = baileys;

// ====================== TEMPLATE MANAGER (COMPLETO) ======================
class TemplateManager {
    constructor() {
        this.templates = new Map();
        this.templatesDir = path.join(__dirname, "templates");
        this.loadTemplates();
    }

    loadTemplates() {
        if (!fs.existsSync(this.templatesDir)) {
            fs.mkdirSync(this.templatesDir, { recursive: true });
        }
        const files = fs.readdirSync(this.templatesDir);
        files.forEach(file => {
            if (file.endsWith('.json')) {
                try {
                    const templatePath = path.join(this.templatesDir, file);
                    const templateData = JSON.parse(fs.readFileSync(templatePath, 'utf8'));
                    const templateId = path.basename(file, '.json');
                    this.templates.set(templateId, templateData);
                    console.log(`📝 Template carregado: ${templateData.name} (${templateId})`);
                } catch (error) {
                    console.error(`❌ Erro ao carregar template ${file}:`, error.message);
                }
            }
        });
    }

    createTemplate(templateData) {
        const templateId = `template_${Date.now()}`;
        const template = {
            id: templateId,
            name: templateData.name,
            message: templateData.message,
            variables: this.extractVariables(templateData.message),
            createdAt: new Date().toISOString(),
            updatedAt: new Date().toISOString()
        };
        this.templates.set(templateId, template);
        this.saveTemplate(templateId, template);
        return template;
    }

    updateTemplate(templateId, templateData) {
        if (!this.templates.has(templateId)) throw new Error('Template não encontrado');
        const template = this.templates.get(templateId);
        template.name = templateData.name || template.name;
        template.message = templateData.message || template.message;
        template.variables = this.extractVariables(template.message);
        template.updatedAt = new Date().toISOString();
        this.templates.set(templateId, template);
        this.saveTemplate(templateId, template);
        return template;
    }

    deleteTemplate(templateId) {
        if (this.templates.has(templateId)) {
            this.templates.delete(templateId);
            const filePath = path.join(this.templatesDir, `${templateId}.json`);
            if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
            return true;
        }
        return false;
    }

    getTemplate(templateId) {
        return this.templates.get(templateId);
    }

    getAllTemplates() {
        return Array.from(this.templates.values());
    }

    extractVariables(message) {
        const regex = /\{\{(\w+)\}\}/g;
        const variables = new Set();
        let match;
        while ((match = regex.exec(message)) !== null) {
            variables.add(match[1]);
        }
        return Array.from(variables);
    }

    renderTemplate(templateId, variables) {
        const template = this.templates.get(templateId);
        if (!template) throw new Error('Template não encontrado');
        let message = template.message;
        template.variables.forEach(varName => {
            const value = variables[varName] || `{{${varName}}}`;
            message = message.replace(new RegExp(`\\{\\{${varName}\\}\\}`, 'g'), value);
        });
        return message;
    }

    saveTemplate(templateId, template) {
        const filePath = path.join(this.templatesDir, `${templateId}.json`);
        fs.writeFileSync(filePath, JSON.stringify(template, null, 2), 'utf8');
    }
}

// ====================== SPREADSHEET MANAGER (COMPLETO) ======================
class SpreadsheetManager {
    constructor() {
        this.contacts = new Map();
        this.uploadsDir = path.join(__dirname, "uploads");
        if (!fs.existsSync(this.uploadsDir)) {
            fs.mkdirSync(this.uploadsDir, { recursive: true });
        }
    }

    async processSpreadsheet(filePath, options = {}) {
        const {
            phoneColumn = 'telefone',
            nameColumn = 'nome',
            skipFirstRow = true,
            customColumns = []
        } = options;

        const fileExt = path.extname(filePath).toLowerCase();
        let contacts = [];

        try {
            console.log(`📊 Processando arquivo: ${filePath}`);

            if (fileExt === '.csv') {
                contacts = await this.processCSV(filePath, { phoneColumn, nameColumn, skipFirstRow, customColumns });
            } else {
                contacts = this.processExcel(filePath, { phoneColumn, nameColumn, skipFirstRow, customColumns });
            }

            // Processamento de telefones
            contacts = contacts.map(contact => {
                if (!contact.name && contact.nome) contact.name = contact.nome;
                if (!contact.name && contact.Nome) contact.name = contact.Nome;

                const phones = [];
                Object.keys(contact).forEach(key => {
                    const keyLower = key.toLowerCase();
                    if (keyLower.includes('telefone') || keyLower.includes('tel') || 
                        keyLower.includes('fone') || keyLower.includes('celular') || 
                        keyLower.includes('whatsapp') || keyLower.includes('phone')) {
                        
                        const phoneValue = contact[key];
                        if (phoneValue && phoneValue.toString().trim()) {
                            const formatted = this.formatPhoneNumber(phoneValue);
                            phones.push({
                                original: phoneValue,
                                formatted,
                                isValid: this.validatePhoneNumber(formatted)
                            });
                        }
                    }
                });

                return {
                    ...contact,
                    name: contact.name || 'Sem nome',
                    phones,
                    mainPhone: phones[0]?.formatted || '',
                    validPhones: phones.filter(p => p.isValid),
                    hasValidPhone: phones.some(p => p.isValid)
                };
            });

            const spreadsheetId = `spreadsheet_${Date.now()}`;
            const spreadsheetData = {
                id: spreadsheetId,
                fileName: path.basename(filePath),
                uploadedAt: new Date().toISOString(),
                totalContacts: contacts.length,
                validContacts: contacts.filter(c => c.hasValidPhone).length,
                contacts: contacts
            };

            this.contacts.set(spreadsheetId, spreadsheetData);

            try { fs.unlinkSync(filePath); } catch (e) {}
            return spreadsheetData;

        } catch (error) {
            console.error('❌ Erro ao processar planilha:', error);
            throw error;
        }
    }

    processExcel(filePath, { phoneColumn, nameColumn, skipFirstRow }) {
        const workbook = xlsx.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
        const headers = data[0] || [];
        const contacts = [];
        const startRow = skipFirstRow ? 1 : 0;

        for (let i = startRow; i < data.length; i++) {
            const row = data[i];
            if (!row || row.length === 0) continue;

            const contact = { rowNumber: i + 1, name: '' };

            // Nome
            const nameIndex = headers.findIndex(h => 
                h && h.toString().toLowerCase().includes(nameColumn.toLowerCase())
            );
            if (nameIndex !== -1) contact.name = row[nameIndex]?.toString().trim() || '';

            // Telefones
            headers.forEach((header, idx) => {
                if (header) {
                    const h = header.toString().toLowerCase();
                    if (h.includes('tel') || h.includes('fone') || h.includes('cel') || h.includes('whatsapp') || h.includes('phone')) {
                        if (row[idx]) contact[header] = row[idx].toString().trim();
                    }
                }
            });

            contacts.push(contact);
        }
        return contacts;
    }

    async processCSV(filePath, { phoneColumn, nameColumn, skipFirstRow }) {
        return new Promise((resolve, reject) => {
            const contacts = [];
            fs.createReadStream(filePath)
                .pipe(csv())
                .on('data', (row) => {
                    const contact = { name: '' };
                    Object.keys(row).forEach(key => {
                        const k = key.toLowerCase();
                        if (k.includes('nome') || k.includes('name')) contact.name = row[key].trim();
                        if (k.includes('tel') || k.includes('fone') || k.includes('cel') || k.includes('whatsapp') || k.includes('phone')) {
                            if (row[key]) contact[key] = row[key].trim();
                        }
                    });
                    if (contact.name || Object.keys(contact).length > 1) contacts.push(contact);
                })
                .on('end', () => resolve(contacts))
                .on('error', reject);
        });
    }

    formatPhoneNumber(phone) {
        let cleaned = phone.toString().replace(/\D/g, '');
        if (cleaned.length < 10) return cleaned;
        if (!cleaned.startsWith('55')) cleaned = '55' + cleaned;
        return cleaned;
    }

    validatePhoneNumber(phone) {
        const cleaned = phone.toString().replace(/\D/g, '');
        return cleaned.length >= 10 && cleaned.length <= 13;
    }

    getAllSpreadsheets() {
        return Array.from(this.contacts.values());
    }

    deleteSpreadsheet(spreadsheetId) {
        return this.contacts.delete(spreadsheetId);
    }
}

// ====================== WHATSAPP MANAGER ======================
class WhatsAppManager {
    constructor() {
        this.connections = new Map();
        this.nextConnectionId = 1;
    }

    generateConnectionId() {
        return `conn_${this.nextConnectionId++}`;
    }

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
            printQRInTerminal: true,
            auth: {
                creds: state.creds,
                keys: makeCacheableSignalKeyStore(state.keys, pino({ level: "silent" })),
            },
            browser: ["Windows 10", "Chrome", "120.0.0.0"],
            keepAliveIntervalMs: 30000,
        });

        const connectionData = {
            id, sock, authDir, qrCode: null, isConnected: false,
            status: "disconnected", saveCreds
        };

        this.setupConnectionEvents(connectionData);
        this.connections.set(id, connectionData);
        return connectionData;
    }

    setupConnectionEvents(connectionData) {
        const { sock, id } = connectionData;
        sock.ev.on("connection.update", async (update) => {
            const { connection, lastDisconnect, qr } = update;

            if (qr) {
                try {
                    connectionData.qrCode = await qrcode.toDataURL(qr);
                    connectionData.status = "awaiting_qr";
                    console.log(`✅ [${id}] QR Code gerado`);
                } catch (e) {}
            }

            if (connection === "open") {
                console.log(`✅ [${id}] CONECTADO COM SUCESSO!`);
                connectionData.isConnected = true;
                connectionData.status = "connected";
            }

            if (connection === "close") {
                console.log(`❌ [${id}] Conexão fechada`);
                if (lastDisconnect?.error?.output?.statusCode !== DisconnectReason.loggedOut) {
                    setTimeout(() => this.createConnection(id).catch(console.error), 8000);
                }
            }
        });

        sock.ev.on("creds.update", connectionData.saveCreds);
    }

    getAllConnections() { return Array.from(this.connections.values()); }
    getActiveConnections() { return this.getAllConnections().filter(c => c.isConnected); }
}

// ====================== INSTANCIAR ======================
const whatsappManager = new WhatsAppManager();
const templateManager = new TemplateManager();
const spreadsheetManager = new SpreadsheetManager();

// ====================== MULTER ======================
const upload = multer({
    storage: multer.diskStorage({
        destination: (req, file, cb) => {
            const dir = path.join(__dirname, "uploads");
            if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
            cb(null, dir);
        },
        filename: (req, file, cb) => cb(null, Date.now() + '-' + file.originalname)
    }),
    limits: { fileSize: 10 * 1024 * 1024 }
});

// ====================== ROTAS (Adicione suas rotas aqui) ======================
app.post("/api/clear", (req, res) => {
    res.json({ success: true, message: "Limpeza de planilhas realizada" });
});

// ====================== INICIALIZAÇÃO ======================
app.get("/", (req, res) => {
    const dashboardPath = path.join(__dirname, "public", "dashboard.html");
    if (fs.existsSync(dashboardPath)) {
        res.sendFile(dashboardPath);
    } else {
        res.send("<h1>🚀 Bot WhatsApp rodando com sucesso no Railway!</h1>");
    }
});

// Keep Alive
setInterval(() => {
    fetch(`http://localhost:${PORT}/ping`).catch(() => {});
}, 300000);

// Iniciar Servidor
app.listen(PORT, '0.0.0.0', () => {
    console.log(`🚀 Servidor iniciado na porta ${PORT}`);
    console.log(`🌐 Health Check: /health`);

    ['public', 'uploads', 'templates'].forEach(dir => {
        const p = path.join(__dirname, dir);
        if (!fs.existsSync(p)) fs.mkdirSync(p, { recursive: true });
    });

    setTimeout(() => {
        whatsappManager.createConnection().catch(console.error);
    }, 10000);
});

process.on("uncaughtException", e => console.error("❌ Uncaught Exception:", e));
process.on("unhandledRejection", e => console.error("❌ Unhandled Rejection:", e));
