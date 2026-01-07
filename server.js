// ============================================
// WHATSAPP BOT - VERSÃO PRODUÇÃO/CLOUD
// ============================================
const express = require("express");
const path = require("path");
const qrcode = require("qrcode");
const pino = require("pino");
const fs = require("fs");

// Carregar Baileys
let baileys;
try {
    baileys = require("@whiskeysockets/baileys");
} catch (error) {
    console.error("❌ ERRO: Baileys não instalado!");
    console.error("Execute: npm install @whiskeysockets/baileys@latest");
    process.exit(1);
}

const { 
    default: makeWASocket, 
    useMultiFileAuthState, 
    DisconnectReason,
    fetchLatestBaileysVersion,
    makeCacheableSignalKeyStore,
    Browsers 
} = baileys;

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware para CORS (importante para cloud)
app.use((req, res, next) => {
    res.header("Access-Control-Allow-Origin", "*");
    res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
    res.header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS");
    next();
});

app.use(express.json({ limit: '10mb' }));
app.use(express.static("public"));

// ============================================
// GERENCIADOR DE CONEXÕES OTIMIZADO PARA CLOUD
// ============================================
class WhatsAppManager {
    constructor() {
        this.connections = new Map();
        this.nextConnectionId = 1;
        this.maxConnections = process.env.MAX_CONNECTIONS || 5;
    }

    generateConnectionId() {
        return `conn_${this.nextConnectionId++}`;
    }

    async createConnection(connectionId = null) {
        // Limitar número de conexões
        if (this.connections.size >= this.maxConnections) {
            throw new Error(`Limite de ${this.maxConnections} conexões atingido`);
        }

        const id = connectionId || this.generateConnectionId();
        
        // Em cloud, usar diretório temporário
        const authDir = process.env.NODE_ENV === 'production' 
            ? path.join('/tmp', `auth_info_${id}`)
            : path.join(__dirname, `auth_info_${id}`);
        
        console.log(`🔄 [${id}] Criando conexão em ${authDir}...`);
        
        try {
            // Criar pasta de autenticação
            if (!fs.existsSync(authDir)) {
                fs.mkdirSync(authDir, { recursive: true });
            }

            const { state, saveCreds } = await useMultiFileAuthState(authDir);
            const { version } = await fetchLatestBaileysVersion();
            
            // CONFIGURAÇÃO OTIMIZADA PARA CLOUD
            const sock = makeWASocket({
                version,
                logger: pino({ level: process.env.LOG_LEVEL || 'info' }),
                printQRInTerminal: true,
                auth: {
                    creds: state.creds,
                    keys: makeCacheableSignalKeyStore(state.keys, pino()),
                },
                browser: Browsers.ubuntu('Chrome'),
                
                // Configurações para cloud
                connectTimeoutMs: 60000,
                defaultQueryTimeoutMs: 60000,
                keepAliveIntervalMs: 25000,
                markOnlineOnConnect: true,
                syncFullHistory: false,
                mobile: false,
                emitOwnEvents: true,
                fireInitQueries: true,
                
                // Otimizações
                generateHighQualityLinkPreview: false,
                shouldSyncHistoryMessage: () => false,
                maxMsgRetryCount: 3,
                
                // Patch para cloud
                patchMessageBeforeSending: (message) => {
                    if (message.messageContextInfo) {
                        delete message.messageContextInfo.deviceListMetadata;
                        delete message.messageContextInfo.deviceListMetadataVersion;
                    }
                    return message;
                }
            });

            const connectionData = {
                id: id,
                sock: sock,
                authDir: authDir,
                qrCode: null,
                isConnected: false,
                status: "connecting",
                userInfo: null,
                saveCreds: saveCreds,
                lastQR: null,
                qrGeneratedAt: null,
                connectionAttempts: 0,
                lastActivity: Date.now()
            };

            this.setupConnectionEvents(connectionData);
            this.connections.set(id, connectionData);
            
            console.log(`✅ [${id}] Conexão criada com sucesso!`);
            return connectionData;
            
        } catch (error) {
            console.error(`❌ [${id}] Erro:`, error.message);
            throw error;
        }
    }

    setupConnectionEvents(connectionData) {
        const { sock, id } = connectionData;

        sock.ev.on("connection.update", async (update) => {
            const { connection, lastDisconnect, qr } = update;
            
            connectionData.lastActivity = Date.now();
            
            // QR Code
            if (qr) {
                console.log(`🎯 [${id}] QR Code gerado`);
                
                try {
                    const qrDataURL = await qrcode.toDataURL(qr);
                    connectionData.qrCode = qrDataURL;
                    connectionData.lastQR = qr;
                    connectionData.status = "awaiting_qr";
                    connectionData.qrGeneratedAt = Date.now();
                    connectionData.isConnected = false;
                } catch (error) {
                    console.error(`❌ [${id}] Erro QR:`, error.message);
                }
            }

            // Conexão aberta
            if (connection === "open") {
                console.log(`✅ [${id}] CONECTADO!`);
                
                connectionData.status = "connected";
                connectionData.qrCode = null;
                connectionData.isConnected = true;
                connectionData.connectionAttempts = 0;
                
                try {
                    const user = sock.user;
                    connectionData.userInfo = {
                        id: user?.id,
                        name: user?.name || user?.verifiedName || "Usuário",
                        phone: user?.id?.split(":")[0] || "Desconhecido",
                        platform: user?.platform || "android"
                    };
                    console.log(`👤 [${id}] Usuário: ${connectionData.userInfo.name}`);
                } catch (error) {
                    console.error(`⚠️ [${id}] Info usuário:`, error.message);
                }
            }

            // Conexão fechada
            if (connection === "close") {
                console.log(`❌ [${id}] Conexão fechada`);
                
                const statusCode = lastDisconnect?.error?.output?.statusCode;
                const shouldReconnect = statusCode !== DisconnectReason.loggedOut;
                
                connectionData.status = "disconnected";
                connectionData.qrCode = null;
                connectionData.isConnected = false;
                connectionData.connectionAttempts++;

                // Auto-reconnect em cloud
                if (shouldReconnect && connectionData.connectionAttempts <= 3) {
                    const delay = connectionData.connectionAttempts * 5000;
                    console.log(`🔄 [${id}] Reconectando em ${delay/1000}s...`);
                    
                    setTimeout(async () => {
                        try {
                            await this.forceNewConnection(id);
                        } catch (error) {
                            console.error(`❌ [${id}] Falha reconexão:`, error.message);
                        }
                    }, delay);
                }
            }
        });

        sock.ev.on("creds.update", connectionData.saveCreds);
    }

    getConnection(connectionId) {
        return this.connections.get(connectionId);
    }

    getAllConnections() {
        return Array.from(this.connections.values());
    }

    getActiveConnections() {
        return this.getAllConnections().filter(conn => conn.isConnected);
    }

    getConnectionsWithQR() {
        return this.getAllConnections().filter(conn => conn.qrCode !== null);
    }

    cleanupOldConnections(maxAge = 3600000) { // 1 hora
        const now = Date.now();
        for (const [id, conn] of this.connections.entries()) {
            if (!conn.isConnected && (now - conn.lastActivity) > maxAge) {
                console.log(`🧹 Limpando conexão antiga: ${id}`);
                this.disconnectConnection(id);
            }
        }
    }

    async disconnectConnection(connectionId) {
        const connection = this.connections.get(connectionId);
        if (connection) {
            try {
                await connection.sock.end();
                console.log(`🔌 [${connectionId}] Desconectado`);
            } catch (error) {
                console.error(`❌ [${connectionId}] Erro:`, error.message);
            }
            
            this.connections.delete(connectionId);
            
            // Limpar pasta em produção
            if (process.env.NODE_ENV === 'production') {
                try {
                    if (fs.existsSync(connection.authDir)) {
                        fs.rmSync(connection.authDir, { recursive: true, force: true });
                    }
                } catch (error) {
                    // Ignorar erros em produção
                }
            }
        }
        return true;
    }

    async forceNewConnection(connectionId) {
        await this.disconnectConnection(connectionId);
        return await this.createConnection(connectionId);
    }
}

// ============================================
// INICIALIZAÇÃO
// ============================================
const whatsappManager = new WhatsAppManager();

// Cleanup automático a cada 30 minutos
setInterval(() => {
    whatsappManager.cleanupOldConnections();
}, 30 * 60 * 1000);

// Delay entre mensagens
function getRandomDelay(min = 5000, max = 30000) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
}

// ============================================
// ROTAS DA API
// ============================================

// Health check (IMPORTANTE para cloud)
app.get("/health", (req, res) => {
    res.json({
        status: "healthy",
        timestamp: new Date().toISOString(),
        connections: whatsappManager.getAllConnections().length,
        active: whatsappManager.getActiveConnections().length,
        version: "1.0.0",
        environment: process.env.NODE_ENV || "development"
    });
});

// Status geral
app.get("/api/status", (req, res) => {
    const connections = whatsappManager.getAllConnections();
    
    res.json({
        success: true,
        data: {
            total: connections.length,
            active: whatsappManager.getActiveConnections().length,
            with_qr: whatsappManager.getConnectionsWithQR().length,
            connections: connections.map(conn => ({
                id: conn.id,
                status: conn.status,
                connected: conn.isConnected,
                user: conn.userInfo,
                has_qr: !!conn.qrCode,
                last_activity: conn.lastActivity
            }))
        }
    });
});

// Criar nova conexão
app.post("/api/connect", async (req, res) => {
    try {
        const { connectionId } = req.body;
        
        const connection = await whatsappManager.createConnection(connectionId);
        
        res.json({
            success: true,
            message: "Conexão iniciada. Aguarde QR Code.",
            connectionId: connection.id,
            hasQR: !!connection.qrCode
        });
        
    } catch (error) {
        res.status(500).json({
            success: false,
            message: `Erro: ${error.message}`
        });
    }
});

// Obter QR Code
app.get("/api/qrcode", async (req, res) => {
    try {
        const { connectionId } = req.query;
        let connection;
        
        if (connectionId) {
            connection = whatsappManager.getConnection(connectionId);
        } else {
            const connections = whatsappManager.getAllConnections();
            connection = connections.length > 0 ? connections[0] : null;
        }
        
        if (!connection) {
            return res.status(404).json({
                success: false,
                message: "Conexão não encontrada"
            });
        }
        
        // QR expirado (> 40 segundos)
        if (connection.qrGeneratedAt && (Date.now() - connection.qrGeneratedAt) > 40000) {
            connection.qrCode = null;
        }
        
        if (connection.qrCode) {
            res.json({
                success: true,
                qr: connection.qrCode,
                connectionId: connection.id,
                expires_in: 40 - Math.floor((Date.now() - connection.qrGeneratedAt) / 1000)
            });
        } else if (connection.isConnected) {
            res.json({
                success: true,
                message: "Já conectado",
                connected: true,
                user: connection.userInfo
            });
        } else {
            res.json({
                success: false,
                message: "Aguardando QR Code...",
                status: connection.status
            });
        }
        
    } catch (error) {
        res.status(500).json({
            success: false,
            message: `Erro: ${error.message}`
        });
    }
});

// Enviar mensagens
app.post("/api/send", async (req, res) => {
    try {
        const { connectionId, numbers, message, delay = true } = req.body;
        
        // Validações
        if (!numbers || !Array.isArray(numbers) || numbers.length === 0) {
            return res.status(400).json({
                success: false,
                message: "Lista de números é obrigatória"
            });
        }
        
        // Encontrar conexão
        let connection;
        if (connectionId) {
            connection = whatsappManager.getConnection(connectionId);
        } else {
            const activeConnections = whatsappManager.getActiveConnections();
            if (activeConnections.length === 0) {
                return res.status(400).json({
                    success: false,
                    message: "Nenhuma conexão ativa"
                });
            }
            connection = activeConnections[0];
        }
        
        if (!connection || !connection.isConnected) {
            return res.status(400).json({
                success: false,
                message: "Conexão não está ativa"
            });
        }
        
        const results = [];
        let successCount = 0;
        
        for (let i = 0; i < numbers.length; i++) {
            const number = numbers[i].replace(/\D/g, "");
            
            if (number.length < 10) {
                results.push({ number: numbers[i], success: false, error: "Número inválido" });
                continue;
            }
            
            try {
                const jid = number.length === 12 ? `${number}@s.whatsapp.net` : `55${number}@s.whatsapp.net`;
                await connection.sock.sendMessage(jid, { text: message });
                
                results.push({ number: numbers[i], success: true, message: "Enviada" });
                successCount++;
                
                // Delay
                if (delay && i < numbers.length - 1) {
                    await new Promise(resolve => setTimeout(resolve, getRandomDelay()));
                }
                
            } catch (error) {
                results.push({ number: numbers[i], success: false, error: error.message });
                
                if (delay && i < numbers.length - 1) {
                    await new Promise(resolve => setTimeout(resolve, 2000));
                }
            }
        }
        
        res.json({
            success: true,
            total: numbers.length,
            sent: successCount,
            results: results,
            connection: connection.id
        });
        
    } catch (error) {
        res.status(500).json({
            success: false,
            message: `Erro: ${error.message}`
        });
    }
});

// Limpar conexões
app.post("/api/clear", async (req, res) => {
    try {
        const connections = whatsappManager.getAllConnections();
        
        for (const connection of connections) {
            await whatsappManager.disconnectConnection(connection.id);
        }
        
        res.json({
            success: true,
            message: `Todas as ${connections.length} conexões limpas`
        });
        
    } catch (error) {
        res.status(500).json({
            success: false,
            message: `Erro: ${error.message}`
        });
    }
});

// Página principal
app.get("/", (req, res) => {
    res.sendFile(path.join(__dirname, "public", "index.html"));
});

// Para SPA (Single Page Application)
app.get("*", (req, res) => {
    res.redirect("/");
});

// ============================================
// INICIAR SERVIDOR
// ============================================
app.listen(PORT, () => {
    console.log(`
╔══════════════════════════════════════════╗
║     WHATSAPP BOT - VERSÃO CLOUD          ║
║     Porta: ${PORT}                          ║
║     Modo: ${process.env.NODE_ENV || 'development'}                    ║
║     Max conexões: ${process.env.MAX_CONNECTIONS || 5}                    ║
╚══════════════════════════════════════════╝
    `);
    
    console.log(`🚀 Servidor online!`);
    console.log(`🌐 Acesse a interface web`);
    console.log(`📡 Health check: /health`);
    console.log(`🔗 API Status: /api/status\n`);
});
