// ============================================
// BOT WHATSAPP - MULTI CONEXÕES COM PLANILHAS E TEMPLATES
// VERSÃO OTIMIZADA PARA RAILWAY
// ============================================

// PRIMEIRO: Importar todas as dependências
const express = require("express");
const { Boom } = require("@hapi/boom");
const path = require("path");
const qrcode = require("qrcode");
const pino = require("pino");
const fs = require("fs");
const multer = require("multer");
const xlsx = require("xlsx");
const csv = require("csv-parser");
const fetch = require("node-fetch"); // ADICIONADO: necessário para keep-alive

// SEGUNDO: Criar a aplicação Express
const app = express();
const PORT = process.env.PORT || 3000; // Railway usa a porta da variável de ambiente

// TERCEIRO: Configurar o Express ANTES de qualquer outra coisa
app.use(express.json());
app.use(express.static("public"));
app.use(express.urlencoded({ extended: true }));
app.use('/uploads', express.static('uploads'));

// QUARTO: Configurações específicas para o Railway
// ============================================
// CONFIGURAÇÕES RAILWAY - ADICIONADO
// ============================================

// Middleware para logging de requisições (útil para debug no Railway)
app.use((req, res, next) => {
    console.log(`📨 [${new Date().toISOString()}] ${req.method} ${req.url}`);
    next();
});

// Rota de health check para o Railway (OBRIGATÓRIO)
app.get("/health", (req, res) => {
    res.status(200).json({ 
        status: "healthy", 
        timestamp: new Date().toISOString(),
        connections: whatsappManager?.getAllConnections().length || 0,
        activeConnections: whatsappManager?.getActiveConnections().length || 0
    });
});

// Rota de ping para keep-alive
app.get("/ping", (req, res) => {
    res.status(200).send("pong");
});

// ============================================
// CONTINUA SEU CÓDIGO ORIGINAL
// ============================================

// QUINTO: Carregar o Baileys (depois do Express)
let baileys;
try {
    baileys = require("@whiskeysockets/baileys");
} catch (error) {
    console.error("❌ ERRO: Não foi possível carregar @whiskeysockets/baileys");
    console.error("Execute: npm install @whiskeysockets/baileys --ignore-scripts");
    process.exit(1);
}

const { 
    default: makeWASocket, 
    useMultiFileAuthState, 
    DisconnectReason,
    fetchLatestBaileysVersion,
    makeCacheableSignalKeyStore 
} = baileys;

// ============================================
// CONFIGURAÇÃO DE UPLOAD
// ============================================
const upload = multer({
    storage: multer.diskStorage({
        destination: (req, file, cb) => {
            const uploadDir = path.join(__dirname, "uploads");
            if (!fs.existsSync(uploadDir)) {
                fs.mkdirSync(uploadDir, { recursive: true });
            }
            cb(null, uploadDir);
        },
        filename: (req, file, cb) => {
            const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
            cb(null, file.fieldname + '-' + uniqueSuffix + path.extname(file.originalname));
        }
    }),
    fileFilter: (req, file, cb) => {
        const allowedExtensions = ['.xlsx', '.xls', '.csv'];
        const ext = path.extname(file.originalname).toLowerCase();
        
        if (allowedExtensions.includes(ext)) {
            cb(null, true);
        } else {
            cb(new Error('Formato de arquivo inválido. Use .xlsx, .xls ou .csv'));
        }
    },
    limits: {
        fileSize: 10 * 1024 * 1024 // 10MB
    }
});

// ============================================
// GERENCIADOR DE TEMPLATES
// ============================================
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
        if (!this.templates.has(templateId)) {
            throw new Error('Template não encontrado');
        }

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
            if (fs.existsSync(filePath)) {
                fs.unlinkSync(filePath);
            }
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
        if (!template) {
            throw new Error('Template não encontrado');
        }

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

// ============================================
// GERENCIADOR DE PLANILHAS
// ============================================
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
            console.log(`📋 Opções:`, { phoneColumn, nameColumn, skipFirstRow, customColumns });

            if (fileExt === '.csv') {
                contacts = await this.processCSV(filePath, { phoneColumn, nameColumn, skipFirstRow, customColumns });
            } else {
                contacts = this.processExcel(filePath, { phoneColumn, nameColumn, skipFirstRow, customColumns });
            }

            console.log(`📊 Total de linhas lidas: ${contacts.length}`);

            // Processar contatos com múltiplos telefones e garantir que o nome seja capturado
            contacts = contacts.map(contact => {
                // Garantir que o nome está presente
                if (!contact.name && contact.nome) {
                    contact.name = contact.nome;
                }
                if (!contact.name && contact.Nome) {
                    contact.name = contact.Nome;
                }
                if (!contact.name && contact.cliente) {
                    contact.name = contact.cliente;
                }
                if (!contact.name && contact.Cliente) {
                    contact.name = contact.Cliente;
                }
                
                // Array para armazenar todos os telefones encontrados
                const phones = [];
                
                // Se tiver phone/telefone principal, adicionar primeiro
                if (contact.phone || contact.telefone) {
                    const mainPhone = contact.phone || contact.telefone;
                    if (mainPhone && mainPhone.toString().trim()) {
                        const formattedPhone = this.formatPhoneNumber(mainPhone);
                        phones.push({
                            original: mainPhone,
                            formatted: formattedPhone,
                            isValid: this.validatePhoneNumber(formattedPhone),
                            column: 'principal'
                        });
                    }
                }
                
                // Coletar todos os campos que parecem ser telefone (incluindo telefone2, telefone3, etc)
                Object.keys(contact).forEach(key => {
                    // Ignorar o telefone principal já adicionado
                    if (key === 'phone' || key === 'telefone') return;
                    
                    const keyLower = key.toLowerCase();
                    if (keyLower.includes('telefone') || 
                        keyLower.includes('tel') || 
                        keyLower.includes('fone') || 
                        keyLower.includes('celular') ||
                        keyLower.includes('whatsapp') ||
                        keyLower.includes('phone')) {
                        
                        const phoneValue = contact[key];
                        if (phoneValue && phoneValue.toString().trim()) {
                            const formattedPhone = this.formatPhoneNumber(phoneValue);
                            // Evitar duplicatas
                            if (!phones.some(p => p.formatted === formattedPhone)) {
                                phones.push({
                                    original: phoneValue,
                                    formatted: formattedPhone,
                                    isValid: this.validatePhoneNumber(formattedPhone),
                                    column: key
                                });
                            }
                        }
                    }
                });

                // Remover duplicatas (mesmo número formatado)
                const uniquePhones = [];
                const seenFormatted = new Set();
                phones.forEach(phone => {
                    if (!seenFormatted.has(phone.formatted) && phone.formatted) {
                        seenFormatted.add(phone.formatted);
                        uniquePhones.push(phone);
                    }
                });

                return {
                    ...contact,
                    name: contact.name || 'Sem nome',
                    phones: uniquePhones,
                    mainPhone: uniquePhones.length > 0 ? uniquePhones[0].formatted : '',
                    allPhones: uniquePhones.map(p => p.formatted),
                    validPhones: uniquePhones.filter(p => p.isValid),
                    invalidPhones: uniquePhones.filter(p => !p.isValid),
                    hasValidPhone: uniquePhones.some(p => p.isValid)
                };
            });

            const validContacts = contacts.filter(c => c.hasValidPhone);
            const invalidContacts = contacts.filter(c => !c.hasValidPhone);

            console.log(`✅ Contatos com pelo menos 1 telefone válido: ${validContacts.length}`);
            console.log(`❌ Contatos sem telefone válido: ${invalidContacts.length}`);
            
            // Estatísticas de telefones
            const totalPhones = contacts.reduce((sum, c) => sum + c.phones.length, 0);
            const totalValidPhones = contacts.reduce((sum, c) => sum + c.validPhones.length, 0);
            console.log(`📞 Total de telefones encontrados: ${totalPhones}`);
            console.log(`📞 Telefones válidos: ${totalValidPhones}`);
            
            // Mostrar primeiros contatos para debug
            if (validContacts.length > 0) {
                console.log(`📋 Primeiro contato:`, {
                    name: validContacts[0].name,
                    phones: validContacts[0].phones.map(p => ({ original: p.original, formatted: p.formatted }))
                });
            }

            const spreadsheetId = `spreadsheet_${Date.now()}`;
            
            const spreadsheetData = {
                id: spreadsheetId,
                fileName: path.basename(filePath),
                filePath: filePath,
                uploadedAt: new Date().toISOString(),
                totalContacts: contacts.length,
                validContacts: validContacts.length,
                invalidContacts: invalidContacts.length,
                totalPhones: totalPhones,
                totalValidPhones: totalValidPhones,
                contacts: validContacts,
                invalidList: invalidContacts,
                columns: contacts.length > 0 ? Object.keys(contacts[0]) : []
            };

            this.contacts.set(spreadsheetId, spreadsheetData);
            
            try {
                fs.unlinkSync(filePath);
                console.log(`🗑️ Arquivo temporário removido: ${filePath}`);
            } catch (err) {
                console.log(`⚠️ Não foi possível remover arquivo: ${err.message}`);
            }

            return spreadsheetData;

        } catch (error) {
            console.error('❌ Erro ao processar planilha:', error);
            try {
                if (fs.existsSync(filePath)) {
                    fs.unlinkSync(filePath);
                }
            } catch (err) {}
            throw error;
        }
    }

    processExcel(filePath, { phoneColumn, nameColumn, skipFirstRow, customColumns }) {
        try {
            console.log(`📖 Lendo arquivo Excel: ${filePath}`);
            const workbook = xlsx.readFile(filePath);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
            
            if (data.length === 0) {
                throw new Error('Arquivo vazio');
            }

            console.log(`📏 Dimensões da planilha: ${data.length} linhas`);

            const headers = data[0] || [];
            console.log(`📋 Cabeçalhos encontrados:`, headers);

            // Encontrar índice da coluna de nome (priorizar o nameColumn fornecido)
            let nameIndex = -1;
            
            // Primeiro tentar encontrar a coluna exata especificada
            nameIndex = headers.findIndex(h => 
                h && h.toString().toLowerCase() === nameColumn.toLowerCase()
            );
            
            // Se não encontrar, tentar encontrar coluna que contenha o termo
            if (nameIndex === -1) {
                nameIndex = headers.findIndex(h => 
                    h && h.toString().toLowerCase().includes(nameColumn.toLowerCase())
                );
            }
            
            // Se ainda não encontrar, procurar por variações comuns
            if (nameIndex === -1) {
                const nameVariations = ['nome', 'cliente', 'nome completo', 'cliente nome', 'name', 'customer', 'contact'];
                for (const variation of nameVariations) {
                    nameIndex = headers.findIndex(h => 
                        h && h.toString().toLowerCase().includes(variation)
                    );
                    if (nameIndex !== -1) {
                        console.log(`👤 Coluna de nome encontrada por variação: ${headers[nameIndex]}`);
                        break;
                    }
                }
            }

            console.log(`👤 Coluna de nome: ${nameIndex !== -1 ? headers[nameIndex] : 'não encontrada'}`);

            // Encontrar TODAS as colunas que podem conter telefones
            const phoneIndices = [];
            const phoneColumnNames = [];

            headers.forEach((header, index) => {
                if (header) {
                    const headerStr = header.toString().toLowerCase();
                    // Verificar se é uma coluna de telefone
                    if (headerStr.includes('telefone') || 
                        headerStr.includes('tel') || 
                        headerStr.includes('fone') || 
                        headerStr.includes('celular') ||
                        headerStr.includes('whatsapp') ||
                        headerStr.includes('phone') ||
                        headerStr.match(/tel[^\w]?\d*/i)) { // Para "tel 1", "tel2", etc
                        
                        phoneIndices.push(index);
                        phoneColumnNames.push(header);
                    }
                }
            });

            // Se não encontrar nenhuma, tentar a coluna padrão
            if (phoneIndices.length === 0) {
                const phoneIndex = headers.findIndex(h => 
                    h && h.toString().toLowerCase().includes(phoneColumn.toLowerCase())
                );
                if (phoneIndex !== -1) {
                    phoneIndices.push(phoneIndex);
                    phoneColumnNames.push(headers[phoneIndex]);
                }
            }

            console.log(`📞 Colunas de telefone encontradas:`, phoneColumnNames);

            if (phoneIndices.length === 0) {
                throw new Error(`Nenhuma coluna de telefone encontrada. Colunas disponíveis: ${headers.join(', ')}`);
            }

            const contacts = [];
            const startRow = skipFirstRow ? 1 : 0;

            for (let i = startRow; i < data.length; i++) {
                const row = data[i];
                if (!row || row.length === 0) continue;

                const contact = {};
                
                // Capturar nome
                if (nameIndex !== -1) {
                    contact.name = row[nameIndex]?.toString().trim() || '';
                    // Também manter como nome para compatibilidade
                    contact.nome = contact.name;
                } else {
                    contact.name = '';
                    contact.nome = '';
                }
                
                contact.rowNumber = i + 1;

                // Adicionar TODOS os telefones encontrados como campos separados
                phoneIndices.forEach((index, idx) => {
                    const phoneValue = row[index]?.toString().trim() || '';
                    if (phoneValue) {
                        // Manter o nome original da coluna
                        const originalColumnName = headers[index].toString().trim();
                        contact[originalColumnName] = phoneValue;
                        
                        // Também adicionar como telefone, telefone2, telefone3, etc
                        if (idx === 0) {
                            contact.telefone = phoneValue;
                            contact.phone = phoneValue;
                        } else {
                            contact[`telefone${idx + 1}`] = phoneValue;
                            contact[`phone${idx + 1}`] = phoneValue;
                        }
                    }
                });

                // Adicionar colunas personalizadas
                if (customColumns && customColumns.length > 0) {
                    customColumns.forEach(col => {
                        const colTrim = col.trim();
                        if (colTrim) {
                            const index = headers.findIndex(h => 
                                h && h.toString().toLowerCase().includes(colTrim.toLowerCase())
                            );
                            if (index !== -1) {
                                contact[colTrim] = row[index]?.toString().trim() || '';
                            }
                        }
                    });
                }

                // Só adicionar se tiver pelo menos um telefone
                if (Object.keys(contact).some(key => 
                    key.toLowerCase().includes('telefone') || 
                    key.toLowerCase().includes('tel') || 
                    key === 'phone' || 
                    key.startsWith('phone'))) {
                    contacts.push(contact);
                }
            }

            console.log(`📊 Total de contatos extraídos: ${contacts.length}`);
            if (contacts.length > 0) {
                console.log(`📋 Exemplo do primeiro contato:`, {
                    name: contacts[0].name,
                    telefone: contacts[0].telefone,
                    telefone2: contacts[0].telefone2,
                    telefone3: contacts[0].telefone3
                });
            }
            return contacts;

        } catch (error) {
            console.error('❌ Erro no processamento Excel:', error);
            throw error;
        }
    }

    async processCSV(filePath, { phoneColumn, nameColumn, skipFirstRow, customColumns }) {
        return new Promise((resolve, reject) => {
            const contacts = [];
            let headers = [];
            let rowCount = 0;
            
            console.log(`📖 Lendo arquivo CSV: ${filePath}`);
            
            fs.createReadStream(filePath)
                .pipe(csv())
                .on('headers', (headerList) => {
                    headers = headerList;
                    console.log(`📋 Cabeçalhos CSV:`, headers);
                })
                .on('data', (row) => {
                    rowCount++;
                    
                    const contact = {};

                    // Encontrar o nome
                    let nameFound = false;
                    
                    // Primeiro tentar a coluna especificada
                    const nameKey = Object.keys(row).find(key => 
                        key.toLowerCase().includes(nameColumn.toLowerCase())
                    );
                    
                    if (nameKey) {
                        contact.name = row[nameKey]?.trim() || '';
                        contact.nome = contact.name;
                        nameFound = true;
                    }
                    
                    // Se não encontrou, procurar por variações
                    if (!nameFound) {
                        const nameVariations = ['nome', 'cliente', 'nome completo', 'cliente nome', 'name', 'customer', 'contact'];
                        for (const variation of nameVariations) {
                            const foundKey = Object.keys(row).find(key => 
                                key.toLowerCase().includes(variation)
                            );
                            if (foundKey) {
                                contact.name = row[foundKey]?.trim() || '';
                                contact.nome = contact.name;
                                break;
                            }
                        }
                    }
                    
                    if (!contact.name) {
                        contact.name = '';
                        contact.nome = '';
                    }

                    // Encontrar TODOS os campos de telefone
                    let phoneFound = false;
                    let phoneCount = 0;
                    
                    Object.keys(row).forEach(key => {
                        const keyLower = key.toLowerCase();
                        const value = row[key]?.trim() || '';
                        
                        if (value) {
                            // Verificar se é um campo de telefone
                            if (keyLower.includes('telefone') || 
                                keyLower.includes('tel') || 
                                keyLower.includes('fone') || 
                                keyLower.includes('celular') ||
                                keyLower.includes('whatsapp') ||
                                keyLower.includes('phone')) {
                                
                                contact[key] = value;
                                phoneFound = true;
                                phoneCount++;
                                
                                // Mapear para telefone, telefone2, telefone3, etc
                                if (!contact.telefone) {
                                    contact.telefone = value;
                                    contact.phone = value;
                                } else {
                                    contact[`telefone${phoneCount}`] = value;
                                    contact[`phone${phoneCount}`] = value;
                                }
                            }
                        }
                    });

                    // Se não encontrou por detecção automática, tentar a coluna padrão
                    if (!phoneFound) {
                        const phoneKey = Object.keys(row).find(key => 
                            key.toLowerCase().includes(phoneColumn.toLowerCase())
                        );
                        
                        if (phoneKey) {
                            const phoneValue = row[phoneKey]?.trim() || '';
                            if (phoneValue) {
                                contact.telefone = phoneValue;
                                contact.phone = phoneValue;
                                contact[phoneKey] = phoneValue;
                                phoneFound = true;
                            }
                        }
                    }

                    // Adicionar colunas personalizadas
                    if (customColumns && customColumns.length > 0) {
                        customColumns.forEach(col => {
                            const colTrim = col.trim();
                            if (colTrim) {
                                const colKey = Object.keys(row).find(key => 
                                    key.toLowerCase().includes(colTrim.toLowerCase())
                                );
                                if (colKey) {
                                    contact[colTrim] = row[colKey]?.trim() || '';
                                }
                            }
                        });
                    }

                    if (phoneFound) {
                        contacts.push(contact);
                    }
                })
                .on('end', () => {
                    console.log(`📊 Total de linhas lidas: ${rowCount}`);
                    console.log(`📊 Total de contatos extraídos: ${contacts.length}`);
                    if (contacts.length > 0) {
                        console.log(`📋 Exemplo do primeiro contato:`, {
                            name: contacts[0].name,
                            telefone: contacts[0].telefone,
                            telefone2: contacts[0].telefone2,
                            telefone3: contacts[0].telefone3
                        });
                    }
                    resolve(contacts);
                })
                .on('error', (error) => {
                    console.error('❌ Erro no processamento CSV:', error);
                    reject(error);
                });
        });
    }

    formatPhoneNumber(phone) {
        if (!phone) return '';
        
        let cleaned = phone.toString().replace(/\D/g, '');
        
        if (!cleaned) return '';
        
        if (cleaned.length < 10) return cleaned;
        
        if (!cleaned.startsWith('55') && (cleaned.length === 10 || cleaned.length === 11)) {
            cleaned = '55' + cleaned;
        }
        
        if (cleaned.length === 13 && cleaned.startsWith('55')) {
            const ddd = cleaned.substring(2, 4);
            const number = cleaned.substring(4);
            if (number.startsWith('9') && number.length === 9) {
                cleaned = '55' + ddd + number.substring(1);
            }
        }
        
        return cleaned;
    }

    validatePhoneNumber(phone) {
        if (!phone) return false;
        const cleaned = phone.toString().replace(/\D/g, '');
        return cleaned.length >= 10 && cleaned.length <= 13;
    }

    getContacts(spreadsheetId) {
        return this.contacts.get(spreadsheetId);
    }

    getAllSpreadsheets() {
        return Array.from(this.contacts.values());
    }

    deleteSpreadsheet(spreadsheetId) {
        return this.contacts.delete(spreadsheetId);
    }
}

// ============================================
// GERENCIADOR DE CONEXÕES WHATSAPP (MODIFICADO PARA RAILWAY)
// ============================================
class WhatsAppManager {
    constructor() {
        this.connections = new Map();
        this.nextConnectionId = 1;
        this.reconnectAttempts = new Map(); // Controlar tentativas de reconexão
    }

    generateConnectionId() {
        return `conn_${this.nextConnectionId++}`;
    }

    async createConnection(connectionId = null) {
        const id = connectionId || this.generateConnectionId();
        // Usar /tmp no Railway para arquivos temporários (persistência limitada)
        const authDir = process.env.RAILWAY_ENV 
            ? path.join('/tmp', `auth_info_${id}`) 
            : path.join(__dirname, `auth_info_${id}`);
        
        console.log(`🔄 [${id}] Criando conexão WhatsApp...`);
        console.log(`📁 Diretório auth: ${authDir}`);
        
        try {
            if (!fs.existsSync(authDir)) {
                fs.mkdirSync(authDir, { recursive: true });
            }

            const { state, saveCreds } = await useMultiFileAuthState(authDir);
            const { version, isLatest } = await fetchLatestBaileysVersion();
            
            const sock = makeWASocket({
                version,
                logger: pino({ level: "silent" }),
                printQRInTerminal: true,
                auth: {
                    creds: state.creds,
                    keys: makeCacheableSignalKeyStore(state.keys, pino({ level: "silent" })),
                },
                browser: ["Windows 10", "Chrome", "120.0.0.0"],
                syncFullHistory: false,
                markOnlineOnConnect: false,
                generateHighQualityLinkPreview: false,
                emitOwnEvents: false,
                defaultQueryTimeoutMs: 60000,
                // Importante para o Railway: keep alive mais frequente
                keepAliveIntervalMs: 30000,
                // Tentar reconectar mais rápido
                retryRequestDelayMs: 5000,
            });

            const connectionData = {
                id: id,
                sock: sock,
                authDir: authDir,
                qrCode: null,
                isConnected: false,
                status: "disconnected",
                userInfo: null,
                saveCreds: saveCreds,
                lastQR: null,
                qrGeneratedAt: null,
                reconnectCount: 0
            };

            this.setupConnectionEvents(connectionData);
            this.connections.set(id, connectionData);
            
            console.log(`✅ [${id}] Conexão criada com sucesso!`);
            return connectionData;
            
        } catch (error) {
            console.error(`❌ [${id}] Erro ao criar conexão:`, error.message);
            throw error;
        }
    }

    setupConnectionEvents(connectionData) {
        const { sock, id } = connectionData;

        sock.ev.on("connection.update", async (update) => {
            const { connection, lastDisconnect, qr } = update;
            
            console.log(`📡 [${id}] Status: ${connection || "unknown"}`);
            
            if (qr) {
                console.log(`🎯 [${id}] QR Code recebido!`);
                
                try {
                    const qrDataURL = await qrcode.toDataURL(qr);
                    connectionData.qrCode = qrDataURL;
                    connectionData.lastQR = qr;
                    connectionData.status = "awaiting_qr";
                    connectionData.qrGeneratedAt = Date.now();
                    connectionData.isConnected = false;
                    
                    console.log(`✅ [${id}] QR Code convertido para imagem`);
                    
                } catch (error) {
                    console.error(`❌ [${id}] Erro ao gerar QR Code:`, error.message);
                }
            }

            if (connection === "close") {
                console.log(`❌ [${id}] Conexão fechada`);
                const statusCode = lastDisconnect?.error?.output?.statusCode;
                const shouldReconnect = statusCode !== DisconnectReason.loggedOut;
                
                connectionData.status = "disconnected";
                connectionData.qrCode = null;
                connectionData.isConnected = false;
                connectionData.userInfo = null;

                console.log(`📊 [${id}] Status code: ${statusCode}, Reconectar: ${shouldReconnect}`);

                if (shouldReconnect) {
                    // Incrementar contador de reconexão
                    connectionData.reconnectCount = (connectionData.reconnectCount || 0) + 1;
                    
                    // Backoff exponencial para reconexões
                    const delay = Math.min(3000 * Math.pow(1.5, connectionData.reconnectCount - 1), 60000);
                    
                    console.log(`🔄 [${id}] Tentativa ${connectionData.reconnectCount} de reconexão em ${Math.round(delay/1000)}s...`);
                    
                    setTimeout(() => {
                        console.log(`🔄 [${id}] Iniciando reconexão...`);
                        this.createConnection(id).catch(err => {
                            console.error(`❌ [${id}] Falha na reconexão:`, err.message);
                        });
                    }, delay);
                } else {
                    console.log(`🚫 [${id}] Logged out, não reconectando automaticamente`);
                    this.connections.delete(id);
                    
                    // Limpar diretório auth se logged out
                    try {
                        if (fs.existsSync(connectionData.authDir)) {
                            fs.rmSync(connectionData.authDir, { recursive: true, force: true });
                            console.log(`🗑️ [${id}] Pasta auth removida`);
                        }
                    } catch (error) {
                        console.error(`❌ [${id}] Erro ao limpar auth:`, error.message);
                    }
                }
            } 
            else if (connection === "open") {
                console.log(`✅ [${id}] CONECTADO COM SUCESSO!`);
                connectionData.status = "connected";
                connectionData.qrCode = null;
                connectionData.lastQR = null;
                connectionData.isConnected = true;
                connectionData.reconnectCount = 0; // Reset contador de reconexão
                
                try {
                    connectionData.userInfo = {
                        id: sock.user?.id,
                        name: sock.user?.name || sock.user?.verifiedName || "Usuário",
                        phone: sock.user?.id?.split(":")[0] || "Desconhecido",
                        platform: sock.user?.platform || "android"
                    };
                    console.log(`👤 [${id}] Usuário:`, connectionData.userInfo);
                } catch (error) {
                    console.error(`❌ [${id}] Erro ao obter info:`, error.message);
                    connectionData.userInfo = { name: "Usuário", phone: "Desconhecido" };
                }
            } 
            else if (connection === "connecting") {
                console.log(`🔗 [${id}] Conectando...`);
                connectionData.status = "connecting";
            }
        });

        sock.ev.on("creds.update", connectionData.saveCreds);
        
        sock.ev.on("messages.upsert", (m) => {
            if (m.type === "notify") {
                console.log(`📨 [${id}] Nova mensagem recebida`);
            }
        });
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

    async disconnectConnection(connectionId) {
        const connection = this.connections.get(connectionId);
        if (connection) {
            try {
                await connection.sock.end();
                console.log(`🔌 [${connectionId}] Desconectado`);
            } catch (error) {
                console.error(`❌ [${connectionId}] Erro ao desconectar:`, error.message);
            }
            
            this.connections.delete(connectionId);
            
            try {
                if (fs.existsSync(connection.authDir)) {
                    fs.rmSync(connection.authDir, { recursive: true, force: true });
                    console.log(`🗑️ [${connectionId}] Pasta auth removida`);
                }
            } catch (error) {
                console.error(`❌ [${connectionId}] Erro ao limpar auth:`, error.message);
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
// INICIALIZAÇÃO DOS GERENCIADORES
// ============================================
const whatsappManager = new WhatsAppManager();
const templateManager = new TemplateManager();
const spreadsheetManager = new SpreadsheetManager();

// Delay aleatório entre mensagens
function getRandomDelay(min = 5000, max = 30000) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
}

// ============================================
// ROTAS DA API (TODAS AS SUAS ROTAS ORIGINAIS AQUI)
// ============================================

// [TODAS AS SUAS ROTAS PERMANECEM IGUAIS]
// /api/test, /api/spreadsheet/upload, /api/spreadsheets, etc.

// ============================================
// ROTA CORRIGIDA: LIMPAR APENAS PLANILHAS
// ============================================
app.post("/api/clear", async (req, res) => {
    try {
        // Limpar apenas planilhas, não conexões
        const spreadsheets = spreadsheetManager.getAllSpreadsheets();
        let deletedCount = 0;
        
        for (const spreadsheet of spreadsheets) {
            const deleted = spreadsheetManager.deleteSpreadsheet(spreadsheet.id);
            if (deleted) deletedCount++;
        }
        
        // Opcional: também limpar arquivos temporários da pasta uploads
        const uploadsDir = path.join(__dirname, "uploads");
        if (fs.existsSync(uploadsDir)) {
            const files = fs.readdirSync(uploadsDir);
            for (const file of files) {
                try {
                    fs.unlinkSync(path.join(uploadsDir, file));
                    console.log(`🗑️ Arquivo temporário removido: ${file}`);
                } catch (err) {
                    console.log(`⚠️ Erro ao remover ${file}: ${err.message}`);
                }
            }
        }
        
        res.json({
            success: true,
            message: `${deletedCount} planilha(s) foram limpas com sucesso`,
            deletedCount: deletedCount
        });
        
    } catch (error) {
        console.error("❌ Erro ao limpar planilhas:", error);
        res.status(500).json({
            success: false,
            message: `Erro ao limpar planilhas: ${error.message}`
        });
    }
});

// ============================================
// ROTA PARA LIMPAR APENAS CONEXÕES
// ============================================
app.post("/api/clear-connections", async (req, res) => {
    try {
        const connections = whatsappManager.getAllConnections();
        
        for (const connection of connections) {
            await whatsappManager.disconnectConnection(connection.id);
        }
        
        res.json({
            success: true,
            message: `Todas as ${connections.length} conexões foram limpas`
        });
        
    } catch (error) {
        res.status(500).json({
            success: false,
            message: `Erro: ${error.message}`
        });
    }
});

// ============================================
// PÁGINA PRINCIPAL
// ============================================
app.get("/", (req, res) => {
    const dashboardPath = path.join(__dirname, "public", "dashboard.html");
    if (fs.existsSync(dashboardPath)) {
        res.sendFile(dashboardPath);
    } else {
        res.send(`
            <!DOCTYPE html>
            <html>
            <head>
                <title>WhatsApp Bot</title>
                <meta charset="utf-8">
                <meta name="viewport" content="width=device-width, initial-scale=1">
                <style>
                    body {
                        font-family: Arial, sans-serif;
                        margin: 0;
                        padding: 20px;
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                        color: white;
                        text-align: center;
                    }
                    .container {
                        max-width: 800px;
                        margin: 0 auto;
                    }
                    h1 { font-size: 2.5em; margin-bottom: 20px; }
                    .card {
                        background: rgba(255,255,255,0.1);
                        padding: 30px;
                        border-radius: 10px;
                        backdrop-filter: blur(10px);
                        margin-top: 30px;
                    }
                    .btn {
                        background: white;
                        color: #764ba2;
                        padding: 12px 30px;
                        border: none;
                        border-radius: 25px;
                        font-size: 16px;
                        font-weight: bold;
                        cursor: pointer;
                        margin: 10px;
                        transition: transform 0.3s;
                    }
                    .btn:hover { transform: scale(1.05); }
                </style>
            </head>
            <body>
                <div class="container">
                    <h1>🚀 WhatsApp Bot</h1>
                    <div class="card">
                        <h2>Servidor rodando com sucesso!</h2>
                        <p>O arquivo dashboard.html não foi encontrado.</p>
                        <p>Crie o arquivo em: public/dashboard.html</p>
                        <button class="btn" onclick="location.reload()">🔄 Recarregar</button>
                    </div>
                </div>
            </body>
            </html>
        `);
    }
});

// ============================================
// SISTEMA DE KEEP-ALIVE PARA RAILWAY
// ============================================
function startKeepAlive() {
    console.log("💓 Iniciando sistema keep-alive...");
    
    // Ping a cada 5 minutos para manter o servidor ativo
    setInterval(() => {
        const url = `http://localhost:${PORT}/ping`;
        fetch(url)
            .then(() => console.log(`💓 Keep-alive ping enviado às ${new Date().toLocaleTimeString()}`))
            .catch(err => console.log(`⚠️ Keep-alive falhou: ${err.message}`));
    }, 300000); // 5 minutos
    
    // Verificar conexões a cada 10 minutos
    setInterval(() => {
        const connections = whatsappManager.getAllConnections();
        const activeConnections = whatsappManager.getActiveConnections();
        
        console.log(`📊 Status das conexões: ${activeConnections.length}/${connections.length} ativas`);
        
        // Se não houver conexões ativas mas houver conexões totais, tentar reconectar
        if (activeConnections.length === 0 && connections.length > 0) {
            console.log("⚠️ Nenhuma conexão ativa detectada. Tentando reconectar...");
            connections.forEach(conn => {
                if (conn.status !== "connected") {
                    whatsappManager.forceNewConnection(conn.id).catch(err => {
                        console.log(`❌ Falha ao reconectar ${conn.id}:`, err.message);
                    });
                }
            });
        }
    }, 600000); // 10 minutos
}

// ============================================
// INICIAR SERVIDOR (MODIFICADO)
// ============================================
const server = app.listen(PORT, '0.0.0.0', () => { // Importante: ouvir em 0.0.0.0 no Railway
    console.log(`\n══════════════════════════════════════════`);
    console.log(`🚀 WHATSAPP BOT INICIADO!`);
    console.log(`📡 Porta: ${PORT}`);
    console.log(`🌐 Acesse: http://localhost:${PORT}`);
    console.log(`🌍 Railway URL: https://whatsapp-bot-production-5443.up.railway.app`);
    console.log(`📱 Suporte multi-conexões`);
    console.log(`📊 Upload de planilhas (Excel/CSV) com MÚLTIPLOS TELEFONES`);
    console.log(`📝 Templates personalizados`);
    console.log(`⏰ Delays: 5-30 segundos entre mensagens`);
    console.log(`📞 Suporte a múltiplos números por contato (enviando para TODOS)`);
    console.log(`🧹 /api/clear - Limpa apenas planilhas (NÃO conexões)`);
    console.log(`🩺 /health - Health check para Railway`);
    console.log(`══════════════════════════════════════════\n`);
    
    const dirs = ['public', 'uploads', 'templates'];
    dirs.forEach(dir => {
        const dirPath = path.join(__dirname, dir);
        if (!fs.existsSync(dirPath)) {
            fs.mkdirSync(dirPath, { recursive: true });
            console.log(`📁 Pasta criada: ${dir}`);
        }
    });
    
    // Iniciar keep-alive
    startKeepAlive();
    
    // Iniciar conexão automática com delay maior
    setTimeout(async () => {
        try {
            console.log(`🔄 Iniciando conexão automática...`);
            await whatsappManager.createConnection();
            console.log(`✅ Conexão automática iniciada!`);
        } catch (error) {
            console.log(`⚠️ Conexão automática falhou, mas o bot está rodando`);
        }
    }, 5000); // Aumentei para 5 segundos
});

// Timeout do servidor (importante para Railway)
server.timeout = 120000; // 2 minutos

// ============================================
// TRATAMENTO DE ERROS NÃO CAPTURADOS
// ============================================
process.on("uncaughtException", (error) => {
    console.error("❌ Erro não capturado:", error);
    // Não encerrar o processo, apenas logar
});

process.on("unhandledRejection", (error) => {
    console.error("❌ Promise rejeitada não tratada:", error);
    // Não encerrar o processo, apenas logar
});

process.on("SIGINT", async () => {
    console.log(`\n🔴 Encerrando bot...`);
    const connections = whatsappManager.getAllConnections();
    
    for (const connection of connections) {
        await whatsappManager.disconnectConnection(connection.id);
    }
    
    console.log(`✅ Bot encerrado corretamente`);
    process.exit(0);
});

process.on("SIGTERM", async () => {
    console.log(`\n🔴 Encerrando bot...`);
    const connections = whatsappManager.getAllConnections();
    
    for (const connection of connections) {
        await whatsappManager.disconnectConnection(connection.id);
    }
    
    console.log(`✅ Bot encerrado corretamente`);
    process.exit(0);
});

module.exports = app;