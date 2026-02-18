// ============================================
// BOT WHATSAPP - MULTI CONEXÕES COM PLANILHAS E TEMPLATES
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

// SEGUNDO: Criar a aplicação Express
const app = express();
const PORT = process.env.PORT || 3000;

// TERCEIRO: Configurar o Express ANTES de qualquer outra coisa
app.use(express.json());
app.use(express.static("public"));
app.use(express.urlencoded({ extended: true }));
app.use('/uploads', express.static('uploads'));

// QUARTO: Carregar o Baileys (depois do Express)
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
// GERENCIADOR DE PLANILHAS (CORRIGIDO)
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
// GERENCIADOR DE CONEXÕES WHATSAPP
// ============================================
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
        const authDir = path.join(__dirname, `auth_info_${id}`);
        
        console.log(`🔄 [${id}] Criando conexão WhatsApp...`);
        
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
                qrGeneratedAt: null
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
                    console.log(`🔄 [${id}] Tentando reconectar em 3 segundos...`);
                    setTimeout(() => {
                        console.log(`🔄 [${id}] Iniciando reconexão...`);
                        this.createConnection(id).catch(err => {
                            console.error(`❌ [${id}] Falha na reconexão:`, err.message);
                        });
                    }, 3000);
                }
            } 
            else if (connection === "open") {
                console.log(`✅ [${id}] CONECTADO COM SUCESSO!`);
                connectionData.status = "connected";
                connectionData.qrCode = null;
                connectionData.lastQR = null;
                connectionData.isConnected = true;
                
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
// ROTAS DA API
// ============================================

// Rota de teste para verificar se o servidor está rodando
app.get("/api/test", (req, res) => {
    res.json({ success: true, message: "Servidor está rodando!" });
});

// Rota de upload de planilha
app.post("/api/spreadsheet/upload", (req, res) => {
    upload.single('file')(req, res, async (err) => {
        if (err) {
            console.error('❌ Erro no upload:', err);
            return res.status(400).json({
                success: false,
                message: `Erro no upload: ${err.message}`
            });
        }

        try {
            if (!req.file) {
                return res.status(400).json({
                    success: false,
                    message: "Nenhum arquivo enviado"
                });
            }

            console.log(`📁 Arquivo recebido:`, {
                originalname: req.file.originalname,
                size: req.file.size,
                path: req.file.path
            });

            const options = {
                phoneColumn: req.body.phoneColumn || 'telefone',
                nameColumn: req.body.nameColumn || 'nome',
                skipFirstRow: req.body.skipFirstRow !== 'false',
                customColumns: req.body.customColumns ? req.body.customColumns.split(',').map(c => c.trim()) : []
            };

            console.log(`⚙️ Opções:`, options);

            const spreadsheetData = await spreadsheetManager.processSpreadsheet(req.file.path, options);
            
            // Preview mostrando múltiplos telefones
            const previewContacts = spreadsheetData.contacts.slice(0, 10).map(c => ({
                name: c.name || 'Sem nome',
                phones: c.phones.map(p => ({
                    original: p.original,
                    formatted: p.formatted,
                    isValid: p.isValid,
                    column: p.column
                })),
                validPhonesCount: c.validPhones.length,
                allPhonesCount: c.phones.length
            }));

            res.json({
                success: true,
                message: `Planilha processada com sucesso. ${spreadsheetData.validContacts} contatos com ${spreadsheetData.totalValidPhones} telefones válidos.`,
                data: {
                    id: spreadsheetData.id,
                    fileName: spreadsheetData.fileName,
                    totalContacts: spreadsheetData.totalContacts,
                    validContacts: spreadsheetData.validContacts,
                    invalidContacts: spreadsheetData.invalidContacts,
                    totalPhones: spreadsheetData.totalPhones,
                    totalValidPhones: spreadsheetData.totalValidPhones,
                    contacts: previewContacts,
                    totalPreview: previewContacts.length,
                    columns: spreadsheetData.columns
                }
            });

        } catch (error) {
            console.error('❌ Erro ao processar planilha:', error);
            
            if (req.file && req.file.path && fs.existsSync(req.file.path)) {
                try {
                    fs.unlinkSync(req.file.path);
                } catch (err) {}
            }

            res.status(500).json({
                success: false,
                message: `Erro ao processar planilha: ${error.message}`
            });
        }
    });
});

// Listar planilhas
app.get("/api/spreadsheets", (req, res) => {
    const spreadsheets = spreadsheetManager.getAllSpreadsheets().map(s => ({
        id: s.id,
        fileName: s.fileName,
        uploadedAt: s.uploadedAt,
        totalContacts: s.totalContacts,
        validContacts: s.validContacts,
        invalidContacts: s.invalidContacts,
        totalPhones: s.totalPhones,
        totalValidPhones: s.totalValidPhones
    }));
    
    res.json({
        success: true,
        data: spreadsheets
    });
});

// Obter contatos de uma planilha
app.get("/api/spreadsheet/:spreadsheetId/contacts", (req, res) => {
    const { spreadsheetId } = req.params;
    const { page = 1, limit = 100 } = req.query;
    
    const spreadsheet = spreadsheetManager.getContacts(spreadsheetId);
    
    if (!spreadsheet) {
        return res.status(404).json({
            success: false,
            message: "Planilha não encontrada"
        });
    }

    const startIndex = (parseInt(page) - 1) * parseInt(limit);
    const endIndex = startIndex + parseInt(limit);
    const paginatedContacts = spreadsheet.contacts.slice(startIndex, endIndex).map(c => ({
        name: c.name,
        phones: c.phones,
        mainPhone: c.mainPhone,
        allPhones: c.allPhones,
        validPhones: c.validPhones,
        invalidPhones: c.invalidPhones,
        hasValidPhone: c.hasValidPhone
    }));

    res.json({
        success: true,
        data: {
            id: spreadsheet.id,
            fileName: spreadsheet.fileName,
            totalContacts: spreadsheet.totalContacts,
            validContacts: spreadsheet.validContacts,
            invalidContacts: spreadsheet.invalidContacts,
            totalPhones: spreadsheet.totalPhones,
            totalValidPhones: spreadsheet.totalValidPhones,
            contacts: paginatedContacts,
            currentPage: parseInt(page),
            totalPages: Math.ceil(spreadsheet.contacts.length / parseInt(limit)),
            limit: parseInt(limit)
        }
    });
});

// Deletar planilha
app.delete("/api/spreadsheet/:spreadsheetId", (req, res) => {
    const { spreadsheetId } = req.params;
    const deleted = spreadsheetManager.deleteSpreadsheet(spreadsheetId);
    
    if (deleted) {
        res.json({
            success: true,
            message: "Planilha deletada com sucesso"
        });
    } else {
        res.status(404).json({
            success: false,
            message: "Planilha não encontrada"
        });
    }
});

// Templates
app.get("/api/templates", (req, res) => {
    const templates = templateManager.getAllTemplates();
    res.json({
        success: true,
        data: templates
    });
});

app.post("/api/templates", (req, res) => {
    try {
        const { name, message } = req.body;
        
        if (!name || !message) {
            return res.status(400).json({
                success: false,
                message: "Nome e mensagem são obrigatórios"
            });
        }

        const template = templateManager.createTemplate({ name, message });
        
        res.json({
            success: true,
            message: "Template criado com sucesso",
            data: template
        });
    } catch (error) {
        res.status(500).json({
            success: false,
            message: `Erro: ${error.message}`
        });
    }
});

app.delete("/api/templates/:templateId", (req, res) => {
    try {
        const { templateId } = req.params;
        const deleted = templateManager.deleteTemplate(templateId);
        
        if (deleted) {
            res.json({
                success: true,
                message: "Template deletado com sucesso"
            });
        } else {
            res.status(404).json({
                success: false,
                message: "Template não encontrado"
            });
        }
    } catch (error) {
        res.status(500).json({
            success: false,
            message: `Erro: ${error.message}`
        });
    }
});

// Status das conexões WhatsApp
app.get("/api/status", (req, res) => {
    const connections = whatsappManager.getAllConnections();
    const active = whatsappManager.getActiveConnections().length;
    const withQR = whatsappManager.getConnectionsWithQR().length;
    
    res.json({
        success: true,
        data: {
            total: connections.length,
            active: active,
            with_qr: withQR,
            connections: connections.map(conn => ({
                id: conn.id,
                status: conn.status,
                connected: conn.isConnected,
                user: conn.userInfo,
                has_qr: !!conn.qrCode
            }))
        }
    });
});

// Criar nova conexão WhatsApp
app.post("/api/connect", async (req, res) => {
    try {
        const { connectionId } = req.body;
        
        console.log(`🔗 Solicitando conexão: ${connectionId || "nova"}`);
        const connection = await whatsappManager.createConnection(connectionId);
        
        res.json({
            success: true,
            message: "Conexão iniciada. Aguarde QR Code.",
            connectionId: connection.id,
            hasQR: !!connection.qrCode
        });
        
    } catch (error) {
        console.error("❌ Erro na conexão:", error.message);
        res.status(500).json({
            success: false,
            message: `Erro: ${error.message}`,
            connectionId: null
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
        
        if (connection.qrCode) {
            res.json({
                success: true,
                qr: connection.qrCode,
                connectionId: connection.id,
                timestamp: connection.qrGeneratedAt
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
                message: "QR Code não disponível. Aguarde ou reinicie a conexão.",
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

// Forçar nova conexão
app.post("/api/restart", async (req, res) => {
    try {
        const { connectionId } = req.body;
        
        if (!connectionId) {
            return res.status(400).json({
                success: false,
                message: "connectionId é obrigatório"
            });
        }
        
        console.log(`🔄 Reiniciando conexão: ${connectionId}`);
        const connection = await whatsappManager.forceNewConnection(connectionId);
        
        res.json({
            success: true,
            message: "Conexão reiniciada",
            connectionId: connection.id
        });
        
    } catch (error) {
        res.status(500).json({
            success: false,
            message: `Erro: ${error.message}`
        });
    }
});

// Desconectar
app.post("/api/disconnect", async (req, res) => {
    try {
        const { connectionId } = req.body;
        
        if (!connectionId) {
            return res.status(400).json({
                success: false,
                message: "connectionId é obrigatório"
            });
        }
        
        await whatsappManager.disconnectConnection(connectionId);
        
        res.json({
            success: true,
            message: "Conexão desconectada"
        });
        
    } catch (error) {
        res.status(500).json({
            success: false,
            message: `Erro: ${error.message}`
        });
    }
});

// Enviar mensagens (CORRIGIDO)
app.post("/api/send", async (req, res) => {
    try {
        const { 
            connectionId, 
            numbers, 
            message, 
            delay = true,
            sendToAllPhones = true // Agora padrão é true para enviar para todos
        } = req.body;
        
        if (!numbers || !Array.isArray(numbers) || numbers.length === 0) {
            return res.status(400).json({
                success: false,
                message: "Lista de números é obrigatória"
            });
        }
        
        if (!message || message.trim() === "") {
            return res.status(400).json({
                success: false,
                message: "Mensagem é obrigatória"
            });
        }
        
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
        
        console.log(`📤 [${connection.id}] Enviando mensagens...`);
        
        const results = [];
        let successCount = 0;
        let totalAttempts = 0;
        
        for (let i = 0; i < numbers.length; i++) {
            const item = numbers[i];
            
            // Verificar se é um contato com múltiplos telefones ou apenas um número
            if (typeof item === 'object' && item.phones && item.phones.length > 0) {
                // É um contato com múltiplos telefones
                const phonesToSend = sendToAllPhones ? item.validPhones : [item.validPhones[0]];
                
                console.log(`📞 [${connection.id}] Contato: ${item.name || 'Sem nome'} - ${phonesToSend.length} telefone(s) para enviar`);
                
                for (let j = 0; j < phonesToSend.length; j++) {
                    const phoneObj = phonesToSend[j];
                    if (!phoneObj || !phoneObj.formatted) continue;
                    
                    totalAttempts++;
                    
                    try {
                        // Personalizar mensagem com o nome do contato se tiver {{nome}} no template
                        let personalizedMessage = message;
                        if (item.name && message.includes('{{nome}}')) {
                            personalizedMessage = message.replace(/\{\{nome\}\}/g, item.name);
                        }
                        if (item.name && message.includes('{{name}}')) {
                            personalizedMessage = message.replace(/\{\{name\}\}/g, item.name);
                        }
                        
                        const jid = `${phoneObj.formatted}@s.whatsapp.net`;
                        await connection.sock.sendMessage(jid, { text: personalizedMessage });
                        
                        results.push({
                            name: item.name || 'Sem nome',
                            phone: phoneObj.original,
                            formattedPhone: phoneObj.formatted,
                            phoneIndex: j + 1,
                            totalPhonesForContact: item.validPhones.length,
                            success: true,
                            message: "Enviada com sucesso"
                        });
                        successCount++;
                        
                        console.log(`✅ [${connection.id}] ${item.name || 'Sem nome'} - Telefone ${j + 1}/${phonesToSend.length}: ${phoneObj.original}`);
                        
                    } catch (error) {
                        console.error(`❌ [${connection.id}] Erro para ${item.name} - Telefone ${j + 1}:`, error.message);
                        results.push({
                            name: item.name || 'Sem nome',
                            phone: phoneObj.original,
                            formattedPhone: phoneObj.formatted,
                            phoneIndex: j + 1,
                            totalPhonesForContact: item.validPhones.length,
                            success: false,
                            error: error.message
                        });
                    }
                    
                    if (delay && (i < numbers.length - 1 || j < phonesToSend.length - 1)) {
                        const waitTime = getRandomDelay();
                        await new Promise(resolve => setTimeout(resolve, waitTime));
                    }
                }
            } else {
                // É apenas um número de telefone simples
                totalAttempts++;
                const number = item.toString().replace(/\D/g, "");
                
                if (number.length < 10) {
                    results.push({
                        number: item,
                        success: false,
                        error: "Número inválido"
                    });
                    continue;
                }
                
                try {
                    const jid = `${number}@s.whatsapp.net`;
                    await connection.sock.sendMessage(jid, { text: message });
                    
                    results.push({
                        number: item,
                        formattedNumber: number,
                        success: true,
                        message: "Enviada com sucesso"
                    });
                    successCount++;
                    
                    console.log(`✅ [${connection.id}] ${i+1}/${numbers.length}: ${number}`);
                    
                } catch (error) {
                    console.error(`❌ [${connection.id}] Erro para ${number}:`, error.message);
                    results.push({
                        number: item,
                        formattedNumber: number,
                        success: false,
                        error: error.message
                    });
                }
                
                if (delay && i < numbers.length - 1) {
                    const waitTime = getRandomDelay();
                    await new Promise(resolve => setTimeout(resolve, waitTime));
                }
            }
        }
        
        console.log(`🎉 [${connection.id}] Envio concluído: ${successCount}/${totalAttempts} mensagens enviadas`);
        
        res.json({
            success: true,
            totalContacts: numbers.length,
            totalAttempts: totalAttempts,
            sent: successCount,
            failed: totalAttempts - successCount,
            results: results,
            connection: connection.id,
            user: connection.userInfo
        });
        
    } catch (error) {
        console.error("❌ Erro no envio:", error);
        res.status(500).json({
            success: false,
            message: `Erro: ${error.message}`
        });
    }
});

// Envio em lote com planilha (CORRIGIDO)
app.post("/api/send/bulk", async (req, res) => {
    try {
        const { 
            connectionId,
            spreadsheetId,
            templateId,
            customMessage,
            delay = true,
            preview = false,
            sendToAllPhones = true // Agora padrão é true para enviar para todos
        } = req.body;

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

        const spreadsheet = spreadsheetManager.getContacts(spreadsheetId);
        if (!spreadsheet) {
            return res.status(404).json({
                success: false,
                message: "Planilha não encontrada"
            });
        }

        const contacts = spreadsheet.contacts;
        if (contacts.length === 0) {
            return res.status(400).json({
                success: false,
                message: "Nenhum contato válido na planilha"
            });
        }

        if (preview) {
            const previewContacts = contacts.slice(0, 5).map(contact => ({
                name: contact.name || 'Sem nome',
                phones: contact.phones.map(p => ({
                    original: p.original,
                    formatted: p.formatted,
                    isValid: p.isValid
                })),
                totalPhones: contact.phones.length,
                validPhones: contact.validPhones.length
            }));

            let previewMessage = '';
            if (templateId) {
                const template = templateManager.getTemplate(templateId);
                if (template) {
                    previewMessage = templateManager.renderTemplate(templateId, contacts[0] || {});
                }
            } else if (customMessage) {
                previewMessage = customMessage;
            }

            return res.json({
                success: true,
                preview: true,
                totalContacts: contacts.length,
                totalPhones: spreadsheet.totalPhones,
                totalValidPhones: spreadsheet.totalValidPhones,
                previewContacts,
                previewMessage: previewMessage || 'Nenhuma mensagem definida',
                template: templateId ? templateManager.getTemplate(templateId) : null,
                sendToAllPhones: sendToAllPhones
            });
        }

        console.log(`📤 [${connection.id}] Enviando lote de ${contacts.length} contatos...`);
        console.log(`📞 Modo: ${sendToAllPhones ? 'Enviar para TODOS os telefones' : 'Enviar apenas para o primeiro telefone'}`);
        
        const results = [];
        let successCount = 0;
        let totalAttempts = 0;
        let contactsWithMultiplePhones = 0;

        for (let i = 0; i < contacts.length; i++) {
            const contact = contacts[i];
            
            // Determinar quais telefones enviar
            const phonesToSend = sendToAllPhones ? contact.validPhones : [contact.validPhones[0]];
            
            if (contact.validPhones.length > 1) {
                contactsWithMultiplePhones++;
            }
            
            console.log(`📞 [${connection.id}] Contato ${i+1}/${contacts.length}: ${contact.name || 'Sem nome'} - ${phonesToSend.length} telefone(s) válido(s)`);
            
            for (let j = 0; j < phonesToSend.length; j++) {
                const phoneObj = phonesToSend[j];
                if (!phoneObj) continue;
                
                totalAttempts++;
                
                try {
                    let message;
                    if (templateId) {
                        message = templateManager.renderTemplate(templateId, contact);
                    } else if (customMessage) {
                        message = customMessage;
                        // Substituir variáveis no formato {{nome}}
                        Object.keys(contact).forEach(key => {
                            message = message.replace(new RegExp(`\\{\\{${key}\\}\\}`, 'g'), contact[key] || '');
                        });
                    } else {
                        message = `Olá ${contact.name || ''}! Mensagem automática.`;
                    }

                    const jid = `${phoneObj.formatted}@s.whatsapp.net`;
                    await connection.sock.sendMessage(jid, { text: message });
                    
                    results.push({
                        name: contact.name || 'Sem nome',
                        phone: phoneObj.original,
                        formattedPhone: phoneObj.formatted,
                        phoneIndex: j + 1,
                        totalPhonesForContact: contact.validPhones.length,
                        success: true,
                        message: "Enviada com sucesso"
                    });
                    successCount++;
                    
                    console.log(`✅ [${connection.id}] ${i+1}/${contacts.length} - ${contact.name || 'Sem nome'} (Tel ${j + 1}/${phonesToSend.length}): ${phoneObj.original}`);
                    
                } catch (error) {
                    console.error(`❌ [${connection.id}] Erro para ${contact.name || 'Sem nome'} - Tel ${j + 1}:`, error.message);
                    results.push({
                        name: contact.name || 'Sem nome',
                        phone: phoneObj.original,
                        formattedPhone: phoneObj.formatted,
                        phoneIndex: j + 1,
                        totalPhonesForContact: contact.validPhones.length,
                        success: false,
                        error: error.message
                    });
                }
                
                if (delay && (i < contacts.length - 1 || j < phonesToSend.length - 1)) {
                    const waitTime = getRandomDelay();
                    await new Promise(resolve => setTimeout(resolve, waitTime));
                }
            }
        }

        console.log(`🎉 [${connection.id}] Envio concluído:`);
        console.log(`   📊 Contatos processados: ${contacts.length}`);
        console.log(`   📞 Contatos com múltiplos telefones: ${contactsWithMultiplePhones}`);
        console.log(`   📨 Total de tentativas: ${totalAttempts}`);
        console.log(`   ✅ Sucessos: ${successCount}`);
        console.log(`   ❌ Falhas: ${totalAttempts - successCount}`);

        res.json({
            success: true,
            totalContacts: contacts.length,
            contactsWithMultiplePhones: contactsWithMultiplePhones,
            totalAttempts: totalAttempts,
            sent: successCount,
            failed: totalAttempts - successCount,
            results: results.slice(0, 50),
            totalResults: results.length,
            connection: connection.id,
            user: connection.userInfo,
            spreadsheet: {
                id: spreadsheet.id,
                fileName: spreadsheet.fileName,
                totalPhones: spreadsheet.totalPhones,
                totalValidPhones: spreadsheet.totalValidPhones
            },
            sendToAllPhones: sendToAllPhones
        });

    } catch (error) {
        console.error("❌ Erro no envio em lote:", error);
        res.status(500).json({
            success: false,
            message: `Erro: ${error.message}`
        });
    }
});

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
// INICIAR SERVIDOR
// ============================================
app.listen(PORT, () => {
    console.log(`\n══════════════════════════════════════════`);
    console.log(`🚀 WHATSAPP BOT INICIADO!`);
    console.log(`📡 Porta: ${PORT}`);
    console.log(`🌐 Acesse: http://localhost:${PORT}`);
    console.log(`📱 Suporte multi-conexões`);
    console.log(`📊 Upload de planilhas (Excel/CSV) com MÚLTIPLOS TELEFONES`);
    console.log(`📝 Templates personalizados`);
    console.log(`⏰ Delays: 5-30 segundos entre mensagens`);
    console.log(`📞 Suporte a múltiplos números por contato (enviando para TODOS)`);
    console.log(`🧹 /api/clear - Limpa apenas planilhas (NÃO conexões)`);
    console.log(`══════════════════════════════════════════\n`);
    
    const dirs = ['public', 'uploads', 'templates'];
    dirs.forEach(dir => {
        const dirPath = path.join(__dirname, dir);
        if (!fs.existsSync(dirPath)) {
            fs.mkdirSync(dirPath, { recursive: true });
            console.log(`📁 Pasta criada: ${dir}`);
        }
    });
    
    setTimeout(async () => {
        try {
            console.log(`🔄 Iniciando conexão automática...`);
            await whatsappManager.createConnection();
            console.log(`✅ Conexão automática iniciada!`);
        } catch (error) {
            console.log(`⚠️ Conexão automática falhou, mas o bot está rodando`);
        }
    }, 2000);
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