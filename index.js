const { default: makeWASocket, useMultiFileAuthState, downloadMediaMessage, DisconnectReason } = require('@whiskeysockets/baileys');
const express = require('express');
const QRCode = require('qrcode');
const pino = require('pino');
const fs = require('fs');
const xlsx = require('xlsx');
const { GoogleGenAI, Type } = require('@google/genai');
const { google } = require('googleapis');

const app = express();
const PORT = process.env.PORT || 10000;

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
const CALENDAR_ID = process.env.CALENDAR_ID || 'primary';
const GOOGLE_CREDENTIALS_JSON = process.env.GOOGLE_CREDENTIALS_JSON;
const EXCEL_PATH = './base_dados.xlsx';
const MODEL_NAME = 'gemini-3.8-flash';

let qrCodeAtual = null;
let statusConexao = 'Aguardando geração do QR Code...';

app.get('/', async (req, res) => {
    if (statusConexao === 'Conectado!') {
        return res.send(`
            <div style="font-family: sans-serif; text-align: center; margin-top: 50px;">
                <h1 style="color: #25D366;">✅ WhatsApp Conectado com Sucesso!</h1>
                <p>O seu Agente IA Privado já está ativo.</p>
            </div>
        `);
    }
    if (qrCodeAtual) {
        const qrImage = await QRCode.toDataURL(qrCodeAtual);
        return res.send(`
            <div style="font-family: sans-serif; text-align: center; margin-top: 40px;">
                <h2>Escaneie o QR Code abaixo com o seu WhatsApp</h2>
                <img src="${qrImage}" style="width: 300px; height: 300px; border: 2px solid #ccc; border-radius: 10px;"/>
            </div>
        `);
    }
    res.send(`<div style="font-family: sans-serif; text-align: center; margin-top: 50px;"><h2>Status: ${statusConexao}</h2></div>`);
});

app.listen(PORT, '0.0.0.0', () => console.log(`Servidor na porta ${PORT}`));

function obterAgenda() {
    if (!GOOGLE_CREDENTIALS_JSON) return null;
    const credentials = JSON.parse(GOOGLE_CREDENTIALS_JSON);
    const auth = new google.auth.GoogleAuth({
        credentials,
        scopes: ['https://www.googleapis.com/auth/calendar']
    });
    return google.calendar({ version: 'v3', auth });
}

async function listarEventosAgenda(dias = 7) {
    try {
        const calendar = obterAgenda();
        if (!calendar) return "Agenda ainda não configurada nas variáveis de ambiente.";
        const agora = new Date();
        const limite = new Date();
        limite.setDate(agora.getDate() + dias);

        const res = await calendar.events.list({
            calendarId: CALENDAR_ID,
            timeMin: agora.toISOString(),
            timeMax: limite.toISOString(),
            singleEvents: true,
            orderBy: 'startTime',
        });

        const eventos = res.data.items;
        if (!eventos || eventos.length === 0) return `Nenhum compromisso para os próximos ${dias} dias.`;
        return eventos.map(e => `- ${e.start.dateTime || e.start.date}: ${e.summary}`).join('\n');
    } catch (err) {
        return `Erro ao consultar agenda: ${err.message}`;
    }
}

async function criarEventoAgenda(titulo, dataHoraInicio, duracaoMinutos = 60, descricao = "") {
    try {
        const calendar = obterAgenda();
        if (!calendar) return "Agenda ainda não configurada.";
        const inicio = new Date(dataHoraInicio);
        const fim = new Date(inicio.getTime() + duracaoMinutos * 60000);

        await calendar.events.insert({
            calendarId: CALENDAR_ID,
            requestBody: {
                summary: titulo,
                description: descricao,
                start: { dateTime: inicio.toISOString(), timeZone: 'America/Sao_Paulo' },
                end: { dateTime: fim.toISOString(), timeZone: 'America/Sao_Paulo' }
            }
        });
        return `✅ Compromisso "${titulo}" agendado com sucesso para ${dataHoraInicio}!`;
    } catch (err) {
        return `Erro ao criar evento: ${err.message}`;
    }
}

function consultarPlanilhaExcel(termoBusca = "") {
    if (!fs.existsSync(EXCEL_PATH)) {
        return "Nenhuma folha de cálculo foi enviada ainda. Envie um ficheiro .xlsx aqui no chat do WhatsApp.";
    }
    try {
        const workbook = xlsx.readFile(EXCEL_PATH);
        let resultado = [];

        workbook.SheetNames.forEach(nomeAba => {
            const dados = xlsx.utils.sheet_to_json(workbook.Sheets[nomeAba], { defval: "" });
            let filtrados = dados;

            if (termoBusca) {
                const termo = termoBusca.toLowerCase();
                filtrados = dados.filter(linha =>
                    Object.values(linha).some(val => String(val).toLowerCase().includes(termo))
                );
            } else {
                filtrados = dados.slice(0, 40);
            }

            const colunas = dados.length > 0 ? Object.keys(dados[0]).join(', ') : 'Vazia';
            resultado.push(`--- ABA: ${nomeAba} (Total: ${dados.length} linhas | Colunas: ${colunas}) ---\n` + JSON.stringify(filtrados, null, 2));
        });

        return resultado.join('\n\n').substring(0, 14000);
    } catch (err) {
        return `Erro ao ler Excel: ${err.message}`;
    }
}

const ferramentas = [{
    functionDeclarations: [
        {
            name: "consultarPlanilhaExcel",
            description: "Lê e pesquisa dados na folha de cálculo Excel enviada pelo utilizador no WhatsApp.",
            parameters: {
                type: Type.OBJECT,
                properties: {
                    termoBusca: { type: Type.STRING, description: "Palavra-chave, loja, nome ou estado para filtrar (ou vazio para ver resumo geral)." }
                }
            }
        },
        {
            name: "listarEventosAgenda",
            description: "Lista os compromissos do Google Calendar dos próximos dias.",
            parameters: {
                type: Type.OBJECT,
                properties: {
                    dias: { type: Type.NUMBER, description: "Quantidade de dias para procurar (padrão 7)." }
                }
            }
        },
        {
            name: "criarEventoAgenda",
            description: "Cria um novo compromisso no Google Calendar.",
            parameters: {
                type: Type.OBJECT,
                properties: {
                    titulo: { type: Type.STRING, description: "Título do compromisso" },
                    dataHoraInicio: { type: Type.STRING, description: "Data e hora no formato ISO YYYY-MM-DDTHH:MM:SS" },
                    duracaoMinutos: { type: Type.NUMBER, description: "Duração em minutos (padrão 60)" },
                    descricao: { type: Type.STRING, description: "Descrição opcional" }
                },
                required: ["titulo", "dataHoraInicio"]
            }
        }
    ]
}];

// Função auxiliar que repete automaticamente se o servidor da Google der erro 503 (alta procura)
async function chamarGeminiComRetry(parametros, tentativas = 3) {
    for (let i = 0; i < tentativas; i++) {
        try {
            return await ai.models.generateContent(parametros);
        } catch (err) {
            if ((err.message?.includes('503') || err.message?.includes('UNAVAILABLE')) && i < tentativas - 1) {
                await new Promise(r => setTimeout(r, 2500));
                continue;
            }
            throw err;
        }
    }
}

async function processarComIA(textoUsuario) {
    try {
        const agora = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
        const config = {
            systemInstruction: `Você é um assistente pessoal inteligente no WhatsApp. Data/hora atual: ${agora}. Você pode ler a planilha Excel enviada, consultar/agendar no Google Calendar e dar sugestões práticas sobre qualquer pergunta. Responda de forma objetiva e formatada para WhatsApp.`,
            tools: ferramentas,
            temperature: 0.4
        };

        let response = await chamarGeminiComRetry({
            model: MODEL_NAME,
            contents: [{ role: 'user', parts: [{ text: textoUsuario }] }],
            config
        });

        if (response.functionCalls && response.functionCalls.length > 0) {
            const call = response.functionCalls[0];
            let resultadoFerramenta = "";

            if (call.name === "consultarPlanilhaExcel") {
                resultadoFerramenta = consultarPlanilhaExcel(call.args.termoBusca || "");
            } else if (call.name === "listarEventosAgenda") {
                resultadoFerramenta = await listarEventosAgenda(call.args.dias || 7);
            } else if (call.name === "criarEventoAgenda") {
                resultadoFerramenta = await criarEventoAgenda(call.args.titulo, call.args.dataHoraInicio, call.args.duracaoMinutos, call.args.descricao);
            }

            response = await chamarGeminiComRetry({
                model: MODEL_NAME,
                contents: [
                    { role: 'user', parts: [{ text: textoUsuario }] },
                    response.candidates[0].content,
                    { role: 'user', parts: [{ functionResponse: { name: call.name, response: { result: resultadoFerramenta } } }] }
                ],
                config
            });
        }

        return response.text || "Processado com sucesso.";
    } catch (err) {
        return `⚠️ A IA está com alta procura neste segundo. Tente novamente em instantes.`;
    }
}

async function iniciarWhatsApp() {
    const { state, saveCreds } = await useMultiFileAuthState('./auth_info');
    const sock = makeWASocket({
        auth: state,
        logger: pino({ level: 'silent' }),
        printQRInTerminal: true
    });

    sock.ev.on('creds.update', saveCreds);

    sock.ev.on('connection.update', (update) => {
        const { connection, lastDisconnect, qr } = update;
        if (qr) {
            qrCodeAtual = qr;
            statusConexao = 'Aguardando leitura do QR Code...';
        }
        if (connection === 'open') {
            qrCodeAtual = null;
            statusConexao = 'Conectado!';
            console.log('WhatsApp conectado!');
        }
        if (connection === 'close') {
            const deveReconectar = lastDisconnect?.error?.output?.statusCode !== DisconnectReason.loggedOut;
            statusConexao = 'Desconectado. Reconectando...';
            if (deveReconectar) iniciarWhatsApp();
        }
    });

    sock.ev.on('messages.upsert', async ({ messages }) => {
        for (const msg of messages) {
            if (!msg.message) continue;

            // REGRA DE OURO: Ignora 100% das mensagens enviadas por terceiros! Só aceita mensagens enviadas por si mesmo.
            if (!msg.key.fromMe) continue;

            const remetente = msg.key.remoteJid;
            if (!remetente || remetente.includes('@g.us') || remetente === 'status@broadcast') continue;

            const texto = msg.message.conversation ||
                          msg.message.extendedTextMessage?.text ||
                          msg.message.ephemeralMessage?.message?.extendedTextMessage?.text ||
                          msg.message.ephemeralMessage?.message?.conversation || "";

            if (texto.startsWith('🤖 *Assistente IA:*') || texto.startsWith('📊 *Planilha')) continue;

            const doc = msg.message.documentMessage || msg.message.documentWithCaptionMessage?.message?.documentMessage;
            const meuNumero = sock.user?.id ? sock.user.id.split(':')[0] + '@s.whatsapp.net' : remetente;
            const destino = remetente.includes('@lid') ? meuNumero : remetente;

            if (doc && (doc.fileName?.endsWith('.xlsx') || doc.fileName?.endsWith('.xls'))) {
                const buffer = await downloadMediaMessage(msg, 'buffer', {});
                fs.writeFileSync(EXCEL_PATH, buffer);
                const wb = xlsx.readFile(EXCEL_PATH);
                const primeiraAba = wb.Sheets[wb.SheetNames[0]];
                const linhas = xlsx.utils.sheet_to_json(primeiraAba);
                const colunas = linhas.length > 0 ? Object.keys(linhas[0]).join(', ') : 'Nenhuma';

                await sock.sendMessage(destino, {
                    text: `📊 *Planilha "${doc.fileName}" salva com sucesso!*\n• *Abas:* ${wb.SheetNames.join(', ')}\n• *Linhas na 1ª aba:* ${linhas.length}\n• *Colunas:* ${colunas}\n\nPara consultar, envie mensagens começando com *!* (Ex: *!analise a planilha*).`
                });
                continue;
            }

            // Só aciona a IA se você começar a frase com "!"
            if (!texto.startsWith('!')) continue;

            const perguntaLimpa = texto.slice(1).trim();
            if (!perguntaLimpa) continue;

            const resposta = await processarComIA(perguntaLimpa);
            await sock.sendMessage(destino, { text: `🤖 *Assistente IA:*\n${resposta}` });
        }
    });
}

iniciarWhatsApp();
