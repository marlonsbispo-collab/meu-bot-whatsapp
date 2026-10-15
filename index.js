const { default: makeWASocket, useMultiFileAuthState, downloadMediaMessage, DisconnectReason, fetchLatestBaileysVersion } = require('@whiskeysockets/baileys');
const express = require('express');
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
const MEU_NUMERO = process.env.MEU_NUMERO || ''; // Opcional: 5511999999999
const RENDER_EXTERNAL_URL = process.env.RENDER_EXTERNAL_URL || '';

const EXCEL_PATH = './base_dados.xlsx';
const PDF_PATH = './documento_atual.pdf';
const IMAGE_PATH = './imagem_atual.jpg';

const LISTA_MODELOS = [
    'gemini-3.8-flash',
    'gemini-3.8-flash-lite-preview',
    'gemini-3.8-pro-preview',
    'gemini-2.0-flash',
    'gemini-2.0-flash-lite'
];

let statusConexao = 'Iniciando...';
let codigoPareamento = null;
let sockGlobal = null;

// Rota rápida para manter o servidor do Render acordado 24/7
app.get('/ping', (req, res) => res.status(200).send('OK'));

// Rota para gerar o código de 8 dígitos direto pelo navegador
app.get('/gerar-codigo', async (req, res) => {
    const num = (req.query.numero || MEU_NUMERO).replace(/\D/g, '');
    if (!num || num.length < 10) {
        return res.send('<h3>⚠️ Digite o número completo com DDI 55 e DDD (Ex: 5511999999999). <a href="/">Voltar</a></h3>');
    }
    try {
        if (!sockGlobal) return res.send('<h3>O WhatsApp ainda está iniciando. Aguarde 5 segundos e tente novamente. <a href="/">Voltar</a></h3>');
        codigoPareamento = await sockGlobal.requestPairingCode(num);
        res.redirect('/');
    } catch (err) {
        res.send(`<h3>Erro ao gerar código: ${err.message}. <a href="/">Voltar</a></h3>`);
    }
});

// Painel Principal (Sem QR Code, usando Código de 8 Dígitos)
app.get('/', (req, res) => {
    if (statusConexao === 'Conectado!') {
        return res.send(`
            <div style="font-family: sans-serif; text-align: center; margin-top: 50px;">
                <h1 style="color: #25D366;">🔒 WHATSAPP CONECTADO E ESTÁVEL!</h1>
                <p>Modo Privado Ativo (Voz + Imagem + PDF + Excel + Agenda + Web).</p>
            </div>
        `);
    }

    res.send(`
        <div style="font-family: sans-serif; max-width: 500px; margin: 40px auto; text-align: center; padding: 25px; border: 1px solid #ddd; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.08);">
            <h2>📱 Conectar sem QR Code (Código de 8 Dígitos)</h2>
            <p>Status atual: <strong>${statusConexao}</strong></p>
            
            ${codigoPareamento ? `
                <div style="background: #e8f5e9; padding: 20px; border-radius: 10px; margin: 20px 0; border: 2px dashed #25D366;">
                    <p style="margin: 0 0 10px 0; font-size: 14px; color: #333;">Digite este código no seu WhatsApp:</p>
                    <h1 style="letter-spacing: 6px; color: #1b5e20; margin: 0; font-size: 36px;">${codigoPareamento}</h1>
                </div>
                <p style="font-size: 13px; color: #555; text-align: left; line-height: 1.5;">
                    1. No WhatsApp do celular, vá em <b>Aparelhos Conectados > Conectar um aparelho</b>.<br>
                    2. Na parte inferior da tela da câmera, toque em <b>"Conectar com número de telefone"</b>.<br>
                    3. Digite os 8 dígitos acima!
                </p>
            ` : `
                <form action="/gerar-codigo" method="GET" style="margin-top: 20px;">
                    <p style="font-size: 14px; color: #555;">Digite seu número com <b>55 + DDD + Número</b> (apenas números):</p>
                    <input type="text" name="numero" placeholder="Ex: 5511999999999" value="${MEU_NUMERO}" required style="padding: 12px; width: 80%; font-size: 16px; border: 1px solid #ccc; border-radius: 6px; margin-bottom: 12px; text-align: center;" />
                    <br>
                    <button type="submit" style="padding: 12px 24px; font-size: 16px; background: #25D366; color: white; border: none; border-radius: 6px; cursor: pointer; font-weight: bold;">Gerar Código de 8 Dígitos</button>
                </form>
            `}
        </div>
    `);
});

app.listen(PORT, '0.0.0.0', () => {
    console.log(`MODO PARING CODE + ANTI-QUEDA ATIVO na porta ${PORT}`);
    // Auto-Ping interno a cada 10 minutos para evitar hibernação
    if (RENDER_EXTERNAL_URL) {
        setInterval(() => {
            fetch(`${RENDER_EXTERNAL_URL}/ping`).catch(() => {});
        }, 10 * 60 * 1000);
    }
});

async function chamarGeminiComRetry(parametros) {
    let ultimoErro = null;

    for (let tentativa = 1; tentativa <= 3; tentativa++) {
        try {
            return await ai.models.generateContent({
                ...parametros,
                model: 'gemini-3.8-flash'
            });
        } catch (err) {
            ultimoErro = err;
            const msg = err.message || "";
            if (msg.includes('503') || msg.includes('UNAVAILABLE') || msg.includes('high demand') || msg.includes('429')) {
                await new Promise(r => setTimeout(r, tentativa * 2500));
                continue;
            }
            break;
        }
    }

    for (const modeloReserva of LISTA_MODELOS.slice(1)) {
        try {
            return await ai.models.generateContent({
                ...parametros,
                model: modeloReserva
            });
        } catch (err) {
            ultimoErro = err;
            await new Promise(r => setTimeout(r, 1500));
        }
    }

    throw ultimoErro;
}

function obterAgenda() {
    if (!GOOGLE_CREDENTIALS_JSON) return null;
    try {
        const credentials = JSON.parse(GOOGLE_CREDENTIALS_JSON);
        const auth = new google.auth.GoogleAuth({
            credentials,
            scopes: ['https://www.googleapis.com/auth/calendar']
        });
        return google.calendar({ version: 'v3', auth });
    } catch (err) {
        return null;
    }
}

async function listarEventosAgenda(dias = 7) {
    try {
        const calendar = obterAgenda();
        if (!calendar) return "⚠️ A Agenda do Google ainda não está configurada nas variáveis de ambiente (GOOGLE_CREDENTIALS_JSON e CALENDAR_ID) lá no painel do Render.";
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
        if (!eventos || eventos.length === 0) return `✅ Agenda conectada! Nenhum compromisso encontrado para os próximos ${dias} dias.`;
        return `📅 *Compromissos dos próximos ${dias} dias:*\n` + eventos.map(e => `- ${e.start.dateTime || e.start.date}: ${e.summary}`).join('\n');
    } catch (err) {
        return `Erro ao consultar agenda: ${err.message}`;
    }
}

async function criarEventoAgenda(titulo, dataHoraInicio, duracaoMinutos = 60, descricao = "") {
    try {
        const calendar = obterAgenda();
        if (!calendar) return "⚠️ Agenda ainda não configurada no menu Environment do Render.";
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
        return "Nenhuma planilha foi enviada ainda. Envie um arquivo .xlsx aqui no chat do WhatsApp.";
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

async function consultarDocumentoPDF(pergunta = "Faça um resumo executivo completo deste documento PDF.") {
    if (!fs.existsSync(PDF_PATH)) {
        return "Nenhum arquivo PDF foi enviado ainda. Envie um documento .pdf aqui no chat do WhatsApp.";
    }
    try {
        const base64Pdf = fs.readFileSync(PDF_PATH).toString('base64');
        const res = await chamarGeminiComRetry({
            contents: [
                {
                    role: 'user',
                    parts: [
                        { inlineData: { mimeType: 'application/pdf', data: base64Pdf } },
                        { text: `Analise este documento PDF com máxima atenção e responda ao seguinte pedido de forma clara, estruturada em tópicos e objetiva para WhatsApp: ${pergunta}` }
                    ]
                }
            ]
        });
        return res.text || "Não foi possível extrair informações do PDF.";
    } catch (err) {
        return `Erro ao analisar o PDF: ${err.message}`;
    }
}

async function consultarImagem(pergunta = "Analise esta imagem em detalhes, extraia todos os textos, números ou valores visíveis e faça um resumo objetivo.") {
    if (!fs.existsSync(IMAGE_PATH)) {
        return "Nenhuma imagem ou foto foi enviada ainda. Envie uma foto ou print aqui no chat do WhatsApp.";
    }
    try {
        const base64Img = fs.readFileSync(IMAGE_PATH).toString('base64');
        const res = await chamarGeminiComRetry({
            contents: [
                {
                    role: 'user',
                    parts: [
                        { inlineData: { mimeType: 'image/jpeg', data: base64Img } },
                        { text: `Analise esta imagem com máxima atenção e responda ao seguinte pedido de forma clara, estruturada e objetiva para WhatsApp: ${pergunta}` }
                    ]
                }
            ]
        });
        return res.text || "Não foi possível analisar a imagem.";
    } catch (err) {
        return `Erro ao analisar a imagem: ${err.message}`;
    }
}

async function pesquisarNaInternet(consulta) {
    try {
        let resultadosTexto = [];

        const urlNews = `https://news.google.com/rss/search?q=${encodeURIComponent(consulta)}&hl=pt-BR&gl=BR&ceid=BR:pt-419`;
        const resNews = await fetch(urlNews);
        const xml = await resNews.text();
        const items = [...xml.matchAll(/<item>[\s\S]*?<title>(.*?)<\/title>[\s\S]*?<pubDate>(.*?)<\/pubDate>[\s\S]*?<\/item>/g)].slice(0, 8);
        if (items.length > 0) {
            resultadosTexto.push("NOTÍCIAS E DADOS RECENTES DO GOOGLE:\n" + items.map(m => `- ${m[1]} (${m[2]})`).join('\n'));
        }

        const urlWeb = `https://html.duckduckgo.com/html/?q=${encodeURIComponent(consulta)}`;
        const resWeb = await fetch(urlWeb, { headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' } });
        const html = await resWeb.text();
        const snippets = [...html.matchAll(/class="result__snippet[^>]*>([\s\S]*?)<\/a>/g)]
            .slice(0, 6)
            .map(m => m[1].replace(/<[^>]+>/g, '').trim());
        if (snippets.length > 0) {
            resultadosTexto.push("RESULTADOS DA WEB:\n" + snippets.map(s => `- ${s}`).join('\n'));
        }

        return resultadosTexto.join('\n\n') || "Nenhum resultado encontrado na web.";
    } catch (err) {
        return `Erro ao procurar na web: ${err.message}`;
    }
}

const ferramentas = [{
    functionDeclarations: [
        {
            name: "pesquisarNaInternet",
            description: "Pesquisa na internet em tempo real sobre jogos, clima, previsão do tempo, notícias, cotações ou qualquer informação atualizada.",
            parameters: {
                type: Type.OBJECT,
                properties: {
                    consulta: { type: Type.STRING, description: "Termos de busca objetivos (ex: 'proximo jogo Corinthians data horario', 'clima Sao Paulo agora')" }
                },
                required: ["consulta"]
            }
        },
        {
            name: "consultarPlanilhaExcel",
            description: "Lê e pesquisa dados na planilha Excel enviada pelo usuário no WhatsApp.",
            parameters: {
                type: Type.OBJECT,
                properties: {
                    termoBusca: { type: Type.STRING, description: "Palavra-chave, loja, nome ou status para filtrar (ou vazio para ver resumo geral)." }
                }
            }
        },
        {
            name: "consultarDocumentoPDF",
            description: "Lê, resume ou responde perguntas sobre o último arquivo PDF enviado pelo usuário no WhatsApp.",
            parameters: {
                type: Type.OBJECT,
                properties: {
                    pergunta: { type: Type.STRING, description: "O que procurar, resumir ou responder com base no documento PDF." }
                },
                required: ["pergunta"]
            }
        },
        {
            name: "consultarImagem",
            description: "Lê, extrai textos/valores ou responde perguntas sobre a última foto, nota fiscal, print ou imagem enviada pelo usuário no WhatsApp.",
            parameters: {
                type: Type.OBJECT,
                properties: {
                    pergunta: { type: Type.STRING, description: "O que analisar, ler ou extrair da imagem." }
                },
                required: ["pergunta"]
            }
        },
        {
            name: "listarEventosAgenda",
            description: "Lista os compromissos do Google Calendar dos próximos dias.",
            parameters: {
                type: Type.OBJECT,
                properties: {
                    dias: { type: Type.NUMBER, description: "Quantidade de dias para buscar (padrão 7)." }
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

const INSTRUCAO_SISTEMA = () => {
    const agora = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
    return `Você é um Assistente Executivo Pessoal de elite no WhatsApp. Data/hora atual: ${agora}.

SUAS CAPACIDADES PRINCIPAIS:
1. LEITURA E TRANSCRIÇÃO DE ÁUDIOS:
   - Quando receber um áudio, sempre inicie a resposta mostrando:
     📝 *Transcrição do Áudio:* "[texto exato falado no áudio]"
   - Se o áudio contiver um pedido (ex: agendar algo, consultar planilha, consultar o PDF, analisar a última foto/imagem, buscar na web, reescrever um e-mail), execute a ação logo abaixo da transcrição!
   - Se for apenas um áudio encaminhado ou informativo, entregue a transcrição completa e um resumo em tópicos.

2. REESCRITA DE TEXTOS E E-MAILS:
   - Sempre que o usuário pedir para reescrever, melhorar ou redigir um texto, comunicado ou e-mail num tom específico (ex: executivo, formal, persuasivo, diplomático, direto), adapte o texto perfeitamente ao tom solicitado, pronto para copiar e colar.

3. FERRAMENTAS (WEB, EXCEL, PDF, IMAGEM E AGENDA):
   - Para clima, jogos, notícias ou fatos atuais: use "pesquisarNaInternet".
   - Para dados da planilha Excel enviada: use "consultarPlanilhaExcel".
   - Para perguntas sobre o documento PDF enviado: use "consultarDocumentoPDF".
   - Para perguntas sobre a última foto/print/imagem enviada: use "consultarImagem".
   - Para compromissos na agenda: use "listarEventosAgenda" ou "criarEventoAgenda".

Responda de forma organizada e formatada para leitura rápida no WhatsApp.`;
};

async function processarComIA(partesConteudo, textoReferencia = "") {
    try {
        const config = {
            systemInstruction: INSTRUCAO_SISTEMA(),
            tools: ferramentas,
            temperature: 0.35
        };

        let response = await chamarGeminiComRetry({
            contents: [{ role: 'user', parts: partesConteudo }],
            config
        });

        if (response.functionCalls && response.functionCalls.length > 0) {
            const call = response.functionCalls[0];
            let resultadoFerramenta = "";

            if (call.name === "pesquisarNaInternet") {
                resultadoFerramenta = await pesquisarNaInternet(call.args.consulta || textoReferencia);
            } else if (call.name === "consultarPlanilhaExcel") {
                resultadoFerramenta = consultarPlanilhaExcel(call.args.termoBusca || "");
            } else if (call.name === "consultarDocumentoPDF") {
                resultadoFerramenta = await consultarDocumentoPDF(call.args.pergunta || textoReferencia);
            } else if (call.name === "consultarImagem") {
                resultadoFerramenta = await consultarImagem(call.args.pergunta || textoReferencia);
            } else if (call.name === "listarEventosAgenda") {
                resultadoFerramenta = await listarEventosAgenda(call.args.dias || 7);
            } else if (call.name === "criarEventoAgenda") {
                resultadoFerramenta = await criarEventoAgenda(call.args.titulo, call.args.dataHoraInicio, call.args.duracaoMinutos, call.args.descricao);
            }

            await new Promise(r => setTimeout(r, 1500));

            response = await chamarGeminiComRetry({
                contents: [
                    { role: 'user', parts: partesConteudo },
                    response.candidates[0].content,
                    { role: 'user', parts: [{ functionResponse: { name: call.name, response: { result: resultadoFerramenta } } }] }
                ],
                config
            });
        }

        return response.text || "Processado com sucesso.";
    } catch (err) {
        return `⚠️ Detalhe técnico da API: ${err.message?.substring(0, 120)}`;
    }
}

async function iniciarWhatsApp() {
    const { state, saveCreds } = await useMultiFileAuthState('./auth_info');
    const { version } = await fetchLatestBaileysVersion();

    const sock = makeWASocket({
        version,
        auth: state,
        logger: pino({ level: 'silent' }),
        printQRInTerminal: false,
        browser: ['Ubuntu', 'Chrome', '20.0.04'], // Obrigatório para o Pairing Code funcionar com estabilidade
        keepAliveIntervalMs: 30000,               // Mantém o túnel do WhatsApp vivo a cada 30s
        syncFullHistory: false,                   // Economiza memória RAM do Render gratuito
        markOnlineOnConnect: false
    });

    sockGlobal = sock;
    sock.ev.on('creds.update', saveCreds);

    sock.ev.on('connection.update', async (update) => {
        const { connection, lastDisconnect, qr } = update;

        if (qr && !sock.authState.creds.registered) {
            statusConexao = 'Pronto para gerar código de 8 dígitos';
            // Se você definiu MEU_NUMERO no Environment do Render, ele já gera o código sozinho!
            if (MEU_NUMERO && !codigoPareamento) {
                try {
                    const numLimpo = MEU_NUMERO.replace(/\D/g, '');
                    codigoPareamento = await sock.requestPairingCode(numLimpo);
                    console.log(`========================================`);
                    console.log(`SEU CÓDIGO DE PAREAMENTO: ${codigoPareamento}`);
                    console.log(`========================================`);
                } catch (e) {
                    console.error("Erro ao gerar código automático:", e.message);
                }
            }
        }

        if (connection === 'open') {
            codigoPareamento = null;
            statusConexao = 'Conectado!';
            console.log('WhatsApp conectado e blindado!');
        }

        if (connection === 'close') {
            const statusCode = lastDisconnect?.error?.output?.statusCode;
            const deveReconectar = statusCode !== DisconnectReason.loggedOut;
            statusConexao = 'Reconectando...';
            if (deveReconectar) {
                setTimeout(() => iniciarWhatsApp(), 4000);
            } else {
                statusConexao = 'Desconectado (Sessão encerrada no celular).';
            }
        }
    });

    sock.ev.on('messages.upsert', async ({ messages }) => {
        for (const msg of messages) {
            if (!msg.message) continue;
            if (!msg.key.fromMe) continue;

            const remetente = msg.key.remoteJid || "";
            if (remetente.includes('@g.us') || remetente === 'status@broadcast') continue;

            const meuNumeroLimpo = sock.user?.id ? sock.user.id.split(':')[0].split('@')[0] : "";
            const meuLidLimpo = sock.user?.lid ? sock.user.lid.split(':')[0].split('@')[0] : "";
            const ehConversaComigoMesmo =
                (meuNumeroLimpo && remetente.startsWith(meuNumeroLimpo)) ||
                (meuLidLimpo && remetente.startsWith(meuLidLimpo)) ||
                remetente.endsWith('@lid');

            if (!ehConversaComigoMesmo) continue;

            const destino = meuNumeroLimpo ? `${meuNumeroLimpo}@s.whatsapp.net` : remetente;
            const conteudoMsg = msg.message.ephemeralMessage?.message || msg.message;

            const texto = conteudoMsg.conversation ||
                          conteudoMsg.extendedTextMessage?.text || "";

            if (texto.startsWith('🤖 *Assistente IA') || texto.startsWith('📊 *Planilha') || texto.startsWith('📄 *PDF') || texto.startsWith('📅 *Compromissos')) continue;

            const doc = conteudoMsg.documentMessage || conteudoMsg.documentWithCaptionMessage?.message?.documentMessage;

            // 1. PLANILHA EXCEL
            if (doc && (doc.fileName?.toLowerCase().endsWith('.xlsx') || doc.fileName?.toLowerCase().endsWith('.xls'))) {
                const buffer = await downloadMediaMessage(msg, 'buffer', {});
                fs.writeFileSync(EXCEL_PATH, buffer);
                const wb = xlsx.readFile(EXCEL_PATH);
                const primeiraAba = wb.Sheets[wb.SheetNames[0]];
                const linhas = xlsx.utils.sheet_to_json(primeiraAba);
                const colunas = linhas.length > 0 ? Object.keys(linhas[0]).join(', ') : 'Nenhuma';

                await sock.sendMessage(destino, {
                    text: `📊 *Planilha "${doc.fileName}" salva com sucesso!*\n• *Abas:* ${wb.SheetNames.join(', ')}\n• *Linhas na 1ª aba:* ${linhas.length}\n• *Colunas:* ${colunas}\n\nVocê pode consultá-la por texto com *!* ou enviando um *áudio*!`
                });
                continue;
            }

            // 2. ARQUIVO PDF
            if (doc && (doc.fileName?.toLowerCase().endsWith('.pdf') || doc.mimetype === 'application/pdf')) {
                try {
                    const bufferPdf = await downloadMediaMessage(msg, 'buffer', {});
                    fs.writeFileSync(PDF_PATH, bufferPdf);

                    const legendaPdf = doc.caption?.trim() || "Faça um resumo executivo completo deste documento PDF, destacando os pontos principais, datas, valores ou conclusões importantes.";

                    await sock.sendMessage(destino, {
                        text: `📄 *PDF "${doc.fileName || 'documento.pdf'}" recebido!*\n_Analisando o conteúdo e gerando o resumo..._`
                    });

                    const analisePdf = await consultarDocumentoPDF(legendaPdf);
                    await sock.sendMessage(destino, {
                        text: `🤖 *Assistente IA (Análise do PDF):*\n${analisePdf}\n\n💡 _O PDF continua salvo! Se quiser fazer perguntas específicas sobre ele, basta usar *!* ou mandar um áudio._`
                    });
                } catch (err) {
                    await sock.sendMessage(destino, { text: `⚠️ Erro ao processar o PDF: ${err.message}` });
                }
                continue;
            }

            // 3. IMAGEM / FOTO
            const imgMsg = conteudoMsg.imageMessage;
            if (imgMsg) {
                try {
                    const bufferImg = await downloadMediaMessage(msg, 'buffer', {});
                    fs.writeFileSync(IMAGE_PATH, bufferImg);

                    const legendaImg = imgMsg.caption?.trim();
                    const instrucaoImg = legendaImg
                        ? (legendaImg.startsWith('!') ? legendaImg.slice(1).trim() : legendaImg)
                        : "Analise esta imagem com atenção: descreva o que vê, extraia todos os textos, tabelas, números ou valores importantes e apresente um resumo organizado.";

                    const respostaImg = await consultarImagem(instrucaoImg);
                    await sock.sendMessage(destino, {
                        text: `🤖 *Assistente IA (Análise de Imagem):*\n${respostaImg}\n\n💡 _A foto ficou salva! Você pode fazer mais perguntas sobre ela usando *!* ou enviando um áudio._`
                    });
                } catch (err) {
                    await sock.sendMessage(destino, { text: `⚠️ Erro ao processar a imagem: ${err.message}` });
                }
                continue;
            }

            // 4. ÁUDIO / VOZ
            const audioMsg = conteudoMsg.audioMessage;
            if (audioMsg) {
                try {
                    const bufferAudio = await downloadMediaMessage(msg, 'buffer', {});
                    const base64Audio = bufferAudio.toString('base64');
                    const mimeType = (audioMsg.mimetype || 'audio/ogg').split(';')[0];

                    const partesAudio = [
                        { inlineData: { mimeType: mimeType, data: base64Audio } },
                        { text: "Transcreva este áudio em português. Em seguida, se houver algum pedido, pergunta, consulta de planilha Excel, consulta do documento PDF, análise da imagem enviada, agendamento, busca na internet ou pedido para reescrever texto/e-mail contido no áudio, execute-o completamente!" }
                    ];

                    const respostaAudio = await processarComIA(partesAudio, "Comando por áudio");
                    await sock.sendMessage(destino, { text: `🤖 *Assistente IA:*\n${respostaAudio}` });
                } catch (err) {
                    await sock.sendMessage(destino, { text: `⚠️ Erro ao processar áudio: ${err.message}` });
                }
                continue;
            }

            // 5. TEXTO COM "!"
            if (!texto.startsWith('!')) continue;

            const perguntaLimpa = texto.slice(1).trim();
            if (!perguntaLimpa) continue;

            if (perguntaLimpa.toLowerCase() === 'agenda' || perguntaLimpa.toLowerCase() === 'conectar na minha agenda') {
                const statusAgenda = await listarEventosAgenda(7);
                await sock.sendMessage(destino, { text: `🤖 *Assistente IA:*\n${statusAgenda}` });
                continue;
            }

            const resposta = await processarComIA([{ text: perguntaLimpa }], perguntaLimpa);
            await sock.sendMessage(destino, { text: `🤖 *Assistente IA:*\n${resposta}` });
        }
    });
}

iniciarWhatsApp();
