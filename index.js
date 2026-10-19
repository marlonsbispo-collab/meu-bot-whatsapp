const { default: makeWASocket, useMultiFileAuthState, downloadMediaMessage, DisconnectReason, fetchLatestBaileysVersion } = require('@whiskeysockets/baileys');
const express = require('express');
const pino = require('pino');
const fs = require('fs');
const xlsx = require('xlsx');
const { GoogleGenAI, Type } = require('@google/genai');

const app = express();
const PORT = process.env.PORT || 10000;

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
const MEU_NUMERO = process.env.MEU_NUMERO || '';
const RENDER_EXTERNAL_URL = process.env.RENDER_EXTERNAL_URL || '';

const EXCEL_PATH = './base_dados.xlsx';
const PDF_PATH = './documento_atual.pdf';
const IMAGE_PATH = './imagem_atual.jpg';
const AGENDA_LOCAL_PATH = './agenda_local.json';

// Lista oficial com os 8 modelos reais de texto da sua chave Google
const LISTA_MODELOS_REAIS = [
    'gemini-3.8-flash',
    'gemini-3.7-flash',
    'gemini-3.6-flash',
    'gemini-3.5-flash',
    'gemini-3.5-flash-lite',
    'gemini-3.1-flash-lite',
    'gemini-omni-1.1-flash',
    'gemma-4-31b-it'
];

// Lista oficial dos modelos de Voz (TTS) ativos na sua chave Google
const LISTA_MODELOS_TTS = [
    'gemini-3.8-flash-tts',
    'gemini-3.8-flash-lite-tts',
    'gemini-2.5-flash-preview-tts',
    'gemini-3.1-flash-tts-preview'
];

let statusConexao = 'Iniciando...';
let codigoPareamento = null;
let sockGlobal = null;
let meuDestinoGlobal = null;

// ==========================================
// MELHORIA 1: MEMÓRIA DE CONVERSA CONTÍNUA
// ==========================================
let historicoConversa = [];
const MAX_HISTORICO = 10; // Guarda as últimas 10 mensagens trocadas

function adicionarAoHistorico(role, texto) {
    if (!texto) return;
    historicoConversa.push({
        role,
        texto: texto.substring(0, 2500),
        hora: new Date().toLocaleTimeString('pt-BR', { timeZone: 'America/Sao_Paulo' })
    });
    if (historicoConversa.length > MAX_HISTORICO) {
        historicoConversa.shift();
    }
}

function obterResumoHistorico() {
    if (historicoConversa.length === 0) return "Nenhuma conversa anterior nesta sessão.";
    return historicoConversa.map(h => `[${h.hora}] ${h.role === 'user' ? 'Usuário' : 'Assistente'}: ${h.texto}`).join('\n');
}

app.get('/ping', (req, res) => res.status(200).send('OK'));

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

app.get('/', (req, res) => {
    if (statusConexao === 'Conectado!') {
        return res.send(`
            <div style="font-family: sans-serif; text-align: center; margin-top: 50px;">
                <h1 style="color: #25D366;">🎙️ ASSISTENTE IA PRO (MEMÓRIA + VOZ TTS + AGENDA COMPLETA) ATIVA!</h1>
                <p>Memória de Contexto + Resposta por Voz + Gestão de Agenda + PDF + Excel + Imagem + Web.</p>
            </div>
        `);
    }

    res.send(`
        <div style="font-family: sans-serif; max-width: 500px; margin: 40px auto; text-align: center; padding: 25px; border: 1px solid #ddd; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.08);">
            <h2>📱 Conectar sem QR Code (Código de 8 Dígitos)</h2>
            <p>Status atual: <strong>${statusConexao}</strong></p>${codigoPareamento ? `
                <div style="background: #e8f5e9; padding: 20px; border-radius: 10px; margin: 20px 0; border: 2px dashed #25D366;">
                    <p style="margin: 0 0 10px 0; font-size: 14px; color: #333;">Digite este código no seu WhatsApp:</p>
                    <h1 style="letter-spacing: 6px; color: #1b5e20; margin: 0; font-size: 36px;">${codigoPareamento}</h1>
                </div>
            ` : `
                <form action="/gerar-codigo" method="GET" style="margin-top: 20px;">
                    <p style="font-size: 14px; color: #555;">Digite o seu número com <b>55 + DDD + Número</b>:</p>
                    <input type="text" name="numero" placeholder="Ex: 5511999999999" value="${MEU_NUMERO}" required style="padding: 12px; width: 80%; font-size: 16px; border: 1px solid #ccc; border-radius: 6px; margin-bottom: 12px; text-align: center;" />
                    <br>
                    <button type="submit" style="padding: 12px 24px; font-size: 16px; background: #25D366; color: white; border: none; border-radius: 6px; cursor: pointer; font-weight: bold;">Gerar Código de 8 Dígitos</button>
                </form>
            `}
        </div>
    `);
});

app.listen(PORT, '0.0.0.0', () => {
    console.log(`ASSISTENTE PRO (MEMÓRIA + VOZ TTS + AGENDA COMPLETA) ATIVA na porta ${PORT}`);
    if (RENDER_EXTERNAL_URL) {
        setInterval(() => {
            fetch(`${RENDER_EXTERNAL_URL}/ping`).catch(() => {});
        }, 10 * 60 * 1000);
    }
});

// Sistema Multi-Modelos em Cascata para Texto
async function chamarGeminiComRetry(parametros) {
    let ultimoErro = null;

    for (const modeloAtual of LISTA_MODELOS_REAIS) {
        try {
            return await ai.models.generateContent({
                ...parametros,
                model: modeloAtual
            });
        } catch (err) {
            ultimoErro = err;
            console.log(`[Troca Automática] Modelo ${modeloAtual} ocupado. Pulando para o próximo...`);
            await new Promise(r => setTimeout(r, 600));
        }
    }
    throw ultimoErro;
}

// ==========================================
// MELHORIA 4: RESPOSTA EM VOZ (GEMINI TTS)
// ==========================================
// Converte o áudio PCM 24kHz bruto do Gemini para um arquivo WAV reproduzível no WhatsApp
function converterPcmParaWav(pcmBuffer, sampleRate = 24000) {
    const numChannels = 1;
    const bitsPerSample = 16;
    const byteRate = (sampleRate * numChannels * bitsPerSample) / 8;
    const blockAlign = (numChannels * bitsPerSample) / 8;
    const dataSize = pcmBuffer.length;
    const buffer = Buffer.alloc(44 + dataSize);

    buffer.write('RIFF', 0);
    buffer.writeUInt32LE(36 + dataSize, 4);
    buffer.write('WAVE', 8);
    buffer.write('fmt ', 12);
    buffer.writeUInt32LE(16, 16);
    buffer.writeUInt16LE(1, 20); // PCM
    buffer.writeUInt16LE(numChannels, 22);
    buffer.writeUInt32LE(sampleRate, 24);
    buffer.writeUInt32LE(byteRate, 28);
    buffer.writeUInt16LE(blockAlign, 32);
    buffer.writeUInt16LE(bitsPerSample, 34);
    buffer.write('data', 36);
    buffer.writeUInt32LE(dataSize, 40);

    pcmBuffer.copy(buffer, 44);
    return buffer;
}

async function gerarAudioRespostaTTS(textoResposta) {
    try {
        // Limpa emojis, links, transcrições e formatações Markdown para a voz soar natural
        let textoLimpo = textoResposta
            .replace(/📝 \*Transcrição do Áudio:\* "[\s\S]*?"\n*/gi, '')
            .replace(/https?:\/\/\S+/g, '')
            .replace(/[*_~`#>-]/g, '')
            .replace(/\s+/g, ' ')
            .trim();

        if (!textoLimpo) return null;
        if (textoLimpo.length > 650) {
            textoLimpo = textoLimpo.substring(0, 650) + "... Veja os detalhes completos na mensagem de texto acima.";
        }

        for (const modeloTts of LISTA_MODELOS_TTS) {
            try {
                const res = await ai.models.generateContent({
                    model: modeloTts,
                    contents: [{ role: 'user', parts: [{ text: `Leia em voz alta em português do Brasil de forma natural, clara e executiva: ${textoLimpo}` }] }],
                    config: {
                        responseModalities: ["AUDIO"],
                        speechConfig: {
                            voiceConfig: {
                                prebuiltVoiceConfig: {
                                    voiceName: "Aoede" // Voz feminina executiva e natural
                                }
                            }
                        }
                    }
                });

                const partAudio = res.candidates?.[0]?.content?.parts?.find(p => p.inlineData?.data);
                if (partAudio) {
                    const rawBuffer = Buffer.from(partAudio.inlineData.data, 'base64');
                    const mime = partAudio.inlineData.mimeType || '';
                    if (mime.includes('pcm') || mime.includes('L16') || !mime.includes('wav')) {
                        return converterPcmParaWav(rawBuffer, 24000);
                    }
                    return rawBuffer;
                }
            } catch (e) {
                console.log(`[TTS] Modelo ${modeloTts} indisponível (${e.message?.substring(0, 40)}). Tentando próximo...`);
            }
        }
        return null;
    } catch (err) {
        console.error("Erro geral no TTS:", err.message);
        return null;
    }
}

// ==========================================
// MELHORIA 5: GESTÃO COMPLETA DA AGENDA
// (Listar, Criar, Cancelar e Reagendar)
// ==========================================
function lerAgendaLocal() {
    if (!fs.existsSync(AGENDA_LOCAL_PATH)) return [];
    try {
        return JSON.parse(fs.readFileSync(AGENDA_LOCAL_PATH, 'utf8'));
    } catch (e) {
        return [];
    }
}

function salvarAgendaLocal(eventos) {
    fs.writeFileSync(AGENDA_LOCAL_PATH, JSON.stringify(eventos, null, 2));
}

function formatarDataBR(dateObj) {
    return dateObj.toLocaleString('pt-BR', {
        timeZone: 'America/Sao_Paulo',
        day: '2-digit',
        month: '2-digit',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
    });
}

function converterParaDateSP(dataHoraIso) {
    const isoComFuso = dataHoraIso.includes('-03:00') || dataHoraIso.endsWith('Z')
        ? dataHoraIso
        : `${dataHoraIso.slice(0, 19)}-03:00`;
    return new Date(isoComFuso);
}

async function listarEventosAgenda() {
    const eventos = lerAgendaLocal();
    if (eventos.length === 0) {
        return "📅 *Sua Agenda está vazia no momento.*\nPara agendar algo, basta pedir por áudio ou texto (Ex: *!agende reunião amanhã às 14h*).";
    }

    eventos.sort((a, b) => new Date(a.isoInicio) - new Date(b.isoInicio));
    salvarAgendaLocal(eventos);

    const listaTexto = eventos.map((e, i) => {
        const desc = e.descricao ? `\n   _${e.descricao}_` : '';
        return `*${i + 1}.* 🕒 *${e.dataFormatada}* — *${e.titulo}*${desc}`;
    }).join('\n\n');

    return `📅 *Seus Compromissos Agendados:*\n\n${listaTexto}\n\n💡 _Você pode pedir para cancelar ou reagendar qualquer item pelo número ou pelo nome (Ex: "!cancele o item 1" ou "!mude a reunião com TI para as 15h")._`;
}

async function criarEventoAgenda(titulo, dataHoraInicio, duracaoMinutos = 60, descricao = "") {
    try {
        const inicio = converterParaDateSP(dataHoraInicio);
        const fim = new Date(inicio.getTime() + (duracaoMinutos || 60) * 60000);
        const dataFormatada = formatarDataBR(inicio);

        const eventos = lerAgendaLocal();
        eventos.push({
            id: Date.now(),
            titulo,
            isoInicio: inicio.toISOString(),
            dataFormatada,
            duracaoMinutos: duracaoMinutos || 60,
            descricao: descricao || "",
            notificado15min: false
        });
        eventos.sort((a, b) => new Date(a.isoInicio) - new Date(b.isoInicio));
        salvarAgendaLocal(eventos);

        const fmt = (d) => d.toISOString().replace(/[-:]/g, '').split('.')[0] + 'Z';
        const linkGoogleCalendar = `https://calendar.google.com/calendar/render?action=TEMPLATE&text=${encodeURIComponent(titulo)}&dates=${fmt(inicio)}/${fmt(fim)}&details=${encodeURIComponent(descricao || 'Agendado via Assistente IA WhatsApp')}`;

        return `✅ *Compromisso Agendado com Sucesso!*\n\n📌 *Assunto:* ${titulo}\n🗓️ *Data e Hora:* ${dataFormatada}\n⏱️ *Duração:* ${duracaoMinutos || 60} min\n🔔 *Lembrete:* Vou te mandar um aviso aqui no WhatsApp 15 minutos antes!\n\n📲 *Salvar também no calendário do celular:*\n${linkGoogleCalendar}`;
    } catch (err) {
        return `Erro ao agendar compromisso: ${err.message}`;
    }
}

async function cancelarEventoAgenda(identificador) {
    const eventos = lerAgendaLocal();
    if (eventos.length === 0) return "📅 Sua agenda já está vazia, não há compromissos para cancelar.";

    eventos.sort((a, b) => new Date(a.isoInicio) - new Date(b.isoInicio));
    const termo = String(identificador).trim().toLowerCase();
    const numIndex = parseInt(termo.replace(/\D/g, ''), 10);

    let indiceRemover = -1;

    // 1. Tenta pelo número na lista (ex: "1", "item 2")
    if (!isNaN(numIndex) && numIndex >= 1 && numIndex <= eventos.length && (termo.length <= 3 || termo.includes('item') || termo.includes('compromisso') || termo.includes('numero') || termo.includes('número'))) {
        indiceRemover = numIndex - 1;
    } else {
        // 2. Tenta buscar por palavra-chave no título ou data
        indiceRemover = eventos.findIndex(e =>
            e.titulo.toLowerCase().includes(termo) ||
            e.dataFormatada.toLowerCase().includes(termo)
        );
    }

    if (indiceRemover === -1) {
        return `⚠️ Não encontrei nenhum compromisso correspondente a *"${identificador}"*.\n\nEnvie *!agenda* para ver a lista numerada dos seus compromissos atuais.`;
    }

    const removido = eventos.splice(indiceRemover, 1)[0];
    salvarAgendaLocal(eventos);

    return `🗑️ *Compromisso Cancelado com Sucesso!*\n\n• *Removido:* ${removido.titulo}\n• *Horário que estava marcado:* ${removido.dataFormatada}\n\nRestam *${eventos.length}* compromisso(s) na sua agenda.`;
}

async function reagendarEventoAgenda(identificador, novaDataHoraInicio, novoTitulo = "") {
    const eventos = lerAgendaLocal();
    if (eventos.length === 0) return "📅 Sua agenda está vazia, não há compromissos para reagendar.";

    eventos.sort((a, b) => new Date(a.isoInicio) - new Date(b.isoInicio));
    const termo = String(identificador).trim().toLowerCase();
    const numIndex = parseInt(termo.replace(/\D/g, ''), 10);

    let indiceEditar = -1;

    if (!isNaN(numIndex) && numIndex >= 1 && numIndex <= eventos.length && (termo.length <= 3 || termo.includes('item') || termo.includes('compromisso') || termo.includes('numero') || termo.includes('número'))) {
        indiceEditar = numIndex - 1;
    } else {
        indiceEditar = eventos.findIndex(e =>
            e.titulo.toLowerCase().includes(termo) ||
            e.dataFormatada.toLowerCase().includes(termo)
        );
    }

    if (indiceEditar === -1) {
        return `⚠️ Não encontrei nenhum compromisso correspondente a *"${identificador}"* para reagendar.\nEnvie *!agenda* para conferir os itens.`;
    }

    const horarioAntigo = eventos[indiceEditar].dataFormatada;
    const novoInicio = converterParaDateSP(novaDataHoraInicio);
    const novaDataFormatada = formatarDataBR(novoInicio);

    eventos[indiceEditar].isoInicio = novoInicio.toISOString();
    eventos[indiceEditar].dataFormatada = novaDataFormatada;
    eventos[indiceEditar].notificado15min = false;
    if (novoTitulo && novoTitulo.trim() !== "") {
        eventos[indiceEditar].titulo = novoTitulo.trim();
    }

    const itemAtualizado = eventos[indiceEditar];
    eventos.sort((a, b) => new Date(a.isoInicio) - new Date(b.isoInicio));
    salvarAgendaLocal(eventos);

    const fim = new Date(novoInicio.getTime() + (itemAtualizado.duracaoMinutos || 60) * 60000);
    const fmt = (d) => d.toISOString().replace(/[-:]/g, '').split('.')[0] + 'Z';
    const linkGoogleCalendar = `https://calendar.google.com/calendar/render?action=TEMPLATE&text=${encodeURIComponent(itemAtualizado.titulo)}&dates=${fmt(novoInicio)}/${fmt(fim)}&details=${encodeURIComponent(itemAtualizado.descricao || 'Reagendado via Assistente IA')}`;

    return `🔄 *Compromisso Reagendado com Sucesso!*\n\n📌 *Assunto:* ${itemAtualizado.titulo}\n❌ *Horário anterior:* ${horarioAntigo}\n✅ *Novo Horário:* ${novaDataFormatada}\n🔔 *Novo alerta programado para 15 min antes!*\n\n📲 *Atualizar no calendário do celular:*\n${linkGoogleCalendar}`;
}

// Despertador automático que verifica a cada 60 segundos se há compromisso próximo
setInterval(async () => {
    if (!sockGlobal || statusConexao !== 'Conectado!' || !meuDestinoGlobal) return;
    try {
        const eventos = lerAgendaLocal();
        if (eventos.length === 0) return;

        const agora = Date.now();
        let alterou = false;

        for (const ev of eventos) {
            const tempoEvento = new Date(ev.isoInicio).getTime();
            const diferencaMinutos = (tempoEvento - agora) / 60000;

            if (diferencaMinutos > 0 && diferencaMinutos <= 16 && !ev.notificado15min) {
                ev.notificado15min = true;
                alterou = true;
                await sockGlobal.sendMessage(meuDestinoGlobal, {
                    text: `🔔 *LEMBRETE DA SUA AGENDA!*\n\nO seu compromisso começa em cerca de *${Math.round(diferencaMinutos)} minutos*:\n📌 *${ev.titulo}*\n🕒 *Horário:* ${ev.dataFormatada}`
                });
            }
        }

        if (alterou) salvarAgendaLocal(eventos);
    } catch (e) {}
}, 60 * 1000);

// ==========================================
// FUNÇÕES DE PLANILHA EXCEL, PDF, IMAGEM E WEB
// ==========================================
async function analisarPlanilhaComIA(perguntaUsuario, termoBusca = "") {
    if (!fs.existsSync(EXCEL_PATH)) {
        return "⚠️ Nenhuma planilha foi enviada ainda. Envie um arquivo .xlsx aqui no chat do WhatsApp.";
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
                filtrados = dados.slice(0, 50);
            }

            const colunas = dados.length > 0 ? Object.keys(dados[0]).join(', ') : 'Vazia';
            resultado.push(`--- ABA: ${nomeAba} (Total: ${dados.length} linhas | Colunas: ${colunas}) ---\n` + JSON.stringify(filtrados, null, 2));
        });

        const dadosTexto = resultado.join('\n\n').substring(0, 14000);
        const res = await chamarGeminiComRetry({
            contents: [{ role: 'user', parts: [{ text: `Histórico recente:\n${obterResumoHistorico()}\n\nDados da planilha Excel:\n${dadosTexto}\n\nResponda ao pedido do usuário: ${perguntaUsuario}` }] }]
        });
        return res.text || dadosTexto;
    } catch (err) {
        return `Erro ao ler Excel: ${err.message}`;
    }
}

async function consultarDocumentoPDF(pergunta = "Faça um resumo executivo completo deste documento PDF.") {
    if (!fs.existsSync(PDF_PATH)) {
        return "⚠️ Nenhum arquivo PDF foi enviado ainda. Envie um documento .pdf aqui no chat do WhatsApp.";
    }
    try {
        const base64Pdf = fs.readFileSync(PDF_PATH).toString('base64');
        const res = await chamarGeminiComRetry({
            contents: [
                {
                    role: 'user',
                    parts: [
                        { inlineData: { mimeType: 'application/pdf', data: base64Pdf } },
                        { text: `Histórico da conversa:\n${obterResumoHistorico()}\n\nAnalise este documento PDF com máxima atenção e responda ao seguinte pedido de forma clara, estruturada em tópicos e objetiva para WhatsApp: ${pergunta}` }
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
        return "⚠️ Nenhuma imagem ou foto foi enviada ainda. Envie uma foto ou print aqui no chat do WhatsApp.";
    }
    try {
        const base64Img = fs.readFileSync(IMAGE_PATH).toString('base64');
        const res = await chamarGeminiComRetry({
            contents: [
                {
                    role: 'user',
                    parts: [
                        { inlineData: { mimeType: 'image/jpeg', data: base64Img } },
                        { text: `Histórico da conversa:\n${obterResumoHistorico()}\n\nAnalise esta imagem com máxima atenção e responda ao seguinte pedido de forma clara, estruturada e objetiva para WhatsApp: ${pergunta}` }
                    ]
                }
            ]
        });
        return res.text || "Não foi possível analisar a imagem.";
    } catch (err) {
        return `Erro ao analisar a imagem: ${err.message}`;
    }
}

async function pesquisarEResponderWeb(perguntaUsuario, consultaBusca) {
    try {
        let resultadosTexto = [];

        const urlNews = `https://news.google.com/rss/search?q=${encodeURIComponent(consultaBusca)}&hl=pt-BR&gl=BR&ceid=BR:pt-419`;
        const resNews = await fetch(urlNews);
        const xml = await resNews.text();
        const items = [...xml.matchAll(/<item>[\s\S]*?<title>(.*?)<\/title>[\s\S]*?<pubDate>(.*?)<\/pubDate>[\s\S]*?<\/item>/g)].slice(0, 8);
        if (items.length > 0) {
            resultadosTexto.push("NOTÍCIAS E DADOS RECENTES DO GOOGLE:\n" + items.map(m => `- ${m[1]} (${m[2]})`).join('\n'));
        }

        const urlWeb = `https://html.duckduckgo.com/html/?q=${encodeURIComponent(consultaBusca)}`;
        const resWeb = await fetch(urlWeb, { headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' } });
        const html = await resWeb.text();
        const snippets = [...html.matchAll(/class="result__snippet[^>]*>([\s\S]*?)<\/a>/g)]
            .slice(0, 6)
            .map(m => m[1].replace(/<[^>]+>/g, '').trim());
        if (snippets.length > 0) {
            resultadosTexto.push("RESULTADOS DA WEB:\n" + snippets.map(s => `- ${s}`).join('\n'));
        }

        const contextoWeb = resultadosTexto.join('\n\n') || "Nenhum resultado encontrado.";
        const agora = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });

        const res = await chamarGeminiComRetry({
            contents: [{ role: 'user', parts: [{ text: `Data/Hora atual: ${agora}.\nHistórico recente:\n${obterResumoHistorico()}\n\nResultados da pesquisa na internet:\n${contextoWeb}\n\nCom base nesses dados atualizados, responda de forma direta e clara à pergunta: ${perguntaUsuario}` }] }]
        });
        return res.text || contextoWeb;
    } catch (err) {
        return `Erro ao pesquisar na web: ${err.message}`;
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
                    consulta: { type: Type.STRING, description: "Termos de busca objetivos (ex: 'proximo jogo Sao Paulo FC data horario', 'clima Sao Paulo agora')" }
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
            description: "Lista todos os compromissos agendados pelo usuário.",
            parameters: {
                type: Type.OBJECT,
                properties: {
                    dias: { type: Type.NUMBER, description: "Quantidade de dias (padrão 7)." }
                }
            }
        },
        {
            name: "criarEventoAgenda",
            description: "Cria e salva um novo compromisso na agenda do usuário.",
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
        },
        {
            name: "cancelarEventoAgenda",
            description: "Cancela, remove ou apaga um compromisso específico da agenda pelo número do item ou pelo nome/assunto.",
            parameters: {
                type: Type.OBJECT,
                properties: {
                    identificador: { type: Type.STRING, description: "Número do item na agenda (ex: '1', '2') ou palavra-chave do título do compromisso (ex: 'TI', 'reunião')." }
                },
                required: ["identificador"]
            }
        },
        {
            name: "reagendarEventoAgenda",
            description: "Altera a data/hora (reagenda) ou o título de um compromisso já existente na agenda.",
            parameters: {
                type: Type.OBJECT,
                properties: {
                    identificador: { type: Type.STRING, description: "Número do item na agenda (ex: '1') ou palavra-chave do compromisso atual (ex: 'TI', 'reunião')." },
                    novaDataHoraInicio: { type: Type.STRING, description: "Nova data e hora de início no formato ISO YYYY-MM-DDTHH:MM:SS" },
                    novoTitulo: { type: Type.STRING, description: "Novo título opcional caso o usuário queira mudar o nome do evento." }
                },
                required: ["identificador", "novaDataHoraInicio"]
            }
        }
    ]
}];

const INSTRUCAO_SISTEMA = () => {
    const agora = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
    const resumoAgendaAtual = lerAgendaLocal().map((e, i) => `${i + 1}. ${e.dataFormatada} - ${e.titulo}`).join(' | ') || 'Vazia';

    return `Você é um Assistente Executivo Pessoal de elite no WhatsApp.
Data/hora atual em São Paulo: ${agora}.
Compromissos atuais na agenda do usuário: ${resumoAgendaAtual}.

HISTÓRICO RECENTE DA CONVERSA (Use para manter o contexto de pedidos anteriores!):
${obterResumoHistorico()}

SUAS CAPACIDADES PRINCIPAIS:
1. MEMÓRIA CONTÍNUA:
   - Leve sempre em conta o HISTÓRICO RECENTE DA CONVERSA acima quando o usuário fizer pedidos curtos como "deixe mais formal", "e amanhã?", "mude para as 15h", etc.

2. LEITURA E TRANSCRIÇÃO DE ÁUDIOS:
   - Quando receber um áudio, sempre inicie a resposta mostrando:
     📝 *Transcrição do Áudio:* "[texto exato falado no áudio]"
   - Se o áudio contiver um pedido (ex: criar/cancelar/reagendar compromisso, consultar planilha, consultar PDF, analisar foto, buscar na web, reescrever um texto), acione a ferramenta correspondente ou execute a ação logo abaixo da transcrição!

3. GESTÃO COMPLETA DA AGENDA:
   - Para listar compromissos: use "listarEventosAgenda".
   - Para criar novo compromisso: use "criarEventoAgenda".
   - Para cancelar/apagar um compromisso específico: use "cancelarEventoAgenda".
   - Para mudar o horário/data de um compromisso existente: use "reagendarEventoAgenda".

4. REESCRITA DE TEXTOS E E-MAILS:
   - Sempre que o usuário pedir para reescrever, melhorar ou redigir um texto, comunicado ou e-mail num tom específico (executivo, formal, persuasivo, diplomático, direto), entregue o texto pronto para copiar e colar.

5. FERRAMENTAS (WEB, EXCEL, PDF E IMAGEM):
   - Para clima, jogos, notícias ou fatos atuais: use "pesquisarNaInternet".
   - Para dados da planilha Excel enviada: use "consultarPlanilhaExcel".
   - Para perguntas sobre o documento PDF enviado: use "consultarDocumentoPDF".
   - Para perguntas sobre a última foto/print/imagem enviada: use "consultarImagem".

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

        let respostaFinal = "";

        if (response.functionCalls && response.functionCalls.length > 0) {
            const call = response.functionCalls[0];

            if (call.name === "pesquisarNaInternet") {
                respostaFinal = await pesquisarEResponderWeb(textoReferencia, call.args.consulta || textoReferencia);
            } else if (call.name === "consultarPlanilhaExcel") {
                respostaFinal = await analisarPlanilhaComIA(textoReferencia, call.args.termoBusca || "");
            } else if (call.name === "consultarDocumentoPDF") {
                respostaFinal = await consultarDocumentoPDF(call.args.pergunta || textoReferencia);
            } else if (call.name === "consultarImagem") {
                respostaFinal = await consultarImagem(call.args.pergunta || textoReferencia);
            } else if (call.name === "listarEventosAgenda") {
                respostaFinal = await listarEventosAgenda();
            } else if (call.name === "criarEventoAgenda") {
                respostaFinal = await criarEventoAgenda(call.args.titulo, call.args.dataHoraInicio, call.args.duracaoMinutos, call.args.descricao);
            } else if (call.name === "cancelarEventoAgenda") {
                respostaFinal = await cancelarEventoAgenda(call.args.identificador);
            } else if (call.name === "reagendarEventoAgenda") {
                respostaFinal = await reagendarEventoAgenda(call.args.identificador, call.args.novaDataHoraInicio, call.args.novoTitulo);
            }
        } else {
            respostaFinal = response.text || "Processado com sucesso.";
        }

        // Salva a pergunta e a resposta na Memória de Conversa (Melhoria 1)
        adicionarAoHistorico('user', textoReferencia);
        adicionarAoHistorico('assistant', respostaFinal);

        return respostaFinal;
    } catch (err) {
        return `⚠️ Erro técnico: ${err.message?.substring(0, 150)}`;
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
        browser: ['Ubuntu', 'Chrome', '20.0.04'],
        keepAliveIntervalMs: 30000,
        syncFullHistory: false,
        markOnlineOnConnect: false
    });

    sockGlobal = sock;
    sock.ev.on('creds.update', saveCreds);

    sock.ev.on('connection.update', async (update) => {
        const { connection, lastDisconnect, qr } = update;

        if (qr && !sock.authState.creds.registered) {
            statusConexao = 'Pronto para gerar código de 8 dígitos';
            if (MEU_NUMERO && !codigoPareamento) {
                try {
                    const numLimpo = MEU_NUMERO.replace(/\D/g, '');
                    codigoPareamento = await sock.requestPairingCode(numLimpo);
                } catch (e) {}
            }
        }

        if (connection === 'open') {
            codigoPareamento = null;
            statusConexao = 'Conectado!';
            const meuNum = sock.user?.id ? sock.user.id.split(':')[0].split('@')[0] : "";
            if (meuNum) meuDestinoGlobal = `${meuNum}@s.whatsapp.net`;
            console.log('WhatsApp conectado com Memória + Voz TTS + Agenda Completa!');
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
            meuDestinoGlobal = destino;
            const conteudoMsg = msg.message.ephemeralMessage?.message || msg.message;

            const texto = conteudoMsg.conversation ||
                          conteudoMsg.extendedTextMessage?.text || "";

            if (texto.startsWith('🤖 *Assistente IA') || texto.startsWith('🔔 *LEMBRETE') || texto.startsWith('📊 *Planilha') || texto.startsWith('📄 *PDF') || texto.startsWith('📅 *Seus Compromissos')) continue;

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
                    text: `📊 *Planilha "${doc.fileName}" salva com sucesso!*\n• *Abas:* ${wb.SheetNames.join(', ')}\n• *Linhas na 1ª aba:* ${linhas.length}\n• *Colunas:* ${colunas}\n\nPode consultá-la por texto com *!* ou enviando um *áudio*!`
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
                    adicionarAoHistorico('user', `[Enviou PDF: ${doc.fileName}] ${legendaPdf}`);
                    adicionarAoHistorico('assistant', analisePdf);

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
                    adicionarAoHistorico('user', `[Enviou Imagem] ${instrucaoImg}`);
                    adicionarAoHistorico('assistant', respostaImg);

                    await sock.sendMessage(destino, {
                        text: `🤖 *Assistente IA (Análise de Imagem):*\n${respostaImg}\n\n💡 _A foto ficou salva! Pode fazer mais perguntas sobre ela usando *!* ou enviando um áudio._`
                    });
                } catch (err) {
                    await sock.sendMessage(destino, { text: `⚠️ Erro ao processar a imagem: ${err.message}` });
                }
                continue;
            }

            // 4. ÁUDIO / VOZ (Ignora os próprios áudios de resposta gerados pelo bot)
            const audioMsg = conteudoMsg.audioMessage;
            if (audioMsg) {
                // Evita loop infinito: o áudio enviado pelo bot é 'audio/wav', enquanto o gravado no WhatsApp é 'audio/ogg' ou 'audio/mp4'
                if (audioMsg.mimetype && audioMsg.mimetype.includes('wav')) continue;

                try {
                    const bufferAudio = await downloadMediaMessage(msg, 'buffer', {});
                    const base64Audio = bufferAudio.toString('base64');
                    const mimeType = (audioMsg.mimetype || 'audio/ogg').split(';')[0];

                    const partesAudio = [
                        { inlineData: { mimeType: mimeType, data: base64Audio } },
                        { text: "Transcreva este áudio em português. Em seguida, considerando o histórico da conversa, se houver algum pedido, pergunta, consulta de planilha Excel, consulta de PDF, análise de imagem, agendamento/cancelamento/reagendamento de compromisso ou pedido para reescrever texto contido no áudio, execute-o completamente!" }
                    ];

                    const respostaAudio = await processarComIA(partesAudio, "[Mensagem de Voz do Usuário]");
                    await sock.sendMessage(destino, { text: `🤖 *Assistente IA:*\n${respostaAudio}` });

                    // MELHORIA 4: Gera e envia também a resposta falada em Áudio!
                    const bufferVoz = await gerarAudioRespostaTTS(respostaAudio);
                    if (bufferVoz) {
                        await sock.sendMessage(destino, {
                            audio: bufferVoz,
                            mimetype: 'audio/wav',
                            ptt: true
                        });
                    }
                } catch (err) {
                    await sock.sendMessage(destino, { text: `⚠️ Erro ao processar áudio: ${err.message}` });
                }
                continue;
            }

            // 5. TEXTO COM "!"
            if (!texto.startsWith('!')) continue;

            const perguntaLimpa = texto.slice(1).trim();
            if (!perguntaLimpa) continue;

            if (perguntaLimpa.toLowerCase() === 'limpar memoria' || perguntaLimpa.toLowerCase() === 'limpar memória') {
                historicoConversa = [];
                await sock.sendMessage(destino, { text: `🤖 *Assistente IA:*\n🧠 Memória de conversa reiniciada com sucesso!` });
                continue;
            }

            if (perguntaLimpa.toLowerCase() === 'limpar agenda') {
                salvarAgendaLocal([]);
                await sock.sendMessage(destino, { text: `🤖 *Assistente IA:*\n🗑️ Todos os compromissos da agenda foram apagados com sucesso!` });
                continue;
            }

            if (perguntaLimpa.toLowerCase() === 'agenda') {
                const statusAgenda = await listarEventosAgenda();
                await sock.sendMessage(destino, { text: `🤖 *Assistente IA:*\n${statusAgenda}` });
                continue;
            }

            // Verifica se o usuário pediu resposta em voz por texto (ex: "!voz ..." ou "responda em áudio")
            const pediuVozNoTexto = perguntaLimpa.toLowerCase().startsWith('voz ') ||
                                    perguntaLimpa.toLowerCase().includes('em áudio') ||
                                    perguntaLimpa.toLowerCase().includes('em audio') ||
                                    perguntaLimpa.toLowerCase().includes('por voz');

            const queryFinal = perguntaLimpa.toLowerCase().startsWith('voz ')
                ? perguntaLimpa.slice(4).trim()
                : perguntaLimpa;

            const resposta = await processarComIA([{ text: queryFinal }], queryFinal);
            await sock.sendMessage(destino, { text: `🤖 *Assistente IA:*\n${resposta}` });

            if (pediuVozNoTexto) {
                const bufferVoz = await gerarAudioRespostaTTS(resposta);
                if (bufferVoz) {
                    await sock.sendMessage(destino, {
                        audio: bufferVoz,
                        mimetype: 'audio/wav',
                        ptt: true
                    });
                }
            }
        }
    });
}

iniciarWhatsApp();
