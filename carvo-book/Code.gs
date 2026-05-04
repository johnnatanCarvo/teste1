// ============================================================
// CARVO BOOK — Configurações
// ============================================================
const GEMINI_API_KEY = 'COLE_SUA_CHAVE_AQUI'; // aistudio.google.com → Get API Key (gratuito)
const NOME_PLANILHA  = 'Carvo Book - Relatórios';

// ============================================================
// ROTA PRINCIPAL — serve o site
// ============================================================
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Carvo Book')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// PROCESSA TRANSCRIÇÃO COM GEMINI (gratuito)
// ============================================================
function processarTranscricao(transcricao) {
  if (!transcricao || transcricao.trim().length < 10) {
    throw new Error('Transcrição muito curta. Grave a reunião antes de gerar o relatório.');
  }

  const prompt = `Você é um assistente especializado em análise de reuniões de negócios em português brasileiro.

Analise a transcrição abaixo e gere um relatório estruturado em português com:

1. RESUMO EXECUTIVO — síntese clara da reunião em 3 a 5 linhas
2. PONTOS DISCUTIDOS — lista dos principais temas abordados
3. DECISÕES TOMADAS — o que foi definido ou acordado
4. PRÓXIMOS PASSOS — ações concretas a executar (com responsável se mencionado)
5. SUGESTÕES DA IA — recomendações estratégicas baseadas no conteúdo discutido

Transcrição:
${transcricao}

Retorne APENAS um JSON válido, sem texto adicional, sem blocos de código markdown, com esta estrutura exata:
{
  "resumo": "texto do resumo",
  "pontos_discutidos": ["ponto 1", "ponto 2"],
  "decisoes": ["decisão 1", "decisão 2"],
  "proximos_passos": ["passo 1", "passo 2"],
  "sugestoes_ia": ["sugestão 1", "sugestão 2"]
}`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${GEMINI_API_KEY}`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.2, maxOutputTokens: 2048 }
  };

  const resposta = UrlFetchApp.fetch(url, {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const dados = JSON.parse(resposta.getContentText());

  if (dados.error) {
    throw new Error('Erro Gemini: ' + dados.error.message);
  }

  const texto = dados.candidates[0].content.parts[0].text;
  const jsonLimpo = texto.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();

  try {
    return JSON.parse(jsonLimpo);
  } catch (e) {
    throw new Error('Não foi possível interpretar a resposta da IA. Tente novamente.');
  }
}

// ============================================================
// SALVA NO SHEETS E ENVIA E-MAIL
// ============================================================
function salvarEEnviar(transcricao, relatorio) {
  const dataHora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
  const sheetUrl = salvarNoSheets(relatorio, transcricao, dataHora);
  enviarEmail(relatorio, dataHora, sheetUrl);
  return { sucesso: true, sheetUrl: sheetUrl };
}

function salvarNoSheets(relatorio, transcricao, dataHora) {
  let planilha;
  const arquivos = DriveApp.getFilesByName(NOME_PLANILHA);

  if (arquivos.hasNext()) {
    planilha = SpreadsheetApp.open(arquivos.next());
  } else {
    planilha = SpreadsheetApp.create(NOME_PLANILHA);
  }

  let aba = planilha.getSheetByName('Relatórios');
  if (!aba) {
    aba = planilha.insertSheet('Relatórios');
    const cabecalho = ['Data/Hora', 'Resumo Executivo', 'Pontos Discutidos', 'Decisões', 'Próximos Passos', 'Sugestões IA', 'Transcrição'];
    aba.appendRow(cabecalho);
    const estilo = aba.getRange(1, 1, 1, cabecalho.length);
    estilo.setFontWeight('bold').setBackground('#1a73e8').setFontColor('white').setFontSize(11);
    aba.setFrozenRows(1);
    aba.setColumnWidth(1, 140);
    aba.setColumnWidth(2, 260);
    [3,4,5,6].forEach(c => aba.setColumnWidth(c, 220));
    aba.setColumnWidth(7, 300);
  }

  const formatarLista = arr => arr.map((i, n) => `${n + 1}. ${i}`).join('\n');

  aba.appendRow([
    dataHora,
    relatorio.resumo,
    formatarLista(relatorio.pontos_discutidos),
    formatarLista(relatorio.decisoes),
    formatarLista(relatorio.proximos_passos),
    formatarLista(relatorio.sugestoes_ia),
    transcricao
  ]);

  const linha = aba.getLastRow();
  aba.getRange(linha, 1, 1, 7).setWrap(true).setVerticalAlignment('top');
  aba.setRowHeight(linha, 130);

  // Linha alternada
  if (linha % 2 === 0) {
    aba.getRange(linha, 1, 1, 7).setBackground('#f1f3f4');
  }

  return planilha.getUrl();
}

function enviarEmail(relatorio, dataHora, sheetUrl) {
  const email = Session.getActiveUser().getEmail();

  const listaHtml = arr => arr.map(i => `<li style="margin-bottom:6px">${i}</li>`).join('');

  const corpo = `<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:680px;margin:0 auto;padding:20px;color:#333;background:#f5f5f5">

  <div style="background:#1a73e8;padding:24px 28px;border-radius:10px 10px 0 0">
    <h1 style="color:white;margin:0;font-size:22px;letter-spacing:.5px">Carvo Book</h1>
    <p style="color:#c2d7ff;margin:6px 0 0;font-size:14px">Relatório de Reunião &mdash; ${dataHora}</p>
  </div>

  <div style="background:white;padding:24px 28px;border-radius:0 0 10px 10px;box-shadow:0 2px 6px rgba(0,0,0,.08)">

    <div style="border-left:4px solid #1a73e8;padding-left:14px;margin-bottom:22px">
      <h2 style="color:#1a73e8;margin:0 0 8px;font-size:16px">Resumo Executivo</h2>
      <p style="margin:0;line-height:1.6">${relatorio.resumo}</p>
    </div>

    <div style="border-left:4px solid #34a853;padding-left:14px;margin-bottom:22px">
      <h2 style="color:#34a853;margin:0 0 8px;font-size:16px">Pontos Discutidos</h2>
      <ul style="margin:0;padding-left:18px">${listaHtml(relatorio.pontos_discutidos)}</ul>
    </div>

    <div style="border-left:4px solid #f57c00;padding-left:14px;margin-bottom:22px">
      <h2 style="color:#f57c00;margin:0 0 8px;font-size:16px">Decisões Tomadas</h2>
      <ul style="margin:0;padding-left:18px">${listaHtml(relatorio.decisoes)}</ul>
    </div>

    <div style="border-left:4px solid #ea4335;padding-left:14px;margin-bottom:22px">
      <h2 style="color:#ea4335;margin:0 0 8px;font-size:16px">Próximos Passos</h2>
      <ul style="margin:0;padding-left:18px">${listaHtml(relatorio.proximos_passos)}</ul>
    </div>

    <div style="border-left:4px solid #9c27b0;padding-left:14px;margin-bottom:28px">
      <h2 style="color:#9c27b0;margin:0 0 8px;font-size:16px">Sugestões da IA</h2>
      <ul style="margin:0;padding-left:18px">${listaHtml(relatorio.sugestoes_ia)}</ul>
    </div>

    <div style="text-align:center">
      <a href="${sheetUrl}" style="background:#1a73e8;color:white;padding:13px 28px;border-radius:6px;text-decoration:none;font-weight:bold;font-size:14px;display:inline-block">
        Abrir Planilha Completa
      </a>
    </div>

  </div>

  <p style="color:#aaa;font-size:11px;text-align:center;margin-top:16px">
    Gerado automaticamente pelo Carvo Book
  </p>

</body></html>`;

  GmailApp.sendEmail(email, `[Carvo Book] Reunião de ${dataHora}`, '', {
    htmlBody: corpo,
    name: 'Carvo Book'
  });
}
