// 🔑 Constantes fixas
const WEBAPP_URL = "https://script.google.com/macros/s/AKfycbznXr_UYCcq4-6rmGsxCxXJs7Df0ABzA3EkBgajbXC42NvH6c0AQdUDc6WZ15XV_3JW/exec";
const SHEET_URL  = "https://docs.google.com/spreadsheets/d/1Pd7jw-1OkUYKvPCw5Z20ckoeDRe211UMn4w4Bud0CKY/edit";
const TZ         = "America/Sao_Paulo";

// === Config extra para coletar comentários por e-mail (NÃO altera seu onFormSubmit) ===
const SHEET_NAME = 'Respostas ao formulário 1'; // nome da aba da sua imagem
const COMMENTS_COL = 10;             // J
const ID_COL = 11;                   // K
// Aprovadores que podem clicar em Aprovar / Reprovar
const APROVADORES = [
  'rober@grupoorion.com.br',
  'lucas.garcia@grupoorion.com.br',
  'qualidade.orion.sp@gmail.com'
];

// Versão em minúsculas para comparação
const APROVADORES_NORMALIZADOS = APROVADORES.map(e => e.toLowerCase());

// Se ainda quiser ter um "aprovador principal" (para receber o e-mail direto)
const APROVADOR_EMAIL = APROVADORES[0]; // continua sendo o Rober

const GMAIL_LABEL_PROCESSED = 'APP-APROVACAO/PROCESSADO';

function onFormSubmit(e) {
  const sh  = e.range.getSheet();
  const row = e.range.getRow();
  const r   = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];

  // Colunas do formulário
  const carimbo = r[0]; // A
  const nome    = r[2]; // C
  const app     = r[3]; // D
  const desc    = r[4]; // E
  const obs     = r[5]; // F

  // ✅ Gerar ID numérico na coluna K (coluna 11), iniciando em 250
  const idCol  = 11; // K
  const idCell = sh.getRange(row, idCol);
  if (!idCell.getValue()) {
    const nextId = 249 + (row - 1); // linha 2 → 250, linha 3 → 251...
    idCell.setValue(nextId);
  }
  const idValue = idCell.getValue();

  // Links de aprovação
  const linkAprovar  = `${WEBAPP_URL}?acao=aprovar&linha=${row}`;
  const linkReprovar = `${WEBAPP_URL}?acao=reprovar&linha=${row}`;

  // E-mails
  const emailGestor = "rober@grupoorion.com.br";
  const emailsCopia = "lucas.garcia@grupoorion.com.br, cilene.silva@grupoorion.com.br";
  const assunto     = `Solicitação de Aprovação - ${app} (${idValue})`;

  // Helper para botões
  const btn = (href, label, bg) => `
    <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="display:inline-block;margin-right:10px;">
      <tr>
        <td bgcolor="${bg}" style="border-radius:6px;">
          <a href="${href}" target="_blank"
             style="font-family:Arial,sans-serif;font-size:14px;line-height:14px;text-decoration:none;
                    color:#ffffff;padding:12px 18px;display:inline-block;">${label}</a>
        </td>
      </tr>
    </table>`;

  // Logo
  const logoUrl  = "https://drive.google.com/uc?export=download&id=1RqQxKlSmnBnDPbg2HoLkV6QbRVgXr0il";
  const logoBlob = UrlFetchApp.fetch(logoUrl).getBlob().setName("logo.png");

  // Data formatada
  const dataFmt = Utilities.formatDate(new Date(carimbo), TZ, "dd/MM/yyyy HH:mm");

  // ✅ Ajuste de formatação para preservar quebras de linha
  const descFmt = (desc || "-").replace(/\n/g, "<br>");
  const obsFmt  = (obs  || "-").replace(/\n/g, "<br>");

  // Corpo do e-mail
  const corpoHtml = `
  <div style="font-family:Arial,Helvetica,sans-serif;color:#333;max-width:760px;">
    <div style="margin:12px 0 20px;">
      <img src="cid:logo" alt="Orion" style="height:40px;display:block;">
    </div>

    <h2 style="color:#2E86C1;margin:0 0 16px;">Solicitação de Aprovação - ${app} (${idValue})</h2>

    <p>Olá Rober, tudo bem?</p>
    <p>Segue os detalhes da nova solicitação referente ao APP:</p>

    <ul style="margin:0 0 16px;padding-left:20px;">
      <li><b>ID:</b> ${idValue}</li>
      <li><b>Data/Hora:</b> ${dataFmt}</li>
      <li><b>Solicitante:</b> ${nome || "-"}</li>
      <li><b>Aplicativo:</b> ${app || "-"}</li>
    </ul>

    <p><b>Descrição:</b><br>${descFmt}</p>
    <p><b>Observação:</b><br>${obsFmt}</p>

    <p>Clique abaixo para registrar a decisão:</p>

    <div style="margin:6px 0 16px;">
      ${btn(linkAprovar, "Aprovar",  "#28a745")}
      ${btn(linkReprovar, "Reprovar", "#dc3545")}
    </div>

    <p style="margin:20px 0 0;">
      <a href="${SHEET_URL}" style="color:#1a73e8;text-decoration:none;">Acessar a planilha</a>
    </p>

    <p style="margin-top:16px;font-size:13px;color:#555;">
      Caso deseje adicionar um comentário, basta abrir a planilha pelo link acima
      e escrever diretamente na coluna <b>Comentários</b> da linha correspondente.
    </p>

    <p style="margin-top:16px;font-size:13px;color:#555;">
      Caso seja exibida alguma mensagem de erro do Google ao clicar em
      <b>“Aprovar”</b> ou <b>“Reprovar”</b>, não se preocupe.<br>
      A decisão será registrada corretamente na planilha, trata-se apenas de uma
      notificação de verificação de domínio da plataforma Google.<br>
      Para confirmar, basta acessar a planilha e verificar o resultado da aprovação.
    </p>

    <p style="margin-top:24px;">Atenciosamente,<br>Equipe de Qualidade Orion.</p>
  </div>`;

  // Envio do e-mail (com cópia para Lucas e Cilene)
  GmailApp.sendEmail(
    emailGestor,
    assunto,
    " ",
    {
      htmlBody: corpoHtml,
      inlineImages: { logo: logoBlob },
      cc: emailsCopia
    }
  );

}

function reenviarPorLinha() {
  const TZ = "America/Sao_Paulo";
  const WEBAPP_URL = "https://script.google.com/macros/s/AKfycbznXr_UYCcq4-6rmGsxCxXJs7Df0ABzA3EkBgajbXC42NvH6c0AQdUDc6WZ15XV_3JW/exec";
  const SHEET_URL  = "https://docs.google.com/spreadsheets/d/1Pd7jw-1OkUYKvPCw5Z20ckoeDRe211UMn4w4Bud0CKY/edit";
  const SHEET_NAME = "Respostas ao formulário 1";

  // IDs das solicitações a reenviar (coluna K)
  const IDS_REENVIAR = [271, 272];

  // --- Envio oficial ---
  const APROVADOR_EMAIL = "rober@grupoorion.com.br";
  const emailsCopia = "lucas.garcia@grupoorion.com.br, cilene.silva@grupoorion.com.br";


  const ss = SpreadsheetApp.openByUrl(SHEET_URL);
  const sh = ss.getSheetByName(SHEET_NAME);
  const logoUrl = "https://drive.google.com/uc?export=download&id=1RqQxKlSmnBnDPbg2HoLkV6QbRVgXr0il";
  const logoBlob = UrlFetchApp.fetch(logoUrl).getBlob().setName("logo.png");

  let reenviados = 0;

  // 🔍 Para cada ID, localizar a linha correspondente
  for (const id of IDS_REENVIAR) {
    const ultimaLinha = sh.getLastRow();
    const idsColuna = sh.getRange(2, 11, ultimaLinha - 1).getValues().flat();
    const linha = idsColuna.indexOf(id) + 2; // +2 pois a contagem começa na linha 2

    if (linha < 2) {
      Logger.log(`❌ ID ${id} não encontrado na planilha.`);
      continue;
    }

    const r = sh.getRange(linha, 1, 1, sh.getLastColumn()).getValues()[0];
    const carimbo = r[0];
    const nome = r[2];
    const app = r[3];
    const desc = r[4];
    const obs = r[5];
    const status = (r[6] || "").toString().trim();
    const idValue = r[10]; // K

    // Se já estiver aprovado/reprovado, pula
    if (status.toLowerCase() === "aprovado" || status.toLowerCase() === "reprovado") continue;

    const linkAprovar  = `${WEBAPP_URL}?acao=aprovar&linha=${linha}`;
    const linkReprovar = `${WEBAPP_URL}?acao=reprovar&linha=${linha}`;
    const assunto = `Reenvio: Solicitação de Aprovação - ${app} (${idValue})`;
    const dataFmt = Utilities.formatDate(new Date(carimbo), TZ, "dd/MM/yyyy HH:mm");

    const btn = (href, label, bg) => `
      <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="display:inline-block;margin-right:10px;">
        <tr>
          <td bgcolor="${bg}" style="border-radius:6px;">
            <a href="${href}" target="_blank"
               style="font-family:Arial,sans-serif;font-size:14px;line-height:14px;text-decoration:none;
                      color:#ffffff;padding:12px 18px;display:inline-block;">${label}</a>
          </td>
        </tr>
      </table>`;

    const descFmt = (desc || "-").replace(/\n/g, "<br>");
    const obsFmt  = (obs  || "-").replace(/\n/g, "<br>");

    const corpoHtml = `
    <div style="font-family:Arial,Helvetica,sans-serif;color:#333;max-width:760px;">
      <div style="margin:12px 0 20px;">
        <img src="cid:logo" alt="Orion" style="height:40px;display:block;">
      </div>

      <h2 style="color:#2E86C1;margin:0 0 16px;">Solicitação de Aprovação - ${app} (${idValue})</h2>

      <p>Olá Rober, tudo bem?</p>
      <p>Este é um <b>reenvio automático</b> de uma solicitação de aprovação pendente.</p>

      <ul style="margin:0 0 16px;padding-left:20px;">
        <li><b>ID:</b> ${idValue}</li>
        <li><b>Data/Hora:</b> ${dataFmt}</li>
        <li><b>Solicitante:</b> ${nome || "-"}</li>
        <li><b>Aplicativo:</b> ${app || "-"}</li>
      </ul>

      <p><b>Descrição:</b><br>${descFmt}</p>
      <p><b>Observação:</b><br>${obsFmt}</p>

      <p>Clique abaixo para registrar a decisão:</p>

      <div style="margin:6px 0 16px;">
        ${btn(linkAprovar, "Aprovar",  "#28a745")}
        ${btn(linkReprovar, "Reprovar", "#dc3545")}
      </div>

      <p style="margin:20px 0 0;">
        <a href="${SHEET_URL}" style="color:#1a73e8;text-decoration:none;">Acessar a planilha</a>
      </p>

      <p style="margin-top:16px;font-size:13px;color:#555;">
        Caso apareça alguma mensagem de erro ao clicar em <b>“Aprovar”</b> ou <b>“Reprovar”</b>,
        a decisão ainda será gravada corretamente na planilha.
      </p>

      <p style="margin-top:24px;">Atenciosamente,<br>Equipe de Qualidade Orion.</p>
    </div>`;

    GmailApp.sendEmail(
      APROVADOR_EMAIL,
      assunto,
      " ",
      {
        htmlBody: corpoHtml,
        inlineImages: { logo: logoBlob },
        cc: emailsCopia
      }
    );

    const hoje = Utilities.formatDate(new Date(), TZ, "dd/MM/yyyy HH:mm");
    sh.getRange(linha, 8).setValue(`Reenviado em ${hoje}`);
    reenviados++;
  }

  SpreadsheetApp.getActive().toast(`✅ Foram reenviadas ${reenviados} solicitações (IDs ${IDS_REENVIAR.join(", ")})`);
  Logger.log(`Total reenviado: ${reenviados}`);
}


// ——— Helpers ———

// Pega apenas o texto digitado sem assinatura do Outlook (Orion)
function limparComentario_(plain) {
  let txt = (plain || '').replace(/\r/g, '').trim();

  // 1) Corta no início de citação/resposta
  const quoteCuts = [
    /^Em .*?escreveu:/mi,       // "Em 30/09/2025, Fulano escreveu:"
    /^From:\s*/mi,              // "From:"
    /^De:\s*/mi,                // "De:"
    /^> /m,                     // linhas iniciadas com ">"
    /^-----Mensagem original-----/mi
  ];
  for (const re of quoteCuts) {
    if (re.test(txt)) { txt = txt.split(re)[0]; break; }
  }

  // 2) Corta em separadores de assinatura comuns
  const sigCuts = [
    /^--\s*$/m,                 // "--"
    /^_{5,}\s*$/m,              // "_____"
    /^[-–—]{5,}\s*$/m,          // "-----" "———"
    /^Atenciosamente[,:]?\s*$/mi,
    /^Att[.,:]?\s*$/mi,
    /^Assinatura eletr[oô]nica/mi
  ];
  for (const re of sigCuts) {
    if (re.test(txt)) { txt = txt.split(re)[0]; break; }
  }

  // 3) Remove linhas típicas de assinatura/rodapé
  let lines = txt.split('\n');

  const dropLine = (s) => {
    const t = s.trim();
    if (!t) return false;
    if (/^\s*(https?:\/\/|www\.)/i.test(t)) return true;                // URLs puras
    if (/\[cid:[^\]]+\]/i.test(t) || /<cid:[^>]+>/i.test(t)) return true; // imagens inline cid
    if (/<https?:\/\/[^>]+>/i.test(t)) return true;                     // <https://...>
    if (/bookwithme|outlook\.office\.com\/bookwithme/i.test(t)) return true; // link de agenda
    if (/agendar|agenda|reserv(ar|e) um hor[aá]rio/i.test(t)) return true;   // frases de agenda
    if (/^enviado do meu/i.test(t)) return true;                        // “Enviado do meu iPhone…”
    return false;
  };

  lines = lines.filter(line => !dropLine(line));

  // 4) Pega só o primeiro bloco de texto (para não puxar assinatura longa)
  const resLines = [];
  let blankStreak = 0;
  for (const line of lines) {
    const isBlank = line.trim() === '';
    if (isBlank) blankStreak++; else blankStreak = 0;
    if (blankStreak >= 2) break;  // parou após duas linhas em branco seguidas
    resLines.push(line);
  }

  let res = resLines.join('\n').trim();

  // 5) Normaliza quebras e limita tamanho
  res = res.replace(/\n{3,}/g, '\n\n');
  if (res.length > 1200) res = res.slice(0, 1200) + '…';

  // 6) Filtra autorrespostas
  if (/^out of office|^auto.?reply/i.test(res)) return '';

  return res;
}

// encontra a linha com o ID na coluna informada (retorna nº da linha ou null)
function encontrarLinhaPorId_(sh, id, col) {
  const last = sh.getLastRow();
  if (last < 2) return null;
  const vals = sh.getRange(2, col, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (Number(vals[i][0]) === Number(id)) return i + 2;
  }
  return null;
}

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

// ——— Coletor: lê Gmail, pega a ÚLTIMA resposta do Rober e grava na coluna J ———
function processarRespostasPorEmail() {
  // procura conversas recentes que provavelmente são do fluxo
const threads = GmailApp.search(
  `from:${APROVADOR_EMAIL} -label:"${GMAIL_LABEL_PROCESSED}"`,
  0,
  100
);



  const ss = SpreadsheetApp.openByUrl(SHEET_URL);
  const sh = ss.getSheetByName(SHEET_NAME);
  const processedLabel = getOrCreateLabel_(GMAIL_LABEL_PROCESSED);

  threads.forEach(thread => {
    const msgs = thread.getMessages();

    // 1) Descobrir o ID pelo ASSUNTO ATUAL que você já envia: "... (123)"
    // tenta primeiro "(123)" no fim; se não achar, procura qualquer número entre parênteses
    let subject = (msgs[msgs.length - 1].getSubject() || '').trim();
    let m = subject.match(/\((\d+)\)\s*$/);           // ... (123)
    if (!m) m = subject.match(/\((\d+)\)/);           // fallback: qualquer (123)
    if (!m) return;
    const idNum = Number(m[1]);

    // 2) Achar a ÚLTIMA mensagem do aprovador nesse thread
   let ultimaDoAprovador = null;
for (let i = msgs.length - 1; i >= 0; i--) {
  const msg = msgs[i];
  if (msg.isDraft()) continue;

  // Captura o campo "From" e converte para minúsculas
  const fromHdr = (msg.getFrom() || '').toLowerCase();

  // 🔍 Log para depuração — mostra exatamente como o Gmail retornou o remetente
  Logger.log("FROM detectado: " + fromHdr);

  // Extrai e-mail se estiver entre < >, senão usa o texto completo
  const emailMatch = fromHdr.match(/<([^>]+)>/);
  const fromEmail = (emailMatch ? emailMatch[1] : fromHdr).trim();

  // ✅ Correção: usa includes() em vez de igualdade rígida
  if (fromHdr.includes(APROVADOR_EMAIL.toLowerCase()) || fromEmail === APROVADOR_EMAIL.toLowerCase()) {
    ultimaDoAprovador = msg;
    break;
  }
}

    if (!ultimaDoAprovador) return; // Ninguém relevante respondeu

    // 3) Vai extrair apenas o que ele digitou acima da citação!
    const raw = ultimaDoAprovador.getPlainBody() || '';
    const comentario = limparComentario_(raw);
    if (!comentario) { ultimaDoAprovador.markRead(); thread.addLabel(processedLabel); return; }

    // 4) Gravar SEMPRE sobrescrevendo a coluna J da linha com ID em K
    const row = encontrarLinhaPorId_(sh, idNum, ID_COL);
    if (row) {
      sh.getRange(row, COMMENTS_COL).setValue(comentario);

      // opcional: atualizar "Data Atendimento" (H = 8)
      const hoje = Utilities.formatDate(new Date(), TZ, "dd/MM/yyyy");
      sh.getRange(row, 8).setValue(hoje);
    }

    // 5) Organização (opcional)
    ultimaDoAprovador.markRead();
    thread.addLabel(processedLabel);
    // thread.moveToArchive(); // se quiser arquivar
  });
}

function doGet(e) {
  try {
    // ===== Configuração básica =====
    const versao = "Versão 3.1.3 - Equipe de Qualidade Orion";
    const LOGO_URL = "https://lh3.googleusercontent.com/d/1B8bS5fljAR-OdHEVQjE4Eb-Wpa9wndty=w600";
    const TZ = "America/Sao_Paulo";
    const SHEET_URL = "https://docs.google.com/spreadsheets/d/1Pd7jw-1OkUYKvPCw5Z20ckoeDRe211UMn4w4Bud0CKY/edit";
    const SHEET_NAME = "Respostas ao formulário 1";
    const novoWebApp = "https://script.google.com/macros/s/AKfycbznXr_UYCcq4-6rmGsxCxXJs7Df0ABzA3EkBgajbXC42NvH6c0AQdUDc6WZ15XV_3JW/exec";

    // ===== Redireciona se for link antigo =====
    const query = e ? e.queryString : "";
    const currentUrl = ScriptApp.getService().getUrl ? ScriptApp.getService().getUrl() : "";
    if (currentUrl && currentUrl.includes("AKfycbzx8K2ET4_")) {
      return HtmlService.createHtmlOutput(
        `<meta http-equiv="refresh" content="0; url='${novoWebApp}?${query}'" />`
      );
    }

    // ===== Template de layout =====
    function gerarHtml(title, content, corTitulo = "#111827") {
      return `
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <title>${title}</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    body {
      margin: 0;
      font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
      background: #f3f4f6;
      color: #111827;
      display: flex;
      align-items: center;
      justify-content: center;
      min-height: 100vh;
    }
    .card {
      background: #fff;
      padding: 24px 28px;
      border-radius: 10px;
      box-shadow: 0 10px 25px rgba(15,23,42,.08);
      max-width: 460px;
      width: 100%;
      box-sizing: border-box;
      text-align: center;
    }
    .logo {
      margin-bottom: 20px;
      opacity: 0;
      animation: fadeIn 1.2s ease-out forwards;
    }
    @keyframes fadeIn {
      from {opacity: 0; transform: translateY(-6px);}
      to {opacity: 1; transform: translateY(0);}
    }
    .logo img { height: 48px; }
    h1 { margin: 0 0 8px; font-size: 20px; color: ${corTitulo}; }
    p { margin: 4px 0; font-size: 14px; color: #4b5563; }
    .btn-row { margin-top: 18px; display: flex; gap: 10px; flex-wrap: wrap; justify-content: center; }
    .btn { padding: 9px 16px; border-radius: 999px; font-size: 14px; font-weight: 500; text-decoration: none; border: none; cursor: pointer; }
    .btn-primary { background: #2563eb; color: #fff; }
    .small { margin-top: 10px; font-size: 12px; color: #9ca3af; }
  </style>
</head>
<body>
  <div class="card">
    <div class="logo">
      <img src="${LOGO_URL}" alt="Logo Orion" referrerpolicy="no-referrer">
    </div>
    ${content}
    <p class="small">${versao}</p>
  </div>
</body>
</html>`;
    }

    // ===== Parâmetros =====
    e = e || {};
    const params = e.parameter || {};
    const acao = String(params.acao || "").toLowerCase();
    const linha = Number(params.linha || 0);

    // ===== Validação =====
    if (!acao || !Number.isInteger(linha) || linha < 2) {
      const content = `
        <h1 style="color:#b45309;">Link inválido</h1>
        <p>Não foi possível identificar a ação ou a linha da solicitação.</p>
        <p>Volte ao e-mail original e clique novamente em <b>Aprovar</b> ou <b>Reprovar</b>.</p>`;
      return HtmlService.createHtmlOutput(gerarHtml("Requisição inválida", content, "#b45309"));
    }

    // ===== Atualiza planilha =====
    const ss = SpreadsheetApp.openByUrl(SHEET_URL);
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      const content = `<h1 style="color:#b91c1c;">Erro</h1><p>Aba <b>${SHEET_NAME}</b> não encontrada.</p>`;
      return HtmlService.createHtmlOutput(gerarHtml("Erro na planilha", content, "#b91c1c"));
    }

    const statusCol = 7; // G
    const dataCol = 8;   // H
    const idCol = 11;    // K
    let statusTxt = "";

   // Lê o comentário do gestor na coluna J (Comentários)
const comentarioGestorRaw = sh.getRange(linha, COMMENTS_COL).getValue();

// Normaliza o comentário para uso no e-mail
const comentarioGestorTxt = (comentarioGestorRaw || "").toString().trim();

const comentarioGestor =
  comentarioGestorTxt && comentarioGestorTxt !== "--------------------"
    ? comentarioGestorTxt
    : "-";



    if (acao === "aprovar") statusTxt = "Aprovado";
    else if (acao === "reprovar") statusTxt = "Reprovado";
    else {
      const content = `<h1 style="color:#b91c1c;">Ação inválida</h1><p>Ação recebida: <b>${acao}</b></p>`;
      return HtmlService.createHtmlOutput(gerarHtml("Ação inválida", content, "#b91c1c"));
    }

    // ===== Atualiza células =====
    sh.getRange(linha, statusCol).setValue(statusTxt);
    const agoraFmt = Utilities.formatDate(new Date(), TZ, "dd/MM/yyyy HH:mm");
    sh.getRange(linha, dataCol).setValue(agoraFmt);
    const idValue = sh.getRange(linha, idCol).getValue();
    if (acao === "aprovar" || acao === "reprovar") {
  const nome = sh.getRange(linha, 3).getValue();      // C
  const app  = sh.getRange(linha, 4).getValue();      // D
  const descricao = sh.getRange(linha, 5).getValue(); // E

  enviarEmailConfirmacaoAprovacao_({
    app,
    idValue,
    nome,
    descricao,
    statusTxt,              // "Aprovado" ou "Reprovado"
    dataHora: agoraFmt,
    sheetUrl: SHEET_URL
  });
}



    // ===== Tela de sucesso =====
    const content = `
      <h1 style="color:#16a34a;">Decisão registrada!</h1>
      <p>Solicitação <b>(${idValue})</b> marcada como <b>${statusTxt}</b>.</p>
      <p>Data/hora: <b>${agoraFmt}</b></p>
      <div class="btn-row">
        <a class="btn btn-primary" href="${SHEET_URL}" target="_blank">Abrir planilha</a>
      </div>
      <p class="small">Verifique sua resposta na planilha e feche a aba com segurança!</p>`;
    return HtmlService.createHtmlOutput(gerarHtml("Decisão registrada", content, "#16a34a"));

  } catch (err) {
    Logger.log("Erro no doGet: " + err);
    const content = `
      <h1 style="color:#b91c1c;">Erro inesperado</h1>
      <p>${err.message || "Ocorreu um erro ao processar sua solicitação."}</p>`;
    return HtmlService.createHtmlOutput(gerarHtml("Erro inesperado", content, "#b91c1c"));
  }
}

// ===== MENU PERSONALIZADO =====
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🔧 Reenvio de Solicitação")
    .addItem("Reenviar Solicitação 📝", "abrirReenvioManual_")
    .addToUi();
}

// Função chamada pelo menu
function abrirReenvioManual_() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Reenviar Solicitação", "Digite o número do ID da solicitação (coluna K):", ui.ButtonSet.OK_CANCEL);

  if (resp.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Ação cancelada.");
    return;
  }

  const idStr = resp.getResponseText().trim();
  const idNum = Number(idStr);
  if (!idNum || isNaN(idNum)) {
    ui.alert("Por favor, digite um número de ID válido.");
    return;
  }

  const confirm = ui.alert(
    "Confirmação",
    `Deseja realmente reenviar a solicitação de ID ${idNum} para Rober?`,
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    ui.alert("Envio cancelado.");
    return;
  }

  // Chama o reenvio individual
  const ok = reenviarSolicitacaoPorId_(idNum);
  if (ok) {
    ui.alert(`✅ Solicitação ${idNum} reenviada com sucesso para Rober.`);
  } else {
    ui.alert(`⚠️ Não foi possível reenviar a solicitação ${idNum}. Verifique se o ID existe ou se já foi aprovado/reprovado.`);
  }

}// ===== MENU PERSONALIZADO =====
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🔧 Reenviar Solicitação")
    .addItem("Reenviar Solicitação 📝", "abrirReenvioManual_")
    .addToUi();
}

function abrirReenvioManual_() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Respostas ao formulário 1");
  if (!sh) {
    ui.alert("❌ Aba 'Respostas ao formulário 1' não encontrada na planilha.");
    return;
  }

  const lastRow = sh.getLastRow();
  const ids = sh.getRange(2, 11, lastRow - 1).getValues().flat(); // Coluna K

  while (true) {
    const resp = ui.prompt("Reenviar Solicitação", "Digite o número do ID da solicitação (coluna K):", ui.ButtonSet.OK_CANCEL);

    if (resp.getSelectedButton() !== ui.Button.OK) {
      ui.alert("Ação cancelada.");
      return;
    }

    const idStr = resp.getResponseText().trim();
    const idNum = Number(idStr);

    if (!idNum || isNaN(idNum)) {
      const retry = ui.alert("ID inválido", "Por favor, digite um número de ID válido.", ui.ButtonSet.YES_NO);
      if (retry === ui.Button.NO) return;
      else continue;
    }

    // 🔍 Verifica se o ID existe na planilha
    const linha = ids.indexOf(idNum) + 2;
    if (linha < 2) {
      const retry = ui.alert(
        "ID não encontrado",
        `⚠️ O ID ${idNum} não foi localizado na coluna K.\n\nDeseja tentar novamente?`,
        ui.ButtonSet.YES_NO
      );
      if (retry === ui.Button.NO) {
        ui.alert("Ação cancelada.");
        return;
      } else {
        continue; // repete o loop para pedir outro ID
      }
    }

    // ✅ Confirma reenvio
    const confirm = ui.alert(
      "Confirmação",
      `Deseja realmente reenviar a solicitação de ID ${idNum} para Rober?`,
      ui.ButtonSet.YES_NO
    );

    if (confirm !== ui.Button.YES) {
      ui.alert("Envio cancelado.");
      return;
    }

    const ok = reenviarSolicitacaoPorId_(idNum);
    if (ok) {
      ui.alert(`✅ Solicitação ${idNum} reenviada com sucesso para Rober.`);
    } else {
      ui.alert(`⚠️ Não foi possível reenviar a solicitação ${idNum}. Verifique se o ID existe ou se já foi aprovado/reprovado.`);
    }
    break; // encerra o loop após envio
  }
}

function reenviarSolicitacaoPorId_(id) {
  const TZ = "America/Sao_Paulo";
  const WEBAPP_URL = "https://script.google.com/macros/s/AKfycbznXr_UYCcq4-6rmGsxCxXJs7Df0ABzA3EkBgajbXC42NvH6c0AQdUDc6WZ15XV_3JW/exec";
  const SHEET_URL = "https://docs.google.com/spreadsheets/d/1Pd7jw-1OkUYKvPCw5Z20ckoeDRe211UMn4w4Bud0CKY/edit";
  const SHEET_NAME = "Respostas ao formulário 1";
  const APROVADOR_EMAIL = "rober@grupoorion.com.br";
  const emailsCopia = "lucas.garcia@grupoorion.com.br, cilene.silva@grupoorion.com.br";

  const ss = SpreadsheetApp.openByUrl(SHEET_URL);
  const sh = ss.getSheetByName(SHEET_NAME);
  const logoUrl = "https://drive.google.com/uc?export=download&id=1RqQxKlSmnBnDPbg2HoLkV6QbRVgXr0il";
  const logoBlob = UrlFetchApp.fetch(logoUrl).getBlob().setName("logo.png");

  // Procura a linha do ID
  const lastRow = sh.getLastRow();
  const ids = sh.getRange(2, 11, lastRow - 1).getValues().flat();
  const linha = ids.indexOf(id) + 2;
  if (linha < 2) return false;

  const r = sh.getRange(linha, 1, 1, sh.getLastColumn()).getValues()[0];
  const carimbo = r[0];
  const nome = r[2];
  const app = r[3];
  const desc = r[4];
  const obs = r[5];
  const status = (r[6] || "").toString().trim();
  const idValue = r[10];

  if (status.toLowerCase() === "aprovado" || status.toLowerCase() === "reprovado") return false;

  const linkAprovar = `${WEBAPP_URL}?acao=aprovar&linha=${linha}`;
  const linkReprovar = `${WEBAPP_URL}?acao=reprovar&linha=${linha}`;
  const assunto = `Reenvio: Solicitação de Aprovação - ${app} (${idValue})`;
  const dataFmt = Utilities.formatDate(new Date(carimbo), TZ, "dd/MM/yyyy HH:mm");

  const btn = (href, label, bg) => `
    <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="display:inline-block;margin-right:10px;">
      <tr>
        <td bgcolor="${bg}" style="border-radius:6px;">
          <a href="${href}" target="_blank"
             style="font-family:Arial,sans-serif;font-size:14px;line-height:14px;text-decoration:none;
                    color:#ffffff;padding:12px 18px;display:inline-block;">${label}</a>
        </td>
      </tr>
    </table>`;

  const descFmt = (desc || "-").replace(/\n/g, "<br>");
  const obsFmt  = (obs  || "-").replace(/\n/g, "<br>");

  const corpoHtml = `
  <div style="font-family:Arial,Helvetica,sans-serif;color:#333;max-width:760px;">
    <div style="margin:12px 0 20px;">
      <img src="cid:logo" alt="Orion" style="height:40px;display:block;">
    </div>

    <h2 style="color:#2E86C1;margin:0 0 16px;">Solicitação de Aprovação - ${app} (${idValue})</h2>

    <p>Olá Rober, tudo bem?</p>
    <p>Este é um <b>reenvio manual</b> de uma solicitação de aprovação pendente.</p>

    <ul style="margin:0 0 16px;padding-left:20px;">
      <li><b>ID:</b> ${idValue}</li>
      <li><b>Data/Hora:</b> ${dataFmt}</li>
      <li><b>Solicitante:</b> ${nome || "-"}</li>
      <li><b>Aplicativo:</b> ${app || "-"}</li>
    </ul>

    <p><b>Descrição:</b><br>${descFmt}</p>
    <p><b>Observação:</b><br>${obsFmt}</p>

    <div style="margin:6px 0 16px;">
      ${btn(linkAprovar, "Aprovar",  "#28a745")}
      ${btn(linkReprovar, "Reprovar", "#dc3545")}
    </div>

    <p style="margin:20px 0 0;">
      <a href="${SHEET_URL}" style="color:#1a73e8;text-decoration:none;">Acessar a planilha</a>
    </p>

    <p style="margin-top:16px;font-size:13px;color:#555;">
      Caso apareça alguma mensagem de erro ao clicar em <b>“Aprovar”</b> ou <b>“Reprovar”</b>,
      a decisão ainda será gravada corretamente na planilha.
    </p>

    <p style="margin-top:24px;">Atenciosamente,<br>Equipe de Qualidade Orion.</p>
  </div>`;

  GmailApp.sendEmail(
    APROVADOR_EMAIL,
    assunto,
    " ",
    {
      htmlBody: corpoHtml,
      inlineImages: { logo: logoBlob },
      cc: emailsCopia
    }
  );

  const hoje = Utilities.formatDate(new Date(), TZ, "dd/MM/yyyy HH:mm");
  sh.getRange(linha, 8).setValue(`Reenviado manualmente em ${hoje}`);

  return true;
}

function enviarEmailConfirmacaoAprovacao_(dados) {
  // ✅ Apenas Lucas receberá durante a fase de validação
  const EMAIL_CILENE = "cilene.silva@grupoorion.com.br";
  const EMAIL_LUCAS  = "lucas.garcia@grupoorion.com.br";



  const LOGO_URL = "https://lh3.googleusercontent.com/d/1B8bS5fljAR-OdHEVQjE4Eb-Wpa9wndty=w300";

  const {
  app,
  idValue,
  nome,
  statusTxt,
  dataHora,
  sheetUrl,
  descricao,
  comentarioGestor   // ✅ NOVO
} = dados;

    const comentarioFinal = comentarioGestor ? comentarioGestor : "";

    const comentarioHtml = comentarioFinal
  ? comentarioFinal
  : `<span style="color:#9ca3af;">
       Nenhum comentário adicional foi informado.
     </span>`;



  const aprovado = statusTxt === "Aprovado";

  const titulo = aprovado ? "Solicitação Aprovada" : "Solicitação Reprovada";
  const corTitulo = aprovado ? "#16a34a" : "#dc2626";

  const textoPrincipal = aprovado
    ? `Sua solicitação <b>(ID ${idValue})</b> acabou de ser aprovada.`
    : `Sua solicitação <b>(ID ${idValue})</b> infelizmente foi reprovada.`;

  const assunto = `${titulo} – ${app} (ID ${idValue})`;

  const corpoHtml = `
<table width="100%" cellpadding="0" cellspacing="0"
       style="background:#f3f4f6;padding:20px;font-family:Arial,sans-serif;">
  <tr>
    <td align="center">
      <table width="600" cellpadding="0" cellspacing="0"
             style="background:#ffffff;border-radius:8px;padding:24px;">

        <!-- LOGO -->
        <tr>
          <td align="center" style="padding-bottom:20px;">
            <img src="${LOGO_URL}" alt="Orion" style="height:48px;display:block;">
          </td>
        </tr>

        <!-- TÍTULO -->
        <tr>
          <td style="font-size:20px;color:${corTitulo};font-weight:bold;padding-bottom:16px;">
            ${titulo}
          </td>
        </tr>

        <!-- TEXTO PRINCIPAL -->
        <tr>
          <td style="font-size:14px;color:#333;padding-bottom:20px;">
            Olá, <b>${nome || ""}</b>!<br><br>
            ${textoPrincipal}
          </td>
        </tr>

        <!-- DADOS DA SOLICITAÇÃO -->
        <tr>
          <td>
            <table width="100%" cellpadding="6" cellspacing="0"
                   style="font-size:14px;color:#333;">
              <tr>
                <td width="160"><b>ID:</b></td>
                <td>${idValue}</td>
              </tr>
              <tr>
                <td><b>Aplicativo:</b></td>
                <td>${app}</td>
              </tr>
              <tr>
                <td style="vertical-align:top;"><b>Descrição:</b></td>
                <td>${descricao || "-"}</td>
              </tr>
              <tr>
                <td><b>Solicitante:</b></td>
                <td>${nome || "-"}</td>
              </tr>
              <tr>
                <td><b>Data / Hora:</b></td>
                <td>${dataHora}</td>
              </tr>

              <!-- DIVISÓRIA -->
              <tr>
                <td colspan="2" style="padding:12px 0;">
                  <hr style="border:none;border-top:1px solid #e5e7eb;">
                </td>
              </tr>

              <!-- COMENTÁRIO DO GESTOR -->
              <tr>
                <td style="vertical-align:top;">
                  <b>Comentário do Gestor:</b>
                </td>
                <td style="color:#6b7280;">
                  ${comentarioHtml}
                </td>
              </tr>


            </table>
          </td>
        </tr>

        <!-- BOTÃO -->
        <tr>
          <td align="center" style="padding:24px 0;">
            <a href="${sheetUrl}"
               style="background:#2563eb;color:#ffffff;text-decoration:none;
                      padding:12px 24px;border-radius:999px;
                      font-size:14px;display:inline-block;">
              Acessar planilha
            </a>
          </td>
        </tr>

        <!-- RODAPÉ -->
        <tr>
          <td style="font-size:12px;color:#6b7280;text-align:center;
                     border-top:1px solid #e5e7eb;
                     padding-top:12px;">
            Este e-mail é enviado automaticamente pelo Setor de Qualidade.<br>
            Por favor, não responda.
          </td>
        </tr>

      </table>
    </td>
  </tr>
</table>
`;


  GmailApp.sendEmail(
  EMAIL_CILENE,
  assunto,
  " ",
  {
    htmlBody: corpoHtml,
    cc: EMAIL_LUCAS
  }
);
}

