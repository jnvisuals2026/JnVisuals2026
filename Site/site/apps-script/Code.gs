const OWNER_EMAIL = "jnvisuals2026@gmail.com";
const SHEET_NAME = "Adesoes";

const BRAND = {
  navy: "#0b1730",
  blue: "#1d4ed8",
  lightBlue: "#eff6ff",
  green: "#15803d",
  lightGreen: "#f0fdf4",
  red: "#b91c1c",
  lightRed: "#fef2f2",
  gray: "#64748b",
  lightGray: "#f8fafc",
  border: "#e2e8f0",
  white: "#ffffff"
};

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents || "{}");

    const required = ["nome", "email", "atleta", "imagem", "assinatura"];
    required.forEach(k => {
      if (!d[k]) throw new Error("Campo obrigatório em falta: " + k);
    });

    const ss =
      SpreadsheetApp.getActiveSpreadsheet() ||
      SpreadsheetApp.create("JN Visuals — Adesões");

    let sh = ss.getSheetByName(SHEET_NAME);

    if (!sh) sh = ss.insertSheet(SHEET_NAME);

    if (sh.getLastRow() === 0) {
      sh.appendRow([
        "Data",
        "ID",
        "Nome",
        "Email",
        "Telefone",
        "Localidade",
        "Atleta",
        "Clube",
        "Escalão",
        "Número/posição",
        "Imagem",
        "Observações",
        "Assinatura",
        "Estado",
        "Token",
        "Data decisão",
        "PDF URL",
        "PDF File ID"
      ]);
    }

    const id = "JNV-" + Utilities.getUuid().slice(0, 8).toUpperCase();
    const token = Utilities.getUuid();
    const now = new Date();

    const row = sh.getLastRow() + 1;

    sh.getRange(row, 1, 1, 18).setValues([[
      now,
      id,
      d.nome,
      d.email,
      d.telefone || "",
      d.localidade || "",
      d.atleta,
      d.clube || "",
      d.escalao || "",
      d.numero || "",
      d.imagem,
      d.observacoes || "",
      d.assinatura,
      "PENDENTE",
      token,
      "",
      "",
      ""
    ]]);

    const base = ScriptApp.getService().getUrl();

    const approve =
      base +
      "?action=approve&token=" +
      encodeURIComponent(token);

    const reject =
      base +
      "?action=reject&token=" +
      encodeURIComponent(token);

    // EMAIL PARA O RESPONSÁVEL JN VISUALS
    MailApp.sendEmail({
      to: OWNER_EMAIL,
      subject: "JN Visuals | Nova adesão — " + id,
      htmlBody: ownerNewAdhesionEmail_(d, id, now, approve, reject)
    });

    // EMAIL PARA O CLIENTE
    MailApp.sendEmail({
      to: d.email,
      subject: "JN Visuals | Pedido recebido — " + id,
      htmlBody: clientPendingEmail_(d, id, now)
    });

    return json_({
      ok: true,
      id: id
    });

  } catch (err) {
    return json_({
      ok: false,
      error: String(err.message || err)
    });
  }
}


function doGet(e) {
  const action = e.parameter.action;
  const token = e.parameter.token;

  if (!action || !token) {
    return pageMessage_(
      "Link inválido",
      "O link utilizado não é válido ou está incompleto.",
      "error"
    );
  }

  const sh =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);

  if (!sh) {
    return pageMessage_(
      "Erro",
      "A folha de adesões não foi encontrada.",
      "error"
    );
  }

  const vals = sh.getDataRange().getValues();

  let r = -1;
  let data = null;

  for (let i = 1; i < vals.length; i++) {
    if (vals[i][14] === token) {
      r = i + 1;
      data = vals[i];
      break;
    }
  }

  if (r < 0) {
    return pageMessage_(
      "Pedido não encontrado",
      "Não foi possível localizar esta adesão.",
      "error"
    );
  }

  if (data[13] !== "PENDENTE") {
    return pageMessage_(
      "Pedido já processado",
      "Este pedido já foi processado anteriormente.",
      "info"
    );
  }


  // =========================
  // RECUSAR
  // =========================

  if (action === "reject") {

    sh.getRange(r, 14, 1, 2)
      .setValues([["RECUSADO", token]]);

    MailApp.sendEmail({
      to: data[3],
      subject: "JN Visuals | Adesão recusada — " + data[1],
      htmlBody: clientRejectedEmail_(data)
    });

    return pageMessage_(
      "Pedido recusado",
      "A adesão foi recusada e o responsável foi notificado por email.",
      "error"
    );
  }


  // =========================
  // APROVAR
  // =========================

  if (action === "approve") {

    const pdf = makePdf_(data);

    sh.getRange(r, 14, 1, 4)
      .setValues([[
        "APROVADO",
        token,
        new Date(),
        pdf.url
      ]]);

    sh.getRange(r, 18)
      .setValue(pdf.id);


    // EMAIL PARA O CLIENTE
    MailApp.sendEmail({
      to: data[3],
      subject: "JN Visuals | Adesão aprovada — " + data[1],
      htmlBody: clientApprovedEmail_(data, pdf),
      attachments: [pdf.blob]
    });


    // EMAIL PARA O RESPONSÁVEL
    MailApp.sendEmail({
      to: OWNER_EMAIL,
      subject: "JN Visuals | Contrato aprovado — " + data[1],
      htmlBody: ownerApprovedEmail_(data, pdf),
      attachments: [pdf.blob]
    });


    return pageMessage_(
      "Adesão aprovada",
      "A adesão foi aprovada. O contrato PDF foi gerado, arquivado e enviado por email.",
      "success"
    );
  }


  return pageMessage_(
    "Ação inválida",
    "A ação solicitada não é reconhecida.",
    "error"
  );
}


// =====================================================
// EMAIL — NOVA ADESÃO PARA O RESPONSÁVEL
// =====================================================

function ownerNewAdhesionEmail_(d, id, date, approve, reject) {

  return emailShell_(
    "NOVA ADESÃO",
    "Existe uma nova adesão a aguardar a tua decisão.",
    `
      ${statusBadge_("PENDENTE", BRAND.blue)}

      ${card_(
        "Dados da adesão",
        `
          ${infoRow_("ID", id)}
          ${infoRow_("Responsável", d.nome)}
          ${infoRow_("Email", d.email)}
          ${infoRow_("Telefone", d.telefone || "—")}
          ${infoRow_("Localidade", d.localidade || "—")}
          ${infoRow_("Atleta", d.atleta)}
          ${infoRow_("Clube / Equipa", d.clube || "—")}
          ${infoRow_("Escalão", d.escalao || "—")}
          ${infoRow_("Número / Posição", d.numero || "—")}
          ${infoRow_("Direito de imagem", d.imagem)}
          ${infoRow_("Data", formatDate_(date))}
        `
      )}

      ${d.observacoes ? card_(
        "Observações",
        `<p style="margin:0;color:${BRAND.gray};line-height:1.7;">
          ${esc_(d.observacoes)}
        </p>`
      ) : ""}

      <div style="margin:28px 0;">
        <p style="font-size:13px;color:${BRAND.gray};margin:0 0 14px;">
          DECISÃO NECESSÁRIA
        </p>

        <a href="${approve}"
           style="
             display:inline-block;
             background:${BRAND.green};
             color:#fff;
             text-decoration:none;
             padding:15px 28px;
             border-radius:10px;
             font-weight:700;
             margin-right:8px;
           ">
          APROVAR ADESÃO
        </a>

        <a href="${reject}"
           style="
             display:inline-block;
             background:${BRAND.red};
             color:#fff;
             text-decoration:none;
             padding:15px 28px;
             border-radius:10px;
             font-weight:700;
           ">
          RECUSAR
        </a>
      </div>

      <div style="
        background:${BRAND.lightBlue};
        border:1px solid #dbeafe;
        border-radius:12px;
        padding:16px;
        font-size:13px;
        color:#334155;
      ">
        <strong>Nota:</strong>
        ao aprovar, o sistema gera automaticamente o contrato PDF,
        arquiva-o no Google Drive e envia-o por email.
      </div>
    `
  );
}


// =====================================================
// EMAIL — CLIENTE / PEDIDO PENDENTE
// =====================================================

function clientPendingEmail_(d, id, date) {

  return emailShell_(
    "PEDIDO RECEBIDO",
    "Recebemos a tua adesão ao Plano Anual JN Visuals.",
    `
      ${statusBadge_("EM ANÁLISE", BRAND.blue)}

      <p style="
        color:#334155;
        line-height:1.8;
        font-size:15px;
      ">
        Olá <strong>${esc_(d.nome)}</strong>,
      </p>

      <p style="
        color:#475569;
        line-height:1.8;
        font-size:15px;
      ">
        O teu pedido foi recebido com sucesso e encontra-se
        atualmente pendente de aprovação.
      </p>

      ${card_(
        "Resumo do pedido",
        `
          ${infoRow_("ID da adesão", id)}
          ${infoRow_("Atleta", d.atleta)}
          ${infoRow_("Clube / Equipa", d.clube || "—")}
          ${infoRow_("Plano", "Plano Anual JN Visuals")}
          ${infoRow_("Valor", "20 € / ano")}
          ${infoRow_("Data", formatDate_(date))}
        `
      )}

      <div style="
        background:${BRAND.lightBlue};
        border:1px solid #dbeafe;
        border-radius:12px;
        padding:18px;
        margin-top:22px;
        color:#334155;
        line-height:1.7;
        font-size:14px;
      ">
        A assinatura eletrónica submetida através do formulário
        foi registada juntamente com os restantes dados da adesão.
      </div>

      <p style="
        color:${BRAND.gray};
        font-size:13px;
        line-height:1.7;
        margin-top:24px;
      ">
        Quando a adesão for aprovada, receberás um novo email
        com o contrato PDF final em anexo.
      </p>
    `
  );
}


// =====================================================
// EMAIL — CLIENTE / APROVADO
// =====================================================

function clientApprovedEmail_(data, pdf) {

  return emailShell_(
    "ADESÃO APROVADA",
    "O teu Plano Anual JN Visuals está oficialmente aprovado.",
    `
      <div style="
        border:1px solid #bbf7d0;
        background:${BRAND.lightGreen};
        border-radius:14px;
        padding:22px;
        margin-bottom:24px;
      ">
        <div style="
          font-size:12px;
          font-weight:800;
          letter-spacing:1.5px;
          color:${BRAND.green};
          margin-bottom:8px;
        ">
          DOCUMENTO OFICIAL
        </div>

        <div style="
          font-size:24px;
          font-weight:800;
          color:#166534;
        ">
          ADESÃO APROVADA
        </div>

        <div style="
          margin-top:8px;
          color:#166534;
          font-size:14px;
        ">
          ID: ${esc_(data[1])}
        </div>
      </div>

      <p style="
        color:#334155;
        line-height:1.8;
        font-size:15px;
      ">
        Olá <strong>${esc_(data[2])}</strong>,
      </p>

      <p style="
        color:#475569;
        line-height:1.8;
        font-size:15px;
      ">
        Informamos que a tua adesão ao Plano Anual de Fotografia
        Desportiva JN Visuals foi aprovada.
      </p>

      ${card_(
        "Dados contratuais",
        `
          ${infoRow_("ID", data[1])}
          ${infoRow_("Responsável", data[2])}
          ${infoRow_("Atleta", data[6])}
          ${infoRow_("Clube / Equipa", data[7] || "—")}
          ${infoRow_("Escalão", data[8] || "—")}
          ${infoRow_("Plano", "Anual")}
          ${infoRow_("Valor", "20 € / ano")}
          ${infoRow_("Estado", "APROVADO")}
        `
      )}

      <div style="text-align:center;margin:30px 0;">
        <a href="${pdf.url}"
           style="
             display:inline-block;
             background:${BRAND.navy};
             color:#fff;
             text-decoration:none;
             padding:15px 28px;
             border-radius:10px;
             font-weight:700;
           ">
          ABRIR CONTRATO PDF
        </a>
      </div>

      <div style="
        border-top:1px solid ${BRAND.border};
        padding-top:20px;
        color:${BRAND.gray};
        font-size:13px;
        line-height:1.7;
      ">
        O contrato PDF final segue também anexado a este email.
        Guarda este documento para os teus registos.
      </div>
    `
  );
}


// =====================================================
// EMAIL — RESPONSÁVEL / CONTRATO APROVADO
// =====================================================

function ownerApprovedEmail_(data, pdf) {

  return emailShell_(
    "CONTRATO APROVADO",
    "A adesão foi aprovada e o contrato foi gerado automaticamente.",
    `
      ${statusBadge_("APROVADO", BRAND.green)}

      ${card_(
        "Contrato",
        `
          ${infoRow_("ID", data[1])}
          ${infoRow_("Responsável", data[2])}
          ${infoRow_("Email", data[3])}
          ${infoRow_("Atleta", data[6])}
          ${infoRow_("Clube / Equipa", data[7] || "—")}
          ${infoRow_("Escalão", data[8] || "—")}
          ${infoRow_("Estado", "APROVADO")}
          ${infoRow_("PDF", "Gerado e arquivado")}
        `
      )}

      <div style="text-align:center;margin:30px 0;">
        <a href="${pdf.url}"
           style="
             display:inline-block;
             background:${BRAND.navy};
             color:#fff;
             text-decoration:none;
             padding:15px 30px;
             border-radius:10px;
             font-weight:700;
           ">
          ABRIR PDF NO GOOGLE DRIVE
        </a>
      </div>

      <div style="
        background:${BRAND.lightBlue};
        border:1px solid #dbeafe;
        border-radius:12px;
        padding:18px;
        color:#334155;
        font-size:13px;
        line-height:1.7;
      ">
        O documento foi criado automaticamente, guardado no Google Drive
        e enviado ao cliente como anexo.
      </div>
    `
  );
}


// =====================================================
// EMAIL — CLIENTE / RECUSADO
// =====================================================

function clientRejectedEmail_(data) {

  return emailShell_(
    "ADESÃO NÃO APROVADA",
    "Informação sobre o estado da tua adesão.",
    `
      ${statusBadge_("NÃO APROVADA", BRAND.red)}

      <p style="
        color:#334155;
        line-height:1.8;
        font-size:15px;
      ">
        Olá <strong>${esc_(data[2])}</strong>,
      </p>

      <p style="
        color:#475569;
        line-height:1.8;
        font-size:15px;
      ">
        Informamos que o pedido de adesão
        <strong>${esc_(data[1])}</strong>
        não foi aprovado.
      </p>

      ${card_(
        "Resumo",
        `
          ${infoRow_("ID", data[1])}
          ${infoRow_("Atleta", data[6])}
          ${infoRow_("Clube / Equipa", data[7] || "—")}
          ${infoRow_("Estado", "NÃO APROVADA")}
        `
      )}

      <p style="
        color:${BRAND.gray};
        font-size:13px;
        line-height:1.7;
        margin-top:24px;
      ">
        Caso necessites de algum esclarecimento relativamente ao pedido,
        poderás contactar a JN Visuals através dos canais disponibilizados.
      </p>
    `
  );
}


// =====================================================
// ESTRUTURA BASE DOS EMAILS
// =====================================================

function emailShell_(eyebrow, title, content) {

  return `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width,initial-scale=1">
  </head>

  <body style="
    margin:0;
    padding:0;
    background:#f1f5f9;
    font-family:Arial,Helvetica,sans-serif;
  ">

    <div style="
      max-width:680px;
      margin:0 auto;
      padding:35px 16px;
    ">

      <div style="
        background:${BRAND.navy};
        border-radius:18px 18px 0 0;
        padding:30px;
        color:#fff;
      ">

        <div style="
          font-size:12px;
          letter-spacing:2px;
          font-weight:800;
          color:#93c5fd;
          margin-bottom:10px;
        ">
          JN VISUALS
        </div>

        <div style="
          font-size:28px;
          line-height:1.2;
          font-weight:800;
        ">
          ${title}
        </div>

        <div style="
          margin-top:10px;
          font-size:14px;
          line-height:1.6;
          color:#cbd5e1;
        ">
          ${eyebrow}
        </div>

      </div>


      <div style="
        background:#ffffff;
        padding:30px;
        border-left:1px solid ${BRAND.border};
        border-right:1px solid ${BRAND.border};
      ">

        ${content}

      </div>


      <div style="
        background:#ffffff;
        border:1px solid ${BRAND.border};
        border-top:0;
        border-radius:0 0 18px 18px;
        padding:22px 30px;
        text-align:center;
      ">

        <div style="
          font-size:14px;
          font-weight:800;
          color:${BRAND.navy};
        ">
          JN VISUALS
        </div>

        <div style="
          margin-top:6px;
          font-size:12px;
          color:${BRAND.gray};
        ">
          Grandes histórias começam com um click
        </div>

        <div style="
          margin-top:14px;
          font-size:11px;
          color:#94a3b8;
        ">
          Este email foi gerado automaticamente pelo sistema JN Visuals.
        </div>

      </div>

    </div>

  </body>
  </html>
  `;
}


// =====================================================
// COMPONENTES VISUAIS
// =====================================================

function card_(title, content) {

  return `
    <div style="
      border:1px solid ${BRAND.border};
      border-radius:14px;
      overflow:hidden;
      margin:20px 0;
      background:#fff;
    ">

      <div style="
        background:${BRAND.lightGray};
        border-bottom:1px solid ${BRAND.border};
        padding:14px 18px;
        font-size:13px;
        font-weight:800;
        color:${BRAND.navy};
        letter-spacing:.3px;
      ">
        ${title}
      </div>

      <div style="
        padding:18px;
      ">
        ${content}
      </div>

    </div>
  `;
}


function infoRow_(label, value) {

  return `
    <div style="
      padding:9px 0;
      border-bottom:1px solid #f1f5f9;
    ">

      <span style="
        display:inline-block;
        width:145px;
        vertical-align:top;
        color:${BRAND.gray};
        font-size:13px;
      ">
        ${label}
      </span>

      <strong style="
        color:#1e293b;
        font-size:13px;
      ">
        ${esc_(value)}
      </strong>

    </div>
  `;
}


function statusBadge_(text, color) {

  return `
    <div style="margin-bottom:20px;">

      <span style="
        display:inline-block;
        background:${color};
        color:#fff;
        padding:8px 13px;
        border-radius:999px;
        font-size:11px;
        font-weight:800;
        letter-spacing:.8px;
      ">
        ${text}
      </span>

    </div>
  `;
}


// =====================================================
// PÁGINAS DE RESULTADO
// =====================================================

function pageMessage_(title, message, type) {

  let accent = BRAND.blue;

  if (type === "success") accent = BRAND.green;
  if (type === "error") accent = BRAND.red;

  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <meta name="viewport" content="width=device-width,initial-scale=1">
      <style>
        body{
          margin:0;
          background:#f1f5f9;
          font-family:Arial,Helvetica,sans-serif;
        }

        .wrap{
          max-width:620px;
          margin:80px auto;
          padding:20px;
        }

        .box{
          background:#fff;
          border-radius:18px;
          padding:40px;
          text-align:center;
          box-shadow:0 10px 40px rgba(15,23,42,.08);
        }

        .brand{
          color:${BRAND.navy};
          font-size:13px;
          font-weight:800;
          letter-spacing:2px;
          margin-bottom:25px;
        }

        .line{
          width:55px;
          height:4px;
          background:${accent};
          border-radius:5px;
          margin:0 auto 25px;
        }

        h2{
          margin:0 0 12px;
          color:${BRAND.navy};
          font-size:28px;
        }

        p{
          color:${BRAND.gray};
          line-height:1.7;
          font-size:15px;
        }
      </style>
    </head>

    <body>

      <div class="wrap">

        <div class="box">

          <div class="brand">
            JN VISUALS
          </div>

          <div class="line"></div>

          <h2>${esc_(title)}</h2>

          <p>${esc_(message)}</p>

        </div>

      </div>

    </body>
    </html>
  `);
}


// =====================================================
// PDF
// =====================================================

function makePdf_(d) {

  const doc =
    DocumentApp.create(
      "JN Visuals — Contrato " + d[1]
    );

  const body = doc.getBody();

  body.clear();

  body
    .appendParagraph("JN VISUALS")
    .setHeading(
      DocumentApp.ParagraphHeading.TITLE
    );

  body
    .appendParagraph(
      "CONTRATO / ADESÃO — PLANO ANUAL 20 €"
    )
    .setHeading(
      DocumentApp.ParagraphHeading.HEADING1
    );

  body.appendParagraph(
    "ID: " + d[1] + " · Estado: APROVADO"
  );


  body
    .appendParagraph("1. IDENTIFICAÇÃO")
    .setHeading(
      DocumentApp.ParagraphHeading.HEADING2
    );

  body.appendParagraph(
    "Responsável: " + d[2] +
    "\nEmail: " + d[3] +
    "\nTelefone: " + d[4] +
    "\nLocalidade: " + d[5] +
    "\nAtleta: " + d[6] +
    "\nClube/Equipa: " + d[7] +
    "\nEscalão: " + d[8] +
    "\nNúmero/posição: " + d[9]
  );


  body
    .appendParagraph("2. OBJETO E PREÇO")
    .setHeading(
      DocumentApp.ParagraphHeading.HEADING2
    );

  body.appendParagraph(
    "Adesão ao Plano Anual de Fotografia Desportiva JN Visuals pelo valor de 20 € por ano."
  );


  body
    .appendParagraph("3. DISPONIBILIDADE E COBERTURA")
    .setHeading(
      DocumentApp.ParagraphHeading.HEADING2
    );

  body.appendParagraph(
    "A cobertura depende da disponibilidade do fotógrafo, das condições do local e da autorização de acesso. Não existe garantia de cobertura de todos os jogos, treinos ou eventos nem de um número mínimo de fotografias."
  );


  body
    .appendParagraph("4. FOTOGRAFIAS E DIREITOS")
    .setHeading(
      DocumentApp.ParagraphHeading.HEADING2
    );

  body.appendParagraph(
    "As fotografias são selecionadas e editadas segundo critérios técnicos e criativos. A adesão não transfere os direitos de autor para o cliente."
  );


  body
    .appendParagraph("5. DIREITO DE IMAGEM")
    .setHeading(
      DocumentApp.ParagraphHeading.HEADING2
    );

  body.appendParagraph(
    "Opção registada: " + d[10] +
    ". A autorização, quando escolhida, abrange publicação em Instagram, Facebook, TikTok, website e portefólio para divulgação do trabalho."
  );


  body
    .appendParagraph("6. REEMBOLSO E LIVRE RESOLUÇÃO")
    .setHeading(
      DocumentApp.ParagraphHeading.HEADING2
    );

  body.appendParagraph(
    "É possível solicitar reembolso nos termos do plano, sem prejuízo dos direitos legais. Nos contratos à distância, quando aplicável, é observado o regime legal de livre resolução."
  );


  body
    .appendParagraph("7. PRIVACIDADE")
    .setHeading(
      DocumentApp.ParagraphHeading.HEADING2
    );

  body.appendParagraph(
    "Os dados são tratados para gestão da adesão, comunicação, execução do serviço, faturação quando aplicável e cumprimento de obrigações legais."
  );


  body
    .appendParagraph("8. ASSINATURA ELETRÓNICA")
    .setHeading(
      DocumentApp.ParagraphHeading.HEADING2
    );

  body.appendParagraph(
    "Assinatura submetida eletronicamente pelo responsável através do formulário JN Visuals. A assinatura é um registo eletrónico de aceitação e não é apresentada como assinatura digital qualificada."
  );


  try {

    const b64 =
      d[12].split(",")[1];

    const bytes =
      Utilities.base64Decode(b64);

    const img =
      body.appendImage(
        Utilities.newBlob(
          bytes,
          "image/png",
          "assinatura.png"
        )
      );

    img.setWidth(250);

  } catch (err) {}


  body.appendParagraph(
    "Data/hora da submissão: " + d[0]
  );


  doc.saveAndClose();


  const pdf =
    DriveApp.createFile(
      DriveApp
        .getFileById(doc.getId())
        .getAs(MimeType.PDF)
    );

  DriveApp
    .getFileById(doc.getId())
    .setTrashed(true);


  return {
    id: pdf.getId(),
    url: pdf.getUrl(),
    blob: pdf.getBlob()
  };
}


// =====================================================
// SEGURANÇA / FORMATAÇÃO
// =====================================================

function esc_(value) {

  if (value === null || value === undefined) {
    return "";
  }

  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}


function formatDate_(date) {

  try {

    return Utilities.formatDate(
      new Date(date),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy HH:mm"
    );

  } catch (err) {

    return String(date);

  }
}


function json_(o) {

  return ContentService
    .createTextOutput(
      JSON.stringify(o)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}