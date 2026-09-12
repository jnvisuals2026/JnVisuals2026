const OWNER_EMAIL = "jnvisuals2026@gmail.com";
const SHEET_NAME = "Adesoes";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || "{}");
    const required = ["nome", "email", "atleta", "imagem"];
    required.forEach(k => {
      if (!data[k]) throw new Error("Campo obrigatório em falta: " + k);
    });

    const id = "JNV-" + Utilities.getUuid().slice(0, 8).toUpperCase();
    const ss = getSheet_();
    ss.appendRow([
      new Date(), id, data.nome || "", data.email || "", data.telefone || "",
      data.atleta || "", data.clube || "", data.escalao || "",
      data.imagem || "", data.observacoes || "", "PENDENTE"
    ]);

    const scriptUrl = ScriptApp.getService().getUrl();
    const approve = scriptUrl + "?action=approve&id=" + encodeURIComponent(id);
    const reject = scriptUrl + "?action=reject&id=" + encodeURIComponent(id);

    MailApp.sendEmail({
      to: OWNER_EMAIL,
      subject: "JN Visuals — Nova adesão " + id,
      htmlBody:
        "<h2>Nova adesão JN Visuals</h2>" +
        "<p><b>ID:</b> " + esc_(id) + "</p>" +
        "<p><b>Nome:</b> " + esc_(data.nome) + "</p>" +
        "<p><b>Email:</b> " + esc_(data.email) + "</p>" +
        "<p><b>Atleta:</b> " + esc_(data.atleta) + "</p>" +
        "<p><b>Clube:</b> " + esc_(data.clube || "") + "</p>" +
        "<p><b>Imagem:</b> " + esc_(data.imagem) + "</p>" +
        "<p><a href='" + approve + "'>APROVAR ADESÃO</a></p>" +
        "<p><a href='" + reject + "'>RECUSAR ADESÃO</a></p>"
    });

    MailApp.sendEmail({
      to: data.email,
      subject: "JN Visuals — Recebemos a tua adesão",
      htmlBody:
        "<p>Olá " + esc_(data.nome) + ",</p>" +
        "<p>Recebemos a tua adesão ao Plano Anual JN Visuals de 20 €.</p>" +
        "<p><b>ID do pedido:</b> " + esc_(id) + "</p>" +
        "<p>O pedido está pendente de aprovação. Receberás uma nova mensagem quando houver uma decisão.</p>"
    });

    return json_({ok:true, id:id});
  } catch(err) {
    return json_({ok:false, error:String(err.message || err)});
  }
}

function doGet(e) {
  const action = e.parameter.action;
  const id = e.parameter.id;
  if (!action || !id) return HtmlService.createHtmlOutput("JN Visuals — pedido inválido.");

  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][1] === id) {
      const row = i + 1;
      const data = {
        id:id, nome:values[i][2], email:values[i][3], telefone:values[i][4],
        atleta:values[i][5], clube:values[i][6], escalao:values[i][7],
        imagem:values[i][8], observacoes:values[i][9], estado:values[i][10]
      };
      if (action === "approve") {
        sheet.getRange(row, 11).setValue("APROVADO");
        MailApp.sendEmail({
          to:data.email,
          subject:"JN Visuals — Adesão aprovada — " + id,
          htmlBody:"<p>Olá " + esc_(data.nome) + ",</p><p>A tua adesão ao Plano Anual JN Visuals foi aprovada.</p><p><b>ID:</b> " + esc_(id) + "</p><p>O serviço tem o valor de 20 € por ano. A cobertura depende da disponibilidade e do acesso/permissão para fotografar.</p>"
        });
        return HtmlService.createHtmlOutput("Adesão " + esc_(id) + " aprovada. O cliente foi notificado.");
      }
      if (action === "reject") {
        sheet.getRange(row, 11).setValue("RECUSADO");
        MailApp.sendEmail({
          to:data.email,
          subject:"JN Visuals — Decisão sobre a adesão — " + id,
          htmlBody:"<p>Olá " + esc_(data.nome) + ",</p><p>A adesão com o ID " + esc_(id) + " foi recusada.</p><p>Se tiveres alguma questão, contacta a JN Visuals.</p>"
        });
        return HtmlService.createHtmlOutput("Adesão " + esc_(id) + " recusada. O cliente foi notificado.");
      }
    }
  }
  return HtmlService.createHtmlOutput("Pedido não encontrado.");
}

function getSheet_() {
  const files = DriveApp.getFilesByName("Adesoes");
  let ss;
  if (files.hasNext()) ss = SpreadsheetApp.open(files.next());
  else ss = SpreadsheetApp.create("Adesoes");
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(["Data","ID","Nome","Email","Telefone","Atleta","Clube","Escalão","Imagem","Observações","Estado"]);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function esc_(v) {
  return String(v ?? "").replace(/[&<>"']/g, s => ({
    "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"
  }[s]));
}
