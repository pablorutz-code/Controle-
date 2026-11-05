var SENHA = "TROQUE_ESTA_SENHA";

function doPost(e) {
  var r;
  try {
    var b = JSON.parse(e.postData.contents);
    if (b.senha !== SENHA) throw new Error("senha incorreta");
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try { r = tratar(b); } finally { lock.releaseLock(); }
  } catch (err) {
    r = { erro: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(r))
    .setMimeType(ContentService.MimeType.JSON);
}

function aba() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Empresas");
  if (!sh) {
    sh = ss.insertSheet("Empresas");
    sh.appendRow(["ID", "Empresa", "Tipo", "CNPJ/CAEPF", "Cidade", "Ativa", "Última alteração", "Dados (não editar)"]);
    sh.setFrozenRows(1);
    sh.getRange("D:D").setNumberFormat("@");
  }
  return sh;
}

function tratar(b) {
  var sh = aba(), n = sh.getLastRow();
  var ids = n > 1 ? sh.getRange(2, 1, n - 1, 1).getValues().map(function (x) { return String(x[0]); }) : [];
  if (b.acao == "listar") {
    var dados = n > 1 ? sh.getRange(2, 8, n - 1, 1).getValues().map(function (x) {
      try { return JSON.parse(x[0]); } catch (e) { return null; }
    }).filter(Boolean) : [];
    return { dados: dados };
  }
  if (b.acao == "salvar") {
    var c = b.emp;
    var linha = [c.id, c.n || "", c.tipo || "CNPJ", c.cnpj || "", c.cid || "", c.ativa || "Sim",
      (c.ult || "") + " em " + (c.ultD || ""), JSON.stringify(c)];
    var i = ids.indexOf(String(c.id));
    if (i >= 0) sh.getRange(i + 2, 1, 1, 8).setValues([linha]); else sh.appendRow(linha);
    return { ok: true };
  }
  if (b.acao == "excluir") {
    var j = ids.indexOf(String(b.id));
    if (j >= 0) sh.deleteRow(j + 2);
    return { ok: true };
  }
  throw new Error("ação inválida");
}
