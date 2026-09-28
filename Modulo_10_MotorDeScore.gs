/**
 * =============================================================================
 * MÓDULO 10: MOTOR DE SCORE E GATILHOS — MONITOR DE HONORÁRIOS
 * Novo arquivo .gs a ser ADICIONADO ao mesmo projeto Apps Script do Módulo 7-9
 * (usa as constantes/funções globais já existentes: ABA_DESTINO, ABA_HISTORICO,
 * limparCNPJ, buscarIndiceColuna, SLACK_WEBHOOK_URL). Não altera nada do que
 * já existe — apenas adiciona abas e funções novas.
 * =============================================================================
 */

var ABA_CONFIG = "CONFIG_REGRAS";
var ABA_GATILHOS = "GATILHOS_ALERTAS";
var ABA_TRATATIVAS = "TRATATIVAS";

var CONFIG_PADRAO = [
  ["FATURAMENTO_ALERTA_PCT", 0.20, "percentual", "Aumento mínimo de faturamento (usado no gatilho forte)"],
  ["FATURAMENTO_FORTE_PCT", 0.30, "percentual", "Aumento de faturamento isolado (gatilho financeiro, sem funcionários)"],
  ["FUNCIONARIOS_ALERTA_PCT", 0.20, "percentual", "Aumento mínimo de funcionários (usado no gatilho forte)"],
  ["FUNCIONARIOS_ALERTA_ABS", 3, "numero", "Aumento absoluto mínimo de funcionários (usado no gatilho forte)"],
  ["FUNCIONARIOS_FORTE_PCT", 0.30, "percentual", "Aumento de funcionários isolado (gatilho de folha)"],
  ["FUNCIONARIOS_FORTE_ABS", 5, "numero", "Aumento absoluto isolado de funcionários (gatilho de folha)"],
  ["NOTAS_ALERTA_PCT", 0.20, "percentual", "Notas do mês acima da média móvel de 6 meses"],
  ["MESES_RECORRENCIA_1", 2, "numero", "Meses consecutivos com gatilho para recorrência nível 1"],
  ["MESES_RECORRENCIA_2", 3, "numero", "Meses consecutivos com gatilho para recorrência nível 2"],
  ["REAJUSTE_ANTIGO_MESES_1", 24, "numero", "Meses sem reajuste — alerta"],
  ["REAJUSTE_ANTIGO_MESES_2", 36, "numero", "Meses sem reajuste — crítico"],
  ["SCORE_MINIMO_REGISTRO", 30, "numero", "Score mínimo para gravar em GATILHOS_ALERTAS"],
  ["SCORE_MINIMO_ALERTA", 50, "numero", "Score mínimo para disparo de alerta Slack/Gmail"],
  ["PONTOS_GATILHO_FORTE", 50, "pontos", "🔴 Faturamento + Funcionários"],
  ["PONTOS_GATILHO_FINANCEIRO", 25, "pontos", "🟠 Faturamento forte isolado"],
  ["PONTOS_GATILHO_FOLHA", 20, "pontos", "🟡 Aumento de funcionários isolado"],
  ["PONTOS_GATILHO_NOTA_VOLUME", 30, "pontos", "🔵 Volume de notas acima da média móvel"],
  ["PONTOS_GATILHO_NOTA_AUSENTE", 15, "pontos", "🔵 Faturamento sem nota emitida no mês"],
  ["PONTOS_RECORRENCIA_1", 20, "pontos", "🟣 Recorrente (2 meses)"],
  ["PONTOS_RECORRENCIA_2", 35, "pontos", "🟣 Recorrente (3 meses)"],
  ["PONTOS_REAJUSTE_1", 20, "pontos", "Sem reajuste há mais de 24 meses"],
  ["PONTOS_REAJUSTE_2", 30, "pontos", "Sem reajuste há mais de 36 meses"]
];

function onOpen_Modulo10() {
  // Não declarar outro onOpen() neste arquivo (colide com o já existente).
  // Ver instrução de patch manual no onOpen do Módulo 7-9.
}

/** Cria (se não existirem) as abas CONFIG_REGRAS, GATILHOS_ALERTAS e TRATATIVAS,
 *  e garante as colunas novas em EMPRESAS/HISTORICO_MENSAL. Idempotente. */
function criarAbasAuxiliares() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var abaConfig = ss.getSheetByName(ABA_CONFIG);
  if (!abaConfig) {
    abaConfig = ss.insertSheet(ABA_CONFIG);
    abaConfig.getRange(1, 1, 1, 4).setValues([["Regra", "Valor", "Tipo", "Descrição"]])
      .setFontWeight("bold").setBackground("#274e13").setFontColor("#ffffff");
    abaConfig.getRange(2, 1, CONFIG_PADRAO.length, 4).setValues(CONFIG_PADRAO);
  }

  var abaGatilhos = ss.getSheetByName(ABA_GATILHOS);
  if (!abaGatilhos) {
    abaGatilhos = ss.insertSheet(ABA_GATILHOS);
    abaGatilhos.getRange(1, 1, 1, 11).setValues([[
      "ID_Alerta", "CNPJ", "Empresa", "Competência Referência", "Gatilhos_Ativados",
      "Score_Calculado", "Classificação", "Status", "Responsável", "Data_Geração", "Observação"
    ]]).setFontWeight("bold").setBackground("#274e13").setFontColor("#ffffff");
  }

  var abaTratativas = ss.getSheetByName(ABA_TRATATIVAS);
  if (!abaTratativas) {
    abaTratativas = ss.insertSheet(ABA_TRATATIVAS);
    abaTratativas.getRange(1, 1, 1, 10).setValues([[
      "ID_Tratativa", "ID_Alerta", "CNPJ", "Data_Ação", "Responsável", "Etapa_Funil",
      "Honorário_Anterior", "Honorário_Novo", "Incremento_Mensal", "Observação"
    ]]).setFontWeight("bold").setBackground("#274e13").setFontColor("#ffffff");
  }

  garantirColunasScoreEmpresas();
  garantirColunasHistorico();
  var msg = "✅ Abas auxiliares OK (CONFIG_REGRAS, GATILHOS_ALERTAS, TRATATIVAS) + colunas novas em EMPRESAS/HISTORICO_MENSAL.";
  alertarSeUiDisponivel_(msg);
  return { ok: true, mensagem: msg };
}

/** Adiciona "Score Atual" e "Status Comercial" ao final de EMPRESAS, se não existirem. */
function garantirColunasScoreEmpresas() {
  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_DESTINO);
  var headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
  ["Score Atual", "Status Comercial"].forEach(function(nome) {
    if (buscarIndiceColuna(headers, nome) === -1) {
      var novaCol = aba.getLastColumn() + 1;
      aba.getRange(1, novaCol).setValue(nome).setFontWeight("bold").setBackground("#274e13").setFontColor("#ffffff");
      headers.push(nome);
    }
  });
}

/** Adiciona "Var_Funcionários_%" e "Var_Notas_%" ao final de HISTORICO_MENSAL, se não existirem.
 *  Não mexe na coluna "Var_Funcionários_Qtd" já existente. */
function garantirColunasHistorico() {
  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_HISTORICO);
  var headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
  ["Var_Funcionários_%", "Var_Notas_%"].forEach(function(nome) {
    if (buscarIndiceColuna(headers, nome) === -1) {
      var novaCol = aba.getLastColumn() + 1;
      aba.getRange(1, novaCol).setValue(nome).setFontWeight("bold").setBackground("#274e13").setFontColor("#ffffff");
      headers.push(nome);
    }
  });
}

/** Lê CONFIG_REGRAS numa vez; cai para CONFIG_PADRAO se a aba/linha não existir. */
function obterConfig_() {
  var cfg = {};
  CONFIG_PADRAO.forEach(function(l) { cfg[l[0]] = l[1]; });
  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_CONFIG);
  if (aba) {
    var dados = aba.getDataRange().getValues();
    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0]) cfg[String(dados[i][0]).trim()] = dados[i][1];
    }
  }
  return cfg;
}

/** "MM / AAAA" -> 202501 (chave numérica p/ ordenar corretamente). */
function competenciaParaChave_(competencia) {
  var partes = String(competencia).split("/").map(function(p) { return p.trim(); });
  if (partes.length !== 2) return 0;
  var mes = parseInt(partes[0], 10), ano = parseInt(partes[1], 10);
  if (!mes || !ano) return 0;
  return ano * 100 + mes;
}

/**
 * MOTOR DE SCORE — varre EMPRESAS x HISTORICO_MENSAL, calcula o score de cada
 * empresa a partir do histórico mais recente e grava em GATILHOS_ALERTAS
 * (upsert por CNPJ+Competência) quando Score_Calculado >= SCORE_MINIMO_REGISTRO.
 * Atualiza também "Score Atual" em EMPRESAS (sempre) e "Status Comercial"
 * (somente se ainda estiver vazio — não sobrescreve tratativa em andamento).
 */
function calcularScoreEGatilhos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaEmpresas = ss.getSheetByName(ABA_DESTINO);
  var abaHist = ss.getSheetByName(ABA_HISTORICO);
  if (!abaEmpresas || !abaHist) {
    var msgErro1 = "❌ Rode antes as sincronizações (opções 1 e 2).";
    alertarSeUiDisponivel_(msgErro1);
    return { ok: false, erro: msgErro1 };
  }

  var cfg = obterConfig_();
  garantirColunasScoreEmpresas();

  var dadosEmpresas = abaEmpresas.getDataRange().getValues();
  var headersEmpresas = dadosEmpresas[0];
  var idxEmp = {
    cnpj: 0, nome: buscarIndiceColuna(headersEmpresas, "Empresa"),
    dataReajuste: buscarIndiceColuna(headersEmpresas, "Data Último Reajuste"),
    scoreAtual: buscarIndiceColuna(headersEmpresas, "Score Atual"),
    statusComercial: buscarIndiceColuna(headersEmpresas, "Status Comercial")
  };

  var dadosHist = abaHist.getDataRange().getValues();
  var headersHist = dadosHist[0];
  var idxH = {
    cnpj: 0,
    competencia: buscarIndiceColuna(headersHist, "Competência"),
    faturamento: buscarIndiceColuna(headersHist, "Faturamento"),
    funcionarios: buscarIndiceColuna(headersHist, "Qtd_Funcionários"),
    notas: buscarIndiceColuna(headersHist, "Nota_Emitida")
  };

  // Agrupa histórico por CNPJ, já ordenado por competência (chave YYYYMM).
  var historicoPorCnpj = {};
  for (var h = 1; h < dadosHist.length; h++) {
    var linha = dadosHist[h];
    var cnpjH = limparCNPJ(linha[idxH.cnpj]);
    if (!cnpjH) continue;
    if (!historicoPorCnpj[cnpjH]) historicoPorCnpj[cnpjH] = [];
    historicoPorCnpj[cnpjH].push({
      chave: competenciaParaChave_(linha[idxH.competencia]),
      competencia: linha[idxH.competencia],
      faturamento: Number(linha[idxH.faturamento]) || 0,
      funcionarios: (linha[idxH.funcionarios] === "" || linha[idxH.funcionarios] === null) ? null : Number(linha[idxH.funcionarios]),
      notas: (linha[idxH.notas] === "" || linha[idxH.notas] === null) ? null : Number(linha[idxH.notas])
    });
  }
  Object.keys(historicoPorCnpj).forEach(function(cnpj) {
    historicoPorCnpj[cnpj].sort(function(a, b) { return a.chave - b.chave; });
  });

  // Índice das linhas já existentes em GATILHOS_ALERTAS (upsert por ID_Alerta).
  var abaGatilhos = ss.getSheetByName(ABA_GATILHOS);
  if (!abaGatilhos) {
    var msgErro2 = "❌ Rode antes '📐 Criar Abas Auxiliares'.";
    alertarSeUiDisponivel_(msgErro2);
    return { ok: false, erro: msgErro2 };
  }
  var dadosGatilhos = abaGatilhos.getDataRange().getValues();
  var mapaGatilhos = {};
  for (var g = 1; g < dadosGatilhos.length; g++) mapaGatilhos[dadosGatilhos[g][0]] = g + 1;

  var agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  var linhasNovas = [];
  var gravados = 0, avaliados = 0;

  for (var e = 1; e < dadosEmpresas.length; e++) {
    var cnpjBruto = dadosEmpresas[e][idxEmp.cnpj];
    var cnpjLimpo = limparCNPJ(cnpjBruto);
    var hist = historicoPorCnpj[cnpjLimpo];
    if (!hist || hist.length < 2) continue; // precisa de ao menos 2 competências p/ comparar
    avaliados++;

    var n = hist.length - 1;
    var resultadoAtual = calcularScoreBaseNoMes_(hist, n, cfg);
    var resultadoM1 = (n - 1 >= 1) ? calcularScoreBaseNoMes_(hist, n - 1, cfg) : null;
    var resultadoM2 = (n - 2 >= 1) ? calcularScoreBaseNoMes_(hist, n - 2, cfg) : null;

    var score = resultadoAtual.score;
    var gatilhos = resultadoAtual.gatilhos.slice();

    if (resultadoAtual.score > 0 && resultadoM1 && resultadoM1.score > 0) {
      if (resultadoM2 && resultadoM2.score > 0) {
        score += cfg.PONTOS_RECORRENCIA_2;
        gatilhos.push("🟣 Recorrente 3 meses");
      } else {
        score += cfg.PONTOS_RECORRENCIA_1;
        gatilhos.push("🟣 Recorrente 2 meses");
      }
    }

    var dataReajuste = idxEmp.dataReajuste !== -1 ? dadosEmpresas[e][idxEmp.dataReajuste] : null;
    if (dataReajuste instanceof Date) {
      var meses = (new Date().getTime() - dataReajuste.getTime()) / (1000 * 60 * 60 * 24 * 30.44);
      if (meses >= cfg.REAJUSTE_ANTIGO_MESES_2) { score += cfg.PONTOS_REAJUSTE_2; gatilhos.push("⏰ Sem reajuste > 36 meses"); }
      else if (meses >= cfg.REAJUSTE_ANTIGO_MESES_1) { score += cfg.PONTOS_REAJUSTE_1; gatilhos.push("⏰ Sem reajuste > 24 meses"); }
    }

    // Score Atual sempre atualizado; Status Comercial só se ainda vazio (não sobrescreve tratativa em curso).
    if (idxEmp.scoreAtual !== -1) abaEmpresas.getRange(e + 1, idxEmp.scoreAtual + 1).setValue(score);
    if (score >= cfg.SCORE_MINIMO_REGISTRO && idxEmp.statusComercial !== -1 && !dadosEmpresas[e][idxEmp.statusComercial]) {
      abaEmpresas.getRange(e + 1, idxEmp.statusComercial + 1).setValue("🆕 Novo alerta");
    }

    if (score < cfg.SCORE_MINIMO_REGISTRO) continue;

    var idAlerta = cnpjLimpo + "-" + hist[n].chave;
    var linhaGravar = [
      idAlerta, cnpjBruto, dadosEmpresas[e][idxEmp.nome], hist[n].competencia,
      gatilhos.join(", "), score, classificarScore_(score), "🆕 Novo", "", agora, ""
    ];
    if (mapaGatilhos[idAlerta]) {
      // Upsert: preserva Status/Responsável/Observação já lançados manualmente (colunas 8, 9, 11).
      var linhaExistente = mapaGatilhos[idAlerta];
      abaGatilhos.getRange(linhaExistente, 5, 1, 3).setValues([[gatilhos.join(", "), score, classificarScore_(score)]]);
      abaGatilhos.getRange(linhaExistente, 10).setValue(agora);
    } else {
      linhasNovas.push(linhaGravar);
    }
    gravados++;
  }

  if (linhasNovas.length > 0) {
    abaGatilhos.getRange(abaGatilhos.getLastRow() + 1, 1, linhasNovas.length, 11).setValues(linhasNovas);
  }
  var msgFinal = "✅ Score calculado.\nEmpresas avaliadas: " + avaliados + "\nCom gatilho (score ≥ " + cfg.SCORE_MINIMO_REGISTRO + "): " + gravados;
  alertarSeUiDisponivel_(msgFinal);
  return { ok: true, avaliados: avaliados, gravados: gravados, scoreMinimo: cfg.SCORE_MINIMO_REGISTRO };
}

function classificarScore_(score) {
  if (score >= 70) return "🔴 Alta prioridade";
  if (score >= 50) return "🟠 Analisar";
  if (score >= 30) return "🟡 Monitorar";
  return "🟢 Normal";
}

/** Calcula o score-base (sem recorrência/reajuste) do mês hist[i] vs hist[i-1]. Requer i>=1. */
function calcularScoreBaseNoMes_(hist, i, cfg) {
  var atual = hist[i], anterior = hist[i - 1];
  var score = 0, gatilhos = [];

  var varFat = anterior.faturamento > 0 ? (atual.faturamento - anterior.faturamento) / anterior.faturamento : null;
  var funcOk = atual.funcionarios !== null && anterior.funcionarios !== null;
  var varFuncPct = (funcOk && anterior.funcionarios > 0) ? (atual.funcionarios - anterior.funcionarios) / anterior.funcionarios : null;
  var varFuncAbs = funcOk ? (atual.funcionarios - anterior.funcionarios) : null;

  if (varFat !== null) {
    var funcSubiu = funcOk && ((varFuncPct !== null && varFuncPct >= cfg.FUNCIONARIOS_ALERTA_PCT) || varFuncAbs >= cfg.FUNCIONARIOS_ALERTA_ABS);
    if (varFat >= cfg.FATURAMENTO_ALERTA_PCT && funcSubiu) {
      score += cfg.PONTOS_GATILHO_FORTE; gatilhos.push("🔴 Faturamento + Funcionários");
    } else if (varFat >= cfg.FATURAMENTO_FORTE_PCT) {
      score += cfg.PONTOS_GATILHO_FINANCEIRO; gatilhos.push("🟠 Faturamento forte");
    }
  }

  if (funcOk && ((varFuncPct !== null && varFuncPct >= cfg.FUNCIONARIOS_FORTE_PCT) || varFuncAbs >= cfg.FUNCIONARIOS_FORTE_ABS)) {
    score += cfg.PONTOS_GATILHO_FOLHA; gatilhos.push("🟡 Aumento de funcionários");
  }

  // Notas: volume vs média móvel de até 6 meses anteriores (excluindo o mês atual).
  if (atual.notas !== null) {
    var janela = hist.slice(Math.max(0, i - 6), i).map(function(r) { return r.notas; }).filter(function(v) { return v !== null; });
    if (janela.length > 0) {
      var media = janela.reduce(function(a, b) { return a + b; }, 0) / janela.length;
      if (media > 0 && atual.notas >= media * (1 + cfg.NOTAS_ALERTA_PCT)) {
        score += cfg.PONTOS_GATILHO_NOTA_VOLUME; gatilhos.push("🔵 Volume de notas acima da média");
      }
    }
    if (atual.faturamento > 0 && atual.notas === 0) {
      score += cfg.PONTOS_GATILHO_NOTA_AUSENTE; gatilhos.push("🔵 Faturamento sem nota emitida");
    }
  }

  return { score: score, gatilhos: gatilhos };
}

/**
 * Envia alerta (Slack + Gmail) com os gatilhos "🆕 Novo" de score alto,
 * gerados na última rodada de calcularScoreEGatilhos(). Não mexe na
 * enviarAlertasGatilhos() já existente (Módulo 9) — função nova, independente.
 * Rode manualmente ou agende (ver criarGatilhoAlertaScore()).
 */
function enviarAlertasScore() {
  var cfg = obterConfig_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaGatilhos = ss.getSheetByName(ABA_GATILHOS);
  if (!abaGatilhos) return { ok: false, erro: "Rode antes '📐 Criar Abas Auxiliares'." };
  var dados = abaGatilhos.getDataRange().getValues();

  var alertas = [];
  for (var i = 1; i < dados.length; i++) {
    var linha = dados[i];
    var score = Number(linha[5]);
    var status = String(linha[7]);
    if (status === "🆕 Novo" && score >= cfg.SCORE_MINIMO_ALERTA) {
      alertas.push({ empresa: linha[2], cnpj: linha[1], score: score, classificacao: linha[6], gatilhos: linha[4] });
    }
  }
  if (alertas.length === 0) return { ok: true, enviados: 0 };
  alertas.sort(function(a, b) { return b.score - a.score; });

  var linhasTexto = alertas.map(function(a) {
    return "• " + a.empresa + " (" + a.cnpj + ") — Score " + a.score + " " + a.classificacao + "\n   " + a.gatilhos;
  });
  var corpo = "🎯 Monitor de Honorários — " + alertas.length + " cliente(s) com score ≥ " + cfg.SCORE_MINIMO_ALERTA + ":\n\n" + linhasTexto.join("\n\n");

  if (typeof SLACK_WEBHOOK_URL !== "undefined" && SLACK_WEBHOOK_URL) {
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, { method: "post", contentType: "application/json", payload: JSON.stringify({ text: corpo }) });
  }
  if (typeof ALERTA_EMAIL_DESTINATARIOS !== "undefined" && ALERTA_EMAIL_DESTINATARIOS) {
    MailApp.sendEmail({ to: ALERTA_EMAIL_DESTINATARIOS, subject: "🎯 " + alertas.length + " cliente(s) prioritário(s) — Monitor de Honorários", body: corpo });
  }
  return { ok: true, enviados: alertas.length };
}

/** Rode uma vez manualmente: agenda o motor de score diariamente às 07h (após sync das 06h). */
function criarGatilhoScoreAutomatico() {
  ScriptApp.newTrigger('calcularScoreEGatilhos').timeBased().everyDays(1).atHour(7).create();
  ScriptApp.newTrigger('enviarAlertasScore').timeBased().everyDays(1).atHour(7).nearMinute(15).create();
}
