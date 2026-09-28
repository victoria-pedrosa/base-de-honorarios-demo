/**
 * =============================================================================
 * MÓDULO 11: FORMULÁRIO DE PREENCHIMENTO (WEB APP) — DADOS DO PILOTO
 * Novo arquivo .gs, mesmo projeto do Módulo 7-9 e do Módulo 10.
 * Plano A: um app na web (HtmlService), publicado a partir deste mesmo projeto,
 * que grava DIRETO nas abas EMPRESAS e HISTORICO_MENSAL — nenhuma aba nova,
 * nenhuma coluna renomeada. Usa Qtd_Funcionários, Nota_Emitida, Honorário Atual,
 * Data Último Reajuste, Motivo Última Alteração e Honorário Sugerido, que já
 * existem, mais Var_Funcionários_Qtd/% e Var_Notas_% (adicionadas no Módulo 10).
 * Arquivo de tela correspondente: FormularioPreenchimento.html (arquivo separado).
 *
 * IMPORTANTE: funções de menu do Módulo 7-9 e do Módulo 10 que chamam
 * SpreadsheetApp.getUi() (para mostrar um alerta) QUEBRAM quando chamadas a
 * partir da tela web (o app publicado não tem "menu" pra mostrar um alerta).
 * Por isso essas funções foram ajustadas para usar alertarSeUiDisponivel_()
 * abaixo, que tenta mostrar o alerta e simplesmente ignora se não der — e
 * agora todas retornam um objeto (ok/erro/contadores) que os botões do
 * formulário usam para mostrar o resultado na tela.
 * =============================================================================
 */

/** Mostra o alerta clássico do Sheets SE a função estiver rodando a partir do
 *  menu/editor; se estiver rodando a partir do app web publicado (sem menu),
 *  simplesmente ignora — nunca deixa a função quebrar por causa do alerta. */
function alertarSeUiDisponivel_(mensagem) {
  try { SpreadsheetApp.getUi().alert(mensagem); } catch (e) { /* chamado fora do editor — ok, ignora */ }
}

/** Publicar como Web App: Implantar > Nova implantação > Aplicativo da Web. */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('FormularioPreenchimento')
    .setTitle('Preenchimento — Monitor de Honorários')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Atalho: abre a mesma tela como painel lateral dentro do próprio Sheets (sem publicar nada). */
function abrirFormularioSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('FormularioPreenchimento')
    .setTitle('Preenchimento — Monitor de Honorários');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Lista as empresas ativas (para a aba "Cadastro" do formulário). */
function listarEmpresasParaFormulario() {
  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_DESTINO);
  var dados = aba.getDataRange().getValues();
  var headers = dados[0];
  var idx = {
    empresa: buscarIndiceColuna(headers, "Empresa"),
    situacao: buscarIndiceColuna(headers, "Situação Tabelão"),
    honorarioAtual: buscarIndiceColuna(headers, "Honorário Atual"),
    dataReajuste: buscarIndiceColuna(headers, "Data Último Reajuste"),
    motivo: buscarIndiceColuna(headers, "Motivo Última Alteração"),
    honorarioSugerido: buscarIndiceColuna(headers, "Honorário Sugerido")
  };
  var lista = [];
  for (var i = 1; i < dados.length; i++) {
    var linha = dados[i];
    if (idx.situacao !== -1 && String(linha[idx.situacao]).trim() === "Suspensa") continue;
    var dataReajuste = linha[idx.dataReajuste];
    lista.push({
      cnpj: linha[0],
      empresa: linha[idx.empresa],
      honorarioAtual: linha[idx.honorarioAtual] || "",
      dataUltimoReajuste: (dataReajuste instanceof Date) ? Utilities.formatDate(dataReajuste, Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
      motivoUltimaAlteracao: linha[idx.motivo] || "",
      honorarioSugerido: linha[idx.honorarioSugerido] || ""
    });
  }
  lista.sort(function(a, b) { return String(a.empresa).localeCompare(String(b.empresa)); });
  return lista;
}

/** Salva os campos de cadastro (uso exclusivo Colaborador 2) de UMA empresa. Upsert por CNPJ. */
function salvarCadastroEmpresa(dados) {
  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_DESTINO);
  var valores = aba.getDataRange().getValues();
  var headers = valores[0];
  var idx = {
    honorarioAtual: buscarIndiceColuna(headers, "Honorário Atual"),
    dataReajuste: buscarIndiceColuna(headers, "Data Último Reajuste"),
    motivo: buscarIndiceColuna(headers, "Motivo Última Alteração"),
    honorarioSugerido: buscarIndiceColuna(headers, "Honorário Sugerido")
  };
  var cnpjAlvo = limparCNPJ(dados.cnpj);
  for (var i = 1; i < valores.length; i++) {
    if (limparCNPJ(valores[i][0]) === cnpjAlvo) {
      var linha = i + 1;
      if (dados.honorarioAtual !== "" && dados.honorarioAtual !== undefined) {
        aba.getRange(linha, idx.honorarioAtual + 1).setValue(Number(dados.honorarioAtual));
      }
      if (dados.dataUltimoReajuste) {
        aba.getRange(linha, idx.dataReajuste + 1).setValue(new Date(dados.dataUltimoReajuste + "T00:00:00"));
      }
      if (dados.motivoUltimaAlteracao !== undefined) {
        aba.getRange(linha, idx.motivo + 1).setValue(dados.motivoUltimaAlteracao);
      }
      if (dados.honorarioSugerido !== "" && dados.honorarioSugerido !== undefined) {
        aba.getRange(linha, idx.honorarioSugerido + 1).setValue(Number(dados.honorarioSugerido));
      }
      return { ok: true };
    }
  }
  return { ok: false, erro: "CNPJ não encontrado: " + dados.cnpj };
}

/** CNPJ + Código Dominio de TODAS as empresas (inclusive suspensas) — usado só no
 *  navegador para casar as linhas de um arquivo importado (Domínio/Nota Focus)
 *  com a empresa certa, seja por Código Dominio ou por CNPJ. */
function obterMapaIdentificadores() {
  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_DESTINO);
  var dados = aba.getDataRange().getValues();
  var headers = dados[0];
  var idxCodigo = buscarIndiceColuna(headers, "Código Dominio");
  var idxNome = buscarIndiceColuna(headers, "Empresa");
  var lista = [];
  for (var i = 1; i < dados.length; i++) {
    lista.push({
      cnpj: dados[i][0],
      empresa: dados[i][idxNome],
      codigoDominio: idxCodigo !== -1 ? String(dados[i][idxCodigo]).replace(/\.0$/, "").trim() : ""
    });
  }
  return lista;
}

/** Competências já existentes em HISTORICO_MENSAL, da mais recente para a mais antiga. */
function obterCompetenciasHistorico() {
  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_HISTORICO);
  var dados = aba.getDataRange().getValues();
  var idxComp = buscarIndiceColuna(dados[0], "Competência");
  var set = {};
  for (var i = 1; i < dados.length; i++) {
    if (dados[i][idxComp]) set[String(dados[i][idxComp]).trim()] = true;
  }
  var lista = Object.keys(set);
  lista.sort(function(a, b) { return competenciaParaChave_(b) - competenciaParaChave_(a); });
  return lista;
}

/** Funcionários/Notas de todas as empresas ativas para UMA competência (para a grade). */
function obterDadosMensais(competencia) {
  var abaEmpresas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_DESTINO);
  var dadosEmpresas = abaEmpresas.getDataRange().getValues();
  var headersEmp = dadosEmpresas[0];
  var idxNome = buscarIndiceColuna(headersEmp, "Empresa");
  var idxSituacao = buscarIndiceColuna(headersEmp, "Situação Tabelão");
  var nomePorCnpj = {}, ativos = [];
  for (var e = 1; e < dadosEmpresas.length; e++) {
    var c = limparCNPJ(dadosEmpresas[e][0]);
    if (!c) continue;
    nomePorCnpj[c] = dadosEmpresas[e][idxNome];
    if (idxSituacao === -1 || String(dadosEmpresas[e][idxSituacao]).trim() !== "Suspensa") ativos.push(c);
  }

  var abaHist = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_HISTORICO);
  var dadosHist = abaHist.getDataRange().getValues();
  var headersHist = dadosHist[0];
  var idxH = {
    competencia: buscarIndiceColuna(headersHist, "Competência"),
    funcionarios: buscarIndiceColuna(headersHist, "Qtd_Funcionários"),
    notas: buscarIndiceColuna(headersHist, "Nota_Emitida")
  };
  var existente = {};
  for (var h = 1; h < dadosHist.length; h++) {
    if (String(dadosHist[h][idxH.competencia]).trim() === competencia) {
      existente[limparCNPJ(dadosHist[h][0])] = {
        funcionarios: dadosHist[h][idxH.funcionarios],
        notas: dadosHist[h][idxH.notas]
      };
    }
  }

  var lista = ativos.map(function(cnpj) {
    var reg = existente[cnpj];
    return {
      cnpj: cnpj,
      empresa: nomePorCnpj[cnpj],
      qtdFuncionarios: (reg && reg.funcionarios !== "") ? reg.funcionarios : "",
      notaEmitida: (reg && reg.notas !== "") ? reg.notas : ""
    };
  });
  lista.sort(function(a, b) { return String(a.empresa).localeCompare(String(b.empresa)); });
  return lista;
}

/**
 * Salva em lote Qtd_Funcionários/Nota_Emitida de uma competência.
 * Upsert por CNPJ+Competência (cria linha em HISTORICO_MENSAL se não existir) e
 * recalcula Var_Funcionários_Qtd/% e Var_Notas_% comparando com o mês anterior.
 */
function salvarDadosMensais(competencia, linhas) {
  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_HISTORICO);
  var dados = aba.getDataRange().getValues();
  var headers = dados[0];
  var idx = {
    competencia: buscarIndiceColuna(headers, "Competência"),
    faturamento: buscarIndiceColuna(headers, "Faturamento"),
    funcionarios: buscarIndiceColuna(headers, "Qtd_Funcionários"),
    notas: buscarIndiceColuna(headers, "Nota_Emitida"),
    sincronizacao: buscarIndiceColuna(headers, "Última Sincronização"),
    varFuncQtd: buscarIndiceColuna(headers, "Var_Funcionários_Qtd"),
    varFuncPct: buscarIndiceColuna(headers, "Var_Funcionários_%"),
    varNotasPct: buscarIndiceColuna(headers, "Var_Notas_%")
  };

  var mapaLinhaCompetenciaAtual = {};   // cnpj -> nº da linha, só na competência que estamos salvando
  var historicoPorCnpj = {};            // cnpj -> { chaveCompetencia: {funcionarios, notas} }
  for (var i = 1; i < dados.length; i++) {
    var cnpjL = limparCNPJ(dados[i][0]);
    var compL = String(dados[i][idx.competencia]).trim();
    if (compL === competencia) mapaLinhaCompetenciaAtual[cnpjL] = i + 1;
    if (!historicoPorCnpj[cnpjL]) historicoPorCnpj[cnpjL] = {};
    historicoPorCnpj[cnpjL][competenciaParaChave_(compL)] = {
      funcionarios: dados[i][idx.funcionarios], notas: dados[i][idx.notas]
    };
  }

  var chaveAtual = competenciaParaChave_(competencia);
  var mesAtual = chaveAtual % 100, anoAtual = Math.floor(chaveAtual / 100);
  var mesAnt = mesAtual - 1, anoAnt = anoAtual;
  if (mesAnt < 1) { mesAnt = 12; anoAnt -= 1; }
  var chaveAnterior = anoAnt * 100 + mesAnt;

  var agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  var linhasNovas = [];
  var atualizados = 0, criados = 0;

  linhas.forEach(function(l) {
    var cnpjLimpo = limparCNPJ(l.cnpj);
    var anterior = (historicoPorCnpj[cnpjLimpo] && historicoPorCnpj[cnpjLimpo][chaveAnterior]) || null;
    var funcAnterior = (anterior && anterior.funcionarios !== "") ? Number(anterior.funcionarios) : null;
    var notasAnterior = (anterior && anterior.notas !== "") ? Number(anterior.notas) : null;

    var funcAtual = (l.qtdFuncionarios === "" || l.qtdFuncionarios === null || l.qtdFuncionarios === undefined) ? "" : Number(l.qtdFuncionarios);
    var notasAtual = (l.notaEmitida === "" || l.notaEmitida === null || l.notaEmitida === undefined) ? "" : Number(l.notaEmitida);

    var varFuncQtd = (funcAtual !== "" && funcAnterior !== null) ? funcAtual - funcAnterior : "";
    var varFuncPct = (funcAtual !== "" && funcAnterior) ? (funcAtual - funcAnterior) / funcAnterior : "";
    var varNotasPct = (notasAtual !== "" && notasAnterior) ? (notasAtual - notasAnterior) / notasAnterior : "";

    if (mapaLinhaCompetenciaAtual[cnpjLimpo]) {
      var linhaDestino = mapaLinhaCompetenciaAtual[cnpjLimpo];
      if (funcAtual !== "") aba.getRange(linhaDestino, idx.funcionarios + 1).setValue(funcAtual);
      if (notasAtual !== "") aba.getRange(linhaDestino, idx.notas + 1).setValue(notasAtual);
      if (idx.varFuncQtd !== -1 && varFuncQtd !== "") aba.getRange(linhaDestino, idx.varFuncQtd + 1).setValue(varFuncQtd);
      if (idx.varFuncPct !== -1 && varFuncPct !== "") aba.getRange(linhaDestino, idx.varFuncPct + 1).setValue(varFuncPct);
      if (idx.varNotasPct !== -1 && varNotasPct !== "") aba.getRange(linhaDestino, idx.varNotasPct + 1).setValue(varNotasPct);
      aba.getRange(linhaDestino, idx.sincronizacao + 1).setValue(agora);
      atualizados++;
    } else {
      var novaLinha = new Array(headers.length).fill("");
      novaLinha[0] = l.cnpj;
      novaLinha[idx.competencia] = competencia;
      novaLinha[idx.faturamento] = 0;
      novaLinha[idx.funcionarios] = funcAtual;
      novaLinha[idx.notas] = notasAtual;
      novaLinha[idx.sincronizacao] = agora;
      if (idx.varFuncQtd !== -1) novaLinha[idx.varFuncQtd] = varFuncQtd;
      if (idx.varFuncPct !== -1) novaLinha[idx.varFuncPct] = varFuncPct;
      if (idx.varNotasPct !== -1) novaLinha[idx.varNotasPct] = varNotasPct;
      linhasNovas.push(novaLinha);
      criados++;
    }
  });

  if (linhasNovas.length > 0) {
    aba.getRange(aba.getLastRow() + 1, 1, linhasNovas.length, headers.length).setValues(linhasNovas);
  }
  return { ok: true, atualizados: atualizados, criados: criados };
}
