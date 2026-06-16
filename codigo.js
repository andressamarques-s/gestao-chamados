var NOME_ABA = "CHAMADOS SUPORTE - 2026"

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Suporte de Sistemas")
    .addItem("Abrir sistema de chamados", "abrirSistema")
    .addToUi();
}

function abrirSistema() {
  let html = HtmlService.createTemplateFromFile("Index").evaluate()
    .setWidth(700)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Sistema de Chamados")
}

function proximoNumeroChamado(origem) {
  let ultimaLinha = origem.getLastRow();
  if (ultimaLinha < 2) return 1;
  let numeros = origem.getRange(2, 1, ultimaLinha - 1, 1).getValues()
    .map(function(r){ return Number(r[0]) || 0; });
  return Math.max.apply(null, numeros) + 1;
}

// Código para tela de Registro de Chamados
function registrarChamado(dados) {
  let origem = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA);
  let linhaVazia = origem.getLastRow() + 1;

  origem.getRange(linhaVazia, 1).setValue(proximoNumeroChamado(origem));
  origem.getRange(linhaVazia, 2).setValue(dados.concluido ? "Concluído" : "Em andamento");
  origem.getRange(linhaVazia, 3).setValue(dados.suporteInicial);
  let agora = new Date();
  origem.getRange(linhaVazia, 4).setValue(Utilities.formatDate(agora, Session.getScriptTimeZone(), "dd/MM/yyyy"));
  origem.getRange(linhaVazia, 5).setValue(Utilities.formatDate(agora, Session.getScriptTimeZone(), "HH:mm:ss"));
  origem.getRange(linhaVazia, 6).setValue(dados.solicitante);
  origem.getRange(linhaVazia, 7).setValue(dados.instituicao);
  origem.getRange(linhaVazia, 8).setValue(dados.demanda);
  origem.getRange(linhaVazia, 9).setValue(dados.motivo);
  origem.getRange(linhaVazia, 10).setValue(dados.resolucao);
  origem.getRange(linhaVazia, 11).setValue(dados.categoria);
  origem.getRange(linhaVazia, 12).setValue(dados.suporteFinal);

  if (dados.concluido) {
    origem.getRange(linhaVazia, 13).setValue(dados.dataFinal);
    origem.getRange(linhaVazia, 14).setValue(dados.horaFinal);
  }

  origem.getRange(linhaVazia, 20).setValue(dados.observacao);
}

function concluirChamado(dados){
  let origem = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA);
  let linha = dados.linha;
  origem.getRange(linha, 2).setValue("Concluído");
  origem.getRange(linha, 10).setValue(dados.resolucao);
  origem.getRange(linha, 11).setValue(dados.categoria);
  origem.getRange(linha, 12).setValue(dados.suporteFinal);

  let agora = new Date();
  let dataFinal = dados.dataFinal && dados.dataFinal !== ""
    ? dados.dataFinal
    : Utilities.formatDate(agora, Session.getScriptTimeZone(), "dd/MM/yyyy");
  let horaFinal = dados.horaFinal && dados.horaFinal !== ""
    ? dados.horaFinal
    : Utilities.formatDate(agora, Session.getScriptTimeZone(), "HH:mm:ss");

  origem.getRange(linha, 13).setValue(dataFinal);
  origem.getRange(linha, 14).setValue(horaFinal);
  origem.getRange(linha, 20).setValue(dados.observacao);
}

function obterChamados(){
  let origem = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA);
  let ultimaLinha = origem.getLastRow();
  let chamados = [];

  if (ultimaLinha < 2) return chamados;

  let dados = origem.getRange(2, 1, ultimaLinha - 1, 20).getValues();

  dados.forEach(function(linha, index){
    let status = String(linha[1]);
    if (status == "Em andamento"){
      chamados.push({
        linha: index + 2,
        numero: linha[0],
        status: linha[1],
        suporteInicial: linha[2],
        dataInicial: formatarData(linha[3]),
        horaInicial: formatarHora(linha[4]),
        solicitante: linha[5],
        instituicao: linha[6],
        demanda: linha[7],
        motivo: linha[8],
        resolucao: linha[9],
        categoria: linha[10],
        suporteFinal: linha[11],
        dataFinal: formatarData(linha[12]),
        horaFinal: formatarHora(linha[13]),
        observacao: linha[19]
      });
    }
  });

  return chamados;
}

function definirAgrupamento(dataInicio, dataFim) {
  if (!dataInicio || !dataFim) return "dia";
  let diffMs   = dataFim.getTime() - dataInicio.getTime();
  let diffDias = diffMs / (1000 * 60 * 60 * 24);
  if (diffDias <= 31)  return "dia";
  if (diffDias <= 90)  return "semana";
  return "mes";
}

function chaveAgrupamento(dataObj, agrupamento, tz) {
  if (agrupamento === "mes") {
    return {
      ordem:    Utilities.formatDate(dataObj, tz, "yyyy-MM"),
      exibicao: Utilities.formatDate(dataObj, tz, "MMM/yyyy")
    };
  }
  if (agrupamento === "semana") {
    let d = new Date(dataObj.getTime());
    d.setDate(d.getDate() - d.getDay());
    return {
      ordem:    Utilities.formatDate(d, tz, "yyyy-MM-dd"),
      exibicao: "Sem " + Utilities.formatDate(d, tz, "dd/MM")
    };
  }
  // dia
  return {
    ordem:    Utilities.formatDate(dataObj, tz, "yyyy-MM-dd"),
    exibicao: Utilities.formatDate(dataObj, tz, "dd/MM")
  };
}

// Painel
function obterDadosPainel(dataInicioStr, dataFimStr){
  let origem = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA);
  let ultimaLinha = origem.getLastRow();

  let resultado = {
    abertos: [],
    totalAbertos: 0,
    totalConcluidos: 0,
    totalPeriodo: 0,
    porDia: [],          
    porAtendente: [],
    agrupamento: "dia"   
  };

  if (ultimaLinha < 2) return resultado;

  let dados = origem.getRange(2, 1, ultimaLinha - 1, 20).getValues();

  let dataInicio = dataInicioStr ? new Date(dataInicioStr + "T00:00:00") : null;
  let dataFim    = dataFimStr    ? new Date(dataFimStr    + "T23:59:59") : null;


  let agrupamento = definirAgrupamento(dataInicio, dataFim);
  resultado.agrupamento = agrupamento;

  let tz = Session.getScriptTimeZone();
  let porDiaMapa       = {};
  let porAtendenteMapa = {};

  dados.forEach(function(linha, index){
    let status      = String(linha[1]);
    let dataChamado = linha[3];

    let dataObj = null;
    if (dataChamado instanceof Date) {
      dataObj = dataChamado;
    } else if (typeof dataChamado === "string" && dataChamado.indexOf("/") > -1) {
      let partes = dataChamado.split("/");
      dataObj = new Date(partes[2], partes[1] - 1, partes[0]);
    }

    let dentroDoPeriodo = true;
    if (dataObj) {
      if (dataInicio && dataObj < dataInicio) dentroDoPeriodo = false;
      if (dataFim    && dataObj > dataFim)    dentroDoPeriodo = false;
    }

    if (status == "Em andamento") {
      resultado.totalAbertos++;
      resultado.abertos.push({
        linha: index + 2,
        numero: linha[0],
        status: linha[1],
        suporteInicial: linha[2],
        dataInicial: formatarData(linha[3]),
        horaInicial: formatarHora(linha[4]),
        solicitante: linha[5],
        instituicao: linha[6],
        demanda: linha[7],
        motivo: linha[8],
        resolucao: linha[9],
        categoria: linha[10],
        suporteFinal: linha[11],
        dataFinal: formatarData(linha[12]),
        horaFinal: formatarHora(linha[13]),
        observacao: linha[19]
      });
    }

    if (dentroDoPeriodo) {
      resultado.totalPeriodo++;

      if (status == "Concluído") {
        resultado.totalConcluidos++;
      }

      if (dataObj) {
        let chave = chaveAgrupamento(dataObj, agrupamento, tz);
        if (!porDiaMapa[chave.ordem]) {
          porDiaMapa[chave.ordem] = { label: chave.exibicao, total: 0 };
        }
        porDiaMapa[chave.ordem].total++;
      }

      let atendente = linha[2];
      if (atendente) {
        porAtendenteMapa[atendente] = (porAtendenteMapa[atendente] || 0) + 1;
      }
    }
  });

  resultado.porDia = Object.keys(porDiaMapa).sort().map(function(chave){
    return { data: porDiaMapa[chave].label, total: porDiaMapa[chave].total };
  });

  resultado.porAtendente = Object.keys(porAtendenteMapa)
    .map(function(nome){ return { nome: nome, total: porAtendenteMapa[nome] }; })
    .sort(function(a, b){ return b.total - a.total; });

  return resultado;
}

function obterListas(){
  let dados = PropertiesService.getScriptProperties().getProperty('Listas');
  if (!dados){
    return {
      suporte: [],
      instituicao: [],
      demanda: [],
      categoria: []
    };
  }
  return JSON.parse(dados);
}

function salvarListas(listas){
  PropertiesService.getScriptProperties().setProperty('Listas', JSON.stringify(listas));
}

function testarRegistro() {
  let dados = {
    suporteInicial: "Andressa",
    solicitante: "Teste",
    instituicao: "Clínica Arqué Saúde",
    demanda: "SOC",
    motivo: "Teste de registro",
    resolucao: "Testando",
    categoria: "Acesso ao SOC",
    suporteFinal: "Andressa",
    observacao: "",
    concluido: false
  };
  registrarChamado(dados);
}

function testarObterChamados() {
  let origem = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_ABA);
  Logger.log("Aba encontrada: " + (origem !== null));
  Logger.log("Última linha: " + origem.getLastRow());
  let resultado = obterChamados();
  Logger.log("Total encontrados: " + resultado.length);
}

function obterChamadosDebug(){
  try {
    let r = obterChamados();
    return JSON.stringify(r);
  } catch(e) {
    return "ERRO: " + e.message;
  }
}

function formatarData(valor) {
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  return valor;
}

function formatarHora(valor) {
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), "HH:mm:ss");
  }
  return valor;
}
