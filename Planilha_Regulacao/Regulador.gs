function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏥 Sistema Regulação')
    .addItem('Abrir Painel de Controle', 'abrirPainelRegulador')
    .addSeparator() 
    .addItem('Ver Painel Geral (Dashboard)', 'abrirDashboardRegulacao')
    .addToUi();
}

function abrirPainelRegulador() {
  var html = HtmlService.createHtmlOutputFromFile('PainelRegulador')
      .setTitle('Sistema de Regulação Médica - Painel de Controle')
      .setWidth(1150)  
      .setHeight(680); 
      
  SpreadsheetApp.getUi().showModelessDialog(html, 'Painel de Regulação');
}

function abrirDashboardRegulacao() {
  var html = HtmlService.createHtmlOutputFromFile('Dashboard')
      .setTitle('SUS - Dashboard de Sala de Situação da Regulação')
      .setWidth(1200)
      .setHeight(750);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Regulação');
}

function buscarDemandasPaciente(cpfRaw) {
  var retorno = {
    encontrado: false,
    nome: "",
    cpf: cpfRaw,
    demandas: []
  };
  
  if (!cpfRaw) return retorno;
  
  var cpfLimpo = cpfRaw.toString().replace(/\D/g, "");
  cpfLimpo = cpfLimpo.padStart(11, "0"); 
  
  if (cpfLimpo.length !== 11) return retorno;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaRegular = ss.getSheetByName("Regular");
  var abaHistorico = ss.getSheetByName("Historico");
  
  if (abaRegular) {
    var dadosReg = abaRegular.getDataRange().getValues();
    for (var i = 1; i < dadosReg.length; i++) {
      if (!dadosReg[i][1]) continue;
      
      var cpfLinha = dadosReg[i][1].toString().replace(/\D/g, "").padStart(11, "0");
      
      if (cpfLinha === cpfLimpo) {
        retorno.encontrado = true;
        retorno.nome = dadosReg[i][2]; // Nome do Paciente
        
        retorno.demandas.push({
          idObjeto: dadosReg[i][9], 
          statusAtual: dadosReg[i][6] ? dadosReg[i][6].toString().toUpperCase() : "EM FILA ATIVA",
          tipoSolicitacao: dadosReg[i][12] ? dadosReg[i][12].toString().toUpperCase() : (dadosReg[i][10] ? "DEMANDA RETORNO" : "CASO NOVO"),
          profSolicitante: dadosReg[i][0] ? dadosReg[i][0].toString().toUpperCase() : "NÃO INFORMADO",
          especialidade: dadosReg[i][3],
          cid: dadosReg[i][4],
          risco: dadosReg[i][6],
          motivo: dadosReg[i][5],
          codigoReg: dadosReg[i][9],
          origemUnidade: dadosReg[i][8] || "Unidade Não Mapeada",
          data: dadosReg[i][7] instanceof Date ? Utilities.formatDate(dadosReg[i][7], Session.getScriptTimeZone(), "dd/MM/yyyy") : dadosReg[i][7],
          jaFinalizado: false 
        });
      }
    }
  }
  
  if (!retorno.encontrado && abaHistorico) {
    var dadosHist = abaHistorico.getDataRange().getValues();
    for (var j = 1; j < dadosHist.length; j++) {
      if (!dadosHist[j][1]) continue;
      
      var cpfHist = dadosHist[j][1].toString().replace(/\D/g, "").padStart(11, "0");
      
      if (cpfHist === cpfLimpo) {
        retorno.encontrado = true;
        retorno.nome = dadosHist[j][2]; 
        
        retorno.demandas.push({
          idObjeto: dadosHist[j][9],
          statusAtual: "ARQUIVADO / REGULADO",
          tipoSolicitacao: dadosHist[j][12] ? dadosHist[j][12].toString().toUpperCase() : "CONCLUÍDO",
          profSolicitante: dadosHist[j][0] ? dadosHist[j][0].toString().toUpperCase() : "NÃO INFORMADO",
          especialidade: dadosHist[j][3],
          cid: dadosHist[j][4],
          risco: dadosHist[j][6],
          motivo: dadosHist[j][5],
          codigoReg: dadosHist[j][9],
          origemUnidade: dadosHist[j][8] || "Unidade Não Mapeada",
          data: dadosHist[j][7] instanceof Date ? Utilities.formatDate(dadosHist[j][7], Session.getScriptTimeZone(), "dd/MM/yyyy") : dadosHist[j][7],
          sisregSalvo: dadosHist[j][10], 
          dataConclusao: dadosHist[j][11] instanceof Date ? Utilities.formatDate(dadosHist[j][11], Session.getScriptTimeZone(), "dd/MM/yyyy") : dadosHist[j][11],
          diasEspera: dadosHist[j][13],
          jaFinalizado: true 
        });
      }
    }
  }
  
  if (retorno.encontrado) {
    retorno.cpf = cpfLimpo.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, "$1.$2.$3-$4");
  }
  
  return retorno;
}

function concluirDemandaDoPainel(codigoIdentificador, numSisreg, decisaoStatus) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaRegular = ss.getSheetByName("Regular"); 
    var abaHistorico = ss.getSheetByName("Historico");
    
    if (!codigoIdentificador || codigoIdentificador === "N/A") return "Erro: Código único de identificação inválido.";
    if (!abaRegular) return "Erro: Aba 'Regular' não encontrada.";
    
    var sisregTexto = numSisreg ? numSisreg.toString().trim() : "";
    if (sisregTexto === "") {
      return "Erro: O preenchimento do código SISREG é obrigatório para concluir a regulação.";
    }
    
    var ID_PLANILHA_EQUIPES = "1VkuzYhlOnxyizkxe49VXt-h0JaU0Js5eGbY3sF4zxVE";
    
    var dadosReg = abaRegular.getDataRange().getValues();
    var totalLinhasNaTabela = dadosReg.length; 
    var linhaAlvo = -1;
    
    for (var i = 1; i < dadosReg.length; i++) {
      if (dadosReg[i][9] && dadosReg[i][9].toString().trim() === codigoIdentificador.toString().trim()) {
        linhaAlvo = i + 1; 
        break;
      }
    }
    
    if (linhaAlvo !== -1) {
      var rangeCompleto = abaRegular.getRange(linhaAlvo, 1, 1, abaRegular.getLastColumn());
      var dadosLinha = rangeCompleto.getValues()[0];
      
      var cpfTratadoHistorico = "'" + dadosLinha[1].toString().replace(/\D/g, "");
      
      var dataSolicitacao = dadosLinha[7]; 
      var dataConclusao = new Date();     
      var tempoTotalDias = 0;
      
      if (dataSolicitacao instanceof Date) {
        var d1 = new Date(dataSolicitacao.getFullYear(), dataSolicitacao.getMonth(), dataSolicitacao.getDate());
        var d2 = new Date(dataConclusao.getFullYear(), dataConclusao.getMonth(), dataConclusao.getDate());
        var diferencaEmMilissegundos = d2.getTime() - d1.getTime();
        tempoTotalDias = Math.floor(diferencaEmMilissegundos / (1000 * 60 * 60 * 24));
        if (tempoTotalDias < 0) tempoTotalDias = 0; 
      }
      
      if (abaHistorico) {
        var linhaHistorico = [
          dadosLinha[0], cpfTratadoHistorico, dadosLinha[2], dadosLinha[3], dadosLinha[4], 
          dadosLinha[5], dadosLinha[6], dataSolicitacao, dadosLinha[8], dadosLinha[9], 
          sisregTexto, dataConclusao, decisaoStatus, tempoTotalDias
        ];
        abaHistorico.appendRow(linhaHistorico);
      }
      
      try {
        var ssEquipes = SpreadsheetApp.openById(ID_PLANILHA_EQUIPES);
        var abaMeusEncaminhamentos = ssEquipes.getSheetByName("Meus_Encaminhamentos");
        
        if (abaMeusEncaminhamentos) {
          var dadosEquipes = abaMeusEncaminhamentos.getDataRange().getValues();
          var linhaEquipeAlvo = -1;
          
          for (var j = 1; j < dadosEquipes.length; j++) {
            if (dadosEquipes[j][9] && dadosEquipes[j][9].toString().trim() === codigoIdentificador.toString().trim()) {
              linhaEquipeAlvo = j + 1;
              break;
            }
          }
          
          if (linhaEquipeAlvo !== -1) {
            if (dadosEquipes.length <= 2) {
              abaMeusEncaminhamentos.getRange(linhaEquipeAlvo, 1, 1, abaMeusEncaminhamentos.getLastColumn()).clearContent();
            } else {
              abaMeusEncaminhamentos.deleteRow(linhaEquipeAlvo);
            }
          }
        }
      } catch (erroEquipes) {
        Logger.log("Aviso: Não foi possível limpar a aba das equipes. " + erroEquipes.message);
      }
      
      if (totalLinhasNaTabela <= 2) {
        abaRegular.getRange(linhaAlvo, 1, 1, abaRegular.getLastColumn()).clearContent();
      } else {
        abaRegular.deleteRow(linhaAlvo);
      }
      
      return "Sucesso! Registro arquivado com o código: " + sisregTexto;
    }
    
    return "Erro: Registro não localizado na fila ativa.";
  } catch(err) {
    return "Erro crítico no servidor: " + err.message;
  }
}

function obterDadosDashboard() {
  var resultado = {
    contadores: { emergencia: 0, urgente: 0, prioritario: 0, eletivo: 0, total: 0 },
    pacientes: [],
    profissionais: [] 
  };
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaRegular = ss.getSheetByName("Regular");
  
  if (!abaRegular) return resultado;
  
  var dados = abaRegular.getDataRange().getValues();
  if (dados.length <= 1) return resultado; 
  
  var listaProfissionaisSet = {}; 
  
  for (var i = 1; i < dados.length; i++) {
    var linha = dados[i];
    
    var profSolicitante = linha[0] ? linha[0].toString().trim().toUpperCase() : "NÃO INFORMADO"; // Coluna A
    var cpf             = linha[1]; // Coluna B
    var paciente        = linha[2]; // Coluna C
    var especialid      = linha[3]; // Coluna D
    var motivo          = linha[5]; // Coluna F
    var riscoBruto      = linha[6]; // Coluna G
    var dataEntrada     = linha[7]; // Coluna H
    var origem          = linha[8]; // Coluna I
    
    if (!paciente || paciente.toString().trim() === "") continue;

    var riscoTexto = riscoBruto ? riscoBruto.toString().trim().toUpperCase() : "";
    var cat = "eletivo";

    if (riscoTexto.indexOf("EMERG") !== -1 || riscoTexto.indexOf("VERMELH") !== -1) {
      cat = "emergencia";
      resultado.contadores.emergencia++;
    } else if (riscoTexto.indexOf("URG") !== -1 || riscoTexto.indexOf("AMAREL") !== -1) {
      cat = "urgente";
      resultado.contadores.urgente++;
    } else if (riscoTexto.indexOf("PRIOR") !== -1 || riscoTexto.indexOf("VERD") !== -1) {
      cat = "prioritario";
      resultado.contadores.prioritario++;
    } else {
      resultado.contadores.eletivo++;
    }

    resultado.contadores.total++;

    var dataFormatada = "-";
    if (dataEntrada) {
      try {
        dataFormatada = (dataEntrada instanceof Date) 
          ? Utilities.formatDate(dataEntrada, Session.getScriptTimeZone(), "dd/MM/yyyy") 
          : dataEntrada.toString();
      } catch (e) {
        dataFormatada = dataEntrada.toString();
      }
    }

    if (profSolicitante !== "NÃO INFORMADO" && profSolicitante !== "") {
      listaProfissionaisSet[profSolicitante] = true;
    }

    resultado.pacientes.push({
      data: dataFormatada,
      nome: paciente.toString(),
      cpf: cpf ? cpf.toString() : "-",
      risco: riscoBruto ? riscoBruto.toString() : "-",
      categoria: cat,
      origem: origem ? origem.toString() : "-",
      especialidade: especialid ? especialid.toString() : "-",
      motivo: motivo ? motivo.toString() : "-",
      profSolicitante: profSolicitante 
    });
  }

  resultado.profissionais = Object.keys(listaProfissionaisSet).sort();

  return resultado;
}
