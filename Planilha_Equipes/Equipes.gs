var CACHE_ENVIO = CacheService.getScriptCache();

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🏥 Sistema USF')
    .addItem('Abrir Painel de Consulta', 'abrirPainelEquipes')
    .addItem('➕ Nova Solicitação', 'abrirModalInclusao')
    .addToUi();
}

function abrirPainelEquipes() {
  var html = HtmlService.createHtmlOutputFromFile('PainelEquipes')
      .setTitle('Consulta de Situação');
  SpreadsheetApp.getUi().showSidebar(html);
}

function abrirModalInclusao() {
  var html = HtmlService.createHtmlOutputFromFile('ModalInclusao')
      .setWidth(620)
      .setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, '🏥 SUS - Formulário de Encaminhamento');
}

function buscarProfissionais() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("Cadastro_Profissionais");
    if (!aba) return [];
    
    var emailUsuarioLogado = Session.getActiveUser().getEmail().toString().trim().toLowerCase();
    
    var dados = aba.getDataRange().getValues();
    var lista = [];
    
    for (var i = 1; i < dados.length; i++) {
      if (!dados[i][0] || !dados[i][1]) continue;
      
      var emailCadastrado = dados[i][3] ? dados[i][3].toString().trim().toLowerCase() : "";
      
      if (emailUsuarioLogado !== "" && emailCadastrado === emailUsuarioLogado) {
        lista.push({
          equipe: dados[i][0].toString().trim(),
          nome: dados[i][1].toString().trim(),
          conselho: dados[i][2] ? dados[i][2].toString().trim() : "",
          email: emailCadastrado
        });
      }
    }
    
    if (lista.length === 0) {
      Logger.log("Aviso: E-mail logado ('" + emailUsuarioLogado + "') não encontrado ou inacessível. Retornando lista completa.");
      for (var j = 1; j < dados.length; j++) {
        if (!dados[j][0] || !dados[j][1]) continue;
        lista.push({
          equipe: dados[j][0].toString().trim(),
          nome: dados[j][1].toString().trim(),
          conselho: dados[j][2] ? dados[j][2].toString().trim() : "",
          email: dados[j][3] ? dados[j][3].toString().trim().toLowerCase() : ""
        });
      }
    }
    
    return lista;
  } catch(e) {
    Logger.log("Erro ao buscar profissionais: " + e.message);
    return [];
  }
}

function processarInclusaoFormulario(dadosForm) {
  try {
    var ssLocal = SpreadsheetApp.getActiveSpreadsheet();
    var nomeAbaAtiva = ssLocal.getActiveSheet().getName(); 
    
    var emailUsuarioLogado = Session.getActiveUser().getEmail();
    var emailFinal = (emailUsuarioLogado && emailUsuarioLogado !== "") ? emailUsuarioLogado : (dadosForm.email || "não.informado@gmail.com");

    var responsavelCompleto = dadosForm.responsavel ? dadosForm.responsavel.toUpperCase() : "";
    if (dadosForm.conselho && dadosForm.conselho.trim() !== "") {
      responsavelCompleto += " (" + dadosForm.conselho.toUpperCase() + ")";
    }

    var pacoteDados = {
      linhaOrigem: null, 
      nomeAba: nomeAbaAtiva,
      responsavel: responsavelCompleto,
      emailSolicitante: emailFinal,
      cpf: dadosForm.cpf ? dadosForm.cpf.replace(/\D/g, "") : "",
      nome: dadosForm.nome ? dadosForm.nome.toUpperCase() : "",
      espec: dadosForm.especialidade ? dadosForm.especialidade.toUpperCase() : "",
      cid: dadosForm.cid ? dadosForm.cid.toUpperCase() : "",
      motivo: dadosForm.motivo || "",
      grauRisco: dadosForm.risco ? dadosForm.risco.toUpperCase() : "",
      dataInc: new Date()
    };

    if (!pacoteDados.responsavel) {
      return { status: "erro", msg: "⚠️ Selecione o Profissional Responsável!" };
    }
    if (!pacoteDados.cpf || pacoteDados.cpf.length !== 11) {
      return { status: "erro", msg: "⚠️ Informe um CPF válido com 11 dígitos!" };
    }
    if (!pacoteDados.nome) {
      return { status: "erro", msg: "⚠️ Nome do Paciente é obrigatório!" };
    }

    var infoDuplicado = checarDuplicidadeBancos(pacoteDados.cpf, pacoteDados.espec);
    
    if (infoDuplicado.existe) {
      CACHE_ENVIO.put("dados_pendentes", JSON.stringify(pacoteDados), 300);
      
      var template = HtmlService.createTemplateFromFile('JanelaDuplicidade');
      template.dados = infoDuplicado;
      var html = template.evaluate().setWidth(500).setHeight(530);
      SpreadsheetApp.getUi().showModalDialog(html, '🚨 Alerta de Registro Duplicado');
      
      return { status: "duplicado", msg: "🚨 Duplicidade detectada! Verifique a janela de alerta." };
    } else {
      executarEnvioDireto(pacoteDados, "NOVA SOLICITACAO");
      return { status: "sucesso", msg: "✅ Demanda cadastrada com sucesso por " + pacoteDados.responsavel + "!" };
    }
  } catch(e) {
    return { status: "erro", msg: "❌ Erro no servidor: " + e.message };
  }
}

function checarDuplicidadeBancos(cpfRaw, especRaw) {
  var apenasNumerosBusca = cpfRaw.toString().replace(/\D/g, "");
  var especBusca = especRaw.toString().trim().toUpperCase();
  var ssLocal = SpreadsheetApp.getActiveSpreadsheet();
  var ID_REGULACAO = "18NpAeC98GNl-lGEK_6HJBolMo74n5kVsO8Z0EK9GdNE";
  
  try {
    var ssReg = SpreadsheetApp.openById(ID_REGULACAO);
    var abaRegular = ssReg.getSheetByName("Regular");
    if (abaRegular) {
      var dadosReg = abaRegular.getDataRange().getValues();
      for (var i = 1; i < dadosReg.length; i++) {
        if (!dadosReg[i][1] || !dadosReg[i][3]) continue;
        var cpfPlanilha = dadosReg[i][1].toString().replace(/\D/g, "");
        var especPlanilha = dadosReg[i][3].toString().trim().toUpperCase();
        
        if (cpfPlanilha === apenasNumerosBusca && especPlanilha === especBusca) {
          var dt = dadosReg[i][7] instanceof Date ? Utilities.formatDate(dadosReg[i][7], Session.getScriptTimeZone(), "dd/MM/yyyy") : dadosReg[i][7];
          return {
            existe: true, local: "Planilha de Regulação (Fila Geral)", nome: dadosReg[i][2], cpf: cpfRaw,
            espec: dadosReg[i][3], cid: dadosReg[i][4], motivo: dadosReg[i][5], data: dt, origem: dadosReg[i][8], sisreg: "Aguardando Regulador"
          };
        }
      }
    }
  } catch(e) {
    Logger.log("Aviso ao checar banco externo: " + e.message);
  }

  var abaLocal = ssLocal.getSheetByName("Meus_Encaminhamentos");
  if (abaLocal) {
    var dadosLoc = abaLocal.getDataRange().getValues();
    for (var j = 1; j < dadosLoc.length; j++) {
      if (!dadosLoc[j][1] || !dadosLoc[j][3]) continue;
      var cpfLoc = dadosLoc[j][1].toString().replace(/\D/g, "");
      var especLoc = dadosLoc[j][3].toString().trim().toUpperCase();
      
      if (cpfLoc === apenasNumerosBusca && especLoc === especBusca) {
        var dtL = dadosLoc[j][7] instanceof Date ? Utilities.formatDate(dadosLoc[j][7], Session.getScriptTimeZone(), "dd/MM/yyyy") : dadosLoc[j][7];
        return {
          existe: true, local: "Aba Local (Meus Encaminhamentos)", nome: dadosLoc[j][2], cpf: cpfRaw,
          espec: dadosLoc[j][3], cid: dadosLoc[j][4], motivo: dadosLoc[j][5], data: dtL, origem: dadosLoc[j][8], sisreg: "Pendente na sua fila"
        };
      }
    }
  }
  return { existe: false };
}

function processarDecisaoHTML(tipoStatus) {
  var cache = CacheService.getScriptCache();
  var stringDados = cache.get("dados_pendentes");
  if (!stringDados) return;
  var pacoteDados = JSON.parse(stringDados);
  executarEnvioDireto(pacoteDados, tipoStatus);
  cache.remove("dados_pendentes");
}

function executarEnvioDireto(dados, tipoStatus) {
  var trava = LockService.getScriptLock();
  try {
    trava.waitLock(30000); 
  } catch (e) {
    throw new Error("O sistema estava muito ocupado processando outros envios. Por favor, tente gravar novamente em alguns instantes.");
  }

  try {
    var ssLocal = SpreadsheetApp.getActiveSpreadsheet();
    var abaMeusEncaminhamentos = ssLocal.getSheetByName("Meus_Encaminhamentos");
    
    var ID_REGULACAO = "18NpAeC98GNl-lGEK_6HJBolMo74n5kVsO8Z0EK9GdNE"; 
    var ssRegulacao = SpreadsheetApp.openById(ID_REGULACAO);
    var abaRegular = ssRegulacao.getSheetByName("Regular"); 
    
    if (!abaRegular || !abaMeusEncaminhamentos) {
      throw new Error("Abas de destino não encontradas.");
    }

    var codigoIdentificador = "REG-" + Math.floor(100000 + Math.random() * 900000);
    var dataFormatada = new Date(dados.dataInc);
    if (isNaN(dataFormatada.getTime())) { dataFormatada = new Date(); }

    var cpfProtegido = "'" + dados.cpf.toString().replace(/\D/g, ""); 

    var novaLinhaReg = [
      dados.responsavel, cpfProtegido, dados.nome, dados.espec, dados.cid, dados.motivo, dados.grauRisco, dataFormatada, dados.nomeAba, codigoIdentificador
    ];
    abaRegular.appendRow(novaLinhaReg);
    
    var novaLinhaLocal = [
      dados.responsavel, cpfProtegido, dados.nome, dados.espec, dados.cid, dados.motivo, dados.grauRisco, dataFormatada, dados.nomeAba, codigoIdentificador, "", "", tipoStatus
    ];
    abaMeusEncaminhamentos.appendRow(novaLinhaLocal);
    
    ssLocal.toast("Encaminhamento gravado com sucesso!", "Sucesso", 4);

  } finally {
    trava.releaseLock();
  }
}

function buscarStatusRemoto(cpfRaw) {
  var retorno = { encontrado: false, nome: "", cpf: cpfRaw, status: "", dados: null };
  if (!cpfRaw) return JSON.stringify(retorno);
  var cpfLimpo = cpfRaw.toString().replace(/\D/g, "");
  if (cpfLimpo.length < 11) return JSON.stringify(retorno);

  var ID_PLANILHA_REGULACAO = "18NpAeC98GNl-lGEK_6HJBolMo74n5kVsO8Z0EK9GdNE";

  try {
    var ssReg = SpreadsheetApp.openById(ID_PLANILHA_REGULACAO);
    var abaRegular = ssReg.getSheetByName("Regular");
    var abaHistorico = ssReg.getSheetByName("Historico");

    if (abaRegular) {
      var dadosReg = abaRegular.getDataRange().getValues();
      for (var i = 1; i < dadosReg.length; i++) {
        if (!dadosReg[i][1]) continue;
        var cpfLinha = dadosReg[i][1].toString().replace(/\D/g, "");
        if (cpfLinha === cpfLimpo) {
          retorno.encontrado = true;
          retorno.nome = dadosReg[i][2];
          retorno.status = "FILA_ATIVA";
          retorno.dados = {
            profSolicitante: dadosReg[i][0] || "Não informado", 
            especialidade: dadosReg[i][3],                       
            cid: dadosReg[i][4],                                 
            motivo: dadosReg[i][5],                              
            risco: dadosReg[i][6],                               
            data: dadosReg[i][7] instanceof Date ? Utilities.formatDate(dadosReg[i][7], Session.getScriptTimeZone(), "dd/MM/yyyy") : dadosReg[i][7], 
            origemUnidade: dadosReg[i][8] || "Não informada"     
          };
          break;
        }
      }
    }

    if (!retorno.encontrado && abaHistorico) {
      var dadosHist = abaHistorico.getDataRange().getValues();
      for (var j = 1; j < dadosHist.length; j++) {
        if (!dadosHist[j][1]) continue;
        var cpfHist = dadosHist[j][1].toString().replace(/\D/g, "");
        if (cpfHist === cpfLimpo) {
          retorno.encontrado = true;
          retorno.nome = dadosHist[j][2];
          retorno.status = "CONCLUIDO";
          retorno.dados = {
            profSolicitante: dadosHist[j][0] || "Não informado", 
            especialidade: dadosHist[j][3],                       
            cid: dadosHist[j][4],                                 
            motivo: dadosHist[j][5],                              
            risco: dadosHist[j][6],                               
            data: dadosHist[j][7] instanceof Date ? Utilities.formatDate(dadosHist[j][7], Session.getScriptTimeZone(), "dd/MM/yyyy") : dadosHist[j][7], 
            origemUnidade: dadosHist[j][8] || "Não informada",    
            sisreg: dadosHist[j][10] || "Não informado",          
            dataConclusao: dadosHist[j][11] instanceof Date ? Utilities.formatDate(dadosHist[j][11], Session.getScriptTimeZone(), "dd/MM/yyyy") : dadosHist[j][11], 
            tempoEspera: dadosHist[j][13] !== undefined ? dadosHist[j][13] : "-" 
          };
          break;
        }
      }
    }

    if (retorno.encontrado) {
      retorno.cpf = cpfLimpo.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, "$1.$2.$3-$4");
    }
  } catch (err) {
    Logger.log("Erro na ligação remota: " + err.message);
  }
  return JSON.stringify(retorno);
}
