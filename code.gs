/**
 * Projeto: Automação de Cotações e Supply Chain
 * Backend: Google Apps Script
 */

// [ATENÇÃO] Substitua pela URL da sua planilha que servirá como Banco de Dados
const PLANILHA_URL = "[COLOQUE_A_URL_DA_SUA_PLANILHA_AQUI]";

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Cotação de Materiais')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function conectarBancoDados() {
  try { return SpreadsheetApp.openByUrl(PLANILHA_URL); } 
  catch (erro) { throw new Error("Falha na conexão com a planilha."); }
}

// ==========================================
// 1. FORNECEDORES (CRUD)
// ==========================================

function obterFornecedores() {
  try {
    const aba = conectarBancoDados().getSheetByName('Fornecedores');
    if (!aba || aba.getLastRow() <= 1) return [];
    
    const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 4).getValues();
    return dados.map(linha => ({ nome: linha[0], email: linha[1], tipoMaterial: linha[2], status: linha[3] || 'Ativo' }));
  } catch (erro) { return []; }
}

function adicionarFornecedor(dados) {
  conectarBancoDados().getSheetByName('Fornecedores').appendRow([dados.nome, dados.email, dados.tipoMaterial, dados.status]);
  return "Fornecedor adicionado com sucesso!";
}

function editarFornecedor(emailAntigo, dados) {
  const aba = conectarBancoDados().getSheetByName('Fornecedores');
  const dadosAba = aba.getDataRange().getValues();
  for (let i = 1; i < dadosAba.length; i++) {
    if (dadosAba[i][1] === emailAntigo) {
      aba.getRange(i + 1, 1, 1, 4).setValues([[dados.nome, dados.email, dados.tipoMaterial, dados.status]]);
      return "Fornecedor atualizado com sucesso!";
    }
  }
  throw new Error("Fornecedor não encontrado.");
}

function excluirFornecedor(email) {
  const aba = conectarBancoDados().getSheetByName('Fornecedores');
  const dadosAba = aba.getDataRange().getValues();
  for (let i = 1; i < dadosAba.length; i++) {
    if (dadosAba[i][1] === email) { aba.deleteRow(i + 1); return "Excluído com sucesso!"; }
  }
  throw new Error("Fornecedor não encontrado.");
}

// ==========================================
// 2. MATERIAIS / PLANILHAS PADRÕES (CRUD)
// ==========================================

function obterItens() {
  try {
    const aba = conectarBancoDados().getSheetByName('Itens_Cotacao');
    if (!aba || aba.getLastRow() <= 1) return [];
    
    const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 5).getValues();
    return dados.map(linha => ({
      material: linha[0],
      txtBreve: linha[1],
      tipoMaterial: linha[2],
      especificacao: linha[3],
      quantidade: linha[4]
    }));
  } catch (erro) { return []; }
}

function adicionarItem(dados) {
  conectarBancoDados().getSheetByName('Itens_Cotacao').appendRow([
    dados.material, dados.txtBreve, dados.tipoMaterial, dados.especificacao, dados.quantidade
  ]);
  return "Material adicionado com sucesso!";
}

function editarItem(materialAntigo, dados) {
  try {
    const aba = conectarBancoDados().getSheetByName('Itens_Cotacao');
    const dadosAba = aba.getDataRange().getValues();
    for (let i = 1; i < dadosAba.length; i++) {
      if (dadosAba[i][0].toString() === materialAntigo.toString()) {
        aba.getRange(i + 1, 1, 1, 5).setValues([[
          dados.material, dados.txtBreve, dados.tipoMaterial, dados.especificacao, dados.quantidade
        ]]);
        return "Material atualizado com sucesso!";
      }
    }
    throw new Error("Material não encontrado.");
  } catch (e) {
    throw new Error("Erro ao editar material: " + e.message);
  }
}

function excluirItem(materialId) {
  const aba = conectarBancoDados().getSheetByName('Itens_Cotacao');
  const dadosAba = aba.getDataRange().getValues();
  for (let i = 1; i < dadosAba.length; i++) {
    if (dadosAba[i][0].toString() === materialId.toString()) { 
      aba.deleteRow(i + 1); return "Material excluído!"; 
    }
  }
  throw new Error("Material não encontrado.");
}

// ==========================================
// 3. ENVIO DE E-MAILS COM EXCEL GERADO
// ==========================================

function dispararEmailsCotacao(fornecedoresSelecionados) {
  try {
    const planilha = conectarBancoDados();
    const abaMateriais = planilha.getSheetByName('Itens_Cotacao'); 
    if (!abaMateriais) throw new Error("Aba 'Itens_Cotacao' não encontrada.");
    
    const todosOsMateriais = abaMateriais.getDataRange().getValues();
    const cabeçalhoMateriais = todosOsMateriais[0];
    
    const abaFornecedores = planilha.getSheetByName('Fornecedores');
    const dadosFornecedores = abaFornecedores.getDataRange().getValues();

    let enviados = 0;

    fornecedoresSelecionados.forEach(emailDestino => {
      let tipoMaterialFornecedor = "";
      let nomeFornecedor = "";
      
      for (let i = 1; i < dadosFornecedores.length; i++) {
        if (dadosFornecedores[i][1] === emailDestino) {
          nomeFornecedor = dadosFornecedores[i][0];
          tipoMaterialFornecedor = dadosFornecedores[i][2];
          break;
        }
      }
      
      let itensFiltrados = [cabeçalhoMateriais]; 
      for (let j = 1; j < todosOsMateriais.length; j++) {
        if (todosOsMateriais[j][2] === tipoMaterialFornecedor) {
          let linha = todosOsMateriais[j];
          let linhaParaFornecedor = [
            linha[0], linha[1], linha[2], linha[3], linha[4], 
            "", "", "", nomeFornecedor, "", ""
          ];
          itensFiltrados.push(linhaParaFornecedor);
        }
      }

      if (itensFiltrados.length > 1) {
        // Cria arquivo temporário no Drive para gerar a planilha
        let planilhaTemp = SpreadsheetApp.create("Temp_Cotacao_" + nomeFornecedor);
        let abaTemp = planilhaTemp.getSheets()[0];
        abaTemp.getRange(1, 1, itensFiltrados.length, itensFiltrados[0].length).setValues(itensFiltrados);
        SpreadsheetApp.flush();

        // Converte para XLSX
        let urlConversao = "https://docs.google.com/spreadsheets/d/" + planilhaTemp.getId() + "/export?format=xlsx";
        let tokenAutenticacao = ScriptApp.getOAuthToken();
        let respostaConversao = UrlFetchApp.fetch(urlConversao, { headers: { 'Authorization': 'Bearer ' + tokenAutenticacao }, muteHttpExceptions: true });
        
        let anexoExcel = respostaConversao.getBlob().setName("Cotacao_" + tipoMaterialFornecedor + ".xlsx");

        MailApp.sendEmail({
          to: emailDestino,
          subject: `Cotação de Materiais - ${tipoMaterialFornecedor}`,
          htmlBody: `<p>Olá ${nomeFornecedor},</p><p>Segue em anexo a planilha de cotação. Por favor, preencha as colunas de valor, marca e prazo e faça o upload no nosso portal.</p>`,
          attachments: [anexoExcel]
        });

        // Limpa o lixo temporário do Drive
        DriveApp.getFileById(planilhaTemp.getId()).setTrashed(true);
        enviados++;
      }
    });

    return `Sucesso! Ficheiros Excel gerados e ${enviados} e-mails enviados.`;
  } catch (erro) {
    throw new Error("Erro ao gerar envios: " + erro.message);
  }
}

// ==========================================
// 4. RETORNOS E UPLOAD (BASE64 PARA CSV)
// ==========================================

function processarUpload(arquivoBase64, nomeArquivo, mesAno, rodada) {
  try {
    const abaRetornos = conectarBancoDados().getSheetByName('Retornos');
    const blob = Utilities.newBlob(Utilities.base64Decode(arquivoBase64), MimeType.CSV, nomeArquivo);
    const dadosCsv = Utilities.parseCsv(blob.getDataAsString(), ","); 
    
    if (dadosCsv.length === 0) throw new Error("Ficheiro vazio.");
    
    if (isNaN(dadosCsv[0][0]) || dadosCsv[0][0].toString().toLowerCase().includes("material")) {
      dadosCsv.shift(); 
    }
    
    const dataAtual = new Date();
    
    const linhasParaInserir = dadosCsv.map(linha => {
      let idMaterial = linha[0];
      let fornecedorReal = linha[8] ? linha[8].toString().trim() : "Fornecedor Desconhecido"; 
      let valorUnitario = linha[6];
      let valorTotal = linha[5];

      return [ idMaterial, fornecedorReal, valorUnitario, valorTotal, dataAtual, mesAno, rodada ];
    });
    
    if (linhasParaInserir.length > 0) {
      abaRetornos.getRange(abaRetornos.getLastRow() + 1, 1, linhasParaInserir.length, 7).setValues(linhasParaInserir);
    }
    
    return "Ficheiro importado com sucesso! Dados reais do fornecedor gravados.";
  } catch (e) { 
    throw new Error("Erro ao ler ficheiro: " + e.message); 
  }
}

// ==========================================
// 5. DASHBOARD E RELATÓRIOS
// ==========================================

function obterFiltrosDashboard() {
  try {
    const planilha = conectarBancoDados();
    const abaRetornos = planilha.getSheetByName('Retornos');
    const abaMateriais = planilha.getSheetByName('Itens_Cotacao');
    
    if (!abaRetornos || abaRetornos.getLastRow() <= 1) return { meses: [], materiais: [], rodadas: [] };

    const dadosRetornos = abaRetornos.getRange(2, 1, abaRetornos.getLastRow() - 1, 7).getValues();
    let mesesUnicos = [...new Set(dadosRetornos.map(linha => {
      let valorMes = linha[5];
      if (!valorMes) return "";
      if (valorMes instanceof Date) return Utilities.formatDate(valorMes, "GMT-3", "yyyy-MM");
      return valorMes.toString().trim();
    }))].filter(String);

    let rodadasUnicas = [...new Set(dadosRetornos.map(linha => linha[6] ? parseInt(linha[6]) : null))].filter(Boolean);

    let tiposUnicos = [];
    if (abaMateriais && abaMateriais.getLastRow() > 1) {
      const dadosMateriais = abaMateriais.getRange(2, 3, abaMateriais.getLastRow() - 1, 1).getValues();
      tiposUnicos = [...new Set(dadosMateriais.map(linha => linha[0] ? linha[0].toString().trim() : ""))].filter(String);
    }

    return { 
      meses: mesesUnicos.sort(), 
      materiais: tiposUnicos.sort(), 
      rodadas: rodadasUnicas.sort((a,b) => a-b) 
    };
  } catch (e) {
    return { meses: [], materiais: [], rodadas: [] };
  }
}

function processarDadosDashboard(mesAnoSelecionado, tipoMaterialSelecionado, rodadaSelecionada) {
  try {
    const planilha = conectarBancoDados();
    const abaRetornos = planilha.getSheetByName('Retornos');
    const abaMateriais = planilha.getSheetByName('Itens_Cotacao');
    
    if (!abaRetornos || abaRetornos.getLastRow() <= 1) return null;

    let mapaIdParaTipo = {};
    if (abaMateriais && abaMateriais.getLastRow() > 1) {
      const dadosMat = abaMateriais.getRange(2, 1, abaMateriais.getLastRow() - 1, 3).getValues();
      dadosMat.forEach(r => {
        mapaIdParaTipo[r[0].toString().trim()] = r[2] ? r[2].toString().trim() : "";
      });
    }

    const dados = abaRetornos.getRange(2, 1, abaRetornos.getLastRow() - 1, 7).getValues();
    
    let retornosFiltrados = dados.filter(linha => {
      let mes = linha[5];
      if (mes instanceof Date) mes = Utilities.formatDate(mes, "GMT-3", "yyyy-MM");
      else mes = mes ? mes.toString().trim() : "";
      
      let idMat = linha[0] ? linha[0].toString().trim() : "";
      let tipoDoMaterial = mapaIdParaTipo[idMat] || "";
      let rodadaLinha = linha[6] ? parseInt(linha[6]) : 1;
      
      let matchMes = mes === mesAnoSelecionado.toString().trim();
      let matchTipo = !tipoMaterialSelecionado || tipoDoMaterial === tipoMaterialSelecionado.toString().trim();
      let matchRodada = !rodadaSelecionada || rodadaLinha === parseInt(rodadaSelecionada);
      
      return matchMes && matchTipo && matchRodada;
    });

    if (retornosFiltrados.length === 0) return null;

    retornosFiltrados.forEach(linha => {
      let strPreco = linha[3] ? linha[3].toString().trim() : (linha[2] ? linha[2].toString().trim() : "0");
      if(strPreco.includes(',') && strPreco.includes('.')) {
          strPreco = strPreco.replace(/\./g, '').replace(',', '.');
      } else if (strPreco.includes(',')) {
          strPreco = strPreco.replace(',', '.');
      }
      linha.valorParaSoma = parseFloat(strPreco) || 0;
    });

    let mapaAgrupado = {};
    retornosFiltrados.forEach(linha => {
      let fornec = linha[1] || "Desconhecido";
      let rodada = parseInt(linha[6]) || 1;
      let chave = fornec + "_" + rodada;
      
      if (!mapaAgrupado[chave]) {
        mapaAgrupado[chave] = { fornecedor: fornec, rodada: rodada, precoTotal: 0 };
      }
      mapaAgrupado[chave].precoTotal += linha.valorParaSoma;
    });

    let listaAgrupada = Object.values(mapaAgrupado);

    let mapaExibicaoTabela = {};
    listaAgrupada.forEach(item => {
      if (!rodadaSelecionada) {
        if (!mapaExibicaoTabela[item.fornecedor] || item.rodada > mapaExibicaoTabela[item.fornecedor].rodada) {
          mapaExibicaoTabela[item.fornecedor] = item;
        }
      } else {
        mapaExibicaoTabela[item.fornecedor] = item;
      }
    });

    let tabelaDecisao = Object.values(mapaExibicaoTabela).sort((a, b) => a.precoTotal - b.precoTotal);
    if (tabelaDecisao.length === 0) return null;

    let menorPrecoGeral = tabelaDecisao[0].precoTotal;

    tabelaDecisao = tabelaDecisao.map(forn => {
      let dif = forn.precoTotal - menorPrecoGeral;
      forn.percReducao = forn.precoTotal > 0 ? (dif / forn.precoTotal) * 100 : 0;
      return forn;
    });

    let graficoEvolutivo = [];
    let dadosParaGraficoGeral = dados.filter(linha => {
      let mes = linha[5];
      if (mes instanceof Date) mes = Utilities.formatDate(mes, "GMT-3", "yyyy-MM");
      else mes = mes ? mes.toString().trim() : "";
      let idMat = linha[0] ? linha[0].toString().trim() : "";
      let tipoDoMaterial = mapaIdParaTipo[idMat] || "";
      return mes === mesAnoSelecionado.toString().trim() && (!tipoMaterialSelecionado || tipoDoMaterial === tipoMaterialSelecionado.toString().trim());
    });

    let mapaGraficoAgrupado = {};
    dadosParaGraficoGeral.forEach(linha => {
      let fornec = linha[1] || "Desconhecido";
      let rodada = parseInt(linha[6]) || 1;
      let strP = linha[3] ? linha[3].toString().trim() : (linha[2] ? linha[2].toString().trim() : "0");
      if(strP.includes(',') && strP.includes('.')) strP = strP.replace(/\./g, '').replace(',', '.');
      else if (strP.includes(',')) strP = strP.replace(',', '.');
      let val = parseFloat(strP) || 0;

      let chave = fornec + "_" + rodada;
      if (!mapaGraficoAgrupado[chave]) mapaGraficoAgrupado[chave] = { rodada: rodada, total: 0 };
      mapaGraficoAgrupado[chave].total += val;
    });

    let listaGrafico = Object.values(mapaGraficoAgrupado);
    let rodadasDoGrafico = [...new Set(listaGrafico.map(l => l.rodada))].sort((a,b) => a-b);
    
    rodadasDoGrafico.forEach(rNum => {
      let itensDaRodada = listaGrafico.filter(l => l.rodada === rNum);
      if(itensDaRodada.length > 0) {
        let menorPrecoDaRodada = Math.min(...itensDaRodada.map(l => l.total));
        graficoEvolutivo.push(["Rodada " + rNum, menorPrecoDaRodada]);
      }
    });

    return {
      kpis: { menorPreco: menorPrecoGeral, melhorFornecedor: tabelaDecisao[0].fornecedor },
      grafico: graficoEvolutivo,
      tabela: tabelaDecisao
    };

  } catch (e) {
    throw new Error(e.message);
  }
}