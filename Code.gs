// Arquivo: Code.gs (Versão 6.0 - Base Refatorada)
// Foco: Clareza, Manutenibilidade e base para os princípios SOLID.

const CONFIG = Object.freeze({
  SHEET_ROTAS: "Rotas",
  SHEET_ALUNOS: "Alunos",
  SHEET_MONITORES: "Monitores",
  SHEET_FREQUENCIA: "Frequencia",
  SHEET_PONTO_FACULTATIVO: "Ponto_Facultativo",
  SHEET_REPOSICAO: "Reposicao",
  SHEET_ATIVIDADE_EXTRA: "Atividade_Extracurricular",
  CACHE_KEY: 'SIGTE_WEBAPP_DATA_V22_REFACTORED', // Chave de cache atualizada
  CACHE_EXPIRATION_SECONDS: 3600 // 1 hora
});

/**
 * Função principal que serve o Web App quando um usuário acessa a URL.
 * @param {object} e O objeto de evento do Apps Script.
 * @returns {HtmlOutput} O HTML renderizado da página principal.
 */
function doGet(e) {
  try {
    let template = HtmlService.createTemplateFromFile('Index');
    let htmlOutput = template.evaluate();
    htmlOutput
      .setTitle('SIG-TE - Painel de Gestão (v22.0 - Refatorado)')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    return htmlOutput;
  } catch (err) {
    Logger.log(`[doGet] ERRO CRÍTICO: ${err.message}\n${err.stack}`);
    return HtmlService.createHtmlOutput(`<h2>Erro Crítico ao Carregar o Sistema.</h2><p>Por favor, contate o administrador. Detalhes: ${err.message}</p>`);
  }
}

/**
 * Inclui o conteúdo de outros arquivos HTML no template principal (ex: CSS, JavaScript).
 * @param {string} filename O nome do arquivo a ser incluído.
 * @returns {string} O conteúdo do arquivo.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Ponto de entrada de dados para o Web App. Busca todos os dados necessários das planilhas.
 * Esta função intencionalmente ignora o cache para garantir que o botão "Atualizar" do frontend
 * sempre receba os dados mais recentes.
 * @returns {object} Um objeto com o status do sucesso e os dados.
 */
function getDadosParaWebApp() {
  try {
    const dados = _getDadosDasPlanilhas();
    if (!dados || Object.keys(dados).length === 0) {
      throw new Error("Nenhuma escola encontrada. Verifique se a aba 'Rotas' está preenchida corretamente.");
    }
    return { success: true, data: dados };
  } catch (err) {
    return handleError_('getDadosParaWebApp', err);
  }
}

/**
 * Processa um lote de submissões vindas do frontend.
 * Isso melhora a performance ao reduzir o número de chamadas entre cliente e servidor.
 * @param {Array<object>} lote Um array de objetos de submissão.
 * @returns {object} Um objeto com o resultado geral do processamento do lote.
 */
function processarLoteDeSubmissoes(lote) {
  if (!lote || !Array.isArray(lote) || lote.length === 0) {
    return { success: false, message: "Lote de submissão inválido ou vazio.", results: [] };
  }

  const results = [];
  let successCount = 0;

  lote.forEach((submissao, idx) => {
    let resultado;
    const submissionId = submissao.submissionId || `idx_${idx}`;
    try {
      switch (submissao.type) {
        case 'frequencia': resultado = processarFrequenciaWebApp(submissao.data); break;
        case 'pontoFacultativo': resultado = processarPontoFacultativo(submissao.data); break;
        case 'reposicao': resultado = processarReposicao(submissao.data); break;
        case 'atividadeExtra': resultado = processarAtividadeExtra(submissao.data); break;
        case 'adicionarAluno': resultado = adicionarAluno(submissao.data); break;
        case 'excluirAluno': resultado = excluirAluno(submissao.data); break;
        case 'adicionarRota': resultado = adicionarRota(submissao.data); break;
        case 'excluirRota': resultado = excluirRota(submissao.data); break;
        default: throw new Error(`Tipo de submissão desconhecido: ${submissao.type}`);
      }
      if (resultado.success) successCount++;
      results.push({ submissionId, ...resultado });
    } catch (err) {
      results.push({ submissionId, success: false, message: `Falha na submissão tipo '${submissao.type}': ${err.message}` });
    }
  });
  
  // Se qualquer operação de escrita foi bem-sucedida, invalida o cache para outros usuários.
  if (successCount > 0) {
    _invalidarCache();
  }
  
  const message = `Processamento do lote concluído. ${successCount} de ${lote.length} operações bem-sucedidas.`;
  return { success: successCount > 0, message, results };
}

// --- FUNÇÕES DE ESCRITA (PROCESSAMENTO DE DADOS) ---

function processarFrequenciaWebApp(dados) {
  _getSheet(CONFIG.SHEET_FREQUENCIA).appendRow([
    new Date(),
    dados.data,
    dados.idRota,
    dados.monitor.id,
    Session.getActiveUser().getEmail(),
    (dados.presentes || []).length,
    (dados.ausentes || []).length,
    (dados.presentes || []).map(a => `${a.nome} (${a.id})`).join(', '),
    (dados.ausentes || []).map(a => `${a.nome} (${a.id})`).join(', ')
  ]);
  return { success: true, message: "Frequência registrada com sucesso!" };
}

function processarPontoFacultativo(dados) {
  _getSheet(CONFIG.SHEET_PONTO_FACULTATIVO).appendRow([new Date(), dados.nomeEscola, dados.data, Session.getActiveUser().getEmail()]);
  return { success: true, message: "Ponto Facultativo registrado com sucesso!" };
}

function processarReposicao(dados) {
  _getSheet(CONFIG.SHEET_REPOSICAO).appendRow([new Date(), dados.nomeEscola, dados.data, dados.qtdOnibus, Session.getActiveUser().getEmail()]);
  return { success: true, message: "Reposição registrada com sucesso!" };
}

function processarAtividadeExtra(dados) {
  _getSheet(CONFIG.SHEET_ATIVIDADE_EXTRA).appendRow([new Date(), dados.nomeEscola, dados.data, dados.descricao, dados.qtdAlunos, dados.linkItinerario || '', Session.getActiveUser().getEmail(), dados.destino]);
  return { success: true, message: "Atividade Extracurricular registrada com sucesso!" };
}

function adicionarAluno(dados) {
  if (!dados || !dados.nomeCompleto || !dados.idRota) throw new Error("Nome e Rota são obrigatórios.");
  const sheet = _getSheet(CONFIG.SHEET_ALUNOS);
  sheet.appendRow([_generateUniqueId("ALU"), dados.nomeCompleto, dados.cpf || '', dados.idRota]);
  return { success: true, message: `Aluno "${dados.nomeCompleto}" adicionado.` };
}

function excluirAluno(dados) {
  if (!dados || !dados.idAluno) throw new Error("ID do aluno não fornecido.");
  const sheet = _getSheet(CONFIG.SHEET_ALUNOS);
  const rowIndex = _findRowIndexById(sheet, dados.idAluno, "ID_Aluno");
  if (rowIndex > -1) {
    sheet.deleteRow(rowIndex);
    return { success: true, message: `Aluno ID "${dados.idAluno}" excluído.` };
  }
  throw new Error(`Aluno ID "${dados.idAluno}" não encontrado.`);
}

function adicionarRota(dados) {
  if (!dados || !dados.nomeEscola || !dados.itinerario || !dados.turno || !dados.idMonitor || !Array.isArray(dados.diasSemana) || dados.diasSemana.length === 0) {
    throw new Error("Todos os campos da rota, incluindo dias da semana, são obrigatórios.");
  }
  const sheet = _getSheet(CONFIG.SHEET_ROTAS);
  sheet.appendRow([_generateUniqueId("ROTA"), dados.nomeEscola, dados.turno, dados.itinerario, dados.idMonitor, dados.diasSemana.join(', ')]);
  return { success: true, message: `Rota "${dados.itinerario}" adicionada.` };
}

function excluirRota(dados) {
  if (!dados || !dados.idRota) throw new Error("ID da rota não fornecido.");
  const sheet = _getSheet(CONFIG.SHEET_ROTAS);
  const rowIndex = _findRowIndexById(sheet, dados.idRota, "ID_Rota");
  if (rowIndex > -1) {
    sheet.deleteRow(rowIndex);
    return { success: true, message: `Rota ID "${dados.idRota}" excluída.` };
  }
  throw new Error(`Rota ID "${dados.idRota}" não encontrada.`);
}


// --- FUNÇÕES INTERNAS E AUXILIARES ---

/**
 * Orquestra a busca de dados de todas as planilhas e os estrutura em um único objeto.
 * @returns {object} O objeto de dados consolidado para o Web App.
 */
function _getDadosDasPlanilhas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const monitoresMap = _getMonitoresMap(ss);
  const alunosPorRota = _getAlunosPorRotaMap(ss);
  
  const dataRotas = _getSheetData(_getSheet(CONFIG.SHEET_ROTAS));
  const monitorPadrao = { id: 'N/A', nome: 'Monitor não atribuído' };
  
  const escolas = dataRotas.reduce((acc, row) => {
    if (row.ID_Rota && row.Nome_Escola) {
      const monitor = monitoresMap.get(String(row.ID_Monitor)) || monitorPadrao;
      const rotaInfo = { 
        id: row.ID_Rota, 
        nome: row.Itinerario_Padrao || 'Sem nome', 
        turno: row.Turno || 'N/A', 
        monitor: monitor, 
        alunos: alunosPorRota[row.ID_Rota] || [],
        diasSemana: row.Dias_da_Semana || ''
      };
      (acc[row.Nome_Escola] = acc[row.Nome_Escola] || { rotas: [] }).rotas.push(rotaInfo);
    }
    return acc;
  }, {});

  // Adiciona a lista de monitores ao objeto final para uso nos formulários.
  escolas.monitores = Array.from(monitoresMap.values());
  return escolas;
}

function _invalidarCache() {
  try {
    CacheService.getScriptCache().remove(CONFIG.CACHE_KEY);
    Logger.log('Cache do sistema invalidado com sucesso.');
  } catch (e) {
    Logger.log(`Falha ao invalidar o cache: ${e.message}`);
  }
}

function _getSheet(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`A aba obrigatória "${sheetName}" não foi encontrada.`);
  return sheet;
}

function _getSheetData(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => h.toString().trim().replace(/\s+/g, '_'));
  return values.slice(1).map(row => headers.reduce((obj, header, i) => { obj[header] = row[i]; return obj; }, {}));
}

function _getMonitoresMap(ss) {
  const data = _getSheetData(_getSheet(CONFIG.SHEET_MONITORES));
  return data.reduce((map, row) => {
    if (row.ID_Monitor && row.Nome_Monitor) {
      map.set(String(row.ID_Monitor), { id: row.ID_Monitor, nome: row.Nome_Monitor });
    }
    return map;
  }, new Map());
}

function _getAlunosPorRotaMap(ss) {
  const data = _getSheetData(_getSheet(CONFIG.SHEET_ALUNOS));
  return data.reduce((acc, row) => {
    if (row.ID_Rota && row.ID_Aluno && row.Nome_Completo) {
      (acc[row.ID_Rota] = acc[row.ID_Rota] || []).push({ id: row.ID_Aluno, nome: row.Nome_Completo, cpf: row.CPF || '' });
    }
    return acc;
  }, {});
}

function _generateUniqueId(prefix) {
  return `${prefix}${new Date().getTime()}${Math.floor(Math.random() * 1000)}`;
}

function _findRowIndexById(sheet, idToFind, idColumnName) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim().replace(/\s+/g, '_'));
  const idColumnIndex = headers.indexOf(idColumnName);
  if (idColumnIndex === -1) {
    throw new Error(`Coluna de ID "${idColumnName}" não encontrada na aba "${sheet.getName()}".`);
  }
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idColumnIndex]) === String(idToFind)) {
      return i + 1; // Retorna o índice da linha na planilha (base 1)
    }
  }
  return -1;
}

function handleError_(functionName, error) {
  const errorMessage = `[ERRO] na função ${functionName}: ${error.message}`;
  Logger.log(`${errorMessage}\nStack: ${error.stack}`);
  return { success: false, message: `Ocorreu um erro no servidor. Detalhes: ${error.message}` };
}


// --- FUNÇÕES DO MENU DA PLANILHA E BUSCA ---

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SIG-TE')
    .addItem('🔎 Localizar Conteúdo...', 'mostrarDialogoBusca')
    .addSeparator()
    .addItem('🔗 Obter URL do Web App', 'mostrarUrlDoWebApp')
    .addItem('🧹 Limpar Cache do Sistema', 'limparCacheManualmente')
    .addItem('✔️ Verificar Estrutura da Planilha', 'verificarEstruturaPlanilha')
    .addToUi();
}

function limparCacheManualmente() {
  _invalidarCache();
  SpreadsheetApp.getUi().alert('Sucesso', 'O cache do sistema foi limpo. Os usuários do Web App verão os dados atualizados na próxima vez que o abrirem ou atualizarem.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function mostrarDialogoBusca() {
  const html = HtmlService.createTemplateFromFile('SearchDialog').evaluate().setWidth(800).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Localizar Conteúdo - SIG-TE');
}

function mostrarUrlDoWebApp() {
  const url = ScriptApp.getService().getUrl();
  if (url) {
    const htmlOutput = HtmlService.createHtmlOutput(
      `A URL do seu aplicativo é:<br><br><input type="text" value="${url}" style="width: 95%;" onclick="this.select();" readonly><br><br><a href="${url}" target="_blank">Abrir Web App</a>`
    ).setWidth(400).setHeight(120);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'URL do Web App');
  } else {
    SpreadsheetApp.getUi().alert('Aviso', 'O script ainda não foi publicado como um Web App.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function verificarEstruturaPlanilha() {
  const requiredSheets = Object.values(CONFIG).filter(v => typeof v === 'string' && v.startsWith('SHEET_') === false);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = ss.getSheets().map(s => s.getName());
  const missingSheets = requiredSheets.filter(name => !sheetNames.includes(name));
  
  if (missingSheets.length > 0) {
    SpreadsheetApp.getUi().alert('ERRO DE ESTRUTURA', `As seguintes abas obrigatórias não foram encontradas:\n\n- ${missingSheets.join('\n- ')}`, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('SUCESSO', 'A estrutura da planilha parece correta. Todas as abas necessárias foram encontradas.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function buscarEmTodaPlanilha(options) {
  try {
    let { termo, caseSensitive, wholeWords, regexSearch } = options;
    if (!termo || termo.length < 2) return { success: true, data: [] };
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const results = [];
    
    if (wholeWords && !regexSearch) {
      termo = `\\b${termo.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`;
      regexSearch = true;
    }
    
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      let textFinder = sheet.createTextFinder(termo);
      if (regexSearch) {
        textFinder.useRegularExpression(true);
      } else {
        textFinder.matchCase(caseSensitive);
      }
      const foundRanges = textFinder.findAll();
      foundRanges.forEach(range => {
        results.push({
          aba: sheetName,
          celula: range.getA1Notation(),
          valor: range.getValue().toString()
        });
      });
    });
    return { success: true, data: results };
  } catch (err) {
    return handleError_('buscarEmTodaPlanilha', err);
  }
}

function navegarParaCelula(sheetName, cellA1) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const range = sheet.getRange(cellA1);
      ss.setActiveSheet(sheet);
      ss.setActiveRange(range);
    }
  } catch (err) {
    Logger.log(`Falha ao navegar para ${sheetName}!${cellA1}: ${err.message}`);
  }
}