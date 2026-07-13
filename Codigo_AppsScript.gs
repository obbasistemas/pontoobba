/**
 * OBBA - SISTEMA DE PONTO ELETRÔNICO
 * Backend em Google Apps Script
 *
 * COMO INSTALAR:
 * 1. Crie uma planilha nova no Google Sheets (pode chamar "OBBA - Controle de Ponto")
 * 2. Menu Extensões > Apps Script
 * 3. Apague o conteúdo do Code.gs e cole todo este arquivo
 * 4. Rode a função "inicializarPlanilha" uma vez (menu Executar > inicializarPlanilha)
 *    - Isso vai criar as 4 abas (Funcionarios, Jornadas, RegistroPonto, Feriados)
 *      com cabeçalhos e um funcionário de exemplo (o ENOQUE, do PDF que você me mandou)
 * 5. Menu Implantar > Nova implantação > tipo "App da Web"
 *    - Executar como: Eu (seu e-mail)
 *    - Quem pode acessar: Qualquer pessoa
 * 6. Copie a URL gerada e cole no arquivo app_ponto.html, na constante APPS_SCRIPT_URL
 *
 * Depois de editar dados diretamente na planilha (novo funcionário, nova jornada,
 * feriados), não precisa reimplantar - é só editar a planilha mesmo.
 */

const TIMEZONE = 'America/Sao_Paulo'; // mesmo horário de São Luís - MA (UTC-3)
const EMPRESA_PADRAO = {
  nome: 'OBBA - SOLUCAO EM GESTAO CONDOMINIAL LTD',
  documento: '37.674.554/0002-65',
  endereco: 'JERONIMO DE ALBUQUERQUE MARANHAO, 25',
  bairro: 'VINHAIS I',
  cidade: 'SAO LUIS',
  uf: 'MA',
  cep: '65-074-199'
};

const ABA_FUNCIONARIOS = 'Funcionarios';
const ABA_JORNADAS = 'Jornadas';
const ABA_REGISTRO = 'RegistroPonto';
const ABA_FERIADOS = 'Feriados';
const ABA_ADMINISTRADORES = 'Administradores';
const CABECALHOS_FUNCIONARIOS = [
  'Codigo',
  'Nome',
  'PIN',
  'Funcao',
  'JornadaID',
  'Status',
  'Depto / Setor / Secao',
  'Endereco',
  'Bairro',
  'Cidade',
  'Estado',
  'CEP',
  'Hr/Mes',
  'CTPS',
  'CBO'
];

// ---------- SETUP ----------

function inicializarPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  criarAbaSeNaoExiste(ss, ABA_FUNCIONARIOS,
    CABECALHOS_FUNCIONARIOS);
  criarAbaSeNaoExiste(ss, ABA_JORNADAS,
    ['JornadaID', 'Descricao', 'SegSex_Entrada', 'SegSex_IntervaloInicio', 'SegSex_IntervaloFim', 'SegSex_Saida', 'Sab_Entrada', 'Sab_Saida', 'ToleranciaMin']);
  criarAbaSeNaoExiste(ss, ABA_REGISTRO,
    ['ID', 'Codigo', 'Nome', 'Data', 'Tipo', 'Hora', 'Latitude', 'Longitude', 'Timestamp']);
  criarAbaSeNaoExiste(ss, ABA_FERIADOS,
    ['Data', 'Descricao']);
  criarAbaSeNaoExiste(ss, ABA_ADMINISTRADORES,
    ['Usuario', 'Senha', 'Nome']);

  // Administrador de exemplo - TROQUE A SENHA depois de testar!
  const admins = ss.getSheetByName(ABA_ADMINISTRADORES);
  if (admins.getLastRow() < 2) {
    admins.appendRow(['admin', 'obba2026', 'Administração OBBA']);
  }

  garantirColunasFuncionariosExtras();

  // Jornada padrão, igual ao horário do modelo em PDF
  const jornadas = ss.getSheetByName(ABA_JORNADAS);
  if (jornadas.getLastRow() < 2) {
    jornadas.appendRow(['PADRAO', 'Seg-Sex 08-18 (int 12-14) / Sáb 08-12', '08:00', '12:00', '14:00', '18:00', '08:00', '12:00', 10]);
  }

  // Funcionário de exemplo, com base no PDF enviado
  const funcionarios = ss.getSheetByName(ABA_FUNCIONARIOS);
  if (funcionarios.getLastRow() < 2) {
    funcionarios.appendRow([
      '8',
      'ENOQUE SANCHES BORGES',
      '1234',
      'GERAL',
      'PADRAO',
      'Ativo',
      '1 - GERAL',
      'MA 703, 17',
      'ARACAGY',
      'SAO LUIS',
      'MA',
      '65-136-000',
      '220,00',
      '7666468/6287',
      '410105,000000'
    ]);
  }

  corrigirFormatoColunas();

  SpreadsheetApp.getUi().alert('Planilha inicializada! Abas criadas: Funcionarios, Jornadas, RegistroPonto, Feriados.');
}

// Trava as colunas Data e Hora da aba RegistroPonto como TEXTO PURO, pra o Sheets
// nunca mais converter sozinho "11/07/2026" ou "10:42:21" em valores de Data/Hora.
// Rode essa função manualmente se já tiver uma planilha criada antes dessa correção.
function corrigirFormatoColunas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_REGISTRO);
  aba.getRange('D:D').setNumberFormat('@'); // coluna Data
  aba.getRange('F:F').setNumberFormat('@'); // coluna Hora
}

function criarAbaSeNaoExiste(ss, nome, cabecalhos) {
  let aba = ss.getSheetByName(nome);
  if (!aba) {
    aba = ss.insertSheet(nome);
    aba.appendRow(cabecalhos);
    aba.setFrozenRows(1);
    aba.getRange(1, 1, 1, cabecalhos.length).setFontWeight('bold');
  }
  return aba;
}

function garantirColunasFuncionariosExtras() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_FUNCIONARIOS);
  if (!aba) return;

  const ultimaColuna = Math.max(aba.getLastColumn(), 1);
  const cabecalhos = aba.getRange(1, 1, 1, ultimaColuna).getValues()[0];
  const mapa = mapearCabecalhos(cabecalhos);

  for (let i = 0; i < CABECALHOS_FUNCIONARIOS.length; i++) {
    const nomeCabecalho = CABECALHOS_FUNCIONARIOS[i];
    if (mapa[normalizarCabecalho(nomeCabecalho)] === undefined) {
      aba.getRange(1, aba.getLastColumn() + 1).setValue(nomeCabecalho);
    }
  }

  aba.getRange(1, 1, 1, aba.getLastColumn()).setFontWeight('bold');
}

// ---------- ENTRY POINTS ----------

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    let resultado;

    if (action === 'login') {
      resultado = fazerLogin(body.codigo, body.pin);
    } else if (action === 'status') {
      resultado = getStatusDoDia(body.codigo);
    } else if (action === 'registrar') {
      resultado = registrarPonto(body.codigo, body.tipo, body.latitude, body.longitude);
    } else if (action === 'loginAdmin') {
      resultado = fazerLoginAdmin(body.usuario, body.senha);
    } else if (action === 'statusGeralHoje') {
      resultado = getStatusGeralHoje();
    } else if (action === 'listarFuncionarios') {
      resultado = listarFuncionarios();
    } else if (action === 'listarRegistros') {
      resultado = listarRegistros(body.codigo, body.dataInicio, body.dataFim);
    } else if (action === 'gerarFolhaMensal') {
      resultado = gerarFolhaMensal(body.codigo, body.competencia);
    } else if (action === 'editarRegistro') {
      resultado = editarRegistro(body.id, body.novaData, body.novoTipo, body.novaHora);
    } else if (action === 'excluirRegistro') {
      resultado = excluirRegistro(body.id);
    } else if (action === 'listarFeriados') {
      resultado = listarFeriados();
    } else if (action === 'adicionarFeriado') {
      resultado = adicionarFeriado(body.data, body.descricao);
    } else if (action === 'removerFeriado') {
      resultado = removerFeriado(body.data);
    } else {
      resultado = { ok: false, erro: 'Ação inválida.' };
    }

    return ContentService.createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, erro: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---------- LÓGICA ----------

function normalizarCabecalho(texto) {
  return String(texto || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

function mapearCabecalhos(cabecalhos) {
  const mapa = {};
  for (let i = 0; i < cabecalhos.length; i++) {
    mapa[normalizarCabecalho(cabecalhos[i])] = i;
  }
  return mapa;
}

function obterValorCabecalho(linha, mapa, nomes, padrao) {
  for (let i = 0; i < nomes.length; i++) {
    const idx = mapa[normalizarCabecalho(nomes[i])];
    if (idx !== undefined) {
      const valor = linha[idx];
      if (valor !== '' && valor !== null && valor !== undefined) return valor;
    }
  }
  return padrao || '';
}

function criarFuncionarioAPartirDaLinha(linha, mapa) {
  return {
    codigo: String(obterValorCabecalho(linha, mapa, ['Codigo'], linha[0])).trim(),
    nome: obterValorCabecalho(linha, mapa, ['Nome'], linha[1]),
    pin: String(obterValorCabecalho(linha, mapa, ['PIN'], linha[2])).trim(),
    funcao: obterValorCabecalho(linha, mapa, ['Funcao'], linha[3]),
    jornadaId: obterValorCabecalho(linha, mapa, ['JornadaID'], linha[4]),
    status: obterValorCabecalho(linha, mapa, ['Status'], linha[5]),
    departamento: obterValorCabecalho(linha, mapa, ['Depto / Setor / Secao', 'Depto', 'Departamento', 'Setor', 'Secao'], ''),
    endereco: obterValorCabecalho(linha, mapa, ['Endereco'], ''),
    bairro: obterValorCabecalho(linha, mapa, ['Bairro'], ''),
    cidade: obterValorCabecalho(linha, mapa, ['Cidade'], ''),
    estado: obterValorCabecalho(linha, mapa, ['Estado', 'UF'], ''),
    cep: obterValorCabecalho(linha, mapa, ['CEP'], ''),
    hrMes: obterValorCabecalho(linha, mapa, ['Hr/Mes', 'HrMes'], '220,00'),
    ctps: obterValorCabecalho(linha, mapa, ['CTPS'], ''),
    cbo: obterValorCabecalho(linha, mapa, ['CBO'], '')
  };
}

function listarFuncionariosCompletos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_FUNCIONARIOS);
  const dados = aba.getDataRange().getValues();
  if (dados.length < 2) return [];
  const mapa = mapearCabecalhos(dados[0]);
  const lista = [];
  for (let i = 1; i < dados.length; i++) {
    lista.push(criarFuncionarioAPartirDaLinha(dados[i], mapa));
  }
  return lista;
}

function getFuncionarioPorCodigo(codigo) {
  const lista = listarFuncionariosCompletos();
  for (let i = 0; i < lista.length; i++) {
    if (String(lista[i].codigo) === String(codigo)) {
      return lista[i];
    }
  }
  return null;
}

function getJornada(jornadaId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_JORNADAS);
  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(jornadaId)) {
      return {
        id: dados[i][0], descricao: dados[i][1],
        segSexEntrada: dados[i][2], segSexIntInicio: dados[i][3], segSexIntFim: dados[i][4], segSexSaida: dados[i][5],
        sabEntrada: dados[i][6], sabSaida: dados[i][7], toleranciaMin: dados[i][8]
      };
    }
  }
  return null;
}

function fazerLogin(codigo, pin) {
  const func = getFuncionarioPorCodigo(codigo);
  if (!func) return { ok: false, erro: 'Código de funcionário não encontrado.' };
  if (func.status !== 'Ativo') return { ok: false, erro: 'Funcionário inativo.' };
  if (func.pin !== String(pin)) return { ok: false, erro: 'PIN incorreto.' };
  return { ok: true, funcionario: { codigo: func.codigo, nome: func.nome, funcao: func.funcao } };
}

// Retorna a sequência de marcações esperadas para hoje, de acordo com o dia da semana e feriados
function sequenciaDoDia(diaSemana, dataStr) {
  if (dataStr && ehFeriado(dataStr)) return []; // feriado - folga, não importa o dia da semana
  // diaSemana: 0=domingo, 1=segunda ... 6=sábado
  if (diaSemana === 0) return []; // domingo - folga
  if (diaSemana === 6) return ['Entrada', 'Saída Final']; // sábado, sem intervalo
  return ['Entrada', 'Saída Almoço', 'Volta Almoço', 'Saída Final']; // seg-sex
}

function ehFeriado(dataStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_FERIADOS);
  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (normalizarData(dados[i][0]) === dataStr) return true;
  }
  return false;
}

// O Sheets às vezes converte sozinho um texto tipo "11/07/2026" para um valor de Data real.
// Essa função normaliza qualquer um dos dois formatos para o mesmo texto "dd/MM/yyyy",
// garantindo que a comparação de "isso é hoje?" sempre funcione.
function normalizarData(valor) {
  if (Object.prototype.toString.call(valor) === '[object Date]') {
    return Utilities.formatDate(valor, TIMEZONE, 'dd/MM/yyyy');
  }
  return String(valor).trim();
}

function normalizarHora(valor) {
  if (Object.prototype.toString.call(valor) === '[object Date]') {
    return Utilities.formatDate(valor, TIMEZONE, 'HH:mm:ss');
  }
  return String(valor).trim();
}

function getStatusDoDia(codigo) {
  const func = getFuncionarioPorCodigo(codigo);
  if (!func) return { ok: false, erro: 'Funcionário não encontrado.' };

  const agora = new Date();
  const dataHoje = Utilities.formatDate(agora, TIMEZONE, 'dd/MM/yyyy');
  const diaSemana = agora.getDay();
  const sequencia = sequenciaDoDia(diaSemana, dataHoje);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_REGISTRO);
  const dados = aba.getDataRange().getValues();

  const marcacoesHoje = [];
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][1]) === String(codigo) && normalizarData(dados[i][3]) === dataHoje) {
      marcacoesHoje.push({ tipo: dados[i][4], hora: normalizarHora(dados[i][5]) });
    }
  }

  let proximaMarcacao = null;
  if (sequencia.length === 0) {
    proximaMarcacao = null; // domingo, sem marcação prevista
  } else if (marcacoesHoje.length < sequencia.length) {
    proximaMarcacao = sequencia[marcacoesHoje.length];
  }

  return {
    ok: true,
    funcionario: { codigo: func.codigo, nome: func.nome, funcao: func.funcao },
    dataHoje: dataHoje,
    ehFolga: sequencia.length === 0,
    marcacoesHoje: marcacoesHoje,
    proximaMarcacao: proximaMarcacao,
    diaCompleto: sequencia.length > 0 && marcacoesHoje.length >= sequencia.length
  };
}

function registrarPonto(codigo, tipo, latitude, longitude) {
  const func = getFuncionarioPorCodigo(codigo);
  if (!func) return { ok: false, erro: 'Funcionário não encontrado.' };

  const status = getStatusDoDia(codigo);
  if (!status.ok) return status;
  if (status.ehFolga) return { ok: false, erro: 'Hoje é domingo (folga) - não há marcação prevista.' };
  if (status.diaCompleto) return { ok: false, erro: 'Todas as marcações de hoje já foram feitas.' };
  if (status.proximaMarcacao !== tipo) {
    return { ok: false, erro: 'Marcação fora de ordem. A próxima esperada é: ' + status.proximaMarcacao };
  }

  const agora = new Date();
  const dataHoje = Utilities.formatDate(agora, TIMEZONE, 'dd/MM/yyyy');
  const horaAgora = Utilities.formatDate(agora, TIMEZONE, 'HH:mm:ss');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_REGISTRO);
  const novoId = aba.getLastRow(); // simples, baseado na linha

  aba.appendRow([novoId, func.codigo, func.nome, dataHoje, tipo, horaAgora, latitude || '', longitude || '', agora.toISOString()]);

  return { ok: true, tipo: tipo, hora: horaAgora, data: dataHoje };
}

// ---------- PAINEL DE ADMINISTRAÇÃO OBBA ----------

function fazerLoginAdmin(usuario, senha) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_ADMINISTRADORES);
  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(usuario)) {
      if (String(dados[i][1]) === String(senha)) {
        return { ok: true, admin: { usuario: dados[i][0], nome: dados[i][2] } };
      }
      return { ok: false, erro: 'Senha incorreta.' };
    }
  }
  return { ok: false, erro: 'Usuário não encontrado.' };
}

function listarFuncionarios() {
  const dados = listarFuncionariosCompletos();
  const lista = [];
  for (let i = 0; i < dados.length; i++) {
    lista.push({
      codigo: String(dados[i].codigo),
      nome: dados[i].nome,
      funcao: dados[i].funcao,
      status: dados[i].status,
      departamento: dados[i].departamento || '',
      endereco: dados[i].endereco || '',
      bairro: dados[i].bairro || '',
      cidade: dados[i].cidade || '',
      estado: dados[i].estado || '',
      cep: dados[i].cep || '',
      hrMes: dados[i].hrMes || '',
      ctps: dados[i].ctps || '',
      cbo: dados[i].cbo || ''
    });
  }
  return { ok: true, funcionarios: lista };
}

// Mostra, pra cada funcionário ativo, quais marcações de hoje já foram feitas e qual falta
function getStatusGeralHoje() {
  const funcs = listarFuncionarios();
  const resultado = [];
  funcs.funcionarios.filter(f => f.status === 'Ativo').forEach(f => {
    const status = getStatusDoDia(f.codigo);
    resultado.push({
      codigo: f.codigo,
      nome: f.nome,
      funcao: f.funcao,
      ehFolga: status.ehFolga,
      marcacoesHoje: status.marcacoesHoje,
      proximaMarcacao: status.proximaMarcacao,
      diaCompleto: status.diaCompleto
    });
  });
  return { ok: true, dataHoje: Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy'), funcionarios: resultado };
}

// Converte "dd/MM/yyyy" em algo comparável (yyyyMMdd) pra poder filtrar por intervalo de datas
function dataParaComparacao(dataStr) {
  const partes = String(dataStr).split('/');
  if (partes.length !== 3) return '';
  return partes[2] + partes[1] + partes[0]; // yyyy + MM + dd
}

function pad2(valor) {
  return String(valor).padStart(2, '0');
}

function parseCompetencia(competencia) {
  const agora = new Date();
  let ano = agora.getFullYear();
  let mes = agora.getMonth() + 1;

  const texto = String(competencia || '').trim();
  if (/^\d{4}-\d{2}$/.test(texto)) {
    ano = Number(texto.slice(0, 4));
    mes = Number(texto.slice(5, 7));
  } else if (/^\d{2}\/\d{4}$/.test(texto)) {
    mes = Number(texto.slice(0, 2));
    ano = Number(texto.slice(3, 7));
  }

  if (!mes || mes < 1 || mes > 12) mes = agora.getMonth() + 1;
  return { ano: ano, mes: mes, competencia: pad2(mes) + '/' + ano };
}

function montarLinhaFolhaDoDia(dataObj, dataStr, registrosDoDia) {
  const linha = {
    data: dataStr,
    dia: pad2(dataObj.getDate()),
    entrada1: '',
    saida1: '',
    entrada2: '',
    saida2: '',
    extras: '',
    visto: ''
  };

  const registrosOrdenados = (registrosDoDia || []).slice().sort((a, b) => String(a.hora).localeCompare(String(b.hora)));
  for (let i = 0; i < registrosOrdenados.length; i++) {
    const registro = registrosOrdenados[i];
    if (registro.tipo === 'Entrada' && !linha.entrada1) linha.entrada1 = registro.hora;
    else if (registro.tipo === 'Saída Almoço' && !linha.saida1) linha.saida1 = registro.hora;
    else if (registro.tipo === 'Volta Almoço' && !linha.entrada2) linha.entrada2 = registro.hora;
    else if (registro.tipo === 'Saída Final') {
      if (!linha.saida1 && dataObj.getDay() === 6) linha.saida1 = registro.hora;
      else if (!linha.saida2) linha.saida2 = registro.hora;
    }
  }

  if (ehFeriado(dataStr)) linha.visto = 'FERIADO';
  else if (dataObj.getDay() === 6) linha.visto = 'SABADO';
  else if (dataObj.getDay() === 0) linha.visto = 'DOMINGO';

  return linha;
}

function gerarFolhaMensal(codigo, competencia) {
  if (!codigo) return { ok: false, erro: 'Selecione um funcionário.' };

  const funcionario = getFuncionarioPorCodigo(codigo);
  if (!funcionario) return { ok: false, erro: 'Funcionário não encontrado.' };

  const jornada = getJornada(funcionario.jornadaId) || {};
  const periodo = parseCompetencia(competencia);
  const ultimoDia = new Date(periodo.ano, periodo.mes, 0).getDate();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_REGISTRO);
  const dados = aba.getDataRange().getValues();
  const registrosPorData = {};

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    if (String(linha[1]) !== String(codigo)) continue;

    const dataLinha = normalizarData(linha[3]);
    const partes = dataLinha.split('/');
    if (partes.length !== 3) continue;
    if (Number(partes[1]) !== periodo.mes || Number(partes[2]) !== periodo.ano) continue;

    if (!registrosPorData[dataLinha]) registrosPorData[dataLinha] = [];
    registrosPorData[dataLinha].push({ tipo: linha[4], hora: normalizarHora(linha[5]) });
  }

  const dias = [];
  for (let dia = 1; dia <= ultimoDia; dia++) {
    const dataObj = new Date(periodo.ano, periodo.mes - 1, dia);
    const dataStr = pad2(dia) + '/' + pad2(periodo.mes) + '/' + periodo.ano;
    dias.push(montarLinhaFolhaDoDia(dataObj, dataStr, registrosPorData[dataStr] || []));
  }

  return {
    ok: true,
    folha: {
      competencia: periodo.competencia,
      funcionario: {
        codigo: funcionario.codigo,
        nome: funcionario.nome || '',
        funcao: funcionario.funcao || '',
        departamento: funcionario.departamento || '',
        endereco: funcionario.endereco || '',
        bairro: funcionario.bairro || '',
        cidade: funcionario.cidade || '',
        estado: funcionario.estado || '',
        cep: funcionario.cep || '',
        hrMes: funcionario.hrMes || '220,00',
        ctps: funcionario.ctps || '',
        cbo: funcionario.cbo || ''
      },
      empresa: EMPRESA_PADRAO,
      horarios: [
        {
          dia: 'Segunda à Sexta',
          expediente: (jornada.segSexEntrada || '08:00') + ' às ' + (jornada.segSexSaida || '18:00'),
          intervalo: (jornada.segSexIntInicio || '12:00') + ' às ' + (jornada.segSexIntFim || '14:00')
        },
        {
          dia: 'Sábado',
          expediente: (jornada.sabEntrada || '08:00') + ' às ' + (jornada.sabSaida || '12:00'),
          intervalo: 'Não Possui'
        },
        {
          dia: 'Domingo',
          expediente: 'Folga',
          intervalo: ''
        }
      ],
      dias: dias
    }
  };
}

// Lista registros de ponto, com filtro opcional por funcionário e/ou intervalo de datas (dd/MM/yyyy)
function listarRegistros(codigo, dataInicio, dataFim) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_REGISTRO);
  const dados = aba.getDataRange().getValues();

  const inicioComp = dataInicio ? dataParaComparacao(dataInicio) : null;
  const fimComp = dataFim ? dataParaComparacao(dataFim) : null;

  const registros = [];
  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    if (codigo && String(linha[1]) !== String(codigo)) continue;

    const dataLinha = normalizarData(linha[3]);
    const dataComp = dataParaComparacao(dataLinha);
    if (inicioComp && dataComp < inicioComp) continue;
    if (fimComp && dataComp > fimComp) continue;

    registros.push({
      id: linha[0], codigo: String(linha[1]), nome: linha[2], data: dataLinha,
      tipo: linha[4], hora: normalizarHora(linha[5]), latitude: linha[6], longitude: linha[7]
    });
  }
  // mais recentes primeiro
  registros.sort((a, b) => (a.data + a.hora < b.data + b.hora) ? 1 : -1);
  return { ok: true, registros: registros };
}

// Corrige uma marcação (data, tipo e/ou hora). Usado pelo painel de administração.
function editarRegistro(id, novaData, novoTipo, novaHora) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_REGISTRO);
  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(id)) {
      const linha = i + 1; // getValues é 0-indexado, planilha é 1-indexada
      if (novaData) aba.getRange(linha, 4).setValue(novaData);
      if (novoTipo) aba.getRange(linha, 5).setValue(novoTipo);
      if (novaHora) aba.getRange(linha, 6).setValue(novaHora);
      return { ok: true };
    }
  }
  return { ok: false, erro: 'Registro não encontrado.' };
}

function excluirRegistro(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_REGISTRO);
  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(id)) {
      aba.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, erro: 'Registro não encontrado.' };
}

function listarFeriados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_FERIADOS);
  const dados = aba.getDataRange().getValues();
  const lista = [];
  for (let i = 1; i < dados.length; i++) {
    lista.push({ data: normalizarData(dados[i][0]), descricao: dados[i][1] });
  }
  lista.sort((a, b) => (dataParaComparacao(a.data) > dataParaComparacao(b.data)) ? 1 : -1);
  return { ok: true, feriados: lista };
}

function adicionarFeriado(data, descricao) {
  if (!data) return { ok: false, erro: 'Informe uma data.' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_FERIADOS);
  aba.getRange('A:A').setNumberFormat('@'); // trava como texto, mesmo motivo do RegistroPonto
  aba.appendRow([data, descricao || '']);
  return { ok: true };
}

function removerFeriado(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_FERIADOS);
  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (normalizarData(dados[i][0]) === data) {
      aba.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, erro: 'Feriado não encontrado.' };
}
