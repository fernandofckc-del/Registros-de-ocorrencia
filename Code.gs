/**
 * REGISTRO DE OCORRÊNCIAS - IRRIGAÇÃO
 * Backend em Google Apps Script.
 *
 * O QUE ISSO FAZ:
 * - Recebe os registros que o app manda (quando o celular está online)
 * - Salva os dados numa aba do Google Sheets (a mesma planilha onde você colar este script)
 * - Salva as fotos (até várias por ocorrência) numa pasta do Google Drive
 * - Guarda a lista de "Tipos de Ocorrência", editável apenas por quem sabe a senha de administrador
 * - Deixa você (ou qualquer app) puxar todos os dados já sincronizados a qualquer momento
 *
 * COMO INSTALAR - veja o arquivo INSTRUCOES.md que veio junto com este arquivo.
 */

// ==== TROQUE ESTE TOKEN por uma palavra secreta sua antes de publicar ====
// Esse token TODO ENCARREGADO vai usar no celular dele pra poder sincronizar.
var SECRET = 'TROQUE_ESTE_TOKEN_123';

// ==== TROQUE ESTA SENHA por outra, diferente da de cima ====
// Essa senha SÓ VOCÊ deve saber. É o que protege a edição dos "Tipos de Ocorrência".
var ADMIN_SECRET = 'TROQUE_ESTA_SENHA_ADMIN_456';

var SHEET_NAME = 'Ocorrencias';
var FOLDER_NAME = 'Fotos_Ocorrencias_Irrigacao';
var FOTO_SEP = ' || ';

var TIPOS_PROP_KEY = 'TIPOS_OCORRENCIA';
var TIPOS_PADRAO = [
  'Mangueira Desconectada',
  'Vazamento',
  'Conector Quebrado',
  'Mangueira Furada (Espinho/Roseta)',
  'Registro/Válvula com Defeito',
  'Gotejador Entupido',
  'Outro'
];

var CAUSAS_PROP_KEY = 'CAUSAS_OCORRENCIA';
var CAUSAS_PADRAO = [
  'Erro Operacional',
  'Desgaste Natural',
  'Espinho/Roseta',
  'Animal',
  'Falha de Material',
  'Pressão Excessiva',
  'Outro'
];

// ---- Programação Semanal de Atividades da Equipe de Irrigação ----
// Uma linha por atividade atribuída a um funcionário num dia. O horário é
// sempre calculado em cascata (igual à Irrigação Cacau), pulando por cima
// do horário de almoço automaticamente quando uma atividade passaria por
// cima dele - o "expediente" fixo é 06:42-12:00 e 13:30-17:00.
//
// Cada atividade cadastrada pode ter um "rendimento" opcional: quantos
// hectares uma quantidade de pessoas cobre num dia útil inteiro (ex.:
// 4 pessoas fazem 18ha/dia). Quando a atividade tem isso configurado, o app
// calcula a duração sozinho a partir da área da vez e de quantas pessoas
// foram escaladas dessa vez - proporcional (menos gente = mais tempo pra
// mesma área). Atividades sem rendimento configurado continuam com duração
// digitada manualmente, igual antes.
var ATIVIDADES_PROP_KEY = 'ATIVIDADES_EQUIPE';
var ATIVIDADES_PADRAO = [
  { nome: 'Correção e conserto do sistema de irrigação', rendimentoHa: 0, rendimentoPessoas: 0 },
  { nome: 'Manutenção preventiva', rendimentoHa: 0, rendimentoPessoas: 0 },
  { nome: 'Instalação de nova válvula', rendimentoHa: 0, rendimentoPessoas: 0 },
  { nome: 'Limpeza de filtro', rendimentoHa: 0, rendimentoPessoas: 0 },
  { nome: 'Troca de mangueira/gotejador', rendimentoHa: 0, rendimentoPessoas: 0 },
  { nome: 'Outro', rendimentoHa: 0, rendimentoPessoas: 0 }
];
var FUNCIONARIOS_PROP_KEY = 'FUNCIONARIOS_EQUIPE';
var FUNCIONARIOS_PADRAO = [
  { nome: 'Encarregado 1', metaHa: 0, turno: 'diurno' },
  { nome: 'Encarregado 2', metaHa: 0, turno: 'diurno' }
];
var SETORES_ATIV_PROP_KEY = 'SETORES_ATIVIDADES';
var SETORES_ATIV_PADRAO = ['TA1', 'TA2', 'TA3', 'TB1', 'TB2', 'D1'];

// Rendimento REAL de fim de dia por funcionário - anotação livre e rápida
// (não precisa bater com nenhuma atividade específica), pra registrar o que
// de fato foi produzido no dia, pra comparar depois com o planejado.
var RENDDIA_SHEET_NAME = 'RendimentoDiario';
var RENDDIA_COLUNAS = ['id', 'data', 'funcionario', 'rendimento', 'excluido', 'criado_em'];

// Banco de horas - falta desconta, plantão/trabalho extra credita. "horas"
// fica guardado em MINUTOS (positivo = crédito, negativo = débito) pra dar
// pra somar certinho sem erro de arredondamento de hora fracionada.
var BANCOHORAS_SHEET_NAME = 'BancoHoras';
var BANCOHORAS_COLUNAS = ['id', 'data', 'funcionario', 'tipo', 'minutos', 'observacao', 'excluido', 'criado_em'];

var PROG_ATIV_SHEET_NAME = 'ProgramacaoAtividades';
var PROG_ATIV_COLUNAS = ['id', 'data', 'funcionario', 'atividade', 'setor', 'duracao_min', 'area_ha',
  'horario_inicio', 'horario_fim', 'observacao', 'concluida', 'excluido', 'criado_em'];

var COLUNAS = ['id', 'data', 'setor', 'tipo', 'causa', 'encarregado', 'descricao',
               'acao', 'status', 'foto_url', 'excluido', 'criado_em', 'cultura'];
var CULTURA_COL = 13;

// Garante que planilhas criadas antes dessa atualização também ganhem a
// coluna "cultura", sem bagunçar nada que já existe - e corrige qualquer
// outro rótulo da linha 1 que tenha ficado diferente do que o script espera
// (ex.: coluna inserida/reordenada manualmente na planilha), sem mexer nos
// dados abaixo.
function ensureCulturaColuna_(sheet) {
  var header = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), CULTURA_COL)).getValues()[0];
  if (header[CULTURA_COL - 1] !== 'cultura') {
    sheet.getRange(1, CULTURA_COL).setValue('cultura');
  }
  var precisaCorrigir = false;
  for (var i = 0; i < COLUNAS.length; i++) {
    if (header[i] !== COLUNAS[i]) { precisaCorrigir = true; break; }
  }
  if (precisaCorrigir) {
    sheet.getRange(1, 1, 1, COLUNAS.length).setValues([COLUNAS]);
  }
}

// ---- Manutenção Preventiva ----
// "Cultura" define o intervalo de reentrada (ex: Coco=15 dias, Cacau=20 dias).
// "Talhão" é o setor real onde a manutenção acontece, e pertence a uma cultura,
// herdando o intervalo dela.
var MANUT_SHEET_NAME = 'Manutencoes';
var MANUT_COLUNAS = ['id', 'setor', 'data', 'encarregado', 'observacao', 'excluido', 'criado_em'];

// ---- Bombeamento (captação de água - 4 bombas) ----
// Uma linha por Bomba por dia. "leitura_inicial"/"leitura_final" são o número
// que aparece no visor do hidrômetro (sem multiplicar) - o consumo real em
// m³ é sempre (leitura_final - leitura_inicial) × 10, calculado e guardado
// já pronto na coluna consumo_m3 pra não precisar recalcular toda vez.
// foto_url_inicio/foto_url_fim ficam em colunas SEPARADAS (não juntas com
// separador numa célula só) justamente pra cada uma virar um link limpo e
// clicável de verdade na planilha - uma célula com "URL1 || URL2" junto faz
// o Google Sheets só reconhecer direito a última URL como link.
var BOMB_SHEET_NAME = 'Bombeamento';
var BOMB_FOLDER_NAME = 'Fotos_Bombeamento';
var BOMB_COLUNAS = ['id', 'data', 'bomba', 'horario_inicio', 'horario_fim',
  'leitura_inicial', 'leitura_final', 'consumo_m3', 'foto_url_inicio', 'foto_url_fim', 'observacao', 'excluido', 'criado_em',
  // 'responsavel' entra no FIM de propósito: inserir no meio deslocaria os
  // dados das linhas antigas (foi o que já causou o problema das colunas
  // trocadas antes). No fim, as linhas antigas só ficam com ele em branco.
  'responsavel'];
var BOMB_COL_RESPONSAVEL = 14;

// ---- Rondas Noturnas (monitoria do plantão) ----
// Uma linha por BOMBA por RONDA - ex.: 4 rondas na noite × 4 bombas = 16
// linhas. O objetivo é justamente ter o rastro de que a conferência foi
// mesmo feita (horario_real gravado na hora do registro, não digitado),
// pra pegar caso de bomba parada/reservatório esvaziando sem ninguém ver.
// O painel de cada bomba pode ser SOFT STARTER (foto + leituras digitadas)
// ou INVERSOR (só foto) - o tipo é configurável por bomba no admin.
var RONDA_SHEET_NAME = 'RondasNoturnas';
var RONDA_COLUNAS = ['id', 'data', 'horario_previsto', 'horario_real', 'plantonista', 'bomba',
  'bomba_ligada', 'nivel_reservatorio', 'painel_tipo', 'painel_status',
  'corrente', 'tensao', 'horimetro', 'foto_painel_url', 'motivo_parada', 'observacao', 'excluido', 'criado_em'];

// Horários das rondas - editável só pelo admin (senha), porque é regra de
// operação, não algo que o plantonista deva mudar.
var RONDAS_HORARIOS_PROP_KEY = 'RONDAS_HORARIOS';
var RONDAS_HORARIOS_PADRAO = ['20:00', '23:00', '02:00', '05:00'];

// Tipo do painel de cada bomba: 'soft' (soft starter) ou 'inversor'.
var PAINEIS_BOMBAS_PROP_KEY = 'PAINEIS_BOMBAS';
var PAINEIS_BOMBAS_PADRAO = [
  { bomba: 1, tipo: 'inversor' },
  { bomba: 2, tipo: 'soft' },
  { bomba: 3, tipo: 'inversor' },
  { bomba: 4, tipo: 'soft' }
];

// Nomes que aparecem no campo "quem registrou" - a dupla do plantão noturno
// e quem mais faça leitura de hidrômetro durante o dia.
var PLANTONISTAS_PROP_KEY = 'PLANTONISTAS_BOMBEAMENTO';
var PLANTONISTAS_PADRAO = ['Alex', 'Elivelto', 'André', 'Ivonaldo'];

var CULTURAS_PROP_KEY = 'CULTURAS_MANUTENCAO';
var CULTURAS_PADRAO = [
  { nome: 'Coco', intervalo: 15, kc: 1.15, modo: 'microaspersor' },
  { nome: 'Cacau', intervalo: 20, kc: 0.94, modo: 'gotejo' }
];

var TALHOES_PROP_KEY = 'TALHOES_MANUTENCAO';
var TALHOES_PADRAO = [
  { nome: 'D1', cultura: 'Coco', valvulas: [] },
  { nome: 'D2', cultura: 'Coco', valvulas: [] },
  { nome: 'TA1', cultura: 'Cacau', valvulas: [] },
  { nome: 'TA2', cultura: 'Cacau', valvulas: [] }
];

// ---- Grupos de Irrigação (separado dos Talhões de Manutenção, de propósito) ----
// "Grupo" aqui = Área/Grupo macro de cálculo (ex.: TA1, T2+T3, T1). É o que
// carrega Kc/KL/Eficiência e agora também uma "descrição" livre (usada na
// coluna Observação do relatório).
var GRUPOS_IRRIG_PROP_KEY = 'GRUPOS_IRRIGACAO';
var GRUPOS_IRRIG_PADRAO = [
  { nome: 'D', cultura: 'Coco', valvulas: [], descricao: 'Área mais velha (espaçamento 6,5 x 6m)' },
  { nome: 'ABC', cultura: 'Coco', valvulas: [], descricao: 'Área mais nova (espaçamento 7,75 x 7m)' },
  { nome: 'TA1', cultura: 'Cacau', valvulas: [], descricao: '' },
  { nome: 'T2+T3', cultura: 'Cacau', valvulas: [], descricao: '' },
  { nome: 'T1', cultura: 'Cacau', valvulas: [], descricao: '' }
];

function getGruposIrrig_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(GRUPOS_IRRIG_PROP_KEY);
  if (!raw) {
    props.setProperty(GRUPOS_IRRIG_PROP_KEY, JSON.stringify(GRUPOS_IRRIG_PADRAO));
    return GRUPOS_IRRIG_PADRAO;
  }
  try {
    var arr = JSON.parse(raw);
    return (arr && arr.length) ? arr : GRUPOS_IRRIG_PADRAO;
  } catch (e) {
    return GRUPOS_IRRIG_PADRAO;
  }
}

// ---- Setores de Irrigação (unidade física/válvula que entra na programação) ----
// Cada Setor pertence a um Grupo (Área/Grupo macro, de onde vem Kc/KL/Eficiência)
// e a um Talhão Integrado (usado só pra exibição/agrupamento no relatório).
var SETORES_IRRIG_PROP_KEY = 'SETORES_IRRIGACAO';
var SETORES_IRRIG_PADRAO = [
  { nome: 'TA1', grupo: 'TA1', talhaoIntegrado: '' },
  { nome: 'TA2', grupo: 'T2+T3', talhaoIntegrado: '' },
  { nome: 'TA3', grupo: 'T2+T3', talhaoIntegrado: '' },
  { nome: 'TA4+TB4', grupo: 'T1', talhaoIntegrado: '' },
  { nome: 'TA5+TB5', grupo: 'T1', talhaoIntegrado: '' },
  { nome: 'TB1+TD1', grupo: 'T1', talhaoIntegrado: '' },
  { nome: 'TB1+TC1', grupo: 'T1', talhaoIntegrado: '' },
  { nome: 'TB2', grupo: 'T2+T3', talhaoIntegrado: '' },
  { nome: 'TB3', grupo: 'T2+T3', talhaoIntegrado: '' },
  { nome: 'TC1+TD2', grupo: 'T1', talhaoIntegrado: '' },
  { nome: 'TC2', grupo: 'T2+T3', talhaoIntegrado: '' },
  { nome: 'TC3', grupo: 'T2+T3', talhaoIntegrado: '' },
  { nome: 'TD1+TD2+QV', grupo: 'T1', talhaoIntegrado: '' },
  { nome: 'TA3+VA4', grupo: 'T2+T3', talhaoIntegrado: '' },
  { nome: 'TB3+VA5', grupo: 'T2+T3', talhaoIntegrado: '' },
  { nome: 'TC3+VA6', grupo: 'T2+T3', talhaoIntegrado: '' },
  { nome: 'QV', grupo: 'T1', talhaoIntegrado: '' }
];

function getSetoresIrrig_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(SETORES_IRRIG_PROP_KEY);
  if (!raw) {
    props.setProperty(SETORES_IRRIG_PROP_KEY, JSON.stringify(SETORES_IRRIG_PADRAO));
    return SETORES_IRRIG_PADRAO;
  }
  try {
    var arr = JSON.parse(raw);
    return (arr && arr.length) ? arr : SETORES_IRRIG_PADRAO;
  } catch (e) {
    return SETORES_IRRIG_PADRAO;
  }
}

// ---- Turmas de Irrigação do Coco (rodízio fixo de válvulas) ----
// Diferente do Cacau, uma Turma aqui NÃO pertence a um único Grupo/Área
// macro (D ou ABC) - ela é só um conjunto fixo de válvulas que liga junto,
// e pode misturar válvulas de D e de ABC na mesma Turma (isso é normal e
// esperado). O prefixo da válvula (ex.: "VD02" = Setor D, "VA10"/"VB.."/
// "VC.." = Setor ABC) é o que decide a vazão usada no cálculo de m³, lá no
// front-end. G1-G4 têm válvulas fixas (confirmadas com o Fernando); G5 e G6
// ficam com a lista vazia até ele preencher, mas todas continuam editáveis
// pelo painel de admin (ele pode trocar válvulas de qualquer Turma quando
// quiser).
var TURMAS_COCO_PROP_KEY = 'TURMAS_COCO';
var TURMAS_COCO_PADRAO = [
  { nome: 'G1', valvulas: ['VA10', 'VB11', 'VB12', 'VC01', 'VC10', 'VC11', 'VD02', 'VD03', 'VD12'] },
  { nome: 'G2', valvulas: ['VA07', 'VA08', 'VB09', 'VC02', 'VC03', 'VC12', 'VD01', 'VD10', 'VD11'] },
  { nome: 'G3', valvulas: ['VA01', 'VA12', 'VB10', 'VC07', 'VC08', 'VC09', 'VD07', 'VD08', 'VD09'] },
  { nome: 'G4', valvulas: ['VA09', 'VB07', 'VB08', 'VC04', 'VC05', 'VC06', 'VD04', 'VD05', 'VD06'] },
  { nome: 'G5', valvulas: [] },
  { nome: 'G6', valvulas: [] }
];

function getTurmasCoco_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(TURMAS_COCO_PROP_KEY);
  if (!raw) {
    props.setProperty(TURMAS_COCO_PROP_KEY, JSON.stringify(TURMAS_COCO_PADRAO));
    return TURMAS_COCO_PADRAO;
  }
  try {
    var arr = JSON.parse(raw);
    return (arr && arr.length) ? arr : TURMAS_COCO_PADRAO;
  } catch (e) {
    return TURMAS_COCO_PADRAO;
  }
}

// ---- Horário de Irrigação (ET0 / lâmina calculada) ----
var IRRIG_SHEET_NAME = 'Irrigacao';
var IRRIG_COLUNAS = ['id', 'data', 'cultura', 'grupo', 'setor', 'talhao_integrado', 'modo', 'eto', 'kc_usado', 'kl_usado', 'lamina_mm',
                      'tempo_recomendado_min', 'litros_planta', 'm3_valvula',
                      'ordem', 'horario', 'horario_fim', 'tempo_min',
                      'tanque_a', 'tanque_b', 'tanque_c', 'tanque_d', 'tanque_e', 'tanque_f',
                      'finalizada', 'observacao', 'excluido', 'criado_em'];
// Índice (1-based) da coluna "excluido" na planilha de Irrigação - usado na
// exclusão lógica. Mantido como constante pra não depender de contar campos.
var IRRIG_EXCLUIDO_COL = 27;

// Planilhas criadas antes desta atualização não tinham as colunas "setor" e
// "talhao_integrado" (elas entravam junto do "grupo"), e o app também foi
// ganhando colunas novas ao longo do tempo (m3_valvula, tempo_min etc.).
// Antes, essa função só conferia UMA posição (a 5ª coluna) - se alguém
// mexesse em qualquer outra coluna direto na planilha (inserir, apagar,
// reordenar manualmente), os rótulos da linha 1 ficavam desalinhados com os
// dados de verdade (que o script sempre escreve na ordem de IRRIG_COLUNAS),
// e aí uma coluna chamada "Kc" podia aparecer cheia de horário, por exemplo.
// Agora ela CONFERE a linha 1 inteira contra IRRIG_COLUNAS e reescreve
// qualquer nome que estiver diferente - sem tocar nos dados, só o cabeçalho.
function ensureIrrigColunas_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), IRRIG_COLUNAS.length);
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // Migração antiga: se a coluna "setor" nem existe ainda (planilha bem
  // velha), abre espaço pra ela e pra "talhao_integrado" antes de mais nada.
  if (header[4] !== 'setor' && header[4] !== 'talhao_integrado') {
    sheet.insertColumnsAfter(4, 2);
    header = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), IRRIG_COLUNAS.length)).getValues()[0];
  }

  // Garante colunas suficientes pra caber todo o IRRIG_COLUNAS.
  if (sheet.getMaxColumns() < IRRIG_COLUNAS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), IRRIG_COLUNAS.length - sheet.getMaxColumns());
  }

  var precisaCorrigir = false;
  for (var i = 0; i < IRRIG_COLUNAS.length; i++) {
    if (header[i] !== IRRIG_COLUNAS[i]) { precisaCorrigir = true; break; }
  }
  if (precisaCorrigir) {
    sheet.getRange(1, 1, 1, IRRIG_COLUNAS.length).setValues([IRRIG_COLUNAS]);
  }

  // Sem isso, o Google Sheets às vezes "adivinha" que um texto tipo "06:30"
  // ou "2026-08-01" é um horário/data de verdade e converte a célula pra um
  // valor interno de data - aí quando o app lê de volta, vem um Date
  // esquisito tipo 1899-12-31 em vez do texto original. Forçando essas
  // colunas como "Texto simples" (@), o Sheets para de fazer essa conversão
  // automática nas escritas futuras. Não desfaz o que já foi convertido
  // (isso é corrigido na leitura, em getIrrigacoesFormatadas_).
  var TEXTO_COLS_IRRIG_ = [1, 2, 3, 4, 5, 6, 7, 16, 17, 26]; // id, data, cultura, grupo, setor, talhao_integrado, modo, horario, horario_fim, observacao
  var totalLinhas = Math.max(sheet.getMaxRows() - 1, 1);
  TEXTO_COLS_IRRIG_.forEach(function (col) {
    sheet.getRange(2, col, totalLinhas, 1).setNumberFormat('@');
  });
}

function getIrrigSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(IRRIG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(IRRIG_SHEET_NAME);
    sheet.appendRow(IRRIG_COLUNAS);
    sheet.setFrozenRows(1);
  }
  ensureIrrigColunas_(sheet);
  return sheet;
}

// A aba "Irrigacao" é a fonte de verdade (é nela que o app se baseia pra
// sincronizar/editar/apagar por id) e por isso mistura Cacau e Coco, com
// várias colunas que só fazem sentido pra uma cultura ou pra outra (ex.:
// tanques de fertirrigação só valem pro Cacau, m3_valvula só pro Coco).
// Essas duas abas aqui são só uma FOTO somente-leitura, gerada de novo a
// cada sincronização, com só as colunas relevantes pra cada cultura - pra
// ficar fácil de olhar/filtrar direto no Drive sem misturar as duas coisas.
var IRRIG_TAB_CACAU_NOME = 'Irrigação - Cacau';
var IRRIG_TAB_COCO_NOME = 'Irrigação - Coco';
var IRRIG_TAB_CACAU_COLS_ = ['data', 'setor', 'grupo', 'talhao_integrado', 'eto', 'kc_usado', 'kl_usado',
  'lamina_mm', 'litros_planta', 'horario', 'horario_fim', 'tempo_min',
  'tanque_a', 'tanque_b', 'tanque_c', 'tanque_d', 'tanque_e', 'tanque_f', 'finalizada', 'observacao'];
var IRRIG_TAB_COCO_COLS_ = ['data', 'setor', 'eto', 'kc_usado', 'kl_usado', 'lamina_mm',
  'horario', 'horario_fim', 'tempo_min', 'm3_valvula', 'finalizada', 'observacao'];

function syncIrrigTabsPorCultura_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var isheet = getIrrigSheet_();
  var lastRow = isheet.getLastRow();
  var dados = lastRow > 1 ? isheet.getRange(2, 1, lastRow - 1, IRRIG_COLUNAS.length).getValues() : [];
  var idx = {};
  IRRIG_COLUNAS.forEach(function (nome, i) { idx[nome] = i; });

  function escreverAba_(nomeAba, colunas, cultura) {
    var aba = ss.getSheetByName(nomeAba);
    if (!aba) aba = ss.insertSheet(nomeAba);
    aba.clear();
    var cabecalho = colunas.map(function (c) { return c === 'setor' ? (cultura === 'Coco' ? 'turma' : 'setor') : c; });
    aba.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]);
    aba.setFrozenRows(1);
    var linhas = dados
      .filter(function (r) { return String(r[idx.cultura]) === cultura && !r[idx.excluido]; })
      .sort(function (a, b) {
        var da = String(a[idx.data]), db = String(b[idx.data]);
        if (da !== db) return da < db ? -1 : 1;
        return (parseFloat(a[idx.ordem]) || 0) - (parseFloat(b[idx.ordem]) || 0);
      })
      .map(function (r) { return colunas.map(function (c) { return r[idx[c]]; }); });
    if (linhas.length) aba.getRange(2, 1, linhas.length, colunas.length).setValues(linhas);
  }

  escreverAba_(IRRIG_TAB_CACAU_NOME, IRRIG_TAB_CACAU_COLS_, 'Cacau');
  escreverAba_(IRRIG_TAB_COCO_NOME, IRRIG_TAB_COCO_COLS_, 'Coco');
}

function ensureManutColunas_(sheet) {
  var header = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), MANUT_COLUNAS.length)).getValues()[0];
  var precisaCorrigir = false;
  for (var i = 0; i < MANUT_COLUNAS.length; i++) {
    if (header[i] !== MANUT_COLUNAS[i]) { precisaCorrigir = true; break; }
  }
  if (precisaCorrigir) {
    sheet.getRange(1, 1, 1, MANUT_COLUNAS.length).setValues([MANUT_COLUNAS]);
  }
}

function getManutSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MANUT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MANUT_SHEET_NAME);
    sheet.appendRow(MANUT_COLUNAS);
    sheet.setFrozenRows(1);
  }
  ensureManutColunas_(sheet);
  return sheet;
}

function ensureBombColunas_(sheet) {
  var header = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), BOMB_COLUNAS.length)).getValues()[0];
  var precisaCorrigir = false;
  for (var i = 0; i < BOMB_COLUNAS.length; i++) {
    if (header[i] !== BOMB_COLUNAS[i]) { precisaCorrigir = true; break; }
  }
  if (precisaCorrigir) {
    sheet.getRange(1, 1, 1, BOMB_COLUNAS.length).setValues([BOMB_COLUNAS]);
  }
}

// Linhas criadas ANTES da correção que separou "foto_url" em duas colunas
// (foto_url_inicio/foto_url_fim) ainda têm as duas URLs juntas numa célula
// só, com " || " entre elas - e o resto dos dados dessa linha (observação,
// excluído, criado_em) ficou uma coluna pra trás do que a linha 1 diz agora.
// Isso detecta esse formato antigo (reconhece pelo " || " sobrando na
// coluna 9) e reorganiza os valores pras colunas certas, automaticamente.
function migrarLinhasBombAntigas_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var range = sheet.getRange(2, 1, lastRow - 1, BOMB_COLUNAS.length);
  var dados = range.getValues();
  var mudou = false;
  dados.forEach(function (r) {
    var col9 = String(r[8] || '');
    if (col9.indexOf(FOTO_SEP) === -1) return; // já está no formato novo, não mexe
    var partes = col9.split(FOTO_SEP);
    var fotoInicioAntiga = (partes[0] || '').trim();
    var fotoFimAntiga = (partes[1] || '').trim();
    // no formato antigo: col9=foto_url junto, col10=observacao, col11=excluido, col12=criado_em
    var observacaoAntiga = r[9];
    var excluidoAntigo = r[10];
    var criadoEmAntigo = r[11];
    r[8] = fotoInicioAntiga;      // foto_url_inicio
    r[9] = fotoFimAntiga;         // foto_url_fim
    r[10] = observacaoAntiga;     // observacao
    r[11] = excluidoAntigo;       // excluido
    r[12] = criadoEmAntigo;       // criado_em
    mudou = true;
  });
  if (mudou) range.setValues(dados);
}

function getBombSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(BOMB_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(BOMB_SHEET_NAME);
    sheet.appendRow(BOMB_COLUNAS);
    sheet.setFrozenRows(1);
  }
  ensureBombColunas_(sheet);
  migrarLinhasBombAntigas_(sheet);
  return sheet;
}

function ensureRondaColunas_(sheet) {
  var header = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), RONDA_COLUNAS.length)).getValues()[0];
  var precisaCorrigir = false;
  for (var i = 0; i < RONDA_COLUNAS.length; i++) {
    if (header[i] !== RONDA_COLUNAS[i]) { precisaCorrigir = true; break; }
  }
  if (precisaCorrigir) {
    sheet.getRange(1, 1, 1, RONDA_COLUNAS.length).setValues([RONDA_COLUNAS]);
  }
}

function getRondaSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RONDA_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RONDA_SHEET_NAME);
    sheet.appendRow(RONDA_COLUNAS);
    sheet.setFrozenRows(1);
  }
  ensureRondaColunas_(sheet);
  return sheet;
}

function getRondasHorarios_() { return getListaConfig_(RONDAS_HORARIOS_PROP_KEY, RONDAS_HORARIOS_PADRAO); }
function getPaineisBombas_() { return getListaConfig_(PAINEIS_BOMBAS_PROP_KEY, PAINEIS_BOMBAS_PADRAO); }
function getPlantonistas_() { return getListaConfig_(PLANTONISTAS_PROP_KEY, PLANTONISTAS_PADRAO); }

function getFolderBomb_() {
  var folders = DriveApp.getFoldersByName(BOMB_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(BOMB_FOLDER_NAME);
}

// Igual ao saveFoto_/processFotos_ já existentes, só que salvando na pasta
// separada do Bombeamento em vez da pasta de Ocorrências.
function saveFotoBomb_(dataUrl, id) {
  var match = dataUrl.match(/^data:(image\/\w+);base64,(.*)$/);
  if (!match) return '';
  var mime = match[1];
  var base64 = match[2];
  var bytes = Utilities.base64Decode(base64);
  var blob = Utilities.newBlob(bytes, mime, id + '.jpg');
  var folder = getFolderBomb_();
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?id=' + file.getId();
}


function getCulturas_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(CULTURAS_PROP_KEY);
  if (!raw) {
    props.setProperty(CULTURAS_PROP_KEY, JSON.stringify(CULTURAS_PADRAO));
    return CULTURAS_PADRAO;
  }
  try {
    var arr = JSON.parse(raw);
    return (arr && arr.length) ? arr : CULTURAS_PADRAO;
  } catch (e) {
    return CULTURAS_PADRAO;
  }
}

function getTalhoes_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(TALHOES_PROP_KEY);
  if (!raw) {
    props.setProperty(TALHOES_PROP_KEY, JSON.stringify(TALHOES_PADRAO));
    return TALHOES_PADRAO;
  }
  try {
    var arr = JSON.parse(raw);
    return (arr && arr.length) ? arr : TALHOES_PADRAO;
  } catch (e) {
    return TALHOES_PADRAO;
  }
}

// ---- Parâmetros técnicos de irrigação por Grupo (só admin edita) ----
// Objeto: { "NomeDoGrupo": { ...campos dependendo do modo da cultura... } }
var PARAMS_IRRIG_PROP_KEY = 'PARAMS_IRRIGACAO_GRUPOS';

// KL (coeficiente de localização) de cada Grupo do Cacau, por idade da área:
// TA1 = talhão mais velho, T2+T3 = área intermediária, T1 = talhão mais novo.
// Antes esses parâmetros nunca eram pré-preenchidos (só ficavam vazios até
// alguém digitar no painel de admin) - e como o modo gotejo calcula o KL pela
// geometria da planta quando ele está vazio, um Grupo sem geometria cadastrada
// e sem KL manual calculava lâmina 0 (silenciosamente). Isso pré-preenche os
// 3 grupos do Cacau com o KL manual certo assim que a planilha é criada,
// exatamente como os outros getters de configuração (getCulturas_,
// getTalhoes_ etc.) já fazem.
var PARAMS_IRRIG_PADRAO = {
  'TA1': { kl: 1 },
  'T2+T3': { kl: 0.67 },
  'T1': { kl: 0.47 },
  // Coco (microaspersor): só a vazão da válvula veio pronta (43 m³/h nas
  // válvulas A/B/C, 45 m³/h nas D). KL, Kc, mm/hora, Fator de Manejo e
  // Eficiência ficam em branco até o admin preencher no painel de
  // Parâmetros de Irrigação (mesma tela usada pelo Cacau).
  'ABC': { vazaoValvula: 43 },
  'D': { vazaoValvula: 45 }
};

function getParamsIrrigacao_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(PARAMS_IRRIG_PROP_KEY);
  if (!raw) {
    props.setProperty(PARAMS_IRRIG_PROP_KEY, JSON.stringify(PARAMS_IRRIG_PADRAO));
    return PARAMS_IRRIG_PADRAO;
  }
  try {
    var parsed = JSON.parse(raw) || {};
    // Garante o KL desses 3 grupos mesmo que o grupo já exista nos parâmetros
    // salvos (ex.: o admin já salvou o painel uma vez, mesmo sem mexer nesses
    // campos - nesse caso o grupo já "existe" só que vazio, e um simples
    // "só preenche se não existir" nunca entrava). Só toca no campo kl
    // quando ele estiver ausente/inválido; nunca sobrescreve um KL que o
    // admin já tenha digitado manualmente, nem mexe nos outros campos do
    // grupo (kc, eficiência, geometria etc.).
    var mudou = false;
    Object.keys(PARAMS_IRRIG_PADRAO).forEach(function (nome) {
      if (!parsed[nome]) { parsed[nome] = {}; mudou = true; }
      var klAtual = parseFloat(parsed[nome].kl);
      if (isNaN(klAtual) || klAtual <= 0) {
        parsed[nome].kl = PARAMS_IRRIG_PADRAO[nome].kl;
        mudou = true;
      }
    });
    if (mudou) props.setProperty(PARAMS_IRRIG_PROP_KEY, JSON.stringify(parsed));
    return parsed;
  } catch (e) {
    return {};
  }
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(COLUNAS);
    sheet.setFrozenRows(1);
  }
  ensureCulturaColuna_(sheet);
  return sheet;
}

function getFolder_() {
  var folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(FOLDER_NAME);
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function findRowById_(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2; // linha real na planilha
  }
  return -1;
}

function saveFoto_(dataUrl, id) {
  var match = dataUrl.match(/^data:(image\/\w+);base64,(.*)$/);
  if (!match) return '';
  var mime = match[1];
  var base64 = match[2];
  var bytes = Utilities.base64Decode(base64);
  var blob = Utilities.newBlob(bytes, mime, id + '.jpg');
  var folder = getFolder_();
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?id=' + file.getId();
}

// Recebe uma lista (array) onde cada item é OU uma foto nova (base64 "data:image/...")
// OU uma URL que já existia. Faz upload só das novas, e devolve a lista final de URLs.
function processFotos_(fotos, baseId) {
  if (!fotos || !fotos.length) return [];
  var urls = [];
  fotos.forEach(function (item, i) {
    if (!item) return;
    if (String(item).indexOf('data:image') === 0) {
      var url = saveFoto_(item, baseId + '_' + i + '_' + new Date().getTime());
      if (url) urls.push(url);
    } else {
      urls.push(item); // já era uma URL, mantém
    }
  });
  return urls;
}

function getListaConfig_(propKey, padrao) {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(propKey);
  if (!raw) {
    props.setProperty(propKey, JSON.stringify(padrao));
    return padrao;
  }
  try {
    var arr = JSON.parse(raw);
    return (arr && arr.length) ? arr : padrao;
  } catch (e) {
    return padrao;
  }
}
function getTipos_() { return getListaConfig_(TIPOS_PROP_KEY, TIPOS_PADRAO); }
function getCausas_() { return getListaConfig_(CAUSAS_PROP_KEY, CAUSAS_PADRAO); }
function getAtividadesLista_() { return getListaConfig_(ATIVIDADES_PROP_KEY, ATIVIDADES_PADRAO); }
function getFuncionariosLista_() { return getListaConfig_(FUNCIONARIOS_PROP_KEY, FUNCIONARIOS_PADRAO); }
function getSetoresAtivLista_() { return getListaConfig_(SETORES_ATIV_PROP_KEY, SETORES_ATIV_PADRAO); }

function ensureProgAtivColunas_(sheet) {
  var header = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), PROG_ATIV_COLUNAS.length)).getValues()[0];
  var precisaCorrigir = false;
  for (var i = 0; i < PROG_ATIV_COLUNAS.length; i++) {
    if (header[i] !== PROG_ATIV_COLUNAS[i]) { precisaCorrigir = true; break; }
  }
  if (precisaCorrigir) {
    sheet.getRange(1, 1, 1, PROG_ATIV_COLUNAS.length).setValues([PROG_ATIV_COLUNAS]);
  }
}

function getProgAtivSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PROG_ATIV_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PROG_ATIV_SHEET_NAME);
    sheet.appendRow(PROG_ATIV_COLUNAS);
    sheet.setFrozenRows(1);
  }
  ensureProgAtivColunas_(sheet);
  return sheet;
}

function ensureRendDiaColunas_(sheet) {
  var header = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), RENDDIA_COLUNAS.length)).getValues()[0];
  var precisaCorrigir = false;
  for (var i = 0; i < RENDDIA_COLUNAS.length; i++) {
    if (header[i] !== RENDDIA_COLUNAS[i]) { precisaCorrigir = true; break; }
  }
  if (precisaCorrigir) {
    sheet.getRange(1, 1, 1, RENDDIA_COLUNAS.length).setValues([RENDDIA_COLUNAS]);
  }
}

function getRendDiaSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RENDDIA_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RENDDIA_SHEET_NAME);
    sheet.appendRow(RENDDIA_COLUNAS);
    sheet.setFrozenRows(1);
  }
  ensureRendDiaColunas_(sheet);
  return sheet;
}

function ensureBancoHorasColunas_(sheet) {
  var header = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), BANCOHORAS_COLUNAS.length)).getValues()[0];
  var precisaCorrigir = false;
  for (var i = 0; i < BANCOHORAS_COLUNAS.length; i++) {
    if (header[i] !== BANCOHORAS_COLUNAS[i]) { precisaCorrigir = true; break; }
  }
  if (precisaCorrigir) {
    sheet.getRange(1, 1, 1, BANCOHORAS_COLUNAS.length).setValues([BANCOHORAS_COLUNAS]);
  }
}

function getBancoHorasSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(BANCOHORAS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(BANCOHORAS_SHEET_NAME);
    sheet.appendRow(BANCOHORAS_COLUNAS);
    sheet.setFrozenRows(1);
  }
  ensureBancoHorasColunas_(sheet);
  return sheet;
}

// O Google Sheets às vezes decide sozinho que um texto tipo "06:30" ou
// "2026-08-01" é uma data/horário de verdade e guarda a célula como um valor
// interno de data - aí getValues() devolve um objeto Date em vez do texto
// original. Isso "recupera" o texto certo: valores com ano <= 1900 são
// horário-só (Sheets usa 30/12/1899 como o "dia zero" pra isso) e viram
// "HH:mm"; o resto vira "yyyy-MM-dd". Qualquer outro valor passa direto.
function sanitizarValorPlanilha_(valor, timeZone) {
  if (Object.prototype.toString.call(valor) !== '[object Date]' || isNaN(valor.getTime())) return valor;
  if (valor.getFullYear() <= 1900) {
    return Utilities.formatDate(valor, timeZone, 'HH:mm');
  }
  return Utilities.formatDate(valor, timeZone, 'yyyy-MM-dd');
}

// GET -> devolve todos os registros + a lista atual de tipos de ocorrência
function doGet(e) {
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  var headers = data.shift();
  var rows = data.map(function (r) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = sanitizarValorPlanilha_(r[i], tz); });
    obj.fotos = obj.foto_url ? String(obj.foto_url).split(FOTO_SEP).filter(function (x) { return x; }) : [];
    return obj;
  }).filter(function (r) { return !r.excluido; });

  var manutSheet = getManutSheet_();
  var manutData = manutSheet.getDataRange().getValues();
  var manutHeaders = manutData.shift();
  var manutRows = manutData.map(function (r) {
    var obj = {};
    manutHeaders.forEach(function (h, i) { obj[h] = sanitizarValorPlanilha_(r[i], tz); });
    return obj;
  }).filter(function (r) { return !r.excluido; });

  var bombSheet = getBombSheet_();
  var bombData = bombSheet.getDataRange().getValues();
  var bombHeaders = bombData.shift();
  var bombRows = bombData.map(function (r) {
    var obj = {};
    bombHeaders.forEach(function (h, i) { obj[h] = sanitizarValorPlanilha_(r[i], tz); });
    obj.fotos = [obj.foto_url_inicio, obj.foto_url_fim].filter(function (x) { return x; });
    return obj;
  }).filter(function (r) { return !r.excluido; });

  var rondaSheet = getRondaSheet_();
  var rondaData = rondaSheet.getDataRange().getValues();
  var rondaHeaders = rondaData.shift();
  var rondaRows = rondaData.map(function (r) {
    var obj = {};
    rondaHeaders.forEach(function (h, i) { obj[h] = sanitizarValorPlanilha_(r[i], tz); });
    return obj;
  }).filter(function (r) { return !r.excluido; });

  var irrigSheet = getIrrigSheet_();
  var irrigData = irrigSheet.getDataRange().getValues();
  var irrigHeaders = irrigData.shift();

  // Linhas bem antigas (de antes do conceito de "id" existir nessa aba)
  // ficaram com a coluna id em branco. Sem id único, o botão de apagar de
  // várias linhas diferentes aponta pro mesmo "vazio" - clicar numa nunca
  // apaga exatamente a que foi clicada. Isso preenche um id novo, de uma
  // vez só, permanentemente, pra cada linha que ainda não tinha.
  var precisaPreencherIds = false;
  irrigData.forEach(function (r, i) {
    // Só preenche id em linha que tem dado de verdade (data preenchida) -
    // linha totalmente vazia (sobra de edição manual na planilha) fica
    // vazia mesmo, sem gerar um id "fantasma" à toa.
    if (!r[0] && r[1]) { r[0] = 'legado_' + Utilities.getUuid().slice(0, 8) + '_' + i; precisaPreencherIds = true; }
  });
  if (precisaPreencherIds) {
    irrigSheet.getRange(2, 1, irrigData.length, 1).setValues(irrigData.map(function (r) { return [r[0]]; }));
  }

  var irrigRows = irrigData.map(function (r) {
    var obj = {};
    irrigHeaders.forEach(function (h, i) { obj[h] = sanitizarValorPlanilha_(r[i], tz); });
    return obj;
  }).filter(function (r) { return !r.excluido; });

  var progAtivSheet = getProgAtivSheet_();
  var progAtivData = progAtivSheet.getDataRange().getValues();
  var progAtivHeaders = progAtivData.shift();
  var progAtivRows = progAtivData.map(function (r) {
    var obj = {};
    progAtivHeaders.forEach(function (h, i) { obj[h] = sanitizarValorPlanilha_(r[i], tz); });
    return obj;
  }).filter(function (r) { return !r.excluido; });

  var rendDiaSheet = getRendDiaSheet_();
  var rendDiaData = rendDiaSheet.getDataRange().getValues();
  var rendDiaHeaders = rendDiaData.shift();
  var rendDiaRows = rendDiaData.map(function (r) {
    var obj = {};
    rendDiaHeaders.forEach(function (h, i) { obj[h] = sanitizarValorPlanilha_(r[i], tz); });
    return obj;
  }).filter(function (r) { return !r.excluido; });

  var bancoHorasSheet = getBancoHorasSheet_();
  var bancoHorasData = bancoHorasSheet.getDataRange().getValues();
  var bancoHorasHeaders = bancoHorasData.shift();
  var bancoHorasRows = bancoHorasData.map(function (r) {
    var obj = {};
    bancoHorasHeaders.forEach(function (h, i) { obj[h] = sanitizarValorPlanilha_(r[i], tz); });
    return obj;
  }).filter(function (r) { return !r.excluido; });

  return jsonOut_({
    ok: true, rows: rows, tipos: getTipos_(), causas: getCausas_(),
    manutencoes: manutRows, culturas: getCulturas_(), talhoes: getTalhoes_(),
    gruposIrrigacao: getGruposIrrig_(),
    setoresIrrigacao: getSetoresIrrig_(),
    parametrosIrrigacao: getParamsIrrigacao_(),
    turmasCoco: getTurmasCoco_(),
    irrigacoes: irrigRows,
    bombeamentos: bombRows,
    rondasNoturnas: rondaRows,
    rondasHorarios: getRondasHorarios_(),
    paineisBombas: getPaineisBombas_(),
    plantonistas: getPlantonistas_(),
    atividadesLista: getAtividadesLista_(),
    funcionariosLista: getFuncionariosLista_(),
    setoresAtivLista: getSetoresAtivLista_(),
    programacaoAtividades: progAtivRows,
    rendimentosDiarios: rendDiaRows,
    bancoHoras: bancoHorasRows
  });
}

// POST -> recebe um registro novo, edição completa, mudança de status, exclusão, ou atualização dos tipos
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action || 'create';

    // Ação especial de administrador: não usa o token normal, usa a senha de admin
    if (action === 'update_tipos') {
      if (body.adminToken !== ADMIN_SECRET) {
        return jsonOut_({ ok: false, error: 'senha de administrador inválida' });
      }
      var novosTipos = (body.tipos || []).map(function (t) { return String(t).trim(); }).filter(function (t) { return t; });
      if (!novosTipos.length) {
        return jsonOut_({ ok: false, error: 'lista de tipos vazia' });
      }
      PropertiesService.getScriptProperties().setProperty(TIPOS_PROP_KEY, JSON.stringify(novosTipos));
      return jsonOut_({ ok: true, tipos: novosTipos });
    }

    if (action === 'update_causas') {
      if (body.adminToken !== ADMIN_SECRET) {
        return jsonOut_({ ok: false, error: 'senha de administrador inválida' });
      }
      var novasCausas = (body.causas || []).map(function (t) { return String(t).trim(); }).filter(function (t) { return t; });
      if (!novasCausas.length) {
        return jsonOut_({ ok: false, error: 'lista de causas vazia' });
      }
      PropertiesService.getScriptProperties().setProperty(CAUSAS_PROP_KEY, JSON.stringify(novasCausas));
      return jsonOut_({ ok: true, causas: novasCausas });
    }

    if (action === 'update_atividades_lista') {
      if (body.adminToken !== ADMIN_SECRET) {
        return jsonOut_({ ok: false, error: 'senha de administrador inválida' });
      }
      var novasAtividades = (body.atividades || []).map(function (t) {
        return {
          nome: String(t.nome || '').trim(),
          rendimentoHa: parseFloat(t.rendimentoHa) || 0,
          rendimentoPessoas: parseFloat(t.rendimentoPessoas) || 0
        };
      }).filter(function (t) { return t.nome; });
      if (!novasAtividades.length) {
        return jsonOut_({ ok: false, error: 'lista de atividades vazia' });
      }
      PropertiesService.getScriptProperties().setProperty(ATIVIDADES_PROP_KEY, JSON.stringify(novasAtividades));
      return jsonOut_({ ok: true, atividades: novasAtividades });
    }

    if (action === 'update_funcionarios_lista') {
      if (body.adminToken !== ADMIN_SECRET) {
        return jsonOut_({ ok: false, error: 'senha de administrador inválida' });
      }
      var novosFuncionarios = (body.funcionarios || []).map(function (t) {
        return {
          nome: String(t.nome || '').trim(), metaHa: parseFloat(t.metaHa) || 0,
          turno: (String(t.turno || '').trim().toLowerCase() === 'noturno') ? 'noturno' : 'diurno'
        };
      }).filter(function (t) { return t.nome; });
      if (!novosFuncionarios.length) {
        return jsonOut_({ ok: false, error: 'lista de funcionários vazia' });
      }
      PropertiesService.getScriptProperties().setProperty(FUNCIONARIOS_PROP_KEY, JSON.stringify(novosFuncionarios));
      return jsonOut_({ ok: true, funcionarios: novosFuncionarios });
    }

    if (action === 'update_setores_ativ_lista') {
      if (body.adminToken !== ADMIN_SECRET) {
        return jsonOut_({ ok: false, error: 'senha de administrador inválida' });
      }
      var novosSetores = (body.setores || []).map(function (t) { return String(t).trim(); }).filter(function (t) { return t; });
      if (!novosSetores.length) {
        return jsonOut_({ ok: false, error: 'lista de setores vazia' });
      }
      PropertiesService.getScriptProperties().setProperty(SETORES_ATIV_PROP_KEY, JSON.stringify(novosSetores));
      return jsonOut_({ ok: true, setores: novosSetores });
    }

    if (action === 'delete_atividade_prog') {
      if (body.adminToken !== ADMIN_SECRET) {
        return jsonOut_({ ok: false, error: 'senha de administrador inválida' });
      }
      var pasheet = getProgAtivSheet_();
      var paids = pasheet.getRange(2, 1, Math.max(pasheet.getLastRow() - 1, 0), 1).getValues();
      var parow = -1;
      for (var pai = 0; pai < paids.length; pai++) {
        if (String(paids[pai][0]) === String(body.id)) { parow = pai + 2; break; }
      }
      if (parow === -1) {
        return jsonOut_({ ok: false, error: 'atividade não encontrada (id: ' + body.id + ')' });
      }
      pasheet.getRange(parow, 12).setValue(true); // coluna 12 = excluido
      return jsonOut_({ ok: true });
    }

    if (action === 'delete_banco_horas') {
      if (body.adminToken !== ADMIN_SECRET) {
        return jsonOut_({ ok: false, error: 'senha de administrador inválida' });
      }
      var bhsheetD = getBancoHorasSheet_();
      var bhidsD = bhsheetD.getRange(2, 1, Math.max(bhsheetD.getLastRow() - 1, 0), 1).getValues();
      var bhrowD = -1;
      for (var bhiD = 0; bhiD < bhidsD.length; bhiD++) {
        if (String(bhidsD[bhiD][0]) === String(body.id)) { bhrowD = bhiD + 2; break; }
      }
      if (bhrowD === -1) return jsonOut_({ ok: false, error: 'lançamento não encontrado (id: ' + body.id + ')' });
      bhsheetD.getRange(bhrowD, 7).setValue(true); // coluna 7 = excluido
      return jsonOut_({ ok: true });
    }

    if (action === 'update_culturas') {
      if (body.adminToken !== ADMIN_SECRET) {
        return jsonOut_({ ok: false, error: 'senha de administrador inválida' });
      }
      var novasCulturas = (body.culturas || []).map(function (s) {
        return {
          nome: String(s.nome || '').trim(),
          intervalo: parseInt(s.intervalo, 10) || 15,
          kc: parseFloat(s.kc) || 0.85,
          modo: (String(s.modo || '').trim().toLowerCase() === 'microaspersor') ? 'microaspersor' : 'gotejo'
        };
      }).filter(function (s) { return s.nome; });
      if (!novasCulturas.length) {
        return jsonOut_({ ok: false, error: 'lista de culturas vazia' });
      }
      PropertiesService.getScriptProperties().setProperty(CULTURAS_PROP_KEY, JSON.stringify(novasCulturas));
      return jsonOut_({ ok: true, culturas: novasCulturas });
    }

    if (action === 'update_talhoes') {
      if (body.adminToken !== ADMIN_SECRET) {
        return jsonOut_({ ok: false, error: 'senha de administrador inválida' });
      }
      var novosTalhoes = (body.talhoes || []).map(function (t) {
        return {
          nome: String(t.nome || '').trim(),
          cultura: String(t.cultura || '').trim(),
          valvulas: Array.isArray(t.valvulas) ? t.valvulas.map(function (v) { return String(v).trim(); }).filter(function (v) { return v; }) : []
        };
      }).filter(function (t) { return t.nome && t.cultura; });
      if (!novosTalhoes.length) {
        return jsonOut_({ ok: false, error: 'lista de talhões vazia' });
      }
      PropertiesService.getScriptProperties().setProperty(TALHOES_PROP_KEY, JSON.stringify(novosTalhoes));
      return jsonOut_({ ok: true, talhoes: novosTalhoes });
    }

    if (action === 'update_grupos_irrigacao') {
      if (body.adminToken !== ADMIN_SECRET) {
        return jsonOut_({ ok: false, error: 'senha de administrador inválida' });
      }
      var novosGrupos = (body.grupos || []).map(function (t) {
        return {
          nome: String(t.nome || '').trim(),
          cultura: String(t.cultura || '').trim(),
          valvulas: Array.isArray(t.valvulas) ? t.valvulas.map(function (v) { return String(v).trim(); }).filter(function (v) { return v; }) : [],
          descricao: String(t.descricao || '').trim()
        };
      }).filter(function (t) { return t.nome && t.cultura; });
      if (!novosGrupos.length) {
        return jsonOut_({ ok: false, error: 'lista de grupos vazia' });
      }
      PropertiesService.getScriptProperties().setProperty(GRUPOS_IRRIG_PROP_KEY, JSON.stringify(novosGrupos));
      return jsonOut_({ ok: true, gruposIrrigacao: novosGrupos });
    }

    if (action === 'update_setores_irrigacao') {
      if (body.adminToken !== ADMIN_SECRET) {
        return jsonOut_({ ok: false, error: 'senha de administrador inválida' });
      }
      var novosSetores = (body.setores || []).map(function (s) {
        return {
          nome: String(s.nome || '').trim(),
          grupo: String(s.grupo || '').trim(),
          talhaoIntegrado: String(s.talhaoIntegrado || '').trim()
        };
      }).filter(function (s) { return s.nome && s.grupo; });
      PropertiesService.getScriptProperties().setProperty(SETORES_IRRIG_PROP_KEY, JSON.stringify(novosSetores));
      return jsonOut_({ ok: true, setoresIrrigacao: novosSetores });
    }

    if (action === 'update_turmas_coco') {
      if (body.adminToken !== ADMIN_SECRET) {
        return jsonOut_({ ok: false, error: 'senha de administrador inválida' });
      }
      var novasTurmas = (body.turmas || []).map(function (t) {
        return {
          nome: String(t.nome || '').trim(),
          valvulas: Array.isArray(t.valvulas) ? t.valvulas.map(function (v) { return String(v).trim(); }).filter(function (v) { return v; }) : []
        };
      }).filter(function (t) { return t.nome; });
      if (!novasTurmas.length) {
        return jsonOut_({ ok: false, error: 'lista de turmas vazia' });
      }
      PropertiesService.getScriptProperties().setProperty(TURMAS_COCO_PROP_KEY, JSON.stringify(novasTurmas));
      return jsonOut_({ ok: true, turmasCoco: novasTurmas });
    }

    if (action === 'update_parametros_irrigacao') {
      if (body.adminToken !== ADMIN_SECRET) {
        return jsonOut_({ ok: false, error: 'senha de administrador inválida' });
      }
      var novosParams = body.parametros && typeof body.parametros === 'object' ? body.parametros : {};
      PropertiesService.getScriptProperties().setProperty(PARAMS_IRRIG_PROP_KEY, JSON.stringify(novosParams));
      return jsonOut_({ ok: true, parametrosIrrigacao: novosParams });
    }

    if (action === 'delete_manutencao') {
      if (body.adminToken !== ADMIN_SECRET) {
        return jsonOut_({ ok: false, error: 'senha de administrador inválida' });
      }
      var msheet2 = getManutSheet_();
      var idsM = msheet2.getRange(2, 1, Math.max(msheet2.getLastRow() - 1, 0), 1).getValues();
      var rowIdxM = -1;
      for (var im = 0; im < idsM.length; im++) {
        if (String(idsM[im][0]) === String(body.id)) { rowIdxM = im + 2; break; }
      }
      // Antes isso sempre respondia ok:true mesmo sem achar a linha - fazia o
      // app "pensar" que apagou quando na verdade não apagou nada, e o registro
      // voltava sozinho na sincronização seguinte.
      if (rowIdxM === -1) {
        return jsonOut_({ ok: false, error: 'registro de manutenção não encontrado (id: ' + body.id + ')' });
      }
      msheet2.getRange(rowIdxM, 6).setValue(true); // coluna 6 = excluido
      return jsonOut_({ ok: true });
    }

    if (action === 'delete_bombeamento') {
      if (body.adminToken !== ADMIN_SECRET) {
        return jsonOut_({ ok: false, error: 'senha de administrador inválida' });
      }
      var bsheet2 = getBombSheet_();
      var idsB = bsheet2.getRange(2, 1, Math.max(bsheet2.getLastRow() - 1, 0), 1).getValues();
      var rowIdxB = -1;
      for (var ib = 0; ib < idsB.length; ib++) {
        if (String(idsB[ib][0]) === String(body.id)) { rowIdxB = ib + 2; break; }
      }
      if (rowIdxB === -1) {
        return jsonOut_({ ok: false, error: 'registro de bombeamento não encontrado (id: ' + body.id + ')' });
      }
      bsheet2.getRange(rowIdxB, 12).setValue(true); // coluna 12 = excluido
      return jsonOut_({ ok: true });
    }

    if (action === 'update_rondas_config') {
      if (body.adminToken !== ADMIN_SECRET) {
        return jsonOut_({ ok: false, error: 'senha de administrador inválida' });
      }
      var novosHorarios = (body.horarios || []).map(function (h) { return String(h).trim(); })
        .filter(function (h) { return /^\d{1,2}:\d{2}$/.test(h); });
      if (!novosHorarios.length) {
        return jsonOut_({ ok: false, error: 'informe ao menos um horário válido (formato HH:MM)' });
      }
      var novosPaineis = (body.paineis || []).map(function (p) {
        return { bomba: parseInt(p.bomba, 10) || 0,
                 tipo: (String(p.tipo || '').toLowerCase() === 'inversor') ? 'inversor' : 'soft' };
      }).filter(function (p) { return p.bomba > 0; });
      var props = PropertiesService.getScriptProperties();
      props.setProperty(RONDAS_HORARIOS_PROP_KEY, JSON.stringify(novosHorarios));
      if (novosPaineis.length) props.setProperty(PAINEIS_BOMBAS_PROP_KEY, JSON.stringify(novosPaineis));
      var novosPlant = (body.plantonistas || []).map(function (n) { return String(n).trim(); }).filter(function (n) { return n; });
      if (novosPlant.length) props.setProperty(PLANTONISTAS_PROP_KEY, JSON.stringify(novosPlant));
      return jsonOut_({ ok: true, horarios: novosHorarios,
        paineis: novosPaineis.length ? novosPaineis : getPaineisBombas_(),
        plantonistas: novosPlant.length ? novosPlant : getPlantonistas_() });
    }

    if (action === 'delete_ronda') {
      if (body.adminToken !== ADMIN_SECRET) {
        return jsonOut_({ ok: false, error: 'senha de administrador inválida' });
      }
      var rsheetD = getRondaSheet_();
      var idsR = rsheetD.getRange(2, 1, Math.max(rsheetD.getLastRow() - 1, 0), 1).getValues();
      var rowIdxR = -1;
      for (var ir = 0; ir < idsR.length; ir++) {
        if (String(idsR[ir][0]) === String(body.id)) { rowIdxR = ir + 2; break; }
      }
      if (rowIdxR === -1) return jsonOut_({ ok: false, error: 'ronda não encontrada (id: ' + body.id + ')' });
      rsheetD.getRange(rowIdxR, 17).setValue(true); // coluna 17 = excluido
      return jsonOut_({ ok: true });
    }

    if (action === 'delete_irrigacao') {
      if (body.adminToken !== ADMIN_SECRET) {
        return jsonOut_({ ok: false, error: 'senha de administrador inválida' });
      }
      var isheet2 = getIrrigSheet_();
      var idsI = isheet2.getRange(2, 1, Math.max(isheet2.getLastRow() - 1, 0), 1).getValues();
      var linhasEncontradasI = [];
      for (var ii = 0; ii < idsI.length; ii++) {
        if (String(idsI[ii][0]) === String(body.id)) { linhasEncontradasI.push(ii + 2); }
      }
      // Idem: sem esse if, respondia ok:true mesmo não achando a linha, e o
      // app removia da tela um registro que continuava intacto na planilha -
      // reaparecia na sincronização seguinte.
      if (!linhasEncontradasI.length) {
        return jsonOut_({ ok: false, error: 'registro de irrigação não encontrado (id: ' + body.id + ')' });
      }
      linhasEncontradasI.forEach(function (rowIdxI) {
        isheet2.getRange(rowIdxI, IRRIG_EXCLUIDO_COL).setValue(true); // coluna 27 = excluido
      });
      syncIrrigTabsPorCultura_();
      return jsonOut_({ ok: true });
    }

    // Remover um Setor da Programação Cacau do dia a dia não exige senha de
    // administrador (mesmo padrão de create_irrigacao/update_irrigacao) -
    // é uso normal e frequente, diferente de apagar histórico já fechado.
    if (action === 'delete_irrigacao_sessao') {
      var isheetDS = getIrrigSheet_();
      var idsDS = isheetDS.getRange(2, 1, Math.max(isheetDS.getLastRow() - 1, 0), 1).getValues();
      var linhasEncontradasDS = [];
      for (var ids = 0; ids < idsDS.length; ids++) {
        if (String(idsDS[ids][0]) === String(body.id)) { linhasEncontradasDS.push(ids + 2); }
      }
      // Essa era a principal causa do "Setor apagado sempre volta": respondia
      // ok:true mesmo quando a linha não existia na planilha (ex.: o "criar"
      // daquele setor ainda não tinha terminado de sincronizar quando o
      // pedido de exclusão chegou). O app confiava no ok:true, tirava da
      // tela, mas a linha continuava lá - e voltava a cada sincronização.
      if (!linhasEncontradasDS.length) {
        return jsonOut_({ ok: false, notFound: true, error: 'sessão de irrigação não encontrada (id: ' + body.id + ')' });
      }
      // Marca TODAS as linhas com esse id, não só a primeira - existem
      // linhas antigas duplicadas (mesmo id repetido, geralmente uma versão
      // já excluída e outra não), e marcar só a primeira deixava a segunda
      // cópia visível pra sempre, parecendo que o apagar não funcionava.
      linhasEncontradasDS.forEach(function (rowIdxDS) {
        isheetDS.getRange(rowIdxDS, IRRIG_EXCLUIDO_COL).setValue(true);
      });
      syncIrrigTabsPorCultura_();
      return jsonOut_({ ok: true });
    }

    if (action === 'delete') {
      if (body.adminToken !== ADMIN_SECRET) {
        return jsonOut_({ ok: false, error: 'senha de administrador inválida' });
      }
      var sheetDel = getSheet_();
      var rowIdxDel = findRowById_(sheetDel, body.id);
      if (rowIdxDel === -1) {
        return jsonOut_({ ok: false, error: 'ocorrência não encontrada (id: ' + body.id + ')' });
      }
      sheetDel.getRange(rowIdxDel, 11).setValue(true); // coluna 11 = excluido
      return jsonOut_({ ok: true });
    }

    if (body.token !== SECRET) {
      return jsonOut_({ ok: false, error: 'token inválido' });
    }
    var sheet = getSheet_();

    if (action === 'create_manutencao') {
      var msheet = getManutSheet_();
      msheet.appendRow([body.id, body.setor, body.data, body.encarregado, body.observacao || '', false, new Date()]);
      return jsonOut_({ ok: true });
    }

    if (action === 'create_atividade_prog') {
      var pasheetC = getProgAtivSheet_();
      pasheetC.appendRow([
        body.id, body.data, body.funcionario, body.atividade, body.setor || '', body.duracaoMin || 0, body.areaHa || '',
        body.horarioInicio || '', body.horarioFim || '', body.observacao || '', !!body.concluida, false, new Date()
      ]);
      return jsonOut_({ ok: true });
    }

    if (action === 'update_atividade_prog') {
      var pasheetU = getProgAtivSheet_();
      var paidsU = pasheetU.getRange(2, 1, Math.max(pasheetU.getLastRow() - 1, 0), 1).getValues();
      var parowU = -1;
      for (var paiU = 0; paiU < paidsU.length; paiU++) {
        if (String(paidsU[paiU][0]) === String(body.id)) { parowU = paiU + 2; break; }
      }
      if (parowU === -1) return jsonOut_({ ok: false, notFound: true, error: 'atividade não encontrada para atualizar' });
      pasheetU.getRange(parowU, 1, 1, 11).setValues([[
        body.id, body.data, body.funcionario, body.atividade, body.setor || '', body.duracaoMin || 0, body.areaHa || '',
        body.horarioInicio || '', body.horarioFim || '', body.observacao || '', !!body.concluida
      ]]);
      return jsonOut_({ ok: true });
    }

    if (action === 'create_rendimento_diario') {
      var rdsheetC = getRendDiaSheet_();
      rdsheetC.appendRow([body.id, body.data, body.funcionario, body.rendimento || '', false, new Date()]);
      return jsonOut_({ ok: true });
    }

    if (action === 'update_rendimento_diario') {
      var rdsheetU = getRendDiaSheet_();
      var rdidsU = rdsheetU.getRange(2, 1, Math.max(rdsheetU.getLastRow() - 1, 0), 1).getValues();
      var rdrowU = -1;
      for (var rdiU = 0; rdiU < rdidsU.length; rdiU++) {
        if (String(rdidsU[rdiU][0]) === String(body.id)) { rdrowU = rdiU + 2; break; }
      }
      if (rdrowU === -1) return jsonOut_({ ok: false, notFound: true, error: 'rendimento não encontrado para atualizar' });
      rdsheetU.getRange(rdrowU, 1, 1, 4).setValues([[body.id, body.data, body.funcionario, body.rendimento || '']]);
      return jsonOut_({ ok: true });
    }

    if (action === 'create_banco_horas') {
      var bhsheetC = getBancoHorasSheet_();
      bhsheetC.appendRow([body.id, body.data, body.funcionario, body.tipo || '', body.minutos || 0, body.observacao || '', false, new Date()]);
      return jsonOut_({ ok: true });
    }

    if (action === 'update_banco_horas') {
      var bhsheetU = getBancoHorasSheet_();
      var bhidsU = bhsheetU.getRange(2, 1, Math.max(bhsheetU.getLastRow() - 1, 0), 1).getValues();
      var bhrowU = -1;
      for (var bhiU = 0; bhiU < bhidsU.length; bhiU++) {
        if (String(bhidsU[bhiU][0]) === String(body.id)) { bhrowU = bhiU + 2; break; }
      }
      if (bhrowU === -1) return jsonOut_({ ok: false, notFound: true, error: 'lançamento não encontrado para atualizar' });
      bhsheetU.getRange(bhrowU, 1, 1, 6).setValues([[body.id, body.data, body.funcionario, body.tipo || '', body.minutos || 0, body.observacao || '']]);
      return jsonOut_({ ok: true });
    }

    if (action === 'create_ronda') {
      var rsheetC = getRondaSheet_();
      if (findRowById_(rsheetC, body.id) > -1) return jsonOut_({ ok: true, duplicado: true });
      var urlPainel = body.fotoPainel
        ? (String(body.fotoPainel).indexOf('data:image') === 0
            ? saveFotoBomb_(body.fotoPainel, body.id + '_painel_' + new Date().getTime())
            : body.fotoPainel)
        : '';
      rsheetC.appendRow([
        body.id, body.data, body.horarioPrevisto || '', body.horarioReal || '', body.plantonista || '',
        body.bomba, !!body.bombaLigada, body.nivelReservatorio || '', body.painelTipo || '',
        body.painelStatus || '', body.corrente || '', body.tensao || '', body.horimetro || '',
        urlPainel, body.motivoParada || '', body.observacao || '', false, new Date()
      ]);
      return jsonOut_({ ok: true, fotoPainelUrl: urlPainel });
    }

    if (action === 'update_ronda') {
      var rsheetU = getRondaSheet_();
      var rowIdxRU = findRowById_(rsheetU, body.id);
      if (rowIdxRU === -1) return jsonOut_({ ok: false, notFound: true, error: 'ronda não encontrada para atualizar' });
      var urlPainelU = body.fotoPainel !== undefined
        ? (body.fotoPainel
            ? (String(body.fotoPainel).indexOf('data:image') === 0
                ? saveFotoBomb_(body.fotoPainel, body.id + '_painel_' + new Date().getTime())
                : body.fotoPainel)
            : '')
        : rsheetU.getRange(rowIdxRU, 14).getValue();
      rsheetU.getRange(rowIdxRU, 1, 1, 16).setValues([[
        body.id, body.data, body.horarioPrevisto || '', body.horarioReal || '', body.plantonista || '',
        body.bomba, !!body.bombaLigada, body.nivelReservatorio || '', body.painelTipo || '',
        body.painelStatus || '', body.corrente || '', body.tensao || '', body.horimetro || '',
        urlPainelU, body.motivoParada || '', body.observacao || ''
      ]]);
      return jsonOut_({ ok: true, fotoPainelUrl: urlPainelU });
    }

    if (action === 'create_bombeamento') {
      var bsheet = getBombSheet_();
      // evita duplicar se o app reenviar o mesmo registro por engano
      if (findRowById_(bsheet, body.id) > -1) {
        return jsonOut_({ ok: true, duplicado: true });
      }
      var urlInicioB = body.fotoInicio ? (String(body.fotoInicio).indexOf('data:image') === 0 ? saveFotoBomb_(body.fotoInicio, body.id + '_inicio_' + new Date().getTime()) : body.fotoInicio) : '';
      var urlFimB = body.fotoFim ? (String(body.fotoFim).indexOf('data:image') === 0 ? saveFotoBomb_(body.fotoFim, body.id + '_fim_' + new Date().getTime()) : body.fotoFim) : '';
      var temAmbasLeiturasB = body.leituraInicial !== '' && body.leituraInicial !== undefined && body.leituraInicial !== null &&
        body.leituraFinal !== '' && body.leituraFinal !== undefined && body.leituraFinal !== null;
      var consumoB = temAmbasLeiturasB ? (parseFloat(body.leituraFinal) - parseFloat(body.leituraInicial)) * 10 : '';
      bsheet.appendRow([
        body.id, body.data, body.bomba, body.horarioInicio || '', body.horarioFim || '',
        body.leituraInicial !== undefined && body.leituraInicial !== '' ? body.leituraInicial : '',
        body.leituraFinal !== undefined && body.leituraFinal !== '' ? body.leituraFinal : '',
        consumoB, urlInicioB, urlFimB, body.observacao || '', false, new Date()
      ]);
      // responsavel fica numa coluna no fim da planilha - gravada separada
      // justamente pra não deslocar as colunas antigas.
      bsheet.getRange(bsheet.getLastRow(), BOMB_COL_RESPONSAVEL).setValue(body.responsavel || '');
      return jsonOut_({ ok: true, fotoUrlInicio: urlInicioB, fotoUrlFim: urlFimB });
    }

    if (action === 'update_bombeamento') {
      var bsheetU = getBombSheet_();
      var rowIdxBU = findRowById_(bsheetU, body.id);
      if (rowIdxBU === -1) return jsonOut_({ ok: false, notFound: true, error: 'registro de bombeamento não encontrado' });
      var temAmbasLeiturasU = body.leituraInicial !== '' && body.leituraInicial !== undefined && body.leituraInicial !== null &&
        body.leituraFinal !== '' && body.leituraFinal !== undefined && body.leituraFinal !== null;
      var consumoU = temAmbasLeiturasU ? (parseFloat(body.leituraFinal) - parseFloat(body.leituraInicial)) * 10 : '';
      var urlInicioU = body.fotoInicio !== undefined
        ? (body.fotoInicio ? (String(body.fotoInicio).indexOf('data:image') === 0 ? saveFotoBomb_(body.fotoInicio, body.id + '_inicio_' + new Date().getTime()) : body.fotoInicio) : '')
        : bsheetU.getRange(rowIdxBU, 9).getValue();
      var urlFimU = body.fotoFim !== undefined
        ? (body.fotoFim ? (String(body.fotoFim).indexOf('data:image') === 0 ? saveFotoBomb_(body.fotoFim, body.id + '_fim_' + new Date().getTime()) : body.fotoFim) : '')
        : bsheetU.getRange(rowIdxBU, 10).getValue();
      bsheetU.getRange(rowIdxBU, 1, 1, 11).setValues([[
        body.id, body.data, body.bomba, body.horarioInicio || '', body.horarioFim || '',
        body.leituraInicial !== undefined && body.leituraInicial !== '' ? body.leituraInicial : '',
        body.leituraFinal !== undefined && body.leituraFinal !== '' ? body.leituraFinal : '',
        consumoU, urlInicioU, urlFimU, body.observacao || ''
      ]]);
      if (body.responsavel !== undefined) {
        bsheetU.getRange(rowIdxBU, BOMB_COL_RESPONSAVEL).setValue(body.responsavel || '');
      }
      return jsonOut_({ ok: true, fotoUrlInicio: urlInicioU, fotoUrlFim: urlFimU });
    }

    if (action === 'create_irrigacao') {
      var isheet = getIrrigSheet_();
      // Evita duplicar se o app reenviar o mesmo registro por engano (ex.: a
      // recriação automática de uma sessão que "sumiu" mandar o create de
      // novo bem na hora que a linha já tinha voltado a existir).
      var idsCI = isheet.getRange(2, 1, Math.max(isheet.getLastRow() - 1, 0), 1).getValues();
      for (var ici = 0; ici < idsCI.length; ici++) {
        if (String(idsCI[ici][0]) === String(body.id)) {
          return jsonOut_({ ok: true, duplicado: true });
        }
      }
      isheet.appendRow([
        body.id, body.data, body.cultura, body.grupo || '', body.setor || '', body.talhaoIntegrado || '', body.modo || '',
        body.eto, body.kcUsado, body.klUsado || '', body.laminaMm,
        body.tempoRecomendadoMin || '', body.litrosPlanta || '', body.m3Valvula || '',
        body.ordem !== undefined ? body.ordem : '', body.horario || '', body.horarioFim || '', body.tempoMin || '',
        body.tanqueA || 0, body.tanqueB || 0, body.tanqueC || 0, body.tanqueD || 0, body.tanqueE || 0, body.tanqueF || 0,
        !!body.finalizada, body.observacao || '', false, new Date()
      ]);
      syncIrrigTabsPorCultura_();
      return jsonOut_({ ok: true });
    }

    // Permite editar uma sessão de irrigação que já tinha sido sincronizada
    // (setor, horário, duração, tanques, etc.) - identifica a linha pelo id.
    if (action === 'update_irrigacao') {
      var isheetU = getIrrigSheet_();
      var idsU = isheetU.getRange(2, 1, Math.max(isheetU.getLastRow() - 1, 0), 1).getValues();
      var rowIdxU = -1;
      // Usa a ÚLTIMA linha com esse id, não a primeira - quando existe uma
      // duplicata antiga (mesmo id, uma cópia já excluída e outra não), a
      // mais recente é sempre a de verdade; atualizar a primeira por engano
      // podia mexer na cópia errada (já excluída).
      for (var iu = 0; iu < idsU.length; iu++) {
        if (String(idsU[iu][0]) === String(body.id)) { rowIdxU = iu + 2; }
      }
      if (rowIdxU === -1) {
        return jsonOut_({ ok: false, notFound: true, error: 'sessão de irrigação não encontrada para atualizar' });
      }
      isheetU.getRange(rowIdxU, 1, 1, IRRIG_COLUNAS.length).setValues([[
        body.id, body.data, body.cultura, body.grupo || '', body.setor || '', body.talhaoIntegrado || '', body.modo || '',
        body.eto, body.kcUsado, body.klUsado || '', body.laminaMm,
        body.tempoRecomendadoMin || '', body.litrosPlanta || '', body.m3Valvula || '',
        body.ordem !== undefined ? body.ordem : '', body.horario || '', body.horarioFim || '', body.tempoMin || '',
        body.tanqueA || 0, body.tanqueB || 0, body.tanqueC || 0, body.tanqueD || 0, body.tanqueE || 0, body.tanqueF || 0,
        !!body.finalizada, body.observacao || '', false, isheetU.getRange(rowIdxU, IRRIG_COLUNAS.length).getValue() || new Date()
      ]]);
      syncIrrigTabsPorCultura_();
      return jsonOut_({ ok: true });
    }

    if (action === 'create') {
      // evita duplicar se o app reenviar o mesmo registro por engano
      if (findRowById_(sheet, body.id) > -1) {
        return jsonOut_({ ok: true, duplicado: true });
      }
      var urls = processFotos_(body.fotos, body.id);
      sheet.appendRow([
        body.id, body.data, body.setor, body.tipo, body.causa,
        body.encarregado, body.descricao, body.acao || '', body.status || 'Aberto',
        urls.join(FOTO_SEP), false, new Date(), body.cultura || ''
      ]);
      return jsonOut_({ ok: true, fotoUrls: urls });
    }

    if (action === 'update') {
      var rowIdxU = findRowById_(sheet, body.id);
      if (rowIdxU === -1) return jsonOut_({ ok: false, error: 'registro não encontrado' });
      sheet.getRange(rowIdxU, 1, 1, 9).setValues([[
        body.id, body.data, body.setor, body.tipo, body.causa,
        body.encarregado, body.descricao, body.acao || '', body.status || 'Aberto'
      ]]);
      if (body.cultura !== undefined) {
        sheet.getRange(rowIdxU, CULTURA_COL).setValue(body.cultura);
      }
      var updatedUrls;
      if (body.fotos !== undefined) {
        updatedUrls = processFotos_(body.fotos, body.id);
        sheet.getRange(rowIdxU, 10).setValue(updatedUrls.join(FOTO_SEP));
      }
      return jsonOut_({ ok: true, fotoUrls: updatedUrls });
    }

    if (action === 'update_status') {
      var rowIdx = findRowById_(sheet, body.id);
      if (rowIdx > -1) sheet.getRange(rowIdx, 9).setValue(body.status); // coluna 9 = status
      return jsonOut_({ ok: true });
    }

    return jsonOut_({ ok: false, error: 'ação desconhecida: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

// ============================================================
// EMBELEZAMENTO DA PLANILHA - roda uma vez (ou sempre que quiser
// reaplicar a formatação). Acessível pelo menu "⚙️ Registro de
// Ocorrências" que aparece no topo da planilha depois de recarregar.
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Registro de Ocorrências')
    .addItem('Formatar planilha (Ocorrências)', 'formatarPlanilha')
    .addItem('Criar/Atualizar painel Resumo (Ocorrências)', 'criarPainelResumo')
    .addSeparator()
    .addItem('Formatar planilha (Manutenções)', 'formatarPlanilhaManutencao')
    .addItem('Criar/Atualizar painel Manutenção', 'criarPainelManutencao')
    .addSeparator()
    .addItem('Formatar planilha (Irrigação)', 'formatarPlanilhaIrrigacao')
    .addItem('Criar/Atualizar Painel de Irrigação (Cacau) do dia...', 'criarPainelIrrigacaoCacau')
    .addToUi();
}

function formatarPlanilha() {
  var sheet = getSheet_();
  var lastCol = COLUNAS.length;
  var lastRow = Math.max(sheet.getLastRow(), 2);

  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setBackground('#1B3A34').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);

  var widths = { 1: 90, 2: 95, 3: 110, 4: 180, 5: 160, 6: 120, 7: 260, 8: 220, 9: 115, 10: 260, 11: 70, 12: 150 };
  Object.keys(widths).forEach(function (c) { sheet.setColumnWidth(parseInt(c, 10), widths[c]); });

  sheet.getRange(2, 2, lastRow - 1, 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(2, 12, lastRow - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');

  var fullRange = sheet.getRange(1, 1, lastRow, lastCol);
  fullRange.getBandings().forEach(function (b) { b.remove(); });
  fullRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);

  var statusRange = sheet.getRange(2, 9, Math.max(lastRow - 1, 1), 1);
  var rules = sheet.getConditionalFormatRules().filter(function (r) {
    return !r.getRanges().some(function (rg) { return rg.getColumn() === 9; });
  });
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Resolvido').setBackground('#DCEEE1').setFontColor('#2F5C3D').setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Em Andamento').setBackground('#FBF0D9').setFontColor('#8A6414').setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Aberto').setBackground('#F5E3DC').setFontColor('#96351B').setRanges([statusRange]).build());
  sheet.setConditionalFormatRules(rules);

  sheet.showColumns(1, lastCol);
  sheet.hideColumns(1);  // id (técnico)
  sheet.hideColumns(11); // excluido (técnico)

  var existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  sheet.getRange(1, 1, lastRow, lastCol).createFilter();

  SpreadsheetApp.getUi().alert('Planilha formatada! Pode rodar de novo quando quiser reaplicar.');
}

function criarPainelResumo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var old = ss.getSheetByName('Resumo');
  if (old) ss.deleteSheet(old);
  var resumo = ss.insertSheet('Resumo');
  resumo.setTabColor('#C1501F');

  resumo.getRange('B2').setValue('Painel de Ocorrências').setFontSize(18).setFontWeight('bold').setFontColor('#1B3A34');
  resumo.getRange('B3').setValue('Se adapta sozinho aos tipos/causas cadastrados. Reabra esta aba para recalcular.')
    .setFontColor('#6B6153').setFontStyle('italic');

  resumo.getRange('B5').setValue('Total de Ocorrências').setFontWeight('bold');
  resumo.getRange('C5').setFormula('=COUNTIFS(Ocorrencias!G2:G;"<>";Ocorrencias!K2:K;"<>TRUE")').setFontWeight('bold').setFontSize(14);

  resumo.getRange('B7').setValue('Por Tipo').setFontWeight('bold');
  resumo.getRange('E7').setValue('Por Causa Provável').setFontWeight('bold');
  resumo.getRange('H7').setValue('Por Status').setFontWeight('bold');

  resumo.getRange('B8').setFormula('=IFERROR(UNIQUE(FILTER(Ocorrencias!D2:D;Ocorrencias!D2:D<>"";Ocorrencias!K2:K<>TRUE));"")');
  resumo.getRange('C8').setFormula('=ARRAYFORMULA(IF(B8:B27="";"";COUNTIFS(Ocorrencias!D2:D;B8:B27;Ocorrencias!K2:K;"<>TRUE")))');

  resumo.getRange('E8').setFormula('=IFERROR(UNIQUE(FILTER(Ocorrencias!E2:E;Ocorrencias!E2:E<>"";Ocorrencias!K2:K<>TRUE));"")');
  resumo.getRange('F8').setFormula('=ARRAYFORMULA(IF(E8:E27="";"";COUNTIFS(Ocorrencias!E2:E;E8:E27;Ocorrencias!K2:K;"<>TRUE")))');

  resumo.getRange('H8').setFormula('=IFERROR(UNIQUE(FILTER(Ocorrencias!I2:I;Ocorrencias!I2:I<>"";Ocorrencias!K2:K<>TRUE));"")');
  resumo.getRange('I8').setFormula('=ARRAYFORMULA(IF(H8:H27="";"";COUNTIFS(Ocorrencias!I2:I;H8:H27;Ocorrencias!K2:K;"<>TRUE")))');

  resumo.setColumnWidths(2, 1, 190);
  resumo.setColumnWidths(3, 1, 70);
  resumo.setColumnWidths(5, 1, 190);
  resumo.setColumnWidths(6, 1, 70);
  resumo.setColumnWidths(8, 1, 150);
  resumo.setColumnWidths(9, 1, 70);

  SpreadsheetApp.flush();

  var chart1 = resumo.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(resumo.getRange('B8:C27'))
    .setPosition(30, 2, 0, 0)
    .setOption('title', 'Ocorrências por Tipo')
    .setOption('legend', { position: 'none' })
    .setOption('colors', ['#2E6690'])
    .setOption('width', 480)
    .setOption('height', 280)
    .build();
  resumo.insertChart(chart1);

  var chart2 = resumo.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(resumo.getRange('E8:F27'))
    .setPosition(30, 6, 0, 0)
    .setOption('title', 'Ocorrências por Causa Provável')
    .setOption('width', 480)
    .setOption('height', 280)
    .build();
  resumo.insertChart(chart2);

  var chart3 = resumo.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(resumo.getRange('H8:I27'))
    .setPosition(52, 2, 0, 0)
    .setOption('title', 'Ocorrências por Status')
    .setOption('legend', { position: 'none' })
    .setOption('colors', ['#C1501F'])
    .setOption('width', 480)
    .setOption('height', 280)
    .build();
  resumo.insertChart(chart3);

  ss.setActiveSheet(resumo);
  SpreadsheetApp.getUi().alert('Painel Resumo criado! Já tem gráficos de Tipo, Causa e Status.');
}

// ============================================================
// EMBELEZAMENTO DA PLANILHA DE MANUTENÇÕES
// ============================================================

function formatarPlanilhaManutencao() {
  var sheet = getManutSheet_();
  var lastCol = MANUT_COLUNAS.length; // id, setor, data, encarregado, observacao, excluido, criado_em
  var lastRow = Math.max(sheet.getLastRow(), 2);

  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setBackground('#2F7DB0').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);

  var widths = { 1: 90, 2: 110, 3: 95, 4: 130, 5: 260, 6: 70, 7: 150 };
  Object.keys(widths).forEach(function (c) { sheet.setColumnWidth(parseInt(c, 10), widths[c]); });

  sheet.getRange(2, 3, lastRow - 1, 1).setNumberFormat('dd/mm/yyyy'); // data (coluna C)
  sheet.getRange(2, 7, lastRow - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm'); // criado_em (coluna G)

  var fullRange = sheet.getRange(1, 1, lastRow, lastCol);
  fullRange.getBandings().forEach(function (b) { b.remove(); });
  fullRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);

  sheet.showColumns(1, lastCol);
  sheet.hideColumns(1); // id (técnico)
  sheet.hideColumns(6); // excluido (técnico)

  var existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  sheet.getRange(1, 1, lastRow, lastCol).createFilter();

  SpreadsheetApp.getUi().alert('Planilha de Manutenções formatada!');
}

// Painel que mostra, para cada talhão, a última manutenção e a próxima data
// prevista de reentrada - calculado direto (não é fórmula, então rode de novo
// sempre que quiser atualizar os números).
function criarPainelManutencao() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var old = ss.getSheetByName('Painel Manutenção');
  if (old) ss.deleteSheet(old);
  var painel = ss.insertSheet('Painel Manutenção');
  painel.setTabColor('#2F7DB0');

  var culturas = getCulturas_();
  var talhoes = getTalhoes_();
  var culturaByNome = {};
  culturas.forEach(function (c) { culturaByNome[c.nome] = c; });

  var msheet = getManutSheet_();
  var mdata = msheet.getDataRange().getValues();
  mdata.shift(); // remove cabeçalho
  var porTalhao = {}; // nome -> maior data (Date)
  mdata.forEach(function (r) {
    var setor = r[1], dataStr = r[2], excluido = r[5];
    if (excluido) return;
    var d = parseDataFlexivel_(dataStr);
    if (!d) return;
    if (!porTalhao[setor] || d.getTime() > porTalhao[setor].getTime()) {
      porTalhao[setor] = d;
    }
  });

  painel.getRange('B2').setValue('Painel de Manutenção Preventiva').setFontSize(18).setFontWeight('bold').setFontColor('#1B3A34');
  painel.getRange('B3').setValue('Gerado em ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + ' — rode de novo pra atualizar.')
    .setFontColor('#6B6153').setFontStyle('italic');

  var hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  var linhas = talhoes.map(function (t) {
    var cultura = culturaByNome[t.cultura] || { intervalo: 15 };
    var ultima = porTalhao[t.nome] || null;
    var proxima = ultima ? addDias_(ultima, cultura.intervalo) : null;
    var diasAte = proxima ? Math.round((proxima.getTime() - hoje.getTime()) / 86400000) : null;
    var situacao, ordem;
    if (!ultima) { situacao = 'Nunca registrada'; ordem = 2; }
    else if (diasAte < 0) { situacao = 'Atrasado ' + Math.abs(diasAte) + ' dia(s)'; ordem = 0; }
    else if (diasAte <= 3) { situacao = 'Próximo do prazo'; ordem = 1; }
    else { situacao = 'Em dia'; ordem = 3; }
    return { cultura: t.cultura, talhao: t.nome, intervalo: cultura.intervalo, ultima: ultima, proxima: proxima, situacao: situacao, ordem: ordem };
  });
  linhas.sort(function (a, b) { return a.ordem - b.ordem; });

  var qtdAtrasado = linhas.filter(function (l) { return l.ordem === 0; }).length;
  var qtdProximo = linhas.filter(function (l) { return l.ordem === 1; }).length;
  var qtdNunca = linhas.filter(function (l) { return l.ordem === 2; }).length;
  var qtdEmDia = linhas.filter(function (l) { return l.ordem === 3; }).length;

  // ---- KPIs (números grandes no topo) ----
  var kpis = [
    { label: 'Total de Talhões', valor: linhas.length, col: 2, cor: '#1B3A34' },
    { label: 'Atrasados', valor: qtdAtrasado, col: 4, cor: '#C1501F' },
    { label: 'Próx. do Prazo', valor: qtdProximo, col: 6, cor: '#D6A233' },
    { label: 'Em Dia', valor: qtdEmDia, col: 8, cor: '#4C7C59' }
  ];
  kpis.forEach(function (k) {
    painel.getRange(5, k.col).setValue(k.valor).setFontSize(26).setFontWeight('bold').setFontColor(k.cor).setHorizontalAlignment('center');
    painel.getRange(6, k.col).setValue(k.label).setFontColor('#6B6153').setFontSize(10).setHorizontalAlignment('center');
  });

  // ---- Tabela de detalhe por talhão ----
  var headers = ['Cultura', 'Talhão', 'Intervalo (dias)', 'Última Manutenção', 'Próxima Entrada', 'Situação'];
  var headerRow = 9;
  painel.getRange(headerRow, 2, 1, headers.length).setValues([headers])
    .setBackground('#2F7DB0').setFontColor('#FFFFFF').setFontWeight('bold');

  var rows = linhas.map(function (l) {
    return [l.cultura, l.talhao, l.intervalo, l.ultima || '—', l.proxima || '—', l.situacao];
  });
  if (rows.length) {
    painel.getRange(headerRow + 1, 2, rows.length, headers.length).setValues(rows);
    painel.getRange(headerRow + 1, 5, rows.length, 1).setNumberFormat('dd/mm/yyyy');
    painel.getRange(headerRow + 1, 6, rows.length, 1).setNumberFormat('dd/mm/yyyy');

    // cor de fundo por situação
    for (var i = 0; i < linhas.length; i++) {
      var rowNum = headerRow + 1 + i;
      var color = linhas[i].ordem === 0 ? '#F5E3DC' : linhas[i].ordem === 1 ? '#FBF0D9' : linhas[i].ordem === 2 ? '#EFEAE0' : '#DCEEE1';
      painel.getRange(rowNum, 2, 1, headers.length).setBackground(color);
    }
  }

  // ---- Tabela auxiliar (escondida) pra alimentar o gráfico de situação ----
  var chartDataRow = headerRow + rows.length + 4;
  painel.getRange(chartDataRow, 2).setValue('Situação');
  painel.getRange(chartDataRow, 3).setValue('Quantidade');
  var situacaoRows = [
    ['Atrasado', qtdAtrasado],
    ['Próximo do prazo', qtdProximo],
    ['Nunca registrada', qtdNunca],
    ['Em dia', qtdEmDia]
  ];
  painel.getRange(chartDataRow + 1, 2, situacaoRows.length, 2).setValues(situacaoRows);

  var chart = painel.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(painel.getRange(chartDataRow, 2, situacaoRows.length + 1, 2))
    .setPosition(headerRow + rows.length + 3, 2, 0, 0)
    .setOption('title', 'Distribuição de Situação dos Talhões')
    .setOption('colors', ['#C1501F', '#D6A233', '#8A8072', '#4C7C59'])
    .setOption('width', 480)
    .setOption('height', 280)
    .build();
  painel.insertChart(chart);

  painel.setColumnWidth(2, 110);
  painel.setColumnWidth(3, 90);
  painel.setColumnWidth(4, 110);
  painel.setColumnWidth(5, 130);
  painel.setColumnWidth(6, 130);
  painel.setColumnWidth(7, 150);
  painel.setColumnWidth(8, 110);
  painel.setFrozenRows(headerRow);

  ss.setActiveSheet(painel);
  SpreadsheetApp.getUi().alert('Painel de Manutenção atualizado!');
}

function parseDataFlexivel_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return v;
  var s = String(v);
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  var m2 = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (m2) return new Date(parseInt(m2[3], 10), parseInt(m2[2], 10) - 1, parseInt(m2[1], 10));
  return null;
}

function addDias_(date, dias) {
  var d = new Date(date.getTime());
  d.setDate(d.getDate() + dias);
  return d;
}

function mesmoDia_(d1, d2) {
  if (!d1 || !d2) return false;
  return d1.getFullYear() === d2.getFullYear() && d1.getMonth() === d2.getMonth() && d1.getDate() === d2.getDate();
}

// ============================================================
// EMBELEZAMENTO DA PLANILHA DE IRRIGAÇÃO
// ============================================================

function formatarPlanilhaIrrigacao() {
  var sheet = getIrrigSheet_();
  var lastCol = IRRIG_COLUNAS.length;
  var lastRow = Math.max(sheet.getLastRow(), 2);

  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setBackground('#1F6B44').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2); // mantém id (escondido) + data visíveis ao rolar pro lado

  // larguras por nome de coluna, na ordem de IRRIG_COLUNAS
  var largurasPorNome = {
    id: 90, data: 95, cultura: 85, grupo: 75, setor: 110, talhao_integrado: 95,
    modo: 100, eto: 55, kc_usado: 65, kl_usado: 65, lamina_mm: 70,
    tempo_recomendado_min: 105, litros_planta: 85, m3_valvula: 85,
    ordem: 55, horario: 65, horario_fim: 65, tempo_min: 75,
    tanque_a: 65, tanque_b: 65, tanque_c: 65, tanque_d: 65, tanque_e: 65, tanque_f: 65,
    finalizada: 75, observacao: 220, excluido: 70, criado_em: 140
  };
  IRRIG_COLUNAS.forEach(function (nome, i) {
    if (largurasPorNome[nome]) sheet.setColumnWidth(i + 1, largurasPorNome[nome]);
  });

  var colData = IRRIG_COLUNAS.indexOf('data') + 1;
  var colCriadoEm = IRRIG_COLUNAS.indexOf('criado_em') + 1;
  sheet.getRange(2, colCriadoEm, lastRow - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');

  var fullRange = sheet.getRange(1, 1, lastRow, lastCol);
  fullRange.getBandings().forEach(function (b) { b.remove(); });
  fullRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);

  // Cor de fundo por Grupo (mesmas cores usadas no relatório em PDF)
  var colGrupo = IRRIG_COLUNAS.indexOf('grupo') + 1;
  var grupoRange = sheet.getRange(2, colGrupo, Math.max(lastRow - 1, 1), 1);
  var rules = sheet.getConditionalFormatRules().filter(function (r) {
    return !r.getRanges().some(function (rg) { return rg.getColumn() === colGrupo; });
  });
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('TA1').setBackground('#DCE9F5').setFontColor('#2F5FA8').setRanges([grupoRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('T2+T3').setBackground('#EDE1F5').setFontColor('#7A3FA0').setRanges([grupoRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('T1').setBackground('#DCEEE1').setFontColor('#1B7A4D').setRanges([grupoRange]).build());
  sheet.setConditionalFormatRules(rules);

  sheet.showColumns(1, lastCol);
  sheet.hideColumns(IRRIG_COLUNAS.indexOf('id') + 1);
  sheet.hideColumns(IRRIG_COLUNAS.indexOf('excluido') + 1);

  var existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  sheet.getRange(1, 1, lastRow, lastCol).createFilter();

  SpreadsheetApp.getUi().alert('Planilha de Irrigação formatada! Pode rodar de novo quando quiser reaplicar.');
}

// ============================================================
// PAINEL DE IRRIGAÇÃO (CACAU) - mesmo formato do relatório em PDF do
// app, só que direto na planilha. Pede a data, monta a tabela de
// Cálculo da Necessidade Hídrica (por Grupo) e a Programação por Setor
// daquele dia. Não é fórmula viva - rode de novo quando quiser
// atualizar (ex.: depois de mudar algo na programação daquele dia).
// ============================================================

function criarPainelIrrigacaoCacau() {
  var ui = SpreadsheetApp.getUi();
  var hoje = new Date();
  var sugestao = Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var resp = ui.prompt('Painel de Irrigação (Cacau)', 'Data da programação (dd/mm/aaaa):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var dataEscolhida = parseDataFlexivel_(resp.getResponseText().trim() || sugestao);
  if (!dataEscolhida) { ui.alert('Data inválida. Use o formato dd/mm/aaaa.'); return; }

  var isheet = getIrrigSheet_();
  var lastRow = isheet.getLastRow();
  var dados = lastRow > 1 ? isheet.getRange(2, 1, lastRow - 1, IRRIG_COLUNAS.length).getValues() : [];
  var idx = {};
  IRRIG_COLUNAS.forEach(function (nome, i) { idx[nome] = i; });

  var sessoes = dados.filter(function (r) {
    return r[idx.cultura] === 'Cacau' && !r[idx.excluido] && mesmoDia_(parseDataFlexivel_(r[idx.data]), dataEscolhida);
  }).sort(function (a, b) { return (a[idx.ordem] || 0) - (b[idx.ordem] || 0); });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var old = ss.getSheetByName('Painel Irrigação');
  if (old) ss.deleteSheet(old);
  var painel = ss.insertSheet('Painel Irrigação');
  painel.setTabColor('#1F6B44');

  var dataFmt = Utilities.formatDate(dataEscolhida, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  painel.getRange('B2').setValue('Programação de Irrigação · Cacau').setFontSize(18).setFontWeight('bold').setFontColor('#164D31');
  painel.getRange('B3').setValue(dataFmt + ' — gerado em ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + '. Rode de novo pra atualizar.')
    .setFontColor('#6B6153').setFontStyle('italic');

  if (!sessoes.length) {
    painel.getRange('B5').setValue('Nenhuma sessão de irrigação encontrada para essa data.').setFontColor('#96351B');
    painel.setColumnWidth(2, 320);
    ss.setActiveSheet(painel);
    ui.alert('Não achei nenhuma sessão de Cacau em ' + dataFmt + '.');
    return;
  }

  var CORES_GRUPO = { 'TA1': '#2F5FA8', 'T2+T3': '#7A3FA0', 'T1': '#1B7A4D' };

  // ---- Cálculo da Necessidade Hídrica (uma linha por Grupo do dia) ----
  var gruposDoDia = [];
  sessoes.forEach(function (r) {
    var g = r[idx.grupo] || r[idx.setor];
    if (g && gruposDoDia.indexOf(g) === -1) gruposDoDia.push(g);
  });

  var calcRow = 5;
  painel.getRange(calcRow, 2).setValue('💧 Cálculo da Necessidade Hídrica').setFontWeight('bold').setFontSize(13).setFontColor('#164D31');
  calcRow++;
  var calcHeaders = ['Grupo', 'ET0', 'KL', 'Kc', 'mm/Planta', 'L/Planta'];
  painel.getRange(calcRow, 2, 1, calcHeaders.length).setValues([calcHeaders])
    .setBackground('#1F6B44').setFontColor('#FFFFFF').setFontWeight('bold');
  calcRow++;
  var calcStartRow = calcRow;
  gruposDoDia.forEach(function (g) {
    var linha = sessoes.filter(function (r) { return (r[idx.grupo] || r[idx.setor]) === g; })[0];
    painel.getRange(calcRow, 2, 1, calcHeaders.length).setValues([[
      g, linha[idx.eto] || '', linha[idx.kl_usado] || '', linha[idx.kc_usado] || '',
      linha[idx.lamina_mm] || '', linha[idx.litros_planta] || ''
    ]]);
    painel.getRange(calcRow, 2).setFontColor(CORES_GRUPO[g] || '#5C6B73').setFontWeight('bold');
    calcRow++;
  });
  painel.getRange(calcStartRow, 2, gruposDoDia.length, calcHeaders.length).setBorder(true, true, true, true, true, true, '#DCD3BC', SpreadsheetApp.BorderStyle.SOLID);

  // ---- Programação por Setor ----
  var TODOS_TANQUES = ['tanque_a', 'tanque_b', 'tanque_c', 'tanque_d', 'tanque_e', 'tanque_f'];
  var LABEL_TANQUE = { tanque_a: 'A', tanque_b: 'B', tanque_c: 'C', tanque_d: 'D', tanque_e: 'E', tanque_f: 'F' };
  var tanquesUsados = TODOS_TANQUES.filter(function (t) {
    return sessoes.some(function (r) { return parseFloat(r[idx[t]]) > 0; });
  });

  var progRow = calcRow + 2;
  painel.getRange(progRow, 2).setValue('🕐 Programação por Setor').setFontWeight('bold').setFontSize(13).setFontColor('#164D31');
  progRow++;
  var progHeaders = ['Ordem', 'Setor', 'Talhão', 'Início', 'Fim', 'Duração (min)'].concat(
    tanquesUsados.map(function (t) { return 'Tanque ' + LABEL_TANQUE[t] + ' (L)'; })
  );
  painel.getRange(progRow, 2, 1, progHeaders.length).setValues([progHeaders])
    .setBackground('#276B48').setFontColor('#FFFFFF').setFontWeight('bold');
  progRow++;
  var progStartRow = progRow;

  var linhasProg = sessoes.map(function (r, i) {
    var base = [i + 1, r[idx.setor] || r[idx.grupo], r[idx.talhao_integrado] || '—', r[idx.horario] || '—', r[idx.horario_fim] || '—', r[idx.tempo_min] || 0];
    var tanques = tanquesUsados.map(function (t) { return parseFloat(r[idx[t]]) || 0; });
    return base.concat(tanques);
  });
  painel.getRange(progRow, 2, linhasProg.length, progHeaders.length).setValues(linhasProg);

  // zebra + cor do talhão
  for (var i = 0; i < linhasProg.length; i++) {
    var rowNum = progStartRow + i;
    if (i % 2 === 1) painel.getRange(rowNum, 2, 1, progHeaders.length).setBackground('#F3EEE0');
    var talhaoDaLinha = sessoes[i][idx.talhao_integrado];
    painel.getRange(rowNum, 4).setFontColor(CORES_GRUPO[talhaoDaLinha] || '#2F5FA8').setFontWeight('bold');
  }
  painel.getRange(progStartRow, 2, linhasProg.length, progHeaders.length).setBorder(true, true, true, true, true, true, '#DCD3BC', SpreadsheetApp.BorderStyle.SOLID);

  // ---- larguras e ajustes finais ----
  painel.setColumnWidth(2, 130);
  painel.setColumnWidth(3, 110);
  for (var c = 4; c <= 3 + progHeaders.length; c++) painel.setColumnWidth(c, 90);
  painel.setFrozenRows(0);

  ss.setActiveSheet(painel);
  SpreadsheetApp.getUi().alert('Painel de Irrigação (Cacau) de ' + dataFmt + ' criado! ' + sessoes.length + ' setor(es).');
}
