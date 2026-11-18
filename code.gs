const SPREADSHEET_ID = '19qboiD0IOYFrnWwDRoorMW1olSaP-iRcVeJ369IThts';

// Web App interno da Cresol que publica as regras do Checklist de Crédito.
// Projeto Apps Script separado — as regras de documentos e garantias NÃO estão
// neste repositório. A chamada envia o token OAuth da execução corrente
// (ScriptApp.getOAuthToken) no header Authorization, portanto esta URL só pode
// apontar para um deployment sob controle da Cresol.
const URL_API_CHECKLIST = 'https://script.google.com/a/macros/cresolsicoper.com.br/s/AKfycbx0qsJ6ZhgZnVzUmtWrYL6XiMcIu7Vq9lY7Eo6wU2vMg_7LAHQfpj9uAq_IDQ_O8YHCvg/exec?action=getRegras';

function doGet() {
  setupInicial();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Simulador de Crédito - Cresol')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// SETUP INICIAL E CRIAÇÃO DE PARÂMETROS
// ==========================================
function setupInicial() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    let sheetConf = ss.getSheetByName('Configuracoes');
    if (!sheetConf) {
      sheetConf = ss.insertSheet('Configuracoes');
      sheetConf.appendRow(["Chave", "Valor"]);
      sheetConf.appendRow(["TAGS_ATIVAS", "Veículos, Imóveis, Consórcio, Investimento, Capital de Giro"]);
      sheetConf.appendRow(["CDI_ANUAL", "10.50"]);
      // Matrizes Iniciais de Regras de Negócio (Em formato JSON para o motor)
      sheetConf.appendRow(["SEGURO_FAIXAS", JSON.stringify([{ max: 65, taxa: 0.89 }, { max: 70, taxa: 1.1518 }, { max: 80, taxa: 2.5724 }])]);
      sheetConf.appendRow(["TAC_FAIXAS_PJ", JSON.stringify([{ max: 50000, taxa: 2.00 }, { max: 100000, taxa: 1.50 }, { max: 500000, taxa: 1.00 }, { max: 999999999, taxa: 0.50 }])]);
      sheetConf.appendRow(["TAC_EXCECOES", JSON.stringify([{ cod: "7669", taxa: 3.00 }, { cod: "7668", taxa: 3.00 }])]); // Ex: PNMPO
      formatarCabecalho(sheetConf, 2);
    }

    let sheetUsuarios = ss.getSheetByName('Usuarios');
    if (!sheetUsuarios) {
      sheetUsuarios = ss.insertSheet('Usuarios');
      const colunasUsuarios = ["Nome", "Email", "Perfil", "Agência", "Status"];
      sheetUsuarios.appendRow(colunasUsuarios);
      sheetUsuarios.appendRow(["Anderson Luiz Traesel Weiss", "anderson.weiss@cresolsicoper.com.br", "ADMIN", "Todas", "Ativo"]);
      formatarCabecalho(sheetUsuarios, colunasUsuarios.length);
    }

    let sheetAssociados = ss.getSheetByName('Associados');
    if (!sheetAssociados) {
      sheetAssociados = ss.insertSheet('Associados');
      const colunasAssociados = ["Nome Coop", "Associado", "Conta", "Score", "Pessoa", "Idade"];
      sheetAssociados.appendRow(colunasAssociados);
      formatarCabecalho(sheetAssociados, colunasAssociados.length);
      const maxCols = sheetAssociados.getMaxColumns();
      if (maxCols > colunasAssociados.length) sheetAssociados.deleteColumns(colunasAssociados.length + 1, maxCols - colunasAssociados.length);
    }

    let sheetModalidades = ss.getSheetByName('modalidades');
    if (!sheetModalidades) {
      sheetModalidades = ss.insertSheet('modalidades');
      const colunasModalidades = ["Perfil", "Código", "Modalidade", "Finalidade", "Parcela", "Carência Máx", "Prazo Máx", "Taxa Baixíssimo", "Taxa Baixo", "Taxa Médio", "Taxa Alto", "Público-Alvo", "Garantia", "Descrição", "Tags", "Amortização"];
      sheetModalidades.appendRow(colunasModalidades);
      formatarCabecalho(sheetModalidades, colunasModalidades.length);
    }

    let sheetLog = ss.getSheetByName('Log Simulacoes');
    if (!sheetLog) {
      sheetLog = ss.insertSheet('Log Simulacoes');
      const colunasLog = ["Data/Hora", "Operador", "Email Operador", "Conta", "Associado", "Modalidade", "Valor Financiado", "Carência (Dias)", "Periodicidade", "Prazo", "Valor TAC", "Valor IOF", "Seguro Prestamista", "Cota Capital", "Outras Custas", "Total de Custas", "Valor Líquido", "Taxa Mínima", "Taxa Efetiva", "1ª Parcela", "Última Parcela", "Agência", "Score", "Tipo Simulação", "Grupo Simulação", "ID Simulação", "Versão", "Parâmetros", "Foto Comprometimento", "Foto Checklist"];
      sheetLog.appendRow(colunasLog);
      formatarCabecalho(sheetLog, colunasLog.length);
    }

    let sheetAgencias = ss.getSheetByName('Agencias');
    if (!sheetAgencias) {
      sheetAgencias = ss.insertSheet('Agencias');
      const colunasAgencias = ["Nome", "Status"];
      sheetAgencias.appendRow(colunasAgencias);
      sheetAgencias.appendRow(["Agência Sede", "Ativo"]);
      formatarCabecalho(sheetAgencias, colunasAgencias.length);
    }

    let sheetNovidades = ss.getSheetByName('Novidades');
    if (!sheetNovidades) {
      sheetNovidades = ss.insertSheet('Novidades');
      const colunasNovidades = ["Data", "Resumo", "Detalhes", "Autor"];
      sheetNovidades.appendRow(colunasNovidades);
      formatarCabecalho(sheetNovidades, colunasNovidades.length);
      sheetNovidades.appendRow([Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"), "Sistema de Novidades Implementado!", "Agora você pode acompanhar as melhorias do simulador diretamente na tela principal.", "Agente Antigravity"]);
    }

    let sheetGrupo = ss.getSheetByName('GrupoEconomico');
    if (!sheetGrupo) {
      sheetGrupo = ss.insertSheet('GrupoEconomico');
      const colunasGrupo = ["nr_cpf_cnpj_agrupador", "nr_cpf_cnpj_integrante"];
      sheetGrupo.appendRow(colunasGrupo);
      formatarCabecalho(sheetGrupo, colunasGrupo.length);
      const maxCols = sheetGrupo.getMaxColumns();
      if (maxCols > colunasGrupo.length) sheetGrupo.deleteColumns(colunasGrupo.length + 1, maxCols - colunasGrupo.length);
    }

  } catch (e) { Logger.log("Erro no setup: " + e.message); }
}

function formatarCabecalho(sheet, numColunas) {
  const range = sheet.getRange(1, 1, 1, numColunas);
  range.setFontWeight("bold");
  range.setBackground("#D9EAD3");
  sheet.setFrozenRows(1);
}

// ==========================================
// SEGURANÇA E CONTROLO DE ACESSO
// ==========================================
// ─── ACESSO PELO HUB — CENTRAL DE FERRAMENTAS (15/10/2026) ────────────────────
// O acesso deixou de vir da aba local `Usuarios` e passou a seguir a política
// cadastrada no Hub Cresol Cooperar (Central de Ferramentas), como nas demais
// aplicações (padrão do projeto "Gestão de Contas com Restrição"):
//   - lemos a planilha do Hub SOMENTE para leitura, com cache de HUB_CACHE_TTL;
//   - o Simulador é um PAINEL do Hub: escopo, classes de cargo, agências e
//     bloqueios cadastrados lá decidem quem entra;
//   - perfil: admin do Hub → ADMIN; colaborador de agência marcada como Sede →
//     SEDE ADMINISTRATIVA; demais colaboradores → AGÊNCIA (decisão de 15/10/2026);
//   - a agência do Hub ("CRESOL PAC HORIZONTINA") é traduzida para o nome usado
//     no Simulador ("Horizontina") por MAPA_AGENCIAS_HUB — é por esse nome que o
//     histórico filtra o perfil AGÊNCIA e que os logs antigos foram gravados;
//   - falha ao consultar o Hub NEGA o acesso. Para não ficar trancado, o e-mail
//     na Propriedade do script EMERGENCY_ADMIN entra sempre como ADMIN.
// PRÉ-REQUISITO: a conta que publica o Web App ("Executar como: eu") precisa de
// acesso de LEITURA à planilha do Hub.
const HUB_SPREADSHEET_ID = '1fSSV2QwswGaC2kjYSBh3g9Y7UNg4NqDADxSMjuZPlhw';
// Painel do Simulador no Hub. Precedência: chave HUB_ID_PAINEL da aba
// Configuracoes > esta constante > busca pelo nome HUB_NOME_PAINEL.
// Para descobrir o ID, rode listarPaineisHub() no editor.
const HUB_ID_PAINEL = '';
const HUB_NOME_PAINEL = 'Simulador de Crédito';
const HUB_ABAS = {
  paineis: 'db_paineis',
  colaboradores: 'db_colaboradores',
  admins: 'db_admins',
  agencias: 'cfg_agencias',
  classes: 'cfg_classes_cargo',
  mapaCargoClasse: 'map_cargo_classe'
};
const HUB_ESCOPO = { TODOS: 'TODOS', TODAS_AGENCIAS: 'TODAS_AGENCIAS', SOMENTE_SEDE: 'SOMENTE_SEDE' };
const HUB_CACHE_TTL = 300; // segundos — mudança feita no Hub vale aqui em até 5 min

// Nome da agência no Hub (cfg_agencias.Nome_Agencia) → nome no Simulador.
// As chaves passam por _normalizarNomeAgencia_ (maiúsculas, sem acento).
// Agência do Hub que não estiver aqui entra com o nome do próprio Hub (e um aviso no log).
const MAPA_AGENCIAS_HUB = {
  'CRESOL COOPERAR': 'Todas',
  'CRESOL PAC HORIZONTINA': 'Horizontina',
  'CRESOL PAC SAO MARTINHO': 'Sao Martinho',
  'CRESOL PAC HUMAITA': 'Humaita',
  'CRESOL PAC VENANCIO AIRES': 'Venancio Aires',
  'CRESOL PAC SANTA CLARA DO SUL': 'Santa Clara do Sul',
  'CRESOL PAC BOA VISTA DO BURICA': 'Boa Vista do Burica',
  'CRESOL PA CONECTA COOPERAR': 'Conecta',
  'CRESOL PAC RIO PARDO': 'Rio Pardo',
  'CRESOL PAC NOVA CANDELARIA': 'Nova Candelaria',
  'CRESOL PAC VALE DO SOL': 'Vale do Sol',
  'CRESOL PAC CRISSIUMAL': 'Crissiumal',
  'CRESOL PAC VERA CRUZ': 'Vera Cruz',
  'CRESOL PAC CORONEL BICACO': 'Coronel Bicaco',
  'CRESOL PAC CANDELARIA': 'Candelaria',
  'CRESOL PAC SANTA CRUZ DO SUL': 'Santa Cruz Do Sul (Centro)',
  'CRESOL PAC SEDE NOVA': 'Sede Nova',
  'CRESOL PAC ESPERANCA DO SUL': 'Esperanca do Sul',
  'CRESOL PAC TIRADENTES DO SUL': 'Tiradentes do Sul',
  'CRESOL PAC TAQUARI': 'Taquari',
  'CRESOL PAC SANTA CRUZ DO SUL ARROIO GRAN': 'Sta Cruz do Sul B. Arroio Grande',
  'CRESOL PAC SAO JERONIMO': 'São Jerônimo',
  // Na tabela enviada em 15/10/2026 estas duas vieram trocadas; confirmado que
  // cada uma corresponde ao próprio nome.
  'CRESOL PAC PASSO DO SOBRADO': 'Passo do Sobrado',
  'CRESOL PAC BOQUEIRAO DO LEAO': 'Boqueirão do Leão'
};

function _normalizarNomeAgencia_(nome) {
  return String(nome || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toUpperCase().replace(/\s+/g, ' ').trim();
}

/** Nome da agência no Simulador a partir do nome no Hub. */
function _agenciaSimuladorDoHub_(nomeHub) {
  const chave = _normalizarNomeAgencia_(nomeHub);
  if (MAPA_AGENCIAS_HUB[chave]) return MAPA_AGENCIAS_HUB[chave];
  if (chave) Logger.log('[HUB] Agência sem correspondência em MAPA_AGENCIAS_HUB: "' + nomeHub + '"');
  return String(nomeHub || '').trim();
}

/**
 * Lê uma aba do Hub como lista de objetos {cabeçalho: valor}, com cache do
 * script. `contexto.ss` guarda a planilha aberta para as leituras seguintes
 * da mesma chamada (abre só se alguma aba não estiver em cache).
 */
function _hubLerAba_(nomeAba, contexto) {
  const cache = CacheService.getScriptCache();
  const chave = 'hub_aba_' + nomeAba;
  const emCache = cache.get(chave);
  if (emCache) { try { return JSON.parse(emCache); } catch (e) { /* cache corrompido: relê */ } }
  if (!contexto.ss) {
    try { contexto.ss = SpreadsheetApp.openById(HUB_SPREADSHEET_ID); }
    catch (e) { throw new Error('não foi possível abrir a planilha do Hub (a conta que publica o app precisa de acesso de leitura). ' + e.message); }
  }
  const sheet = contexto.ss.getSheetByName(nomeAba);
  if (!sheet) throw new Error('aba "' + nomeAba + '" não encontrada na planilha do Hub.');
  const valores = sheet.getDataRange().getDisplayValues(); // texto: nada de Date indo para o cache/cliente
  if (valores.length < 2) return [];
  const cab = valores[0].map(h => String(h).trim());
  const linhas = [];
  for (let i = 1; i < valores.length; i++) {
    const obj = {};
    cab.forEach((c, j) => { obj[c] = valores[i][j]; });
    linhas.push(obj);
  }
  try { cache.put(chave, JSON.stringify(linhas), HUB_CACHE_TTL); } catch (e) { /* grande demais para o cache: segue sem */ }
  return linhas;
}

function _hubLista_(valor) {
  return String(valor || '').split(',').map(v => v.trim()).filter(v => v);
}

function _hubVerdadeiro_(valor) {
  const v = String(valor || '').trim().toUpperCase();
  return v === 'SIM' || v === 'TRUE' || v === 'VERDADEIRO' || v === 'S';
}

function _hubAtivo_(valor) {
  return String(valor || '').trim().toUpperCase() === 'ATIVO';
}

/** Contexto do usuário no Hub: admin?, colaborador ativo?, classe, agência, setores. */
function _hubContextoUsuario_(email, contexto) {
  const emailNorm = String(email || '').trim().toLowerCase();
  const ctx = {
    email: emailNorm, isAdmin: false, colaborador: false, nome: '', idClasse: '', ehSoberano: false,
    idAgencia: '', ehSede: false, agencia: '', idsSetoresAdmin: []
  };
  if (!emailNorm) return ctx;

  ctx.isAdmin = _hubLerAba_(HUB_ABAS.admins, contexto).some(a =>
    String(a.Email || '').trim().toLowerCase() === emailNorm && _hubAtivo_(a.Status));

  const colab = _hubLerAba_(HUB_ABAS.colaboradores, contexto).find(c =>
    String(c.Email || '').trim().toLowerCase() === emailNorm && _hubAtivo_(c.Status));
  if (!colab) return ctx;

  ctx.colaborador = true;
  ctx.nome = String(colab.Nome || '').trim();
  ctx.idAgencia = String(colab.ID_Agencia || '').trim();
  ctx.idsSetoresAdmin = _hubLista_(colab.IDs_Setores_Administrativos);

  const cargo = String(colab.Cargo_Bruto || '').trim();
  const mapa = _hubLerAba_(HUB_ABAS.mapaCargoClasse, contexto).find(m =>
    String(m.Cargo_Bruto || '').trim() === cargo && (String(m.Status || '').trim() === '' || _hubAtivo_(m.Status)));
  if (mapa) ctx.idClasse = String(mapa.ID_Classe || '').trim();
  if (ctx.idClasse) {
    const classe = _hubLerAba_(HUB_ABAS.classes, contexto).find(cl => String(cl.ID_Classe || '').trim() === ctx.idClasse);
    if (classe) ctx.ehSoberano = _hubVerdadeiro_(classe.Acesso_Soberano);
  }
  if (ctx.idAgencia) {
    const ag = _hubLerAba_(HUB_ABAS.agencias, contexto).find(a => String(a.ID_Agencia || '').trim() === ctx.idAgencia);
    if (ag) {
      ctx.ehSede = _hubVerdadeiro_(ag.Eh_Sede);
      ctx.agencia = _agenciaSimuladorDoHub_(ag.Nome_Agencia);
    }
  }
  return ctx;
}

/** Linha do Simulador em db_paineis (por ID configurado; senão pelo nome). */
function _hubPainel_(idConfigurado, contexto) {
  const paineis = _hubLerAba_(HUB_ABAS.paineis, contexto);
  const id = String(idConfigurado || HUB_ID_PAINEL || '').trim();
  if (id) {
    const p = paineis.find(x => String(x.ID_Painel || '').trim() === id);
    if (p) return p;
  }
  return paineis.find(x => String(x.Nome || '').trim() === HUB_NOME_PAINEL) || null;
}

/** Mesma regra de visibilidade do Hub. Inclusão vazia = sem restrição; exclusão sempre vence. */
function _hubPodeAcessarPainel_(ctx, painel) {
  if (!painel) return false;
  if (String(painel.Status || '').trim() !== 'Publicado') return false;
  if (ctx.isAdmin || ctx.ehSoberano) return true;

  const escopo = String(painel.Escopo || HUB_ESCOPO.TODOS).trim();
  if (escopo === HUB_ESCOPO.SOMENTE_SEDE && !ctx.ehSede) return false;
  if (escopo === HUB_ESCOPO.TODAS_AGENCIAS && ctx.ehSede) return false;

  const classes = _hubLista_(painel.Classes_Permitidas);
  if (classes.length && classes.indexOf(ctx.idClasse) === -1) return false;
  const agencias = _hubLista_(painel.Agencias_Especificas);
  if (agencias.length && agencias.indexOf(ctx.idAgencia) === -1) return false;
  const setores = _hubLista_(painel.Setores_Administrativos_Especificos);
  if (setores.length && !ctx.idsSetoresAdmin.some(s => setores.indexOf(s) !== -1)) return false;

  if (_hubLista_(painel.Classes_Bloqueadas).indexOf(ctx.idClasse) !== -1) return false;
  if (_hubLista_(painel.Agencias_Bloqueadas).indexOf(ctx.idAgencia) !== -1) return false;
  const setoresBloq = _hubLista_(painel.Setores_Administrativos_Bloqueados);
  if (ctx.idsSetoresAdmin.some(s => setoresBloq.indexOf(s) !== -1)) return false;
  return true;
}

/** Lê chaves da aba Configuracoes deste app. */
function _lerConfiguracoes_(chaves) {
  const resultado = {};
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Configuracoes');
  if (!sheet) return resultado;
  sheet.getDataRange().getValues().forEach(r => {
    if (chaves.indexOf(r[0]) !== -1) resultado[r[0]] = r[1];
  });
  return resultado;
}

/** E-mails dos admins ativos do Hub (= ADMINs do Simulador). Usado nos avisos do sync. */
function _emailsAdminsHub_() {
  return _hubLerAba_(HUB_ABAS.admins, {})
    .filter(a => _hubAtivo_(a.Status))
    .map(a => String(a.Email || '').trim())
    .filter(e => e);
}

/**
 * Valida o usuário da sessão pelo Hub. Contrato mantido para a tela:
 * { autorizado, nome, email, perfil, agencia } ou { autorizado: false, erro }.
 */
function verificarAcesso() {
  try {
    const emailLogado = Session.getActiveUser().getEmail();
    if (!emailLogado) return { autorizado: false, erro: "E-mail não identificado." };
    const emailNorm = String(emailLogado).trim().toLowerCase();

    // Acesso de emergência: entra como ADMIN mesmo com o Hub fora do ar.
    const emergencia = PropertiesService.getScriptProperties().getProperty('EMERGENCY_ADMIN');
    if (emergencia && emergencia.trim().toLowerCase() === emailNorm) {
      const acessoEmergencia = { autorizado: true, nome: 'Admin (emergência)', email: emailLogado, perfil: 'ADMIN', agencia: 'Todas' };
      _guardarAcessoEmCache_(emailLogado, acessoEmergencia);
      return acessoEmergencia;
    }

    const contexto = {};
    const ctx = _hubContextoUsuario_(emailNorm, contexto);
    let acesso;

    if (ctx.isAdmin) {
      // Admin do Hub ignora as regras do painel e o modo manutenção, como antes.
      acesso = { autorizado: true, nome: ctx.nome || emailLogado, email: emailLogado, perfil: 'ADMIN', agencia: ctx.agencia || 'Todas' };
    } else {
      if (!ctx.colaborador) {
        return { autorizado: false, erro: `Acesso negado. O e-mail (${emailLogado}) não está ativo na Central de Ferramentas.` };
      }
      const config = _lerConfiguracoes_(['MANUTENCAO_ATIVA', 'HUB_ID_PAINEL']);
      const painel = _hubPainel_(config.HUB_ID_PAINEL, contexto);
      if (!painel) {
        return { autorizado: false, erro: "O Simulador de Crédito não foi localizado na Central de Ferramentas. Avise um administrador." };
      }
      if (!_hubPodeAcessarPainel_(ctx, painel)) {
        return { autorizado: false, erro: "Seu cargo ou agência não tem acesso ao Simulador de Crédito na Central de Ferramentas." };
      }
      if (String(config.MANUTENCAO_ATIVA || '').trim().toUpperCase() === "SIM") {
        return { autorizado: false, erro: "SISTEMA EM MANUTENÇÃO. Por favor, aguarde alguns instantes." };
      }
      acesso = {
        autorizado: true, nome: ctx.nome || emailLogado, email: emailLogado,
        perfil: ctx.ehSede ? 'SEDE ADMINISTRATIVA' : 'AGÊNCIA',
        agencia: ctx.agencia
      };
    }

    _guardarAcessoEmCache_(emailLogado, acesso);
    return acesso;
  } catch (e) {
    // Fail-closed: sem resposta confiável do Hub, ninguém entra (exceto EMERGENCY_ADMIN).
    return { autorizado: false, erro: "Erro ao validar o acesso na Central de Ferramentas: " + e.message };
  }
}

/**
 * Painéis do Hub e o painel em uso pelo Simulador. Usada pela tela Admin
 * (Gestão de Usuários) e também pelo editor — registra a lista no log. Exige ADMIN.
 */
function listarPaineisHub() {
  try {
    validarAcessoAdmin();
    const paineis = _hubLerAba_(HUB_ABAS.paineis, {});
    const idConfigurado = String(_lerConfiguracoes_(['HUB_ID_PAINEL']).HUB_ID_PAINEL || '').trim();
    const atual = _hubPainel_(idConfigurado, {});
    const resumir = p => ({
      id: String(p.ID_Painel || '').trim(), nome: String(p.Nome || '').trim(),
      status: String(p.Status || '').trim(), escopo: String(p.Escopo || 'TODOS').trim()
    });
    Logger.log('PAINÉIS DO HUB (ID | Nome | Status | Escopo)\n' +
      paineis.map(p => { const r = resumir(p); return `${r.id} | ${r.nome} | ${r.status} | escopo: ${r.escopo}`; }).join('\n'));
    Logger.log('Painel usado hoje pelo Simulador: ' + (atual ? `${atual.ID_Painel} | ${atual.Nome} | ${atual.Status}` : 'NENHUM'));
    return {
      sucesso: true,
      paineis: paineis.map(resumir).filter(p => p.id),
      idConfigurado: idConfigurado,
      nomePadrao: HUB_NOME_PAINEL,
      painelAtual: atual ? resumir(atual) : null
    };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Define o painel do Simulador no Hub (chave HUB_ID_PAINEL). Id vazio volta à
 * busca automática pelo nome HUB_NOME_PAINEL. Exige ADMIN.
 * Mudar o painel muda QUEM pode entrar no Simulador — por isso invalida o
 * cache de acesso: a nova regra vale na próxima ação de cada usuário.
 */
function salvarPainelHub(idPainel) {
  try {
    validarAcessoAdmin();
    const id = String(idPainel || '').trim();
    let painel = null;
    if (id) {
      painel = _hubLerAba_(HUB_ABAS.paineis, {}).find(p => String(p.ID_Painel || '').trim() === id);
      if (!painel) return { sucesso: false, erro: 'Painel não encontrado na Central de Ferramentas.' };
    }
    const gravacao = _setConfigInterna('HUB_ID_PAINEL', id);
    if (!gravacao || !gravacao.sucesso) return { sucesso: false, erro: (gravacao && gravacao.erro) || 'Falha ao gravar a configuração.' };
    _invalidarCacheAcesso_();
    const resolvido = _hubPainel_(id, {});
    return {
      sucesso: true,
      painelAtual: resolvido ? { id: String(resolvido.ID_Painel || '').trim(), nome: String(resolvido.Nome || '').trim(), status: String(resolvido.Status || '').trim() } : null
    };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Confere MAPA_AGENCIAS_HUB contra as agências reais do Hub — rode pelo editor.
 * Mostra cada agência do Hub e o nome que ela recebe no Simulador ("SEM MAPA" =
 * usa o nome do próprio Hub; normal para agência nova, sem histórico). A mesma
 * informação aparece na tela Admin → Agências. Exige ADMIN.
 */
function conferirMapaAgenciasHub() {
  validarAcessoAdmin();
  const agenciasHub = _hubLerAba_(HUB_ABAS.agencias, {});
  const linhas = agenciasHub.map(a => {
    const chave = _normalizarNomeAgencia_(a.Nome_Agencia);
    const noMapa = !!MAPA_AGENCIAS_HUB[chave];
    const nome = _agenciaSimuladorDoHub_(a.Nome_Agencia);
    return `${noMapa ? 'ok      ' : 'SEM MAPA'} | ${a.Nome_Agencia} → ${nome}${_hubVerdadeiro_(a.Eh_Sede) ? ' (Sede)' : ''}`;
  });
  Logger.log('AGÊNCIAS DO HUB → NOME NO SIMULADOR\n' + linhas.join('\n'));
  return { sucesso: true, total: linhas.length };
}

// ─── CACHE DA VALIDAÇÃO DE ACESSO (desempenho — 14/10/2026) ──────────────────
// Toda chamada do app validava o acesso relendo as abas Usuarios e
// Configuracoes — numa busca de PJ, 3 vezes. Agora o resultado POSITIVO de
// verificarAcesso() fica guardado por ACESSO_CACHE_SEGUNDOS e os guards o
// reaproveitam. O login (verificarAcesso direto) sempre consulta a planilha.
//
// - Cache do SCRIPT, com o e-mail na chave. NÃO trocar por getUserCache(): com
//   a implantação "Executar como: eu", o "usuário" daquele cache é o dono do
//   script, e todos os operadores dividiriam a mesma entrada.
// - Invalidação: a gravação de MANUTENCAO_ATIVA troca ACESSO_VERSAO (Propriedades
//   do script), o que deixa órfãs todas as entradas. Mudanças feitas no Hub
//   (colaborador inativado, cargo, agência, painel) valem em até ~10 min: até
//   5 min do cache das abas do Hub (HUB_CACHE_TTL) + até 5 min deste cache.
// - Os nomes TERMINAM em "_": no Apps Script isso torna a função privada
//   (google.script.run não a alcança). Sem isso, qualquer usuário poderia
//   gravar no cache uma entrada dizendo-se ADMIN.
const ACESSO_CACHE_SEGUNDOS = 300;

function _chaveCacheAcesso_(email) {
  const versao = PropertiesService.getScriptProperties().getProperty('ACESSO_VERSAO') || '0';
  return 'ACESSO_' + versao + '_' + String(email).trim().toLowerCase();
}

function _guardarAcessoEmCache_(email, acesso) {
  try {
    CacheService.getScriptCache().put(_chaveCacheAcesso_(email), JSON.stringify(acesso), ACESSO_CACHE_SEGUNDOS);
  } catch (e) { /* o cache é só otimização: falhar aqui não pode bloquear o acesso */ }
}

function _invalidarCacheAcesso_() {
  try {
    // UUID e não Date.now(): duas invalidações no mesmo milissegundo gerariam a
    // mesma versão, e a segunda não invalidaria nada.
    PropertiesService.getScriptProperties().setProperty('ACESSO_VERSAO', Utilities.getUuid());
  } catch (e) { Logger.log('[AVISO] Não foi possível invalidar o cache de acesso: ' + e.message); }
}

/** Acesso do usuário da sessão: do cache, se houver; senão verificação completa. */
function _acessoAtual_() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (email) {
      const salvo = CacheService.getScriptCache().get(_chaveCacheAcesso_(email));
      if (salvo) return JSON.parse(salvo);
    }
  } catch (e) { /* sem cache: segue para a verificação completa */ }
  return verificarAcesso();
}

function validarAcessoAdmin() {
  const acesso = _acessoAtual_();
  if (!acesso.autorizado || String(acesso.perfil).toUpperCase() !== "ADMIN") {
    throw new Error("Acesso negado. Apenas administradores podem executar esta ação.");
  }
}

/**
 * Valida se há uma sessão de usuário ADMIN ativa.
 * Retorna true se há sessão válida de admin, false se não há sessão (contexto de trigger).
 * Nunca lança exceção — use para decisão condicional, não para barrar execução.
 */
function _isContextoTrigger() {
  try {
    const email = Session.getActiveUser().getEmail();
    return !email; // Sem e-mail = sem sessão de usuário = contexto de trigger
  } catch (e) {
    return true; // Erro ao obter sessão também indica contexto de trigger
  }
}

/**
 * Guarda para funções operacionais — exige usuário cadastrado e ATIVO na aba
 * Usuarios, de qualquer perfil.
 *
 * IMPORTANTE: toda função de topo deste arquivo é invocável por qualquer
 * usuário do domínio via google.script.run, tenha ou não botão na interface.
 * Esconder o menu no frontend NÃO protege nada.
 *
 * Lança exceção quando não autorizado (o withFailureHandler do cliente trata).
 *
 * @returns {Object} { autorizado, nome, email, perfil, agencia } do usuário logado
 */
function validarAcessoOperacional() {
  const acesso = _acessoAtual_();
  if (!acesso.autorizado) {
    throw new Error(acesso.erro || "Acesso negado.");
  }
  return acesso;
}

/**
 * Guarda para funções que rodam TANTO pelo trigger agendado (sem sessão)
 * QUANTO pelo painel Admin (com sessão).
 *
 * Substitui o antigo parâmetro `bypassAuth`, que era uma afirmação do CLIENTE
 * sobre o contexto de execução e portanto burlável — bastava passar `true`.
 * Aqui quem decide é o SERVIDOR, consultando a sessão.
 *
 * Sem sessão  → é o trigger, libera.
 * Com sessão  → tem de ser ADMIN.
 */
function validarContextoInterno() {
  if (_isContextoTrigger()) return;
  validarAcessoAdmin();
}

/*
 * FUNÇÕES INTENCIONALMENTE SEM GUARD — não acrescente efeito colateral nelas
 * sem antes adicionar um guard, porque TODA função de topo é exposta por RPC.
 *
 * Rodam antes de existir sessão:
 *   doGet, verificarAcesso, setupInicial
 *   → setupInicial é chamada por doGet a cada carregamento. Um guard aqui
 *     derruba o login de toda a rede. É idempotente (só cria aba faltante).
 *
 * Helpers puros, sem efeito colateral e sem leitura de dado sensível:
 *   formatarCabecalho, _isContextoTrigger, _calcularIdade, _normalizarPessoa,
 *   _normalizarBoolean, _parseCurrencyRobust, _garantirCabecalho_,
 *   _valorCelulaParaTexto_, _getSheetPesquisas_, _getSheetRespostas_,
 *   _gerarIdPesquisa_, _getDadosUsuarioLogado_, _buscarDadosComprometimento
 *   (os helpers dos limites pré-aprovados terminam em `_` e são privados)
 *   → _buscarDadosComprometimento exige um objeto Spreadsheet como 1º
 *     argumento, que não é serializável pelo cliente; só é alcançável através
 *     de calcularComprometimentoRendaGlobal, que tem guard operacional.
 *
 * Funções cujo nome TERMINA em "_" são privadas no Apps Script: o
 * google.script.run não as alcança, então não precisam de guard. É o caso dos
 * helpers de cache de acesso (_acessoAtual_, _guardarAcessoEmCache_...) e de
 * localização rápida (_localizarPrimeira_, _localizarPorDigitos_...). Atenção:
 * o "_" no INÍCIO do nome (ex.: _setConfigInterna) NÃO torna a função privada.
 */

function alternarModoManutencao(estado) {
  try {
    validarAcessoAdmin();
    setConfigGlobal('MANUTENCAO_ATIVA', estado);
    return { sucesso: true };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

// A gestão de usuários passou para o Hub (Central de Ferramentas) em 15/10/2026.
// listarUsuarios devolve a marca `gerenciadoNoHub` para a tela mostrar o aviso;
// salvarUsuario e alternarStatusUsuario recusam — gravar na aba local Usuarios
// não daria nem tiraria acesso de ninguém.
function listarUsuarios() {
  try {
    validarContextoInterno();
    return { sucesso: true, dados: [], gerenciadoNoHub: true };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

function salvarUsuario(nome, email, perfil, agencia) {
  try {
    validarAcessoAdmin();
    return { sucesso: false, erro: "Os usuários agora são gerenciados na Central de Ferramentas (Hub)." };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

function alternarStatusUsuario(email) {
  try {
    validarAcessoAdmin();
    return { sucesso: false, erro: "Os usuários agora são gerenciados na Central de Ferramentas (Hub)." };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

// Desde 15/10/2026 as agências vêm do Hub (cfg_agencias): agência nova
// cadastrada lá aparece aqui sozinha. `nome` é o nome usado no Simulador
// (MAPA_AGENCIAS_HUB para as que já existiam; senão o próprio nome do Hub) —
// é ele que vai para o log e para o filtro do histórico.
function listarAgencias() {
  try {
    validarAcessoAdmin();
    const agencias = _hubLerAba_(HUB_ABAS.agencias, {})
      .filter(a => String(a.Nome_Agencia || '').trim())
      .map(a => {
        const statusHub = String(a.Status || '').trim();
        return {
          nome: _agenciaSimuladorDoHub_(a.Nome_Agencia),
          nomeHub: String(a.Nome_Agencia).trim(),
          status: (!statusHub || _hubAtivo_(statusHub)) ? 'Ativo' : 'Inativo',
          sede: _hubVerdadeiro_(a.Eh_Sede),
          mapeada: !!MAPA_AGENCIAS_HUB[_normalizarNomeAgencia_(a.Nome_Agencia)]
        };
      })
      .sort((x, y) => x.nome.localeCompare(y.nome, 'pt-BR'));
    return { sucesso: true, dados: agencias, origemHub: true };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

function salvarAgencia(nome, status = "Ativo") {
  try {
    validarAcessoAdmin();
    return { sucesso: false, erro: "As agências agora são cadastradas na Central de Ferramentas (Hub)." };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

// ============================================================================
// GESTÃO DINÂMICA DE CONFIGURAÇÕES E MATRIZES
// ==========================================
function setConfigGlobal(chave, valor) {
  try {
    validarAcessoAdmin();
    return _setConfigInterna(chave, valor);
  } catch (e) {
    Logger.log(`[AVISO] Falha ao gravar configuração global (${chave}): ${e.message}`);
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Grava uma chave em 'Configuracoes' a partir de trigger ou de ADMIN.
 *
 * Usada por triggers automáticos e funções internas de sistema que não possuem
 * sessão de usuário ativa (ex: sincronizarGrupoEconomico, substituirAssociados).
 *
 * ATENÇÃO: o prefixo `_` NÃO torna a função privada no Apps Script — ela é de
 * topo e portanto invocável por google.script.run. Até 23/09/2026 não tinha
 * guard nenhum e anulava por completo o de setConfigGlobal, permitindo gravar
 * MANUTENCAO_ATIVA, CDI_ANUAL, TAC_FAIXAS_PJ e SEGURO_FAIXAS sem ser admin.
 */
function _setConfigInterna(chave, valor) {
  try {
    validarContextoInterno();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheetConf = ss.getSheetByName('Configuracoes');
    if (!sheetConf) { sheetConf = ss.insertSheet('Configuracoes'); sheetConf.appendRow(["Chave", "Valor"]); }

    // Manutenção ligada/desligada precisa valer já, não depois do TTL do cache de
    // acesso. A invalidação vem DEPOIS da gravação: antes dela, uma chamada no
    // intervalo guardaria o estado antigo por mais 5 minutos.
    const invalidarSeManutencao = () => { if (chave === 'MANUTENCAO_ATIVA') _invalidarCacheAcesso_(); };
    let data = sheetConf.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === chave) {
        sheetConf.getRange(i + 1, 2).setValue(valor);
        SpreadsheetApp.flush();
        invalidarSeManutencao();
        return { sucesso: true };
      }
    }
    sheetConf.appendRow([chave, valor]);
    SpreadsheetApp.flush();
    invalidarSeManutencao();
    return { sucesso: true };
  } catch (e) {
    Logger.log(`[AVISO] Falha ao gravar configuração interna (${chave}): ${e.message}`);
    return { sucesso: false, erro: e.message };
  }
}

function buscarConfiguracoesAdmin() {
  try {
    validarAcessoOperacional();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetMod = ss.getSheetByName('modalidades');
    let tagsUnicas = [];

    if (sheetMod && sheetMod.getLastRow() > 1) {
      const data = sheetMod.getDataRange().getValues();
      const headers = data[0];
      const idxTag = headers.indexOf("Tags");
      if (idxTag !== -1) {
        let todasTags = new Set();
        for (let i = 1; i < data.length; i++) {
          let tagsCell = String(data[i][idxTag]).split(',');
          tagsCell.forEach(t => { let tagLimpa = t.trim(); if (tagLimpa) todasTags.add(tagLimpa); });
        }
        tagsUnicas = Array.from(todasTags).sort();
      }
    }

    const sheetConf = ss.getSheetByName('Configuracoes');
    let ativas = []; let cdi = "10.50";
    let seguroFaixas = []; let tacPJ = []; let tacExc = [];
    let dataAssoc = "Não informada"; let dataMod = "Não informada"; let versaoMod = "N/A";
    let dataGrupo = "Não informada";
    let manutencao = "NÃO";
    let instAssoc, instMod, instEnriq, orientCDI, orientSeguro, orientTAC;

    if (sheetConf) {
      const dataConf = sheetConf.getDataRange().getValues();
      for (let i = 0; i < dataConf.length; i++) {
        if (dataConf[i][0] === "TAGS_ATIVAS") ativas = String(dataConf[i][1]).split(',').map(t => t.trim()).filter(t => t);
        if (dataConf[i][0] === "CDI_ANUAL") cdi = String(dataConf[i][1]);
        if (dataConf[i][0] === "SEGURO_FAIXAS") seguroFaixas = JSON.parse(dataConf[i][1] || "[]");
        if (dataConf[i][0] === "TAC_FAIXAS_PJ") tacPJ = JSON.parse(dataConf[i][1] || "[]");
        if (dataConf[i][0] === "TAC_EXCECOES") tacExc = JSON.parse(dataConf[i][1] || "[]");
        if (dataConf[i][0] === "DATA_ATUALIZACAO_ASSOCIADOS") dataAssoc = String(dataConf[i][1]);
        if (dataConf[i][0] === "DATA_ATUALIZACAO_MODALIDADES") dataMod = String(dataConf[i][1]);
        if (dataConf[i][0] === "DATA_ATUALIZACAO_GRUPO_ECONOMICO") dataGrupo = String(dataConf[i][1]);
        if (dataConf[i][0] === "VERSAO_MODALIDADE") versaoMod = String(dataConf[i][1]);
        if (dataConf[i][0] === "MANUTENCAO_ATIVA") manutencao = String(dataConf[i][1]).trim().toUpperCase();

        // Novos campos de textos de orientação
        if (dataConf[i][0] === "INST_ASSOC") instAssoc = String(dataConf[i][1]);
        if (dataConf[i][0] === "INST_MOD") instMod = String(dataConf[i][1]);
        if (dataConf[i][0] === "INST_ENRIQ") instEnriq = String(dataConf[i][1]);
        if (dataConf[i][0] === "ORIENT_CDI") orientCDI = String(dataConf[i][1]);
        if (dataConf[i][0] === "ORIENT_SEGURO") orientSeguro = String(dataConf[i][1]);
        if (dataConf[i][0] === "ORIENT_TAC") orientTAC = String(dataConf[i][1]);
      }
    }

    // Textos padrão se estiverem vazios
    const t = (val, def) => (val && val !== "undefined" && val !== "") ? val : def;
    const finalInstAssoc = t(instAssoc || "", "1. Acesse o sistema CRM/Relatórios.\n2. Gere o relatório de associados em formato .xlsx.\n3. Certifique-se de que as colunas Conta e Nome estão presentes.");
    const finalInstMod = t(instMod || "", "1. Baixe o catálogo de taxas e modalidades vigente.\n2. Verifique se as taxas estão em formato decimal.\n3. Importe o arquivo para atualizar o simulador.");
    const finalInstEnriq = t(instEnriq || "", "1. Use este campo para colunas extras (ex: Periodicidade).\n2. O arquivo deve conter a coluna 'Código' como chave.\n3. Os dados novos serão mesclados aos existentes.");
    const finalOrientCDI = t(orientCDI || "", "Atualize o CDI sempre que houver alteração na meta Selic pelo COPOM.");
    const finalOrientSeguro = t(orientSeguro || "", "Defina as faixas de idade de acordo com a apólice vigente da cooperativa.");
    const finalOrientTAC = t(orientTAC || "", "Configure o escalonamento da TAC conforme as regras de receita de serviços vigentes.");

    return {
      sucesso: true, disponiveis: tagsUnicas, ativas: ativas, cdi: cdi,
      seguroFaixas: seguroFaixas, tacPJ: tacPJ, tacExc: tacExc,
      dataAssoc: dataAssoc, dataMod: dataMod, dataGrupo: dataGrupo,
      versaoMod: versaoMod, manutencao: manutencao,
      orientacoes: {
        instAssoc: finalInstAssoc, instMod: finalInstMod, instEnriq: finalInstEnriq,
        orientCDI: finalOrientCDI, orientSeguro: finalOrientSeguro, orientTAC: finalOrientTAC
      }
    };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

function salvarTextoOrientacao(chave, texto) {
  try {
    validarAcessoAdmin();
    setConfigGlobal(chave, texto);
    return { sucesso: true };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

function salvarParametrosFinanceirosAdmin(cdi, seguroFaixas, tacPJ, tacExc) {
  try {
    validarAcessoAdmin();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheetConf = ss.getSheetByName('Configuracoes');
    if (!sheetConf) { sheetConf = ss.insertSheet('Configuracoes'); sheetConf.appendRow(["Chave", "Valor"]); }

    let data = sheetConf.getDataRange().getValues();

    function atualizarOuInserir(chave, valor) {
      let achou = false;
      for (let i = 0; i < data.length; i++) {
        if (data[i][0] === chave) { sheetConf.getRange(i + 1, 2).setValue(valor); achou = true; break; }
      }
      if (!achou) { sheetConf.appendRow([chave, valor]); data.push([chave, valor]); }
    }

    atualizarOuInserir("CDI_ANUAL", cdi);
    atualizarOuInserir("SEGURO_FAIXAS", JSON.stringify(seguroFaixas));
    atualizarOuInserir("TAC_FAIXAS_PJ", JSON.stringify(tacPJ));
    atualizarOuInserir("TAC_EXCECOES", JSON.stringify(tacExc));

    return { sucesso: true };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

function salvarTagsAdmin(tagsArray) {
  try {
    validarAcessoAdmin();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheetConf = ss.getSheetByName('Configuracoes');
    let data = sheetConf.getDataRange().getValues();
    let achou = false; let stringTags = tagsArray.join(', ');
    for (let i = 0; i < data.length; i++) { if (data[i][0] === "TAGS_ATIVAS") { sheetConf.getRange(i + 1, 2).setValue(stringTags); achou = true; break; } }
    if (!achou) sheetConf.appendRow(["TAGS_ATIVAS", stringTags]);
    return { sucesso: true };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

// ==========================================
// FUNÇÕES DE PESQUISA E SIMULAÇÃO
// ==========================================
// ==========================================
// ==========================================
// LOCALIZAÇÃO RÁPIDA NAS ABAS GRANDES (desempenho — 14/10/2026)
// ==========================================
// Medido em produção (versão 184): busca por conta 15–20 s; seleção na busca
// por nome 11–85 s. A causa era ler abas INTEIRAS — Associados chegava a ser
// lida 2 vezes na mesma busca — para achar uma ou poucas linhas. Estes helpers
// localizam só as linhas necessárias:
//   1. a busca nativa da planilha (TextFinder) pré-seleciona candidatas, sem
//      transferir a aba;
//   2. cada candidata é CONFERIDA pela mesma regra de comparação de antes,
//      então o resultado não muda;
//   3. se nenhuma conferir (ex.: valor exibido num formato que a busca não
//      reconhece), cai na leitura de UMA coluna — ainda bem menor que a aba.
// Os nomes terminam em "_": no Apps Script isso torna a função privada.

/** Só os dígitos (CPF/CNPJ). Mesma normalização de antes: String(v).replace(/\D/g, ''). */
function _apenasDigitos_(valor) {
  return String(valor).replace(/\D/g, '');
}

/** Regex que casa a sequência de dígitos com qualquer pontuação entre eles. */
function _regexDigitos_(digitos) {
  return '^\\D*' + String(digitos).split('').join('\\D*') + '\\D*$';
}

/** Regex que casa o texto exato, admitindo espaços nas pontas (como o .trim() de antes). */
function _regexTextoExato_(texto) {
  return '^\\s*' + String(texto).replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s*$';
}

/** Linha 1 da aba (cabeçalho). */
function _cabecalho_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/** Lê uma linha inteira. Uma chamada por linha — use só com poucas linhas. */
function _lerLinha_(sheet, linha, numCols) {
  return sheet.getRange(linha, 1, 1, numCols || sheet.getLastColumn()).getValues()[0];
}

/**
 * Linhas candidatas (crescentes) pela busca nativa.
 * - `max` finito: findNext, parando ao passar de `max` (devolve null). Para TEXTO
 *   que pode se repetir muito — o findAll monta a lista inteira antes de
 *   devolver, e "Sem conta" (repetida em ~30 mil linhas) estourou a memória no
 *   diagnóstico de 14/10/2026. Custo: ~65 ms por ocorrência (uma ida ao servidor
 *   cada), aceitável porque uma conta tem 1 ocorrência.
 * - `max` = Infinity: findAll, UMA chamada só. Para CPF/CNPJ (regex de dígitos),
 *   que tem poucas ocorrências — e quando tem muitas (agrupador de grupo grande),
 *   avançar uma a uma custaria segundos.
 */
function _candidatas_(sheet, col, regex, max) {
  const ultima = sheet.getLastRow();
  if (ultima < 2) return [];
  const finder = sheet.getRange(2, col, ultima - 1, 1).createTextFinder(regex).useRegularExpression(true);
  if (max === Infinity) return finder.findAll().map(r => r.getRow()).sort((a, b) => a - b);
  const linhas = [];
  let achado;
  while ((achado = finder.findNext())) {
    linhas.push(achado.getRow());
    if (linhas.length > max) return null;
  }
  return linhas.sort((a, b) => a - b);
}

/**
 * Lê várias linhas inteiras. Poucas: uma a uma. Muitas: o trecho entre a
 * primeira e a última, numa leitura só (antes lia a aba INTEIRA — com o Grupo
 * Econômico de 463 mil linhas isso levou 22 s e estourou a memória).
 */
function _lerLinhas_(sheet, linhas, numCols) {
  if (linhas.length <= 30) return linhas.map(linha => ({ linha, row: _lerLinha_(sheet, linha, numCols) }));
  const primeira = linhas[0], ultima = linhas[linhas.length - 1];
  const bloco = sheet.getRange(primeira, 1, ultima - primeira + 1, numCols).getValues();
  return linhas.map(linha => ({ linha, row: bloco[linha - primeira] }));
}

/** Plano B: lê só a coluna `col` e devolve as linhas que conferem (até `limite`). */
function _varrerColuna_(sheet, col, confere, limite) {
  const ultima = sheet.getLastRow();
  if (ultima < 2) return [];
  const valores = sheet.getRange(2, col, ultima - 1, 1).getValues();
  const linhas = [];
  for (let i = 0; i < valores.length && linhas.length < limite; i++) {
    if (confere(valores[i][0])) linhas.push(i + 2);
  }
  return _lerLinhas_(sheet, linhas, sheet.getLastColumn());
}

/**
 * Primeira linha (na ordem da planilha) cuja coluna `col` satisfaz `confere`; null se não houver.
 * Usada na busca por CONTA: aqui o plano B continua valendo quando a busca nativa
 * não confere — "conta não encontrada" é raro (erro de digitação) e barato.
 */
function _localizarPrimeira_(sheet, col, regex, confere) {
  const cand = _candidatas_(sheet, col, regex, 20);
  if (cand) {
    const numCols = sheet.getLastColumn();
    for (const linha of cand) {
      const row = _lerLinha_(sheet, linha, numCols);
      if (confere(row[col - 1])) return { linha, row };
    }
  }
  return _varrerColuna_(sheet, col, confere, 1)[0] || null;
}

/**
 * Todas as linhas (na ordem da planilha) cuja coluna `col` satisfaz `confere`.
 * Usada com CPF/CNPJ (regex de dígitos). "Não achei" da busca nativa é aceito
 * como resposta: o diagnóstico de 14/10/2026 mostrou que os documentos são
 * exibidos por extenso e a busca os encontra; reler a coluna a cada negativa
 * custava 2,6 s no Grupo Econômico — e a negativa é o caso comum (associado
 * sem grupo, sócio que não é associado). Use só com regex de dígitos (ver
 * _candidatas_).
 */
function _localizarTodas_(sheet, col, regex, confere) {
  const cand = _candidatas_(sheet, col, regex, Infinity);
  const conferidas = [];
  const numCols = sheet.getLastColumn();
  _lerLinhas_(sheet, cand, numCols).forEach(a => { if (confere(a.row[col - 1])) conferidas.push(a); });
  return conferidas;
}

/**
 * Para cada sequência de dígitos da lista, a PRIMEIRA linha cuja coluna `col`
 * tem os mesmos dígitos. Devolve Map(dígitos → linha inteira).
 * Até 5 itens: busca nativa um a um, uma chamada cada (negativa aceita, como em _localizarTodas_).
 * Acima disso, uma única leitura da coluna atende todos de uma vez.
 */
function _localizarPorDigitos_(sheet, col, listaDigitos) {
  const resultado = new Map();
  const numCols = sheet.getLastColumn();
  const unicos = [...new Set(listaDigitos.filter(d => d))];
  const pendentes = [];
  if (unicos.length <= 5) {
    unicos.forEach(d => {
      const cand = _candidatas_(sheet, col, _regexDigitos_(d), Infinity);
      for (const linha of cand) {
        const row = _lerLinha_(sheet, linha, numCols);
        if (_apenasDigitos_(row[col - 1]) === d) { resultado.set(d, row); return; }
      }
    });
  } else {
    pendentes.push(...unicos);
  }
  if (pendentes.length && sheet.getLastRow() >= 2) {
    const falta = new Set(pendentes);
    const valores = sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getValues();
    const achadas = []; // [dígitos, linha], na ordem da planilha
    for (let i = 0; i < valores.length && falta.size; i++) {
      const d = _apenasDigitos_(valores[i][0]);
      if (falta.has(d)) { achadas.push([d, i + 2]); falta.delete(d); }
    }
    const linhas = _lerLinhas_(sheet, achadas.map(a => a[1]), numCols);
    achadas.forEach((a, k) => resultado.set(a[0], linhas[k].row));
  }
  return resultado;
}

/**
 * Só responde SE o documento integra algum grupo econômico — é tudo o que a
 * tela usa da busca de associado (index.html: `data.grupoEconomico.temGrupo`).
 * A lista de integrantes só é necessária no comprometimento de renda, que a
 * busca de novo pelo servidor (buscarGrupoEconomico). Montá-la a cada busca de
 * conta levou 22 s e estourou a memória com grupos grandes (diagnóstico de
 * 14/10/2026, aba GrupoEconomico com 463 mil linhas).
 */
function _pertenceAGrupo_(cpfCnpj, ss) {
  const semGrupo = { temGrupo: false, integrantes: [] };
  try {
    const sheetGrupo = ss.getSheetByName('GrupoEconomico');
    if (!sheetGrupo || sheetGrupo.getLastRow() <= 1) return semGrupo;
    const doc = _apenasDigitos_(cpfCnpj);
    if (!doc) return semGrupo;
    // Basta UMA linha válida (agrupador e integrante preenchidos, mesma regra de
    // buscarGrupoEconomico). findNext para na primeira ocorrência — sem montar a
    // lista, que para documentos presentes em muitas linhas custou 21 s e
    // estourou a memória. O agrupador não é devolvido: a tela não o usa, e o
    // "primeiro" daqui poderia diferir do "último" de buscarGrupoEconomico.
    const finder = sheetGrupo.getRange(2, 2, sheetGrupo.getLastRow() - 1, 1)
      .createTextFinder(_regexDigitos_(doc)).useRegularExpression(true);
    const numCols = sheetGrupo.getLastColumn();
    let achado, tentativas = 0;
    while ((achado = finder.findNext()) && tentativas++ < 20) {
      const row = _lerLinha_(sheetGrupo, achado.getRow(), numCols);
      if (_apenasDigitos_(row[1]) === doc && _apenasDigitos_(row[0])) return { temGrupo: true, integrantes: [] };
    }
    return semGrupo;
  } catch (e) {
    Logger.log('Erro em _pertenceAGrupo_: ' + e.message);
    return semGrupo;
  }
}

// BUSCAR DADOS DO ASSOCIADO (ULTRARRÁPIDO, INTELIGENTE E COM NORMALIZAÇÃO)
/**
 * Busca a composição societária de um CNPJ, cruzando com a aba Associados para obter a idade.
 * @param {string} cnpj CNPJ a ser buscado.
 * @returns {Array} Lista de sócios contendo cpf, percentual e idade.
 */
function buscarSociosPorCnpj(cnpj) {
  try {
    validarAcessoOperacional();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetComp = ss.getSheetByName('ComposicaoSocietaria');
    if (!sheetComp) return [];

    if (sheetComp.getLastRow() <= 1) return [];

    // Desempenho (14/10/2026): antes lia a Composição INTEIRA e depois a base de
    // Associados INTEIRA de novo, só para a idade dos sócios. Agora localiza
    // apenas as linhas necessárias — mesma regra de comparação (só dígitos).
    const headersComp = _cabecalho_(sheetComp);
    const idxCnpj = headersComp.findIndex(h => String(h).toUpperCase().includes('CNPJ'));
    const idxCpf = headersComp.findIndex(h => String(h).toUpperCase().includes('CPF'));
    const idxPerc = headersComp.findIndex(h => String(h).toUpperCase().includes('PERCENTUAL'));

    if (idxCnpj === -1 || idxCpf === -1 || idxPerc === -1) return [];

    const cleanCnpj = _apenasDigitos_(cnpj);
    let sociosEncontrados = _localizarTodas_(sheetComp, idxCnpj + 1, _regexDigitos_(cleanCnpj),
      v => _apenasDigitos_(v) === cleanCnpj).map(({ row }) => ({
        cpf: String(row[idxCpf]).trim(),
        percentual: parseFloat(String(row[idxPerc]).replace(',', '.')) || 0,
        idade: 0
      }));

    if (sociosEncontrados.length === 0) return [];

    const sheetAssoc = ss.getSheetByName('Associados');
    if (sheetAssoc && sheetAssoc.getLastRow() > 1) {
      const headersAssoc = _cabecalho_(sheetAssoc);
      const aIdxCpf = headersAssoc.findIndex(h => String(h).toUpperCase().includes('CPF') || String(h).toUpperCase().includes('CNPJ'));
      const aIdxIdade = headersAssoc.findIndex(h => String(h).toUpperCase().includes('IDADE'));

      if (aIdxCpf !== -1 && aIdxIdade !== -1) {
        // Primeira ocorrência de cada CPF, como no laço com `break` de antes.
        const porCpf = _localizarPorDigitos_(sheetAssoc, aIdxCpf + 1, sociosEncontrados.map(s => _apenasDigitos_(s.cpf)));
        sociosEncontrados.forEach(socio => {
          const row = porCpf.get(_apenasDigitos_(socio.cpf));
          if (row) socio.idade = parseInt(row[aIdxIdade]) || 0;
        });
      }
    }

    return sociosEncontrados;
  } catch (e) {
    Logger.log("Erro em buscarSociosPorCnpj: " + e.message);
    return [];
  }
}

// ==========================================
// ==========================================
function buscarAssociado(contaBusca) {
  try {
    validarAcessoOperacional();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Associados');
    if (!sheet) return { sucesso: false, erro: "Aba de Associados não encontrada." };

    // Desempenho (14/10/2026): antes lia a base INTEIRA (15–20 s medidos em
    // produção). Agora lê só o cabeçalho e a linha da conta (_localizarPrimeira_).
    if (sheet.getLastRow() <= 1) return { sucesso: false, erro: "Base de associados está vazia." };

    const headers = _cabecalho_(sheet);

    // 1. Busca da Conta
    const idxConta = headers.findIndex(h => String(h).toUpperCase().trim() === 'CONTA' || String(h).toUpperCase().includes('CONTA'));

    // 2. CORREÇÃO: Busca do Nome Inteligente e Segura
    // Primeiro tenta achar os nomes exatos das colunas padrão
    let idxNome = headers.findIndex(h => {
      let str = String(h).toUpperCase().trim();
      return str === 'ASSOCIADO' || str === 'NOME DO ASSOCIADO' || str === 'NOME COMPLETO' || str === 'NOME';
    });

    // Se não achar o nome exato, faz a busca parcial, mas EXCLUI as colunas de "Nome Coop" e "Nome Agência"
    if (idxNome === -1) {
      idxNome = headers.findIndex(h => {
        let str = String(h).toUpperCase();
        return (str.includes('NOME') || str.includes('ASSOCIADO')) &&
          !str.includes('COOP') && !str.includes('AGÊNCIA') && !str.includes('AGENCIA');
      });
    }

    // Busca das outras variáveis
    const idxIdade = headers.findIndex(h => String(h).toUpperCase().includes('IDADE') || String(h).toUpperCase().includes('NASCIMENTO'));
    const idxPessoa = headers.findIndex(h => String(h).toUpperCase().includes('TIPO') && String(h).toUpperCase().includes('PESSOA'));
    const idxMei = headers.findIndex(h => String(h).toUpperCase().includes('MEI'));
    const idxSimples = headers.findIndex(h => String(h).toUpperCase().includes('SIMPLES'));
    const idxScore = headers.findIndex(h => String(h).toUpperCase().includes('SCORE'));
    const idxCpfCnpj = headers.findIndex(h => String(h).toUpperCase().includes('CPF'));
    const idxRenda = headers.findIndex(h => String(h).toUpperCase().includes('RENDA') && !String(h).toUpperCase().includes('DATA'));
    const idxDataRenda = headers.findIndex(h => String(h).toUpperCase().includes('DATA') && String(h).toUpperCase().includes('RENDA') || String(h).toUpperCase().includes('ATUALIZACAO'));
    const idxColab = headers.findIndex(h => String(h).toUpperCase().includes('COLABORADOR'));
    const idxEndivCP = headers.findIndex(h => String(h).toUpperCase().includes('CURTO'));
    const idxEndivTot = headers.findIndex(h => String(h).toUpperCase().includes('ENDIV') && String(h).toUpperCase().includes('TOTAL'));

    if (idxConta === -1) return { sucesso: false, erro: "Coluna 'Conta' não encontrada na base de dados." };

    // Higienizador de dados (Trata datas e nulos)
    const limparDado = (valor) => {
      if (valor instanceof Date) {
        let d = valor.getDate(); let m = valor.getMonth() + 1; let y = valor.getFullYear();
        return `${d < 10 ? '0' + d : d}/${m < 10 ? '0' + m : m}/${y}`;
      }
      return valor == null ? "" : String(valor).trim();
    };

    let contaProcurada = String(contaBusca).trim();

    // Primeira linha cuja conta confere — mesma comparação de antes (texto, com trim).
    const achado = _localizarPrimeira_(sheet, idxConta + 1, _regexTextoExato_(contaProcurada),
      v => String(v).trim() === contaProcurada);
    if (achado) {
      const linhaDados = achado.row;
      let result = {
        sucesso: true,
        nome: idxNome !== -1 ? limparDado(linhaDados[idxNome]) : "Nome Indisponível",
        idade: idxIdade !== -1 ? limparDado(linhaDados[idxIdade]) : "",
        // Normalização semântica: garante 'PF'/'PJ' e 'SIM'/'NÃO'
        pessoa: idxPessoa !== -1 ? _normalizarPessoa(linhaDados[idxPessoa]) : "PF",
        mei: idxMei !== -1 ? _normalizarBoolean(linhaDados[idxMei]) : "NÃO",
        simples: idxSimples !== -1 ? _normalizarBoolean(linhaDados[idxSimples]) : "NÃO",
        score: idxScore !== -1 ? limparDado(linhaDados[idxScore]) : "",
        cpfCnpj: idxCpfCnpj !== -1 ? limparDado(linhaDados[idxCpfCnpj]) : "",
        renda: idxRenda !== -1 ? limparDado(linhaDados[idxRenda]) : "",
        dataRenda: idxDataRenda !== -1 ? limparDado(linhaDados[idxDataRenda]) : "",
        colaborador: idxColab !== -1 ? _normalizarBoolean(linhaDados[idxColab]) : "NÃO",
        endivCP: idxEndivCP !== -1 ? limparDado(linhaDados[idxEndivCP]) : "",
        endivTotal: idxEndivTot !== -1 ? limparDado(linhaDados[idxEndivTot]) : ""
      };
      const civisProp = _dadosCivis_(headers, linhaDados); // checklist: estado civil do proponente
      result.estadoCivil = civisProp.estadoCivil;
      result.uniaoEstavel = civisProp.uniaoEstavel;

      if (result.pessoa === 'PJ' && result.cpfCnpj) {
        result.sociosPJ = buscarSociosPorCnpj(result.cpfCnpj);
      }

      if (result.cpfCnpj) {
        // T2: passa ss já aberto para buscarGrupoEconomico
        result.grupoEconomico = _pertenceAGrupo_(result.cpfCnpj, ss); // a tela só usa temGrupo
      }

      result.limites = _limitesDaConta_(contaProcurada, ss);

      return result;
    }

    return { sucesso: false, erro: "Conta não encontrada na base." };
  } catch (e) {
    return { sucesso: false, erro: "Erro no servidor: " + e.message };
  }
}

// ==========================================
// PESSOA PELO CPF/CNPJ (avalista do checklist — 26/10/2026)
// ==========================================
/**
 * Dados de uma pessoa da base de Associados pelo CPF/CNPJ, para completar o
 * avalista no checklist: nome e, desde 27/10/2026, estado civil, união estável e
 * CPF do cônjuge (documento por estado civil e cônjuge como avalista). Nada de
 * conta, score ou renda; o tipo PF/PJ o cliente tira do tamanho do documento.
 * Sem registro de consulta e sem restrição por agência (decisões de 26/10/2026).
 * Retorno: { sucesso, encontrado, nome, estadoCivil, uniaoEstavel, docConjuge }.
 */
function buscarPessoaPorDocumento(documento) {
  try {
    validarAcessoOperacional();
    const digitos = _apenasDigitos_(documento);
    if (digitos.length !== 11 && digitos.length !== 14) {
      return { sucesso: false, erro: 'Informe um CPF (11 dígitos) ou um CNPJ (14 dígitos).' };
    }
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Associados');
    if (!sheet || sheet.getLastRow() < 2) return { sucesso: true, encontrado: false };

    const headers = _cabecalho_(sheet);
    const idxCpf = headers.findIndex(h => String(h).toUpperCase().includes('CPF'));
    // Mesma detecção da coluna de nome de buscarAssociado (exclui "Nome Coop" e "Nome Agência").
    let idxNome = headers.findIndex(h => {
      const t = String(h).toUpperCase().trim();
      return t === 'ASSOCIADO' || t === 'NOME DO ASSOCIADO' || t === 'NOME COMPLETO' || t === 'NOME';
    });
    if (idxNome === -1) {
      idxNome = headers.findIndex(h => {
        const t = String(h).toUpperCase();
        return (t.includes('NOME') || t.includes('ASSOCIADO')) && !t.includes('COOP') && !t.includes('AGÊNCIA') && !t.includes('AGENCIA');
      });
    }
    if (idxCpf === -1 || idxNome === -1) return { sucesso: false, erro: 'A base de associados não tem as colunas de CPF/CNPJ e nome.' };

    // Documento guardado como número perde o zero à esquerda: tenta também sem ele.
    const semZeros = digitos.replace(/^0+/, '');
    const mapa = _localizarPorDigitos_(sheet, idxCpf + 1, semZeros !== digitos ? [digitos, semZeros] : [digitos]);
    const row = mapa.get(digitos) || mapa.get(semZeros);
    if (!row) return { sucesso: true, encontrado: false };
    const nome = String(row[idxNome] == null ? '' : row[idxNome]).trim();
    if (!nome) return { sucesso: true, encontrado: false };
    const civis = _dadosCivis_(headers, row);
    return { sucesso: true, encontrado: true, nome, estadoCivil: civis.estadoCivil, uniaoEstavel: civis.uniaoEstavel, docConjuge: civis.docConjuge };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

// ==========================================
// MÓDULOS DE IMPORTAÇÃO E ENRIQUECIMENTO
// ==========================================
// ==========================================
// IMPORTAÇÃO DE ASSOCIADOS (RECEBE CARGA LEVE DO FRONTEND)
// ==========================================
/**
 * Substitui a aba 'Associados' completa usando estratégia de swap atômico
 * para evitar perda de dados em caso de timeout.
 *
 * O controle de acesso é decidido pelo SERVIDOR via validarContextoInterno():
 * trigger (sem sessão) passa; chamada do frontend exige ADMIN.
 *
 * NÃO reintroduza um parâmetro de bypass aqui. Até 23/09/2026 existia um
 * `bypassAuth` que o cliente podia passar como true para desligar o guard e
 * substituir a base inteira de Associados.
 *
 * @param {Array[][]} dadosFiltrados - Array 2D com cabeçalho + dados.
 */
function substituirAssociados(dadosFiltrados) {
  try {
    validarContextoInterno();
    if (!dadosFiltrados || dadosFiltrados.length === 0) return { sucesso: false, erro: "Sem dados para gravar." };

    const TAMANHO_LOTE = 5000; // Máximo de linhas por lote para não exceder o tempo do Apps Script
    const totalLinhas = dadosFiltrados.length;
    const neededCols = dadosFiltrados[0].length;
    const NOME_TEMP = '_Associados_temp';

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ── ESTRATÉGIA ANTI-PERDA DE DADOS ─────────────────────────────────────
    // 1. Gravar PRIMEIRO em uma aba temporária
    // 2. Validar que a gravação foi completa
    // 3. SÓ ENTÃO deletar a aba real e renomear a temporária
    // Isso garante que, se um timeout ocorrer durante a gravação,
    // a aba 'Associados' original permanece intacta com os dados anteriores.
    // ────────────────────────────────────────────────────────────────────────

    // Limpar aba temporária de execuções anteriores que possam ter falhado
    const sheetTemporaria = ss.getSheetByName(NOME_TEMP);
    if (sheetTemporaria) {
      ss.deleteSheet(sheetTemporaria);
    }

    // Determinar a posição da aba real para inserir a temporária ao lado
    let sheetReal = ss.getSheetByName('Associados');
    const posReal = sheetReal ? sheetReal.getIndex() - 1 : 0;

    // Criar a aba temporária em posição adjacente
    let sheetTemp = ss.insertSheet(NOME_TEMP, posReal + 1);

    // Redimensionar colunas excedentes na aba temporária
    const maxCols = sheetTemp.getMaxColumns();
    if (maxCols > neededCols) {
      sheetTemp.deleteColumns(neededCols + 1, maxCols - neededCols);
    }

    // Alocar todas as linhas necessárias de uma só vez antes de gravar
    const maxRows = sheetTemp.getMaxRows();
    if (totalLinhas > maxRows) {
      sheetTemp.insertRowsAfter(maxRows, totalLinhas - maxRows);
    }
    SpreadsheetApp.flush(); // Garante que a estrutura seja processada antes de gravar

    // Gravação em lotes na aba TEMPORÁRIA (base real intacta durante todo esse processo)
    let linhaInicio = 1;
    while (linhaInicio <= totalLinhas) {
      const lote = dadosFiltrados.slice(linhaInicio - 1, linhaInicio - 1 + TAMANHO_LOTE);
      sheetTemp.getRange(linhaInicio, 1, lote.length, neededCols).setValues(lote);
      SpreadsheetApp.flush(); // Força o Sheets a processar o lote antes do próximo
      linhaInicio += TAMANHO_LOTE;
    }

    // Validação: confirmar que a aba temporária tem o número correto de linhas
    const linhasGravadas = sheetTemp.getLastRow();
    if (linhasGravadas < totalLinhas) {
      // Gravação incompleta — remover aba temporária e abortar sem tocar na base real
      ss.deleteSheet(sheetTemp);
      throw new Error(
        `Gravação incompleta na aba temporária: esperado ${totalLinhas} linhas, encontrado ${linhasGravadas}. ` +
        `A base 'Associados' original foi preservada.`
      );
    }

    // ── SWAP ATÔMICO ────────────────────────────────────────────────────────
    // Gravação validada: agora é seguro substituir a aba real
    if (sheetReal) {
      ss.deleteSheet(sheetReal);
    }
    sheetTemp.setName('Associados');
    SpreadsheetApp.flush();
    // ────────────────────────────────────────────────────────────────────────

    _setConfigInterna('DATA_ATUALIZACAO_ASSOCIADOS', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"));

    return { sucesso: true, linhas: totalLinhas - 1 };
  } catch (e) {
    return { sucesso: false, erro: "Erro no servidor ao gravar: " + e.message };
  }
}


/**
 * FALLBACK MANUAL — Atualiza colunas MEI e Simples diretamente na aba Associados.
 * Recebe array 2D com colunas [OptanteMei, bl_optante_simples_nacional, nr_contacorrente].
 * Usado como alternativa quando o Sync Automático não está disponível.
 * (Anteriormente: substituirDadosPJ — agora aponta para Associados)
 */
function atualizarMeiSimplesNaAssociados(dados2D) {
  try {
    validarAcessoAdmin();
    if (!dados2D || dados2D.length < 2) return { sucesso: false, erro: "Nenhum dado recebido." };
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Associados');
    if (!sheet) return { sucesso: false, erro: "Aba 'Associados' não encontrada." };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idxConta = headers.findIndex(h => String(h).toUpperCase().includes('CONTA'));
    const idxMei = headers.findIndex(h => String(h).toUpperCase().includes('MEI'));
    const idxSimples = headers.findIndex(h => String(h).toUpperCase().includes('SIMPLES'));

    if (idxConta === -1 || idxMei === -1 || idxSimples === -1) {
      return { sucesso: false, erro: "Colunas MEI, Simples ou Conta não encontradas na aba Associados." };
    }

    // Monta mapa conta → {mei, simples} a partir dos dados recebidos
    // Espera cabeçalho na linha 0: [OptanteMei, bl_optante_simples_nacional, nr_contacorrente]
    const mapa = {};
    for (let i = 1; i < dados2D.length; i++) {
      const mei = String(dados2D[i][0]).trim();
      const simples = String(dados2D[i][1]).trim();
      const conta = String(dados2D[i][2]).trim();
      if (conta) mapa[conta] = { mei, simples };
    }

    let atualizados = 0;
    for (let i = 1; i < data.length; i++) {
      const conta = String(data[i][idxConta]).trim();
      if (mapa[conta]) {
        sheet.getRange(i + 1, idxMei + 1).setValue(mapa[conta].mei);
        sheet.getRange(i + 1, idxSimples + 1).setValue(mapa[conta].simples);
        atualizados++;
      }
    }

    _setConfigInterna('DATA_ATUALIZACAO_ASSOCIADOS', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"));
    return { sucesso: true, linhas: atualizados };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}


function substituirModalidades(dados2D) {
  try {
    validarAcessoAdmin();
    if (!dados2D || dados2D.length === 0) return { sucesso: false, erro: "Nenhum dado recebido." };
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); const sheet = ss.getSheetByName('modalidades'); if (!sheet) return { sucesso: false, erro: "Aba 'modalidades' não encontrada." };
    sheet.clear();

    let versaoExtraida = "Versão não identificada";
    let idxHeaderOffset = -1;
    for (let r = 0; r < Math.min(dados2D.length, 10); r++) {
      for (let c = 0; c < dados2D[r].length; c++) {
        if (!dados2D[r][c]) continue;
        let celula = String(dados2D[r][c]).toUpperCase();
        if (celula.includes("VERSÃO") || celula.includes("VERSAO")) {
          let match = celula.match(/(VERS[ÃA]O\s*[\d\.]+\s*\|\s*VIG[ÊE]NCIA[^\n]+)/i);
          if (match) versaoExtraida = match[1].trim();
          else {
            let matchFallback = celula.match(/(VERS[ÃA]O[^\n]+)/i);
            if (matchFallback) versaoExtraida = matchFallback[1].trim();
          }
        }
        if (celula.includes("MODALIDADE") && idxHeaderOffset === -1) {
          idxHeaderOffset = r;
        }
      }
    }

    let targetMapping = [];
    if (idxHeaderOffset !== -1) {
      const headerRow = dados2D[idxHeaderOffset].map(c => String(c).toUpperCase().trim());
      targetMapping[0] = headerRow.findIndex(h => h.includes("PERFIL"));
      targetMapping[1] = headerRow.findIndex(h => h.includes("CÓD") || h.includes("COD"));
      targetMapping[2] = headerRow.findIndex(h => h.includes("MODALIDADE"));
      targetMapping[3] = headerRow.findIndex(h => h.includes("FINALIDADE"));
      targetMapping[4] = headerRow.findIndex(h => h.includes("PARCELA"));
      targetMapping[5] = headerRow.findIndex(h => h.includes("CARÊNCIA") || h.includes("CARENCIA"));
      targetMapping[6] = headerRow.findIndex(h => h.includes("PRAZO"));
      targetMapping[7] = headerRow.findIndex(h => h.includes("BAIXÍSSIMO") || h.includes("BAIXISSIMO"));
      targetMapping[8] = headerRow.findIndex(h => h.includes("BAIXO") && !h.includes("BAIXÍSSIMO") && !h.includes("BAIXISSIMO"));
      targetMapping[9] = headerRow.findIndex(h => h.includes("MÉDIO") || h.includes("MEDIO"));
      targetMapping[10] = headerRow.findIndex(h => h.includes("ALTO"));
      targetMapping[11] = headerRow.findIndex(h => h.includes("PÚBLICO") || h.includes("PUBLICO"));
      targetMapping[12] = headerRow.findIndex(h => h.includes("GARANTIA"));
      targetMapping[13] = headerRow.findIndex(h => h.includes("DESCRIÇÃO") || h.includes("DESCRICAO") || h.includes("OBS"));
      targetMapping[14] = headerRow.findIndex(h => h.includes("TAGS"));
      targetMapping[15] = headerRow.findIndex(h => h.includes("AMORTIZAÇÃO") || h.includes("AMORTIZACAO"));
    } else {
      for (let i = 0; i < 16; i++) targetMapping[i] = i;
    }

    let dadosLimpos = idxHeaderOffset !== -1 ? dados2D.slice(idxHeaderOffset + 1) : dados2D;
    // O filtro procura a coluna designada para CODIGO (targetMapping[1]) ou MODALIDADE (targetMapping[2])
    let idxCod = targetMapping[1] !== -1 ? targetMapping[1] : 1;
    let idxMod = targetMapping[2] !== -1 ? targetMapping[2] : 2;
    dadosLimpos = dadosLimpos.filter(linha => (linha[idxCod] && String(linha[idxCod]).trim()) || (linha[idxMod] && String(linha[idxMod]).trim()));

    const headersBase = ["Perfil", "Código", "Modalidade", "Finalidade", "Parcela", "Carência Máx", "Prazo Máx", "Taxa Baixíssimo", "Taxa Baixo", "Taxa Médio", "Taxa Alto", "Público-Alvo", "Garantia", "Descrição", "Tags", "Amortização"];
    let dadosFormatados = [headersBase];

    dadosLimpos.forEach((linha) => {
      let novaLinha = [];
      for (let i = 0; i < 16; i++) {
        let colIdx = targetMapping[i];
        if (colIdx !== -1 && colIdx < linha.length) {
          novaLinha.push(linha[colIdx]);
        } else {
          novaLinha.push("");
        }
      }
      dadosFormatados.push(novaLinha);
    });

    sheet.getRange(1, 1, dadosFormatados.length, 16).setValues(dadosFormatados); formatarCabecalho(sheet, 16);

    _setConfigInterna('DATA_ATUALIZACAO_MODALIDADES', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"));
    _setConfigInterna('VERSAO_MODALIDADE', versaoExtraida);

    return { sucesso: true, linhas: dadosFormatados.length - 1 };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

// ==========================================
// ENRIQUECIMENTO MANUAL DE MODALIDADES (ATUALIZADO)
// ==========================================
function enriquecerModalidades(dados) {
  try {
    validarAcessoAdmin();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Modalidades');
    if (!sheet) return { sucesso: false, erro: "Aba 'Modalidades' não encontrada no banco de dados." };

    const sheetData = sheet.getDataRange().getValues();
    if (sheetData.length <= 1) return { sucesso: false, erro: "A Base de Modalidades está vazia." };

    const sheetHeaders = sheetData[0];
    const uploadHeaders = dados[0];

    // Procura onde está a coluna "Código" no ficheiro Excel enviado
    const idxCodUpload = uploadHeaders.findIndex(h => String(h).toUpperCase().trim() === 'CÓDIGO' || String(h).toUpperCase().trim() === 'CODIGO');
    if (idxCodUpload === -1) return { sucesso: false, erro: "Coluna 'Código' não encontrada no ficheiro de upload." };

    // Procura onde está a coluna "Código" na base de dados (Planilha)
    const idxCodSheet = sheetHeaders.findIndex(h => String(h).toUpperCase().trim() === 'CÓDIGO' || String(h).toUpperCase().trim() === 'CODIGO');
    if (idxCodSheet === -1) return { sucesso: false, erro: "A coluna 'Código' sumiu da aba Modalidades." };

    // Mapeia todas as colunas que vieram no ficheiro de enriquecimento (exceto o Código)
    const colunasEnriquecimento = [];
    for (let i = 0; i < uploadHeaders.length; i++) {
      if (i !== idxCodUpload && String(uploadHeaders[i]).trim() !== "") {
        colunasEnriquecimento.push({ name: String(uploadHeaders[i]).trim(), indexUpload: i });
      }
    }

    // Verifica se as colunas novas (ex: Periodicidades, Regras IOF) existem na Planilha. Se não, cria-as!
    colunasEnriquecimento.forEach(col => {
      let idxSheet = sheetHeaders.indexOf(col.name);
      if (idxSheet === -1) {
        // Cria o novo cabeçalho na última coluna vazia
        sheet.getRange(1, sheetHeaders.length + 1).setValue(col.name);
        sheetHeaders.push(col.name);
        col.indexSheet = sheetHeaders.length - 1;
      } else {
        col.indexSheet = idxSheet;
      }
    });

    let linhasAtualizadas = 0;

    // Percorre as linhas do ficheiro enviado
    for (let i = 1; i < dados.length; i++) {
      let rowUpload = dados[i];
      let codUpload = String(rowUpload[idxCodUpload]).trim();
      if (!codUpload) continue;

      // Procura a modalidade correspondente na base de dados
      for (let j = 1; j < sheetData.length; j++) {
        if (String(sheetData[j][idxCodSheet]).trim() === codUpload) {

          // Atualiza as células com as novas informações
          colunasEnriquecimento.forEach(col => {
            let valor = rowUpload[col.indexUpload];
            // Só atualiza se o Excel enviado tiver algum valor escrito (não apaga dados antigos se a célula vier vazia)
            if (valor !== undefined && valor !== "") {
              sheet.getRange(j + 1, col.indexSheet + 1).setValue(valor);
            }
          });

          linhasAtualizadas++;
          // Sem break: continua a varrer para enriquecer TODAS as linhas com o mesmo código
        }
      }
    }

    return { sucesso: true, linhas: linhasAtualizadas };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

// ==========================================
// REGISTRO DE SIMULAÇÕES — ID E VERSÕES (17/10/2026)
// ==========================================
// Cada simulação tem um ID numérico sequencial (coluna Z) e versões (AA).
// "Simulação de trabalho" (decisão do negócio): recalcular com o mesmo cliente
// no mesmo módulo grava uma VERSÃO nova do mesmo ID; trocar de cliente ou
// "Nova simulação" gera ID novo; as N linhas de um lote têm o mesmo ID e versão;
// editar pelo histórico continua o ID. O registro original nunca é alterado.
//
// O log é gravado por POSIÇÃO: colunas novas só no fim, e LOG_COL é a única
// fonte dos índices. Os JSONs (AB a AD) não trafegam na lista do histórico —
// só em buscarDetalheSimulacao, sob demanda.
const LOG_COL = { CONTA: 4, AGENCIA: 22, GRUPO: 25, ID: 26, VERSAO: 27, PARAMETROS: 28, FOTO_COMPROMETIMENTO: 29, FOTO_CHECKLIST: 30 };
const LOG_TOTAL_COLUNAS = 30;
const LOG_COLUNAS_NOVAS = ['ID Simulação', 'Versão', 'Parâmetros', 'Foto Comprometimento', 'Foto Checklist'];
const PROP_ULTIMO_ID_SIMULACAO = 'ULTIMO_ID_SIMULACAO';
const LOG_LIMITE_JSON = 45000; // célula do Sheets aceita 50 mil caracteres

/** Garante as colunas Z..AD (aba antiga tinha 25) e o cabeçalho delas. */
function _garantirColunasLog_(sheetLog) {
  const max = sheetLog.getMaxColumns();
  if (max < LOG_TOTAL_COLUNAS) sheetLog.insertColumnsAfter(max, LOG_TOTAL_COLUNAS - max);
  const cab = sheetLog.getRange(1, LOG_COL.ID, 1, LOG_COLUNAS_NOVAS.length).getValues()[0];
  if (cab.some((v, i) => String(v).trim() !== LOG_COLUNAS_NOVAS[i])) {
    sheetLog.getRange(1, LOG_COL.ID, 1, LOG_COLUNAS_NOVAS.length).setValues([LOG_COLUNAS_NOVAS]);
  }
}

/** Próximo ID (chamar sob lock). Contador nas Propriedades; sem ele, parte do maior ID da coluna Z. */
function _proximoIdSimulacao_(sheetLog) {
  const props = PropertiesService.getScriptProperties();
  let ultimo = parseInt(props.getProperty(PROP_ULTIMO_ID_SIMULACAO), 10);
  if (isNaN(ultimo)) ultimo = _maiorIdSimulacao_(sheetLog);
  ultimo++;
  props.setProperty(PROP_ULTIMO_ID_SIMULACAO, String(ultimo));
  return ultimo;
}

function _maiorIdSimulacao_(sheetLog) {
  const n = sheetLog.getLastRow();
  if (n < 2 || sheetLog.getMaxColumns() < LOG_COL.ID) return 0;
  let maior = 0;
  sheetLog.getRange(2, LOG_COL.ID, n - 1, 1).getValues().forEach(r => {
    const v = parseInt(r[0], 10);
    if (v > maior) maior = v;
  });
  return maior;
}

/** Linhas (texto exibido, sem Date) de um ID: [{ linha, row }], na ordem da planilha. */
function _linhasDoIdSimulacao_(sheetLog, id) {
  if (sheetLog.getMaxColumns() < LOG_COL.ID) return [];
  const cand = _candidatas_(sheetLog, LOG_COL.ID, _regexDigitos_(String(id)), Infinity);
  if (!cand.length) return [];
  let lidas;
  if (cand.length <= 30) {
    lidas = cand.map(linha => ({ linha, row: sheetLog.getRange(linha, 1, 1, LOG_TOTAL_COLUNAS).getDisplayValues()[0] }));
  } else {
    const primeira = cand[0];
    const bloco = sheetLog.getRange(primeira, 1, cand[cand.length - 1] - primeira + 1, LOG_TOTAL_COLUNAS).getDisplayValues();
    lidas = cand.map(linha => ({ linha, row: bloco[linha - primeira] }));
  }
  return lidas.filter(a => parseInt(_apenasDigitos_(a.row[LOG_COL.ID - 1]), 10) === id);
}

/** Perfil AGÊNCIA só alcança simulações da própria agência (mesma regra do histórico). */
function _podeVerSimulacao_(usuario, row) {
  const perfil = String(usuario.perfil || 'AGÊNCIA').trim().toUpperCase();
  if (perfil !== 'AGÊNCIA') return true;
  return String(row[LOG_COL.AGENCIA - 1] || 'Legado').trim() === String(usuario.agencia || 'Legado').trim();
}

/** Texto do cliente que começa com "=" viraria fórmula na planilha. */
function _textoSeguroLog_(v) {
  return (typeof v === 'string' && /^[=]/.test(v)) ? "'" + v : v;
}

function _jsonLog_(obj) {
  if (obj == null || obj === '') return '';
  const json = JSON.stringify(obj);
  return json.length > LOG_LIMITE_JSON ? JSON.stringify({ erro: 'Conteúdo grande demais para guardar (' + json.length + ' caracteres).' }) : json;
}

/**
 * Grava as linhas de UMA simulação (1 na Simples, N num lote), todas com o
 * mesmo ID e versão. Sem idSimulacao: ID novo, versão 1. Com idSimulacao:
 * versão seguinte desse ID — e o servidor confere se o usuário pode vê-lo.
 * Devolve { sucesso, id, versao }.
 */
function registrarSimulacao(pacotes, idSimulacao) {
  const lock = LockService.getScriptLock();
  try {
    const usuario = validarAcessoOperacional();
    if (!Array.isArray(pacotes) || !pacotes.length) return { sucesso: false, erro: 'Nada a registrar.' };
    if (pacotes.length > 200) return { sucesso: false, erro: 'Simulação com linhas demais.' };
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetLog = ss.getSheetByName('Log Simulacoes');
    if (!sheetLog) return { sucesso: false, erro: "Aba de Log não encontrada." };

    lock.waitLock(20000);
    _garantirColunasLog_(sheetLog);

    let id, versao;
    if (idSimulacao != null && idSimulacao !== '') {
      id = parseInt(idSimulacao, 10);
      if (!(id > 0)) return { sucesso: false, erro: 'ID de simulação inválido.' };
      const existentes = _linhasDoIdSimulacao_(sheetLog, id);
      if (!existentes.length) return { sucesso: false, erro: `Simulação nº ${id} não encontrada.` };
      if (!_podeVerSimulacao_(usuario, existentes[0].row)) return { sucesso: false, erro: 'Você não tem permissão para editar esta simulação.' };
      versao = Math.max.apply(null, existentes.map(e => parseInt(e.row[LOG_COL.VERSAO - 1], 10) || 1)) + 1;
    } else {
      id = _proximoIdSimulacao_(sheetLog);
      versao = 1;
    }

    const dataHora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    const emailOperador = Session.getActiveUser().getEmail();
    const linhas = pacotes.map((dados, i) => [
      dataHora,
      dados.operador,
      emailOperador,
      dados.conta,
      dados.nomeAssociado,
      dados.modalidade,
      dados.valorFinanciado,
      dados.carencia,
      dados.periodicidade,
      dados.prazo,
      dados.tac,
      dados.iof,
      dados.seguro,
      dados.cota,
      dados.custas,
      dados.totalCustas,
      dados.valorLiquido,
      dados.taxaMinima,
      dados.taxaEfetiva,
      dados.primeiraParcela, // Coluna T
      dados.ultimaParcela,    // Coluna U
      // Coluna V — agência do usuário AUTENTICADO, não a enviada pelo navegador:
      // o histórico filtra o perfil AGÊNCIA por esta coluna, então ela não pode
      // ser escolhida pelo cliente. (Até 15/10/2026 o Lote por Modalidades
      // mandava sempre "Sede" — lia um elemento inexistente da tela.)
      usuario.agencia || dados.agencia,
      dados.score,           // Coluna W
      dados.tipoSimulacao || "SIMPLES", // Coluna X: SIMPLES | COMPOSTA | LOTE_PRAZOS | LOTE_MODALIDADES
      dados.grupoSimulacao || "",       // Coluna Y: agrupa as linhas de um lote
      id,                               // Coluna Z
      versao,                           // Coluna AA
      i === 0 ? _jsonLog_(dados.parametros) : '', // Coluna AB: parâmetros (só na 1ª linha da versão)
      '', ''                            // AC/AD: fotos, gravadas depois por registrarFotoSimulacao
    ].map(v => _textoSeguroLog_(v === undefined || v === null ? '' : v)));

    const inicio = sheetLog.getLastRow() + 1;
    const faltam = inicio + linhas.length - 1 - sheetLog.getMaxRows();
    if (faltam > 0) sheetLog.insertRowsAfter(sheetLog.getMaxRows(), faltam);
    sheetLog.getRange(inicio, 1, linhas.length, LOG_TOTAL_COLUNAS).setValues(linhas);
    SpreadsheetApp.flush();
    return { sucesso: true, id, versao };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  } finally {
    lock.releaseLock();
  }
}

/** Compatibilidade: uma linha. O frontend usa registrarSimulacao desde 17/10/2026. */
function registrarLogSimulacao(dados) {
  return registrarSimulacao([dados || {}], dados ? dados.idSimulacao : null);
}

/**
 * Guarda a "foto" do comprometimento ou do checklist numa versão. Os dois
 * costumam ser feitos DEPOIS do registro da simulação, por isso é separado.
 * Comprometimento: lista (uma entrada por parcela avaliada — no lote cada
 * prazo/modalidade tem a sua), a mais recente substitui a da mesma parcela.
 * Checklist: a mais recente substitui a anterior.
 */
function registrarFotoSimulacao(idSimulacao, versao, tipo, foto) {
  const lock = LockService.getScriptLock();
  try {
    const usuario = validarAcessoOperacional();
    if (tipo !== 'comprometimento' && tipo !== 'checklist') return { sucesso: false, erro: 'Tipo de foto inválido.' };
    const id = parseInt(idSimulacao, 10), ver = parseInt(versao, 10);
    if (!(id > 0) || !(ver > 0) || !foto || typeof foto !== 'object') return { sucesso: false, erro: 'Dados da foto inválidos.' };
    const sheetLog = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Log Simulacoes');
    if (!sheetLog) return { sucesso: false, erro: "Aba de Log não encontrada." };

    lock.waitLock(20000);
    _garantirColunasLog_(sheetLog);
    const daVersao = _linhasDoIdSimulacao_(sheetLog, id).filter(a => (parseInt(a.row[LOG_COL.VERSAO - 1], 10) || 1) === ver);
    if (!daVersao.length) return { sucesso: false, erro: `Simulação nº ${id} v${ver} não encontrada.` };
    if (!_podeVerSimulacao_(usuario, daVersao[0].row)) return { sucesso: false, erro: 'Você não tem permissão sobre esta simulação.' };

    const col = tipo === 'comprometimento' ? LOG_COL.FOTO_COMPROMETIMENTO : LOG_COL.FOTO_CHECKLIST;
    const registro = Object.assign({}, foto, {
      registradoEm: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
      registradoPor: usuario.nome || Session.getActiveUser().getEmail()
    });
    let valor = registro;
    if (tipo === 'comprometimento') {
      let lista = [];
      try { lista = JSON.parse(daVersao[0].row[col - 1] || '[]'); } catch (e) { lista = []; }
      if (!Array.isArray(lista)) lista = [];
      lista = lista.filter(f => !(f.parcela === registro.parcela && f.periodicidade === registro.periodicidade));
      lista.push(registro);
      valor = lista.slice(-20);
    }
    sheetLog.getRange(daVersao[0].linha, col).setValue(_jsonLog_(valor));
    SpreadsheetApp.flush();
    return { sucesso: true };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  } finally {
    lock.releaseLock();
  }
}

/** Linha do log (texto exibido) → objeto do histórico. */
function _logDaLinha_(row) {
  return {
    dataHora: row[0], operador: row[1], email: row[2], conta: row[3], nomeAssociado: row[4],
    modalidade: row[5], valorFinanciado: row[6], carencia: row[7], periodicidade: row[8],
    prazo: row[9], tac: row[10], iof: row[11], seguro: row[12], cota: row[13],
    custas: row[14], totalCustas: row[15], valorLiquido: row[16], taxaMinima: row[17], taxaEfetiva: row[18],
    primeiraParcela: row[19] || "---",
    ultimaParcela: row[20] || "---",
    agencia: row[21] || "Legado",
    score: row[22] || "Sem Score",
    tipoSimulacao: row[23] || "SIMPLES",  // Coluna X
    grupoSimulacao: row[24] || "",         // Coluna Y
    idSimulacao: parseInt(_apenasDigitos_(row[25] || ''), 10) || null, // Coluna Z
    versao: parseInt(row[26], 10) || (row[25] ? 1 : null)              // Coluna AA
  };
}

// ==========================================
// BUSCAR HISTÓRICO DE SIMULAÇÕES (COM CONTROLE DE ACESSO)
// ==========================================
function buscarHistoricoSimulacoes() {
  try {
    const usuario = validarAcessoOperacional();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetLog = ss.getSheetByName('Log Simulacoes');
    if (!sheetLog) return { sucesso: false, erro: "Aba 'Log Simulacoes' não encontrada." };

    // Só A..AA: os JSONs (AB..AD) ficam para buscarDetalheSimulacao.
    const ultimaLinha = sheetLog.getLastRow();
    if (ultimaLinha <= 1) return { sucesso: true, dados: [], perfil: "AGÊNCIA" };
    const numCols = Math.min(LOG_COL.VERSAO, sheetLog.getMaxColumns());
    const data = sheetLog.getRange(1, 1, ultimaLinha, numCols).getDisplayValues();

    // Perfil e agência vêm do usuário validado (Hub). O perfil AGÊNCIA só vê as
    // linhas cuja coluna V tem o nome da sua agência no Simulador.
    const perfilAtual = String(usuario.perfil || "AGÊNCIA").trim().toUpperCase();
    const agenciaAtual = String(usuario.agencia || "Legado").trim();

    const logs = [];
    for (let i = data.length - 1; i >= 1; i--) {
      let row = data[i];
      if (perfilAtual === "AGÊNCIA" && String(row[21] || "Legado").trim() !== agenciaAtual) continue;
      logs.push(_logDaLinha_(row));
    }
    return { sucesso: true, dados: logs, perfil: perfilAtual };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Todas as versões de uma simulação, com parâmetros e fotos (JSON), da mais
 * recente para a mais antiga: [{ versao, linhas: [log], parametros, fotoComprometimento, fotoChecklist }].
 */
function buscarDetalheSimulacao(idSimulacao) {
  try {
    const usuario = validarAcessoOperacional();
    const id = parseInt(idSimulacao, 10);
    if (!(id > 0)) return { sucesso: false, erro: 'ID de simulação inválido.' };
    const sheetLog = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Log Simulacoes');
    if (!sheetLog) return { sucesso: false, erro: "Aba 'Log Simulacoes' não encontrada." };
    const lidas = _linhasDoIdSimulacao_(sheetLog, id);
    if (!lidas.length) return { sucesso: false, erro: `Simulação nº ${id} não encontrada.` };
    if (!_podeVerSimulacao_(usuario, lidas[0].row)) return { sucesso: false, erro: 'Você não tem permissão para ver esta simulação.' };

    const lerJson = txt => { if (!txt) return null; try { return JSON.parse(txt); } catch (e) { return null; } };
    const porVersao = {};
    lidas.forEach(({ row }) => {
      const log = _logDaLinha_(row);
      const v = log.versao || 1;
      const item = porVersao[v] || (porVersao[v] = { versao: v, linhas: [], parametros: null, fotoComprometimento: null, fotoChecklist: null });
      item.linhas.push(log);
      item.parametros = item.parametros || lerJson(row[LOG_COL.PARAMETROS - 1]);
      item.fotoComprometimento = item.fotoComprometimento || lerJson(row[LOG_COL.FOTO_COMPROMETIMENTO - 1]);
      item.fotoChecklist = item.fotoChecklist || lerJson(row[LOG_COL.FOTO_CHECKLIST - 1]);
    });
    const versoes = Object.keys(porVersao).map(Number).sort((a, b) => b - a).map(v => porVersao[v]);
    return { sucesso: true, id, versoes };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

/**
 * ADMIN — rodar UMA vez no editor depois de publicar (17/10/2026). Numera as
 * simulações antigas em ordem de gravação: cada linha da Simples ganha um ID;
 * as linhas de um mesmo lote (coluna Y) ganham o mesmo ID; todas viram versão 1.
 * Idempotente: linhas que já têm ID são mantidas. Acerta o contador.
 */
function numerarSimulacoesAntigas() {
  const lock = LockService.getScriptLock();
  try {
    validarAcessoAdmin();
    const sheetLog = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Log Simulacoes');
    if (!sheetLog) return 'Aba Log Simulacoes não encontrada.';
    lock.waitLock(30000);
    _garantirColunasLog_(sheetLog);
    const n = sheetLog.getLastRow();
    if (n < 2) return 'Log vazio.';

    const grupos = sheetLog.getRange(2, LOG_COL.GRUPO, n - 1, 1).getDisplayValues();
    const idVer = sheetLog.getRange(2, LOG_COL.ID, n - 1, 2).getValues();
    let ultimo = 0;
    idVer.forEach(r => { const v = parseInt(r[0], 10); if (v > ultimo) ultimo = v; });
    const idDoGrupo = {};
    idVer.forEach((r, i) => { const g = String(grupos[i][0]).trim(); if (g && parseInt(r[0], 10) > 0) idDoGrupo[g] = parseInt(r[0], 10); });

    let novos = 0;
    const saida = idVer.map((r, i) => {
      const atual = parseInt(r[0], 10);
      if (atual > 0) return [atual, parseInt(r[1], 10) || 1];
      const g = String(grupos[i][0]).trim();
      let id;
      if (g && idDoGrupo[g]) id = idDoGrupo[g];
      else { id = ++ultimo; if (g) idDoGrupo[g] = id; novos++; }
      return [id, 1];
    });
    sheetLog.getRange(2, LOG_COL.ID, n - 1, 2).setValues(saida);
    SpreadsheetApp.flush();
    PropertiesService.getScriptProperties().setProperty(PROP_ULTIMO_ID_SIMULACAO, String(ultimo));
    const msg = `Numeração concluída: ${novos} simulações novas numeradas (${n - 1} linhas no log). Último ID: ${ultimo}.`;
    Logger.log(msg);
    return msg;
  } catch (e) {
    return 'Erro na numeração: ' + e.message;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Função de Ajuste Único: Recupera a agência de logs antigos baseando-se no e-mail do operador.
 */
function ajusteRetroativoAgencias() {
  try {
    validarAcessoAdmin();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetLog = ss.getSheetByName('Log Simulacoes');
    const sheetUsers = ss.getSheetByName('Usuarios');

    if (!sheetLog || !sheetUsers) return "Abas necessárias não encontradas.";

    const logs = sheetLog.getRange(1, 1, sheetLog.getLastRow(), 22).getValues();
    const users = sheetUsers.getRange(1, 1, sheetUsers.getLastRow(), 4).getValues();

    // Criar um mapa de Email -> Agência para busca rápida
    const mapaAgencias = {};
    for (let i = 1; i < users.length; i++) {
      let email = String(users[i][1]).trim().toLowerCase();
      let agencia = String(users[i][3]).trim();
      if (email) mapaAgencias[email] = agencia;
    }

    let logsAtualizados = 0;
    const numLinhas = logs.length;

    for (let i = 1; i < numLinhas; i++) {
      let logEmail = String(logs[i][2]).trim().toLowerCase(); // Coluna C (Índice 2)
      let logAgenciaAtual = String(logs[i][21] || "").trim(); // Coluna V (Índice 21)

      if (!logAgenciaAtual || logAgenciaAtual === "Legado") {
        if (mapaAgencias[logEmail]) {
          sheetLog.getRange(i + 1, 22).setValue(mapaAgencias[logEmail]); // Escreve na coluna V (22)
          logsAtualizados++;
        }
      }
    }

    return `Ajuste concluído. ${logsAtualizados} registros atualizados.`;
  } catch (e) {
    return "Erro no ajuste: " + e.message;
  }
}

// ==========================================
// BUSCAR CATÁLOGO DE MODALIDADES (CORREÇÃO DE TAGS E LEITURA INFINITA BLINDADA)
// ==========================================
function buscarModalidadesGlobais() {
  try {
    validarAcessoOperacional();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Modalidades');
    if (!sheet) return { sucesso: false, erro: "Aba 'Modalidades' não encontrada." };

    // CORREÇÃO MESTRA: getDisplayValues() força a leitura como texto puro, 
    // evitando crashes silenciosos com datas ou formatações do Excel
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return { sucesso: false, dados: [] };

    const headers = data[0];
    const modalidades = [];

    // Leitura dinâmica infinita de colunas
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      let obj = {};

      for (let j = 0; j < headers.length; j++) {
        let colName = String(headers[j]).trim();
        if (colName) {
          obj[colName] = row[j];
        }
      }

      // Só adiciona se a linha tiver um Código válido
      if (obj['Código'] || obj['Codigo']) {
        modalidades.push(obj);
      }
    }

    // Busca das Tags Globais Ativas
    let tagsAtivas = [];
    try {
      const config = buscarConfiguracoesAdmin();
      if (config && config.ativas) {
        tagsAtivas = config.ativas;
      }
    } catch (errTag) {
      // Fallback silencioso
    }

    return { sucesso: true, dados: modalidades, configGlobais: { ativas: tagsAtivas } };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

// ==========================================
// ==========================================
// PESQUISA AVANÇADA POR NOME (PARA O POP-UP)
// ==========================================
function pesquisarNomeAssociado(termo) {
  try {
    validarAcessoOperacional();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Associados');
    if (!sheet) return { sucesso: false, erro: "Aba de Associados não encontrada." };

    const termoUpper = String(termo).toUpperCase().trim();
    if (termoUpper.length < 3) return { sucesso: false, erro: "Digite pelo menos 3 letras para pesquisar." };

    const totalLinhas = sheet.getLastRow();
    if (totalLinhas < 2) return { sucesso: true, dados: [] };

    // Colunas fixas: [A]=Nome Coop, [B]=Associado, [C]=Conta, [D]=Score,
    // [E]=Tipo Pessoa, [F]=Idade, [G]=MEI, [H]=Simples, [I]=CPF/CNPJ.
    // Desempenho: UMA leitura (colunas B a I). Em 14/10/2026 tentei ler só a
    // coluna de nomes e depois os detalhes (3 leituras menores) e ficou PIOR —
    // 7,5 a 12,6 s contra 7,7 s —, porque nesta planilha cada leitura tem custo
    // fixo alto (1,3 a 4 s por coluna). Aqui o número de leituras pesa mais que
    // o número de colunas. (A coluna A, Nome Coop, não é usada.)
    const montar = (linha, r) => ({
      linhaSheet: linha,                          // linha real no Sheets (1-indexed + cabeçalho)
      nome: String(r[0] || 'Nome Indisponível'),
      conta: String(r[1] || ''),
      score: String(r[2] || ''),
      pessoa: _normalizarPessoa(r[3]),
      documento: String(r[7] || '')               // CPF/CNPJ
    });

    // Caminho rápido (15/10/2026): o operador costuma digitar um nome quase
    // completo, com poucas ocorrências. A busca nativa (sem regex, sem
    // diferenciar maiúsculas) acha as linhas sem ler a base; com até
    // NOME_LIMITE_RAPIDO ocorrências, lê só elas. Cada uma é conferida pela
    // mesma regra da leitura completa, então o resultado é igual.
    const NOME_LIMITE_RAPIDO = 15;
    const finder = sheet.getRange(2, 2, totalLinhas - 1, 1).createTextFinder(String(termo).trim()).matchCase(false);
    const linhasRapidas = [];
    let achado;
    while ((achado = finder.findNext())) {
      linhasRapidas.push(achado.getRow());
      if (linhasRapidas.length > NOME_LIMITE_RAPIDO) break;
    }
    if (linhasRapidas.length <= NOME_LIMITE_RAPIDO) {
      linhasRapidas.sort((a, b) => a - b);
      const dados = [];
      linhasRapidas.forEach(linha => {
        const r = sheet.getRange(linha, 2, 1, 8).getValues()[0]; // B..I
        if (String(r[0] || '').toUpperCase().includes(termoUpper)) dados.push(montar(linha, r));
      });
      _marcarLimitesNaLista_(dados, ss);
      return { sucesso: true, dados };
    }

    // Termo genérico (muitas ocorrências): leitura única de B..I, como antes.
    const bloco = sheet.getRange(2, 2, totalLinhas - 1, 8).getValues(); // B..I

    const resultados = [];
    for (let i = 0; i < bloco.length; i++) {
      if (String(bloco[i][0] || '').toUpperCase().includes(termoUpper)) {
        resultados.push({
          linhaSheet: i + 2,                          // linha real no Sheets (1-indexed + cabeçalho)
          nome: String(bloco[i][0] || 'Nome Indisponível'),
          conta: String(bloco[i][1] || ''),
          score: String(bloco[i][2] || ''),
          pessoa: _normalizarPessoa(bloco[i][3]),
          documento: String(bloco[i][7] || '')       // CPF/CNPJ
        });
        if (resultados.length >= 100) break;
      }
    }

    _marcarLimitesNaLista_(resultados, ss);
    return { sucesso: true, dados: resultados };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

// ==========================================
// BUSCA DE ASSOCIADO POR NÚMERO DE LINHA (Fase 2)
// ==========================================
function buscarAssociadoPorLinha(linhaNum) {
  try {
    validarAcessoOperacional();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Associados');
    if (!sheet) return { sucesso: false, erro: "Aba de Associados não encontrada." };

    const numCols = sheet.getLastColumn();
    if (numCols === 0) return { sucesso: false, erro: "Aba Associados vazia." };

    // O índice vem do cliente: validar antes de usar em getRange().
    // Linha 1 é cabeçalho, por isso o mínimo é 2.
    const linha = Number(linhaNum);
    const ultimaLinha = sheet.getLastRow();
    if (!Number.isInteger(linha) || linha < 2 || linha > ultimaLinha) {
      return { sucesso: false, erro: "Registro de associado inválido." };
    }

    const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];
    const row = sheet.getRange(linha, 1, 1, numCols).getValues()[0];

    const limparDado = (valor) => {
      if (valor instanceof Date) {
        let d = valor.getDate(); let m = valor.getMonth() + 1; let y = valor.getFullYear();
        return `${d < 10 ? '0' + d : d}/${m < 10 ? '0' + m : m}/${y}`;
      }
      return valor == null ? "" : String(valor).trim();
    };

    const idxNome = headers.findIndex(h => { const s = String(h).toUpperCase().trim(); return s === 'ASSOCIADO' || s === 'NOME DO ASSOCIADO' || s === 'NOME COMPLETO' || s === 'NOME'; });
    const idxConta = headers.findIndex(h => String(h).toUpperCase().includes('CONTA'));
    const idxScore = headers.findIndex(h => String(h).toUpperCase().includes('SCORE'));
    const idxIdade = headers.findIndex(h => String(h).toUpperCase().includes('IDADE') || String(h).toUpperCase().includes('NASCIMENTO'));
    const idxPessoa = headers.findIndex(h => String(h).toUpperCase().includes('TIPO') && String(h).toUpperCase().includes('PESSOA'));
    const idxMei = headers.findIndex(h => String(h).toUpperCase().includes('MEI'));
    const idxSimples = headers.findIndex(h => String(h).toUpperCase().includes('SIMPLES'));
    const idxCpf = headers.findIndex(h => String(h).toUpperCase().includes('CPF'));
    const idxRenda = headers.findIndex(h => String(h).toUpperCase().includes('RENDA') && !String(h).toUpperCase().includes('DATA'));
    const idxDtRenda = headers.findIndex(h => (String(h).toUpperCase().includes('DATA') && String(h).toUpperCase().includes('RENDA')) || String(h).toUpperCase().includes('ATUALIZACAO'));
    const idxColab = headers.findIndex(h => String(h).toUpperCase().includes('COLABORADOR'));
    const idxEndivCP = headers.findIndex(h => String(h).toUpperCase().includes('CURTO'));
    const idxEndivTt = headers.findIndex(h => String(h).toUpperCase().includes('ENDIV') && String(h).toUpperCase().includes('TOTAL'));

    let result = {
      sucesso: true,
      nome: idxNome !== -1 ? limparDado(row[idxNome]) : "Nome Indisponível",
      conta: idxConta !== -1 ? limparDado(row[idxConta]) : "",
      score: idxScore !== -1 ? limparDado(row[idxScore]) : "",
      idade: idxIdade !== -1 ? limparDado(row[idxIdade]) : "",
      pessoa: idxPessoa !== -1 ? _normalizarPessoa(row[idxPessoa]) : "PF",
      mei: idxMei !== -1 ? _normalizarBoolean(row[idxMei]) : "NÃO",
      simples: idxSimples !== -1 ? _normalizarBoolean(row[idxSimples]) : "NÃO",
      cpfCnpj: idxCpf !== -1 ? limparDado(row[idxCpf]) : "",
      renda: idxRenda !== -1 ? limparDado(row[idxRenda]) : "",
      dataRenda: idxDtRenda !== -1 ? limparDado(row[idxDtRenda]) : "",
      colaborador: idxColab !== -1 ? _normalizarBoolean(row[idxColab]) : "NÃO",
      endivCP: idxEndivCP !== -1 ? limparDado(row[idxEndivCP]) : "",
      endivTotal: idxEndivTt !== -1 ? limparDado(row[idxEndivTt]) : ""
    };
    const civisProp = _dadosCivis_(headers, row); // checklist: estado civil do proponente
    result.estadoCivil = civisProp.estadoCivil;
    result.uniaoEstavel = civisProp.uniaoEstavel;

    if (result.pessoa === 'PJ' && result.cpfCnpj) {
      result.sociosPJ = buscarSociosPorCnpj(result.cpfCnpj);
    }

    if (result.cpfCnpj) {
      // T2: passa ss já aberto para buscarGrupoEconomico
      result.grupoEconomico = _pertenceAGrupo_(result.cpfCnpj, ss); // a tela só usa temGrupo
    }

    if (result.conta) result.limites = _limitesDaConta_(result.conta, ss);

    return result;
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

// ==========================================
// GESTÃO DE NOVIDADES E MELHORIAS
// ==========================================
function buscarNovidades() {
  try {
    validarAcessoOperacional();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Novidades');
    if (!sheet) return { sucesso: true, dados: [] };

    const data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return { sucesso: true, dados: [] };

    // Teto de segurança: só as N mais recentes trafegam para o cliente.
    // Evita que um crescimento anômalo da aba trave o navegador do operador.
    const MAX_NOVIDADES = 50;
    const limiteInferior = Math.max(1, data.length - MAX_NOVIDADES);

    const novidades = [];
    // Retorna do mais recente para o mais antigo
    for (let i = data.length - 1; i >= limiteInferior; i--) {
      novidades.push({
        data: data[i][0],
        resumo: data[i][1],
        detalhes: data[i][2],
        autor: data[i][3]
      });
    }
    return { sucesso: true, dados: novidades };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

function registrarNovidade(resumo, detalhes) {
  try {
    validarContextoInterno();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Novidades');
    if (!sheet) {
      sheet = ss.insertSheet('Novidades');
      sheet.appendRow(["Data", "Resumo", "Detalhes", "Autor"]);
      formatarCabecalho(sheet, 4);
    }

    // ── Guarda anti-duplicata ────────────────────────────────────────────────
    // Se o mesmo resumo já foi publicado, não grava de novo. Protege a aba
    // contra reintrodução acidental de chamadas automáticas/repetidas.
    const existentes = sheet.getDataRange().getValues();
    const resumoNormalizado = String(resumo).trim().toLowerCase();
    for (let i = 1; i < existentes.length; i++) {
      if (String(existentes[i][1]).trim().toLowerCase() === resumoNormalizado) {
        return { sucesso: false, erro: 'Novidade já registrada anteriormente: "' + resumo + '".' };
      }
    }

    const emailOperador = Session.getActiveUser().getEmail();
    let nomeAutor = emailOperador;

    // Nome do autor pela mesma validação de acesso do app (Hub, desde 15/10/2026).
    // Rodando pelo editor ou trigger, cai no e-mail.
    try {
      const acesso = _acessoAtual_();
      if (acesso && acesso.autorizado && acesso.nome) nomeAutor = acesso.nome;
    } catch (eNome) { /* mantém o e-mail */ }

    const dataHoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    sheet.appendRow([dataHoje, resumo, detalhes, nomeAutor]);
    return { sucesso: true };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

// ============================================================================
// PESQUISA DE SATISFAÇÃO LEGADA — REMOVIDA EM 09/09/2026
//
// As funções verificarPesquisa() e salvarRespostaPesquisa(), que atendiam ao
// modal de 6 perguntas fixas, foram removidas após a migração para o módulo
// de Pesquisas Admin (abas PesquisasAdmin / RespostasPesquisas).
//
// A aba "Pesquisa" foi preservada na planilha como histórico e continua sendo
// a origem lida por migrarPesquisaLegadaParaAdmin().
// ============================================================================

// ============================================================================
// SINCRONIZAÇÃO AUTOMÁTICA — BI → GOOGLE DRIVE CSV → GOOGLE SHEETS
// Implementado em: 2026-06-03
// Plano: Rotina Agendada (Sync Noturno) com horário configurável via Admin
// ============================================================================

// --- CONFIGURAÇÕES DO SYNC ---
const SYNC_CONFIG = {
  // ID do arquivo CSV no Google Drive (fixo — não requer alteração)
  CSV_FILE_ID: '15jGftcqh9h9-5p5LN1-LFkUr6M8wQl0q',

  // Separador do CSV
  CSV_SEPARATOR: ',',

  // Encoding do CSV
  CSV_ENCODING: 'UTF-8',

  // Mapeamento exato das colunas do CSV para os campos do simulador
  COLUNAS: {
    conta: 'nr_conta_corrente',
    nome: 'nm_nome',
    nascimento: 'dt_nascimento',             // Data de nascimento → calcular idade
    pessoa: 'ds_pessoa_tipo',             // 'F' → 'PF' | 'J' → 'PJ'
    mei: 'bl_optante_mei',             // '1' → 'SIM' | '0' → 'NÃO'
    simples: 'bl_optante_simples_nacional', // '1' → 'SIM' | '0' → 'NÃO'
    // score: não disponível no CSV — será gravado como vazio
    cpfCnpj: 'nr_cpf_cnpj',
    renda: 'vl_anual_fonte_renda_total',
    dataRenda: 'dt_mes_ano_atualizacao_fonte_renda',
    dataSCR: 'dt_situacao_Restr',
    colaborador: 'bl_colaborador',             // '1' → 'SIM' | '0' → 'NÃO'
    endivMensal: 'vl_mensal_endividamento', // Endividamento mensal
    endivTotal: 'vl_12_total',                // Endividamento total
    // Estado civil e cônjuge (27/10/2026) — checklist: documento por estado civil e
    // cônjuge do avalista. Opcionais: coluna ausente no CSV só gera aviso no log.
    estadoCivil: 'ds_estado_civil',
    uniaoEstavel: 'bl_uniao_estavel',         // '1' → 'SIM' | '0' → 'NÃO'
    cpfConjuge: 'nr_cpf_conjuge',
  }
};

// --- CONFIGURAÇÕES DO SYNC DA COMPOSIÇÃO SOCIETÁRIA ---
const SYNC_COMPOSICAO_CONFIG = {
  CSV_FILE_ID: '1kRZHq5jTSuyphz0G3vJdBtN-XnYkgIg6',
  CSV_SEPARATOR: ',',
  CSV_ENCODING: 'UTF-8',
  COLUNAS: {
    cnpj: 'nr_cnpj',
    cpfSocio: 'nr_cpf_composicao_societaria',
    percentual: 'nr_percentual_participacao'
  }
};

// --- FUNÇÕES AUXILIARES DE TRANSFORMAÇÃO ---

/**
 * Calcula a idade em anos a partir de uma string de data de nascimento.
 * Suporta formatos: 'YYYY-MM-DD', 'DD/MM/YYYY', 'DD-MM-YYYY'.
 * @param {string} valorData - O valor de data vindo do CSV.
 * @returns {string} Idade em anos, ou '' se não for possível calcular.
 */
function _calcularIdade(valorData) {
  if (!valorData) return '';
  let dataNasc;
  const str = String(valorData).trim();

  // Tenta formato ISO: YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
    dataNasc = new Date(str.substring(0, 10));
  }
  // Tenta formato BR: DD/MM/YYYY ou DD-MM-YYYY
  else if (/^\d{2}[\/\-]\d{2}[\/\-]\d{4}/.test(str)) {
    const partes = str.split(/[\/\-]/);
    dataNasc = new Date(parseInt(partes[2]), parseInt(partes[1]) - 1, parseInt(partes[0]));
  }
  else {
    // Último recurso: tenta o construtor nativo
    dataNasc = new Date(str);
  }

  if (isNaN(dataNasc.getTime())) return '';

  const hoje = new Date();
  let idade = hoje.getFullYear() - dataNasc.getFullYear();
  const mesAtual = hoje.getMonth();
  const diaAtual = hoje.getDate();
  const mesNasc = dataNasc.getMonth();
  const diaNasc = dataNasc.getDate();

  // Ajusta se ainda não fez aniversário este ano
  if (mesAtual < mesNasc || (mesAtual === mesNasc && diaAtual < diaNasc)) {
    idade--;
  }

  return idade >= 0 ? String(idade) : '';
}

/**
 * Normaliza o tipo de pessoa: 'F' → 'PF', 'J' → 'PJ'.
 * Aceita também 'PF', 'PJ', 'FÍSICA', 'JURÍDICA' como entrada.
 */
function _normalizarPessoa(valor) {
  const v = String(valor || '').trim().toUpperCase();
  if (v === 'F' || v === 'PF' || v.includes('FISICA') || v.includes('FÍSICA')) return 'PF';
  if (v === 'J' || v === 'PJ' || v.includes('JURIDICA') || v.includes('JURÍDICA')) return 'PJ';
  return 'PF'; // Padrão
}

/**
 * Normaliza booleano de enquadramento tributário: '1'/'S'/'SIM'/'TRUE' → 'SIM', resto → 'NÃO'.
 */
function _normalizarBoolean(valor) {
  const v = String(valor || '').trim().toUpperCase();
  if (v === '1' || v === 'S' || v === 'SIM' || v === 'TRUE' || v === 'T') return 'SIM';
  return 'NÃO';
}

// --- FUNÇÕES PRINCIPAIS ---

/**
 * FUNÇÃO PRINCIPAL — Executada pelo Trigger agendado.
 * Orquestra todo o processo de leitura do CSV e atualização do Sheets.
 */
function sincronizarBaseDosDrive() {
  validarContextoInterno();
  const inicio = new Date();
  Logger.log(`[SYNC] Iniciando sincronização em ${inicio.toLocaleString('pt-BR')}`);

  try {
    // 1. Obter o arquivo CSV do Drive
    const arquivo = DriveApp.getFileById(SYNC_CONFIG.CSV_FILE_ID);
    const conteudoCsv = arquivo.getBlob().getDataAsString(SYNC_CONFIG.CSV_ENCODING);

    // 2. Parse do CSV (Utilities.parseCsv lida corretamente com valores entre aspas)
    const linhas = Utilities.parseCsv(conteudoCsv, SYNC_CONFIG.CSV_SEPARATOR.charCodeAt(0));
    if (!linhas || linhas.length < 2) {
      throw new Error('Arquivo CSV está vazio ou sem dados após o cabeçalho.');
    }

    const headers = linhas[0].map(h => String(h).trim());
    Logger.log(`[SYNC] ${linhas.length - 1} registros encontrados no CSV.`);

    // 3. Mapear índices das colunas configuradas
    const idx = {};
    for (const [campo, nomeColuna] of Object.entries(SYNC_CONFIG.COLUNAS)) {
      idx[campo] = headers.findIndex(h => h.toLowerCase() === nomeColuna.toLowerCase());
      if (idx[campo] === -1) {
        Logger.log(`[SYNC][AVISO] Coluna "${nomeColuna}" (campo: ${campo}) não encontrada no CSV.`);
      }
    }

    // 4. Montar array de Associados
    // Cabeçalho compatível com a lógica de leitura dinâmica do buscarAssociado()
    const headersAssoc = [
      'Nome Coop', 'Associado', 'Conta', 'Score', 'Tipo Pessoa', 'Idade',
      'MEI', 'Simples', 'CPF/CNPJ', 'Renda Mensal', 'Data Atualizacao Renda', 'Data Consulta SCR',
      'Colaborador', 'Endividamento Mensal', 'Endiv. Total',
      // Acrescentadas no fim (27/10/2026). Os nomes evitam "CPF", "MEI", "IDADE" etc.,
      // que as buscas usam para achar outras colunas pelo cabeçalho.
      'Estado Civil', 'União Estável', 'Documento Cônjuge'
    ];
    const dadosAssoc = [headersAssoc];

    for (let i = 1; i < linhas.length; i++) {
      const row = linhas[i];

      const conta = idx.conta !== undefined && idx.conta !== -1
        ? String(row[idx.conta] || '').trim()
        : '';

      const contaFinal = conta || 'Sem conta';

      const nome = idx.nome !== -1 ? String(row[idx.nome] || '').trim() : '';
      const nascim = idx.nascimento !== -1 ? String(row[idx.nascimento] || '').trim() : '';
      const pessoa = idx.pessoa !== -1 ? _normalizarPessoa(row[idx.pessoa]) : 'PF';
      const mei = idx.mei !== -1 ? _normalizarBoolean(row[idx.mei]) : 'NÃO';
      const simples = idx.simples !== -1 ? _normalizarBoolean(row[idx.simples]) : 'NÃO';
      const idade = _calcularIdade(nascim);
      const cpfCnpj = idx.cpfCnpj !== -1 ? String(row[idx.cpfCnpj] || '').trim() : '';
      const renda = idx.renda !== -1 ? String(row[idx.renda] || '').trim() : '';
      const dataRenda = idx.dataRenda !== -1 ? String(row[idx.dataRenda] || '').trim() : '';
      const dataSCR = idx.dataSCR !== -1 ? String(row[idx.dataSCR] || '').trim() : '';
      const colab = idx.colaborador !== -1 ? _normalizarBoolean(row[idx.colaborador]) : 'NÃO';
      const endivMensal = idx.endivMensal !== -1 ? _parseCurrencyRobust(row[idx.endivMensal]) : 0;
      const endivTot = idx.endivTotal !== -1 ? String(row[idx.endivTotal] || '').trim() : '';
      const estadoCivil = idx.estadoCivil !== -1 ? String(row[idx.estadoCivil] || '').trim() : '';
      const uniaoEstavel = idx.uniaoEstavel !== -1 ? _normalizarBoolean(row[idx.uniaoEstavel]) : '';
      const cpfConjuge = idx.cpfConjuge !== -1 ? String(row[idx.cpfConjuge] || '').trim() : '';

      dadosAssoc.push([
        '',          // Nome Coop (não disponível no CSV)
        nome,        // Associado
        contaFinal,  // Conta (ou 'Sem conta')
        '',          // Score — vazio (operador informa manualmente)
        pessoa,      // Tipo Pessoa (PF/PJ)
        idade,       // Idade calculada
        mei,         // MEI (SIM/NÃO)
        simples,     // Simples Nacional (SIM/NÃO)
        cpfCnpj,     // CPF/CNPJ
        renda,       // Renda Mensal
        dataRenda,   // Data Atualização Renda
        dataSCR,     // Data Consulta SCR
        colab,       // Colaborador (SIM/NÃO)
        endivMensal, // Endividamento Mensal
        endivTot,    // Endividamento Total
        estadoCivil, // Estado Civil
        uniaoEstavel, // União Estável (SIM/NÃO)
        cpfConjuge,  // Documento Cônjuge
      ]);
    }

    // 5. Gravar Associados
    // Sem sessão de usuário (contexto de trigger), validarContextoInterno()
    // dentro de substituirAssociados libera a gravação automaticamente.
    const resultAssoc = substituirAssociados(dadosAssoc);
    Logger.log(`[SYNC] Associados: ${JSON.stringify(resultAssoc)}`);

    if (!resultAssoc.sucesso) {
      throw new Error(resultAssoc.erro);
    }

    // 6. Registrar resultado (MEI/Simples já gravados na aba Associados acima)
    const duracao = ((new Date() - inicio) / 1000).toFixed(1);
    const resumo = `Sync automático concluído: ${dadosAssoc.length - 1} associados atualizados. Duração: ${duracao}s.`;
    Logger.log(`[SYNC] ✅ ${resumo}`);

    // Gravar timestamp e status do último sync nas configurações
    _setConfigInterna('SYNC_ULTIMO_STATUS', 'SUCESSO');
    _setConfigInterna('SYNC_ULTIMO_RESULTADO', resumo);

    // Após concluir Associados com sucesso, chama o sync de Composição Societária
    try {
      Logger.log(`[SYNC] Iniciando sincronização em cadeia da Composição Societária...`);
      sincronizarBaseComposicaoSocietaria();
    } catch (eComposicao) {
      Logger.log(`[SYNC][AVISO] Falha ao sincronizar Composição Societária: ${eComposicao.message}`);
    }

    // Após concluir, chama o sync de Grupo Econômico
    try {
      Logger.log(`[SYNC] Iniciando sincronização em cadeia de Grupo Econômico...`);
      sincronizarGrupoEconomico();
    } catch (eGrupo) {
      Logger.log(`[SYNC][AVISO] Falha ao sincronizar Grupo Econômico: ${eGrupo.message}`);
    }

    return { sucesso: true, resumo };

  } catch (e) {
    const mensagemErro = `FALHA na sincronização: ${e.message}`;
    Logger.log(`[SYNC] ❌ ${mensagemErro}`);

    _setConfigInterna('SYNC_ULTIMO_STATUS', 'FALHA');
    _setConfigInterna('SYNC_ULTIMO_RESULTADO', mensagemErro);

    // Notificação por e-mail para todos os ADMINs ativos
    try {
      // Destinatários: admins do Hub (= ADMINs do Simulador desde 15/10/2026).
      const emailsAdmins = _emailsAdminsHub_();
      if (emailsAdmins.length) {
        const emails = emailsAdmins.join(',');
        if (emails) {
          MailApp.sendEmail(
            emails,
            '[Simulador Cresol] Falha no Sync Automático (Associados)',
            mensagemErro
          );
        }
      }
    } catch (errMail) {
      Logger.log(`[SYNC] Erro ao enviar e-mail de falha: ${errMail.message}`);
    }

    return { sucesso: false, erro: e.message };
  }
}

/**
 * SUBSTITUIÇÃO SEGURA DA ABA COMPOSIÇÃO SOCIETÁRIA
 * Usa estratégia de swap atômico com aba temporária + gravação em lotes:
 * 1. Grava em '_ComposicaoSocietaria_temp' (aba real intacta durante gravação)
 * 2. Valida contagem de linhas gravadas
 * 3. SÓ ENTÃO deleta aba real e renomeia a temporária
 * Isso evita perda de dados se um timeout ocorrer durante a gravação.
 */
function substituirComposicaoSocietaria(dadosFiltrados) {
  try {
    validarContextoInterno();
    if (!dadosFiltrados || dadosFiltrados.length === 0) return { sucesso: false, erro: "Sem dados para gravar." };

    const TAMANHO_LOTE = 5000;
    const totalLinhas = dadosFiltrados.length;
    const neededCols = dadosFiltrados[0].length;
    const NOME_TEMP = '_ComposicaoSocietaria_temp';

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Limpar aba temporária de execuções anteriores que possam ter falhado
    const sheetTemporaria = ss.getSheetByName(NOME_TEMP);
    if (sheetTemporaria) {
      ss.deleteSheet(sheetTemporaria);
    }

    // Determinar a posição da aba real
    let sheetReal = ss.getSheetByName('ComposicaoSocietaria');
    const posReal = sheetReal ? sheetReal.getIndex() - 1 : 1;

    // Criar a aba temporária em posição adjacente
    let sheetTemp = ss.insertSheet(NOME_TEMP, posReal + 1);

    // Redimensionar colunas excedentes
    const maxCols = sheetTemp.getMaxColumns();
    if (maxCols > neededCols) {
      sheetTemp.deleteColumns(neededCols + 1, maxCols - neededCols);
    }

    // Alocar todas as linhas necessárias de uma só vez
    const maxRows = sheetTemp.getMaxRows();
    if (totalLinhas > maxRows) {
      sheetTemp.insertRowsAfter(maxRows, totalLinhas - maxRows);
    }
    SpreadsheetApp.flush();

    // Gravação em lotes na aba TEMPORÁRIA (base real intacta durante todo esse processo)
    let linhaInicio = 1;
    while (linhaInicio <= totalLinhas) {
      const lote = dadosFiltrados.slice(linhaInicio - 1, linhaInicio - 1 + TAMANHO_LOTE);
      sheetTemp.getRange(linhaInicio, 1, lote.length, neededCols).setValues(lote);
      SpreadsheetApp.flush();
      linhaInicio += TAMANHO_LOTE;
    }

    // Validação: confirmar contagem de linhas antes do swap
    const linhasGravadas = sheetTemp.getLastRow();
    if (linhasGravadas < totalLinhas) {
      ss.deleteSheet(sheetTemp);
      throw new Error(
        `Gravação incompleta na aba temporária: esperado ${totalLinhas} linhas, encontrado ${linhasGravadas}. ` +
        `A base 'ComposicaoSocietaria' original foi preservada.`
      );
    }

    // ── SWAP ATÔMICO ────────────────────────────────────────────────────────
    if (sheetReal) {
      ss.deleteSheet(sheetReal);
    }
    sheetTemp.setName('ComposicaoSocietaria');
    SpreadsheetApp.flush();
    // ────────────────────────────────────────────────────────────────────────

    return { sucesso: true, linhas: totalLinhas - 1 };
  } catch (e) {
    return { sucesso: false, erro: "Erro no servidor ao gravar composição societária: " + e.message };
  }
}


/**
 * Parser resiliente para valores monetários vindos do CSV
 */
function _parseCurrencyRobust(val) {
  if (!val) return 0;
  let str = String(val).trim();
  if (str.includes('.') && str.includes(',')) {
    str = str.replace(/\./g, '').replace(',', '.');
  } else if (str.includes(',')) {
    str = str.replace(',', '.');
  }
  let floatVal = parseFloat(str);
  return isNaN(floatVal) ? 0 : floatVal;
}

/**
 * ROTINA DE SINCRONIZAÇÃO AUTOMÁTICA DE COMPOSIÇÃO SOCIETÁRIA
 */
function sincronizarBaseComposicaoSocietaria() {
  validarContextoInterno();
  const inicio = new Date();
  Logger.log(`[SYNC-COMPOSICAO] Iniciando sincronização em ${inicio.toLocaleString('pt-BR')}`);

  try {
    // 1. Obter o arquivo CSV do Drive
    const arquivo = DriveApp.getFileById(SYNC_COMPOSICAO_CONFIG.CSV_FILE_ID);
    const conteudoCsv = arquivo.getBlob().getDataAsString(SYNC_COMPOSICAO_CONFIG.CSV_ENCODING);

    // 2. Parse do CSV
    const linhas = Utilities.parseCsv(conteudoCsv, SYNC_COMPOSICAO_CONFIG.CSV_SEPARATOR.charCodeAt(0));
    if (!linhas || linhas.length < 2) {
      throw new Error('Arquivo CSV está vazio ou sem dados após o cabeçalho.');
    }

    const headers = linhas[0].map(h => String(h).trim());
    Logger.log(`[SYNC-COMPOSICAO] ${linhas.length - 1} registros encontrados no CSV.`);

    // 3. Mapear índices
    const idxCnpj = headers.findIndex(h => h.toLowerCase() === SYNC_COMPOSICAO_CONFIG.COLUNAS.cnpj.toLowerCase());
    const idxCpfSocio = headers.findIndex(h => h.toLowerCase() === SYNC_COMPOSICAO_CONFIG.COLUNAS.cpfSocio.toLowerCase());
    const idxPercentual = headers.findIndex(h => h.toLowerCase() === SYNC_COMPOSICAO_CONFIG.COLUNAS.percentual.toLowerCase());

    if (idxCnpj === -1 || idxCpfSocio === -1 || idxPercentual === -1) {
      throw new Error('Colunas obrigatórias não encontradas no CSV de Composição Societária.');
    }

    // 4. Montar array
    const headersFinais = ['nr_cnpj', 'nr_cpf_composicao_societaria', 'nr_percentual_participacao'];
    const dados = [headersFinais];

    // Participação gravada como NÚMERO (27/10/2026). Antes ia como texto e a planilha
    // o reinterpretava pela localidade: um sócio com 100 no CSV ficou com 1 na base,
    // e o seguro prestamista (rateado pela participação) saiu 100 vezes menor.
    for (let i = 1; i < linhas.length; i++) {
      const row = linhas[i];
      const percentualTxt = String(row[idxPercentual] || '').trim();
      dados.push([
        String(row[idxCnpj] || '').trim(),
        String(row[idxCpfSocio] || '').trim(),
        percentualTxt === '' ? '' : _parseCurrencyRobust(percentualTxt)
      ]);
    }

    // 5. Gravar
    const result = substituirComposicaoSocietaria(dados);
    Logger.log(`[SYNC-COMPOSICAO] Gravado: ${JSON.stringify(result)}`);

    const duracao = ((new Date() - inicio) / 1000).toFixed(1);
    const resumo = `Sync Composição Societária concluído: ${dados.length - 1} sócios atualizados. Duração: ${duracao}s.`;
    Logger.log(`[SYNC-COMPOSICAO] ✅ ${resumo}`);

    return { sucesso: true, resumo };
  } catch (e) {
    Logger.log(`[SYNC-COMPOSICAO] ❌ FALHA: ${e.message}`);

    // Notificação por e-mail para todos os ADMINs ativos
    try {
      // Destinatários: admins do Hub (= ADMINs do Simulador desde 15/10/2026).
      const emailsAdmins = _emailsAdminsHub_();
      if (emailsAdmins.length) {
        const emails = emailsAdmins.join(',');
        if (emails) {
          MailApp.sendEmail(
            emails,
            '[Simulador Cresol] Falha no Sync Automático (Composição Societária PJ)',
            `FALHA na sincronização de Composição Societária: ${e.message}`
          );
        }
      }
    } catch (errMail) {
      Logger.log(`[SYNC-COMPOSICAO] Erro ao enviar e-mail de falha: ${errMail.message}`);
    }

    return { sucesso: false, erro: e.message };
  }
}

/**
 * FUNÇÃO AUXILIAR — Instala o Trigger de tempo no projeto.
 * Lê o horário da configuração SYNC_HORA (padrão: 5).
 * Execute MANUALMENTE UMA ÚNICA VEZ pelo editor do Apps Script,
 * ou via botão "Salvar e Aplicar" no painel Admin do simulador.
 */
function instalarTriggerSyncAutomatico() {
  validarAcessoAdmin();
  // Lê o horário configurado no Sheets (ou usa o padrão 5)
  let hora = 5;
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetConf = ss.getSheetByName('Configuracoes');
    if (sheetConf) {
      const dataConf = sheetConf.getDataRange().getValues();
      for (let i = 0; i < dataConf.length; i++) {
        if (dataConf[i][0] === 'SYNC_HORA') {
          const h = parseInt(dataConf[i][1]);
          if (!isNaN(h) && h >= 0 && h <= 23) hora = h;
          break;
        }
      }
    }
  } catch (e) {
    Logger.log('[TRIGGER] Não foi possível ler SYNC_HORA, usando padrão 05h. Erro: ' + e.message);
  }

  // Remove triggers existentes da mesma função (evita duplicatas)
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'sincronizarBaseDosDrive') {
      ScriptApp.deleteTrigger(t);
      Logger.log('[TRIGGER] Trigger anterior removido.');
    }
  });

  // Cria o novo trigger diário
  ScriptApp.newTrigger('sincronizarBaseDosDrive')
    .timeBased()
    .everyDays(1)
    .atHour(hora)
    .create();

  _instalarTriggerLimites_(hora);

  const msg = `Trigger instalado: sincronizarBaseDosDrive — diariamente às ${String(hora).padStart(2, '0')}h.`;
  Logger.log(`[TRIGGER] ✅ ${msg}`);
  setConfigGlobal('SYNC_HORA', hora);
  setConfigGlobal('SYNC_TRIGGER_ATIVO', 'SIM');
  return { sucesso: true, mensagem: msg };
}

/**
 * FUNÇÃO AUXILIAR — Remove o Trigger de sync (para desativar a automação).
 */
function removerTriggerSyncAutomatico() {
  validarAcessoAdmin();
  const triggers = ScriptApp.getProjectTriggers();
  let removidos = 0;
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'sincronizarBaseDosDrive' || t.getHandlerFunction() === LIMITES_HANDLER) {
      ScriptApp.deleteTrigger(t);
      removidos++;
    }
  });
  setConfigGlobal('SYNC_TRIGGER_ATIVO', 'NÃO');
  Logger.log(`[TRIGGER] ${removidos} trigger(s) de sync removido(s).`);
  return { sucesso: true, removidos };
}

/**
 * FUNÇÃO ADMIN — Salva o horário do sync e reinstala o trigger.
 * Chamada pelo painel Admin do simulador via botão "Salvar e Aplicar".
 * @param {number} hora - Hora desejada (0–23).
 */
function salvarHorarioSync(hora) {
  try {
    validarAcessoAdmin();
    const h = parseInt(hora);
    if (isNaN(h) || h < 0 || h > 23) {
      return { sucesso: false, erro: 'Horário inválido. Informe um valor entre 0 e 23.' };
    }
    setConfigGlobal('SYNC_HORA', h);
    return instalarTriggerSyncAutomatico();
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

/**
 * FUNÇÃO ADMIN — Retorna o status atual do sync para exibir no painel Admin.
 */
function buscarStatusSync() {
  try {
    validarAcessoAdmin();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetConf = ss.getSheetByName('Configuracoes');
    let hora = 5, triggerAtivo = 'NÃO', ultimoStatus = '—', ultimoResultado = '—';

    if (sheetConf) {
      const dataConf = sheetConf.getDataRange().getValues();
      dataConf.forEach(row => {
        if (row[0] === 'SYNC_HORA') hora = row[1];
        if (row[0] === 'SYNC_TRIGGER_ATIVO') triggerAtivo = row[1];
        if (row[0] === 'SYNC_ULTIMO_STATUS') ultimoStatus = row[1];
        if (row[0] === 'SYNC_ULTIMO_RESULTADO') ultimoResultado = row[1];
      });
    }

    // Verificar se o trigger realmente existe no Apps Script
    const triggers = ScriptApp.getProjectTriggers();
    const triggerExiste = triggers.some(t => t.getHandlerFunction() === 'sincronizarBaseDosDrive');
    if (!triggerExiste && triggerAtivo === 'SIM') {
      triggerAtivo = 'NÃO';
      setConfigGlobal('SYNC_TRIGGER_ATIVO', 'NÃO');
    }

    return {
      sucesso: true,
      hora: hora,
      triggerAtivo: triggerAtivo,
      ultimoStatus: ultimoStatus,
      ultimoResultado: ultimoResultado
    };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

/**
 * FUNÇÃO ADMIN — Executa a sincronização manualmente (chamada via google.script.run).
 */
function executarSincronizacaoManual() {
  validarAcessoAdmin();
  return sincronizarBaseDosDrive();
}

// ==========================================
// LIMITES PRÉ-APROVADOS E LIMITE INTELIGENTE (15/10/2026)
// ==========================================
// Todo mês chega numa pasta do Drive uma planilha .xlsx com os limites por
// conta (nome do tipo "2032 - 202610 Pré-Aprovado e Limite Inteligente.xlsx").
// Vale sempre a MAIS RECENTE — pelo AAAAMM no nome; sem ele, pela data de
// modificação. Dela usamos 3 colunas (1ª aba, cabeçalho na linha 1):
//   - conta;
//   - Valor Pré-Aprovado (contratável pelo associado no app);
//   - Valor Limite Inteligente (aprovado por renda, para todas as linhas; até 11/2026 "Vlr ...").
// Limite 0 = não tem. Contas sem nenhum dos dois não são gravadas.
//
// O .xlsx é convertido numa planilha temporária pelo serviço avançado
// "Drive API" (precisa estar ativado no projeto) e lido de uma vez; o
// resultado vai para a aba LimitesPreAprovados com o mesmo swap atômico das
// outras cargas. A busca de associado consulta essa aba pela conta (busca
// nativa, ~0,3 s) — nunca o .xlsx.
//
// O trigger diário só recarrega se o arquivo mudou (outro ID ou nova data de
// modificação). "Atualizar agora" no Admin força a recarga.

const LIMITES_ABA = 'LimitesPreAprovados';
const LIMITES_HANDLER = 'sincronizarLimitesPreAprovados';
const LIMITES_PADRAO = {
  PASTA_ID: '1MKhAPiCkPLPGnvpm5Q0Q7MaB1BYafiG4',
  COL_CONTA: 'Nr. Conta Corrente',
  COL_PRE: 'Valor Pré-Aprovado',
  COL_INTELIGENTE: 'Valor Limite Inteligente' // até 11/2026 "Vlr Limite Inteligente" (as duas formas são aceitas)
};
const LIMITES_CHAVES = ['LIMITES_PASTA_ID', 'LIMITES_COL_CONTA', 'LIMITES_COL_PRE', 'LIMITES_COL_INTELIGENTE',
  'LIMITES_ARQUIVO_NOME', 'LIMITES_ARQUIVO_ID', 'LIMITES_ARQUIVO_ATUALIZADO', 'LIMITES_MES_REF', 'LIMITES_QTD',
  'LIMITES_ULTIMA_CARGA', 'LIMITES_ULTIMO_STATUS', 'LIMITES_ULTIMO_RESULTADO'];
const LIMITES_MESES = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];

/** Configuração em uso: o que está em Configuracoes, senão o padrão. */
function _configLimites_() {
  const c = _lerConfiguracoes_(LIMITES_CHAVES);
  const txt = (k, padrao) => String(c[k] == null || c[k] === '' ? (padrao || '') : c[k]).trim();
  return {
    pastaId: txt('LIMITES_PASTA_ID', LIMITES_PADRAO.PASTA_ID),
    colConta: txt('LIMITES_COL_CONTA', LIMITES_PADRAO.COL_CONTA),
    colPre: txt('LIMITES_COL_PRE', LIMITES_PADRAO.COL_PRE),
    colInteligente: txt('LIMITES_COL_INTELIGENTE', LIMITES_PADRAO.COL_INTELIGENTE),
    arquivoNome: txt('LIMITES_ARQUIVO_NOME'),
    arquivoId: txt('LIMITES_ARQUIVO_ID'),
    arquivoAtualizado: txt('LIMITES_ARQUIVO_ATUALIZADO'),
    mesRef: txt('LIMITES_MES_REF'),
    qtd: txt('LIMITES_QTD'),
    ultimaCarga: txt('LIMITES_ULTIMA_CARGA'),
    ultimoStatus: txt('LIMITES_ULTIMO_STATUS'),
    ultimoResultado: txt('LIMITES_ULTIMO_RESULTADO')
  };
}

/** Aceita o ID ou o link da pasta/arquivo do Drive. */
function _extrairIdDrive_(texto) {
  const t = String(texto || '').trim();
  const m = t.match(/\/folders\/([\w-]{10,})/) || t.match(/\/d\/([\w-]{10,})/) || t.match(/[?&]id=([\w-]{10,})/);
  return m ? m[1] : t;
}

/** Nome de coluna comparável: sem acento, minúsculo, espaços simples. */
function _normalizarColuna_(texto) {
  return String(texto || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}

/**
 * Nome de coluna dos limites para comparação: além de _normalizarColuna_, iguala as
 * abreviações (vlr/vl → valor; nr/nro/num → numero), troca pontuação por espaço e tira
 * "de/do/da". Assim "Vlr Limite Inteligente" = "Valor Limite Inteligente" e
 * "Nr. Conta Corrente" = "Número da Conta Corrente".
 */
function _chaveColunaLimites_(texto) {
  return _normalizarColuna_(texto)
    .replace(/[^a-z0-9]+/g, ' ')
    .split(' ')
    .filter(p => p && ['de', 'do', 'da', 'dos', 'das'].indexOf(p) === -1)
    .map(p => ({ vlr: 'valor', vl: 'valor', nr: 'numero', nro: 'numero', num: 'numero', no: 'numero' })[p] || p)
    .join(' ');
}

/** Valor do limite: número da célula, ou texto como "R$ 1.234,56". Inválido = 0. */
function _numeroLimite_(valor) {
  if (typeof valor === 'number') return isFinite(valor) ? valor : 0;
  let t = String(valor == null ? '' : valor).replace(/[^\d,.\-]/g, '');
  if (!t) return 0;
  if (t.indexOf(',') !== -1) t = t.replace(/\./g, '').replace(',', '.');
  else if ((t.match(/\./g) || []).length > 1) t = t.replace(/\./g, '');
  const n = parseFloat(t);
  return isNaN(n) ? 0 : n;
}

/** Conta comparável: só os dígitos, sem zeros à esquerda ("012.345-6" = "123456"). */
function _chaveConta_(conta) {
  return _apenasDigitos_(conta).replace(/^0+/, '');
}

/** Regex (busca nativa) para a célula cuja conta tem exatamente estes dígitos. */
function _regexConta_(chave) {
  return '^\\D*0*' + String(chave).split('').join('\\D*') + '\\D*$';
}

/** AAAAMM do nome do arquivo → { chave: 202610, texto: "out/2026" }; null se não houver. */
function _mesDoNomeLimites_(nome) {
  const m = String(nome || '').match(/(?:^|\D)(20\d{2})(0[1-9]|1[0-2])(?:\D|$)/);
  if (!m) return null;
  return { chave: Number(m[1] + m[2]), texto: LIMITES_MESES[Number(m[2]) - 1] + '/' + m[1] };
}

/** Planilhas (.xlsx ou Google Sheets) da pasta, a mais recente primeiro. */
function _arquivosLimites_(pastaId) {
  const pasta = DriveApp.getFolderById(pastaId);
  const arquivos = [];
  const it = pasta.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const tipo = f.getMimeType();
    const nome = f.getName();
    const ehXlsx = tipo === MimeType.MICROSOFT_EXCEL || /\.xlsx$/i.test(nome);
    if (!ehXlsx && tipo !== MimeType.GOOGLE_SHEETS) continue;
    const mes = _mesDoNomeLimites_(nome);
    arquivos.push({ arquivo: f, nome, id: f.getId(), ehXlsx, mes,
      atualizado: f.getLastUpdated().getTime() });
  }
  arquivos.sort((a, b) => {
    const ma = a.mes ? a.mes.chave : 0, mb = b.mes ? b.mes.chave : 0;
    if (ma !== mb) return mb - ma;
    return b.atualizado - a.atualizado;
  });
  return arquivos;
}

/** Converte o .xlsx numa planilha Google temporária (serviço avançado Drive, v3 ou v2). Devolve o ID. */
function _converterXlsxTemporario_(arquivoId, nome) {
  if (typeof Drive === 'undefined' || !Drive.Files) {
    throw new Error('Ative o serviço avançado "Drive API" no projeto do Apps Script (Editor → Serviços → + → Drive API) para ler arquivos .xlsx.');
  }
  const titulo = '_tmp_limites_' + nome;
  if (typeof Drive.Files.create === 'function') { // v3
    return Drive.Files.copy({ name: titulo, mimeType: MimeType.GOOGLE_SHEETS, parents: ['root'] },
      arquivoId, { supportsAllDrives: true }).id;
  }
  return Drive.Files.copy({ title: titulo, mimeType: MimeType.GOOGLE_SHEETS, parents: [{ id: 'root' }] }, // v2
    arquivoId, { convert: true, supportsAllDrives: true }).id;
}

/** Lê a 1ª aba do arquivo e devolve as linhas [conta, pré, inteligente] com algum limite > 0. */
function _lerArquivoLimites_(item, cfg) {
  let idTemp = null;
  try {
    let idLeitura = item.id;
    if (item.ehXlsx) idLeitura = idTemp = _converterXlsxTemporario_(item.id, item.nome);
    const aba = SpreadsheetApp.openById(idLeitura).getSheets()[0];
    const valores = aba.getDataRange().getValues();
    if (valores.length < 2) throw new Error('A 1ª aba do arquivo está vazia.');

    // Nomes comparados por _chaveColunaLimites_ ("Vlr" = "Valor", "Nr" = "Número"...). Se o
    // nome configurado não existir, aceita a ÚNICA coluna que tenha as palavras-chave: em
    // 11/2026 a Central trocou "Vlr Limite Inteligente" por "Valor Limite Inteligente" e a
    // carga parou. O que foi achado assim vai para o resultado da carga.
    const cab = valores[0].map(_chaveColunaLimites_);
    const ajustes = [];
    const achar = (nomeCol, palavras) => {
      const exato = cab.indexOf(_chaveColunaLimites_(nomeCol));
      if (exato !== -1) return exato;
      const candidatos = [];
      cab.forEach((c, i) => { if (palavras.every(p => c.indexOf(p) !== -1)) candidatos.push(i); });
      if (candidatos.length !== 1) return -1;
      ajustes.push('"' + nomeCol + '" lida da coluna "' + valores[0][candidatos[0]] + '"');
      return candidatos[0];
    };
    const iConta = achar(cfg.colConta, ['conta']);
    const iPre = achar(cfg.colPre, ['pre', 'aprovado']);
    const iInt = achar(cfg.colInteligente, ['limite', 'inteligente']);
    const faltando = [];
    if (iConta === -1) faltando.push(cfg.colConta);
    if (iPre === -1) faltando.push(cfg.colPre);
    if (iInt === -1) faltando.push(cfg.colInteligente);
    if (faltando.length) {
      throw new Error('Coluna(s) não encontrada(s) na 1ª aba: ' + faltando.join(', ') +
        '. Colunas do arquivo: ' + valores[0].filter(String).join(', '));
    }

    const linhas = [];
    let semConta = 0;
    for (let i = 1; i < valores.length; i++) {
      const r = valores[i];
      const conta = typeof r[iConta] === 'number' ? String(Math.round(r[iConta])) : String(r[iConta] || '').trim();
      const pre = Math.round(_numeroLimite_(r[iPre]) * 100) / 100;
      const intel = Math.round(_numeroLimite_(r[iInt]) * 100) / 100;
      if (!(pre > 0) && !(intel > 0)) continue;
      if (!_chaveConta_(conta)) { semConta++; continue; }
      linhas.push([conta, pre > 0 ? pre : 0, intel > 0 ? intel : 0]);
    }
    if (ajustes.length) Logger.log('[LIMITES] Coluna(s) por nome aproximado: ' + ajustes.join('; '));
    return { linhas, totalArquivo: valores.length - 1, semConta, ajustes };
  } finally {
    if (idTemp) {
      try { DriveApp.getFileById(idTemp).setTrashed(true); }
      catch (e) { Logger.log('[LIMITES] Não consegui mover a planilha temporária para a lixeira: ' + e.message); }
    }
  }
}

/** Grava a aba de limites com swap atômico (base anterior intacta se algo falhar). */
function _gravarAbaLimites_(linhas, mesRef) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const NOME_TEMP = '_' + LIMITES_ABA + '_temp';
  const dados = [['Conta', 'Valor Pré-Aprovado', 'Vlr Limite Inteligente', 'Mês Referência']]
    .concat(linhas.map(l => [l[0], l[1], l[2], mesRef]));
  const total = dados.length, numCols = 4;

  const antiga = ss.getSheetByName(NOME_TEMP);
  if (antiga) ss.deleteSheet(antiga);
  const real = ss.getSheetByName(LIMITES_ABA);
  const temp = ss.insertSheet(NOME_TEMP, real ? real.getIndex() : ss.getNumSheets());

  const maxCols = temp.getMaxColumns();
  if (maxCols > numCols) temp.deleteColumns(numCols + 1, maxCols - numCols);
  const maxRows = temp.getMaxRows();
  if (total > maxRows) temp.insertRowsAfter(maxRows, total - maxRows);
  // Conta e mês como TEXTO: o Sheets transformaria "0123" em 123 e "out/2026" em data.
  temp.getRange(1, 1, total, 1).setNumberFormat('@');
  temp.getRange(1, 4, total, 1).setNumberFormat('@');
  SpreadsheetApp.flush();

  const LOTE = 5000;
  for (let ini = 0; ini < total; ini += LOTE) {
    const lote = dados.slice(ini, ini + LOTE);
    temp.getRange(ini + 1, 1, lote.length, numCols).setValues(lote);
    SpreadsheetApp.flush();
  }
  const gravadas = temp.getLastRow();
  if (gravadas < total) {
    ss.deleteSheet(temp);
    throw new Error(`Gravação incompleta (esperado ${total} linhas, gravado ${gravadas}). Os limites anteriores foram preservados.`);
  }
  if (real) ss.deleteSheet(real);
  temp.setName(LIMITES_ABA);
  SpreadsheetApp.flush();
  return total - 1;
}

/**
 * Carrega os limites do arquivo mais recente da pasta.
 * Trigger diário (recebe o objeto do evento): só recarrega se o arquivo mudou.
 * `forcar === true` (botão "Atualizar agora"): recarrega sempre.
 */
function sincronizarLimitesPreAprovados(forcar) {
  validarContextoInterno();
  const forcado = forcar === true;
  const inicio = new Date();
  const agora = () => Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  try {
    const cfg = _configLimites_();
    if (!cfg.pastaId) throw new Error('Pasta dos limites não configurada.');
    const arquivos = _arquivosLimites_(cfg.pastaId);
    if (!arquivos.length) throw new Error('Nenhuma planilha (.xlsx) encontrada na pasta dos limites.');
    const item = arquivos[0];

    if (!forcado && item.id === cfg.arquivoId && String(item.atualizado) === cfg.arquivoAtualizado) {
      Logger.log(`[LIMITES] Sem arquivo novo (${item.nome}). Nada a fazer.`);
      return { sucesso: true, semMudanca: true, arquivo: item.nome };
    }

    const lido = _lerArquivoLimites_(item, cfg);
    if (!lido.linhas.length) throw new Error('O arquivo "' + item.nome + '" não tem nenhuma conta com limite maior que zero.');
    const mesRef = item.mes ? item.mes.texto
      : Utilities.formatDate(new Date(item.atualizado), Session.getScriptTimeZone(), 'MM/yyyy');
    const qtd = _gravarAbaLimites_(lido.linhas, mesRef);

    const duracao = ((new Date() - inicio) / 1000).toFixed(1);
    const resumo = `${qtd} contas com limite (de ${lido.totalArquivo} linhas do arquivo` +
      (lido.semConta ? `; ${lido.semConta} sem conta ignoradas` : '') + `). ${duracao}s.` +
      (lido.ajustes && lido.ajustes.length ? ' Atenção — coluna com nome diferente do configurado: ' + lido.ajustes.join('; ') + '. Ajuste o nome no Admin.' : '');
    _setConfigInterna('LIMITES_ARQUIVO_NOME', item.nome);
    _setConfigInterna('LIMITES_ARQUIVO_ID', item.id);
    _setConfigInterna('LIMITES_ARQUIVO_ATUALIZADO', String(item.atualizado));
    _setConfigInterna('LIMITES_MES_REF', mesRef);
    _setConfigInterna('LIMITES_QTD', String(qtd));
    _setConfigInterna('LIMITES_ULTIMA_CARGA', agora());
    _setConfigInterna('LIMITES_ULTIMO_STATUS', 'SUCESSO');
    _setConfigInterna('LIMITES_ULTIMO_RESULTADO', resumo);
    Logger.log(`[LIMITES] ✅ ${item.nome}: ${resumo}`);
    return { sucesso: true, arquivo: item.nome, mesRef, qtd, resumo };
  } catch (e) {
    Logger.log(`[LIMITES] ❌ FALHA: ${e.message}`);
    try {
      _setConfigInterna('LIMITES_ULTIMO_STATUS', 'FALHA');
      _setConfigInterna('LIMITES_ULTIMO_RESULTADO', agora() + ' — ' + e.message);
    } catch (e2) { /* sem onde registrar */ }
    if (!forcado) { // no trigger ninguém vê a tela: avisa os admins
      try {
        const emails = _emailsAdminsHub_();
        if (emails.length) MailApp.sendEmail(emails.join(','), '[Simulador Cresol] Falha na carga dos Limites Pré-aprovados',
          'FALHA na carga dos limites pré-aprovados: ' + e.message);
      } catch (errMail) { Logger.log('[LIMITES] Erro ao enviar e-mail de falha: ' + errMail.message); }
    }
    return { sucesso: false, erro: e.message };
  }
}

/** Trigger diário dos limites, 1 h depois do sync da base. Recria se já existir. */
function _instalarTriggerLimites_(horaSync) {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === LIMITES_HANDLER) ScriptApp.deleteTrigger(t);
  });
  const h = parseInt(horaSync, 10);
  const hora = isNaN(h) ? 6 : (h + 1) % 24;
  ScriptApp.newTrigger(LIMITES_HANDLER).timeBased().everyDays(1).atHour(hora).create();
  return hora;
}

/**
 * Limites de uma conta, ou null. Busca nativa na coluna A da aba de limites
 * (negativa aceita: a conta está por extenso na célula). Nunca lança: falha
 * aqui não pode impedir a busca do associado.
 */
function _limitesDaConta_(conta, ss) {
  try {
    const chave = _chaveConta_(conta);
    if (!chave) return null;
    const sheet = (ss || SpreadsheetApp.openById(SPREADSHEET_ID)).getSheetByName(LIMITES_ABA);
    if (!sheet || sheet.getLastRow() < 2) return null;
    const cand = _candidatas_(sheet, 1, _regexConta_(chave), 5);
    for (const linha of (cand || [])) {
      const r = sheet.getRange(linha, 1, 1, 4).getValues()[0];
      if (_chaveConta_(r[0]) !== chave) continue;
      const pre = _numeroLimite_(r[1]), intel = _numeroLimite_(r[2]);
      if (!(pre > 0) && !(intel > 0)) return null;
      return { preAprovado: pre > 0 ? pre : 0, limiteInteligente: intel > 0 ? intel : 0, mesRef: String(r[3] || '') };
    }
    return null;
  } catch (e) {
    Logger.log('[LIMITES] Falha ao consultar a conta: ' + e.message);
    return null;
  }
}

/**
 * Marca `item.limite` nos resultados da busca por nome: UMA busca nativa com
 * todas as contas (alternância) e UMA leitura do trecho encontrado. Nunca lança.
 */
function _marcarLimitesNaLista_(itens, ss) {
  try {
    const chaves = {};
    itens.forEach(it => { const c = _chaveConta_(it.conta); if (c) chaves[c] = true; });
    const lista = Object.keys(chaves);
    if (!lista.length) return;
    const sheet = ss.getSheetByName(LIMITES_ABA);
    if (!sheet || sheet.getLastRow() < 2) return;
    const regex = '^\\D*0*(?:' + lista.map(c => c.split('').join('\\D*')).join('|') + ')\\D*$';
    const linhas = _candidatas_(sheet, 1, regex, Infinity);
    if (!linhas.length) return;
    const primeira = linhas[0];
    const bloco = sheet.getRange(primeira, 1, linhas[linhas.length - 1] - primeira + 1, 3).getValues();
    const porConta = {};
    linhas.forEach(l => {
      const r = bloco[l - primeira];
      const c = _chaveConta_(r[0]);
      if (!chaves[c] || porConta[c]) return;
      const pre = _numeroLimite_(r[1]), intel = _numeroLimite_(r[2]);
      if (pre > 0 || intel > 0) porConta[c] = { preAprovado: pre > 0 ? pre : 0, limiteInteligente: intel > 0 ? intel : 0 };
    });
    itens.forEach(it => { const c = _chaveConta_(it.conta); if (porConta[c]) it.limite = porConta[c]; });
  } catch (e) {
    Logger.log('[LIMITES] Falha ao marcar a lista: ' + e.message);
  }
}

/** ADMIN — configuração e situação da carga, para o card do Admin. */
function buscarConfigLimites() {
  try {
    validarAcessoAdmin();
    const cfg = _configLimites_();
    const trigger = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === LIMITES_HANDLER)[0];
    let arquivoMaisRecente = '', erroPasta = '';
    try {
      const arquivos = _arquivosLimites_(cfg.pastaId);
      arquivoMaisRecente = arquivos.length ? arquivos[0].nome : '(nenhuma planilha .xlsx na pasta)';
    } catch (e) { erroPasta = e.message; }
    return {
      sucesso: true,
      pastaId: cfg.pastaId, colConta: cfg.colConta, colPre: cfg.colPre, colInteligente: cfg.colInteligente,
      padrao: LIMITES_PADRAO,
      arquivoCarregado: cfg.arquivoNome, mesRef: cfg.mesRef, qtd: cfg.qtd, ultimaCarga: cfg.ultimaCarga,
      ultimoStatus: cfg.ultimoStatus, ultimoResultado: cfg.ultimoResultado,
      arquivoMaisRecente, erroPasta,
      triggerAtivo: !!trigger,
      driveApiAtiva: typeof Drive !== 'undefined' && !!Drive.Files
    };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

/** ADMIN — salva pasta (ID ou link) e nomes das colunas. Confere o acesso à pasta antes. */
function salvarConfigLimites(cfg) {
  try {
    validarAcessoAdmin();
    cfg = cfg || {};
    const pastaId = _extrairIdDrive_(cfg.pastaId);
    const colConta = String(cfg.colConta || '').trim();
    const colPre = String(cfg.colPre || '').trim();
    const colInteligente = String(cfg.colInteligente || '').trim();
    if (!pastaId) return { sucesso: false, erro: 'Informe a pasta do Drive (link ou ID).' };
    if (!colConta || !colPre || !colInteligente) return { sucesso: false, erro: 'Informe o nome das três colunas.' };
    let nomePasta;
    try { nomePasta = DriveApp.getFolderById(pastaId).getName(); }
    catch (e) { return { sucesso: false, erro: 'Não consegui abrir a pasta. Confira o link e se a conta que publica o app tem acesso a ela.' }; }

    _setConfigInterna('LIMITES_PASTA_ID', pastaId);
    _setConfigInterna('LIMITES_COL_CONTA', colConta);
    _setConfigInterna('LIMITES_COL_PRE', colPre);
    _setConfigInterna('LIMITES_COL_INTELIGENTE', colInteligente);
    // Pasta ou colunas novas: a próxima carga lê de novo mesmo que o arquivo seja o mesmo.
    _setConfigInterna('LIMITES_ARQUIVO_ATUALIZADO', '');
    _instalarTriggerLimites_(_lerConfiguracoes_(['SYNC_HORA']).SYNC_HORA);
    return { sucesso: true, nomePasta };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

/** ADMIN — recarrega já, do arquivo mais recente (e garante o trigger diário). */
function atualizarLimitesAgora() {
  try {
    validarAcessoAdmin();
    const res = sincronizarLimitesPreAprovados(true);
    if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === LIMITES_HANDLER)) {
      _instalarTriggerLimites_(_lerConfiguracoes_(['SYNC_HORA']).SYNC_HORA);
    }
    return res;
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

// ==========================================
// SINCRONIZAÇÃO E BUSCA DE GRUPO ECONÔMICO
// ==========================================

const SYNC_GRUPO_ECONOMICO_CONFIG = {
  CSV_FILE_ID: '1ciKl0ZsJ1cnWDJGfX6lRrkCJy-6Tezyw',
  CSV_SEPARATOR: ',',
  CSV_ENCODING: 'UTF-8'
};

function substituirGrupoEconomico(dadosFiltrados) {
  try {
    validarContextoInterno();
    if (!dadosFiltrados || dadosFiltrados.length === 0) return { sucesso: false, erro: "Sem dados para gravar." };

    const TAMANHO_LOTE = 5000;
    const totalLinhas = dadosFiltrados.length;
    const neededCols = dadosFiltrados[0].length;
    const NOME_TEMP = '_GrupoEconomico_temp';

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Limpar aba temporária de execuções anteriores que possam ter falhado
    const sheetTemporaria = ss.getSheetByName(NOME_TEMP);
    if (sheetTemporaria) {
      ss.deleteSheet(sheetTemporaria);
    }

    // Determinar a posição da aba real
    let sheetReal = ss.getSheetByName('GrupoEconomico');
    const posReal = sheetReal ? sheetReal.getIndex() - 1 : 1;

    // Criar a aba temporária em posição adjacente
    let sheetTemp = ss.insertSheet(NOME_TEMP, posReal + 1);

    // Ajustar colunas excedentes
    const maxCols = sheetTemp.getMaxColumns();
    if (maxCols > neededCols) {
      sheetTemp.deleteColumns(neededCols + 1, maxCols - neededCols);
    }

    // Alocar todas as linhas necessárias de uma vez
    const maxRows = sheetTemp.getMaxRows();
    if (totalLinhas > maxRows) {
      sheetTemp.insertRowsAfter(maxRows, totalLinhas - maxRows);
    }
    SpreadsheetApp.flush();

    // Gravação em lotes na aba TEMPORÁRIA (base real intacta durante todo esse processo)
    let linhaInicio = 1;
    while (linhaInicio <= totalLinhas) {
      const lote = dadosFiltrados.slice(linhaInicio - 1, linhaInicio - 1 + TAMANHO_LOTE);
      sheetTemp.getRange(linhaInicio, 1, lote.length, neededCols).setValues(lote);
      SpreadsheetApp.flush();
      linhaInicio += TAMANHO_LOTE;
    }

    // Validação: confirmar contagem de linhas antes do swap
    const linhasGravadas = sheetTemp.getLastRow();
    if (linhasGravadas < totalLinhas) {
      ss.deleteSheet(sheetTemp);
      throw new Error(
        `Gravação incompleta na aba temporária: esperado ${totalLinhas} linhas, encontrado ${linhasGravadas}. ` +
        `A base 'GrupoEconomico' original foi preservada.`
      );
    }

    // ── SWAP ATÔMICO ────────────────────────────────────────────────────────
    if (sheetReal) {
      ss.deleteSheet(sheetReal);
    }
    sheetTemp.setName('GrupoEconomico');
    SpreadsheetApp.flush();
    // ────────────────────────────────────────────────────────────────────────

    // Usa _setConfigInterna pois esta função é chamada por triggers (sem sessão de usuário)
    _setConfigInterna('DATA_ATUALIZACAO_GRUPO_ECONOMICO', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"));

    return { sucesso: true, linhas: totalLinhas - 1 };
  } catch (e) {
    return { sucesso: false, erro: "Erro no servidor ao gravar grupo econômico: " + e.message };
  }
}


function sincronizarGrupoEconomico() {
  validarContextoInterno();
  const inicio = new Date();
  Logger.log(`[SYNC-GRUPO-ECONOMICO] Iniciando sincronização em ${inicio.toLocaleString('pt-BR')}`);

  try {
    const arquivo = DriveApp.getFileById(SYNC_GRUPO_ECONOMICO_CONFIG.CSV_FILE_ID);
    const conteudoCsv = arquivo.getBlob().getDataAsString(SYNC_GRUPO_ECONOMICO_CONFIG.CSV_ENCODING);

    const linhas = Utilities.parseCsv(conteudoCsv, SYNC_GRUPO_ECONOMICO_CONFIG.CSV_SEPARATOR.charCodeAt(0));
    if (!linhas || linhas.length < 2) {
      throw new Error('Arquivo CSV está vazio ou sem dados após o cabeçalho.');
    }

    const headers = linhas[0].map(h => String(h).trim());
    const idxAgrupador = headers.findIndex(h => h.toLowerCase() === 'nr_cpf_cnpj_agrupador');
    const idxIntegrante = headers.findIndex(h => h.toLowerCase() === 'nr_cpf_cnpj_integrante');

    if (idxAgrupador === -1 || idxIntegrante === -1) {
      throw new Error('Colunas obrigatórias não encontradas no CSV de Grupo Econômico.');
    }

    const headersFinais = ['nr_cpf_cnpj_agrupador', 'nr_cpf_cnpj_integrante'];
    const pares = [];

    for (let i = 1; i < linhas.length; i++) {
      const row = linhas[i];
      pares.push([
        String(row[idxAgrupador] || '').trim(),
        String(row[idxIntegrante] || '').trim()
      ]);
    }

    // O CSV do BI traz uma linha por registro de origem, e só guardamos 2 colunas:
    // o mesmo par se repetia ~14 vezes (463 mil linhas para 33 mil integrantes,
    // diagnóstico de 14/10/2026), deixando o comprometimento e a planilha inteira
    // lentos. Ver _deduplicarGrupo_.
    const dados = [headersFinais].concat(_deduplicarGrupo_(pares));

    const result = substituirGrupoEconomico(dados);
    Logger.log(`[SYNC-GRUPO-ECONOMICO] Gravado: ${JSON.stringify(result)}`);

    const duracao = ((new Date() - inicio) / 1000).toFixed(1);
    const resumo = `Sync Grupo Econômico concluído: ${dados.length - 1} registros atualizados (${pares.length} linhas no CSV, repetidas e vazias removidas). Duração: ${duracao}s.`;
    Logger.log(`[SYNC-GRUPO-ECONOMICO] ✅ ${resumo}`);

    return { sucesso: true, resumo };
  } catch (e) {
    Logger.log(`[SYNC-GRUPO-ECONOMICO] ❌ FALHA: ${e.message}`);
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Remove pares (agrupador, integrante) repetidos e linhas sem um dos dois.
 * Mantém cada par UMA vez, na posição da sua ÚLTIMA ocorrência — assim a regra
 * de buscarGrupoEconomico ("se o integrante aparecer em mais de um grupo, vale
 * a última linha") dá exatamente o mesmo resultado que na aba sem deduplicar.
 * Linhas sem agrupador ou sem integrante já eram ignoradas em todas as buscas.
 * A comparação é por dígitos (como nas buscas); grava o texto da última ocorrência.
 * @param {Array<Array<string>>} pares [[agrupador, integrante], ...] sem cabeçalho
 * @returns {Array<Array<string>>} pares únicos, na ordem das últimas ocorrências
 */
function _deduplicarGrupo_(pares) {
  const ultimaPosicao = new Map(); // "agrupador|integrante" → índice da última ocorrência
  for (let i = 0; i < pares.length; i++) {
    const a = _apenasDigitos_(pares[i][0]), b = _apenasDigitos_(pares[i][1]);
    if (!a || !b) continue;
    ultimaPosicao.set(a + '|' + b, i);
  }
  return [...ultimaPosicao.values()].sort((x, y) => x - y).map(i => pares[i]);
}

/**
 * Busca o grupo econômico de um CPF/CNPJ.
 * T2: Aceita um objeto Spreadsheet (ss) já aberto para evitar dupla abertura.
 * Se não receber ss, abre sozinha (compatibilidade retroativa).
 */
function buscarGrupoEconomico(cpfCnpjBusca, ssParam) {
  try {
    validarAcessoOperacional();
    const ss = ssParam || SpreadsheetApp.openById(SPREADSHEET_ID); // T2
    const sheetGrupo = ss.getSheetByName('GrupoEconomico');
    if (!sheetGrupo) return { temGrupo: false, integrantes: [] };

    if (sheetGrupo.getLastRow() <= 1) return { temGrupo: false, integrantes: [] };

    // Colunas: [A]=agrupador, [B]=integrante. Só valem linhas com os dois preenchidos;
    // se o integrante aparecer em mais de um grupo, vale a ÚLTIMA linha.
    //
    // Desempenho (14/10/2026): esta função deixou de ser chamada a cada busca de
    // associado (que agora usa _pertenceAGrupo_) e só roda no comprometimento,
    // que precisa da lista COMPLETA de integrantes. Para isso não há atalho: a aba
    // tem 463 mil linhas e há agrupadores com integrantes demais para buscar
    // linha a linha (tentativa que levou 18 s e estourou a memória). Fica o
    // método original — uma leitura e índice em memória —, lendo só as duas
    // colunas usadas em vez de getDataRange().
    const dataGrupo = sheetGrupo.getRange(1, 1, sheetGrupo.getLastRow(), 2).getValues();

    const cleanBusca = String(cpfCnpjBusca).replace(/\D/g, '');
    let cpfCnpjAgrupador = null;

    // T3: Criar índice Map uma única vez para lookup O(1)
    // Mapa: integrante → agrupador
    const mapaIntegranteParaAgrupador = new Map();
    // Mapa: agrupador → lista de integrantes
    const mapaAgrupadorParaIntegrantes = new Map();

    for (let i = 1; i < dataGrupo.length; i++) {
      const agrup = String(dataGrupo[i][0]).replace(/\D/g, '');
      const integ = String(dataGrupo[i][1]).replace(/\D/g, '');
      if (agrup && integ) {
        mapaIntegranteParaAgrupador.set(integ, agrup);
        if (!mapaAgrupadorParaIntegrantes.has(agrup)) {
          mapaAgrupadorParaIntegrantes.set(agrup, []);
        }
        mapaAgrupadorParaIntegrantes.get(agrup).push(String(dataGrupo[i][1]).trim());
      }
    }

    // Lookup O(1)
    cpfCnpjAgrupador = mapaIntegranteParaAgrupador.get(cleanBusca) || null;

    if (!cpfCnpjAgrupador) {
      return { temGrupo: false, integrantes: [] };
    }

    const integrantes = mapaAgrupadorParaIntegrantes.get(cpfCnpjAgrupador) || [];

    return {
      temGrupo: true,
      agrupador: cpfCnpjAgrupador,
      integrantes: integrantes
    };
  } catch (e) {
    Logger.log("Erro em buscarGrupoEconomico: " + e.message);
    return { temGrupo: false, integrantes: [] };
  }
}

// ==========================================
// CÁLCULO DE COMPROMETIMENTO DE RENDA
// Otimizações: CacheService + T2 (único ss) + T3 (Map O(1)) + T6 (servidor retorna dados brutos)
// ==========================================

/**
 * Núcleo de busca de dados de comprometimento.
 * T2: Recebe ss já aberto. T3: Usa Map para lookup O(1).
 * T6: Retorna apenas dados brutos (renda, endivCP, detalhamento) — sem calcular percentual.
 */
/**
 * Estado civil, união estável e CPF do cônjuge de uma linha de Associados, pelas
 * colunas acrescentadas em 27/10/2026 (vazias se a aba ainda não as tem).
 */
function _dadosCivis_(headers, row) {
  const norm = h => String(h).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().trim();
  const cab = headers.map(norm);
  const iEstado = cab.indexOf('ESTADO CIVIL');
  const iUniao = cab.findIndex(h => h.includes('UNIAO') && h.includes('ESTAVEL'));
  const iConjuge = cab.findIndex(h => h.includes('CONJUGE'));
  const texto = i => (i === -1 || row[i] == null) ? '' : String(row[i]).trim();
  return {
    estadoCivil: texto(iEstado),
    uniaoEstavel: _normalizarBoolean(texto(iUniao)) === 'SIM',
    docConjuge: texto(iConjuge)
  };
}

/**
 * Idade em anos a partir da coluna Idade de Associados (o sync grava o número;
 * aceita também data de nascimento). null quando não dá para saber. Nunca
 * devolve Date (o retorno vai para o cliente).
 */
function _idadeNumero_(valor) {
  if (typeof valor === 'number') return isFinite(valor) ? Math.floor(valor) : null;
  if (valor instanceof Date) {
    const i = _calcularIdade(Utilities.formatDate(valor, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    return i === '' ? null : Number(i);
  }
  const t = String(valor == null ? '' : valor).trim();
  if (!t) return null;
  if (/^\d{1,3}$/.test(t)) return Number(t);
  const i = _calcularIdade(t);
  return i === '' || isNaN(Number(i)) ? null : Number(i);
}

function _buscarDadosComprometimento(ss, cpfCnpjPrincipal) {
  let detalhamentoPorPessoa = [];
  let rendaTotal = 0;
  let endivCurtoPrazoTotal = 0;

  // T2: Reutiliza ss passado como parâmetro
  let grupoInfo = buscarGrupoEconomico(cpfCnpjPrincipal, ss);
  let cpfsParaBuscar = new Set();

  if (grupoInfo.temGrupo && grupoInfo.integrantes && grupoInfo.integrantes.length > 0) {
    grupoInfo.integrantes.forEach(i => cpfsParaBuscar.add(String(i).replace(/\D/g, '')));
  } else if (cpfCnpjPrincipal) {
    cpfsParaBuscar.add(String(cpfCnpjPrincipal).replace(/\D/g, ''));
  }

  if (cpfsParaBuscar.size > 0) {
    const sheetAssoc = ss.getSheetByName('Associados');
    if (sheetAssoc && sheetAssoc.getLastRow() > 1) {
      const headersAssoc = _cabecalho_(sheetAssoc);

      const idxNome = headersAssoc.findIndex(h => { const s = String(h).toUpperCase().trim(); return s === 'ASSOCIADO' || s === 'NOME DO ASSOCIADO' || s === 'NOME COMPLETO' || s === 'NOME'; });
      const idxCpf = headersAssoc.findIndex(h => String(h).toUpperCase().includes('CPF'));
      const idxPessoa = headersAssoc.findIndex(h => String(h).toUpperCase().includes('TIPO') && String(h).toUpperCase().includes('PESSOA'));
      const idxMei = headersAssoc.findIndex(h => String(h).toUpperCase().includes('MEI'));
      const idxSimples = headersAssoc.findIndex(h => String(h).toUpperCase().includes('SIMPLES'));
      const idxRenda = headersAssoc.findIndex(h => String(h).toUpperCase().includes('RENDA') && !String(h).toUpperCase().includes('DATA'));
      const idxEndivCP = headersAssoc.findIndex(h => String(h).toUpperCase().includes('MENSAL') && String(h).toUpperCase().includes('ENDIVIDAMENTO'));
      const idxDataRenda = headersAssoc.findIndex(h => String(h).toUpperCase().includes('DATA ATUALIZACAO RENDA'));
      const idxDataSCR = headersAssoc.findIndex(h => String(h).toUpperCase().includes('DATA CONSULTA SCR'));
      const idxColab = headersAssoc.findIndex(h => String(h).toUpperCase().includes('COLABORADOR'));
      // Idade: o checklist deixa de fora os menores de idade (o comprometimento os mantém).
      const idxIdade = headersAssoc.findIndex(h => String(h).toUpperCase().includes('IDADE') || String(h).toUpperCase().includes('NASCIMENTO'));

      if (idxCpf !== -1 && idxRenda !== -1 && idxEndivCP !== -1) {
        // Desempenho (14/10/2026): antes indexava a aba Associados INTEIRA. Agora
        // localiza só as linhas dos CPFs do grupo — primeira ocorrência de cada,
        // como o índice de antes (que não sobrescrevia).
        const indiceCpf = _localizarPorDigitos_(sheetAssoc, idxCpf + 1, [...cpfsParaBuscar]);

        // T3: Lookup O(1) por CPF — sem varredura linear
        cpfsParaBuscar.forEach(cpf => {
          const rowData = indiceCpf.get(cpf);
          if (!rowData) return;

          const nome = idxNome !== -1 ? String(rowData[idxNome]).trim() : 'Nome Indisponível';
          const documento = String(rowData[idxCpf]).trim();
          const tipoPessoa = idxPessoa !== -1 ? String(rowData[idxPessoa]).trim() : 'PF';
          const isMei = idxMei !== -1 ? String(rowData[idxMei]).trim() : 'NÃO';
          const isSimples = idxSimples !== -1 ? String(rowData[idxSimples]).trim() : 'NÃO';

          const rawRenda = rowData[idxRenda];
          const renda = typeof rawRenda === 'number' ? rawRenda : parseFloat(String(rawRenda).replace(/\./g, '').replace(',', '.')) || 0;
          const rawEndiv = rowData[idxEndivCP];
          const endiv = typeof rawEndiv === 'number' ? rawEndiv : parseFloat(String(rawEndiv).replace(/\./g, '').replace(',', '.')) || 0;
          const dataRenda = idxDataRenda !== -1 ? String(rowData[idxDataRenda]).trim() : 'N/D';
          const dataSCR = idxDataSCR !== -1 ? String(rowData[idxDataSCR]).trim() : 'N/D';
          const isColab = idxColab !== -1 ? String(rowData[idxColab]).toUpperCase().trim() === 'SIM' : false;
          const idade = idxIdade !== -1 ? _idadeNumero_(rowData[idxIdade]) : null;
          const civis = _dadosCivis_(headersAssoc, rowData); // checklist: estado civil e cônjuge

          rendaTotal += renda;
          endivCurtoPrazoTotal += endiv;
          detalhamentoPorPessoa.push({
            nome,
            cpfCnpj: documento,
            tipoPessoa,
            mei: isMei,
            simples: isSimples,
            renda,
            endivCP: endiv,
            dataRenda,
            dataSCR,
            colaborador: isColab,
            idade,
            estadoCivil: civis.estadoCivil,
            uniaoEstavel: civis.uniaoEstavel,
            docConjuge: civis.docConjuge
          });
        });
      }
    }
  }

  // T6: Retorna apenas dados brutos — o cálculo do percentual fica no frontend
  return {
    sucesso: true,
    rendaTotalGrupo: rendaTotal,
    endivCurtoPrazoTotalGrupo: endivCurtoPrazoTotal,
    detalhamentoPorPessoa: detalhamentoPorPessoa,
    temGrupoEconomico: grupoInfo.temGrupo
  };
}

/**
 * Calcula o comprometimento de renda.
 * CacheService: verifica cache primeiro. T2+T3: via _buscarDadosComprometimento.
 * T6: não calcula o percentual — retorna apenas dados brutos para o frontend processar.
 */
function calcularComprometimentoRendaGlobal(cpfCnpjPrincipal, valorParcela, periodicidade, membrosManuais) {
  try {
    validarAcessoOperacional();
    // Caso Prospect: membros manuais fornecidos pelo frontend — sem consulta à planilha
    if (membrosManuais && Array.isArray(membrosManuais) && membrosManuais.length > 0) {
      let rendaTotal = 0;
      let endivCurtoPrazoTotal = 0;
      const detalhamentoPorPessoa = [];

      membrosManuais.forEach(membro => {
        const renda = parseFloat(String(membro.renda).replace(/\./g, '').replace(',', '.')) || 0;
        const endiv = parseFloat(String(membro.endivCP).replace(/\./g, '').replace(',', '.')) || 0;
        rendaTotal += renda;
        endivCurtoPrazoTotal += endiv;
        detalhamentoPorPessoa.push({ nome: membro.nome || 'Prospect/Membro', renda, endivCP: endiv, colaborador: false });
      });

      // T6: retorna dados brutos — frontend calcula o percentual
      return {
        sucesso: true,
        rendaTotalGrupo: rendaTotal,
        endivCurtoPrazoTotalGrupo: endivCurtoPrazoTotal,
        detalhamentoPorPessoa: detalhamentoPorPessoa,
        temGrupoEconomico: false
      };
    }

    // Caso Associado: verificar CacheService primeiro
    const cache = CacheService.getUserCache();
    const chave = 'COMP3_' + String(cpfCnpjPrincipal || '').replace(/\D/g, ''); // COMP3: inclui estado civil e cônjuge (27/10/2026)
    const cached = cache.get(chave);
    if (cached) {
      // Cache hit: retorno em ~50ms sem abrir a planilha
      return JSON.parse(cached);
    }

    // Cache miss: calcular e armazenar
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); // T2: única abertura
    const resultado = _buscarDadosComprometimento(ss, cpfCnpjPrincipal);

    if (resultado.sucesso) {
      try {
        cache.put(chave, JSON.stringify(resultado), 300); // TTL: 5 minutos
      } catch (cacheErr) { /* Silencioso se objeto > 100KB */ }
    }

    return resultado;
  } catch (e) {
    return { sucesso: false, erro: 'Erro ao calcular comprometimento: ' + e.message };
  }
}

/**
 * Busca as regras do Checklist na API externa.
 * @returns {Object} O JSON de resposta da API
 */
function buscarRegrasChecklistAPI() {
  try {
    validarAcessoOperacional();

    const options = {
      'method': 'get',
      'headers': {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      },
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(URL_API_CHECKLIST, options);

    // Com muteHttpExceptions ativo, um 302/401/500 NÃO lança — vem como corpo.
    // Sem esta checagem, uma página de login do Google chegava ao JSON.parse e
    // virava um SyntaxError genérico, indistinguível de "API fora do ar".
    const status = response.getResponseCode();
    if (status !== 200) {
      return { success: false, erro: `API de Checklist respondeu HTTP ${status}.` };
    }

    let json;
    try {
      json = JSON.parse(response.getContentText());
    } catch (eParse) {
      return { success: false, erro: "API de Checklist retornou conteúdo que não é JSON." };
    }

    // Valida o formato antes de devolver: o frontend renderiza estes campos.
    if (!json || typeof json !== 'object') {
      return { success: false, erro: "API de Checklist retornou formato inesperado." };
    }
    if (json.data !== undefined && !Array.isArray(json.data)) {
      return { success: false, erro: "API de Checklist retornou 'data' em formato inesperado." };
    }

    return json;
  } catch (e) {
    return { success: false, erro: "Erro ao buscar API de Checklist: " + e.message };
  }
}

// ============================================================================
// MÓDULO DE PESQUISAS ADMIN
// Implementado em: 2026-08-31
// Permite ao administrador criar, editar, encerrar e analisar pesquisas.
// Banco de dados: abas "PesquisasAdmin" e "RespostasPesquisas" no Google Sheets.
// ============================================================================

/**
 * Garante que a aba possua a linha de cabeçalho esperada.
 *
 * Todo o módulo de Pesquisas lê os dados com `for (let i = 1; ...)`, ou seja,
 * assume que a linha 1 é cabeçalho. Se a aba existir SEM cabeçalho, o primeiro
 * registro é engolido silenciosamente e o módulo aparenta estar vazio.
 *
 * Cobre os dois cenários problemáticos:
 *   - aba existente e vazia  → grava o cabeçalho;
 *   - aba existente com dados começando na linha 1 → insere o cabeçalho acima,
 *     preservando todos os registros já gravados.
 *
 * @param {Sheet} sheet Aba a validar
 * @param {Array<string>} cols Nomes das colunas esperadas
 */
function _garantirCabecalho_(sheet, cols) {
  const primeiraCelula = String(sheet.getRange(1, 1).getValue()).trim();
  if (primeiraCelula === cols[0]) return; // cabeçalho já presente

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(cols);
  } else {
    // Há dados ocupando a linha 1: abre espaço e reinsere o cabeçalho.
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, cols.length).setValues([cols]);
  }
  formatarCabecalho(sheet, cols.length);
  sheet.setFrozenRows(1);
}

/**
 * Converte um valor lido de célula em texto seguro para trafegar ao cliente.
 *
 * CRÍTICO: o google.script.run NÃO aceita objetos Date no retorno — a requisição
 * falha e o cliente recebe `null` no withSuccessHandler, sem erro no
 * withFailureHandler. Como o Sheets converte automaticamente strings como
 * "09/09/2026 14:19" em datetime, qualquer coluna de data lida com getValues()
 * volta como Date e quebra a chamada inteira.
 *
 * @param {*} valor Valor bruto vindo de getValues()
 * @returns {string} Texto formatado (dd/MM/yyyy HH:mm para datas)
 */
function _valorCelulaParaTexto_(valor) {
  if (valor === null || valor === undefined || valor === '') return '';
  if (Object.prototype.toString.call(valor) === '[object Date]') {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  }
  return String(valor);
}

/**
 * Obtém ou cria a aba de pesquisas e retorna a referência.
 * Colunas: ID | Titulo | Descricao | Status | DataCriacao | DataEncerramento | Perguntas(JSON) | PublicoAlvo | AgenciaAlvo
 */
function _getSheetPesquisas_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const cols = ['ID', 'Titulo', 'Descricao', 'Status', 'DataCriacao', 'DataEncerramento', 'Perguntas', 'PublicoAlvo', 'AgenciaAlvo'];
  let sheet = ss.getSheetByName('PesquisasAdmin');
  if (!sheet) {
    sheet = ss.insertSheet('PesquisasAdmin');
    sheet.appendRow(cols);
    formatarCabecalho(sheet, cols.length);
    // Congelar cabeçalho
    sheet.setFrozenRows(1);
    // Ajustar larguras das colunas principais
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(3, 280);
    sheet.setColumnWidth(7, 400);
  } else {
    _garantirCabecalho_(sheet, cols);
  }
  return sheet;
}

/**
 * Obtém ou cria a aba de respostas e retorna a referência.
 * Colunas: ID_Pesquisa | DataResposta | EmailRespondente | NomeRespondente | Respostas(JSON) | AgenciaRespondente
 */
function _getSheetRespostas_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const cols = ['ID_Pesquisa', 'DataResposta', 'EmailRespondente', 'NomeRespondente', 'Respostas', 'AgenciaRespondente'];
  let sheet = ss.getSheetByName('RespostasPesquisas');
  if (!sheet) {
    sheet = ss.insertSheet('RespostasPesquisas');
    sheet.appendRow(cols);
    formatarCabecalho(sheet, cols.length);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(5, 500);
  } else {
    _garantirCabecalho_(sheet, cols);
  }
  return sheet;
}

/**
 * Gera um ID único simples baseado em timestamp.
 */
function _gerarIdPesquisa_() {
  return 'PQ_' + new Date().getTime();
}

/**
 * Busca o perfil e a agência do usuário logado.
 */
function _getDadosUsuarioLogado_() {
  // Desde 15/10/2026 vem do Hub, pela mesma validação de acesso do app.
  try {
    const acesso = _acessoAtual_();
    if (!acesso || !acesso.autorizado) return { email: '', nome: '', perfil: 'AGÊNCIA', agencia: '' };
    return { email: acesso.email, nome: acesso.nome || acesso.email, perfil: acesso.perfil || 'AGÊNCIA', agencia: acesso.agencia || '' };
  } catch (e) {
    return { email: '', nome: '', perfil: 'AGÊNCIA', agencia: '' };
  }
}

// ----------------------------------------------------------------------------
// FUNÇÕES ADMIN — Gestão de Pesquisas
// ----------------------------------------------------------------------------

/**
 * Lista todas as pesquisas com contador de respostas.
 * @returns {Object} { sucesso: true, pesquisas: [...] }
 */
function listarPesquisasAdmin() {
  try {
    validarAcessoAdmin();
    const sheet = _getSheetPesquisas_();
    const sheetResp = _getSheetRespostas_();
    const data = sheet.getDataRange().getValues();
    const dataResp = sheetResp.getDataRange().getValues();

    // Monta mapa de contagem de respostas por pesquisa
    const contagem = {};
    for (let i = 1; i < dataResp.length; i++) {
      const idP = String(dataResp[i][0]);
      if (!contagem[idP]) contagem[idP] = 0;
      contagem[idP]++;
    }

    const pesquisas = [];
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][0]);
      if (!id) continue;
      pesquisas.push({
        id: id,
        titulo: _valorCelulaParaTexto_(data[i][1]),
        descricao: _valorCelulaParaTexto_(data[i][2]),
        status: _valorCelulaParaTexto_(data[i][3]) || 'Ativa',
        dataCriacao: _valorCelulaParaTexto_(data[i][4]),
        dataEncerramento: _valorCelulaParaTexto_(data[i][5]),
        perguntas: data[i][6] ? JSON.parse(data[i][6]) : [],
        publicoAlvo: _valorCelulaParaTexto_(data[i][7]) || 'TODOS',
        agenciaAlvo: _valorCelulaParaTexto_(data[i][8]),
        totalRespostas: contagem[id] || 0
      });
    }
    // Mais recente primeiro
    pesquisas.sort((a, b) => b.id.localeCompare(a.id));
    return { sucesso: true, pesquisas };
  } catch (e) {
    return { sucesso: false, erro: 'Erro ao listar pesquisas: ' + e.message };
  }
}

/**
 * Cria ou edita uma pesquisa.
 * @param {Object} dados - { id?, titulo, descricao, perguntas, publicoAlvo, agenciaAlvo }
 * @returns {Object} { sucesso: true, id }
 */
function salvarPesquisa(dados) {
  try {
    validarAcessoAdmin();
    const sheet = _getSheetPesquisas_();
    const dataHoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    const perguntasJson = JSON.stringify(dados.perguntas || []);

    if (dados.id) {
      // Edição: localiza a linha pelo ID
      const allData = sheet.getDataRange().getValues();
      for (let i = 1; i < allData.length; i++) {
        if (String(allData[i][0]) === String(dados.id)) {
          const row = i + 1; // 1-indexed
          sheet.getRange(row, 2).setValue(dados.titulo || '');
          sheet.getRange(row, 3).setValue(dados.descricao || '');
          sheet.getRange(row, 7).setValue(perguntasJson);
          sheet.getRange(row, 8).setValue(dados.publicoAlvo || 'TODOS');
          sheet.getRange(row, 9).setValue(dados.agenciaAlvo || '');
          return { sucesso: true, id: dados.id };
        }
      }
      return { sucesso: false, erro: 'Pesquisa não encontrada para edição.' };
    } else {
      // Criação
      const novoId = _gerarIdPesquisa_();
      sheet.appendRow([
        novoId,
        dados.titulo || '',
        dados.descricao || '',
        'Ativa',
        dataHoje,
        '',
        perguntasJson,
        dados.publicoAlvo || 'TODOS',
        dados.agenciaAlvo || ''
      ]);
      return { sucesso: true, id: novoId };
    }
  } catch (e) {
    return { sucesso: false, erro: 'Erro ao salvar pesquisa: ' + e.message };
  }
}

/**
 * Encerra uma pesquisa ativa.
 * @param {string} idPesquisa
 * @returns {Object} { sucesso: true }
 */
function encerrarPesquisa(idPesquisa) {
  try {
    validarAcessoAdmin();
    const sheet = _getSheetPesquisas_();
    const allData = sheet.getDataRange().getValues();
    const dataHoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][0]) === String(idPesquisa)) {
        sheet.getRange(i + 1, 4).setValue('Encerrada');
        sheet.getRange(i + 1, 6).setValue(dataHoje);
        return { sucesso: true };
      }
    }
    return { sucesso: false, erro: 'Pesquisa não encontrada.' };
  } catch (e) {
    return { sucesso: false, erro: 'Erro ao encerrar pesquisa: ' + e.message };
  }
}

/**
 * Exclui uma pesquisa e todas as suas respostas.
 * @param {string} idPesquisa
 * @returns {Object} { sucesso: true }
 */
function excluirPesquisa(idPesquisa) {
  try {
    validarAcessoAdmin();
    // Remove da aba de pesquisas
    const sheet = _getSheetPesquisas_();
    const allData = sheet.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][0]) === String(idPesquisa)) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    // Remove respostas associadas
    const sheetResp = _getSheetRespostas_();
    const dataResp = sheetResp.getDataRange().getValues();
    // Percorre de trás para frente para não deslocar índices
    for (let i = dataResp.length - 1; i >= 1; i--) {
      if (String(dataResp[i][0]) === String(idPesquisa)) {
        sheetResp.deleteRow(i + 1);
      }
    }
    return { sucesso: true };
  } catch (e) {
    return { sucesso: false, erro: 'Erro ao excluir pesquisa: ' + e.message };
  }
}

/**
 * Busca os resultados agregados de uma pesquisa para análise.
 * @param {string} idPesquisa
 * @returns {Object} { sucesso: true, pesquisa, respostas, agregado }
 */
function buscarResultadosPesquisa(idPesquisa) {
  try {
    validarAcessoAdmin();
    // Busca dados da pesquisa
    const sheet = _getSheetPesquisas_();
    const allData = sheet.getDataRange().getValues();
    let pesquisa = null;
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][0]) === String(idPesquisa)) {
        pesquisa = {
          id: _valorCelulaParaTexto_(allData[i][0]),
          titulo: _valorCelulaParaTexto_(allData[i][1]),
          descricao: _valorCelulaParaTexto_(allData[i][2]),
          status: _valorCelulaParaTexto_(allData[i][3]),
          dataCriacao: _valorCelulaParaTexto_(allData[i][4]),
          dataEncerramento: _valorCelulaParaTexto_(allData[i][5]),
          perguntas: allData[i][6] ? JSON.parse(allData[i][6]) : [],
          publicoAlvo: _valorCelulaParaTexto_(allData[i][7]),
          agenciaAlvo: _valorCelulaParaTexto_(allData[i][8])
        };
        break;
      }
    }
    if (!pesquisa) return { sucesso: false, erro: 'Pesquisa não encontrada.' };

    // Busca respostas
    const sheetResp = _getSheetRespostas_();
    const dataResp = sheetResp.getDataRange().getValues();
    const respostas = [];
    for (let i = 1; i < dataResp.length; i++) {
      if (String(dataResp[i][0]) === String(idPesquisa)) {
        respostas.push({
          data: _valorCelulaParaTexto_(dataResp[i][1]),
          email: _valorCelulaParaTexto_(dataResp[i][2]),
          nome: _valorCelulaParaTexto_(dataResp[i][3]),
          respostas: dataResp[i][4] ? JSON.parse(dataResp[i][4]) : {},
          agencia: _valorCelulaParaTexto_(dataResp[i][5])
        });
      }
    }

    // Agrega resultados por pergunta
    const agregado = {};
    pesquisa.perguntas.forEach((perg, idx) => {
      const key = 'p' + idx;
      if (perg.tipo === 'multipla') {
        // Conta ocorrências de cada opção
        const contagem = {};
        perg.opcoes.forEach(op => contagem[op] = 0);
        respostas.forEach(r => {
          const val = r.respostas[key];
          if (val !== undefined && val !== null && val !== '') {
            contagem[val] = (contagem[val] || 0) + 1;
          }
        });
        agregado[key] = { tipo: 'multipla', contagem };
      } else if (perg.tipo === 'escala') {
        // Calcula distribuição e média
        const dist = { '1': 0, '2': 0, '3': 0, '4': 0, '5': 0 };
        let soma = 0, count = 0;
        respostas.forEach(r => {
          const val = r.respostas[key];
          if (val) {
            dist[String(val)] = (dist[String(val)] || 0) + 1;
            soma += Number(val);
            count++;
          }
        });
        agregado[key] = { tipo: 'escala', dist, media: count > 0 ? (soma / count).toFixed(1) : '—' };
      } else {
        // Texto livre: retorna lista de respostas não vazias
        const textos = respostas
          .map(r => ({ texto: r.respostas[key] || '', respondente: r.nome, agencia: r.agencia }))
          .filter(t => t.texto.trim() !== '');
        agregado[key] = { tipo: 'texto', textos };
      }
    });

    return { sucesso: true, pesquisa, respostas, agregado, totalRespostas: respostas.length };
  } catch (e) {
    return { sucesso: false, erro: 'Erro ao buscar resultados: ' + e.message };
  }
}

// ----------------------------------------------------------------------------
// FUNÇÕES PARA OPERADORES — Responder pesquisas ativas
// ----------------------------------------------------------------------------

/**
 * Verifica se há pesquisa ativa que o usuário logado ainda não respondeu
 * e que é direcionada ao seu perfil/agência.
 * @returns {Object} { temPesquisa: true, pesquisa: {...} } ou { temPesquisa: false }
 */
function verificarPesquisaAtiva() {
  try {
    validarAcessoOperacional();
    const usuario = _getDadosUsuarioLogado_();
    const email = usuario.email;
    if (!email) return { temPesquisa: false };

    const sheet = _getSheetPesquisas_();
    const allData = sheet.getDataRange().getValues();

    // Busca emails já respondidos por esse usuário
    const sheetResp = _getSheetRespostas_();
    const dataResp = sheetResp.getDataRange().getValues();
    const jaRespondidas = new Set();
    for (let i = 1; i < dataResp.length; i++) {
      if (String(dataResp[i][2]).toLowerCase() === email.toLowerCase()) {
        jaRespondidas.add(String(dataResp[i][0]));
      }
    }

    // Procura pesquisa ativa não respondida e compatível com o perfil
    for (let i = 1; i < allData.length; i++) {
      const id = String(allData[i][0]);
      const status = String(allData[i][3]);
      if (status !== 'Ativa') continue;
      if (jaRespondidas.has(id)) continue;

      const publicoAlvo = String(allData[i][7] || 'TODOS').toUpperCase();
      const agenciaAlvo = String(allData[i][8] || '').trim();
      const perfilUsuario = String(usuario.perfil || '').toUpperCase();
      const agenciaUsuario = String(usuario.agencia || '').trim();

      // Verifica compatibilidade de público
      let elegivel = false;
      if (publicoAlvo === 'TODOS') {
        elegivel = true;
      } else if (publicoAlvo === 'SEDE ADMINISTRATIVA' && perfilUsuario === 'SEDE ADMINISTRATIVA') {
        elegivel = true;
      } else if (publicoAlvo === 'AGÊNCIA') {
        if (agenciaAlvo === '' || agenciaAlvo.toLowerCase() === agenciaUsuario.toLowerCase()) {
          elegivel = perfilUsuario === 'AGÊNCIA';
        }
      }

      if (elegivel) {
        return {
          temPesquisa: true,
          pesquisa: {
            id: id,
            titulo: _valorCelulaParaTexto_(allData[i][1]),
            descricao: _valorCelulaParaTexto_(allData[i][2]),
            perguntas: allData[i][6] ? JSON.parse(allData[i][6]) : []
          }
        };
      }
    }
    return { temPesquisa: false };
  } catch (e) {
    return { temPesquisa: false };
  }
}

/**
 * Salva a resposta de um operador a uma pesquisa.
 * @param {string} idPesquisa
 * @param {Object} respostas - { p0: valor, p1: valor, ... }
 * @returns {Object} { sucesso: true }
 */
function salvarRespostaPesquisaAdmin(idPesquisa, respostas) {
  try {
    validarAcessoOperacional();
    const usuario = _getDadosUsuarioLogado_();
    const sheetResp = _getSheetRespostas_();
    const dataHoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    sheetResp.appendRow([
      idPesquisa,
      dataHoje,
      usuario.email,
      usuario.nome,
      JSON.stringify(respostas || {}),
      usuario.agencia || ''
    ]);
    return { sucesso: true };
  } catch (e) {
    return { sucesso: false, erro: 'Erro ao salvar resposta: ' + e.message };
  }
}

// ============================================================================
// MIGRAÇÃO DA PESQUISA LEGADA → MÓDULO DE PESQUISAS ADMIN
// Converte a pesquisa hardcoded (aba "Pesquisa") para o formato dinâmico.
// ============================================================================

/**
 * Migra a pesquisa de satisfação legada (Comprometimento + Checklist) para o
 * novo módulo de Pesquisas Admin, preservando todas as respostas existentes.
 *
 * @returns {Object} { sucesso, mensagem, pesquisaCriada, respostasMigradas }
 */
function migrarPesquisaLegadaParaAdmin() {
  try {
    validarAcessoAdmin();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ── 1. Verifica se a pesquisa legada já existe no novo módulo ────────────
    const sheetPesquisas = _getSheetPesquisas_();
    const dadosPesquisas = sheetPesquisas.getDataRange().getValues();
    const ID_LEGADA = 'PQ_LEGADA_SATISFACAO';
    for (let i = 1; i < dadosPesquisas.length; i++) {
      if (String(dadosPesquisas[i][0]) === ID_LEGADA) {
        return {
          sucesso: false,
          mensagem: 'A pesquisa legada já foi migrada anteriormente (ID: ' + ID_LEGADA + '). Verifique o módulo de Pesquisas.'
        };
      }
    }

    // ── 2. Define as perguntas originais da pesquisa legada ──────────────────
    const perguntas = [
      {
        tipo: 'multipla',
        enunciado: '💰 [Comprometimento de Renda] Com que frequência você utiliza essa função?',
        opcoes: ['Nunca utilizei', 'Raramente', 'Às vezes', 'Frequentemente']
      },
      {
        tipo: 'multipla',
        enunciado: '💰 [Comprometimento de Renda] O quanto você considera esta função útil no seu dia a dia?',
        opcoes: ['Nada útil', 'Pouco útil', 'Útil', 'Muito útil']
      },
      {
        tipo: 'multipla',
        enunciado: '💰 [Comprometimento de Renda] Os resultados são compatíveis com o retorno do sistema oficial?',
        opcoes: ['Muito diferente', 'Às vezes difere', 'Bastante próximo', 'Não consigo comparar']
      },
      {
        tipo: 'multipla',
        enunciado: '📄 [Checklist de Documentos] Com que frequência você utiliza o gerador de checklist?',
        opcoes: ['Nunca utilizei', 'Raramente', 'Às vezes', 'Frequentemente']
      },
      {
        tipo: 'multipla',
        enunciado: '📄 [Checklist de Documentos] O quanto você considera esta função útil no seu dia a dia?',
        opcoes: ['Nada útil', 'Pouco útil', 'Útil', 'Muito útil']
      },
      {
        tipo: 'multipla',
        enunciado: '📄 [Checklist de Documentos] O checklist gerado reflete corretamente os documentos exigidos?',
        opcoes: ['Raramente', 'Às vezes', 'Quase sempre', 'Sempre', 'Não usei o suficiente']
      },
      {
        tipo: 'texto',
        enunciado: '💬 Tem alguma sugestão ou crítica sobre o Simulador de Crédito?'
      }
    ];

    const dataHoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

    // ── 3. Cria a entrada na aba PesquisasAdmin ──────────────────────────────
    sheetPesquisas.appendRow([
      ID_LEGADA,
      'Pesquisa de Satisfação — Comprometimento & Checklist (Migrada)',
      'Pesquisa original do sistema sobre uso das funções de Comprometimento de Renda e Checklist de Documentos. Migrada automaticamente do formato legado.',
      'Encerrada',          // Status: já encerrada (legada)
      '01/01/2026 00:00',   // Data de criação aproximada
      dataHoje,             // Data de encerramento = agora
      JSON.stringify(perguntas),
      'TODOS',
      ''
    ]);

    // ── 4. Migra respostas da aba "Pesquisa" → "RespostasPesquisas" ──────────
    // Colunas originais: [Data, Email, Nome, P1, P2, P3, P4, P5, P6, Observacao]
    //                      0     1     2    3   4   5   6   7   8     9
    const sheetLegada = ss.getSheetByName('Pesquisa');
    const sheetRespostas = _getSheetRespostas_();
    let respostasMigradas = 0;

    if (sheetLegada) {
      const dadosLegados = sheetLegada.getDataRange().getValues();
      // Busca emails que já foram migrados (evita duplicatas em re-execuções)
      const dadosRespostas = sheetRespostas.getDataRange().getValues();
      const jaRegistrados = new Set();
      for (let i = 1; i < dadosRespostas.length; i++) {
        if (String(dadosRespostas[i][0]) === ID_LEGADA) {
          jaRegistrados.add(String(dadosRespostas[i][1]) + '_' + String(dadosRespostas[i][2]));
        }
      }

      for (let i = 1; i < dadosLegados.length; i++) {
        const linha = dadosLegados[i];
        const data  = linha[0] || '';
        const email = String(linha[1] || '').trim();
        const nome  = String(linha[2] || '').trim();

        if (!email) continue;
        const chave = String(data) + '_' + email;
        if (jaRegistrados.has(chave)) continue;

        // Mapeia as colunas legadas para o formato dinâmico { p0..p6 }
        const respostas = {
          p0: linha[3] || '',  // P1_FreqComprometimento
          p1: linha[4] || '',  // P2_UtilidadeComprometimento
          p2: linha[5] || '',  // P3_PrecisaoComprometimento
          p3: linha[6] || '',  // P4_FreqChecklist
          p4: linha[7] || '',  // P5_UtilidadeChecklist
          p5: linha[8] || '',  // P6_PrecisaoChecklist
          p6: linha[9] || ''   // Observacao (texto livre)
        };

        // Busca agência do usuário (melhor esforço)
        let agencia = '';
        const sheetUsers = ss.getSheetByName('Usuarios');
        if (sheetUsers) {
          const usersData = sheetUsers.getDataRange().getValues();
          for (let u = 1; u < usersData.length; u++) {
            if (String(usersData[u][1]).trim().toLowerCase() === email.toLowerCase()) {
              agencia = usersData[u][3] || '';
              break;
            }
          }
        }

        sheetRespostas.appendRow([
          ID_LEGADA,
          data,
          email,
          nome,
          JSON.stringify(respostas),
          agencia
        ]);
        respostasMigradas++;
      }
    }

    return {
      sucesso: true,
      mensagem: `Migração concluída! Pesquisa criada com ID "${ID_LEGADA}" e ${respostasMigradas} resposta(s) migrada(s) com sucesso.`,
      pesquisaCriada: ID_LEGADA,
      respostasMigradas: respostasMigradas
    };

  } catch (e) {
    return { sucesso: false, mensagem: 'Erro durante a migração: ' + e.message };
  }
}

// ============================================================================
// ATENÇÃO — NÃO ADICIONE CHAMADAS EXECUTÁVEIS EM ESCOPO GLOBAL NESTE ARQUIVO.
//
// O Apps Script reavalia TODO o Code.gs a cada invocação (doGet, cada
// google.script.run, cada trigger). Qualquer IIFE ou chamada solta aqui roda
// centenas de vezes por dia, por todos os usuários.
//
// Havia aqui uma IIFE chamando registrarNovidade() para anunciar o módulo de
// Pesquisas Admin. Ela gravava uma linha na aba "Novidades" a cada RPC de
// qualquer operador, inflando a aba com milhares de registros idênticos e
// travando o modal de Novidades. Removida em 09/09/2026.
//
// Para publicar uma novidade, execute registrarNovidade(resumo, detalhes)
// manualmente pelo editor do Apps Script.
// ============================================================================

