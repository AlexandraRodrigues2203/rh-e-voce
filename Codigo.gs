
const CONFIG = {
  EMPRESA: 'Ferramentaria Soares',
  EMAIL_RH: 'alexandra.rodrigues@ferramentariasoares.com.br',
  PASTA_RAIZ: 'PORTAL RH - FERRAMENTARIA SOARES',
  PLANILHA: 'BANCO DE DADOS - PORTAL RH',
  FUSO: 'America/Sao_Paulo',
  DURACAO_SESSAO_SEGUNDOS: 21600,
  EMAIL_ADMINISTRADOR: 'alexandra.rodrigues@ferramentariasoares.com.br',
  PIN_INICIAL: '123456'
};

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('RH | e Você | Ferramentaria Soares')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function instalarEtapa5() {
  const pastaRaiz = obterOuCriarPasta_(DriveApp.getRootFolder(), CONFIG.PASTA_RAIZ);
  const pastaBanco = obterOuCriarPasta_(pastaRaiz, '02 - Banco de Dados');
  const pastaAtestados = obterOuCriarPasta_(pastaRaiz, '03 - Atestados');
  const pastaFerias = obterOuCriarPasta_(pastaRaiz, '05 - Férias');
  const pastaComunicados = obterOuCriarPasta_(pastaRaiz, '06 - Comunicados');
  const pastaDocumentos = obterOuCriarPasta_(pastaRaiz, '07 - Documentos');
  const pastaSolicitacoes = obterOuCriarPasta_(pastaRaiz, '10 - Solicitações');
  const pastaHolerites = obterOuCriarPasta_(pastaRaiz, '11 - Holerites');
  const pastaCurriculos = obterOuCriarPasta_(pastaRaiz, '12 - Currículos');
  const pastaSeguranca = obterOuCriarPasta_(pastaRaiz, '13 - Segurança do Trabalho');

  const ss = obterOuCriarPlanilha_(pastaBanco);
  prepararAbas_(ss);

  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    PLANILHA_ID: ss.getId(),
    PASTA_RAIZ_ID: pastaRaiz.getId(),
    PASTA_ATESTADOS_ID: pastaAtestados.getId(),
    PASTA_FERIAS_ID: pastaFerias.getId(),
    PASTA_COMUNICADOS_ID: pastaComunicados.getId(),
    PASTA_DOCUMENTOS_ID: pastaDocumentos.getId(),
    PASTA_SOLICITACOES_ID: pastaSolicitacoes.getId(),
    PASTA_HOLERITES_ID: pastaHolerites.getId(),
    PASTA_CURRICULOS_ID: pastaCurriculos.getId(),
    PASTA_SEGURANCA_ID: pastaSeguranca.getId()
  });

  const acessoInicial = criarAdministradorInicial_(ss);

  Logger.log('ACESSO INICIAL DO ADMINISTRADOR');
  Logger.log('E-mail: ' + acessoInicial.email);
  Logger.log('PIN temporário: ' + acessoInicial.pin);

  return {
    sucesso: true,
    planilhaUrl: ss.getUrl(),
    pastaUrl: pastaRaiz.getUrl(),
    administrador: acessoInicial
  };
}

/* ==========================================================
   AUTENTICAÇÃO E SESSÃO
========================================================== */

function login(identificador, pin) {
  validarInstalacao_();

  identificador = String(identificador || '').trim().toLowerCase();
  pin = String(pin || '').trim();

  if (!identificador) throw new Error('Informe o e-mail ou a matrícula.');
  if (!pin) throw new Error('Informe o PIN.');

  const usuario = buscarUsuarioPorLogin_(identificador);

  if (!usuario) throw new Error('Usuário não cadastrado.');
  if (String(usuario.ativo).toUpperCase() !== 'SIM') {
    throw new Error('Usuário inativo. Procure o RH.');
  }

  if (usuario.pinHash !== gerarHash_(pin)) {
    throw new Error('PIN incorreto.');
  }

  const token = Utilities.getUuid();
  const sessao = {
    email: usuario.email,
    nome: usuario.nome,
    matricula: usuario.matricula,
    perfil: usuario.perfil,
    trocarPin: String(usuario.trocarPin).toUpperCase() === 'SIM'
  };

  CacheService.getScriptCache().put(
    'SESSAO_' + token,
    JSON.stringify(sessao),
    CONFIG.DURACAO_SESSAO_SEGUNDOS
  );

  atualizarUltimoAcesso_(usuario.linha);

  return {
    token,
    usuario: sessao
  };
}

function logout(token) {
  if (token) {
    CacheService.getScriptCache().remove('SESSAO_' + token);
  }
  return true;
}

function obterSessao(token) {
  try {
    return exigirSessao_(token);
  } catch (e) {
    return null;
  }
}

function alterarMeuPin(token, pinAtual, novoPin) {
  const sessao = exigirSessao_(token);

  pinAtual = String(pinAtual || '').trim();
  novoPin = String(novoPin || '').trim();

  if (!/^\d{6}$/.test(novoPin)) {
    throw new Error('O novo PIN deve conter exatamente 6 números.');
  }

  const usuario = buscarUsuarioPorLogin_(sessao.email || sessao.matricula);
  if (!usuario) throw new Error('Usuário não encontrado.');

  if (usuario.pinHash !== gerarHash_(pinAtual)) {
    throw new Error('PIN atual incorreto.');
  }

  const aba = abrirBanco_().getSheetByName('USUARIOS');
  aba.getRange(usuario.linha, 6).setValue(gerarHash_(novoPin));
  aba.getRange(usuario.linha, 8).setValue('NÃO');

  CacheService.getScriptCache().remove('SESSAO_' + token);

  return {
    sucesso: true,
    mensagem: 'PIN alterado. Entre novamente com o novo PIN.'
  };
}

function exigirSessao_(token) {
  if (!token) throw new Error('Sessão não informada.');

  const valor = CacheService.getScriptCache().get('SESSAO_' + token);
  if (!valor) throw new Error('Sessão expirada. Entre novamente.');

  return JSON.parse(valor);
}

function exigirPerfil_(token, perfisPermitidos) {
  const sessao = exigirSessao_(token);

  if (!perfisPermitidos.includes(sessao.perfil)) {
    throw new Error('Você não possui permissão para esta ação.');
  }

  return sessao;
}

/* ==========================================================
   ADMINISTRAÇÃO DE USUÁRIOS
========================================================== */

function cadastrarUsuario(token, dados) {
  exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);

  if (!dados) throw new Error('Dados não recebidos.');

  const email = String(dados.email || '').trim().toLowerCase();
  const nome = String(dados.nome || '').trim();
  const matricula = String(dados.matricula || '').trim();
  const perfil = String(dados.perfil || '').trim().toUpperCase();

  if (!nome) throw new Error('Informe o nome.');
  if (!matricula) throw new Error('Informe a matrícula.');
  if (!['COLABORADOR', 'LIDER', 'RH', 'ADMINISTRADOR'].includes(perfil)) {
    throw new Error('Perfil inválido.');
  }

  if (buscarUsuarioPorLogin_(matricula) || (email && buscarUsuarioPorLogin_(email))) {
    throw new Error('Já existe um usuário com esta matrícula ou e-mail.');
  }

  const pinTemporario = CONFIG.PIN_INICIAL;

  abrirBanco_().getSheetByName('USUARIOS').appendRow([
    Utilities.getUuid(),
    email,
    matricula,
    nome,
    perfil,
    gerarHash_(pinTemporario),
    'SIM',
    'SIM',
    '',
    new Date()
  ]);

  return {
    sucesso: true,
    pinTemporario,
    mensagem: 'Usuário cadastrado com sucesso.'
  };
}

function listarUsuarios(token) {
  exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);

  const aba = abrirBanco_().getSheetByName('USUARIOS');
  if (aba.getLastRow() < 2) return [];

  return aba.getRange(2, 1, aba.getLastRow() - 1, 10)
    .getValues()
    .map((linha, indice) => ({
      linha: indice + 2,
      id: linha[0],
      email: linha[1],
      matricula: linha[2],
      nome: linha[3],
      perfil: linha[4],
      ativo: linha[6],
      trocarPin: linha[7],
      ultimoAcesso: formatarData_(linha[8]),
      criadoEm: formatarData_(linha[9])
    }));
}

function alterarStatusUsuario(token, email, ativo) {
  exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);

  const usuario = buscarUsuarioPorLogin_(email);
  if (!usuario) throw new Error('Usuário não encontrado.');

  abrirBanco_().getSheetByName('USUARIOS')
    .getRange(usuario.linha, 7)
    .setValue(ativo ? 'SIM' : 'NÃO');

  return true;
}

function redefinirPinUsuario(token, email) {
  exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);

  const usuario = buscarUsuarioPorLogin_(email);
  if (!usuario) throw new Error('Usuário não encontrado.');

  const novoPin = CONFIG.PIN_INICIAL;
  const aba = abrirBanco_().getSheetByName('USUARIOS');

  aba.getRange(usuario.linha, 6).setValue(gerarHash_(novoPin));
  aba.getRange(usuario.linha, 8).setValue('SIM');

  return {
    sucesso: true,
    pinTemporario: novoPin
  };
}

/* ==========================================================
   ATESTADOS
========================================================== */

function salvarAtestado(token, dados) {
  const sessao = exigirSessao_(token);
  validarDadosBasicos_(dados);
  if (!dados.tipoDocumento) throw new Error('Informe o tipo de documento.');
  validarArquivo_(dados, true);

  const protocolo = gerarProtocolo_('ATE');
  const pasta = DriveApp.getFolderById(
    PropertiesService.getScriptProperties().getProperty('PASTA_ATESTADOS_ID')
  );
  const arquivo = salvarArquivo_(pasta, dados, protocolo, dados.nome);

  const ss = abrirBanco_();
  ss.getSheetByName('ATESTADOS').appendRow([
    protocolo, new Date(), dados.matricula, dados.nome, dados.setor || '',
    dados.tipoDocumento, dados.dataEmissao || '', dados.dataInicial || '',
    dados.quantidadeDias || '', dados.quantidadeHoras || '',
    dados.observacao || '', arquivo.getId(), arquivo.getUrl(),
    'RECEBIDO', '', ''
  ]);

  registrarHistorico_(ss, dados, protocolo, 'ENVIO', 'ATESTADOS',
    'Documento enviado por ' + sessao.email + '.');

  enviarAvisoRH_(
    '[Portal RH] Novo atestado - ' + protocolo,
    [
      'Novo documento recebido.',
      '',
      'Protocolo: ' + protocolo,
      'Colaborador: ' + dados.nome,
      'Matrícula: ' + dados.matricula,
      'Setor: ' + (dados.setor || ''),
      'Tipo: ' + dados.tipoDocumento
    ].join('\n')
  );

  return { sucesso: true, protocolo, mensagem: 'Documento enviado com sucesso.' };
}

/* ==========================================================
   FÉRIAS
========================================================== */

function salvarFerias(token, dados) {
  const sessao = exigirSessao_(token);
  validarDadosBasicos_(dados);

  if (!dados.setor) throw new Error('Informe o setor.');
  if (!dados.lider) throw new Error('Informe o nome do líder.');
  if (!dados.dataInicial) throw new Error('Informe a data inicial desejada.');
  if (!dados.quantidadeDias) throw new Error('Informe a quantidade de dias.');
  if (!dados.venda10Dias) throw new Error('Informe se deseja vender 10 dias.');
  if (!dados.adiantamento13) throw new Error('Informe se deseja adiantamento do 13º.');

  validarArquivo_(dados, false);

  const protocolo = gerarProtocolo_('FER');
  const pasta = DriveApp.getFolderById(
    PropertiesService.getScriptProperties().getProperty('PASTA_FERIAS_ID')
  );
  let arquivoId = '';
  let arquivoUrl = '';
  if (dados.arquivoBase64) {
    const arquivo = salvarArquivo_(pasta, dados, protocolo, dados.nome);
    arquivoId = arquivo.getId();
    arquivoUrl = arquivo.getUrl();
  }
  const dataFinal = calcularDataFinalFerias_(dados.dataInicial, Number(dados.quantidadeDias));

  const ss = abrirBanco_();
  ss.getSheetByName('FERIAS').appendRow([
    protocolo, new Date(), dados.matricula, dados.nome, dados.setor,
    dados.lider, dados.dataInicial, dados.quantidadeDias, dataFinal,
    dados.venda10Dias, dados.adiantamento13, arquivoId, arquivoUrl,
    'AGUARDANDO_ANALISE_RH', dados.observacao || '', '', ''
  ]);

  registrarHistorico_(ss, dados, protocolo, 'SOLICITACAO', 'FERIAS',
    'Solicitação enviada por ' + sessao.email + '.');

  enviarAvisoRH_(
    '[Portal RH] Nova solicitação de férias - ' + protocolo,
    [
      'Nova solicitação de férias recebida.',
      '',
      'Protocolo: ' + protocolo,
      'Colaborador: ' + dados.nome,
      'Matrícula: ' + dados.matricula,
      'Setor: ' + dados.setor,
      'Líder: ' + dados.lider,
      'Data inicial: ' + dados.dataInicial,
      'Quantidade de dias: ' + dados.quantidadeDias
    ].join('\n')
  );

  return { sucesso: true, protocolo, mensagem: 'Solicitação de férias enviada com sucesso.' };
}

/* ==========================================================
   SOLICITAÇÕES
========================================================== */

function salvarSolicitacao(token, dados) {
  const sessao = exigirSessao_(token);
  validarDadosBasicos_(dados);

  if (!dados.setor) throw new Error('Informe o setor.');
  if (!dados.tipoSolicitacao) throw new Error('Informe o tipo de solicitação.');
  if (!dados.descricao) throw new Error('Descreva a solicitação.');
  if (!dados.urgencia) throw new Error('Informe a urgência.');

  validarArquivo_(dados, false);

  const protocolo = gerarProtocolo_('SOL');
  let arquivoId = '';
  let arquivoUrl = '';

  if (dados.arquivoBase64) {
    const pasta = DriveApp.getFolderById(
      PropertiesService.getScriptProperties().getProperty('PASTA_SOLICITACOES_ID')
    );
    const arquivo = salvarArquivo_(pasta, dados, protocolo, dados.nome);
    arquivoId = arquivo.getId();
    arquivoUrl = arquivo.getUrl();
  }

  const ss = abrirBanco_();
  ss.getSheetByName('SOLICITACOES').appendRow([
    protocolo, new Date(), dados.matricula, dados.nome, dados.setor,
    dados.tipoSolicitacao, dados.descricao, arquivoId, arquivoUrl,
    dados.urgencia, 'ABERTA', '', '', ''
  ]);

  registrarHistorico_(ss, dados, protocolo, 'SOLICITACAO', 'SOLICITACOES',
    'Solicitação enviada por ' + sessao.email + '.');

  enviarAvisoRH_(
    '[Portal RH] Nova solicitação - ' + protocolo,
    [
      'Nova solicitação recebida.',
      '',
      'Protocolo: ' + protocolo,
      'Colaborador: ' + dados.nome,
      'Matrícula: ' + dados.matricula,
      'Setor: ' + dados.setor,
      'Tipo: ' + dados.tipoSolicitacao,
      'Urgência: ' + dados.urgencia,
      '',
      'Descrição:',
      dados.descricao
    ].join('\n')
  );

  return { sucesso: true, protocolo, mensagem: 'Solicitação enviada com sucesso.' };
}

/* ==========================================================
   COMUNICADOS E DOCUMENTOS
========================================================== */

function salvarComunicado(token, dados) {
  const sessao = exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);

  if (!dados) throw new Error('Dados não recebidos.');
  if (!dados.titulo) throw new Error('Informe o título.');
  if (!dados.categoria) throw new Error('Informe a categoria.');
  if (!dados.resumo) throw new Error('Informe o resumo.');
  if (!dados.descricao) throw new Error('Informe a descrição.');
  if (!dados.publico) throw new Error('Informe o público.');
  if (!dados.publicado) throw new Error('Informe se o comunicado está publicado.');

  validarArquivo_(dados, false);

  const id = gerarProtocolo_('COM');
  let arquivoId = '';
  let arquivoUrl = '';

  if (dados.arquivoBase64) {
    const pasta = DriveApp.getFolderById(
      PropertiesService.getScriptProperties().getProperty('PASTA_COMUNICADOS_ID')
    );
    const arquivo = salvarArquivo_(pasta, dados, id, dados.titulo);
    arquivoId = arquivo.getId();
    arquivoUrl = arquivo.getUrl();
  }

  abrirBanco_().getSheetByName('COMUNICADOS').appendRow([
    id, new Date(), dados.titulo, dados.categoria, dados.resumo,
    dados.descricao, dados.publico, dados.dataValidade || '',
    dados.publicado, dados.exigeConfirmacao || 'Não',
    arquivoId, arquivoUrl, sessao.email
  ]);

  return { sucesso: true, protocolo: id, mensagem: 'Comunicado cadastrado com sucesso.' };
}

function listarComunicadosPublicados(token) {
  exigirSessao_(token);

  const aba = abrirBanco_().getSheetByName('COMUNICADOS');
  if (aba.getLastRow() < 2) return [];

  return aba.getRange(2, 1, aba.getLastRow() - 1, 13)
    .getValues()
    .filter(linha => String(linha[8]).toUpperCase() === 'SIM')
    .filter(linha => {
      if (!linha[7]) return true;
      const validade = new Date(linha[7]);
      validade.setHours(23, 59, 59, 999);
      return validade >= new Date();
    })
    .reverse()
    .map(linha => ({
      id: linha[0],
      data: formatarDataSimples_(linha[1]),
      titulo: linha[2],
      categoria: linha[3],
      resumo: linha[4],
      descricao: linha[5],
      publico: linha[6],
      exigeConfirmacao: linha[9],
      arquivoUrl: linha[11]
    }));
}

function salvarDocumento(token, dados) {
  exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);

  if (!dados) throw new Error('Dados não recebidos.');
  if (!dados.nomeDocumento) throw new Error('Informe o nome do documento.');
  if (!dados.categoria) throw new Error('Informe a categoria.');
  if (!dados.descricao) throw new Error('Informe a descrição.');
  if (!dados.publico) throw new Error('Informe o público.');

  validarArquivo_(dados, true);

  const id = gerarProtocolo_('DOC');
  const pasta = DriveApp.getFolderById(
    PropertiesService.getScriptProperties().getProperty('PASTA_DOCUMENTOS_ID')
  );
  const arquivo = salvarArquivo_(pasta, dados, id, dados.nomeDocumento);

  abrirBanco_().getSheetByName('DOCUMENTOS').appendRow([
    id, dados.categoria, dados.nomeDocumento, dados.descricao,
    dados.versao || '1.0', new Date(), dados.publico,
    dados.obrigatorio || 'Não', dados.exigeConfirmacao || 'Não',
    'SIM', arquivo.getId(), arquivo.getUrl()
  ]);

  return { sucesso: true, protocolo: id, mensagem: 'Documento publicado com sucesso.' };
}

function listarDocumentosAtivos(token) {
  exigirSessao_(token);

  const aba = abrirBanco_().getSheetByName('DOCUMENTOS');
  if (aba.getLastRow() < 2) return [];

  return aba.getRange(2, 1, aba.getLastRow() - 1, 12)
    .getValues()
    .filter(linha => String(linha[9]).toUpperCase() === 'SIM')
    .reverse()
    .map(linha => ({
      id: linha[0],
      categoria: linha[1],
      nome: linha[2],
      descricao: linha[3],
      versao: linha[4],
      data: formatarDataSimples_(linha[5]),
      publico: linha[6],
      obrigatorio: linha[7],
      arquivoUrl: linha[11]
    }));
}

/* ==========================================================
   INSTALAÇÃO E BANCO
========================================================== */

function prepararAbas_(ss) {
  const abas = {
    CONFIGURACOES: ['CHAVE', 'VALOR'],
    USUARIOS: [
      'ID', 'EMAIL', 'MATRICULA', 'NOME', 'PERFIL',
      'PIN_HASH', 'ATIVO', 'TROCAR_PIN', 'ULTIMO_ACESSO', 'CRIADO_EM'
    ],
    COLABORADORES: [
      'ID', 'MATRICULA', 'NOME', 'CPF', 'EMAIL', 'TELEFONE',
      'DATA_NASCIMENTO', 'DATA_ADMISSAO', 'CARGO', 'SETOR',
      'LIDER', 'STATUS', 'PERFIL'
    ],
    ATESTADOS: [
      'PROTOCOLO', 'DATA_ENVIO', 'MATRICULA', 'COLABORADOR', 'SETOR',
      'TIPO_DOCUMENTO', 'DATA_EMISSAO', 'DATA_INICIAL',
      'QUANTIDADE_DIAS', 'QUANTIDADE_HORAS', 'OBSERVACAO',
      'ARQUIVO_ID', 'ARQUIVO_URL', 'STATUS',
      'RESPONSAVEL_ANALISE', 'OBSERVACAO_RH'
    ],
    FERIAS: [
      'PROTOCOLO', 'DATA_SOLICITACAO', 'MATRICULA', 'COLABORADOR',
      'SETOR', 'LIDER', 'DATA_INICIAL', 'QUANTIDADE_DIAS',
      'DATA_FINAL_CALCULADA', 'VENDA_10_DIAS', 'ADIANTAMENTO_13',
      'AUTORIZACAO_ARQUIVO_ID', 'AUTORIZACAO_URL', 'STATUS',
      'OBSERVACAO_COLABORADOR', 'RESPOSTA_RH', 'DATA_RESPOSTA'
    ],
    SOLICITACOES: [
      'PROTOCOLO', 'DATA', 'MATRICULA', 'COLABORADOR', 'SETOR',
      'TIPO_SOLICITACAO', 'DESCRICAO', 'ARQUIVO_ID', 'ARQUIVO_URL',
      'URGENCIA', 'STATUS', 'RESPONSAVEL', 'RESPOSTA', 'DATA_CONCLUSAO'
    ],
    COMUNICADOS: [
      'ID', 'DATA_PUBLICACAO', 'TITULO', 'CATEGORIA', 'RESUMO',
      'DESCRICAO', 'PUBLICO', 'DATA_VALIDADE', 'PUBLICADO',
      'EXIGE_CONFIRMACAO', 'ARQUIVO_ID', 'ARQUIVO_URL', 'AUTOR'
    ],
    DOCUMENTOS: [
      'ID', 'CATEGORIA', 'NOME_DOCUMENTO', 'DESCRICAO', 'VERSAO',
      'DATA_PUBLICACAO', 'PUBLICO', 'OBRIGATORIO',
      'EXIGE_CONFIRMACAO', 'ATIVO', 'ARQUIVO_ID', 'ARQUIVO_URL'
    ],
    AVALIACOES: [
      'ID', 'DATA', 'TIPO', 'MATRICULA', 'COLABORADOR',
      'SETOR', 'AVALIADOR', 'NOTA_GERAL', 'CLASSIFICACAO',
      'PONTOS_FORTES', 'PONTOS_MELHORIA', 'PLANO_ACAO',
      'CARGO', 'CICLO', 'DATA_PROXIMA_AVALIACAO', 'STATUS'
    ],
    RESPOSTAS_AVALIACOES: [
      'AVALIACAO_ID', 'ORDEM', 'PERGUNTA', 'NOTA', 'OBSERVACAO'
    ],
    INTEGRACOES: [
      'PROTOCOLO', 'DATA_INICIO', 'MATRICULA', 'COLABORADOR',
      'SETOR', 'CARGO', 'DATA_ADMISSAO', 'RESPONSAVEL_CADASTRO',
      'STATUS', 'DATA_CONCLUSAO', 'TREINAMENTO_CONCLUIDO',
      'ACEITE_REGULAMENTO', 'ACEITE_SEGURANCA',
      'ACEITE_PRIVACIDADE', 'CODIGO_CERTIFICADO'
    ],
    TREINAMENTOS: [
      'ID', 'TITULO', 'TIPO', 'DESCRICAO', 'CARGA_HORARIA',
      'VALIDADE_MESES', 'OBRIGATORIO', 'SETORES', 'CARGOS',
      'STATUS', 'CRIADO_EM', 'CRIADO_POR'
    ],
    TREINAMENTOS_CONCLUIDOS: [
      'PROTOCOLO', 'TREINAMENTO_ID', 'TREINAMENTO', 'MATRICULA',
      'COLABORADOR', 'SETOR', 'CARGO', 'DATA_CONCLUSAO',
      'DATA_VENCIMENTO', 'STATUS_VENCIMENTO', 'NOTA', 'INSTRUTOR',
      'CODIGO_CERTIFICADO', 'ARQUIVO_URL', 'REGISTRADO_EM',
      'REGISTRADO_POR', 'OBSERVACAO'
    ],
    DOCUMENTOS_COLABORADOR: [
      'PROTOCOLO', 'MATRICULA', 'COLABORADOR', 'TIPO',
      'DESCRICAO', 'DATA_EMISSAO', 'DATA_VALIDADE', 'ARQUIVO_URL',
      'REGISTRADO_EM', 'REGISTRADO_POR'
    ],
    EPIS: [
      'PROTOCOLO', 'MATRICULA', 'COLABORADOR', 'EPI', 'CA',
      'QUANTIDADE', 'DATA_ENTREGA', 'DATA_DEVOLUCAO', 'STATUS',
      'TERMO_URL', 'REGISTRADO_EM', 'REGISTRADO_POR', 'OBSERVACAO'
    ],
    OCORRENCIAS_COLABORADOR: [
      'PROTOCOLO', 'MATRICULA', 'COLABORADOR', 'DATA_OCORRENCIA',
      'TIPO', 'TITULO', 'DESCRICAO', 'ACAO_TOMADA', 'STATUS',
      'ARQUIVO_URL', 'REGISTRADO_EM', 'REGISTRADO_POR'
    ],
    HOLERITES: [
      'ID', 'COMPETENCIA', 'MATRICULA', 'COLABORADOR', 'ARQUIVO_ID',
      'ARQUIVO_URL', 'PUBLICADO_EM', 'PUBLICADO_POR', 'VISUALIZADO_EM', 'CONFIRMADO'
    ],
    SEGURANCA_TRABALHO: [
      'PROTOCOLO', 'DATA', 'MATRICULA', 'COLABORADOR', 'SETOR', 'TIPO',
      'DESCRICAO', 'LOCAL', 'URGENCIA', 'ARQUIVO_ID', 'ARQUIVO_URL',
      'STATUS', 'RESPONSAVEL', 'RESPOSTA', 'DATA_CONCLUSAO'
    ],
    CURRICULOS: [
      'PROTOCOLO', 'DATA', 'NOME', 'TELEFONE', 'EMAIL', 'CIDADE',
      'AREA_INTERESSE', 'EXPERIENCIA', 'ARQUIVO_ID', 'ARQUIVO_URL', 'STATUS', 'OBSERVACAO_RH'
    ],
    CARDAPIO: [
      'ID', 'SEMANA_INICIO', 'DIA_SEMANA', 'DATA', 'PRATO_PRINCIPAL',
      'ACOMPANHAMENTOS', 'SALADA', 'SOBREMESA', 'OBSERVACAO', 'PUBLICADO', 'ATUALIZADO_POR'
    ],
    HISTORICO: [
      'DATA', 'HORA', 'MATRICULA', 'COLABORADOR',
      'ACAO', 'MODULO', 'PROTOCOLO', 'DESCRICAO'
    ]
  };

  const primeira = ss.getSheets()[0];
  if (primeira.getName() !== 'CONFIGURACOES' && !ss.getSheetByName('CONFIGURACOES')) {
    primeira.setName('CONFIGURACOES');
  }

  Object.entries(abas).forEach(([nome, cabecalhos]) => {
    let aba = ss.getSheetByName(nome);
    if (!aba) aba = ss.insertSheet(nome);

    if (aba.getLastRow() <= 1) {
      aba.clear();
      aba.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
    } else {
      aba.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
    }

    aba.getRange(1, 1, 1, cabecalhos.length)
      .setBackground('#123B63')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');

    aba.setFrozenRows(1);
    aba.autoResizeColumns(1, cabecalhos.length);
  });

  const config = ss.getSheetByName('CONFIGURACOES');
  if (config.getLastRow() < 2) {
    config.getRange(2, 1, 3, 2).setValues([
      ['NOME_EMPRESA', CONFIG.EMPRESA],
      ['EMAIL_RH', CONFIG.EMAIL_RH],
      ['DATA_INSTALACAO', new Date()]
    ]);
  }
}

function criarAdministradorInicial_(ss) {
  return configurarAdministradorPrincipal_(ss);
}

function configurarAdministradorPrincipal_(ss) {
  const aba = ss.getSheetByName('USUARIOS');
  const email = CONFIG.EMAIL_ADMINISTRADOR.toLowerCase();
  const pin = CONFIG.PIN_INICIAL;
  let linha = 0;

  if (aba.getLastRow() >= 2) {
    const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 10).getValues();
    for (let i = 0; i < dados.length; i++) {
      if (String(dados[i][1]).trim().toLowerCase() === email) {
        linha = i + 2;
        break;
      }
    }
  }

  const valores = [
    linha ? aba.getRange(linha, 1).getValue() : Utilities.getUuid(),
    email,
    'ADMIN',
    'Alexandra Rodrigues',
    'ADMINISTRADOR',
    gerarHash_(pin),
    'SIM',
    'SIM',
    '',
    linha ? aba.getRange(linha, 10).getValue() || new Date() : new Date()
  ];

  if (linha) {
    aba.getRange(linha, 1, 1, 10).setValues([valores]);
  } else {
    aba.appendRow(valores);
    linha = aba.getLastRow();
  }

  CacheService.getScriptCache().removeAll(['SESSAO_ADMIN_RESET']);

  return {
    email: email,
    pin: pin,
    perfil: 'ADMINISTRADOR',
    trocarPin: 'SIM',
    linha: linha
  };
}

function resetarAdministradorPrincipal() {
  validarInstalacao_();
  const ss = abrirBanco_();
  prepararAbas_(ss);
  const acesso = configurarAdministradorPrincipal_(ss);

  Logger.log('ADMINISTRADOR PRINCIPAL REDEFINIDO');
  Logger.log('E-mail: ' + acesso.email);
  Logger.log('PIN temporário: ' + acesso.pin);

  return {
    sucesso: true,
    email: acesso.email,
    pinTemporario: acesso.pin,
    mensagem: 'Administrador redefinido. Entre com o PIN 123456 e altere-o no primeiro acesso.'
  };
}

function buscarUsuarioPorEmail_(email) {
  const aba = abrirBanco_().getSheetByName('USUARIOS');
  if (aba.getLastRow() < 2) return null;

  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 10).getValues();
  const emailBusca = String(email || '').trim().toLowerCase();

  for (let i = 0; i < dados.length; i++) {
    const linha = dados[i];
    if (String(linha[1]).trim().toLowerCase() === emailBusca) {
      return {
        linha: i + 2,
        id: linha[0],
        email: linha[1],
        matricula: linha[2],
        nome: linha[3],
        perfil: linha[4],
        pinHash: linha[5],
        ativo: linha[6],
        trocarPin: linha[7],
        ultimoAcesso: linha[8],
        criadoEm: linha[9]
      };
    }
  }

  return null;
}


function buscarUsuarioPorLogin_(identificador) {
  const aba = abrirBanco_().getSheetByName('USUARIOS');
  if (!aba || aba.getLastRow() < 2) return null;

  const busca = String(identificador || '').trim().toLowerCase();
  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 10).getValues();

  for (let i = 0; i < dados.length; i++) {
    const linha = dados[i];
    const email = String(linha[1] || '').trim().toLowerCase();
    const matricula = String(linha[2] || '').trim().toLowerCase();

    if (email === busca || matricula === busca) {
      return {
        linha: i + 2,
        id: linha[0],
        email: linha[1],
        matricula: linha[2],
        nome: linha[3],
        perfil: linha[4],
        pinHash: linha[5],
        ativo: linha[6],
        trocarPin: linha[7],
        ultimoAcesso: linha[8],
        criadoEm: linha[9]
      };
    }
  }

  return null;
}

function sincronizarUsuariosColaboradores(token, redefinirPins) {
  exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);

  const ss = abrirBanco_();
  const abaColaboradores = ss.getSheetByName('COLABORADORES');
  const abaUsuarios = ss.getSheetByName('USUARIOS');

  if (!abaColaboradores || abaColaboradores.getLastRow() < 2) {
    throw new Error('A aba COLABORADORES está vazia. Preencha os colaboradores antes de sincronizar.');
  }

  const cabecalhos = abaColaboradores.getRange(1, 1, 1, abaColaboradores.getLastColumn())
    .getValues()[0]
    .map(v => String(v || '').trim().toUpperCase());

  const indice = nome => cabecalhos.indexOf(nome);
  const iMatricula = indice('MATRICULA');
  const iNome = indice('NOME');
  const iEmail = indice('EMAIL');
  const iStatus = indice('STATUS');
  const iPerfil = indice('PERFIL');

  if (iMatricula < 0 || iNome < 0) {
    throw new Error('A aba COLABORADORES precisa conter as colunas MATRICULA e NOME.');
  }

  const colaboradores = abaColaboradores
    .getRange(2, 1, abaColaboradores.getLastRow() - 1, abaColaboradores.getLastColumn())
    .getValues();

  const usuarios = abaUsuarios.getLastRow() >= 2
    ? abaUsuarios.getRange(2, 1, abaUsuarios.getLastRow() - 1, 10).getValues()
    : [];

  const porMatricula = {};
  const porEmail = {};
  usuarios.forEach((u, pos) => {
    const matricula = String(u[2] || '').trim().toLowerCase();
    const email = String(u[1] || '').trim().toLowerCase();
    if (matricula) porMatricula[matricula] = pos + 2;
    if (email) porEmail[email] = pos + 2;
  });

  let criados = 0;
  let atualizados = 0;
  let ignorados = 0;
  let pinsRedefinidos = 0;

  colaboradores.forEach(c => {
    const matricula = String(c[iMatricula] || '').trim();
    const nome = String(c[iNome] || '').trim();
    const email = iEmail >= 0 ? String(c[iEmail] || '').trim().toLowerCase() : '';
    const status = iStatus >= 0 ? String(c[iStatus] || 'ATIVO').trim().toUpperCase() : 'ATIVO';
    let perfil = iPerfil >= 0 ? String(c[iPerfil] || 'COLABORADOR').trim().toUpperCase() : 'COLABORADOR';

    if (!['COLABORADOR', 'LIDER', 'RH', 'ADMINISTRADOR'].includes(perfil)) perfil = 'COLABORADOR';
    if (!matricula || !nome || ['INATIVO', 'DESLIGADO', 'DEMITIDO'].includes(status)) {
      ignorados++;
      return;
    }

    const linhaExistente = porMatricula[matricula.toLowerCase()] || (email ? porEmail[email] : 0);

    if (linhaExistente) {
      const atual = abaUsuarios.getRange(linhaExistente, 1, 1, 10).getValues()[0];
      atual[1] = email || atual[1] || '';
      atual[2] = matricula;
      atual[3] = nome;
      atual[4] = perfil;
      atual[6] = 'SIM';
      if (redefinirPins) {
        atual[5] = gerarHash_(CONFIG.PIN_INICIAL);
        atual[7] = 'SIM';
        pinsRedefinidos++;
      }
      abaUsuarios.getRange(linhaExistente, 1, 1, 10).setValues([atual]);
      atualizados++;
    } else {
      abaUsuarios.appendRow([
        Utilities.getUuid(), email, matricula, nome, perfil,
        gerarHash_(CONFIG.PIN_INICIAL), 'SIM', 'SIM', '', new Date()
      ]);
      const novaLinha = abaUsuarios.getLastRow();
      porMatricula[matricula.toLowerCase()] = novaLinha;
      if (email) porEmail[email] = novaLinha;
      criados++;
    }
  });

  return {
    sucesso: true,
    criados,
    atualizados,
    ignorados,
    pinsRedefinidos,
    pinInicial: CONFIG.PIN_INICIAL,
    mensagem: 'Sincronização concluída.'
  };
}

function atualizarUltimoAcesso_(linha) {
  abrirBanco_().getSheetByName('USUARIOS')
    .getRange(linha, 9)
    .setValue(new Date());
}

/* ==========================================================
   UTILITÁRIOS
========================================================== */

function gerarHash_(texto) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    texto,
    Utilities.Charset.UTF_8
  );

  return bytes.map(byte => {
    const valor = byte < 0 ? byte + 256 : byte;
    return ('0' + valor.toString(16)).slice(-2);
  }).join('');
}

function gerarPin_() {
  return String(Math.floor(100000 + Math.random() * 900000));
}

function validarDadosBasicos_(dados) {
  if (!dados) throw new Error('Dados não recebidos.');
  if (!dados.nome) throw new Error('Informe o nome do colaborador.');
  if (!dados.matricula) throw new Error('Informe a matrícula.');
}

function validarArquivo_(dados, obrigatorio) {
  if (obrigatorio && !dados.arquivoBase64) {
    throw new Error('Selecione um arquivo.');
  }
  if (!dados.arquivoBase64) return;

  const mimePermitidos = [
    'application/pdf',
    'image/jpeg',
    'image/png'
  ];

  if (!mimePermitidos.includes(dados.mimeType)) {
    throw new Error('Formato não permitido. Envie PDF, JPG, JPEG ou PNG.');
  }

  const bytes = Utilities.base64Decode(dados.arquivoBase64);
  if (bytes.length > 10 * 1024 * 1024) {
    throw new Error('O arquivo deve ter no máximo 10 MB.');
  }
}

function salvarArquivo_(pasta, dados, protocolo, nomeReferencia) {
  const bytes = Utilities.base64Decode(dados.arquivoBase64);
  const extensao = obterExtensao_(dados.nomeArquivo, dados.mimeType);
  const nomeArquivo = protocolo + ' - ' + sanitizarNome_(nomeReferencia) + extensao;
  return pasta.createFile(Utilities.newBlob(bytes, dados.mimeType, nomeArquivo));
}

function obterOuCriarPlanilha_(pastaBanco) {
  const props = PropertiesService.getScriptProperties();
  const idExistente = props.getProperty('PLANILHA_ID');

  if (idExistente) {
    try {
      return SpreadsheetApp.openById(idExistente);
    } catch (e) {}
  }

  const arquivos = pastaBanco.getFilesByName(CONFIG.PLANILHA);
  if (arquivos.hasNext()) {
    return SpreadsheetApp.openById(arquivos.next().getId());
  }

  const ss = SpreadsheetApp.create(CONFIG.PLANILHA);
  ss.setSpreadsheetTimeZone(CONFIG.FUSO);

  const arquivo = DriveApp.getFileById(ss.getId());
  pastaBanco.addFile(arquivo);
  try {
    DriveApp.getRootFolder().removeFile(arquivo);
  } catch (e) {}

  return ss;
}

function abrirBanco_() {
  return SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty('PLANILHA_ID')
  );
}

function obterOuCriarPasta_(pastaPai, nome) {
  const pastas = pastaPai.getFoldersByName(nome);
  return pastas.hasNext() ? pastas.next() : pastaPai.createFolder(nome);
}

function registrarHistorico_(ss, dados, protocolo, acao, modulo, descricao) {
  ss.getSheetByName('HISTORICO').appendRow([
    new Date(),
    Utilities.formatDate(new Date(), CONFIG.FUSO, 'HH:mm:ss'),
    dados.matricula,
    dados.nome,
    acao,
    modulo,
    protocolo,
    descricao
  ]);
}

function enviarAvisoRH_(assunto, corpo) {
  MailApp.sendEmail({
    to: CONFIG.EMAIL_RH,
    subject: assunto,
    body: corpo
  });
}

function validarInstalacao_() {
  if (!PropertiesService.getScriptProperties().getProperty('PLANILHA_ID')) {
    throw new Error('Execute primeiro a função instalarEtapa5().');
  }
}

function gerarProtocolo_(prefixo) {
  const data = Utilities.formatDate(new Date(), CONFIG.FUSO, 'yyyyMMddHHmmss');
  const aleatorio = Math.floor(Math.random() * 900) + 100;
  return prefixo + '-' + data + '-' + aleatorio;
}

function sanitizarNome_(texto) {
  return String(texto)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Za-z0-9 ]/g, '')
    .trim()
    .replace(/\s+/g, '_');
}

function obterExtensao_(nomeArquivo, mimeType) {
  const nome = String(nomeArquivo || '');
  const posicao = nome.lastIndexOf('.');
  if (posicao >= 0) return nome.substring(posicao).toLowerCase();

  return {
    'application/pdf': '.pdf',
    'image/jpeg': '.jpg',
    'image/png': '.png'
  }[mimeType] || '';
}

function calcularDataFinalFerias_(dataInicial, dias) {
  const data = new Date(dataInicial + 'T12:00:00');
  data.setDate(data.getDate() + dias - 1);
  return Utilities.formatDate(data, CONFIG.FUSO, 'yyyy-MM-dd');
}

function formatarData_(valor) {
  if (!valor) return '';
  return Utilities.formatDate(new Date(valor), CONFIG.FUSO, 'dd/MM/yyyy HH:mm');
}

function formatarDataSimples_(valor) {
  if (!valor) return '';
  return Utilities.formatDate(new Date(valor), CONFIG.FUSO, 'dd/MM/yyyy');
}
/* ==========================================================
   ETAPA 6 - DASHBOARD E CENTRAL DE APROVAÇÕES
========================================================== */

function obterDashboard(token) {
  exigirPerfil_(token, ['RH', 'ADMINISTRADOR', 'LIDER']);

  const ss = abrirBanco_();
  const agora = new Date();
  const mes = agora.getMonth();
  const ano = agora.getFullYear();

  const colaboradores = contarLinhasAtivas_(ss.getSheetByName('COLABORADORES'), 12);
  const usuariosAtivos = contarValor_(ss.getSheetByName('USUARIOS'), 7, 'SIM');

  const atestados = lerDados_(ss.getSheetByName('ATESTADOS'), 16);
  const ferias = lerDados_(ss.getSheetByName('FERIAS'), 17);
  const solicitacoes = lerDados_(ss.getSheetByName('SOLICITACOES'), 14);
  const comunicados = lerDados_(ss.getSheetByName('COMUNICADOS'), 13);

  const atestadosMes = atestados.filter(l => mesmaCompetencia_(l[1], mes, ano)).length;
  const atestadosPendentes = atestados.filter(l =>
    ['RECEBIDO', 'PENDENTE', 'AGUARDANDO_ANALISE_RH'].includes(String(l[13]).toUpperCase())
  ).length;

  const feriasPendentes = ferias.filter(l =>
    ['AGUARDANDO_ANALISE_RH', 'PENDENTE', 'AGUARDANDO_LIDER'].includes(String(l[13]).toUpperCase())
  ).length;

  const solicitacoesAbertas = solicitacoes.filter(l =>
    !['CONCLUIDA', 'CONCLUÍDA', 'FINALIZADA', 'REPROVADA', 'CANCELADA'].includes(String(l[10]).toUpperCase())
  ).length;

  const comunicadosAtivos = comunicados.filter(l =>
    String(l[8]).toUpperCase() === 'SIM'
  ).length;

  const pendenciasTotal = atestadosPendentes + feriasPendentes + solicitacoesAbertas;

  const porTipo = [
    { tipo: 'Atestados', quantidade: atestadosPendentes },
    { tipo: 'Férias', quantidade: feriasPendentes },
    { tipo: 'Solicitações', quantidade: solicitacoesAbertas }
  ];

  const porSetor = {};
  atestados.forEach(l => somarSetor_(porSetor, l[4]));
  ferias.forEach(l => somarSetor_(porSetor, l[4]));
  solicitacoes.forEach(l => somarSetor_(porSetor, l[4]));

  const setores = Object.keys(porSetor)
    .map(setor => ({ setor, quantidade: porSetor[setor] }))
    .sort((a, b) => b.quantidade - a.quantidade)
    .slice(0, 10);

  return {
    indicadores: {
      colaboradores,
      usuariosAtivos,
      atestadosMes,
      atestadosPendentes,
      feriasPendentes,
      solicitacoesAbertas,
      comunicadosAtivos,
      pendenciasTotal
    },
    porTipo,
    setores
  };
}

function listarPendencias(token) {
  exigirPerfil_(token, ['RH', 'ADMINISTRADOR', 'LIDER']);

  const ss = abrirBanco_();
  const lista = [];

  const atestados = lerDados_(ss.getSheetByName('ATESTADOS'), 16);
  atestados.forEach((l, i) => {
    const status = String(l[13]).toUpperCase();
    if (['RECEBIDO', 'PENDENTE', 'AGUARDANDO_ANALISE_RH'].includes(status)) {
      lista.push({
        modulo: 'ATESTADO',
        linha: i + 2,
        protocolo: l[0],
        data: formatarData_(l[1]),
        matricula: l[2],
        colaborador: l[3],
        setor: l[4],
        tipo: l[5],
        status: l[13],
        descricao: l[10] || '',
        arquivoUrl: l[12] || ''
      });
    }
  });

  const ferias = lerDados_(ss.getSheetByName('FERIAS'), 17);
  ferias.forEach((l, i) => {
    const status = String(l[13]).toUpperCase();
    if (['AGUARDANDO_ANALISE_RH', 'PENDENTE', 'AGUARDANDO_LIDER'].includes(status)) {
      lista.push({
        modulo: 'FERIAS',
        linha: i + 2,
        protocolo: l[0],
        data: formatarData_(l[1]),
        matricula: l[2],
        colaborador: l[3],
        setor: l[4],
        tipo: 'Férias - ' + l[7] + ' dias',
        status: l[13],
        descricao: 'Início desejado: ' + formatarDataSimples_(l[6]),
        arquivoUrl: l[12] || ''
      });
    }
  });

  const solicitacoes = lerDados_(ss.getSheetByName('SOLICITACOES'), 14);
  solicitacoes.forEach((l, i) => {
    const status = String(l[10]).toUpperCase();
    if (!['CONCLUIDA', 'CONCLUÍDA', 'FINALIZADA', 'REPROVADA', 'CANCELADA'].includes(status)) {
      lista.push({
        modulo: 'SOLICITACAO',
        linha: i + 2,
        protocolo: l[0],
        data: formatarData_(l[1]),
        matricula: l[2],
        colaborador: l[3],
        setor: l[4],
        tipo: l[5],
        status: l[10],
        descricao: l[6] || '',
        arquivoUrl: l[8] || ''
      });
    }
  });

  return lista.sort((a, b) => String(b.data).localeCompare(String(a.data)));
}

function atualizarPendencia(token, dados) {
  const sessao = exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);

  if (!dados || !dados.modulo || !dados.linha || !dados.acao) {
    throw new Error('Dados incompletos para atualizar a pendência.');
  }

  const acao = String(dados.acao).toUpperCase();
  if (!['APROVAR', 'REPROVAR', 'SOLICITAR_CORRECAO', 'CONCLUIR'].includes(acao)) {
    throw new Error('Ação inválida.');
  }

  const ss = abrirBanco_();
  let aba;
  let colunaStatus;
  let colunaResponsavel;
  let colunaResposta;
  let novoStatus;

  if (dados.modulo === 'ATESTADO') {
    aba = ss.getSheetByName('ATESTADOS');
    colunaStatus = 14;
    colunaResponsavel = 15;
    colunaResposta = 16;
    novoStatus = acao === 'APROVAR' ? 'VALIDADO' :
                 acao === 'REPROVAR' ? 'REPROVADO' :
                 acao === 'CONCLUIR' ? 'CONCLUIDO' : 'CORRECAO_SOLICITADA';
  } else if (dados.modulo === 'FERIAS') {
    aba = ss.getSheetByName('FERIAS');
    colunaStatus = 14;
    colunaResponsavel = 16;
    colunaResposta = 16;
    novoStatus = acao === 'APROVAR' ? 'APROVADA' :
                 acao === 'REPROVAR' ? 'REPROVADA' :
                 acao === 'CONCLUIR' ? 'FINALIZADA' : 'CORRECAO_SOLICITADA';
  } else if (dados.modulo === 'SOLICITACAO') {
    aba = ss.getSheetByName('SOLICITACOES');
    colunaStatus = 11;
    colunaResponsavel = 12;
    colunaResposta = 13;
    novoStatus = acao === 'APROVAR' ? 'APROVADA' :
                 acao === 'REPROVAR' ? 'REPROVADA' :
                 acao === 'CONCLUIR' ? 'CONCLUIDA' : 'CORRECAO_SOLICITADA';
  } else {
    throw new Error('Módulo inválido.');
  }

  const linha = Number(dados.linha);
  aba.getRange(linha, colunaStatus).setValue(novoStatus);
  aba.getRange(linha, colunaResponsavel).setValue(sessao.nome + ' - ' + sessao.email);
  aba.getRange(linha, colunaResposta).setValue(dados.observacao || '');

  if (dados.modulo === 'FERIAS') {
    aba.getRange(linha, 17).setValue(new Date());
  }
  if (dados.modulo === 'SOLICITACAO' && acao === 'CONCLUIR') {
    aba.getRange(linha, 14).setValue(new Date());
  }

  const protocolo = aba.getRange(linha, 1).getValue();
  const matricula = aba.getRange(linha, 3).getValue();
  const colaborador = aba.getRange(linha, 4).getValue();

  ss.getSheetByName('HISTORICO').appendRow([
    new Date(),
    Utilities.formatDate(new Date(), CONFIG.FUSO, 'HH:mm:ss'),
    matricula,
    colaborador,
    acao,
    dados.modulo,
    protocolo,
    (dados.observacao || '') + ' | Responsável: ' + sessao.email
  ]);

  return {
    sucesso: true,
    status: novoStatus,
    mensagem: 'Pendência atualizada com sucesso.'
  };
}

function contarLinhasAtivas_(aba, colunaStatus) {
  if (!aba || aba.getLastRow() < 2) return 0;
  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, aba.getLastColumn()).getValues();
  return dados.filter(l => {
    const status = String(l[colunaStatus - 1] || '').toUpperCase();
    return !['INATIVO', 'DESLIGADO', 'NÃO'].includes(status);
  }).length;
}

function contarValor_(aba, coluna, valor) {
  if (!aba || aba.getLastRow() < 2) return 0;
  return aba.getRange(2, coluna, aba.getLastRow() - 1, 1)
    .getValues()
    .filter(l => String(l[0]).toUpperCase() === valor).length;
}

function lerDados_(aba, colunas) {
  if (!aba || aba.getLastRow() < 2) return [];
  return aba.getRange(2, 1, aba.getLastRow() - 1, colunas).getValues();
}

function mesmaCompetencia_(valor, mes, ano) {
  if (!valor) return false;
  const d = new Date(valor);
  return d.getMonth() === mes && d.getFullYear() === ano;
}

function somarSetor_(objeto, setor) {
  const nome = String(setor || 'Não informado').trim() || 'Não informado';
  objeto[nome] = (objeto[nome] || 0) + 1;
}



/* ==========================================================
   ETAPA 7 - AVALIAÇÃO COMPORTAMENTAL E TÉCNICA
========================================================== */

function obterPerguntasAvaliacao(token, tipo) {
  exigirSessao_(token);

  tipo = String(tipo || '').toUpperCase();

  const comportamentais = [
    'Demonstra responsabilidade e comprometimento com suas atividades?',
    'Mantém bom relacionamento e respeito com colegas e líderes?',
    'Cumpre horários, prazos e procedimentos internos?',
    'Demonstra iniciativa para resolver problemas?',
    'Aceita orientações e feedbacks de forma profissional?',
    'Colabora com a equipe e compartilha informações?',
    'Mantém postura adequada e comunicação respeitosa?',
    'Cuida dos equipamentos, ferramentas e recursos da empresa?',
    'Segue as regras de segurança e utiliza corretamente os EPIs?',
    'Demonstra interesse em aprender e se desenvolver?'
  ];

  const tecnicas = [
    'Possui conhecimento técnico compatível com o cargo?',
    'Executa as atividades conforme desenhos, procedimentos ou orientações?',
    'Entrega o trabalho com qualidade e atenção aos detalhes?',
    'Evita retrabalhos e desperdícios?',
    'Utiliza corretamente máquinas, equipamentos e ferramentas?',
    'Consegue identificar problemas técnicos durante a execução?',
    'Organiza corretamente o posto e os materiais de trabalho?',
    'Mantém produtividade compatível com a função?',
    'Cumpre os padrões de qualidade e inspeção aplicáveis?',
    'Consegue trabalhar com autonomia dentro das responsabilidades do cargo?'
  ];

  if (tipo === 'COMPORTAMENTAL') return embaralharPerguntas_(comportamentais, 8);
  if (tipo === 'TECNICA' || tipo === 'TÉCNICA') return embaralharPerguntas_(tecnicas, 8);

  return {
    comportamental: embaralharPerguntas_(comportamentais, 8),
    tecnica: embaralharPerguntas_(tecnicas, 8)
  };
}

function salvarAvaliacao(token, dados) {
  const sessao = exigirPerfil_(token, ['LIDER', 'RH', 'ADMINISTRADOR']);

  if (!dados) throw new Error('Dados da avaliação não recebidos.');
  if (!dados.matricula) throw new Error('Informe a matrícula.');
  if (!dados.colaborador) throw new Error('Informe o colaborador.');
  if (!dados.setor) throw new Error('Informe o setor.');
  if (!dados.cargo) throw new Error('Informe o cargo.');
  if (!dados.tipo) throw new Error('Informe o tipo de avaliação.');
  if (!dados.respostas || !dados.respostas.length) {
    throw new Error('Responda às perguntas da avaliação.');
  }

  const notas = dados.respostas.map(r => Number(r.nota));
  if (notas.some(n => !n || n < 1 || n > 5)) {
    throw new Error('Todas as notas devem estar entre 1 e 5.');
  }

  const media = notas.reduce((soma, nota) => soma + nota, 0) / notas.length;
  const notaGeral = Math.round(media * 100) / 100;
  const classificacao = classificarNota_(notaGeral);
  const id = gerarProtocolo_('AVA');

  const ss = abrirBanco_();
  const aba = ss.getSheetByName('AVALIACOES');

  aba.appendRow([
    id,
    new Date(),
    String(dados.tipo).toUpperCase(),
    dados.matricula,
    dados.colaborador,
    dados.setor,
    sessao.nome + ' - ' + sessao.email,
    notaGeral,
    classificacao,
    dados.pontosFortes || '',
    dados.pontosMelhoria || '',
    dados.planoAcao || '',
    dados.cargo || '',
    dados.ciclo || '',
    dados.dataProximaAvaliacao || '',
    'FINALIZADA'
  ]);

  const abaRespostas = ss.getSheetByName('RESPOSTAS_AVALIACOES');
  dados.respostas.forEach((resposta, indice) => {
    abaRespostas.appendRow([
      id,
      indice + 1,
      resposta.pergunta,
      Number(resposta.nota),
      resposta.observacao || ''
    ]);
  });

  ss.getSheetByName('HISTORICO').appendRow([
    new Date(),
    Utilities.formatDate(new Date(), CONFIG.FUSO, 'HH:mm:ss'),
    dados.matricula,
    dados.colaborador,
    'AVALIACAO_FINALIZADA',
    'AVALIACOES',
    id,
    'Tipo: ' + dados.tipo + ' | Nota: ' + notaGeral + ' | Avaliador: ' + sessao.email
  ]);

  return {
    sucesso: true,
    protocolo: id,
    notaGeral,
    classificacao,
    mensagem: 'Avaliação registrada com sucesso.'
  };
}

function listarAvaliacoes(token, matriculaFiltro) {
  const sessao = exigirSessao_(token);
  const aba = abrirBanco_().getSheetByName('AVALIACOES');
  if (!aba || aba.getLastRow() < 2) return [];

  let dados = aba.getRange(2, 1, aba.getLastRow() - 1, 16).getValues();

  if (sessao.perfil === 'COLABORADOR') {
    dados = dados.filter(l => String(l[3]) === String(sessao.matricula));
  } else if (matriculaFiltro) {
    dados = dados.filter(l => String(l[3]) === String(matriculaFiltro));
  }

  return dados.reverse().map(l => ({
    id: l[0],
    data: formatarDataSimples_(l[1]),
    tipo: l[2],
    matricula: l[3],
    colaborador: l[4],
    setor: l[5],
    avaliador: l[6],
    notaGeral: l[7],
    classificacao: l[8],
    pontosFortes: l[9],
    pontosMelhoria: l[10],
    planoAcao: l[11],
    cargo: l[12],
    ciclo: l[13],
    proximaAvaliacao: formatarDataSimples_(l[14]),
    status: l[15]
  }));
}

function obterDetalhesAvaliacao(token, idAvaliacao) {
  const sessao = exigirSessao_(token);
  const ss = abrirBanco_();
  const avaliacoes = ss.getSheetByName('AVALIACOES');
  const respostas = ss.getSheetByName('RESPOSTAS_AVALIACOES');

  if (!avaliacoes || avaliacoes.getLastRow() < 2) {
    throw new Error('Avaliação não encontrada.');
  }

  const dados = avaliacoes.getRange(2, 1, avaliacoes.getLastRow() - 1, 16).getValues();
  const avaliacao = dados.find(l => String(l[0]) === String(idAvaliacao));

  if (!avaliacao) throw new Error('Avaliação não encontrada.');

  if (sessao.perfil === 'COLABORADOR' &&
      String(avaliacao[3]) !== String(sessao.matricula)) {
    throw new Error('Você não possui acesso a esta avaliação.');
  }

  let listaRespostas = [];
  if (respostas && respostas.getLastRow() >= 2) {
    listaRespostas = respostas
      .getRange(2, 1, respostas.getLastRow() - 1, 5)
      .getValues()
      .filter(l => String(l[0]) === String(idAvaliacao))
      .map(l => ({
        ordem: l[1],
        pergunta: l[2],
        nota: l[3],
        observacao: l[4]
      }));
  }

  return {
    id: avaliacao[0],
    data: formatarDataSimples_(avaliacao[1]),
    tipo: avaliacao[2],
    matricula: avaliacao[3],
    colaborador: avaliacao[4],
    setor: avaliacao[5],
    avaliador: avaliacao[6],
    notaGeral: avaliacao[7],
    classificacao: avaliacao[8],
    pontosFortes: avaliacao[9],
    pontosMelhoria: avaliacao[10],
    planoAcao: avaliacao[11],
    cargo: avaliacao[12],
    ciclo: avaliacao[13],
    proximaAvaliacao: formatarDataSimples_(avaliacao[14]),
    status: avaliacao[15],
    respostas: listaRespostas
  };
}

function obterResumoAvaliacoes(token) {
  exigirPerfil_(token, ['LIDER', 'RH', 'ADMINISTRADOR']);

  const aba = abrirBanco_().getSheetByName('AVALIACOES');
  if (!aba || aba.getLastRow() < 2) {
    return {
      total: 0,
      mediaGeral: 0,
      excelentes: 0,
      desenvolvimento: 0,
      porSetor: []
    };
  }

  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 16).getValues();
  const notas = dados.map(l => Number(l[7])).filter(n => !isNaN(n));
  const mediaGeral = notas.length
    ? Math.round((notas.reduce((a, b) => a + b, 0) / notas.length) * 100) / 100
    : 0;

  const porSetor = {};
  dados.forEach(l => {
    const setor = String(l[5] || 'Não informado');
    if (!porSetor[setor]) porSetor[setor] = { soma: 0, quantidade: 0 };
    porSetor[setor].soma += Number(l[7]) || 0;
    porSetor[setor].quantidade++;
  });

  return {
    total: dados.length,
    mediaGeral,
    excelentes: dados.filter(l => Number(l[7]) >= 4.5).length,
    desenvolvimento: dados.filter(l => Number(l[7]) < 3).length,
    porSetor: Object.keys(porSetor).map(setor => ({
      setor,
      media: Math.round((porSetor[setor].soma / porSetor[setor].quantidade) * 100) / 100,
      quantidade: porSetor[setor].quantidade
    })).sort((a, b) => b.media - a.media)
  };
}

function classificarNota_(nota) {
  if (nota >= 4.5) return 'EXCELENTE';
  if (nota >= 4) return 'MUITO BOM';
  if (nota >= 3) return 'ADEQUADO';
  if (nota >= 2) return 'EM DESENVOLVIMENTO';
  return 'INSATISFATORIO';
}

function embaralharPerguntas_(lista, quantidade) {
  const copia = lista.slice();

  for (let i = copia.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const temporario = copia[i];
    copia[i] = copia[j];
    copia[j] = temporario;
  }

  return copia.slice(0, quantidade).map((pergunta, indice) => ({
    ordem: indice + 1,
    pergunta
  }));
}



/* ==========================================================
   ETAPA 8 - INTEGRAÇÃO ONLINE
========================================================== */

function obterConfiguracaoIntegracao(token) {
  exigirSessao_(token);

  return {
    linkTreinamento: 'https://script.google.com/macros/s/AKfycbzSKC-hi7fimdtRRYyggTRuJM8jfPVWfy7wt2nnonSYKlvj9gfCk8dmW84VfC5ldRjZYw/exec',
    titulo: 'Integração Online - Ferramentaria Soares',
    orientacao: 'Conclua o treinamento, leia os documentos obrigatórios e registre sua conclusão no portal.'
  };
}

function iniciarIntegracao(token, dados) {
  const sessao = exigirSessao_(token);

  if (!dados) throw new Error('Dados não recebidos.');
  if (!dados.matricula) throw new Error('Informe a matrícula.');
  if (!dados.colaborador) throw new Error('Informe o colaborador.');
  if (!dados.setor) throw new Error('Informe o setor.');
  if (!dados.cargo) throw new Error('Informe o cargo.');
  if (!dados.dataAdmissao) throw new Error('Informe a data de admissão.');

  const ss = abrirBanco_();
  const aba = ss.getSheetByName('INTEGRACOES');

  const existente = localizarIntegracaoAtiva_(aba, dados.matricula);
  if (existente) {
    return {
      sucesso: true,
      protocolo: existente.protocolo,
      status: existente.status,
      mensagem: 'Já existe uma integração cadastrada para esta matrícula.'
    };
  }

  const protocolo = gerarProtocolo_('INT');

  aba.appendRow([
    protocolo,
    new Date(),
    dados.matricula,
    dados.colaborador,
    dados.setor,
    dados.cargo,
    dados.dataAdmissao,
    sessao.email,
    'INICIADA',
    '',
    '',
    '',
    '',
    '',
    ''
  ]);

  ss.getSheetByName('HISTORICO').appendRow([
    new Date(),
    Utilities.formatDate(new Date(), CONFIG.FUSO, 'HH:mm:ss'),
    dados.matricula,
    dados.colaborador,
    'INTEGRACAO_INICIADA',
    'INTEGRACOES',
    protocolo,
    'Integração iniciada por ' + sessao.email
  ]);

  return {
    sucesso: true,
    protocolo,
    status: 'INICIADA',
    mensagem: 'Integração iniciada com sucesso.'
  };
}

function concluirIntegracao(token, dados) {
  const sessao = exigirSessao_(token);

  if (!dados || !dados.protocolo) {
    throw new Error('Informe o protocolo da integração.');
  }

  if (!dados.confirmouTreinamento) {
    throw new Error('Confirme que o treinamento online foi concluído.');
  }

  if (!dados.aceitouRegulamento) {
    throw new Error('Confirme a leitura e o aceite do regulamento interno.');
  }

  if (!dados.aceitouSeguranca) {
    throw new Error('Confirme a leitura das orientações de segurança.');
  }

  if (!dados.aceitouPrivacidade) {
    throw new Error('Confirme o aceite das orientações de privacidade e uso de dados.');
  }

  const ss = abrirBanco_();
  const aba = ss.getSheetByName('INTEGRACOES');
  const linha = localizarLinhaPorProtocolo_(aba, dados.protocolo);

  if (!linha) throw new Error('Integração não encontrada.');

  const matricula = aba.getRange(linha, 3).getValue();
  const colaborador = aba.getRange(linha, 4).getValue();

  if (sessao.perfil === 'COLABORADOR' &&
      String(sessao.matricula) !== String(matricula)) {
    throw new Error('Você não possui acesso a esta integração.');
  }

  const codigoCertificado = 'CERT-' +
    Utilities.formatDate(new Date(), CONFIG.FUSO, 'yyyyMMdd') + '-' +
    String(Math.floor(Math.random() * 9000) + 1000);

  aba.getRange(linha, 9).setValue('CONCLUIDA');
  aba.getRange(linha, 10).setValue(new Date());
  aba.getRange(linha, 11).setValue('SIM');
  aba.getRange(linha, 12).setValue('SIM');
  aba.getRange(linha, 13).setValue('SIM');
  aba.getRange(linha, 14).setValue('SIM');
  aba.getRange(linha, 15).setValue(codigoCertificado);

  ss.getSheetByName('HISTORICO').appendRow([
    new Date(),
    Utilities.formatDate(new Date(), CONFIG.FUSO, 'HH:mm:ss'),
    matricula,
    colaborador,
    'INTEGRACAO_CONCLUIDA',
    'INTEGRACOES',
    dados.protocolo,
    'Integração concluída por ' + sessao.email + ' | Certificado: ' + codigoCertificado
  ]);

  return {
    sucesso: true,
    protocolo: dados.protocolo,
    codigoCertificado,
    mensagem: 'Integração concluída com sucesso.'
  };
}

function listarIntegracoes(token) {
  const sessao = exigirSessao_(token);
  const aba = abrirBanco_().getSheetByName('INTEGRACOES');

  if (!aba || aba.getLastRow() < 2) return [];

  let dados = aba.getRange(2, 1, aba.getLastRow() - 1, 15).getValues();

  if (sessao.perfil === 'COLABORADOR') {
    dados = dados.filter(l => String(l[2]) === String(sessao.matricula));
  }

  return dados.reverse().map(l => ({
    protocolo: l[0],
    dataInicio: formatarData_(l[1]),
    matricula: l[2],
    colaborador: l[3],
    setor: l[4],
    cargo: l[5],
    dataAdmissao: formatarDataSimples_(l[6]),
    responsavelCadastro: l[7],
    status: l[8],
    dataConclusao: formatarData_(l[9]),
    confirmouTreinamento: l[10],
    aceitouRegulamento: l[11],
    aceitouSeguranca: l[12],
    aceitouPrivacidade: l[13],
    codigoCertificado: l[14]
  }));
}

function obterResumoIntegracoes(token) {
  exigirPerfil_(token, ['LIDER', 'RH', 'ADMINISTRADOR']);

  const aba = abrirBanco_().getSheetByName('INTEGRACOES');
  if (!aba || aba.getLastRow() < 2) {
    return {
      total: 0,
      iniciadas: 0,
      concluidas: 0,
      pendentes: 0,
      taxaConclusao: 0,
      porSetor: []
    };
  }

  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 15).getValues();
  const total = dados.length;
  const concluidas = dados.filter(l => String(l[8]).toUpperCase() === 'CONCLUIDA').length;
  const iniciadas = dados.filter(l => String(l[8]).toUpperCase() === 'INICIADA').length;
  const pendentes = total - concluidas;
  const taxaConclusao = total ? Math.round((concluidas / total) * 10000) / 100 : 0;

  const porSetor = {};
  dados.forEach(l => {
    const setor = String(l[4] || 'Não informado');
    if (!porSetor[setor]) porSetor[setor] = { total: 0, concluidas: 0 };
    porSetor[setor].total++;
    if (String(l[8]).toUpperCase() === 'CONCLUIDA') {
      porSetor[setor].concluidas++;
    }
  });

  return {
    total,
    iniciadas,
    concluidas,
    pendentes,
    taxaConclusao,
    porSetor: Object.keys(porSetor).map(setor => ({
      setor,
      total: porSetor[setor].total,
      concluidas: porSetor[setor].concluidas,
      percentual: Math.round((porSetor[setor].concluidas / porSetor[setor].total) * 10000) / 100
    })).sort((a, b) => b.percentual - a.percentual)
  };
}

function gerarDadosCertificado(token, protocolo) {
  const sessao = exigirSessao_(token);
  const aba = abrirBanco_().getSheetByName('INTEGRACOES');
  const linha = localizarLinhaPorProtocolo_(aba, protocolo);

  if (!linha) throw new Error('Integração não encontrada.');

  const valores = aba.getRange(linha, 1, 1, 15).getValues()[0];

  if (sessao.perfil === 'COLABORADOR' &&
      String(sessao.matricula) !== String(valores[2])) {
    throw new Error('Você não possui acesso a este certificado.');
  }

  if (String(valores[8]).toUpperCase() !== 'CONCLUIDA') {
    throw new Error('O certificado somente é disponibilizado após a conclusão.');
  }

  return {
    empresa: CONFIG.EMPRESA,
    colaborador: valores[3],
    matricula: valores[2],
    setor: valores[4],
    cargo: valores[5],
    dataConclusao: formatarDataSimples_(valores[9]),
    codigoCertificado: valores[14],
    protocolo: valores[0]
  };
}

function localizarIntegracaoAtiva_(aba, matricula) {
  if (!aba || aba.getLastRow() < 2) return null;

  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 15).getValues();

  for (let i = 0; i < dados.length; i++) {
    if (String(dados[i][2]) === String(matricula) &&
        !['CANCELADA'].includes(String(dados[i][8]).toUpperCase())) {
      return {
        linha: i + 2,
        protocolo: dados[i][0],
        status: dados[i][8]
      };
    }
  }

  return null;
}

function localizarLinhaPorProtocolo_(aba, protocolo) {
  if (!aba || aba.getLastRow() < 2) return 0;

  const valores = aba.getRange(2, 1, aba.getLastRow() - 1, 1).getValues();

  for (let i = 0; i < valores.length; i++) {
    if (String(valores[i][0]) === String(protocolo)) return i + 2;
  }

  return 0;
}



/* ==========================================================
   ETAPA 9 - GESTÃO DE TREINAMENTOS E CERTIFICADOS
========================================================== */

function salvarTreinamento(token, dados) {
  const sessao = exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);

  if (!dados) throw new Error('Dados não recebidos.');
  if (!dados.titulo) throw new Error('Informe o título do treinamento.');
  if (!dados.tipo) throw new Error('Informe o tipo do treinamento.');
  if (!dados.cargaHoraria) throw new Error('Informe a carga horária.');
  if (!dados.validadeMeses) throw new Error('Informe a validade em meses.');

  const ss = abrirBanco_();
  const aba = ss.getSheetByName('TREINAMENTOS');
  const id = gerarProtocolo_('TRN');

  aba.appendRow([
    id,
    dados.titulo,
    dados.tipo,
    dados.descricao || '',
    Number(dados.cargaHoraria),
    Number(dados.validadeMeses),
    dados.obrigatorio ? 'SIM' : 'NÃO',
    dados.setores || 'TODOS',
    dados.cargos || 'TODOS',
    'ATIVO',
    new Date(),
    sessao.email
  ]);

  ss.getSheetByName('HISTORICO').appendRow([
    new Date(),
    Utilities.formatDate(new Date(), CONFIG.FUSO, 'HH:mm:ss'),
    '',
    '',
    'TREINAMENTO_CADASTRADO',
    'TREINAMENTOS',
    id,
    dados.titulo + ' | Responsável: ' + sessao.email
  ]);

  return {
    sucesso: true,
    id,
    mensagem: 'Treinamento cadastrado com sucesso.'
  };
}

function listarTreinamentos(token) {
  exigirSessao_(token);

  const aba = abrirBanco_().getSheetByName('TREINAMENTOS');
  if (!aba || aba.getLastRow() < 2) return [];

  return aba.getRange(2, 1, aba.getLastRow() - 1, 12).getValues()
    .filter(l => String(l[9]).toUpperCase() === 'ATIVO')
    .map(l => ({
      id: l[0],
      titulo: l[1],
      tipo: l[2],
      descricao: l[3],
      cargaHoraria: l[4],
      validadeMeses: l[5],
      obrigatorio: l[6],
      setores: l[7],
      cargos: l[8],
      status: l[9]
    }));
}

function salvarConclusaoTreinamento(token, dados) {
  const sessao = exigirPerfil_(token, ['LIDER', 'RH', 'ADMINISTRADOR']);

  if (!dados) throw new Error('Dados não recebidos.');
  if (!dados.treinamentoId) throw new Error('Selecione o treinamento.');
  if (!dados.matricula) throw new Error('Informe a matrícula.');
  if (!dados.colaborador) throw new Error('Informe o colaborador.');
  if (!dados.setor) throw new Error('Informe o setor.');
  if (!dados.cargo) throw new Error('Informe o cargo.');
  if (!dados.dataConclusao) throw new Error('Informe a data de conclusão.');

  const ss = abrirBanco_();
  const abaTreinamentos = ss.getSheetByName('TREINAMENTOS');
  const linhaTreinamento = localizarLinhaPorId_(abaTreinamentos, dados.treinamentoId);

  if (!linhaTreinamento) throw new Error('Treinamento não encontrado.');

  const treinamento = abaTreinamentos.getRange(linhaTreinamento, 1, 1, 12).getValues()[0];
  const validadeMeses = Number(treinamento[5]) || 0;
  const dataConclusao = new Date(dados.dataConclusao + 'T12:00:00');
  const dataVencimento = new Date(dataConclusao);

  if (validadeMeses > 0) {
    dataVencimento.setMonth(dataVencimento.getMonth() + validadeMeses);
  }

  const protocolo = gerarProtocolo_('CRT');
  const codigoCertificado = 'CERT-TRN-' +
    Utilities.formatDate(new Date(), CONFIG.FUSO, 'yyyyMMdd') + '-' +
    String(Math.floor(Math.random() * 9000) + 1000);

  const arquivo = salvarArquivoTreinamento_(dados.arquivo, protocolo, dados.colaborador);

  ss.getSheetByName('TREINAMENTOS_CONCLUIDOS').appendRow([
    protocolo,
    dados.treinamentoId,
    treinamento[1],
    dados.matricula,
    dados.colaborador,
    dados.setor,
    dados.cargo,
    dataConclusao,
    validadeMeses > 0 ? dataVencimento : '',
    calcularStatusVencimento_(validadeMeses > 0 ? dataVencimento : ''),
    Number(dados.nota || 0),
    dados.instrutor || '',
    codigoCertificado,
    arquivo ? arquivo.url : '',
    new Date(),
    sessao.email,
    dados.observacao || ''
  ]);

  ss.getSheetByName('HISTORICO').appendRow([
    new Date(),
    Utilities.formatDate(new Date(), CONFIG.FUSO, 'HH:mm:ss'),
    dados.matricula,
    dados.colaborador,
    'TREINAMENTO_CONCLUIDO',
    'TREINAMENTOS_CONCLUIDOS',
    protocolo,
    treinamento[1] + ' | Certificado: ' + codigoCertificado
  ]);

  return {
    sucesso: true,
    protocolo,
    codigoCertificado,
    dataVencimento: validadeMeses > 0 ? formatarDataSimples_(dataVencimento) : 'Sem vencimento',
    mensagem: 'Conclusão registrada com sucesso.'
  };
}

function listarTreinamentosConcluidos(token) {
  const sessao = exigirSessao_(token);
  const aba = abrirBanco_().getSheetByName('TREINAMENTOS_CONCLUIDOS');

  if (!aba || aba.getLastRow() < 2) return [];

  let dados = aba.getRange(2, 1, aba.getLastRow() - 1, 17).getValues();

  if (sessao.perfil === 'COLABORADOR') {
    dados = dados.filter(l => String(l[3]) === String(sessao.matricula));
  }

  return dados.reverse().map(l => ({
    protocolo: l[0],
    treinamentoId: l[1],
    treinamento: l[2],
    matricula: l[3],
    colaborador: l[4],
    setor: l[5],
    cargo: l[6],
    dataConclusao: formatarDataSimples_(l[7]),
    dataVencimento: formatarDataSimples_(l[8]),
    status: calcularStatusVencimento_(l[8]),
    nota: l[10],
    instrutor: l[11],
    codigoCertificado: l[12],
    arquivoUrl: l[13],
    observacao: l[16]
  }));
}

function obterResumoTreinamentos(token) {
  exigirPerfil_(token, ['LIDER', 'RH', 'ADMINISTRADOR']);

  const ss = abrirBanco_();
  const abaTreinamentos = ss.getSheetByName('TREINAMENTOS');
  const abaConcluidos = ss.getSheetByName('TREINAMENTOS_CONCLUIDOS');

  const treinamentosAtivos = abaTreinamentos && abaTreinamentos.getLastRow() >= 2
    ? abaTreinamentos.getRange(2, 1, abaTreinamentos.getLastRow() - 1, 12).getValues()
        .filter(l => String(l[9]).toUpperCase() === 'ATIVO').length
    : 0;

  const concluidos = abaConcluidos && abaConcluidos.getLastRow() >= 2
    ? abaConcluidos.getRange(2, 1, abaConcluidos.getLastRow() - 1, 17).getValues()
    : [];

  let validos = 0;
  let vencendo = 0;
  let vencidos = 0;
  let semVencimento = 0;

  const porSetor = {};

  concluidos.forEach(l => {
    const status = calcularStatusVencimento_(l[8]);
    if (status === 'VÁLIDO') validos++;
    else if (status === 'VENCE EM ATÉ 30 DIAS') vencendo++;
    else if (status === 'VENCIDO') vencidos++;
    else semVencimento++;

    const setor = String(l[5] || 'Não informado');
    if (!porSetor[setor]) porSetor[setor] = { total: 0, vencidos: 0 };
    porSetor[setor].total++;
    if (status === 'VENCIDO') porSetor[setor].vencidos++;
  });

  return {
    treinamentosAtivos,
    registrosConcluidos: concluidos.length,
    validos,
    vencendo,
    vencidos,
    semVencimento,
    porSetor: Object.keys(porSetor).map(setor => ({
      setor,
      total: porSetor[setor].total,
      vencidos: porSetor[setor].vencidos,
      percentualVencido: porSetor[setor].total
        ? Math.round((porSetor[setor].vencidos / porSetor[setor].total) * 10000) / 100
        : 0
    })).sort((a, b) => b.vencidos - a.vencidos)
  };
}

function gerarDadosCertificadoTreinamento(token, protocolo) {
  const sessao = exigirSessao_(token);
  const aba = abrirBanco_().getSheetByName('TREINAMENTOS_CONCLUIDOS');
  const linha = localizarLinhaPorProtocolo_(aba, protocolo);

  if (!linha) throw new Error('Certificado não encontrado.');

  const v = aba.getRange(linha, 1, 1, 17).getValues()[0];

  if (sessao.perfil === 'COLABORADOR' &&
      String(sessao.matricula) !== String(v[3])) {
    throw new Error('Você não possui acesso a este certificado.');
  }

  return {
    empresa: CONFIG.EMPRESA,
    treinamento: v[2],
    colaborador: v[4],
    matricula: v[3],
    setor: v[5],
    cargo: v[6],
    dataConclusao: formatarDataSimples_(v[7]),
    dataVencimento: formatarDataSimples_(v[8]) || 'Sem vencimento',
    cargaHoraria: obterCargaHorariaTreinamento_(v[1]),
    instrutor: v[11] || 'Não informado',
    codigoCertificado: v[12],
    protocolo: v[0]
  };
}

function listarAlertasTreinamentos(token) {
  exigirPerfil_(token, ['LIDER', 'RH', 'ADMINISTRADOR']);

  const aba = abrirBanco_().getSheetByName('TREINAMENTOS_CONCLUIDOS');
  if (!aba || aba.getLastRow() < 2) return [];

  return aba.getRange(2, 1, aba.getLastRow() - 1, 17).getValues()
    .map(l => ({
      protocolo: l[0],
      treinamento: l[2],
      matricula: l[3],
      colaborador: l[4],
      setor: l[5],
      dataVencimento: formatarDataSimples_(l[8]),
      status: calcularStatusVencimento_(l[8]),
      arquivoUrl: l[13]
    }))
    .filter(item => ['VENCIDO', 'VENCE EM ATÉ 30 DIAS'].includes(item.status))
    .sort((a, b) => a.status.localeCompare(b.status));
}

function calcularStatusVencimento_(dataVencimento) {
  if (!dataVencimento) return 'SEM VENCIMENTO';

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  const vencimento = new Date(dataVencimento);
  vencimento.setHours(0, 0, 0, 0);

  const diferenca = Math.ceil((vencimento - hoje) / 86400000);

  if (diferenca < 0) return 'VENCIDO';
  if (diferenca <= 30) return 'VENCE EM ATÉ 30 DIAS';
  return 'VÁLIDO';
}

function localizarLinhaPorId_(aba, id) {
  if (!aba || aba.getLastRow() < 2) return 0;

  const valores = aba.getRange(2, 1, aba.getLastRow() - 1, 1).getValues();

  for (let i = 0; i < valores.length; i++) {
    if (String(valores[i][0]) === String(id)) return i + 2;
  }

  return 0;
}

function obterCargaHorariaTreinamento_(treinamentoId) {
  const aba = abrirBanco_().getSheetByName('TREINAMENTOS');
  const linha = localizarLinhaPorId_(aba, treinamentoId);
  return linha ? aba.getRange(linha, 5).getValue() : '';
}

function salvarArquivoTreinamento_(arquivo, protocolo, colaborador) {
  if (!arquivo || !arquivo.base64 || !arquivo.nome) return null;

  validarArquivo_(arquivo);

  const raiz = obterPastaRaiz_();
  const pasta = obterOuCriarPasta_(raiz, '08 - Treinamentos e Certificados');

  const bytes = Utilities.base64Decode(arquivo.base64);
  const blob = Utilities.newBlob(
    bytes,
    arquivo.tipo || 'application/octet-stream',
    protocolo + ' - ' + colaborador + ' - ' + arquivo.nome
  );

  const file = pasta.createFile(blob);

  return {
    id: file.getId(),
    url: file.getUrl(),
    nome: file.getName()
  };
}



/* ==========================================================
   ETAPA 10 - PAINEL EXECUTIVO DA DIRETORIA
========================================================== */

function obterPainelExecutivo(token) {
  exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);

  const ss = abrirBanco_();
  const hoje = new Date();
  const mesAtual = hoje.getMonth();
  const anoAtual = hoje.getFullYear();

  const colaboradores = lerDadosFlexivel_(ss.getSheetByName('COLABORADORES'));
  const atestados = lerDadosFlexivel_(ss.getSheetByName('ATESTADOS'));
  const ferias = lerDadosFlexivel_(ss.getSheetByName('FERIAS'));
  const solicitacoes = lerDadosFlexivel_(ss.getSheetByName('SOLICITACOES'));
  const avaliacoes = lerDadosFlexivel_(ss.getSheetByName('AVALIACOES'));
  const integracoes = lerDadosFlexivel_(ss.getSheetByName('INTEGRACOES'));
  const treinamentos = lerDadosFlexivel_(ss.getSheetByName('TREINAMENTOS_CONCLUIDOS'));
  const usuarios = lerDadosFlexivel_(ss.getSheetByName('USUARIOS'));

  const totalColaboradores = contarColaboradoresAtivosExecutivo_(colaboradores);
  const usuariosAtivos = usuarios.filter(l => String(l[6] || '').toUpperCase() === 'SIM').length;

  const atestadosMes = atestados.filter(l => mesmaCompetencia_(l[1], mesAtual, anoAtual)).length;
  const diasAtestadosMes = atestados
    .filter(l => mesmaCompetencia_(l[1], mesAtual, anoAtual))
    .reduce((soma, l) => soma + (Number(l[8]) || 0), 0);

  const feriasPendentes = ferias.filter(l => {
    const status = String(l[13] || '').toUpperCase();
    return ['PENDENTE', 'AGUARDANDO_LIDER', 'AGUARDANDO_ANALISE_RH'].includes(status);
  }).length;

  const solicitacoesAbertas = solicitacoes.filter(l => {
    const status = String(l[10] || '').toUpperCase();
    return !['CONCLUIDA', 'CONCLUÍDA', 'FINALIZADA', 'REPROVADA', 'CANCELADA'].includes(status);
  }).length;

  const notas = avaliacoes.map(l => Number(l[7])).filter(n => !isNaN(n) && n > 0);
  const mediaAvaliacoes = notas.length
    ? Math.round((notas.reduce((a, b) => a + b, 0) / notas.length) * 100) / 100
    : 0;

  const integracoesConcluidas = integracoes.filter(l =>
    String(l[8] || '').toUpperCase() === 'CONCLUIDA'
  ).length;
  const taxaIntegracao = integracoes.length
    ? Math.round((integracoesConcluidas / integracoes.length) * 10000) / 100
    : 0;

  let treinamentosValidos = 0;
  let treinamentosVencendo = 0;
  let treinamentosVencidos = 0;

  treinamentos.forEach(l => {
    const status = calcularStatusVencimento_(l[8]);
    if (status === 'VÁLIDO') treinamentosValidos++;
    if (status === 'VENCE EM ATÉ 30 DIAS') treinamentosVencendo++;
    if (status === 'VENCIDO') treinamentosVencidos++;
  });

  const absenteismoEstimado = totalColaboradores > 0
    ? Math.round((diasAtestadosMes / (totalColaboradores * 22)) * 10000) / 100
    : 0;

  const atestadosPorMes = montarSerieMensal_(atestados, 1, 12, false);
  const diasAtestadosPorMes = montarSerieMensal_(atestados, 1, 12, true, 8);
  const avaliacoesPorMes = montarSerieMensal_(avaliacoes, 1, 12, false);
  const integracoesPorMes = montarSerieMensal_(integracoes, 1, 12, false);

  const atestadosPorSetor = agruparPorSetor_(atestados, 4);
  const feriasPorSetor = agruparPorSetor_(ferias, 4);
  const solicitacoesPorTipo = agruparPorCampo_(solicitacoes, 5);
  const avaliacoesPorClassificacao = agruparPorCampo_(avaliacoes, 8);
  const treinamentosPorStatus = [
    { nome: 'Válidos', quantidade: treinamentosValidos },
    { nome: 'Vencem em até 30 dias', quantidade: treinamentosVencendo },
    { nome: 'Vencidos', quantidade: treinamentosVencidos }
  ];

  const alertas = montarAlertasExecutivos_({
    feriasPendentes,
    solicitacoesAbertas,
    treinamentosVencendo,
    treinamentosVencidos,
    mediaAvaliacoes,
    taxaIntegracao,
    absenteismoEstimado
  });

  return {
    atualizadoEm: Utilities.formatDate(hoje, CONFIG.FUSO, 'dd/MM/yyyy HH:mm'),
    indicadores: {
      totalColaboradores,
      usuariosAtivos,
      atestadosMes,
      diasAtestadosMes,
      absenteismoEstimado,
      feriasPendentes,
      solicitacoesAbertas,
      mediaAvaliacoes,
      taxaIntegracao,
      treinamentosValidos,
      treinamentosVencendo,
      treinamentosVencidos
    },
    series: {
      atestadosPorMes,
      diasAtestadosPorMes,
      avaliacoesPorMes,
      integracoesPorMes
    },
    rankings: {
      atestadosPorSetor,
      feriasPorSetor,
      solicitacoesPorTipo,
      avaliacoesPorClassificacao,
      treinamentosPorStatus
    },
    alertas
  };
}

function obterRelatorioExecutivo(token) {
  const dados = obterPainelExecutivo(token);

  return {
    titulo: 'Relatório Executivo de RH',
    empresa: CONFIG.EMPRESA,
    atualizadoEm: dados.atualizadoEm,
    indicadores: dados.indicadores,
    alertas: dados.alertas,
    resumo: [
      'Colaboradores ativos: ' + dados.indicadores.totalColaboradores,
      'Atestados no mês: ' + dados.indicadores.atestadosMes,
      'Dias de afastamento no mês: ' + dados.indicadores.diasAtestadosMes,
      'Absenteísmo estimado: ' + dados.indicadores.absenteismoEstimado + '%',
      'Férias pendentes: ' + dados.indicadores.feriasPendentes,
      'Solicitações abertas: ' + dados.indicadores.solicitacoesAbertas,
      'Média das avaliações: ' + dados.indicadores.mediaAvaliacoes,
      'Taxa de integração: ' + dados.indicadores.taxaIntegracao + '%',
      'Treinamentos vencidos: ' + dados.indicadores.treinamentosVencidos
    ]
  };
}

function montarSerieMensal_(dados, colunaData, meses, somarValor, colunaValor) {
  const resultado = [];
  const hoje = new Date();

  for (let i = meses - 1; i >= 0; i--) {
    const referencia = new Date(hoje.getFullYear(), hoje.getMonth() - i, 1);
    const mes = referencia.getMonth();
    const ano = referencia.getFullYear();

    const filtrados = dados.filter(l => mesmaCompetencia_(l[colunaData], mes, ano));
    const valor = somarValor
      ? filtrados.reduce((soma, l) => soma + (Number(l[colunaValor]) || 0), 0)
      : filtrados.length;

    resultado.push({
      mes: Utilities.formatDate(referencia, CONFIG.FUSO, 'MM/yyyy'),
      valor
    });
  }

  return resultado;
}

function agruparPorSetor_(dados, colunaSetor) {
  return agruparPorCampo_(dados, colunaSetor);
}

function agruparPorCampo_(dados, coluna) {
  const mapa = {};

  dados.forEach(l => {
    const nome = String(l[coluna] || 'Não informado').trim() || 'Não informado';
    mapa[nome] = (mapa[nome] || 0) + 1;
  });

  return Object.keys(mapa)
    .map(nome => ({ nome, quantidade: mapa[nome] }))
    .sort((a, b) => b.quantidade - a.quantidade)
    .slice(0, 10);
}

function montarAlertasExecutivos_(i) {
  const alertas = [];

  if (i.treinamentosVencidos > 0) {
    alertas.push({
      nivel: 'CRITICO',
      titulo: 'Treinamentos vencidos',
      descricao: i.treinamentosVencidos + ' certificado(s) precisam de renovação.'
    });
  }

  if (i.treinamentosVencendo > 0) {
    alertas.push({
      nivel: 'ATENCAO',
      titulo: 'Treinamentos próximos do vencimento',
      descricao: i.treinamentosVencendo + ' certificado(s) vencem em até 30 dias.'
    });
  }

  if (i.feriasPendentes > 0) {
    alertas.push({
      nivel: 'ATENCAO',
      titulo: 'Férias pendentes',
      descricao: i.feriasPendentes + ' solicitação(ões) aguardam análise.'
    });
  }

  if (i.solicitacoesAbertas > 5) {
    alertas.push({
      nivel: 'ATENCAO',
      titulo: 'Volume de solicitações',
      descricao: i.solicitacoesAbertas + ' solicitações estão abertas.'
    });
  }

  if (i.absenteismoEstimado >= 3) {
    alertas.push({
      nivel: 'CRITICO',
      titulo: 'Absenteísmo elevado',
      descricao: 'O índice estimado do mês está em ' + i.absenteismoEstimado + '%.'
    });
  }

  if (i.mediaAvaliacoes > 0 && i.mediaAvaliacoes < 3) {
    alertas.push({
      nivel: 'ATENCAO',
      titulo: 'Desempenho abaixo do esperado',
      descricao: 'A média geral das avaliações está em ' + i.mediaAvaliacoes + '.'
    });
  }

  if (i.taxaIntegracao < 80) {
    alertas.push({
      nivel: 'ATENCAO',
      titulo: 'Integrações pendentes',
      descricao: 'A taxa de conclusão das integrações está em ' + i.taxaIntegracao + '%.'
    });
  }

  if (!alertas.length) {
    alertas.push({
      nivel: 'OK',
      titulo: 'Indicadores sob controle',
      descricao: 'Nenhum alerta crítico identificado no momento.'
    });
  }

  return alertas;
}

function contarColaboradoresAtivosExecutivo_(dados) {
  return dados.filter(l => {
    if (!l.length) return false;

    const statusPossiveis = [
      String(l[11] || '').toUpperCase(),
      String(l[12] || '').toUpperCase(),
      String(l[13] || '').toUpperCase()
    ];

    return !statusPossiveis.some(status =>
      ['INATIVO', 'DESLIGADO', 'DEMITIDO', 'NÃO'].includes(status)
    );
  }).length;
}

function lerDadosFlexivel_(aba) {
  if (!aba || aba.getLastRow() < 2) return [];
  return aba.getRange(2, 1, aba.getLastRow() - 1, aba.getLastColumn()).getValues();
}



/* ==========================================================
   ETAPA 13 - FICHA COMPLETA DO COLABORADOR
========================================================== */

function buscarColaboradoresFicha(token, termo) {
  exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);

  termo = String(termo || '').trim().toLowerCase();
  const aba = abrirBanco_().getSheetByName('COLABORADORES');
  if (!aba || aba.getLastRow() < 2) return [];

  const cabecalhos = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0]
    .map(v => String(v || '').trim().toUpperCase());
  const idx = nome => cabecalhos.indexOf(nome);
  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, aba.getLastColumn()).getValues();

  return dados
    .map((l, i) => ({
      linha: i + 2,
      matricula: idx('MATRICULA') >= 0 ? l[idx('MATRICULA')] || '' : '',
      nome: idx('NOME') >= 0 ? l[idx('NOME')] || '' : '',
      cpf: idx('CPF') >= 0 ? l[idx('CPF')] || '' : '',
      cargo: idx('CARGO') >= 0 ? l[idx('CARGO')] || '' : '',
      setor: idx('SETOR') >= 0 ? l[idx('SETOR')] || '' : '',
      lider: idx('LIDER') >= 0 ? l[idx('LIDER')] || '' : '',
      admissao: idx('DATA_ADMISSAO') >= 0 ? formatarDataSimples_(l[idx('DATA_ADMISSAO')]) : '',
      email: idx('EMAIL') >= 0 ? l[idx('EMAIL')] || '' : '',
      telefone: idx('TELEFONE') >= 0 ? l[idx('TELEFONE')] || '' : '',
      nascimento: idx('DATA_NASCIMENTO') >= 0 ? formatarDataSimples_(l[idx('DATA_NASCIMENTO')]) : '',
      status: idx('STATUS') >= 0 ? l[idx('STATUS')] || 'ATIVO' : 'ATIVO'
    }))
    .filter(item => {
      if (!item.matricula || !item.nome) return false;
      if (!termo) return true;
      return [item.matricula, item.nome, item.cpf, item.cargo, item.setor, item.email]
        .join(' ').toLowerCase().includes(termo);
    })
    .slice(0, 30);
}

function obterFichaCompleta(token, matricula) {
  const sessao = exigirSessao_(token);

  if (!matricula) {
    matricula = sessao.matricula;
  }

  if (sessao.perfil === 'COLABORADOR' &&
      String(sessao.matricula) !== String(matricula)) {
    throw new Error('Você não possui acesso a esta ficha.');
  }

  if (!['COLABORADOR', 'RH', 'ADMINISTRADOR'].includes(sessao.perfil)) {
    throw new Error('Seu perfil não possui acesso à ficha completa.');
  }

  const ss = abrirBanco_();
  const colaborador = localizarColaboradorCompleto_(ss.getSheetByName('COLABORADORES'), matricula);

  if (!colaborador) throw new Error('Colaborador não encontrado.');

  const atestados = filtrarPorMatricula_(ss.getSheetByName('ATESTADOS'), matricula, 2)
    .map(l => ({
      protocolo: l[0], data: formatarData_(l[1]), tipo: l[5] || '',
      inicio: formatarDataSimples_(l[7]), dias: l[8] || '',
      status: l[13] || '', arquivoUrl: l[12] || ''
    }));

  const ferias = filtrarPorMatricula_(ss.getSheetByName('FERIAS'), matricula, 2)
    .map(l => ({
      protocolo: l[0], data: formatarData_(l[1]),
      inicio: formatarDataSimples_(l[6]), dias: l[7] || '',
      status: l[13] || '', observacao: l[11] || ''
    }));

  const solicitacoes = filtrarPorMatricula_(ss.getSheetByName('SOLICITACOES'), matricula, 2)
    .map(l => ({
      protocolo: l[0], data: formatarData_(l[1]), tipo: l[5] || '',
      descricao: l[6] || '', status: l[10] || ''
    }));

  const avaliacoes = filtrarPorMatricula_(ss.getSheetByName('AVALIACOES'), matricula, 3)
    .map(l => ({
      id: l[0], data: formatarDataSimples_(l[1]), tipo: l[2] || '',
      nota: l[7] || '', classificacao: l[8] || '',
      pontosFortes: l[9] || '', pontosMelhoria: l[10] || '',
      planoAcao: l[11] || ''
    }));

  const integracoes = filtrarPorMatricula_(ss.getSheetByName('INTEGRACOES'), matricula, 2)
    .map(l => ({
      protocolo: l[0], inicio: formatarData_(l[1]), status: l[8] || '',
      conclusao: formatarData_(l[9]), certificado: l[14] || ''
    }));

  const treinamentos = filtrarPorMatricula_(ss.getSheetByName('TREINAMENTOS_CONCLUIDOS'), matricula, 3)
    .map(l => ({
      protocolo: l[0], treinamento: l[2] || '',
      conclusao: formatarDataSimples_(l[7]),
      vencimento: formatarDataSimples_(l[8]),
      status: calcularStatusVencimento_(l[8]),
      certificado: l[12] || '', arquivoUrl: l[13] || ''
    }));

  const documentos = filtrarPorMatricula_(ss.getSheetByName('DOCUMENTOS_COLABORADOR'), matricula, 1)
    .map(l => ({
      protocolo: l[0], matricula: l[1], nome: l[2], tipo: l[3],
      descricao: l[4], emissao: formatarDataSimples_(l[5]),
      validade: formatarDataSimples_(l[6]), status: calcularStatusDocumento_(l[6]),
      arquivoUrl: l[7] || '', registradoEm: formatarData_(l[8]),
      registradoPor: l[9] || ''
    }));

  const epis = filtrarPorMatricula_(ss.getSheetByName('EPIS'), matricula, 1)
    .map(l => ({
      protocolo: l[0], epi: l[3], ca: l[4], quantidade: l[5],
      entrega: formatarDataSimples_(l[6]), devolucao: formatarDataSimples_(l[7]),
      status: l[8], termoUrl: l[9] || '', observacao: l[12] || ''
    }));

  const ocorrencias = filtrarPorMatricula_(ss.getSheetByName('OCORRENCIAS_COLABORADOR'), matricula, 1)
    .map(l => ({
      protocolo: l[0], data: formatarDataSimples_(l[3]), tipo: l[4],
      titulo: l[5], descricao: l[6], acao: l[7], status: l[8],
      arquivoUrl: l[9] || '', registradoPor: l[11] || ''
    }));

  const historico = montarLinhaDoTempoColaborador_({
    atestados, ferias, solicitacoes, avaliacoes,
    integracoes, treinamentos, documentos, epis, ocorrencias
  });

  return {
    colaborador,
    resumo: {
      atestados: atestados.length,
      diasAfastamento: atestados.reduce((s, a) => s + (Number(a.dias) || 0), 0),
      ferias: ferias.length,
      solicitacoes: solicitacoes.length,
      avaliacoes: avaliacoes.length,
      mediaAvaliacoes: calcularMediaLista_(avaliacoes.map(a => Number(a.nota))),
      treinamentos: treinamentos.length,
      treinamentosVencidos: treinamentos.filter(t => t.status === 'VENCIDO').length,
      documentos: documentos.length,
      documentosVencidos: documentos.filter(d => d.status === 'VENCIDO').length,
      episAtivos: epis.filter(e => String(e.status).toUpperCase() === 'ENTREGUE').length,
      ocorrencias: ocorrencias.length
    },
    atestados, ferias, solicitacoes, avaliacoes,
    integracoes, treinamentos, documentos, epis, ocorrencias,
    historico
  };
}

function salvarDocumentoColaborador(token, dados) {
  const sessao = exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);

  if (!dados || !dados.matricula || !dados.nome || !dados.tipo) {
    throw new Error('Preencha matrícula, nome e tipo do documento.');
  }

  const protocolo = gerarProtocolo_('DOC');
  const arquivo = salvarArquivoFicha_(dados.arquivo, '09 - Documentos de Colaboradores', protocolo, dados.nome);

  abrirBanco_().getSheetByName('DOCUMENTOS_COLABORADOR').appendRow([
    protocolo, dados.matricula, dados.nome, dados.tipo,
    dados.descricao || '', dados.dataEmissao || '', dados.dataValidade || '',
    arquivo ? arquivo.url : '', new Date(), sessao.email
  ]);

  registrarHistoricoFicha_(dados.matricula, dados.nome, 'DOCUMENTO_REGISTRADO', protocolo, dados.tipo, sessao.email);

  return { sucesso: true, protocolo, mensagem: 'Documento registrado com sucesso.' };
}

function salvarEpiColaborador(token, dados) {
  const sessao = exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);

  if (!dados || !dados.matricula || !dados.nome || !dados.epi || !dados.dataEntrega) {
    throw new Error('Preencha matrícula, nome, EPI e data de entrega.');
  }

  const protocolo = gerarProtocolo_('EPI');
  const arquivo = salvarArquivoFicha_(dados.arquivo, '11 - EPIs', protocolo, dados.nome);

  abrirBanco_().getSheetByName('EPIS').appendRow([
    protocolo, dados.matricula, dados.nome, dados.epi,
    dados.ca || '', Number(dados.quantidade || 1), dados.dataEntrega,
    dados.dataDevolucao || '', dados.status || 'ENTREGUE',
    arquivo ? arquivo.url : '', new Date(), sessao.email,
    dados.observacao || ''
  ]);

  registrarHistoricoFicha_(dados.matricula, dados.nome, 'EPI_REGISTRADO', protocolo, dados.epi, sessao.email);

  return { sucesso: true, protocolo, mensagem: 'Entrega de EPI registrada com sucesso.' };
}

function salvarOcorrenciaColaborador(token, dados) {
  const sessao = exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);

  if (!dados || !dados.matricula || !dados.nome || !dados.tipo ||
      !dados.titulo || !dados.dataOcorrencia) {
    throw new Error('Preencha os campos obrigatórios da ocorrência.');
  }

  const protocolo = gerarProtocolo_('OCR');
  const arquivo = salvarArquivoFicha_(dados.arquivo, '12 - Ocorrências', protocolo, dados.nome);

  abrirBanco_().getSheetByName('OCORRENCIAS_COLABORADOR').appendRow([
    protocolo, dados.matricula, dados.nome, dados.dataOcorrencia,
    dados.tipo, dados.titulo, dados.descricao || '',
    dados.acao || '', dados.status || 'REGISTRADA',
    arquivo ? arquivo.url : '', new Date(), sessao.email
  ]);

  registrarHistoricoFicha_(dados.matricula, dados.nome, 'OCORRENCIA_REGISTRADA', protocolo, dados.titulo, sessao.email);

  return { sucesso: true, protocolo, mensagem: 'Ocorrência registrada com sucesso.' };
}

function localizarColaboradorCompleto_(aba, matricula) {
  if (!aba || aba.getLastRow() < 2) return null;

  const cabecalhos = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0]
    .map(v => String(v || '').trim().toUpperCase());
  const idx = nome => cabecalhos.indexOf(nome);
  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, aba.getLastColumn()).getValues();
  const iMatricula = idx('MATRICULA');
  if (iMatricula < 0) return null;

  const l = dados.find(item => String(item[iMatricula]) === String(matricula));
  if (!l) return null;

  const valor = nome => idx(nome) >= 0 ? l[idx(nome)] || '' : '';

  return {
    matricula: valor('MATRICULA'),
    nome: valor('NOME'),
    cpf: valor('CPF'),
    cargo: valor('CARGO'),
    setor: valor('SETOR'),
    lider: valor('LIDER'),
    admissao: formatarDataSimples_(valor('DATA_ADMISSAO')),
    email: valor('EMAIL'),
    telefone: valor('TELEFONE'),
    nascimento: formatarDataSimples_(valor('DATA_NASCIMENTO')),
    endereco: valor('ENDERECO'),
    status: valor('STATUS') || 'ATIVO'
  };
}

function filtrarPorMatricula_(aba, matricula, indiceMatricula) {
  if (!aba || aba.getLastRow() < 2) return [];
  return aba.getRange(2, 1, aba.getLastRow() - 1, aba.getLastColumn()).getValues()
    .filter(l => String(l[indiceMatricula]) === String(matricula));
}

function calcularMediaLista_(lista) {
  const validos = lista.filter(n => !isNaN(n) && n > 0);
  if (!validos.length) return 0;
  return Math.round((validos.reduce((a, b) => a + b, 0) / validos.length) * 100) / 100;
}

function calcularStatusDocumento_(validade) {
  if (!validade) return 'SEM VALIDADE';
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  const data = new Date(validade);
  data.setHours(0, 0, 0, 0);
  const dias = Math.ceil((data - hoje) / 86400000);
  if (dias < 0) return 'VENCIDO';
  if (dias <= 30) return 'VENCE EM ATÉ 30 DIAS';
  return 'VÁLIDO';
}

function salvarArquivoFicha_(arquivo, pastaNome, protocolo, colaborador) {
  if (!arquivo || !arquivo.base64 || !arquivo.nome) return null;

  validarArquivo_(arquivo);
  const raiz = obterPastaRaiz_();
  const pasta = obterOuCriarPasta_(raiz, pastaNome);
  const bytes = Utilities.base64Decode(arquivo.base64);
  const blob = Utilities.newBlob(
    bytes,
    arquivo.tipo || 'application/octet-stream',
    protocolo + ' - ' + colaborador + ' - ' + arquivo.nome
  );
  const file = pasta.createFile(blob);

  return { id: file.getId(), url: file.getUrl(), nome: file.getName() };
}

function registrarHistoricoFicha_(matricula, nome, acao, protocolo, descricao, email) {
  abrirBanco_().getSheetByName('HISTORICO').appendRow([
    new Date(),
    Utilities.formatDate(new Date(), CONFIG.FUSO, 'HH:mm:ss'),
    matricula, nome, acao, 'FICHA_COLABORADOR', protocolo,
    descricao + ' | Responsável: ' + email
  ]);
}

function montarLinhaDoTempoColaborador_(dados) {
  const itens = [];

  dados.atestados.forEach(x => itens.push({
    data: x.data, modulo: 'Atestado', titulo: x.tipo || 'Atestado',
    descricao: (x.dias || 0) + ' dia(s) - ' + x.status
  }));
  dados.ferias.forEach(x => itens.push({
    data: x.data, modulo: 'Férias', titulo: 'Solicitação de férias',
    descricao: (x.dias || 0) + ' dia(s) - ' + x.status
  }));
  dados.solicitacoes.forEach(x => itens.push({
    data: x.data, modulo: 'Solicitação', titulo: x.tipo,
    descricao: x.status
  }));
  dados.avaliacoes.forEach(x => itens.push({
    data: x.data, modulo: 'Avaliação', titulo: x.tipo,
    descricao: 'Nota ' + x.nota + ' - ' + x.classificacao
  }));
  dados.integracoes.forEach(x => itens.push({
    data: x.inicio, modulo: 'Integração', titulo: 'Integração online',
    descricao: x.status
  }));
  dados.treinamentos.forEach(x => itens.push({
    data: x.conclusao, modulo: 'Treinamento', titulo: x.treinamento,
    descricao: x.status
  }));
  dados.documentos.forEach(x => itens.push({
    data: x.registradoEm, modulo: 'Documento', titulo: x.tipo,
    descricao: x.status
  }));
  dados.epis.forEach(x => itens.push({
    data: x.entrega, modulo: 'EPI', titulo: x.epi,
    descricao: x.status
  }));
  dados.ocorrencias.forEach(x => itens.push({
    data: x.data, modulo: 'Ocorrência', titulo: x.titulo,
    descricao: x.status
  }));

  return itens.sort((a, b) => {
    const da = converterDataBr_(a.data);
    const db = converterDataBr_(b.data);
    return db - da;
  });
}

function converterDataBr_(valor) {
  if (!valor) return new Date(0);
  if (valor instanceof Date) return valor;
  const parte = String(valor).split(' ')[0].split('/');
  if (parte.length === 3) return new Date(parte[2], parte[1] - 1, parte[0]);
  return new Date(valor);
}



/* ==========================================================
   ETAPA 14 - MODO APLICATIVO / PWA
========================================================== */

function obterConfiguracaoAplicativo(token) {
  exigirSessao_(token);

  return {
    nome: 'RH | e Você',
    nomeCurto: 'RH | e Você',
    versao: '2.1',
    tema: '#174A78',
    fundo: '#F3F6F9',
    recursos: [
      'Layout adaptado para celular',
      'Atalho na tela inicial',
      'Indicador de conexão',
      'Cache local da interface',
      'Última ficha consultada disponível durante falha temporária de conexão'
    ]
  };
}



/* ==========================================================
   GESTÃO ADMINISTRATIVA - INCLUSÃO, EDIÇÃO E EXCLUSÃO
========================================================== */

function listarRegistrosAdministrativos(token, modulo) {
  exigirPerfil_(token, ['ADMINISTRADOR']);

  const ss = abrirBanco_();
  modulo = String(modulo || '').toUpperCase();

  const mapa = {
    COMUNICADOS: { aba: 'COMUNICADOS', colunas: 13 },
    DOCUMENTOS: { aba: 'DOCUMENTOS', colunas: 13 },
    TREINAMENTOS: { aba: 'TREINAMENTOS', colunas: 12 },
    AVALIACOES: { aba: 'AVALIACOES', colunas: 16 },
    INTEGRACOES: { aba: 'INTEGRACOES', colunas: 15 },
    USUARIOS: { aba: 'USUARIOS', colunas: 9 }
  };

  if (!mapa[modulo]) throw new Error('Módulo administrativo inválido.');

  const cfg = mapa[modulo];
  const aba = ss.getSheetByName(cfg.aba);
  if (!aba || aba.getLastRow() < 2) return [];

  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, Math.min(cfg.colunas, aba.getLastColumn())).getValues();

  return dados.map((l, i) => {
    if (modulo === 'COMUNICADOS') {
      return { linha:i+2, id:l[0], titulo:l[2], detalhe:l[3], status:l[8], data:formatarData_(l[1]) };
    }
    if (modulo === 'DOCUMENTOS') {
      return { linha:i+2, id:l[0], titulo:l[2], detalhe:l[3], status:l[9] || l[8], data:formatarData_(l[1]) };
    }
    if (modulo === 'TREINAMENTOS') {
      return { linha:i+2, id:l[0], titulo:l[1], detalhe:l[2], status:l[9], data:formatarData_(l[10]) };
    }
    if (modulo === 'AVALIACOES') {
      return { linha:i+2, id:l[0], titulo:l[4], detalhe:l[2] + ' - Nota ' + l[7], status:l[15], data:formatarData_(l[1]) };
    }
    if (modulo === 'INTEGRACOES') {
      return { linha:i+2, id:l[0], titulo:l[3], detalhe:l[4] + ' - ' + l[5], status:l[8], data:formatarData_(l[1]) };
    }
    return { linha:i+2, id:l[0], titulo:l[3] || l[2] || l[1], detalhe:l[1], status:l[6], data:formatarData_(l[8]) };
  }).reverse();
}

function excluirRegistroAdministrativo(token, dados) {
  const sessao = exigirPerfil_(token, ['ADMINISTRADOR']);

  if (!dados || !dados.modulo || !dados.linha) {
    throw new Error('Dados insuficientes para excluir o registro.');
  }

  const modulo = String(dados.modulo).toUpperCase();
  const mapa = {
    COMUNICADOS: 'COMUNICADOS',
    DOCUMENTOS: 'DOCUMENTOS',
    TREINAMENTOS: 'TREINAMENTOS',
    AVALIACOES: 'AVALIACOES',
    INTEGRACOES: 'INTEGRACOES',
    USUARIOS: 'USUARIOS'
  };

  const nomeAba = mapa[modulo];
  if (!nomeAba) throw new Error('Módulo inválido.');

  const aba = abrirBanco_().getSheetByName(nomeAba);
  const linha = Number(dados.linha);

  if (!aba || linha < 2 || linha > aba.getLastRow()) {
    throw new Error('Registro não encontrado.');
  }

  const id = aba.getRange(linha, 1).getValue();
  const descricao = aba.getRange(linha, 2, 1, Math.min(4, aba.getLastColumn())).getValues()[0].join(' | ');

  if (modulo === 'USUARIOS') {
    const email = String(aba.getRange(linha, 2).getValue()).toLowerCase();
    if (email === 'alexandra.rodrigues@ferramentariasoares.com.br') {
      throw new Error('O administrador principal não pode ser excluído.');
    }
  }

  aba.deleteRow(linha);

  abrirBanco_().getSheetByName('HISTORICO').appendRow([
    new Date(),
    Utilities.formatDate(new Date(), CONFIG.FUSO, 'HH:mm:ss'),
    '',
    sessao.nome,
    'REGISTRO_EXCLUIDO',
    modulo,
    id,
    descricao + ' | Excluído por: ' + sessao.email
  ]);

  return { sucesso:true, mensagem:'Registro excluído com sucesso.' };
}

function atualizarRegistroAdministrativo(token, dados) {
  const sessao = exigirPerfil_(token, ['ADMINISTRADOR']);

  if (!dados || !dados.modulo || !dados.linha) {
    throw new Error('Dados insuficientes para atualizar o registro.');
  }

  const modulo = String(dados.modulo).toUpperCase();
  const linha = Number(dados.linha);
  const ss = abrirBanco_();

  if (modulo === 'COMUNICADOS') {
    const aba = ss.getSheetByName('COMUNICADOS');
    aba.getRange(linha, 3).setValue(dados.titulo || '');
    aba.getRange(linha, 4).setValue(dados.detalhe || '');
    aba.getRange(linha, 9).setValue(dados.status || 'SIM');
  } else if (modulo === 'DOCUMENTOS') {
    const aba = ss.getSheetByName('DOCUMENTOS');
    aba.getRange(linha, 3).setValue(dados.titulo || '');
    aba.getRange(linha, 4).setValue(dados.detalhe || '');
    if (aba.getLastColumn() >= 10) aba.getRange(linha, 10).setValue(dados.status || 'ATIVO');
  } else if (modulo === 'TREINAMENTOS') {
    const aba = ss.getSheetByName('TREINAMENTOS');
    aba.getRange(linha, 2).setValue(dados.titulo || '');
    aba.getRange(linha, 3).setValue(dados.detalhe || '');
    aba.getRange(linha, 10).setValue(dados.status || 'ATIVO');
  } else if (modulo === 'USUARIOS') {
    const aba = ss.getSheetByName('USUARIOS');
    aba.getRange(linha, 4).setValue(dados.titulo || '');
    aba.getRange(linha, 7).setValue(dados.status || 'SIM');
  } else {
    throw new Error('Este módulo permite somente exclusão.');
  }

  ss.getSheetByName('HISTORICO').appendRow([
    new Date(),
    Utilities.formatDate(new Date(), CONFIG.FUSO, 'HH:mm:ss'),
    '',
    sessao.nome,
    'REGISTRO_ATUALIZADO',
    modulo,
    dados.id || '',
    'Atualizado por: ' + sessao.email
  ]);

  return { sucesso:true, mensagem:'Registro atualizado com sucesso.' };
}


/* PORTAL RH 2.0 */
function obterVersaoPortal2(token) {
  exigirSessao_(token);
  return {
    versao: '2.1',
    nome: 'RH | e Você',
    empresa: 'Ferramentaria Soares',
    atualizadoEm: Utilities.formatDate(new Date(), CONFIG.FUSO, 'dd/MM/yyyy HH:mm')
  };
}


/* ==========================================================
   PORTAL RH 3.0 - HOLERITES, SEGURANÇA, CURRÍCULOS,
   ANIVERSARIANTES E CARDÁPIO
========================================================== */

function salvarHolerite(token, dados) {
  exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);
  if (!dados || !dados.competencia || !dados.matricula) {
    throw new Error('Informe a competência e a matrícula.');
  }
  validarArquivoPdfHolerite_(dados);

  const colaborador = localizarColaboradorHolerite_(dados.matricula);
  if (!colaborador) {
    throw new Error('Matrícula não encontrada entre os colaboradores ativos.');
  }

  return gravarHolerite_(token, dados.competencia, colaborador, dados, false);
}

function analisarArquivosHolerites(token, competencia, nomesArquivos) {
  exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);
  validarCompetenciaHolerite_(competencia);

  if (!Array.isArray(nomesArquivos) || !nomesArquivos.length) {
    throw new Error('Selecione os arquivos PDF dos holerites.');
  }

  const colaboradores = listarColaboradoresAtivosHolerite_();
  const resultados = [];
  const usados = {};

  nomesArquivos.forEach(nomeArquivo => {
    const nome = String(nomeArquivo || '').trim();
    const item = { nomeArquivo: nome, matricula: '', colaborador: '', status: 'ERRO', mensagem: '' };

    if (!/\.pdf$/i.test(nome)) {
      item.mensagem = 'Arquivo ignorado: não é PDF.';
      resultados.push(item);
      return;
    }

    // Versão 3.4: identificação principal por nome e sobrenome.
    // A matrícula continua sendo usada apenas internamente para vincular o documento ao cadastro correto.
    const candidatos = colaboradores
      .map(c => ({ colaborador: c, score: pontuarNomeArquivoHolerite_(nome, c) }))
      .filter(x => x.score >= 70)
      .sort((a, b) => b.score - a.score || String(a.colaborador.nome).localeCompare(String(b.colaborador.nome), 'pt-BR'));

    if (!candidatos.length) {
      item.mensagem = 'Nome e sobrenome não identificados no nome do arquivo. Selecione o colaborador manualmente.';
    } else {
      const melhorScore = candidatos[0].score;
      const melhores = candidatos.filter(x => x.score === melhorScore);
      if (melhores.length > 1) {
        item.mensagem = 'Há mais de um colaborador compatível com este nome. Confirme manualmente.';
      } else {
        const c = melhores[0].colaborador;
        item.matricula = c.matricula;
        item.colaborador = c.nome;
        if (usados[String(c.matricula)]) {
          item.status = 'DUPLICADO';
          item.mensagem = 'Há mais de um PDF associado a este colaborador.';
        } else {
          item.status = 'PRONTO';
          item.mensagem = 'Nome identificado: ' + c.nome + '.';
          usados[String(c.matricula)] = true;
        }
      }
    }
    resultados.push(item);
  });

  const prontas = resultados.filter(r => r.status === 'PRONTO').length;
  const erros = resultados.filter(r => r.status !== 'PRONTO').length;
  const faltantesSelecao = colaboradores
    .filter(c => !usados[String(c.matricula)])
    .map(c => ({ matricula: c.matricula, nome: c.nome, setor: c.setor }));

  return {
    competencia: competencia,
    totalArquivos: nomesArquivos.length,
    prontos: prontas,
    erros: erros,
    resultados: resultados,
    faltantesSelecao: faltantesSelecao,
    totalColaboradoresAtivos: colaboradores.length,
    colaboradoresDisponiveis: colaboradores
      .slice()
      .sort((a, b) => String(a.nome).localeCompare(String(b.nome), 'pt-BR'))
  };
}

function salvarHoleriteLote(token, competencia, dadosArquivo) {
  exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);
  validarCompetenciaHolerite_(competencia);
  validarArquivoPdfHolerite_(dadosArquivo);

  // VERSAO 4.0
  // A distribuicao automatica e feita EXCLUSIVAMENTE pelo nome completo
  // identificado dentro do recibo. A matricula impressa na folha nao e usada
  // para decidir o destinatario, pois o Portal pode possuir identificadores
  // internos diferentes do sistema de folha.
  const nomeLido = String(dadosArquivo && dadosArquivo.nomeColaborador || '').trim();
  if (!nomeLido) {
    throw new Error('O nome completo do colaborador nao foi identificado no recibo. Envie este item para conferencia manual.');
  }

  const alvo = normalizarNomeHolerite_(nomeLido);
  const encontrados = listarColaboradoresAtivosHolerite_().filter(function(c) {
    return normalizarNomeHolerite_(c.nome) === alvo;
  });

  if (encontrados.length === 0) {
    throw new Error('Nome do recibo nao encontrado no cadastro ativo: ' + nomeLido + '. O holerite nao foi publicado.');
  }
  if (encontrados.length > 1) {
    throw new Error('Existe mais de um cadastro ativo com o nome ' + nomeLido + '. O holerite foi bloqueado para conferencia manual.');
  }

  const colaborador = encontrados[0];
  return gravarHolerite_(token, competencia, colaborador, dadosArquivo, true);
}

function obterColaboradoresParaImportacaoHolerites(token) {
  exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);
  return listarColaboradoresAtivosHolerite_()
    .slice()
    .sort((a, b) => String(a.nome).localeCompare(String(b.nome), 'pt-BR'));
}

function salvarHoleriteManualDivergencia(token, competencia, dadosArquivo) {
  exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);
  validarCompetenciaHolerite_(competencia);
  validarArquivoPdfHolerite_(dadosArquivo);
  if (!dadosArquivo || !dadosArquivo.matricula) {
    throw new Error('Selecione o colaborador para a publicação manual.');
  }
  const colaborador = localizarColaboradorHolerite_(dadosArquivo.matricula);
  if (!colaborador) throw new Error('Colaborador não encontrado entre os ativos.');
  return gravarHolerite_(token, competencia, colaborador, dadosArquivo, false);
}


function obterVersaoPortalHolerites(token) {
  exigirSessao_(token);
  return { versao: '4.7.3', build: '2026-08-11', modulo: 'HOLERITES' };
}

function verificarPublicacaoHolerites(token) {
  const sessao = exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);
  const props = PropertiesService.getScriptProperties();

  const planilhaId = props.getProperty('PLANILHA_ID');
  if (!planilhaId) throw new Error('Banco de dados não configurado. Execute instalarEtapa5().');

  const ss = SpreadsheetApp.openById(planilhaId);
  const aba = ss.getSheetByName('HOLERITES');
  if (!aba) throw new Error('A aba HOLERITES não existe. Execute instalarEtapa5() para atualizar a estrutura.');

  let pasta = null;
  let pastaId = props.getProperty('PASTA_HOLERITES_ID');
  if (pastaId) {
    try { pasta = DriveApp.getFolderById(pastaId); } catch (e) { pasta = null; }
  }

  if (!pasta) {
    let pastaRaiz = null;
    const raizId = props.getProperty('PASTA_RAIZ_ID');
    if (raizId) {
      try { pastaRaiz = DriveApp.getFolderById(raizId); } catch (e) { pastaRaiz = null; }
    }
    if (!pastaRaiz) pastaRaiz = obterOuCriarPasta_(DriveApp.getRootFolder(), CONFIG.PASTA_RAIZ);
    pasta = obterOuCriarPasta_(pastaRaiz, '11 - Holerites');
    props.setProperty('PASTA_HOLERITES_ID', pasta.getId());
  }

  // Teste real de escrita. O arquivo temporário é enviado para a lixeira logo em seguida.
  const teste = pasta.createFile(Utilities.newBlob('teste', 'text/plain', 'teste_publicacao_portal_rh.txt'));
  const pastaUrl = pasta.getUrl();
  try { teste.setTrashed(true); } catch (e) {}

  return {
    sucesso: true,
    mensagem: 'Servidor pronto para publicar holerites.',
    pastaUrl: pastaUrl,
    usuario: sessao.email
  };
}

function localizarColaboradorHoleritePorNomeExato_(nome) {
  const alvo = normalizarNomeHolerite_(nome);
  if (!alvo) return null;
  const encontrados = listarColaboradoresAtivosHolerite_().filter(c => normalizarNomeHolerite_(c.nome) === alvo);
  return encontrados.length === 1 ? encontrados[0] : null;
}

function obterResumoCompetenciaHolerites(token, competencia) {
  exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);
  validarCompetenciaHolerite_(competencia);

  const colaboradores = listarColaboradoresAtivosHolerite_();
  const usuarios = listarUsuariosAtivosHolerite_();
  const aba = abrirBanco_().getSheetByName('HOLERITES');
  const registros = [];

  if (aba && aba.getLastRow() >= 2) {
    aba.getRange(2, 1, aba.getLastRow() - 1, 10).getValues().forEach(function(l) {
      if (normalizarCompetenciaHolerite_(l[1]) === normalizarCompetenciaHolerite_(competencia)) registros.push(l);
    });
  }

  function localizarRegistro(c) {
    const mat = normalizarMatriculaHolerite_(c.matricula);
    const nome = normalizarNomeHolerite_(c.nome);
    return registros.find(function(l) {
      const matReg = normalizarMatriculaHolerite_(l[2]);
      const nomeReg = normalizarNomeHolerite_(l[3]);
      return (mat && matReg && mat === matReg) || (nome && nomeReg && nome === nomeReg);
    }) || null;
  }

  const faltantes = [];
  const naoDisponiveis = [];
  let publicados = 0;
  let disponiveisNoApp = 0;
  let visualizados = 0;

  colaboradores.forEach(function(c) {
    const registro = localizarRegistro(c);
    if (!registro) {
      faltantes.push({ matricula: c.matricula, nome: c.nome, setor: c.setor });
      return;
    }
    publicados++;
    const usuario = localizarUsuarioEntregaHolerite_(c, usuarios);
    if (usuario) {
      disponiveisNoApp++;
      if (String(registro[9] || '').toUpperCase() === 'SIM' || registro[8]) visualizados++;
    } else {
      naoDisponiveis.push({ matricula: c.matricula, nome: c.nome, setor: c.setor });
    }
  });

  return {
    versao: '4.7.3',
    competencia: competencia,
    totalAtivos: colaboradores.length,
    publicados: publicados,
    faltantes: faltantes.length,
    listaFaltantes: faltantes,
    disponiveisNoApp: disponiveisNoApp,
    visualizados: visualizados,
    naoDisponiveisNoApp: naoDisponiveis.length,
    listaNaoDisponiveisNoApp: naoDisponiveis
  };
}

function gravarHolerite_(token, competencia, colaborador, dadosArquivo, origemLote) {
  const sessao = exigirPerfil_(token, ['RH', 'ADMINISTRADOR']);
  validarCompetenciaHolerite_(competencia);

  const pastaRaizId = PropertiesService.getScriptProperties().getProperty('PASTA_HOLERITES_ID');
  if (!pastaRaizId) throw new Error('Pasta de holerites não configurada. Execute instalarEtapa5().');

  const pastaRaiz = DriveApp.getFolderById(pastaRaizId);
  const ano = String(competencia).slice(0, 4);
  const pastaAno = obterOuCriarPasta_(pastaRaiz, ano);
  const pastaMes = obterOuCriarPasta_(pastaAno, String(competencia));
  const id = gerarProtocolo_('HOL');
  const nomeFinal = String(colaborador.matricula) + ' - ' + sanitizarNome_(colaborador.nome) + ' - ' + competencia + '.pdf';
  const bytes = Utilities.base64Decode(dadosArquivo.arquivoBase64);
  const arquivo = pastaMes.createFile(Utilities.newBlob(bytes, 'application/pdf', nomeFinal));

  const aba = abrirBanco_().getSheetByName('HOLERITES');
  let linhaExistente = 0;
  let arquivoAnteriorId = '';
  const matAlvo = normalizarMatriculaHolerite_(colaborador.matricula);
  const nomeAlvo = normalizarNomeHolerite_(colaborador.nome);

  if (aba.getLastRow() >= 2) {
    const valores = aba.getRange(2, 1, aba.getLastRow() - 1, 10).getValues();
    for (let i = 0; i < valores.length; i++) {
      const mesmaCompetencia = normalizarCompetenciaHolerite_(valores[i][1]) === normalizarCompetenciaHolerite_(competencia);
      const mesmaMatricula = matAlvo && normalizarMatriculaHolerite_(valores[i][2]) === matAlvo;
      const mesmoNome = nomeAlvo && normalizarNomeHolerite_(valores[i][3]) === nomeAlvo;
      if (mesmaCompetencia && (mesmaMatricula || mesmoNome)) {
        linhaExistente = i + 2;
        arquivoAnteriorId = valores[i][4] || '';
        break;
      }
    }
  }

  const registro = [
    linhaExistente ? aba.getRange(linhaExistente, 1).getValue() : id,
    competencia, colaborador.matricula, colaborador.nome,
    arquivo.getId(), arquivo.getUrl(), new Date(), sessao.email, '', 'NÃO'
  ];

  let linhaGravada;
  if (linhaExistente) {
    aba.getRange(linhaExistente, 1, 1, 10).setValues([registro]);
    linhaGravada = linhaExistente;
  } else {
    aba.appendRow(registro);
    linhaGravada = aba.getLastRow();
  }
  SpreadsheetApp.flush();

  // 4.7.3: valida a gravação pelo protocolo único, e não pela posição da linha.
  // Em planilhas maiores, getLastRow()/appendRow pode produzir uma leitura imediata
  // inconsistente. Por isso fazemos até 4 tentativas de localização do protocolo.
  let confirmado = null;
  let linhaConfirmada = 0;
  for (let tentativa = 0; tentativa < 4 && !confirmado; tentativa++) {
    SpreadsheetApp.flush();
    if (tentativa > 0) Utilities.sleep(250);
    const ultima = aba.getLastRow();
    if (ultima >= 2) {
      const ids = aba.getRange(2, 1, ultima - 1, 1).getDisplayValues();
      for (let i = ids.length - 1; i >= 0; i--) {
        if (String(ids[i][0]).trim() === String(registro[0]).trim()) {
          linhaConfirmada = i + 2;
          confirmado = aba.getRange(linhaConfirmada, 1, 1, 10).getValues()[0];
          break;
        }
      }
    }
  }

  const gravacaoOk = confirmado &&
    String(confirmado[0]).trim() === String(registro[0]).trim() &&
    normalizarCompetenciaHolerite_(confirmado[1]) === normalizarCompetenciaHolerite_(competencia) &&
    normalizarNomeHolerite_(confirmado[3]) === nomeAlvo &&
    String(confirmado[4]).trim() === String(arquivo.getId()).trim();

  if (!gravacaoOk) {
    // Não apagamos o PDF automaticamente: preservamos o arquivo para auditoria
    // e retornamos um diagnóstico detalhado para o RH.
    const detalhe = confirmado
      ? 'Registro localizado, mas os dados relidos não conferem. Protocolo=' + confirmado[0] +
        ', competência=' + confirmado[1] + ', colaborador=' + confirmado[3] +
        ', arquivoId=' + confirmado[4] + '.'
      : 'O protocolo ' + registro[0] + ' não foi localizado na aba HOLERITES após 4 tentativas.';
    throw new Error('Falha na validação da gravação do holerite. ' + detalhe);
  }

  linhaGravada = linhaConfirmada || linhaGravada;

  if (linhaExistente && arquivoAnteriorId && arquivoAnteriorId !== arquivo.getId()) {
    try { DriveApp.getFileById(arquivoAnteriorId).setTrashed(true); } catch (e) {}
  }

  // 4.7.2: depois que a linha foi gravada e relida com sucesso, falhas em
  // verificações auxiliares NÃO podem transformar uma publicação válida em erro técnico.
  let usuarioEntrega = null;
  let avisoEntrega = '';
  try {
    usuarioEntrega = localizarUsuarioEntregaHolerite_(colaborador);
  } catch (e) {
    avisoEntrega = 'Não foi possível validar o usuário do aplicativo: ' + (e && e.message ? e.message : e);
  }
  const disponivelNoAplicativo = Boolean(usuarioEntrega);

  let avisoHistorico = '';
  try {
    registrarHistorico_(abrirBanco_(), { matricula: colaborador.matricula, nome: colaborador.nome },
      registro[0], linhaExistente ? 'SUBSTITUICAO' : 'PUBLICACAO', 'HOLERITES',
      (origemLote ? 'Importação em lote' : 'Publicação individual') + ' da competência ' + competencia + ' por ' + sessao.email +
      '. Disponível no app: ' + (disponivelNoAplicativo ? 'SIM' : 'NÃO') + '.');
  } catch (e) {
    avisoHistorico = 'O holerite foi salvo, mas o histórico não pôde ser registrado: ' + (e && e.message ? e.message : e);
  }

  return {
    sucesso: true,
    versao: '4.7.3',
    protocolo: registro[0],
    matricula: colaborador.matricula,
    nome: colaborador.nome,
    substituido: Boolean(linhaExistente),
    disponivelNoAplicativo: disponivelNoAplicativo,
    usuarioEntrega: usuarioEntrega ? { matricula: usuarioEntrega.matricula, nome: usuarioEntrega.nome, email: usuarioEntrega.email } : null,
    avisos: [avisoEntrega, avisoHistorico].filter(Boolean),
    mensagem: disponivelNoAplicativo
      ? 'Holerite gravado e validado. Disponível no aplicativo.'
      : 'Holerite gravado e validado. O vínculo com o login ainda precisa ser conferido.'
  };
}

function validarArquivoPdfHolerite_(dados) {
  if (!dados || !dados.arquivoBase64) throw new Error('Selecione um arquivo PDF.');
  const nome = String(dados.nomeArquivo || 'holerite.pdf');
  const mime = String(dados.mimeType || '').toLowerCase();
  if (!/\.pdf$/i.test(nome) && mime !== 'application/pdf') {
    throw new Error('Somente arquivos PDF são permitidos.');
  }
  const bytes = Utilities.base64Decode(dados.arquivoBase64);
  if (bytes.length > 15 * 1024 * 1024) {
    throw new Error('Cada holerite separado deve ter no máximo 15 MB.');
  }
}

function validarCompetenciaHolerite_(competencia) {
  if (!/^\d{4}-\d{2}$/.test(String(competencia || ''))) {
    throw new Error('Informe uma competência válida.');
  }
}

// 4.7.4: a planilha pode converter 2026-07 em uma data (01/07/2026).
// Esta função transforma string, Date e valor de célula no mesmo formato YYYY-MM.
function normalizarCompetenciaHolerite_(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, CONFIG.FUSO, 'yyyy-MM');
  }
  const texto = String(valor == null ? '' : valor).trim();
  if (/^\d{4}-\d{2}$/.test(texto)) return texto;
  const d = new Date(valor);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, CONFIG.FUSO, 'yyyy-MM');
  }
  return texto;
}

function listarColaboradoresAtivosHolerite_() {
  const aba = abrirBanco_().getSheetByName('COLABORADORES');
  if (!aba || aba.getLastRow() < 2) return [];

  const cab = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0]
    .map(v => String(v || '').trim().toUpperCase());
  const iMat = cab.indexOf('MATRICULA');
  const iNome = cab.indexOf('NOME');
  const iSetor = cab.indexOf('SETOR');
  const iStatus = cab.indexOf('STATUS');
  if (iMat < 0 || iNome < 0) throw new Error('A aba COLABORADORES precisa ter MATRICULA e NOME.');

  return aba.getRange(2, 1, aba.getLastRow() - 1, aba.getLastColumn()).getValues()
    .filter(l => {
      const status = iStatus >= 0 ? String(l[iStatus] || 'ATIVO').trim().toUpperCase() : 'ATIVO';
      return l[iMat] && l[iNome] && !['INATIVO', 'DESLIGADO', 'DEMITIDO'].includes(status);
    })
    .map(l => ({
      matricula: String(l[iMat]).trim(),
      nome: String(l[iNome]).trim(),
      setor: iSetor >= 0 ? String(l[iSetor] || '').trim() : ''
    }));
}

function localizarColaboradorHolerite_(matricula) {
  const busca = normalizarMatriculaHolerite_(matricula);
  return listarColaboradoresAtivosHolerite_().find(c =>
    normalizarMatriculaHolerite_(c.matricula) === busca
  ) || null;
}

function normalizarMatriculaHolerite_(valor) {
  const texto = String(valor || '').trim().toUpperCase();
  if (/^\d+$/.test(texto)) return String(Number(texto));
  return texto.replace(/[^A-Z0-9]/g, '');
}

function nomeArquivoContemMatricula_(nomeArquivo, matricula) {
  const nome = String(nomeArquivo || '').replace(/\.pdf$/i, '');
  const mat = String(matricula || '').trim();
  if (!mat) return false;

  if (/^\d+$/.test(mat)) {
    const matNormal = String(Number(mat));
    const grupos = nome.match(/\d+/g) || [];
    return grupos.some(g => String(Number(g)) === matNormal);
  }

  const nomeNormal = nome.toUpperCase().replace(/[^A-Z0-9]/g, '');
  return nomeNormal.indexOf(normalizarMatriculaHolerite_(mat)) >= 0;
}

function normalizarNomeHolerite_(valor) {
  return String(valor || '')
    .replace(/\.pdf$/i, '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/[_\-.]+/g, ' ')
    .replace(/[^A-Z0-9 ]/g, ' ')
    .replace(/\b(HOLERITE|HOLERITES|CONTRACHEQUE|RECIBO|SALARIO|PAGAMENTO|FOLHA)\b/g, ' ')
    .replace(/\b(20\d{2}|0?[1-9]|1[0-2])\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function tokensNomeHolerite_(valor) {
  const ignorar = { DE:true, DA:true, DO:true, DAS:true, DOS:true, E:true };
  return normalizarNomeHolerite_(valor)
    .split(' ')
    .filter(t => t && !ignorar[t] && !/^\d+$/.test(t));
}

function pontuarNomeArquivoHolerite_(nomeArquivo, colaborador) {
  const arquivo = normalizarNomeHolerite_(nomeArquivo);
  const tokensArquivo = tokensNomeHolerite_(nomeArquivo);
  const tokensNome = tokensNomeHolerite_(colaborador.nome);
  if (!arquivo || tokensNome.length < 2) return 0;

  const presentes = tokensNome.filter(t => tokensArquivo.includes(t));
  const primeiro = tokensNome[0];
  const ultimo = tokensNome[tokensNome.length - 1];
  const temPrimeiroUltimo = tokensArquivo.includes(primeiro) && tokensArquivo.includes(ultimo);
  const todos = presentes.length === tokensNome.length;

  if (todos && temPrimeiroUltimo) return 100;
  if (temPrimeiroUltimo) return 90;

  // Quando a coluna MATRÍCULA já usa nome.sobrenome (ex.: ABNER.ANDRADE),
  // ela também pode ajudar na identificação sem exigir matrícula numérica no PDF.
  const matTokens = tokensNomeHolerite_(colaborador.matricula);
  if (matTokens.length >= 2 && matTokens.every(t => tokensArquivo.includes(t))) return 85;

  if (tokensNome.length >= 3 && presentes.length >= 3 && tokensArquivo.includes(primeiro)) return 75;
  return 0;
}

function listarUsuariosAtivosHolerite_() {
  const aba = abrirBanco_().getSheetByName('USUARIOS');
  if (!aba || aba.getLastRow() < 2) return [];
  return aba.getRange(2, 1, aba.getLastRow() - 1, 10).getValues()
    .filter(l => String(l[6] || '').trim().toUpperCase() === 'SIM')
    .map(l => ({
      id: l[0],
      email: String(l[1] || '').trim(),
      matricula: String(l[2] || '').trim(),
      nome: String(l[3] || '').trim(),
      perfil: String(l[4] || '').trim()
    }));
}

function localizarUsuarioEntregaHolerite_(colaborador, usuarios) {
  usuarios = usuarios || listarUsuariosAtivosHolerite_();
  const mat = normalizarMatriculaHolerite_(colaborador && colaborador.matricula);
  const porMatricula = usuarios.filter(u => normalizarMatriculaHolerite_(u.matricula) === mat);
  if (porMatricula.length === 1) return porMatricula[0];

  const nome = normalizarNomeHolerite_(colaborador && colaborador.nome);
  if (!nome) return null;
  const porNome = usuarios.filter(u => normalizarNomeHolerite_(u.nome) === nome);
  return porNome.length === 1 ? porNome[0] : null;
}

function holeritePertenceSessao_(linha, sessao) {
  const matRegistro = normalizarMatriculaHolerite_(linha && linha[2]);
  const matSessao = normalizarMatriculaHolerite_(sessao && sessao.matricula);
  if (matRegistro && matSessao && matRegistro === matSessao) return true;

  const nomeRegistro = normalizarNomeHolerite_(linha && linha[3]);
  const nomeSessao = normalizarNomeHolerite_(sessao && sessao.nome);
  return Boolean(nomeRegistro && nomeSessao && nomeRegistro === nomeSessao);
}

function listarMeusHolerites(token) {
  const sessao = exigirSessao_(token);
  const aba = abrirBanco_().getSheetByName('HOLERITES');
  if (!aba || aba.getLastRow() < 2) return [];
  return aba.getRange(2,1,aba.getLastRow()-1,10).getValues()
    .filter(l => holeritePertenceSessao_(l, sessao))
    .reverse().map(l=>({
      id:String(l[0] || ''),
      // 4.7.5: nunca devolver Date pelo google.script.run. A planilha pode
      // transformar a competência em 01/07/2026; convertemos sempre para YYYY-MM.
      competencia:normalizarCompetenciaHolerite_(l[1]),
      colaborador:String(l[3] || ''),
      arquivoUrl:String(l[5] || ''),
      publicadoEm:formatarData_(l[6]),
      visualizadoEm:formatarData_(l[8]),
      confirmado:String(l[9] || '')
    }));
}

function obterHoleritePdfSeguro(token, id) {
  const sessao = exigirSessao_(token);
  const aba = abrirBanco_().getSheetByName('HOLERITES');
  if (!aba || aba.getLastRow() < 2) throw new Error('Holerite não encontrado para este usuário.');

  const dados = aba.getRange(2,1,aba.getLastRow()-1,10).getValues();
  for (let i = 0; i < dados.length; i++) {
    const linha = dados[i];
    if (String(linha[0]) !== String(id)) continue;
    if (!holeritePertenceSessao_(linha, sessao)) {
      throw new Error('Você não possui permissão para visualizar este holerite.');
    }

    const arquivoId = String(linha[4] || '').trim();
    if (!arquivoId) throw new Error('Arquivo do holerite não localizado.');

    let arquivo;
    try {
      arquivo = DriveApp.getFileById(arquivoId);
    } catch (e) {
      throw new Error('O PDF do holerite não foi encontrado no armazenamento.');
    }

    const blob = arquivo.getBlob();
    const bytes = blob.getBytes();
    if (!bytes || !bytes.length) throw new Error('O PDF do holerite está vazio.');

    // Registra a visualização somente depois que o servidor validou o usuário
    // e conseguiu ler o arquivo privado com sucesso.
    aba.getRange(i + 2, 9).setValue(new Date());
    aba.getRange(i + 2, 10).setValue('SIM');
    SpreadsheetApp.flush();

    return {
      sucesso: true,
      id: String(linha[0] || ''),
      competencia: normalizarCompetenciaHolerite_(linha[1]),
      colaborador: String(linha[3] || ''),
      nomeArquivo: arquivo.getName() || ('Holerite-' + normalizarCompetenciaHolerite_(linha[1]) + '.pdf'),
      mimeType: blob.getContentType() || 'application/pdf',
      arquivoBase64: Utilities.base64Encode(bytes)
    };
  }

  throw new Error('Holerite não encontrado para este usuário.');
}

function confirmarHolerite(token, id) {
  const sessao = exigirSessao_(token);
  const aba = abrirBanco_().getSheetByName('HOLERITES');
  if (!aba || aba.getLastRow() < 2) throw new Error('Holerite não encontrado para este usuário.');
  const dados = aba.getRange(2,1,aba.getLastRow()-1,10).getValues();
  for (let i=0;i<dados.length;i++) {
    if (String(dados[i][0])===String(id) && holeritePertenceSessao_(dados[i], sessao)) {
      aba.getRange(i+2,9).setValue(new Date());
      aba.getRange(i+2,10).setValue('SIM');
      return {sucesso:true,mensagem:'Visualização confirmada.'};
    }
  }
  throw new Error('Holerite não encontrado para este usuário.');
}

function salvarOcorrenciaSeguranca(token, dados) {
  const sessao = exigirSessao_(token);
  validarDadosBasicos_(dados);
  if (!dados.setor || !dados.tipo || !dados.descricao) throw new Error('Preencha setor, tipo e descrição.');
  validarArquivo_(dados, false);
  const protocolo = gerarProtocolo_('SST');
  let arquivoId='', arquivoUrl='';
  if (dados.arquivoBase64) {
    const pasta = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('PASTA_SEGURANCA_ID'));
    const arquivo = salvarArquivo_(pasta,dados,protocolo,dados.nome);
    arquivoId=arquivo.getId(); arquivoUrl=arquivo.getUrl();
  }
  abrirBanco_().getSheetByName('SEGURANCA_TRABALHO').appendRow([
    protocolo,new Date(),dados.matricula,dados.nome,dados.setor,dados.tipo,dados.descricao,
    dados.local||'',dados.urgencia||'Normal',arquivoId,arquivoUrl,'ABERTA','', '', ''
  ]);
  enviarAvisoRH_('[Portal RH] Segurança do Trabalho - '+protocolo,
    'Nova ocorrência de Segurança do Trabalho.\n\nColaborador: '+dados.nome+'\nSetor: '+dados.setor+'\nTipo: '+dados.tipo+'\nDescrição: '+dados.descricao);
  return {sucesso:true,protocolo,mensagem:'Registro enviado com sucesso.'};
}

function salvarCurriculo(dados) {
  if (!dados || !dados.nome || !dados.telefone || !dados.areaInteresse) {
    throw new Error('Informe nome, telefone e área de interesse.');
  }
  validarArquivo_(dados, true);
  const protocolo=gerarProtocolo_('CV');
  const pasta=DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('PASTA_CURRICULOS_ID'));
  const arquivo=salvarArquivo_(pasta,dados,protocolo,dados.nome);
  abrirBanco_().getSheetByName('CURRICULOS').appendRow([
    protocolo,new Date(),dados.nome,dados.telefone,dados.email||'',dados.cidade||'',
    dados.areaInteresse,dados.experiencia||'',arquivo.getId(),arquivo.getUrl(),'RECEBIDO',''
  ]);
  enviarAvisoRH_('[Portal RH] Novo currículo - '+protocolo,
    'Novo currículo recebido.\n\nNome: '+dados.nome+'\nTelefone: '+dados.telefone+'\nÁrea: '+dados.areaInteresse);
  return {sucesso:true,protocolo,mensagem:'Currículo enviado com sucesso.'};
}

function obterAniversariantes(token) {
  exigirSessao_(token);
  const aba=abrirBanco_().getSheetByName('COLABORADORES');
  if (!aba || aba.getLastRow()<2) return {hoje:[],mes:[]};
  const hoje=new Date();
  const lista=aba.getRange(2,1,aba.getLastRow()-1,13).getValues()
    .filter(l=>l[6] && !['INATIVO','DESLIGADO','DEMITIDO'].includes(String(l[11]).toUpperCase()))
    .map(l=>{const d=new Date(l[6]);return {nome:l[2],setor:l[9]||'',dia:d.getDate(),mes:d.getMonth()+1};})
    .filter(x=>x.mes===hoje.getMonth()+1)
    .sort((a,b)=>a.dia-b.dia);
  return {hoje:lista.filter(x=>x.dia===hoje.getDate()),mes:lista};
}

function salvarItemCardapio(token,dados) {
  const sessao=exigirPerfil_(token,['RH','ADMINISTRADOR']);
  if (!dados || !dados.semanaInicio || !dados.diaSemana || !dados.pratoPrincipal) {
    throw new Error('Informe semana, dia e prato principal.');
  }
  const id=gerarProtocolo_('CAR');
  abrirBanco_().getSheetByName('CARDAPIO').appendRow([
    id,dados.semanaInicio,dados.diaSemana,dados.data||'',dados.pratoPrincipal,
    dados.acompanhamentos||'',dados.salada||'',dados.sobremesa||'',dados.observacao||'',
    dados.publicado||'SIM',sessao.email
  ]);
  return {sucesso:true,protocolo:id,mensagem:'Item do cardápio publicado.'};
}

function listarCardapioSemana(token) {
  exigirSessao_(token);
  const aba=abrirBanco_().getSheetByName('CARDAPIO');
  if (!aba || aba.getLastRow()<2) return [];
  const hoje=new Date(); hoje.setHours(0,0,0,0);
  const segunda=new Date(hoje); const dia=segunda.getDay(); segunda.setDate(segunda.getDate()-(dia===0?6:dia-1));
  const chave=Utilities.formatDate(segunda,CONFIG.FUSO,'yyyy-MM-dd');
  return aba.getRange(2,1,aba.getLastRow()-1,11).getValues()
    .filter(l=>String(l[1])===chave && String(l[9]).toUpperCase()==='SIM')
    .map(l=>({id:l[0],diaSemana:l[2],data:l[3],pratoPrincipal:l[4],acompanhamentos:l[5],salada:l[6],sobremesa:l[7],observacao:l[8]}));
}
