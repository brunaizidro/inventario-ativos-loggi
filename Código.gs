/************************************************************
 * SISTEMA DE INVENTÁRIO DE ATIVOS - LOGGI
 * ARQUIVO: Código.gs
 ************************************************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Inventário de Ativos")
    .addItem("Configurar sistema", "configurarSistema")
    .addItem("Corrigir estrutura antiga", "corrigirEstruturaAntiga")
    .addSeparator()
    .addItem("Testar usuário logado", "testarUsuarioLogado")
    .addItem("Abrir sistema", "abrirSistema")
    .addToUi();

  menuAgenteIA_(SpreadsheetApp.getUi());
}


/**
 * Abre Login ou Index.
 */
function doGet(e) {

  const pagina =
    e && e.parameter
      ? e.parameter.pagina
      : "";

  if (pagina === "sistema" || pagina === "inventario") {

    return HtmlService
      .createTemplateFromFile("Index")
      .evaluate()
      .setTitle("Inventário de Ativos")
      .setXFrameOptionsMode(
        HtmlService.XFrameOptionsMode.ALLOWALL
      );

  }

  if (pagina === "dashboard") {

    return HtmlService
      .createTemplateFromFile("Dashboard")
      .evaluate()
      .setTitle("Dashboard - Inventário de Ativos")
      .setXFrameOptionsMode(
        HtmlService.XFrameOptionsMode.ALLOWALL
      );

  }

  return HtmlService
    .createTemplateFromFile("Login")
    .evaluate()
    .setTitle("Login - Inventário de Ativos")
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


/**
 * Configura todas as abas.
 */
function configurarSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const estruturas = {
    BASES: [
      "NOME_BASE",
      "DESCRICAO",
      "STATUS",
      "DATA_CADASTRO"
    ],

    USUARIOS: [
      "EMAIL",
      "NOME",
      "PERFIL",
      "STATUS",
      "DATA_CADASTRO"
    ],

    PERMISSOES: [
      "EMAIL",
      "NOME_BASE",
      "STATUS",
      "DATA_CADASTRO"
    ],

    ATIVOS_ESPERADOS: [
      "CODIGO_ATIVO",
      "TIPO_ATIVO",
      "NOME_BASE",
      "STATUS",
      "ORIGEM",
      "DATA_CADASTRO",
      "ULTIMA_ATUALIZACAO"
    ],

    INVENTARIOS: [
      "NOME_BASE",
      "TIPO_INVENTARIO",
      "SEMANA_REFERENCIA",
      "DATA_INICIO",
      "DATA_FIM",
      "RESPONSAVEL",
      "STATUS",
      "TOTAL_BIPS",
      "OBSERVACAO"
    ],

    BIPS: [
      "NOME_BASE",
      "TIPO_INVENTARIO",
      "SEMANA_REFERENCIA",
      "CODIGO_ATIVO",
      "TIPO_ATIVO",
      "STATUS_ATIVO",
      "DATA_HORA",
      "RESPONSAVEL",
      "RESULTADO"
    ],

    RESULTADOS: [
      "NOME_BASE",
      "SEMANA_REFERENCIA",
      "CODIGO_ATIVO",
      "TIPO_ATIVO",
      "RESULTADO",
      "DATA_PROCESSAMENTO"
    ],

    TRATATIVAS: [
      "NOME_BASE",
      "SEMANA_REFERENCIA",
      "CODIGO_ATIVO",
      "TIPO_ATIVO",
      "TIPO_TRATATIVA",
      "STATUS",
      "RESPONSAVEL",
      "JUSTIFICATIVA",
      "DATA_ABERTURA",
      "DATA_CONCLUSAO"
    ],

    MOVIMENTACOES: [
      "CODIGO_ATIVO",
      "TIPO_ATIVO",
      "BASE_ORIGEM",
      "BASE_DESTINO",
      "STATUS",
      "SOLICITANTE",
      "APROVADOR",
      "JUSTIFICATIVA",
      "DATA_SOLICITACAO",
      "DATA_APROVACAO"
    ],

    CONFIGURACOES: [
      "PARAMETRO",
      "VALOR",
      "DESCRICAO"
    ]
  };

  Object.keys(estruturas).forEach(nomeAba => {
    let aba = ss.getSheetByName(nomeAba);

    if (!aba) {
      aba = ss.insertSheet(nomeAba);
    }

    const cabecalhos = estruturas[nomeAba];
    const totalColunas = Math.max(
      aba.getLastColumn(),
      cabecalhos.length
    );

    aba
      .getRange(1, 1, 1, totalColunas)
      .clearContent()
      .clearFormat();

    aba
      .getRange(1, 1, 1, cabecalhos.length)
      .setValues([cabecalhos]);

    formatarAba_(aba, cabecalhos.length);
  });

  SpreadsheetApp.getUi().alert(
    "Sistema configurado com sucesso!"
  );
}


/**
 * Formata as abas.
 */
function formatarAba_(aba, quantidadeColunas) {
  const cabecalho = aba.getRange(
    1,
    1,
    1,
    1,
    quantidadeColunas
  );

  cabecalho
    .setFontWeight("bold")
    .setBackground("#006aff")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  aba.setFrozenRows(1);
  aba.setRowHeight(1, 30);

  for (let coluna = 1; coluna <= quantidadeColunas; coluna++) {
    aba.autoResizeColumn(coluna);
  }
}


/**
 * Obtém o e-mail do usuário logado.
 */
function obterUsuarioLogado_() {
  const email = Session.getActiveUser().getEmail();

  if (!email) {
    throw new Error(
      "Não foi possível identificar o usuário logado. " +
      "Verifique se o Web App está configurado para exigir login."
    );
  }

  return email.toLowerCase().trim();
}


/**
 * Testa o usuário logado.
 */
function testarUsuarioLogado() {
  const email = obterUsuarioLogado_();

  SpreadsheetApp.getUi().alert(
    "Usuário identificado:\n\n" + email
  );
}


/**
 * Valida o login.
 */
function validarLogin() {
  const email = obterUsuarioLogado_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("USUARIOS");

  if (!aba) {
    throw new Error(
      "A aba USUARIOS não foi encontrada."
    );
  }

  const dados = aba.getDataRange().getValues();

  if (dados.length < 2) {
    return {
      autorizado: false,
      mensagem: "Nenhum usuário foi cadastrado."
    };
  }

  const cabecalho = dados[0].map(valor =>
    String(valor || "")
      .trim()
      .toUpperCase()
  );

  const colunaEmail = cabecalho.indexOf("EMAIL");
  const colunaNome = cabecalho.indexOf("NOME");
  const colunaPerfil = cabecalho.indexOf("PERFIL");
  const colunaStatus = cabecalho.indexOf("STATUS");

  if (colunaEmail === -1 || colunaStatus === -1) {
    throw new Error(
      "A aba USUARIOS precisa conter as colunas EMAIL e STATUS."
    );
  }

  for (let i = 1; i < dados.length; i++) {
    const emailPlanilha = String(
      dados[i][colunaEmail] || ""
    )
      .toLowerCase()
      .trim();

    const status = String(
      dados[i][colunaStatus] || ""
    )
      .toUpperCase()
      .trim();

    if (emailPlanilha === email) {
      if (status !== "ATIVO" && status !== "ATIVA") {
        return {
          autorizado: false,
          mensagem: "Seu usuário está inativo."
        };
      }

      return {
        autorizado: true,
        mensagem: "Acesso autorizado.",
        usuario: {
          email: email,
          nome: colunaNome >= 0
            ? dados[i][colunaNome]
            : "",
          perfil: colunaPerfil >= 0
            ? dados[i][colunaPerfil]
            : ""
        }
      };
    }
  }

  return {
    autorizado: false,
    mensagem:
      "Seu e-mail não está autorizado para acessar o sistema."
  };
}


/**
 * Valida usuário ativo.
 */
function validarUsuarioAtivo_() {
  const login = validarLogin();

  if (!login.autorizado) {
    throw new Error(login.mensagem);
  }

  return login.usuario;
}


/**
 * Retorna todas as bases ativas.
 */
function obterTodasBasesAtivas_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("BASES");

  if (!aba) {
    throw new Error(
      "A aba BASES não foi encontrada."
    );
  }

  const dados = aba.getDataRange().getValues();
  const bases = [];

  for (let i = 1; i < dados.length; i++) {
    const nomeBase = String(
      dados[i][0] || ""
    ).trim();

    const status = String(
      dados[i][2] || ""
    )
      .toUpperCase()
      .trim();

    if (
      nomeBase &&
      (
        status === "ATIVO" ||
        status === "ATIVA"
      )
    ) {
      bases.push(nomeBase);
    }
  }

  return bases;
}


/**
 * Retorna as bases autorizadas.
 */
function obterBasesAutorizadas() {
  const usuario = validarUsuarioAtivo_();
  const basesAtivas = obterTodasBasesAtivas_();

  const perfil = String(
    usuario.perfil || ""
  )
    .toUpperCase()
    .trim();

  if (perfil === "ADMIN" || perfil === "ADMINISTRADOR") {
    return basesAtivas;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("PERMISSOES");

  if (!aba) {
    throw new Error(
      "A aba PERMISSOES não foi encontrada."
    );
  }

  const dados = aba.getDataRange().getValues();
  const basesPermitidas = [];

  for (let i = 1; i < dados.length; i++) {
    const email = String(
      dados[i][0] || ""
    )
      .toLowerCase()
      .trim();

    const nomeBase = String(
      dados[i][1] || ""
    ).trim();

    const status = String(
      dados[i][2] || ""
    )
      .toUpperCase()
      .trim();

    if (
      email === usuario.email &&
      (
        status === "ATIVO" ||
        status === "ATIVA"
      ) &&
      nomeBase
    ) {
      basesPermitidas.push(nomeBase);
    }
  }

  return basesAtivas.filter(base =>
    basesPermitidas.includes(base)
  );
}


/**
 * Dados iniciais do sistema.
 */
function obterDadosInicialSistema() {
  const usuario = validarUsuarioAtivo_();
  const bases = obterBasesAutorizadas();

  return {
    usuario: usuario,
    bases: bases
  };
}


/**
 * Garante que a aba BIPS tenha a coluna STATUS_ATIVO.
 * Compatível com estruturas antigas que tinham 8 colunas.
 */
function garantirColunaStatusAtivo_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("BIPS");

  if (!aba) {
    throw new Error("A aba BIPS não foi encontrada.");
  }

  if (aba.getLastColumn() < 9) {
    aba.insertColumnAfter(5);
  }

  const cabecalhos = aba
    .getRange(1, 1, 1, Math.max(aba.getLastColumn(), 9))
    .getValues()[0]
    .map(v => String(v || "").trim().toUpperCase());

  if (cabecalhos[5] !== "STATUS_ATIVO") {
    aba.insertColumnAfter(5);
  }

  aba.getRange(1, 1, 1, 9).setValues([[
    "NOME_BASE",
    "TIPO_INVENTARIO",
    "SEMANA_REFERENCIA",
    "CODIGO_ATIVO",
    "TIPO_ATIVO",
    "STATUS_ATIVO",
    "DATA_HORA",
    "RESPONSAVEL",
    "RESULTADO"
  ]]);
}


/**
 * Identifica o tipo do ativo pelo cadastro esperado ou pelo prefixo da TAG.
 *
 * Prefixos conhecidos:
 * BIP / BIP-...             -> BIP
 * TERM-SIMP / TERM-SMP-... -> TERMINAL SIMPLES
 * TAB / TAB-...             -> TABLET
 * CEL / CEL-...             -> CELULAR
 * NOTE / NTB / NOTE-...    -> NOTEBOOK
 */
function obterTipoAtivoPorCodigo_(codigoAtivo, nomeBase) {
  const codigo = normalizarCodigoAtivo_(codigoAtivo);

  if (!codigo) {
    return null;
  }

  // Primeiro tenta o cadastro de ativos esperados, quando existir.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaEsperados = ss.getSheetByName("ATIVOS_ESPERADOS");

  if (abaEsperados && abaEsperados.getLastRow() >= 2) {
    const dados = abaEsperados.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
      const codigoLinha = normalizarCodigoAtivo_(dados[i][0]);
      const tipoLinha = String(dados[i][1] || "").trim().toUpperCase();
      const baseLinha = String(dados[i][2] || "").trim().toUpperCase();

      if (codigoLinha === codigo && (!baseLinha || baseLinha === String(nomeBase || "").trim().toUpperCase())) {
        const tipoNormalizado = normalizarTipoAtivo_(tipoLinha);
        if (tipoNormalizado) {
          return tipoNormalizado;
        }
      }
    }
  }

  const mapa = [
    { prefixos: ["TERM-SIMP", "TERM-SMP"], tipo: "TERMINAL SIMPLES", descricao: "Terminal simples" },
    { prefixos: ["BIP"], tipo: "BIP", descricao: "BIP" },
    { prefixos: ["TAB"], tipo: "TABLET", descricao: "Tablet" },
    { prefixos: ["CEL"], tipo: "CELULAR", descricao: "Celular" },
    { prefixos: ["NOTE", "NTB"], tipo: "NOTEBOOK", descricao: "Notebook" }
  ];

  for (let i = 0; i < mapa.length; i++) {
    const item = mapa[i];

    for (let j = 0; j < item.prefixos.length; j++) {
      const prefixo = item.prefixos[j];

      if (
        codigo === prefixo ||
        codigo.indexOf(prefixo + "-") === 0 ||
        codigo.indexOf(prefixo + "_") === 0
      ) {
        return {
          tipo: item.tipo,
          descricao: item.descricao
        };
      }
    }
  }

  return null;
}


/**
 * Normaliza nomes de tipos vindos do cadastro.
 */
function normalizarTipoAtivo_(tipo) {
  const valor = String(tipo || "").trim().toUpperCase();

  const mapa = {
    "TABLET": { tipo: "TABLET", descricao: "Tablet" },
    "TAB": { tipo: "TABLET", descricao: "Tablet" },
    "BIP": { tipo: "BIP", descricao: "BIP" },
    "TERMINAL SIMPLES": { tipo: "TERMINAL SIMPLES", descricao: "Terminal simples" },
    "TERM-SIMP": { tipo: "TERMINAL SIMPLES", descricao: "Terminal simples" },
    "TERM-SMP": { tipo: "TERMINAL SIMPLES", descricao: "Terminal simples" },
    "CELULAR": { tipo: "CELULAR", descricao: "Celular" },
    "NOTEBOOK": { tipo: "NOTEBOOK", descricao: "Notebook" },
    "OUTRO": { tipo: "OUTRO", descricao: "Outro" }
  };

  return mapa[valor] || null;
}


/**
 * Gera um código interno único para ativo sem tag.
 */
function gerarCodigoSemTag_(nomeBase, semanaReferencia) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("BIPS");

  if (!aba) {
    throw new Error("A aba BIPS não foi encontrada.");
  }

  const dados = aba.getDataRange().getValues();
  let maiorNumero = 0;

  for (let i = 1; i < dados.length; i++) {
    const baseLinha = String(dados[i][0] || "").trim().toUpperCase();
    const semanaLinha = String(dados[i][2] || "").trim();
    const codigoLinha = normalizarCodigoAtivo_(dados[i][3]);

    if (
      baseLinha === String(nomeBase || "").trim().toUpperCase() &&
      semanaLinha === semanaReferencia
    ) {
      const match = codigoLinha.match(/^SEM TAG-(\d+)$/);

      if (match) {
        maiorNumero = Math.max(maiorNumero, Number(match[1]));
      }
    }
  }

  const proximo = maiorNumero + 1;
  return "SEM TAG-" + String(proximo).padStart(3, "0");
}


/**
 * Normaliza código de ativo.
 */
function normalizarCodigoAtivo_(codigo) {
  return String(codigo || "")
    .trim()
    .toUpperCase();
}


/**
 * Valida acesso à base.
 */
function validarAcessoBase_(nomeBase) {
  const bases = obterBasesAutorizadas();

  if (!bases.includes(nomeBase)) {
    throw new Error(
      "Você não possui acesso à base: " + nomeBase
    );
  }

  return true;
}


/**
 * Verifica se a bipagem é duplicada.
 */
function verificarBipDuplicado_(
  nomeBase,
  semanaReferencia,
  codigoAtivo
) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("BIPS");

  if (!aba) {
    throw new Error(
      "A aba BIPS não foi encontrada."
    );
  }

  const dados = aba.getDataRange().getValues();
  const codigo = normalizarCodigoAtivo_(codigoAtivo);

  for (let i = 1; i < dados.length; i++) {
    const base = String(
      dados[i][0] || ""
    ).trim();

    const semana = String(
      dados[i][2] || ""
    ).trim();

    const codigoLinha = normalizarCodigoAtivo_(
      dados[i][3]
    );

    if (
      base === nomeBase &&
      semana === semanaReferencia &&
      codigoLinha === codigo
    ) {
      return true;
    }
  }

  return false;
}


/**
 * Registra vários ativos de uma vez.
 * O front-end envia pequenos lotes para reduzir drasticamente o número de
 * chamadas ao servidor e de operações de escrita na planilha.
 */
function registrarBipagens(lista) {
  if (!Array.isArray(lista) || lista.length === 0) {
    throw new Error("Nenhum ativo para sincronizar.");
  }

  const primeiro = lista[0] || {};
  const nomeBase = String(primeiro.nomeBase || "").trim();
  const tipoInventario = String(primeiro.tipoInventario || "").trim();
  const semanaReferencia = String(primeiro.semanaReferencia || "").trim();
  const responsavel = String(primeiro.responsavel || "").trim();

  if (!nomeBase || !tipoInventario || !semanaReferencia || !responsavel) {
    throw new Error("Dados do inventário incompletos.");
  }

  validarAcessoBase_(nomeBase);

  // Bloqueia concorrência para dois usuários não gerarem o mesmo SEM TAG.
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const configuracao = obterConfiguracaoInventario(nomeBase);

    if (tipoInventario !== configuracao.tipoInventario) {
      throw new Error(
        "O tipo de inventário informado não corresponde ao tipo definido para esta base."
      );
    }

    if (semanaReferencia !== configuracao.semanaAtual) {
      throw new Error(
        "A semana do inventário deve ser a semana atual: " +
        configuracao.semanaAtual + "."
      );
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("BIPS");

    if (!aba) {
      throw new Error("A aba BIPS não foi encontrada.");
    }

    garantirColunaStatusAtivo_();

    const valoresExistentes = aba.getDataRange().getValues();
    const chavesExistentes = new Set();
    let maiorSemTag = 0;

    for (let i = 1; i < valoresExistentes.length; i++) {
      const baseLinha = String(valoresExistentes[i][0] || "").trim().toUpperCase();
      const semanaLinha = String(valoresExistentes[i][2] || "").trim();
      const codigoLinha = normalizarCodigoAtivo_(valoresExistentes[i][3]);

      if (baseLinha === nomeBase.toUpperCase() && semanaLinha === semanaReferencia) {
        chavesExistentes.add(baseLinha + "|" + semanaLinha + "|" + codigoLinha);

        const match = codigoLinha.match(/^SEM TAG-(\d+)$/);
        if (match) {
          maiorSemTag = Math.max(maiorSemTag, Number(match[1]));
        }
      }
    }

    const chavesLote = new Set();
    const linhasParaGravar = [];
    const idsProcessados = [];
    const idsDuplicados = [];

    for (let i = 0; i < lista.length; i++) {
      const item = lista[i] || {};
      const id = String(item.id || ("item-" + i));
      const semTag = item.semTag === true || String(item.semTag || "").toUpperCase() === "TRUE";
      let codigoAtivo = normalizarCodigoAtivo_(item.codigoAtivo);
      const tipoInformado = String(item.tipoAtivo || "").trim().toUpperCase();
      const statusAtivo = String(item.statusAtivo || "").trim().toUpperCase();

      if (String(item.nomeBase || "").trim() !== nomeBase || String(item.semanaReferencia || "").trim() !== semanaReferencia) {
        throw new Error("Todos os ativos do lote precisam pertencer ao mesmo inventário.");
      }

      if (!statusAtivo || ["FUNCIONANDO", "ESTRAGADO"].indexOf(statusAtivo) === -1) {
        throw new Error("Existe ativo sem condição válida (FUNCIONANDO ou ESTRAGADO).");
      }

      let tipoAtivo = tipoInformado;

      if (semTag) {
        if (["TABLET", "BIP", "TERMINAL SIMPLES", "CELULAR", "NOTEBOOK", "OUTRO"].indexOf(tipoAtivo) === -1) {
          throw new Error("Tipo de ativo inválido para SEM TAG.");
        }

        do {
          maiorSemTag++;
          codigoAtivo = "SEM TAG-" + String(maiorSemTag).padStart(3, "0");
        } while (chavesExistentes.has(nomeBase.toUpperCase() + "|" + semanaReferencia + "|" + codigoAtivo));
      } else {
        if (!codigoAtivo) {
          throw new Error("Bipe ou informe o código do ativo.");
        }

        const tipoAutomatico = obterTipoAtivoPorCodigo_(codigoAtivo, nomeBase);
        if (!tipoAutomatico) {
          throw new Error(
            "Não foi possível identificar o tipo da tag " + codigoAtivo + "."
          );
        }

        if (tipoAtivo && tipoAtivo !== tipoAutomatico.tipo) {
          throw new Error(
            "O tipo do ativo " + codigoAtivo + " não corresponde ao prefixo/cadastro da tag."
          );
        }

        tipoAtivo = tipoAutomatico.tipo;
      }

      const chave = nomeBase.toUpperCase() + "|" + semanaReferencia + "|" + codigoAtivo;

      if (chavesExistentes.has(chave) || chavesLote.has(chave)) {
        idsDuplicados.push(id);
        continue;
      }

      chavesLote.add(chave);
      linhasParaGravar.push([
        nomeBase,
        tipoInventario,
        semanaReferencia,
        codigoAtivo,
        tipoAtivo,
        statusAtivo,
        new Date(),
        responsavel,
        "REGISTRADO"
      ]);
      idsProcessados.push(id);
    }

    if (linhasParaGravar.length) {
      const primeiraLinha = aba.getLastRow() + 1;
      aba.getRange(primeiraLinha, 1, linhasParaGravar.length, 9).setValues(linhasParaGravar);
      SpreadsheetApp.flush();
    }

    const quantidadeDuplicados = idsDuplicados.length;
    let mensagem = linhasParaGravar.length + (linhasParaGravar.length === 1 ? " ativo sincronizado." : " ativos sincronizados.");
    if (quantidadeDuplicados) {
      mensagem += " " + quantidadeDuplicados + (quantidadeDuplicados === 1 ? " já estava registrado." : " já estavam registrados.");
    }

    return {
      sucesso: true,
      mensagem: mensagem,
      idsProcessados: idsProcessados,
      idsDuplicados: idsDuplicados,
      quantidade: linhasParaGravar.length,
      totalRecebido: lista.length
    };
  } finally {
    lock.releaseLock();
  }
}


/**
 * Registra uma bipagem/ativo no inventário.
 *
 * Regras:
 * - Com TAG: o tipo é identificado automaticamente pelo prefixo/cadastro do ativo.
 * - SEM TAG: o código é gerado automaticamente e o tipo é escolhido manualmente.
 * - Todo ativo precisa ter condição FUNCIONANDO ou ESTRAGADO.
 */
function registrarBipagem(dados) {
  if (!dados) {
    throw new Error("Dados do ativo não informados.");
  }

  const nomeBase = String(dados.nomeBase || "").trim();
  const tipoInventario = String(dados.tipoInventario || "").trim();
  const semanaReferencia = String(dados.semanaReferencia || "").trim();
  const semTag = dados.semTag === true || String(dados.semTag || "").toUpperCase() === "TRUE";
  let codigoAtivo = normalizarCodigoAtivo_(dados.codigoAtivo);
  const tipoAtivoInformado = String(dados.tipoAtivo || "").trim().toUpperCase();
  const statusAtivo = String(dados.statusAtivo || "").trim().toUpperCase();
  const responsavel = String(dados.responsavel || "").trim();

  if (!nomeBase || !tipoInventario || !semanaReferencia) {
    throw new Error("Preencha a base, o tipo e a semana do inventário.");
  }

  if (!responsavel) {
    throw new Error("Responsável não informado.");
  }

  if (!statusAtivo || ["FUNCIONANDO", "ESTRAGADO"].indexOf(statusAtivo) === -1) {
    throw new Error("Informe se o ativo está FUNCIONANDO ou ESTRAGADO.");
  }

  validarAcessoBase_(nomeBase);

  const configuracao = obterConfiguracaoInventario(nomeBase);

  if (tipoInventario !== configuracao.tipoInventario) {
    throw new Error(
      "O tipo de inventário informado não corresponde ao tipo definido para esta base."
    );
  }

  if (semanaReferencia !== configuracao.semanaAtual) {
    throw new Error(
      "A semana do inventário deve ser a semana atual: " + configuracao.semanaAtual + "."
    );
  }

  garantirColunaStatusAtivo_();

  let tipoAtivo = tipoAtivoInformado;

  if (semTag) {
    if (!tipoAtivo) {
      throw new Error("Para SEM TAG, selecione o tipo do ativo.");
    }

    if (["TABLET", "BIP", "TERMINAL SIMPLES", "CELULAR", "NOTEBOOK", "OUTRO"].indexOf(tipoAtivo) === -1) {
      throw new Error("Tipo de ativo inválido.");
    }

    codigoAtivo = gerarCodigoSemTag_(nomeBase, semanaReferencia);
  } else {
    if (!codigoAtivo) {
      throw new Error("Bipe ou informe o código do ativo.");
    }

    const tipoAutomatico = obterTipoAtivoPorCodigo_(codigoAtivo, nomeBase);

    if (!tipoAutomatico) {
      throw new Error(
        "Não foi possível identificar o tipo desta tag. Confira o prefixo do código do ativo."
      );
    }

    if (tipoAtivo && tipoAtivo !== tipoAutomatico.tipo) {
      throw new Error(
        "O tipo do ativo não corresponde ao prefixo/cadastro da tag. Tipo esperado: " +
        tipoAutomatico.descricao + "."
      );
    }

    tipoAtivo = tipoAutomatico.tipo;
  }

  if (verificarBipDuplicado_(nomeBase, semanaReferencia, codigoAtivo)) {
    return {
      sucesso: false,
      mensagem: "Este ativo já foi registrado neste inventário."
    };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("BIPS");

  if (!aba) {
    throw new Error("A aba BIPS não foi encontrada.");
  }

  aba.appendRow([
    nomeBase,
    tipoInventario,
    semanaReferencia,
    codigoAtivo,
    tipoAtivo,
    statusAtivo,
    new Date(),
    responsavel,
    "REGISTRADO"
  ]);

  return {
    sucesso: true,
    mensagem: semTag
      ? "Ativo SEM TAG registrado com sucesso."
      : "Ativo registrado com sucesso.",
    codigoAtivo: codigoAtivo,
    tipoAtivo: tipoAtivo,
    statusAtivo: statusAtivo,
    semTag: semTag
  };
}


/**
 * Atualiza somente a condição (FUNCIONANDO/ESTRAGADO) de um ativo já registrado.
 */
function atualizarStatusAtivo(dados) {
  if (!dados) {
    throw new Error("Dados da atualização não informados.");
  }

  const nomeBase = String(dados.nomeBase || "").trim();
  const semanaReferencia = String(dados.semanaReferencia || "").trim();
  const codigoAtivo = normalizarCodigoAtivo_(dados.codigoAtivo);
  const statusAtivo = String(dados.statusAtivo || "").trim().toUpperCase();
  const responsavel = String(dados.responsavel || "").trim();

  if (!nomeBase || !semanaReferencia || !codigoAtivo) {
    throw new Error("Base, semana e código do ativo são obrigatórios.");
  }

  if (["FUNCIONANDO", "ESTRAGADO"].indexOf(statusAtivo) === -1) {
    throw new Error("A condição deve ser FUNCIONANDO ou ESTRAGADO.");
  }

  if (!responsavel) {
    throw new Error("Responsável não informado.");
  }

  validarAcessoBase_(nomeBase);

  const configuracao = obterConfiguracaoInventario(nomeBase);
  if (semanaReferencia !== configuracao.semanaAtual) {
    throw new Error("A semana do inventário deve ser a semana atual: " + configuracao.semanaAtual + ".");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("BIPS");
  if (!aba) {
    throw new Error("A aba BIPS não foi encontrada.");
  }

  garantirColunaStatusAtivo_();

  const dadosBips = aba.getDataRange().getValues();
  const baseNormalizada = nomeBase.toUpperCase();
  let linhaEncontrada = -1;

  for (let i = 1; i < dadosBips.length; i++) {
    const baseLinha = String(dadosBips[i][0] || "").trim().toUpperCase();
    const semanaLinha = String(dadosBips[i][2] || "").trim();
    const codigoLinha = normalizarCodigoAtivo_(dadosBips[i][3]);

    if (baseLinha === baseNormalizada && semanaLinha === semanaReferencia && codigoLinha === codigoAtivo) {
      linhaEncontrada = i + 1;
      break;
    }
  }

  if (linhaEncontrada === -1) {
    throw new Error("Ativo não encontrado no inventário para edição.");
  }

  aba.getRange(linhaEncontrada, 6).setValue(statusAtivo);

  return {
    sucesso: true,
    mensagem: "Condição atualizada com sucesso.",
    codigoAtivo: codigoAtivo,
    statusAtivo: statusAtivo
  };
}


/**
 * Registra o início do inventário.
 */
function registrarInventario(dados) {
  if (!dados) {
    throw new Error("Dados do inventário não informados.");
  }

  const nomeBase = String(dados.nomeBase || "").trim();
  const tipoInventario = String(dados.tipoInventario || "").trim();
  const semanaReferencia = String(dados.semanaReferencia || "").trim();
  const responsavel = String(dados.responsavel || "").trim();

  if (!nomeBase || !tipoInventario || !semanaReferencia || !responsavel) {
    throw new Error("Preencha todos os campos obrigatórios.");
  }

  validarAcessoBase_(nomeBase);

  const configuracao = obterConfiguracaoInventario(nomeBase);

  if (semanaReferencia !== configuracao.semanaAtual) {
    throw new Error(
      "A semana de referência deve ser a semana atual: " +
      configuracao.semanaAtual
    );
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("INVENTARIOS");

  if (!aba) {
    throw new Error("A aba INVENTARIOS não foi encontrada.");
  }

  // Regra de negócio: cada base pode ter somente 1 inventário por semana.
  // Se estiver EM ANDAMENTO, permitimos retomada do mesmo inventário.
  const dadosInventarios = aba.getDataRange().getValues();
  const baseNormalizada = nomeBase.toUpperCase();

  for (let i = 1; i < dadosInventarios.length; i++) {
    const baseLinha = String(dadosInventarios[i][0] || "")
      .trim()
      .toUpperCase();
    const semanaLinha = String(dadosInventarios[i][2] || "").trim();
    const statusLinha = String(dadosInventarios[i][6] || "")
      .trim()
      .toUpperCase();

    if (
      baseLinha === baseNormalizada &&
      semanaLinha === semanaReferencia
    ) {
      if (statusLinha === "FINALIZADO") {
        throw new Error(
          "Esta base já realizou o inventário da " +
          semanaReferencia + ". Um novo inventário será liberado na próxima semana."
        );
      }

      if (statusLinha === "EM ANDAMENTO") {
        const tipoExistente = String(dadosInventarios[i][1] || "").trim().toUpperCase();

        return {
          sucesso: true,
          continuar: true,
          mensagem: "Inventário em andamento localizado. Continuando o inventário da " + semanaReferencia + ".",
          tipoInventario: tipoExistente || configuracao.tipoInventario,
          semanaReferencia: semanaReferencia
        };
      }

      throw new Error(
        "Esta base já possui um inventário registrado na " +
        semanaReferencia + ". Um novo inventário será liberado na próxima semana."
      );
    }
  }

  if (tipoInventario !== configuracao.tipoInventario) {
    throw new Error(
      "O tipo de inventário é definido automaticamente para esta base: " +
      configuracao.descricao + "."
    );
  }

  aba.appendRow([
    nomeBase,
    configuracao.tipoInventario,
    configuracao.semanaAtual,
    new Date(),
    "",
    responsavel,
    "EM ANDAMENTO",
    0,
    ""
  ]);

  return {
    sucesso: true,
    mensagem:
      "Inventário " + configuracao.descricao + " iniciado com sucesso.",
    tipoInventario: configuracao.tipoInventario,
    semanaReferencia: configuracao.semanaAtual
  };
}


/**
 * Retorna os ativos já registrados em um inventário em andamento.
 * Usado para permitir que o usuário retome o inventário após sair/recarregar a tela.
 */
function obterAtivosInventario(nomeBase, semanaReferencia) {
  if (!nomeBase || !semanaReferencia) {
    throw new Error("Informe a base e a semana do inventário.");
  }

  validarAcessoBase_(nomeBase);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("BIPS");

  if (!aba) {
    throw new Error("A aba BIPS não foi encontrada.");
  }

  const dados = aba.getDataRange().getValues();
  const ativos = [];

  for (let i = 1; i < dados.length; i++) {
    const base = String(dados[i][0] || "").trim();
    const semana = String(dados[i][2] || "").trim();

    if (
      base.toUpperCase() === String(nomeBase).trim().toUpperCase() &&
      semana === String(semanaReferencia).trim()
    ) {
      const codigo = normalizarCodigoAtivo_(dados[i][3]);
      ativos.push({
        id: "server-" + (i + 1),
        nomeBase: base,
        semanaReferencia: semana,
        codigoAtivo: codigo.startsWith("SEM TAG-") ? "" : codigo,
        codigoExibicao: codigo,
        tipoAtivo: String(dados[i][4] || "").trim().toUpperCase(),
        statusAtivo: String(dados[i][5] || "").trim().toUpperCase(),
        semTag: codigo.indexOf("SEM TAG-") === 0,
        tipoInventario: String(dados[i][1] || "").trim().toUpperCase(),
        responsavel: String(dados[i][7] || "").trim()
      });
    }
  }

  return {
    sucesso: true,
    total: ativos.length,
    ativos: ativos
  };
}


/**
 * Finaliza o inventário.
 */
function finalizarInventario(dados) {
  if (!dados) {
    throw new Error(
      "Dados da finalização não informados."
    );
  }

  const nomeBase = String(
    dados.nomeBase || ""
  ).trim();

  const semanaReferencia = String(
    dados.semanaReferencia || ""
  ).trim();

  if (!nomeBase || !semanaReferencia) {
    throw new Error(
      "Informe a base e a semana."
    );
  }

  validarAcessoBase_(nomeBase);

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const abaInventarios =
    ss.getSheetByName("INVENTARIOS");

  const abaBips =
    ss.getSheetByName("BIPS");

  if (!abaInventarios || !abaBips) {
    throw new Error(
      "As abas INVENTARIOS e BIPS precisam existir."
    );
  }

  const inventarios =
    abaInventarios.getDataRange().getValues();

  const bips =
    abaBips.getDataRange().getValues();

  let linhaEncontrada = -1;

  for (let i = 1; i < inventarios.length; i++) {
    const base = String(
      inventarios[i][0] || ""
    ).trim();

    const semana = String(
      inventarios[i][2] || ""
    ).trim();

    const status = String(
      inventarios[i][6] || ""
    )
      .toUpperCase()
      .trim();

    if (
      base === nomeBase &&
      semana === semanaReferencia &&
      status === "EM ANDAMENTO"
    ) {
      linhaEncontrada = i + 1;
      break;
    }
  }

  if (linhaEncontrada === -1) {
    throw new Error(
      "Inventário em andamento não encontrado."
    );
  }

  let totalAtivos = 0;

  for (let i = 1; i < bips.length; i++) {
    const base = String(
      bips[i][0] || ""
    ).trim();

    const semana = String(
      bips[i][2] || ""
    ).trim();

    if (
      base === nomeBase &&
      semana === semanaReferencia
    ) {
      totalAtivos++;
    }
  }

  abaInventarios
    .getRange(linhaEncontrada, 5)
    .setValue(new Date());

  abaInventarios
    .getRange(linhaEncontrada, 7)
    .setValue("FINALIZADO");

  abaInventarios
    .getRange(linhaEncontrada, 8)
    .setValue(totalAtivos);

  return {
    sucesso: true,
    mensagem: "Inventário finalizado com sucesso.",
    totalAtivos: totalAtivos,
    totalBips: totalAtivos
  };
}


/**
 * Retorna o último inventário realizado por base.
 */
function obterUltimoInventarioPorBase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("INVENTARIOS");

  if (!aba) {
    throw new Error(
      "A aba INVENTARIOS não foi encontrada."
    );
  }

  const dados = aba.getDataRange().getValues();
  const resultado = {};

  for (let i = 1; i < dados.length; i++) {
    const base = String(
      dados[i][0] || ""
    ).trim();

    const semana = String(
      dados[i][2] || ""
    ).trim();

    const dataFim = dados[i][4];

    const status = String(
      dados[i][6] || ""
    )
      .toUpperCase()
      .trim();

    if (
      base &&
      status === "FINALIZADO"
    ) {
      resultado[base] = {
        semana: semana,
        dataFim: dataFim,
        status: "REALIZADO"
      };
    }
  }

  return resultado;
}


/**
 * Abre o sistema pelo menu da planilha.
 */
function abrirSistema() {
  const url = ScriptApp.getService().getUrl();

  if (!url) {
    SpreadsheetApp.getUi().alert(
      "O Web App ainda não foi implantado.\n\n" +
      "Vá em Implantar → Nova implantação."
    );

    return;
  }

  const html = HtmlService.createHtmlOutput(
    "<script>" +
    "window.open('" + url + "', '_blank');" +
    "google.script.host.close();" +
    "</script>"
  );

  SpreadsheetApp
    .getUi()
    .showModalDialog(
      html,
      "Abrindo sistema"
    );
}


/**
 * Corrige estruturas antigas.
 */
function corrigirEstruturaAntiga() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const abaBases = ss.getSheetByName("BASES");

  if (abaBases && abaBases.getLastRow() >= 1) {
    const cabecalhos = abaBases
      .getRange(
        1,
        1,
        1,
        abaBases.getLastColumn()
      )
      .getValues()[0]
      .map(valor =>
        String(valor || "")
          .trim()
          .toUpperCase()
      );

    if (cabecalhos[0] === "ID_BASE") {
      const ultimaLinha = abaBases.getLastRow();

      if (ultimaLinha >= 2) {
        const dados = abaBases
          .getRange(2, 1, ultimaLinha - 1, 5)
          .getValues();

        const corrigidos = dados.map(linha => [
          linha[1],
          linha[2],
          linha[3],
          linha[4]
        ]);

        abaBases
          .getRange(2, 1, corrigidos.length, 4)
          .setValues(corrigidos);
      }

      if (abaBases.getMaxColumns() >= 5) {
        abaBases.deleteColumn(5);
      }
    }
  }

  const abaUsuarios = ss.getSheetByName("USUARIOS");

  if (abaUsuarios && abaUsuarios.getLastRow() >= 1) {
    const cabecalhos = abaUsuarios
      .getRange(
        1,
        1,
        1,
        abaUsuarios.getLastColumn()
      )
      .getValues()[0]
      .map(valor =>
        String(valor || "")
          .trim()
          .toUpperCase()
      );

    if (cabecalhos[0] === "ID_USUARIO") {
      const ultimaLinha = abaUsuarios.getLastRow();

      if (ultimaLinha >= 2) {
        const dados = abaUsuarios
          .getRange(2, 1, ultimaLinha - 1, 6)
          .getValues();

        const corrigidos = dados.map(linha => [
          linha[1],
          linha[2],
          linha[3],
          linha[4],
          linha[5]
        ]);

        abaUsuarios
          .getRange(2, 1, corrigidos.length, 5)
          .setValues(corrigidos);
      }

      if (abaUsuarios.getMaxColumns() >= 6) {
        abaUsuarios.deleteColumn(6);
      }
    }
  }

  configurarSistema();

  SpreadsheetApp.getUi().alert(
    "Estrutura antiga corrigida com sucesso!"
  );
}

/**
 * Retorna o endereço oficial do Web App.
 */
function obterUrlSistema() {
  const url = ScriptApp.getService().getUrl();

  if (!url) {
    throw new Error(
      "Não foi possível identificar a URL da implantação do sistema."
    );
  }

  return url + "?pagina=inventario";
}

/**
 * Retorna a URL de uma página interna do sistema.
 * Páginas suportadas: inventario, dashboard.
 */
function obterUrlPagina(pagina) {
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    throw new Error(
      "Não foi possível identificar a URL da implantação do sistema."
    );
  }

  const paginaNormalizada = String(pagina || "")
    .trim()
    .toLowerCase();

  if (paginaNormalizada === "dashboard") {
    return url + "?pagina=dashboard";
  }

  return url + "?pagina=inventario";
}
/**
 * Define automaticamente o tipo de inventário da base.
 * Se a base nunca teve inventário: INICIAL.
 * Caso já tenha algum registro: SEMANAL.
 */
function obterSemanaAtual_() {
  const data = new Date();

  const dataUTC = new Date(Date.UTC(
    data.getFullYear(),
    data.getMonth(),
    data.getDate()
  ));

  const diaSemana = dataUTC.getUTCDay() || 7;

  dataUTC.setUTCDate(
    dataUTC.getUTCDate() + 4 - diaSemana
  );

  const inicioAno = new Date(Date.UTC(
    dataUTC.getUTCFullYear(),
    0,
    1
  ));

  const numeroSemana = Math.ceil(
    (((dataUTC - inicioAno) / 86400000) + 1) / 7
  );

  return "Semana " + numeroSemana;
}


/**
 * Define automaticamente o tipo de inventário da base
 * e informa a semana atual.
 *
 * Primeira vez da base: INICIAL
 * Após o primeiro inventário: SEMANAL
 */
function obterConfiguracaoInventario(nomeBase) {
  if (!nomeBase) {
    throw new Error("Base não informada.");
  }

  validarAcessoBase_(nomeBase);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("INVENTARIOS");

  if (!aba) {
    throw new Error("A aba INVENTARIOS não foi encontrada.");
  }

  const dados = aba.getDataRange().getValues();
  const baseSelecionada = String(nomeBase).trim().toUpperCase();
  const semanaAtual = obterSemanaAtual_();

  let possuiInventarioHistorico = false;
  let inventarioEmAndamentoAtual = null;

  for (let i = 1; i < dados.length; i++) {
    const baseRegistrada = String(dados[i][0] || "")
      .trim()
      .toUpperCase();
    const semanaRegistrada = String(dados[i][2] || "").trim();
    const tipoRegistrado = String(dados[i][1] || "").trim().toUpperCase();
    const statusRegistrado = String(dados[i][6] || "")
      .trim()
      .toUpperCase();

    if (baseRegistrada !== baseSelecionada) {
      continue;
    }

    possuiInventarioHistorico = true;

    // Se o inventário da semana atual ainda estiver aberto,
    // preserva o mesmo tipo (inclusive INICIAL) para permitir continuidade.
    if (
      semanaRegistrada === semanaAtual &&
      statusRegistrado === "EM ANDAMENTO"
    ) {
      inventarioEmAndamentoAtual = tipoRegistrado;
    }
  }

  const tipoInventario = inventarioEmAndamentoAtual ||
    (possuiInventarioHistorico ? "SEMANAL" : "INICIAL");

  return {
    sucesso: true,
    tipoInventario: tipoInventario,
    descricao: tipoInventario === "INICIAL"
      ? "Inventário Inicial"
      : "Inventário Semanal",
    semanaAtual: semanaAtual,
    continuar: Boolean(inventarioEmAndamentoAtual)
  };
}
