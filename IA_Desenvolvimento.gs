/************************************************************
 * AGENTE IA DE DESENVOLVIMENTO - APPS SCRIPT
 *
 * Permite:
 * 1. Ler os arquivos do próprio projeto Apps Script.
 * 2. Enviar o projeto + pedido para a OpenAI Responses API.
 * 3. Receber apenas os arquivos alterados em JSON estruturado.
 * 4. Criar backup antes de aplicar a alteração.
 * 5. Atualizar o projeto via Apps Script API.
 *
 * IMPORTANTE:
 * - Nunca coloque a chave da OpenAI em HTML/JavaScript do navegador.
 * - A chave fica em Script Properties.
 * - A aplicação da alteração exige confirmar=true.
 ************************************************************/

const IA_DEV_CONFIG = {
  OPENAI_API_URL: "https://api.openai.com/v1/responses",
  OPENAI_MODEL: "gpt-5",
  APPS_SCRIPT_API_BASE: "https://script.googleapis.com/v1/projects",
  BACKUP_FOLDER: "Backups - Inventário de Ativos"
};


/**
 * Adiciona o menu do agente de desenvolvimento.
 * Esta função é chamada pelo onOpen principal.
 */
function menuAgenteIA_(ui) {
  ui.createMenu("IA Desenvolvimento")
    .addItem("Configurar chave OpenAI", "configurarChaveOpenAI")
    .addItem("Testar conexão OpenAI", "testarConexaoOpenAI")
    .addItem("Ler projeto atual", "mostrarResumoProjetoIA")
    .addItem("Criar backup do projeto", "criarBackupProjetoIA")
    .addToUi();
}


/**
 * Salva a chave da OpenAI em Script Properties.
 */
function configurarChaveOpenAI() {
  const ui = SpreadsheetApp.getUi();
  const resultado = ui.prompt(
    "Configurar OpenAI",
    "Cole sua chave da API da OpenAI. Ela será armazenada somente nas propriedades do script.",
    ui.ButtonSet.OK_CANCEL
  );

  if (resultado.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const chave = String(resultado.getResponseText() || "").trim();
  if (!chave) {
    ui.alert("Nenhuma chave foi informada.");
    return;
  }

  PropertiesService.getScriptProperties()
    .setProperty("OPENAI_API_KEY", chave);

  ui.alert("Chave da OpenAI configurada com sucesso.");
}


/**
 * Testa a autenticação com a OpenAI.
 */
function testarConexaoOpenAI() {
  const chave = obterChaveOpenAI_();

  const resposta = UrlFetchApp.fetch(
    "https://api.openai.com/v1/models",
    {
      method: "get",
      headers: {
        Authorization: "Bearer " + chave
      },
      muteHttpExceptions: true
    }
  );

  const codigo = resposta.getResponseCode();

  if (codigo < 200 || codigo >= 300) {
    throw new Error(
      "Falha na conexão com a OpenAI (HTTP " + codigo + ").\n\n" +
      resposta.getContentText()
    );
  }

  SpreadsheetApp.getUi().alert("Conexão com a OpenAI realizada com sucesso.");
}


/**
 * Retorna a chave da OpenAI armazenada em Script Properties.
 */
function obterChaveOpenAI_() {
  const chave = PropertiesService.getScriptProperties()
    .getProperty("OPENAI_API_KEY");

  if (!chave) {
    throw new Error(
      "A chave da OpenAI ainda não foi configurada. " +
      "Use o menu IA Desenvolvimento → Configurar chave OpenAI."
    );
  }

  return chave;
}


/**
 * Lê o conteúdo atual do próprio projeto pela Apps Script API.
 */
function obterConteudoProjetoIA_() {
  const scriptId = ScriptApp.getScriptId();
  const token = ScriptApp.getOAuthToken();
  const url = IA_DEV_CONFIG.APPS_SCRIPT_API_BASE +
    "/" + encodeURIComponent(scriptId) + "/content";

  const resposta = UrlFetchApp.fetch(
    url,
    {
      method: "get",
      headers: {
        Authorization: "Bearer " + token
      },
      muteHttpExceptions: true
    }
  );

  const codigo = resposta.getResponseCode();
  const texto = resposta.getContentText();

  if (codigo < 200 || codigo >= 300) {
    throw new Error(
      "Não foi possível ler o projeto pelo Apps Script API (HTTP " +
      codigo + ").\n\n" + texto
    );
  }

  const dados = JSON.parse(texto);

  if (!Array.isArray(dados.files)) {
    throw new Error("A API do Apps Script não retornou a lista de arquivos.");
  }

  return {
    scriptId: scriptId,
    files: dados.files.map(function(file) {
      return {
        name: file.name,
        type: file.type,
        source: file.source || ""
      };
    })
  };
}


/**
 * Mostra um resumo dos arquivos atuais do projeto.
 */
function mostrarResumoProjetoIA() {
  const projeto = obterConteudoProjetoIA_();

  const resumo = projeto.files.map(function(file) {
    return file.name + " | " + file.type + " | " +
      String(file.source || "").length + " caracteres";
  }).join("\n");

  SpreadsheetApp.getUi().alert(
    "Projeto: " + projeto.scriptId + "\n\n" + resumo
  );
}


/**
 * Cria backup completo do projeto antes de qualquer alteração.
 */
function criarBackupProjetoIA() {
  const projeto = obterConteudoProjetoIA_();
  const arquivo = criarArquivoBackupProjetoIA_(projeto);

  SpreadsheetApp.getUi().alert(
    "Backup criado com sucesso.\n\nArquivo: " + arquivo.getName()
  );

  return arquivo.getId();
}


/**
 * Cria o arquivo de backup em uma pasta dedicada do Drive.
 */
function criarArquivoBackupProjetoIA_(projeto) {
  const data = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd_HH-mm-ss"
  );

  const nome = "backup_apps_script_" + data + ".json";
  const conteudo = JSON.stringify(projeto, null, 2);

  const pastas = DriveApp.getFoldersByName(IA_DEV_CONFIG.BACKUP_FOLDER);
  const pasta = pastas.hasNext()
    ? pastas.next()
    : DriveApp.createFolder(IA_DEV_CONFIG.BACKUP_FOLDER);

  return pasta.createFile(
    nome,
    conteudo,
    MimeType.PLAIN_TEXT
  );
}


/**
 * Envia projeto + pedido para a OpenAI.
 * Retorna somente os arquivos que a IA decidiu alterar.
 */
function gerarAlteracaoComIA_(pedido, projeto) {
  const chave = obterChaveOpenAI_();

  const arquivosTexto = projeto.files.map(function(file) {
    return [
      "===== ARQUIVO: " + file.name + " =====",
      "TIPO: " + file.type,
      file.source,
      "===== FIM ARQUIVO ====="
    ].join("\n");
  }).join("\n\n");

  const instrucoes = [
    "Você é um engenheiro de software responsável por manter um projeto Google Apps Script.",
    "Analise o projeto inteiro antes de modificar qualquer coisa.",
    "O pedido do usuário é:",
    pedido,
    "\nRegras obrigatórias:",
    "1. Preserve todas as funcionalidades existentes que não estejam relacionadas ao pedido.",
    "2. Não invente nomes de arquivos nem apague arquivos existentes.",
    "3. Altere somente os arquivos necessários.",
    "4. Para cada arquivo alterado, devolva o conteúdo COMPLETO do arquivo, nunca um trecho.",
    "5. Não retorne markdown nem blocos de código.",
    "6. Não altere o manifest sem necessidade.",
    "7. Caso o pedido não possa ser aplicado com segurança, devolva a lista de alterações vazia e explique em resumo.",
    "8. Mantenha o código compatível com Google Apps Script V8.",
    "\nArquivos atuais do projeto:\n",
    arquivosTexto
  ].join("\n");

  const payload = {
    model: IA_DEV_CONFIG.OPENAI_MODEL,
    input: [
      {
        role: "developer",
        content: instrucoes
      }
    ],
    store: false,
    text: {
      format: {
        type: "json_schema",
        name: "code_changes",
        strict: true,
        schema: {
          type: "object",
          additionalProperties: false,
          properties: {
            resumo: {
              type: "string"
            },
            alteracoes: {
              type: "array",
              items: {
                type: "object",
                additionalProperties: false,
                properties: {
                  name: {
                    type: "string"
                  },
                  type: {
                    type: "string",
                    enum: ["SERVER_JS", "HTML", "JSON"]
                  },
                  source: {
                    type: "string"
                  }
                },
                required: ["name", "type", "source"]
              }
            }
          },
          required: ["resumo", "alteracoes"]
        }
      }
    }
  };

  const resposta = UrlFetchApp.fetch(
    IA_DEV_CONFIG.OPENAI_API_URL,
    {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: "Bearer " + chave
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );

  const codigo = resposta.getResponseCode();
  const texto = resposta.getContentText();

  if (codigo < 200 || codigo >= 300) {
    throw new Error(
      "OpenAI retornou HTTP " + codigo + ":\n\n" + texto
    );
  }

  const dados = JSON.parse(texto);
  const outputText = String(dados.output_text || "").trim();

  if (!outputText) {
    throw new Error("A OpenAI não retornou um conteúdo estruturado.");
  }

  let resultado;
  try {
    resultado = JSON.parse(outputText);
  } catch (erro) {
    throw new Error(
      "A resposta da OpenAI não veio em JSON válido.\n\n" + outputText
    );
  }

  if (!resultado || !Array.isArray(resultado.alteracoes)) {
    throw new Error("A resposta da OpenAI não contém a estrutura esperada.");
  }

  return resultado;
}


/**
 * Aplica uma alteração solicitada à IA.
 *
 * Exemplo:
 * aplicarPedidoComIA(
 *   "Corrija o fluxo de retomada do inventário e permita continuar um inventário em andamento.",
 *   true
 * );
 *
 * confirmar=false retorna a proposta sem alterar o projeto.
 */
function aplicarPedidoComIA(pedido, confirmar) {
  if (!pedido || !String(pedido).trim()) {
    throw new Error("Informe o pedido de alteração.");
  }

  const projetoAtual = obterConteudoProjetoIA_();
  const proposta = gerarAlteracaoComIA_(String(pedido).trim(), projetoAtual);

  if (!proposta.alteracoes.length) {
    return {
      sucesso: true,
      aplicado: false,
      resumo: proposta.resumo,
      alteracoes: []
    };
  }

  validarAlteracoesIA_(projetoAtual.files, proposta.alteracoes);

  if (confirmar !== true) {
    return {
      sucesso: true,
      aplicado: false,
      resumo: proposta.resumo,
      alteracoes: proposta.alteracoes.map(function(item) {
        return item.name;
      })
    };
  }

  // Backup obrigatório antes de alterar qualquer arquivo.
  const backup = criarArquivoBackupProjetoIA_(projetoAtual);

  const mapaAlteracoes = {};
  proposta.alteracoes.forEach(function(item) {
    mapaAlteracoes[item.name] = item;
  });

  const novoProjeto = {
    scriptId: projetoAtual.scriptId,
    files: projetoAtual.files.map(function(file) {
      if (Object.prototype.hasOwnProperty.call(mapaAlteracoes, file.name)) {
        const novo = mapaAlteracoes[file.name];
        return {
          name: novo.name,
          type: novo.type,
          source: novo.source
        };
      }

      return {
        name: file.name,
        type: file.type,
        source: file.source
      };
    })
  };

  atualizarConteudoProjetoIA_(novoProjeto);

  return {
    sucesso: true,
    aplicado: true,
    resumo: proposta.resumo,
    arquivosAlterados: proposta.alteracoes.map(function(item) {
      return item.name;
    }),
    backupId: backup.getId(),
    backupNome: backup.getName()
  };
}


/**
 * Valida a lista de arquivos que a IA quer alterar.
 */
function validarAlteracoesIA_(arquivosAtuais, alteracoes) {
  const atuais = {};

  arquivosAtuais.forEach(function(file) {
    atuais[file.name] = file.type;
  });

  const vistos = {};

  alteracoes.forEach(function(item) {
    if (!item.name || !item.type || typeof item.source !== "string") {
      throw new Error("A IA retornou uma alteração de arquivo inválida.");
    }

    if (!Object.prototype.hasOwnProperty.call(atuais, item.name)) {
      throw new Error(
        "A IA tentou alterar/criar um arquivo inexistente: " + item.name
      );
    }

    if (atuais[item.name] !== item.type) {
      throw new Error(
        "O tipo do arquivo " + item.name + " não corresponde ao projeto atual."
      );
    }

    if (vistos[item.name]) {
      throw new Error("Arquivo duplicado na resposta da IA: " + item.name);
    }

    vistos[item.name] = true;

    if (item.type === "SERVER_JS" && !item.source.trim()) {
      throw new Error("O arquivo SERVER_JS não pode ficar vazio: " + item.name);
    }

    if (item.type === "HTML" && !item.source.trim()) {
      throw new Error("O arquivo HTML não pode ficar vazio: " + item.name);
    }

    if (item.type === "JSON") {
      try {
        JSON.parse(item.source);
      } catch (erro) {
        throw new Error("O manifest retornado pela IA não é um JSON válido.");
      }
    }
  });
}


/**
 * Atualiza o projeto inteiro preservando os arquivos não alterados.
 */
function atualizarConteudoProjetoIA_(projeto) {
  const token = ScriptApp.getOAuthToken();
  const url = IA_DEV_CONFIG.APPS_SCRIPT_API_BASE +
    "/" + encodeURIComponent(projeto.scriptId) + "/content";

  const payload = {
    files: projeto.files.map(function(file) {
      return {
        name: file.name,
        type: file.type,
        source: file.source
      };
    })
  };

  const resposta = UrlFetchApp.fetch(
    url,
    {
      method: "put",
      contentType: "application/json",
      headers: {
        Authorization: "Bearer " + token
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );

  const codigo = resposta.getResponseCode();
  const texto = resposta.getContentText();

  if (codigo < 200 || codigo >= 300) {
    throw new Error(
      "Não foi possível atualizar o projeto pelo Apps Script API (HTTP " +
      codigo + ").\n\n" + texto
    );
  }

  return JSON.parse(texto);
}
