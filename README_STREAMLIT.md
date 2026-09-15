# Migração para Streamlit

A nova interface está em `app.py` e mantém o Google Sheets como fonte de dados. A migração começou sem apagar o Apps Script antigo, permitindo validação paralela.

## O que já foi migrado

- Login por Google/OIDC preparado com `st.login()` + `st.user`.
- Validação do usuário na aba `USUARIOS`.
- Permissões por base via `PERMISSOES`.
- Dashboard e Inventário como módulos dentro do mesmo app.
- Regras `INICIAL` / `SEMANAL`.
- Uma execução por base/semana.
- Retomada de inventário `EM ANDAMENTO`.
- Registro de ativos com identificação por prefixo.
- `SEM TAG` com numeração sequencial.
- Condição `FUNCIONANDO` / `ESTRAGADO`.
- Edição da condição do ativo.
- Finalização do inventário e total de ativos.

## Configuração do Google Sheets

O app usa uma conta de serviço para acessar a planilha. Compartilhe a planilha com o `client_email` da conta de serviço e configure os Secrets do Streamlit com base em `.streamlit/secrets.example.toml`.

Nunca commite um `secrets.toml` real.

## Teste inicial

Para o primeiro teste, é possível usar temporariamente `DEV_EMAIL` nos Secrets. Em produção, a estrutura para Google OIDC já está preparada e deve substituir o modo de desenvolvimento.

## Deploy no Streamlit Community Cloud

Escolha o repositório `brunaizidro/inventario-ativos-loggi`, branch `main` e arquivo `app.py`. Depois, em **Advanced settings > Secrets**, cole os segredos necessários.

O Community Cloud acompanha o GitHub e pode atualizar o app após novos commits.
