# Consolidador Web de Planilhas

Aplicacao Flask para recebimento publico de planilhas e geracao de consolidacoes separadas de Saude e Odonto, com rastreabilidade, resumo e inconsistencias.

## O que a aplicacao entrega

- pagina publica com dois campos de upload: `Saude` e `Odonto`
- area administrativa protegida por senha
- consolidacao separada em Excel para `Saude` e `Odonto`
- rastreabilidade por lote, remetente, horario e arquivo de origem
- historico simples de consolidacoes geradas
- filtro por protocolo, tipo e remetente no painel admin
- filtro por periodo no painel admin
- paginas separadas para `Arquivos recebidos` e `Auditoria simples`
- escopo por lote para gerar consolidado sem misturar testes antigos
- exclusao manual de arquivos individuais e lotes completos
- exclusao em massa de arquivos selecionados no painel admin
- historico de consolidacoes com identificacao de escopo por lote
- download individual de arquivos recebidos
- auditoria basica de uploads, logins e consolidacoes
- mascaramento visual de e-mail e IP no admin
- expiracao da sessao admin por inatividade
- limite de frequencia para uploads por IP
- suporte a senha admin com hash
- criptografia opcional dos arquivos enviados e consolidados em disco

## Arquitetura atual

O projeto foi reorganizado para um formato mais limpo e escalavel:

```text
app.py
consolidador/
  __init__.py          # app factory
  config.py            # configuracao por ambiente
  db.py                # bootstrap e acesso ao SQLite
  models.py            # dataclasses do dominio
  security.py          # CSRF, headers e protecao admin
  blueprints/
    public.py          # upload publico
    admin.py           # autenticacao e painel
  services/
    uploads.py         # validacao e armazenamento
    consolidation.py   # leitura, normalizacao e geracao do xlsx
    repository.py      # consultas e gravacao no banco
    audit.py           # trilha de auditoria
  templates/
  static/
```

## Regra de consolidacao

- Saude usa sua planilha-base e Odonto usa sua propria planilha-base
- por padrao, o projeto procura automaticamente as planilhas-base na raiz do repositorio
- colunas sao normalizadas com remocao de espacos extras
- colunas tecnicamente vazias e sem cabecalho sao descartadas
- linhas totalmente vazias podem ser ignoradas
- duplicados podem ser removidos opcionalmente
- arquivos com estrutura divergente nao entram no consolidado do tipo correspondente
- divergencias e erros de leitura entram na aba `Inconsistencias`
- quando nenhuma aba e informada, o sistema tenta localizar automaticamente a planilha com mais linhas reais aproveitaveis
- o historico de saidas mostra quando o consolidado foi gerado para um protocolo especifico ou para todos os lotes
- uploads, historico e auditoria podem ser lidos com o mesmo recorte de periodo

## Seguranca e validacao

- CSRF em formularios
- comparacao segura da senha admin
- bloqueio temporario apos varias tentativas de login invalido
- expiracao da sessao administrativa por inatividade
- headers de seguranca HTTP
- politica `no-store` para respostas administrativas
- limite de frequencia por IP na pagina publica
- auditoria de downloads administrativos
- criptografia em repouso para uploads e consolidadores quando configurada
- limite de tamanho por requisicao
- validacao basica de e-mail e quantidade de arquivos
- opcao de bloquear arquivos duplicados por hash

## Rodar localmente

```bash
pip install -r requirements.txt
python app.py
```

Acesse:

- upload publico: `http://localhost:5000`
- admin: `http://localhost:5000/admin`

## Variaveis de ambiente

Use `.env.example` como referencia.

- `SECRET_KEY`: chave da sessao Flask
- `DATA_ENCRYPTION_KEY`: chave Fernet para criptografar uploads e outputs em disco

Exemplo para gerar a chave:

```bash
python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"
```
- `ADMIN_PASSWORD`: senha do painel admin
- `ADMIN_PASSWORD_HASH`: hash da senha admin com prioridade sobre a senha em texto puro
- `DATA_DIR`: pasta persistente com `uploads`, `outputs` e `app.db`
- `SHEET_NAME`: aba padrao para leitura; vazio usa a primeira aba
- `INCLUDE_SOURCE_COLUMNS`: inclui colunas de rastreabilidade no consolidado
- `MAX_CONTENT_LENGTH_MB`: limite de upload por requisicao
- `MAX_FILES_PER_UPLOAD`: limite de arquivos por envio
- `BLOCK_DUPLICATE_FILES`: bloqueia arquivos com hash ja existente
- `HEALTH_TEMPLATE_PATH`: caminho da planilha-base de Saude
- `DENTAL_TEMPLATE_PATH`: caminho da planilha-base de Odonto
- `SESSION_COOKIE_SECURE`: use `true` em producao com HTTPS
- `SESSION_LIFETIME_MINUTES`: validade maxima da sessao Flask
- `ADMIN_SESSION_IDLE_MINUTES`: tempo maximo de inatividade antes de exigir novo login
- `LOGIN_MAX_ATTEMPTS`: tentativas invalidas antes de bloqueio
- `LOGIN_LOCK_MINUTES`: janela de bloqueio do login admin
- `UPLOAD_RATE_LIMIT_COUNT`: quantidade maxima de tentativas de upload por IP dentro da janela
- `UPLOAD_RATE_LIMIT_WINDOW_MINUTES`: janela de tempo do limite de uploads
- `PORT`: porta do servidor

## Deploy

### Render

1. publique o repositorio como Web Service Python
2. configure um disco persistente
3. aponte `DATA_DIR` para esse disco, por exemplo `/var/data`
4. defina `SECRET_KEY`, `ADMIN_PASSWORD` e `SESSION_COOKIE_SECURE=true`
5. suba com `gunicorn app:app`

O arquivo `render.yaml` ja traz uma configuracao inicial.

### Railway ou VPS

Tambem funciona em qualquer ambiente Python com armazenamento persistente. O requisito mais importante e preservar a pasta de dados entre reinicios.

## Proximos passos recomendados

- filtros por periodo, remetente e protocolo no painel admin
- exclusao controlada de uploads concluida no painel admin
- autenticacao com usuarios e perfis
- barra de progresso no upload
- armazenamento externo em S3 ou Google Drive
- fila/processamento assincrono para volumes maiores
