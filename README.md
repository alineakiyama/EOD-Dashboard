# EOD Dropscale

Registro diário (EOD) da equipe de CS. Cada pessoa registra o que fez no dia, com
sugestões "com base em ontem", e a chefe acompanha, dá visto e faz perguntas no próprio EOD.

Montado do mesmo jeito que o CS Dashboard, para ser integrado a ele depois:

```
index.html (GitHub Pages, com login)
        │  POST { action, token, ... }
        ▼
Google Apps Script  ──►  Google Sheets (abas Usuarios / Sessoes / Catalogo / Registros / Vistos / Mensagens)
```

Página estática, sem build. Tudo é grátis (GitHub Pages + Google Sheets + Apps Script).

## Arquivos

| Arquivo | O que é |
|---|---|
| `index.html` | O site inteiro: HTML, CSS e JS num arquivo só. |
| `AppsScript.gs` | Código do Apps Script. Referência e backup: o que roda é o que está colado na planilha. |

## Instalação (uma vez)

### 1. Planilha e Apps Script
1. Crie uma planilha nova no Google Sheets (ex.: "EOD Dropscale").
2. Menu **Extensões → Apps Script**.
3. Apague o conteúdo do `Código.gs`, cole todo o `AppsScript.gs` e salve.
4. No topo, escolha a função **configurar** e clique em **Executar**. O Google pede autorização
   (Revisar permissões → sua conta → Avançado → Acessar → Permitir).
5. Abra **Registro de execução**: aparece o **código de primeiro acesso** (8 letras/números). Anote.

### 2. Publicar a API
1. **Implantar → Nova implantação** → engrenagem → **App da Web**.
2. Executar como: **Eu**. Quem pode acessar: **Qualquer pessoa**.
3. **Implantar** e copie a **URL do app da Web** (termina em `/exec`).

"Qualquer pessoa" só deixa o site conversar com a API. Os dados continuam protegidos pelo login:
sem e-mail e senha, a API não devolve nada.

### 3. Ligar o site à API
No `index.html`, cole a URL na linha:

```js
var API_URL = "https://script.google.com/macros/s/.../exec";
```

Faça commit e push pelo GitHub Desktop.

### 4. GitHub Pages
No github.com, no repositório: **Settings → Pages → Branch: main / (root) → Save**.
Em alguns minutos o site fica em `https://<usuario>.github.io/<repositorio>/`.

### 5. Primeiro acesso
Abra o site, clique em **Primeiro acesso**, digite o código do passo 1, seu nome, e-mail e uma senha.
Você vira admin. Depois, em **Pessoas e tarefas**, cadastre a equipe (e-mail, senha inicial, lojas).

## Atualizando o Apps Script
Colar o código novo no editor, salvar e **Implantar → Gerenciar implantações → lápis →
Versão: Nova versão → Implantar**. Sem isso a URL continua rodando a versão antiga.

## Como funciona

- **Meu dia**: busca de tarefas (catálogo compartilhado; se não existe, cria na hora) e o bloco
  **Com base em ontem**, com tudo do último dia registrado. Marcou, entrou no EOD de hoje.
  Tarefa longa que não chegou em **Concluído** continua sendo sugerida nos dias seguintes.
- **Dois tipos de tarefa**: *por quantidade* (e-mails, tickets: só o número) e *longa*
  (planilha, dashboard: status, tempo gasto e quantidade opcional).
- **Travada**: marca que falta algo para terminar e pede "o que falta". Aparece em âmbar para a chefe
  e tem filtro próprio.
- Tudo salva sozinho enquanto a pessoa preenche.
- **Equipe** (admin): quem registrou, quem falta, travadas, botão **Visto** e conversa por pessoa/dia.
  Somatório das tarefas por quantidade, por loja.
- **Histórico**: filtros por período, pessoa, loja, tarefa e só travadas; **Baixar CSV**.
- **Papéis**: *agente* vê só o próprio EOD; *admin* vê a equipe, dá visto, pergunta e cadastra pessoas.
  O corte é feito no Apps Script, não só na tela.

## Integração futura com o CS Dashboard
Mesma arquitetura e mesmos tokens de cor. Para integrar: as funções do `AppsScript.gs` entram no
Apps Script do dashboard (as abas Usuarios/Sessoes já existem lá, então o login passa a ser o mesmo),
as abas Catalogo/Registros/Vistos/Mensagens são copiadas para a planilha dele, e as telas viram uma
aba "EOD" no `index.html`. Depois disso, a quantidade de e-mails pode vir do contador automático.

Pendências de segunda fase: aviso no Slack quando a chefe pergunta; importar o histórico do Slack
para o catálogo inicial de tarefas.
