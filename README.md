# WattWise · Faturas e Contratos

Plataforma interna da WattWise para conferir, toda semana, quais clientes estão sem fatura gerada.

Site: https://danielamotaxp-code.github.io/faturas-contratos/

Ela funciona sozinha. A colaboradora **cola o link da planilha do Google** (ou envia o arquivo .xlsx), e a plataforma confere e gera a planilha nova na hora, dentro do próprio navegador. A planilha enviada, guardada como uma cópia daquele momento, e a gerada ficam salvas para consultas e downloads futuros.

### Envio pelo link do Google

- A planilha precisa estar em **Compartilhar → Acesso geral → Qualquer pessoa com o link (Leitor)**.
- A leitura usa a API oficial do Google Sheets, com uma chave de API cadastrada em **Configurações**. A chave é gratuita, é feita uma única vez e fica restrita a este site e à Google Sheets API. Os passos estão na própria tela.
- A plataforma lê o nome da planilha (de onde tira o mês), todas as abas e os valores exatos.

| Aba | O que recebe | O que gera |
|---|---|---|
| **Faturas Pendentes** | Arquivo da planilha mensal de faturas (ex.: *Clientes Não Faturados – OUT 2026.xlsx*), com abas *Semana 1*, *Semana 2*… | **Faturas Pendentes – Outubro 2026 – Semana N.xlsx**, com todas as semanas. Instalações com fatura do mês ficam com todas as linhas em **azul** e as que ainda não têm ficam em **branco**. Inclui as abas *Resumo* e *Pendentes – Semana N*. |
| **Contratos sem Faturas** | Arquivo com as abas *Contratos – Semana N* e *Faturas – Semana N* (gateway) | **Contratos sem Faturas – Outubro 2026 – Semana N.xlsx**, só com os contratos vinculados que não têm fatura. Tem uma aba por semana e a aba *Resumo*. |
| **Progresso** | Histórico de todas as conferências | Números por semana e por mês (sem fatura, resolvidos, novos, que continuam e cobertura) e a lista de clientes recorrentes sem fatura, que pode ser baixada. |

## Regras

- **O mês vem do nome do arquivo.** O nome precisa ter o mês e o ano, por exemplo *OUT 2026*.
- **O cliente é identificado pela instalação** (*Instalação* / *Número instalação*).
- **Fatura com status *Cancelado* conta como fatura gerada.**
- Cada envio gera uma **planilha nova**. As anteriores continuam guardadas, com as cores delas.
- **CPF/CNPJ sempre completos.** Quando a planilha original guarda o documento como número e perde os zeros do começo, a plataforma recoloca esses zeros e confere os dígitos verificadores. Instalação e unidade consumidora são gravadas como texto, sem notação científica.
- A lista de planilhas é organizada por **mês → Semana 1 a 4 (ou mais)**. Cada semana tem o botão **+ Inserir planilha**, e cada planilha tem o botão **Excluir**, que pede confirmação. Ao excluir um envio, o progresso daquela semana volta para o envio anterior, se houver um.
- No Google Planilhas, baixe o arquivo em **Arquivo → Fazer download → Microsoft Excel (.xlsx)**, para que todas as abas venham juntas.

## Onde fica cada coisa

| Repositório | Visibilidade | Conteúdo |
|---|---|---|
| `faturas-contratos` | público | O site (`index.html`, `assets/`) e o `assets/cofre.json`, que tem só o token **criptografado**. Não guarda nenhum dado de cliente. |
| `faturas-contratos-acessos` | **privado** | Os dados, na pasta `dados/`: planilhas enviadas e geradas (`dados/arquivos/AAAA-MM/`), lista de planilhas (`indice.json`), histórico do progresso (`historico.json`) e cadastro dos clientes pendentes (`clientes.json`). |

### Arquivos do site

- `assets/regras.js`: regras de conferência (cores, cruzamento, progresso).
- `assets/planilhas.js`: lê o arquivo enviado e gera o .xlsx com as cores.
- `assets/armazenamento.js`: grava e lê os arquivos no repositório privado e cuida do cofre de acesso.
- `assets/app.js`: as telas.
- `assets/vendor/`: biblioteca de planilhas [xlsx-js-style](https://github.com/gitbrent/xlsx-js-style), licença Apache 2.0.

## Acesso

O login tem as usuárias **Daniela**, **Paula** e **Ytala**, com uma senha da equipe.

A senha abre o cofre (AES-256-GCM, com chave derivada por PBKDF2 em 600 mil iterações). Dentro dele está o token do GitHub que lê e grava no repositório privado.

### Configuração inicial ou troca de senha

1. No GitHub, acesse **Settings → Developer settings → Personal access tokens → Fine-grained tokens → Generate new token**.
   - Em *Repository access*, escolha **Only select repositories** e marque `faturas-contratos` e `faturas-contratos-acessos`.
   - Em *Permissions → Repository permissions → Contents*, escolha **Read and write**.
   - Escolha a validade mais longa disponível.
2. No site, clique em **Configurar acesso / trocar senha**. Cole o token. A senha da equipe já vem preenchida com **102030**.
3. A plataforma confere o token, criptografa e publica o cofre. Em cerca de 1 minuto, as demais usuárias já conseguem entrar.

Quando o token vencer, repita esses passos. Os dados já guardados não são perdidos.
