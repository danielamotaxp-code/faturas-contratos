# WattWise · Faturas e Contratos

Plataforma interna da WattWise para conferir, toda semana, quais clientes estão sem fatura gerada. Ela funciona sozinha: os colaboradores colam o link da planilha e a plataforma gera a planilha conferida na hora.

| Aba | O que recebe | O que gera |
|---|---|---|
| **Faturas Pendentes** | Link da planilha mensal de faturas (ex.: *Clientes Não Faturados – OUT 2026*), com abas *Semana 1*, *Semana 2*… | Planilha **Faturas Pendentes – Outubro 2026 – Semana N** com todas as semanas. Instalações com fatura do mês ficam com todas as linhas em **azul**, e as que ainda não têm ficam em **branco**. Inclui as abas *Resumo* e *Pendentes – Semana N*. |
| **Contratos sem Faturas** | Link da planilha com *Contratos – Semana N* e *Faturas – Semana N* (gateway) | Planilha **Contratos sem Faturas – Outubro 2026 – Semana N**, só com os contratos vinculados que não têm fatura, uma aba por semana, mais a aba *Resumo*. |
| **Progresso** | Histórico de todas as conferências | Números por semana e por mês: sem fatura, resolvidos, novos, que continuam e cobertura. Também traz a lista de clientes recorrentes sem fatura, que pode ser baixada. |

## Regras

- **O mês vem do nome da planilha.** O nome precisa ter o mês e o ano, por exemplo *OUT 2026*.
- **O cliente é identificado pela instalação** (*Instalação* / *Número instalação*).
- **Fatura com status *Cancelado* conta como fatura gerada.**
- Cada envio gera uma **planilha nova**. A planilha da semana anterior continua salva e com as cores dela.
- Os links da planilha enviada e da planilha gerada ficam salvos na plataforma, organizados por mês.

## Como funciona

- **Site** (esta pasta, publicado pelo GitHub Pages): `index.html` e `assets/`. Não guarda nenhum dado de cliente.
- **Motor** (`apps-script/`): roda no Google Apps Script da conta dona das planilhas. Ele lê as planilhas, gera as novas no Google Drive (pasta *WattWise – Faturas e Contratos*) e guarda o histórico numa planilha-banco.
- **Logins**: ficam no arquivo `Config.gs`, que está no repositório **privado** `faturas-contratos-acessos`. Esse arquivo nunca entra aqui.

## Instalação (uma vez)

1. Acesse <https://script.google.com>, crie um projeto e dê o nome **WattWise – Faturas e Contratos**.
2. Cole o conteúdo de `apps-script/Codigo.gs` no arquivo `Código.gs`.
3. Crie um arquivo de script chamado `Config` e cole nele o `Config.gs` do repositório privado.
4. Em *Configurações do projeto*, marque **"Mostrar arquivo de manifesto appsscript.json"** e cole o conteúdo de `apps-script/appsscript.json`.
5. Selecione a função `configurarPlataforma` e clique em **Executar**. Autorize o acesso ao Drive e ao Sheets. A pasta e a planilha-banco são criadas nesse momento.
6. Clique em **Implantar → Nova implantação → App da Web**, com *Executar como: eu* e *Quem pode acessar: qualquer pessoa*. Copie a URL gerada.
7. Cole a URL em `assets/config.js` (campo `API_URL`).

> Se o motor mudar, faça **Implantar → Gerenciar implantações → editar → Nova versão**. Assim a URL continua a mesma.

**Importante:** a conta que fez a implantação precisa ter acesso às planilhas cujos links forem colados na plataforma.
