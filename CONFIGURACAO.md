# Vegas · Aprovação de Manutenção — Configuração

Sistema com 2 partes:
- **index.html** — a tela (formulário + pendentes + dashboard). Abra localmente com Live Server ou publique no GitHub Pages/Netlify, como nos seus outros projetos.
- **Code.gs** — o backend (Google Apps Script), que grava tudo numa aba "Solicitacoes" de uma planilha Google.

## 1. Publicar o backend (Code.gs)

1. Acesse [script.google.com](https://script.google.com) → **Novo projeto**
   (ou, se preferir vincular a uma planilha específica: abra a planilha → **Extensões → Apps Script**).
2. Apague o conteúdo do arquivo `Code.gs` que abrir e cole o conteúdo do arquivo `Code.gs` deste pacote.
3. No topo do editor, no seletor de funções, escolha `configurarInicial` e clique em **Executar** (▶). Na primeira execução o Google vai pedir autorização — aceite.
   - Isso cria a aba **Solicitacoes** com os cabeçalhos certos e define a senha padrão do dashboard como `vegas2026`.
4. **Implantar → Nova implantação**:
   - Tipo: **App da Web**
   - Executar como: **Eu**
   - Quem pode acessar: **Qualquer pessoa**
5. Copie a **URL do App da Web** gerada (termina em `/exec`).

### Trocar a senha do dashboard
No editor do Apps Script, na função `definirSenha`, rode manualmente passando a nova senha — ou mais simples: cole isto numa nova linha do editor e execute uma vez:
```javascript
definirSenha('SUA_NOVA_SENHA_AQUI');
```
Depois apague essa linha se quiser.

## 2. Configurar o index.html

1. Abra `index.html` num editor de texto.
2. Procure a linha:
   ```javascript
   const API_URL = "COLE_AQUI_A_URL_DO_APP_DA_WEB";
   ```
3. Substitua pelo link que você copiou no passo 1.5 (termina em `/exec`).
4. Mantenha a pasta `assets/logovegas.png` no mesmo diretório do `index.html` — o cabeçalho usa esse caminho relativo (`assets/logovegas.png`).

## 3. Publicar a tela

Igual aos seus outros projetos: pode rodar local com Live Server para testar, e depois subir para GitHub Pages ou Netlify. É só garantir que a pasta `assets/` vá junto.

## Como o fluxo funciona

1. **Qualquer pessoa** abre o índex e preenche "+ Nova Solicitação" com os dados do orçamento (peças, mão de obra, valores, com marcação verde/amarelo/vermelho por fornecedor). Isso vai direto para a planilha com status `Pendente`.
2. A solicitação aparece na lista **"Aguardando Assinatura"**, visível a todos.
3. Um diretor clica em **Aprovar Orçamento**, informa seu nome, a **senha da diretoria**, desenha a assinatura na tela e confirma.
4. Ao confirmar, o registro sai da lista de pendentes e passa para o **Dashboard** (status `Aprovado`, com nome do diretor, data/hora e a assinatura).
5. O Dashboard é protegido pela mesma senha. Lá é possível **exportar a solicitação em PDF** (réplica do formulário original, com as cores dos fornecedores) e **excluir definitivamente** o registro quando a transação estiver concluída.

## Estrutura da planilha (aba "Solicitacoes")

| Coluna | Conteúdo |
|---|---|
| id | identificador único (uuid) |
| numero | Nº sequencial da solicitação (0001, 0002...) |
| timestamp | data/hora de criação |
| data, solicitante, setor, veiculo | campos do formulário |
| descricaoJSON, pecasJSON, maoObraJSON | listas em formato JSON |
| obs | observações |
| status | `Pendente` ou `Aprovado` |
| diretor, dataAprovacao, assinatura | preenchidos apenas na aprovação (assinatura em base64) |

Se quiser trocar o número inicial da contagem ou resetar, edite a propriedade `CONTADOR` em **Configurações do projeto → Propriedades do script**, dentro do editor do Apps Script.
