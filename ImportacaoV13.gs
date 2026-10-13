/**
 * VEGAS VIGILÂNCIA E SEGURANÇA
 * Importação única: planilha CUSTO_MANUTENCAO_OFICIAL_V13 (abas Preventivas + Lançamentos)
 * para a Manutenção Preventiva do sistema.
 *
 * COMO USAR (no editor do Apps Script, mesmo projeto do Code.gs):
 *   1. Arquivos > "+" > Script > nome "ImportacaoV13" > cole este arquivo > Salvar.
 *   2. Rode  simularImportacaoV13   e confira em "Registro de execução" o que será gravado.
 *   3. Rode  importarPreventivasV13  para gravar de verdade (só roda uma vez).
 *   4. Se precisar voltar atrás: rode  desfazerImportacaoV13  (os registros importados
 *      são marcados como excluídos, sem apagar linhas; o hodômetro não volta).
 *   Depois de importado, este arquivo pode ser apagado do projeto.
 *
 * Regras aplicadas:
 *   - Viatura localizada pela placa. Se não existir, é cadastrada (aparece também no
 *     formulário de solicitações). Em viaturas já cadastradas, só preenche marca, modelo,
 *     tipo e ano que estiverem em branco; nunca sobrescreve.
 *   - Km da última troca vem da aba Preventivas; data e valor vêm da aba Lançamentos
 *     quando há lançamento correspondente. Sem correspondência, a troca entra sem data.
 *   - Lançamentos mais recentes que a aba Preventivas entram com a data; quando a planilha
 *     só tinha o odômetro atual naquela linha, o km é marcado como APROXIMADO.
 *   - Correia com 0 km na planilha = correia original de fábrica (próxima troca = intervalo).
 *   - KM ATUAL da aba Preventivas atualiza o hodômetro só se for maior que o do sistema.
 *   - Rodar de novo não duplica: cada registro é conferido antes de gravar.
 */
var IMPORT_V13_ORIGEM = 'Planilha V13';
var IMPORT_V13 = {"veiculos":[{"placa":"FKD6J09","marca":"Fiat","modelo":"Palio","tipo":"carro","ano":"2015","km":202490,"intervalos":{},"nome":"Palio"},{"placa":"FOT6A75","marca":"Fiat","modelo":"Palio","tipo":"carro","ano":"2015","km":170269,"intervalos":{},"nome":"Palio Furgão"},{"placa":"FPC8476","marca":"Fiat","modelo":"Palio","tipo":"carro","ano":"2015","km":197853,"intervalos":{},"nome":"Palio Técnica"},{"placa":"FTM1G47","marca":"Fiat","modelo":"Palio","tipo":"carro","ano":"2014","km":197599,"intervalos":{},"nome":"Palio Fire"},{"placa":"FUG5E67","marca":"Fiat","modelo":"Palio","tipo":"carro","ano":"2014","km":202374,"intervalos":{},"nome":"Palio Técnica"},{"placa":"KWU6034","marca":"Fiat","modelo":"Palio","tipo":"carro","ano":"2015","km":235640,"intervalos":{},"nome":"Palio Técnica"},{"placa":"LLT7B15","marca":"Honda","modelo":"CG Fan","tipo":"moto","ano":"2012","km":13012,"intervalos":{},"nome":"Moto CG Vermelha"},{"placa":"TTV0C50","marca":"Honda","modelo":"CG 160 Start","tipo":"moto","ano":"2015","km":15518,"intervalos":{},"nome":"CG 160 Start"},{"placa":"LMP9A93","marca":"Fiat","modelo":"Mobi","tipo":"carro","ano":"2018","km":188258,"intervalos":{},"nome":"Mobi"},{"placa":"LMP9C24","marca":"Fiat","modelo":"Mobi","tipo":"carro","ano":"2018","km":157261,"intervalos":{},"nome":"Mobi Técnica"},{"placa":"LQV5J10","marca":"Volkswagen","modelo":"Gol","tipo":"carro","ano":"2013","km":146299,"intervalos":{},"nome":"Gol"},{"placa":"LRT5A50","marca":"Fiat","modelo":"Palio","tipo":"carro","ano":"2013","km":160086,"intervalos":{},"nome":"Palio Técnica"},{"placa":"PUM5G69","marca":"Fiat","modelo":"Palio","tipo":"carro","ano":"2014","km":206099,"intervalos":{},"nome":"Palio Técnica"},{"placa":"PXP7J86","marca":"Volkswagen","modelo":"Gol","tipo":"carro","ano":"2016","km":169504,"intervalos":{},"nome":"Gol"},{"placa":"PYN2H36","marca":"Fiat","modelo":"Palio","tipo":"carro","ano":"2016","km":129354,"intervalos":{},"nome":"Palio"},{"placa":"QXO4I98","marca":"Fiat","modelo":"Mobi","tipo":"carro","ano":"2024","km":72827,"intervalos":{},"nome":"Mobi Like"},{"placa":"RFD3B47","marca":"Fiat","modelo":"Uno","tipo":"carro","ano":"2025","km":189715,"intervalos":{"comando":50000},"nome":"Uno"},{"placa":"RGA7G46","marca":"Fiat","modelo":"Mobi","tipo":"carro","ano":"2020","km":105731,"intervalos":{},"nome":"Mobi"},{"placa":"SHR8B73","marca":"Chevrolet","modelo":"Onix","tipo":"carro","ano":"2023","km":67666,"intervalos":{},"nome":"Onix"},{"placa":"STQ1H32","marca":"","modelo":"JTZ","tipo":"moto","ano":"2024","km":19360,"intervalos":{},"nome":"JTZ Preta"},{"placa":"SVH5G66","marca":"","modelo":"JTZ","tipo":"moto","ano":"2024","km":null,"intervalos":{},"nome":"JTZ Preta"},{"placa":"TEB3I22","marca":"Fiat","modelo":"Mobi","tipo":"carro","ano":"2025","km":68437,"intervalos":{},"nome":"Mobi"},{"placa":"LMH5F12","marca":"","modelo":"","tipo":"moto","ano":"2015","km":19357,"intervalos":{},"nome":"Moto"},{"placa":"SRJ1B52","marca":"Honda","modelo":"CG 160","tipo":"moto","ano":"2023","km":54875,"intervalos":{},"nome":"CG 160"},{"placa":"TDY6E08","marca":"Renault","modelo":"Kwid","tipo":"carro","ano":"2025","km":42338,"intervalos":{},"nome":"Kwid"},{"placa":"SRK5E55","marca":"Renault","modelo":"Kwid","tipo":"carro","ano":"2023","km":86003,"intervalos":{},"nome":"Kwid Fiscal Abreu"},{"placa":"SRO1E94","marca":"Honda","modelo":"CG 160 Start","tipo":"moto","ano":"2023","km":31226,"intervalos":{},"nome":"CG 160 Start"},{"placa":"SJG2C16","marca":"Chevrolet","modelo":"Onix","tipo":"carro","ano":"2023","km":69444,"intervalos":{},"nome":"Onix Comercial"},{"placa":"TXN5G41","marca":"Fiat","modelo":"Mobi","tipo":"carro","ano":"2025","km":26790,"intervalos":{},"nome":"Mobi"},{"placa":"SRM8A00","marca":"Renault","modelo":"Kwid","tipo":"carro","ano":"2023","km":73751,"intervalos":{},"nome":"Kwid"},{"placa":"SRH5D02","marca":"Renault","modelo":"Kwid","tipo":"carro","ano":"2023","km":159431,"intervalos":{},"nome":"Kwid Fiscal Alarme"},{"placa":"TXN5G04","marca":"Fiat","modelo":"Mobi","tipo":"carro","ano":"2023","km":44938,"intervalos":{},"nome":"Mobi Supervisão"},{"placa":"TXN5G43","marca":"Fiat","modelo":"Mobi","tipo":"carro","ano":"2023","km":42375,"intervalos":{},"nome":"Mobi Supervisão"},{"placa":"QUA8H36","marca":"Mitsubishi","modelo":"L200","tipo":"carro","ano":"2019","km":110113,"intervalos":{"comando":400000},"nome":"L200"},{"placa":"TEB3I73","marca":"Fiat","modelo":"Mobi","tipo":"carro","ano":"2023","km":52278,"intervalos":{},"nome":"Mobi Supervisão"},{"placa":"TEB3I89","marca":"Fiat","modelo":"Mobi","tipo":"carro","ano":"2023","km":49616,"intervalos":{},"nome":"Mobi Supervisão"},{"placa":"TEB3I83","marca":"Fiat","modelo":"Mobi","tipo":"carro","ano":"2023","km":41651,"intervalos":{},"nome":"Mobi"},{"placa":"RIZ0C40","marca":"Fiat","modelo":"Mobi","tipo":"carro","ano":"2014","km":114675,"intervalos":{},"nome":"Mobi"},{"placa":"TDY6G22","marca":"Renault","modelo":"Kwid","tipo":"carro","ano":"2025","km":23326,"intervalos":{},"nome":"Kwid"}],"registros":[{"placa":"FKD6J09","item":"oleo","data":"","km":198351,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"FKD6J09","item":"comando","data":"","km":151060,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"FKD6J09","item":"comando","data":"2026-09-28","km":202490,"valor":448.22,"oficina":"STILO CAR","descricao":"TROCA DA CORREIA DENTADA","obs":"Importado da planilha V13 (Lançamentos linha 70). KM APROXIMADO: nesta linha a planilha mostra o odômetro atual do veículo, não o do dia do serviço. Confirme.","flag":"km aproximado"},{"placa":"FOT6A75","item":"oleo","data":"2026-06-22","km":167629,"valor":103.6,"oficina":"ROBERTO MECÂNICO","descricao":"Óleo e filtro","obs":"Importado da planilha V13 (Preventivas + Lançamentos linha 14).","flag":""},{"placa":"FOT6A75","item":"comando","data":"2026-08-19","km":168788,"valor":567.09,"oficina":"ROBERTO MECÂNICO","descricao":"ROLAMENTO ESTICADOR DENTADO, CORREIA DE DISTRIBUIÇÃO, CORREIA DO AR CONDICIONADO, QUEROSENE, RETENTOR DO EIXO DE COMANDO E RETENDOR DO MOTOR DIANTEIRO","obs":"Importado da planilha V13 (Preventivas + Lançamentos linha 42).","flag":""},{"placa":"FPC8476","item":"oleo","data":"","km":192260,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"FPC8476","item":"comando","data":"","km":158739,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"FTM1G47","item":"oleo","data":"","km":194362,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"FTM1G47","item":"comando","data":"2026-08-13","km":196842,"valor":706.42,"oficina":"EUCAR AUTO MECÂNICA","descricao":"ROLAMENTO ESTICADOR DENTADO, CORREIA DE DISTRIBUIÇÃO, CORREIA DO AR CONDICIONADO, QUEROSENE, RETENTOR DO EIXO DE COMANDO E RETENDOR DO MOTOR DIANTEIRO","obs":"Importado da planilha V13 (Preventivas + Lançamentos linha 37).","flag":""},{"placa":"FUG5E67","item":"oleo","data":"","km":198154,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"FUG5E67","item":"comando","data":"2026-08-25","km":200561,"valor":567.09,"oficina":"ROBERTO MECÂNICO","descricao":"ROLAMENTO ESTICADOR DENTADO, CORREIA DE DISTRIBUIÇÃO, CORREIA DO AR CONDICIONADO, QUEROSENE, RETENTOR DO EIXO DE COMANDO E RETENDOR DO MOTOR DIANTEIRO","obs":"Importado da planilha V13 (Preventivas + Lançamentos linha 45). Lançamentos registra 200.587 km; mantido o km da aba Preventivas.","flag":""},{"placa":"KWU6034","item":"oleo","data":"","km":235640,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"KWU6034","item":"oleo","data":"2026-07-31","km":230310,"valor":103.6,"oficina":"ROBERTO MECÂNICO","descricao":"Troca de óleo","obs":"Importado da planilha V13 (Lançamentos linha 23).","flag":""},{"placa":"KWU6034","item":"comando","data":"2026-09-29","km":235640,"valor":762.25,"oficina":"STILO CAR","descricao":"TROCA DA BOMBA DE OLEO E CORREIA DENTADA","obs":"Importado da planilha V13 (Preventivas + Lançamentos linha 75).","flag":""},{"placa":"LLT7B15","item":"oleo","data":"2026-08-03","km":11238,"valor":0.0,"oficina":"SILVA MOTOS","descricao":"Pneu traseiro, cabo de velocímetro, freio traseiro e troca de óleo","obs":"Importado da planilha V13 (Preventivas + Lançamentos linha 25). Lançamentos registra 112.382 km (dígito a mais); usado o km da aba Preventivas. O valor refere-se ao serviço completo descrito.","flag":""},{"placa":"TTV0C50","item":"oleo","data":"","km":12540,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"TTV0C50","item":"oleo","data":"2026-09-02","km":15518,"valor":67.51,"oficina":"SILVA MOTOS","descricao":"TROCA DE OLEO E REPARO E LUBRIFICAÇÃO NA CORRENTE","obs":"Importado da planilha V13 (Lançamentos linha 67). KM APROXIMADO: nesta linha a planilha mostra o odômetro atual do veículo, não o do dia do serviço. Confirme.","flag":"km aproximado"},{"placa":"LMP9A93","item":"oleo","data":"","km":187958,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"LMP9A93","item":"comando","data":"","km":147934,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"LMP9C24","item":"comando","data":"","km":0,"valor":"","oficina":"","descricao":"Correia/corrente original de fábrica","obs":"Importado da planilha V13: a aba Preventivas indica 0 km (sem troca desde o veículo novo).","flag":"fabrica"},{"placa":"LQV5J10","item":"oleo","data":"","km":146229,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"LQV5J10","item":"comando","data":"","km":145000,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"LRT5A50","item":"oleo","data":"2026-08-10","km":158694,"valor":103.6,"oficina":"ROBERTO MECÂNICO","descricao":"TROCA DO OLEO E O FILTRO","obs":"Importado da planilha V13 (Preventivas + Lançamentos linha 33). Lançamentos registra 158.639 km; mantido o km da aba Preventivas.","flag":""},{"placa":"LRT5A50","item":"comando","data":"","km":105290,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"LRT5A50","item":"comando","data":"2026-09-28","km":160086,"valor":498.22,"oficina":"ROBERTO MECÂNICO","descricao":"TROCA DA CORREIA DENTADA","obs":"Importado da planilha V13 (Lançamentos linha 74). KM APROXIMADO: nesta linha a planilha mostra o odômetro atual do veículo, não o do dia do serviço. Confirme.","flag":"km aproximado"},{"placa":"PUM5G69","item":"oleo","data":"","km":193230,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"PUM5G69","item":"oleo","data":"2026-09-28","km":206099,"valor":102.1,"oficina":"ROBERTO MECÂNICO","descricao":"TROCA DE OLEO","obs":"Importado da planilha V13 (Lançamentos linha 69). KM APROXIMADO: nesta linha a planilha mostra o odômetro atual do veículo, não o do dia do serviço. Confirme.","flag":"km aproximado"},{"placa":"PUM5G69","item":"comando","data":"","km":190920,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"PXP7J86","item":"oleo","data":"","km":168596,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"PXP7J86","item":"comando","data":"2026-09-11","km":169504,"valor":651.0,"oficina":"EUCAR AUTO MECÂNICA","descricao":"TROCA DA CORREIA DENTADA","obs":"Importado da planilha V13 (Preventivas + Lançamentos linha 54).","flag":""},{"placa":"PYN2H36","item":"oleo","data":"","km":122299,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"PYN2H36","item":"comando","data":"","km":170237,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"QXO4I98","item":"oleo","data":"","km":63677,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"QXO4I98","item":"comando","data":"","km":0,"valor":"","oficina":"","descricao":"Correia/corrente original de fábrica","obs":"Importado da planilha V13: a aba Preventivas indica 0 km (sem troca desde o veículo novo).","flag":"fabrica"},{"placa":"RFD3B47","item":"oleo","data":"","km":188294,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"RFD3B47","item":"comando","data":"","km":150000,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"RGA7G46","item":"oleo","data":"","km":103286,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"RGA7G46","item":"comando","data":"","km":0,"valor":"","oficina":"","descricao":"Correia/corrente original de fábrica","obs":"Importado da planilha V13: a aba Preventivas indica 0 km (sem troca desde o veículo novo).","flag":"fabrica"},{"placa":"TEB3I22","item":"comando","data":"","km":0,"valor":"","oficina":"","descricao":"Correia/corrente original de fábrica","obs":"Importado da planilha V13: a aba Preventivas indica 0 km (sem troca desde o veículo novo).","flag":"fabrica"},{"placa":"LMH5F12","item":"oleo","data":"","km":18876,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"LMH5F12","item":"oleo","data":"2026-08-22","km":19357,"valor":89.99,"oficina":"SILVA MOTOS","descricao":"MOTO ESTAVA SEM FREIO E TROCA DO OLEO","obs":"Importado da planilha V13 (Lançamentos linha 72). KM APROXIMADO: nesta linha a planilha mostra o odômetro atual do veículo, não o do dia do serviço. Confirme. O valor refere-se ao serviço completo descrito.","flag":"km aproximado"},{"placa":"TDY6E08","item":"comando","data":"","km":0,"valor":"","oficina":"","descricao":"Correia/corrente original de fábrica","obs":"Importado da planilha V13: a aba Preventivas indica 0 km (sem troca desde o veículo novo).","flag":"fabrica"},{"placa":"SRK5E55","item":"oleo","data":"","km":80541,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"SRK5E55","item":"oleo","data":"2026-08-10","km":null,"valor":154.81,"oficina":"ROBERTO MECÂNICO","descricao":"TROCA DO OLEO E O FILTRO","obs":"Importado da planilha V13 (Lançamentos linha 32). Km não importado: sem km na planilha.","flag":"sem km"},{"placa":"SRK5E55","item":"comando","data":"","km":0,"valor":"","oficina":"","descricao":"Correia/corrente original de fábrica","obs":"Importado da planilha V13: a aba Preventivas indica 0 km (sem troca desde o veículo novo).","flag":"fabrica"},{"placa":"SRO1E94","item":"oleo","data":"","km":30096,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"SRO1E94","item":"oleo","data":"2026-07-30","km":null,"valor":294.99,"oficina":"SILVA MOTOS","descricao":"Troca de óleo, relação e troca de pneu dianteiro","obs":"Importado da planilha V13 (Lançamentos linha 22). Km não importado: km 96.582 maior que o odômetro atual (31.226), provável erro de digitação. O valor refere-se ao serviço completo descrito.","flag":"sem km"},{"placa":"SRO1E94","item":"oleo","data":"2026-08-14","km":null,"valor":59.2,"oficina":"SILVA MOTOS","descricao":"TROCA DO OLEO DO MOTOR E REPARA NA CORRENTE","obs":"Importado da planilha V13 (Lançamentos linha 71). Km não importado: a planilha mostra o odômetro atual nesta linha, não o do dia do serviço.","flag":"sem km"},{"placa":"SRO1E94","item":"oleo","data":"2026-09-01","km":31226,"valor":67.51,"oficina":"SILVA MOTOS","descricao":"TROCA DE OLEO E REPARO E LUBRIFICAÇÃO NA CORRENTE","obs":"Importado da planilha V13 (Lançamentos linha 66). KM APROXIMADO: nesta linha a planilha mostra o odômetro atual do veículo, não o do dia do serviço. Confirme.","flag":"km aproximado"},{"placa":"SJG2C16","item":"oleo","data":"2026-08-10","km":67830,"valor":212.59,"oficina":"ROBERTO MECÂNICO","descricao":"TROCA DO OLEO E O FILTRO","obs":"Importado da planilha V13 (Preventivas + Lançamentos linha 31). Lançamentos registra 67.240 km; mantido o km da aba Preventivas.","flag":""},{"placa":"SJG2C16","item":"comando","data":"","km":0,"valor":"","oficina":"","descricao":"Correia/corrente original de fábrica","obs":"Importado da planilha V13: a aba Preventivas indica 0 km (sem troca desde o veículo novo).","flag":"fabrica"},{"placa":"TXN5G41","item":"oleo","data":"2026-08-10","km":21376,"valor":104.7,"oficina":"ROBERTO MECÂNICO","descricao":"TROCA DO OLEO E O FILTRO","obs":"Importado da planilha V13 (Preventivas + Lançamentos linha 34).","flag":""},{"placa":"TXN5G41","item":"comando","data":"","km":0,"valor":"","oficina":"","descricao":"Correia/corrente original de fábrica","obs":"Importado da planilha V13: a aba Preventivas indica 0 km (sem troca desde o veículo novo).","flag":"fabrica"},{"placa":"SRM8A00","item":"oleo","data":"","km":64802,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"SRM8A00","item":"comando","data":"","km":0,"valor":"","oficina":"","descricao":"Correia/corrente original de fábrica","obs":"Importado da planilha V13: a aba Preventivas indica 0 km (sem troca desde o veículo novo).","flag":"fabrica"},{"placa":"SRH5D02","item":"oleo","data":"","km":151938,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"SRH5D02","item":"comando","data":"","km":0,"valor":"","oficina":"","descricao":"Correia/corrente original de fábrica","obs":"Importado da planilha V13: a aba Preventivas indica 0 km (sem troca desde o veículo novo).","flag":"fabrica"},{"placa":"TXN5G04","item":"oleo","data":"","km":43270,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"TXN5G04","item":"oleo","data":"2026-08-10","km":null,"valor":104.7,"oficina":"ROBERTO MECÂNICO","descricao":"TROCA DO OLEO E O FILTRO","obs":"Importado da planilha V13 (Lançamentos linha 35). Km não importado: sem km na planilha.","flag":"sem km"},{"placa":"TXN5G04","item":"oleo","data":"2026-09-28","km":44938,"valor":156.35,"oficina":"ROBERTO MECÂNICO","descricao":"TROCA DE OLEO","obs":"Importado da planilha V13 (Lançamentos linha 68). KM APROXIMADO: nesta linha a planilha mostra o odômetro atual do veículo, não o do dia do serviço. Confirme.","flag":"km aproximado"},{"placa":"TXN5G04","item":"comando","data":"","km":0,"valor":"","oficina":"","descricao":"Correia/corrente original de fábrica","obs":"Importado da planilha V13: a aba Preventivas indica 0 km (sem troca desde o veículo novo).","flag":"fabrica"},{"placa":"TXN5G43","item":"oleo","data":"2026-08-10","km":10699,"valor":104.7,"oficina":"ROBERTO MECÂNICO","descricao":"TROCA DO OLEO E O FILTRO","obs":"Importado da planilha V13 (Lançamentos linha 36).","flag":""},{"placa":"TXN5G43","item":"comando","data":"","km":0,"valor":"","oficina":"","descricao":"Correia/corrente original de fábrica","obs":"Importado da planilha V13: a aba Preventivas indica 0 km (sem troca desde o veículo novo).","flag":"fabrica"},{"placa":"QUA8H36","item":"oleo","data":"","km":108900,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"QUA8H36","item":"comando","data":"","km":0,"valor":"","oficina":"","descricao":"Correia/corrente original de fábrica","obs":"Importado da planilha V13: a aba Preventivas indica 0 km (sem troca desde o veículo novo).","flag":"fabrica"},{"placa":"TEB3I73","item":"oleo","data":"","km":41545,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"TEB3I73","item":"comando","data":"","km":0,"valor":"","oficina":"","descricao":"Correia/corrente original de fábrica","obs":"Importado da planilha V13: a aba Preventivas indica 0 km (sem troca desde o veículo novo).","flag":"fabrica"},{"placa":"TEB3I89","item":"oleo","data":"","km":45984,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"TEB3I89","item":"comando","data":"","km":0,"valor":"","oficina":"","descricao":"Correia/corrente original de fábrica","obs":"Importado da planilha V13: a aba Preventivas indica 0 km (sem troca desde o veículo novo).","flag":"fabrica"},{"placa":"TEB3I83","item":"comando","data":"","km":0,"valor":"","oficina":"","descricao":"Correia/corrente original de fábrica","obs":"Importado da planilha V13: a aba Preventivas indica 0 km (sem troca desde o veículo novo).","flag":"fabrica"},{"placa":"RIZ0C40","item":"comando","data":"","km":0,"valor":"","oficina":"","descricao":"Correia/corrente original de fábrica","obs":"Importado da planilha V13: a aba Preventivas indica 0 km (sem troca desde o veículo novo).","flag":"fabrica"},{"placa":"TDY6G22","item":"oleo","data":"","km":22600,"valor":"","oficina":"","descricao":"","obs":"Importado da planilha V13 (aba Preventivas). A planilha não informa a data desta troca.","flag":"sem data"},{"placa":"TDY6G22","item":"comando","data":"","km":0,"valor":"","oficina":"","descricao":"Correia/corrente original de fábrica","obs":"Importado da planilha V13: a aba Preventivas indica 0 km (sem troca desde o veículo novo).","flag":"fabrica"}]};

function simularImportacaoV13() { return executarImportacaoV13_(true); }
function importarPreventivasV13() { return executarImportacaoV13_(false); }

function executarImportacaoV13_(simular) {
  var props = PropertiesService.getScriptProperties();
  var feito = props.getProperty('IMPORT_V13_FEITO');
  if (!simular && feito) {
    throw new Error('A importação da planilha V13 já foi feita em ' + feito + '. Para repetir, rode desfazerImportacaoV13 antes.');
  }
  return withLock_(function () {
    var agora = agoraIso_();
    var shV = sheetVeiculosMnt_();
    var veics = lerVeiculosMnt_(shV);
    if (!simular) { try { sincronizarKmSolicitacoes_(shV, veics); } catch (e) { /* segue */ } }
    var headV = cabecalho_(shV);
    var porPlaca = {};
    veics.forEach(function (v) { if (v.placa) porPlaca[placaNorm_(v.placa)] = v; });

    var shM = sheetManut_();
    garantirCabecalhos_(shM, HEADERS_MANUT.concat(['origem']));
    var jaTem = {};
    lerTabela_(shM).forEach(function (r) {
      if (String(r.excluido) === 'Sim') return;
      jaTem[[placaNorm_(r.placa), r.item, num_(r.km) === null ? '' : num_(r.km), txt_(r.data).split('T')[0]].join('|')] = true;
    });

    var log = [], novasViaturas = [], novasLinhasV = [], kmLeituras = [], linhasM = [];
    var res = { viaturasNovas: 0, viaturasCompletadas: 0, hodometros: 0, registros: 0, repetidos: 0 };
    var proxLinhaV = shV.getLastRow() + 1;

    IMPORT_V13.veiculos.forEach(function (d) {
      var chave = placaNorm_(d.placa);
      var v = porPlaca[chave];
      if (!v) {
        v = { id: Utilities.getUuid(), veiculo: d.nome, placa: d.placa, tipo: d.tipo, kmAtual: null, _row: proxLinhaV++, novo: true };
        novasLinhasV.push({
          veiculo: d.nome, placa: d.placa, timestamp: agora, id: v.id, marca: d.marca, modelo: d.modelo, tipo: d.tipo,
          ano: d.ano, comando: '', intervalosJSON: JSON.stringify(d.intervalos || {}), atualizadoEm: agora
        });
        porPlaca[chave] = v;
        res.viaturasNovas++;
        log.push('NOVA VIATURA  ' + d.placa + '  ' + d.nome + ' (' + (d.marca ? d.marca + ' ' : '') + d.modelo + ', ' + d.tipo + ', ' + d.ano + ')');
      } else {
        var campos = {};
        if (v.inferido.marca && d.marca) campos.marca = d.marca;
        if (v.inferido.modelo && d.modelo) campos.modelo = d.modelo;
        if (v.inferido.tipo && d.tipo) campos.tipo = d.tipo;
        if (!v.ano && d.ano) campos.ano = d.ano;
        var iv = v.intervalos || {}, mudouIv = false;
        Object.keys(d.intervalos || {}).forEach(function (k) { if (!(num_(iv[k]) > 0)) { iv[k] = d.intervalos[k]; mudouIv = true; } });
        if (mudouIv) campos.intervalosJSON = JSON.stringify(iv);
        if (Object.keys(campos).length) {
          campos.atualizadoEm = agora;
          if (!simular) gravarCampos_(shV, headV, v._row, campos);
          res.viaturasCompletadas++;
          log.push('COMPLETA CADASTRO  ' + d.placa + '  ' + v.veiculo + ': ' + Object.keys(campos).filter(function (k) { return k !== 'atualizadoEm'; }).join(', '));
        }
        if (campos.tipo) v.tipo = campos.tipo;
      }
      if (d.km && (v.kmAtual === null || d.km > v.kmAtual)) {
        kmLeituras.push({ v: v, km: d.km, anterior: v.kmAtual });
        log.push('HODÔMETRO  ' + d.placa + ': ' + (v.kmAtual === null ? 'sem leitura' : v.kmAtual + ' km') + ' -> ' + d.km + ' km');
        v.kmAtual = d.km;
        res.hodometros++;
      }
    });

    IMPORT_V13.registros.forEach(function (r) {
      var v = porPlaca[placaNorm_(r.placa)];
      var k = [placaNorm_(r.placa), r.item, r.km === null ? '' : r.km, r.data].join('|');
      if (jaTem[k]) { res.repetidos++; return; }
      jaTem[k] = true;
      linhasM.push({
        id: Utilities.getUuid(), veiculoId: v.id, placa: v.placa || r.placa, veiculo: v.veiculo, item: r.item, posicoes: '',
        descricao: r.descricao, data: r.data, km: r.km === null ? '' : r.km, valor: r.valor, oficina: r.oficina, obs: r.obs,
        comprovante: '', registradoEm: agora, tipoVeiculo: v.tipo || '', intervaloAplicado: '', proximaKm: '',
        excluido: '', origem: IMPORT_V13_ORIGEM
      });
      res.registros++;
      log.push('REGISTRO  ' + r.placa + '  ' + (r.item === 'oleo' ? 'Óleo' : 'Correia') + '  ' + (r.data || 'sem data') + '  ' +
        (r.km === null ? 'km não informado' : r.km + ' km') + (r.valor !== '' ? '  R$ ' + r.valor : '') + (r.flag ? '  [' + r.flag + ']' : ''));
    });

    if (!simular) {
      if (novasLinhasV.length) anexarLinhasLote_(shV, novasLinhasV, TEXTO_VEICULOS_MNT);
      kmLeituras.forEach(function (l) {
        gravarCampos_(shV, headV, l.v._row, { kmAtual: l.km, kmAtualizadoEm: agora, kmOrigem: IMPORT_V13_ORIGEM, kmRegistradoEm: agora });
      });
      if (kmLeituras.length) anexarLinhasLote_(sheetHodometro_(), kmLeituras.map(function (l) {
        return {
          id: Utilities.getUuid(), veiculoId: l.v.id, placa: l.v.placa, data: agora, km: l.km, kmAnterior: l.anterior === null ? '' : l.anterior,
          origem: 'importacao', ref: IMPORT_V13_ORIGEM, responsavel: '', obs: 'KM ATUAL da aba Preventivas da planilha V13 (data da leitura não informada na planilha)',
          regressao: '', autorizadoPor: '', registradoEm: agora
        };
      }), TEXTO_HODOMETRO);
      if (linhasM.length) anexarLinhasLote_(shM, linhasM, TEXTO_MANUT);
      props.setProperty('IMPORT_V13_FEITO', agora);
      painelInterno_(null); // recalcula previsões e abre os alertas
    }
    var titulo = (simular ? 'SIMULAÇÃO (nada foi gravado)' : 'IMPORTAÇÃO CONCLUÍDA') + ': ' +
      res.viaturasNovas + ' viaturas novas, ' + res.viaturasCompletadas + ' cadastros completados, ' +
      res.hodometros + ' hodômetros atualizados, ' + res.registros + ' trocas registradas, ' + res.repetidos + ' já existiam.';
    console.log(titulo + '\n\n' + log.join('\n'));
    res.resumo = titulo;
    return res;
  });
}

// Marca como excluídos (exclusão lógica) os registros que vieram da planilha V13.
function desfazerImportacaoV13() {
  return withLock_(function () {
    var sh = sheetManut_();
    var head = cabecalho_(sh);
    var n = 0, agora = agoraIso_();
    lerTabela_(sh).forEach(function (r) {
      if (String(r.origem) !== IMPORT_V13_ORIGEM || String(r.excluido) === 'Sim') return;
      gravarCampos_(sh, head, r._row, { excluido: 'Sim', excluidoEm: agora, excluidoPor: 'Importação desfeita', motivoExclusao: 'Importação da planilha V13 desfeita' });
      n++;
    });
    PropertiesService.getScriptProperties().deleteProperty('IMPORT_V13_FEITO');
    painelInterno_(null);
    console.log(n + ' registros da planilha V13 marcados como excluídos. Viaturas cadastradas e leituras de hodômetro foram mantidas.');
    return n;
  });
}

function gravarCampos_(sh, head, row, campos) {
  Object.keys(campos).forEach(function (k) {
    var i = head.indexOf(k);
    if (i < 0) return;
    var cel = sh.getRange(row, i + 1);
    if (['kmAtualizadoEm', 'kmRegistradoEm', 'atualizadoEm', 'excluidoEm', 'intervalosJSON', 'ano'].indexOf(k) >= 0) cel.setNumberFormat('@');
    cel.setValue(campos[k] === null || campos[k] === undefined ? '' : campos[k]);
  });
}

function anexarLinhasLote_(sh, objs, colsTexto) {
  var head = cabecalho_(sh);
  var ini = sh.getLastRow() + 1;
  (colsTexto || []).forEach(function (c) {
    var i = head.indexOf(c);
    if (i >= 0) sh.getRange(ini, i + 1, objs.length, 1).setNumberFormat('@');
  });
  var linhas = objs.map(function (o) {
    return head.map(function (h) { return o.hasOwnProperty(h) && o[h] !== null && o[h] !== undefined ? o[h] : ''; });
  });
  sh.getRange(ini, 1, linhas.length, head.length).setValues(linhas);
}
