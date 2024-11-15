/*******************************
 * Script: InserirIngressoPlanilha.gs
 * 
 * Descrição: Automatiza a inserção de ingressantes na planilha "Kit Ingresso".
 * 
 * Configuração:
 * - Substitua os valores de exemplo pelos IDs reais das planilhas.
 *******************************/

// IDs das planilhas
const PLANILHA_ORIGEM_ID = '1abcXyzExampleIdForPublicRepo'; // ID fictício da planilha de origem
const PLANILHA_DESTINO_ID = '2defExampleIdForKitIngressoRepo'; // ID fictício da planilha de destino

// Nome das abas
const ABA_ORIGEM = 'Respostas ao formulário 1';
const ABA_DESTINO = 'Kit Ingresso';

/**
 * Função que insere ingressantes na planilha "Kit Ingresso" com base nas respostas do formulário.
 */
function inserirIngressoPlanilha() {
  // Obter a planilha de origem e aba de origem
  const planilhaOrigem = SpreadsheetApp.openById(PLANILHA_ORIGEM_ID);
  const abaOrigem = planilhaOrigem.getSheetByName(ABA_ORIGEM);

  // Obter a última linha preenchida na aba de origem
  const ultimaLinha = abaOrigem.getLastRow();

  // Obter valores das colunas necessárias (I, A, P)
  const valorCelulaI = abaOrigem.getRange('I' + ultimaLinha).getValue(); // Nome dos ingressantes (coluna I)
  const valorCelulaA = abaOrigem.getRange('A' + ultimaLinha).getValue(); // Data de ingresso (coluna A)
  const valorCelulaP = abaOrigem.getRange('P' + ultimaLinha).getValue(); // Contato do ingressante (coluna P)

  // Verificar se o campo de nome (coluna I) está vazio
  if (!valorCelulaI) {
    return; // Interrompe a execução se o nome estiver vazio
  }

  // Separar os nomes dos ingressantes ignorando informações adicionais após o hífen
  const ingressantes = valorCelulaI.split(',').map(ingressante => {
    return ingressante.split(' - ')[0].trim(); // Retorna apenas o nome antes do hífen
  });

  // Obter a planilha de destino e aba de destino
  const planilhaDestino = SpreadsheetApp.openById(PLANILHA_DESTINO_ID);
  const abaDestino = planilhaDestino.getSheetByName(ABA_DESTINO);

  // Formatar a data de execução do script
  const dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  // Inserir dados para cada ingressante identificado
  ingressantes.forEach(nomeIngressante => {
    const ultimaLinhaDestino = abaDestino.getLastRow() + 1; // Próxima linha vazia

    // Preencher as informações na planilha de destino
    abaDestino.getRange('A' + ultimaLinhaDestino).setValue(nomeIngressante); // Nome (coluna A)
    abaDestino.getRange('B' + ultimaLinhaDestino).setValue(valorCelulaA);   // Data de ingresso (coluna B)
    abaDestino.getRange('C' + ultimaLinhaDestino).setValue(valorCelulaP);   // Contato (coluna C)
    abaDestino.getRange('D' + ultimaLinhaDestino).setValue("VERIFICAR");    // Status (coluna D)
    abaDestino.getRange('E' + ultimaLinhaDestino).setValue("robô");         // Servidor (coluna E)
    abaDestino.getRange('F' + ultimaLinhaDestino).setValue(dataAtual);      // Última atualização (coluna F)
    abaDestino.getRange('G' + ultimaLinhaDestino).setValue(`Dados do ingresso inseridos automaticamente pelo robô em ${dataAtual}.`); // Observações (coluna G)
  });
}
