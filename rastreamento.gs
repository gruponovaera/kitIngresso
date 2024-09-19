/**
 * Função que varre a planilha e consulta rastreamento na API Linketrack.
 */
function atualizarRastreamento() {
  // Abre a planilha / aba
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kit Ingresso");
  var lastRow = sheet.getLastRow(); // Última linha da planilha
  
  // Varre todas as linhas da planilha, começando da segunda
  for (var i = 2; i <= lastRow; i++) {
    var codigoRastreio = sheet.getRange(i, 9).getValue(); // Coluna I (Código de rastreio)
    var statusEnvio = sheet.getRange(i, 4).getValue();    // Coluna D (Status de envio)

    // Verifica se o código de rastreio não está vazio e se o status de envio não é "RECEBIDO"
    if (codigoRastreio && statusEnvio !== "RECEBIDO") {
      
      // Consulta API com o código de rastreio
      var resultado = consultarRastreamento(codigoRastreio);
      
      if (resultado && resultado.eventos && resultado.eventos.length > 0) {
        // Atualiza as informações
        var eventosRastreamento = formatarEventos(resultado.eventos); // Formata eventos rastreio para a coluna J
        var dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yy");
        
        // Atualiza a coluna J com os eventos rastreados
        sheet.getRange(i, 10).setValue(eventosRastreamento); // Coluna J (Eventos do rastreamento)

        // Atualiza a coluna F com a data da última atualização
        sheet.getRange(i, 6).setValue(dataAtual); // Coluna F (Última atualização)
        
        // Atualiza a coluna E adicionando " e robô"
        var servidores = sheet.getRange(i, 5).getValue(); // Coluna E (Servidores)
        if (!servidores.includes("robô")) {
          sheet.getRange(i, 5).setValue(servidores + " e robô");
        }
        
        // Atualiza a coluna G (Observações) com a data atual
        var observacoes = sheet.getRange(i, 7).getValue(); // Coluna G (Observações)
        if (resultado.eventos[0].status.includes("entregue")) {
          sheet.getRange(i, 7).setValue(observacoes + " Confirmado recebimento pelo robô de rastreamento em " + dataAtual + ".");
          // Se o último evento for de entrega, altera o status para "RECEBIDO"
          sheet.getRange(i, 4).setValue("RECEBIDO"); // Coluna D (Status de envio)
        } else {
          sheet.getRange(i, 7).setValue(observacoes + " Atualizado pelo robô em " + dataAtual + ".");
        }

        // Aguarda 3 segundos entre as consultas (respeita o delay) da API
        Utilities.sleep(3000);
      }
    }
  }
}

/**
 * Função que consulta a API de rastreamento Linketrack
 * @param {string} codigoRastreio - O código de rastreamento a ser consultado
 * @return {Object} - Retorno da API de rastreamento com eventos e status
 */
function consultarRastreamento(codigoRastreio) {
  var url = 'https://api.linketrack.com/track/json?user=gruponovaera&token=12345678902582782853589234572843578420951234567890&codigo=' + codigoRastreio;
  // Coloque sua credencial da API linketrack user e token

  try {
    // Faz a requisição GET na API
    var response = UrlFetchApp.fetch(url);
    var resultado = JSON.parse(response.getContentText());
    return resultado;
  } catch (e) {
    Logger.log("Erro na consulta do código de rastreamento: " + codigoRastreio + ". Erro: " + e.message);
    return null;
  }
}

/**
 * Função que formata os eventos retornados pela API para salvar na planilha
 * @param {Array} eventos - Array de eventos do objeto rastreado
 * @return {string} - String formatada com os eventos do rastreamento
 */
function formatarEventos(eventos) {
  var eventosFormatados = eventos.map(function(evento) {
    return evento.data + " " + evento.hora + " " + evento.local + " " + evento.status;
  });
  return eventosFormatados.join("; ");
}
