/**
 * Função que varre a planilha e consulta rastreamento na API Linketrack.
 * Após a atualização, envia mensagens via WhatsApp para o grupo de serviço 
 * e também para o número de contato de WhatsApp do ingressante.
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
        // Formata os eventos rastreio
        var eventosRastreamento = formatarEventos(resultado.eventos);
        
        // Verifica o conteúdo atual da coluna J (Eventos do Rastreamento)
        var eventosAtuais = sheet.getRange(i, 10).getValue(); // Coluna J
        
        // Se houver alteração nos eventos, atualiza a planilha e envia as mensagens
        if (eventosAtuais !== eventosRastreamento) {
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
          
          // Atualiza a coluna G (Observações) com a data atual, evitando duplicação
          var observacoes = sheet.getRange(i, 7).getValue(); // Coluna G (Observações)
          if (resultado.eventos[0].status.includes("entregue")) {
            sheet.getRange(i, 7).setValue(observacoes + " Confirmado recebimento pelo robô de rastreamento em " + dataAtual + ".");
            // Se o último evento for de entrega, altera o status para "RECEBIDO"
            sheet.getRange(i, 4).setValue("RECEBIDO"); // Coluna D (Status de envio)
          } else {
            // Verificar se já termina com "Atualizado pelo robô em" para evitar repetição
            if (observacoes.endsWith("Atualizado pelo robô em " + dataAtual + ".")) {
              observacoes = observacoes.replace(/(\d{2}\/\d{2}\/\d{2})$/, dataAtual);
            } else {
              observacoes += " Atualizado pelo robô em " + dataAtual + ".";
            }
            sheet.getRange(i, 7).setValue(observacoes);
          }

          // Obtém informações para a mensagem
          var nome = sheet.getRange(i, 1).getValue();          // Coluna A (Nome)
          var ingressoData = sheet.getRange(i, 2).getValue();  // Coluna B (Data Ingresso)
          var ingressoFormatado = formatarData(new Date(ingressoData)); // Formata a data de ingresso
          var contato = sheet.getRange(i, 3).getValue();       // Coluna C (Telefone WhatsApp)
          var endereco = sheet.getRange(i, 8).getValue();      // Coluna H (Endereço Completo)
          
          // Substitui ";" por "\n" para formatação de mensagem
          var eventosMensagem = eventosRastreamento.replace(/;/g, "\n");

          // Monta a mensagem do WhatsApp 
          var mensagem = "*GRUPO NOVA ERA ONLINE*\n\n*Nome:* " + nome  + "\n*Contato:* " + contato + "\n*Ingresso:* " + ingressoFormatado + "\n*Código Rastreio:* " + codigoRastreio + "\n*Rastreamento:* \n" + eventosMensagem + "\n*Endereço:* \n" + endereco + "\n\n*Acesse o site para rastrear:* https://rastreamento.correios.com.br";

          // Envia mensagem para o grupo de serviço
          enviarMensagemWhatsApp(mensagem, "5511970861705-1613854533@g.us");
          
          // Envia mensagem para o ingressante (coluna C)
          enviarMensagemWhatsApp(mensagem, contato);

          // Aguarda 3 segundos entre as consultas (respeitar o delay) da API
          Utilities.sleep(3000);
        }
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
  var url = 'https://api.linketrack.com/track/json?user=gruponovaera@email.com&token=12345678902582782853589234572843578420951234567890&codigo=' + codigoRastreio;
  // Coloque sua credencial da API linketrack user e token
  // Envie e-mail para api@linketrack.com solicitando credenciais
  
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
 * @return {string} - String formatada com os eventos do rastreamento separados por ;
 */
function formatarEventos(eventos) {
  var eventosFormatados = eventos.map(function(evento) {
    return evento.data + " " + evento.hora + " " + evento.local + " " + evento.status;
  }).reverse(); // Inverte a ordem dos eventos
  return eventosFormatados.join(";"); // Adiciona ; entre os eventos para salvar na planilha
}

/**
 * Função que formata a data no formato "DiaSemana DD/MM/AAAA"
 * @param {Date} data - Objeto Date a ser formatado
 * @return {string} - Data formatada com o dia da semana
 */
function formatarData(data) {
  var diasDaSemana = ["Domingo", "Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"];
  
  var diaSemana = diasDaSemana[data.getDay()];
  var dia = ("0" + data.getDate()).slice(-2);
  var mes = ("0" + (data.getMonth() + 1)).slice(-2);
  var ano = data.getFullYear();
  
  return diaSemana + " " + dia + "/" + mes + "/" + ano;
}

/**
 * Função que envia mensagem via WhatsApp usando a Evolution API
 * @param {string} mensagem - O texto da mensagem que será enviada
 * @param {string} numero - O número de WhatsApp ou ID do grupo para enviar a mensagem
 */
function enviarMensagemWhatsApp(mensagem, numero) {
  var apiUrl = "https://{url.api}/message/sendText/instance"; // Substituir pela URL da EvolutionAPI
  var apiKey = "Token"; // Substituir pelo seu Token da API

  var payload = {
    "number": numero, // Número ou ID do grupo
    "text": mensagem   // Texto da mensagem a ser enviada
  };  

  var options = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "apikey": apiKey,
    },
    payload: JSON.stringify(payload),
  };

  try {
    var response = UrlFetchApp.fetch(apiUrl, options);
    var responseData = JSON.parse(response.getContentText());

    // Verifica se a mensagem foi enviada com sucesso
    if (responseData.success) {
      Logger.log("Mensagem enviada com sucesso para o WhatsApp via API.");
    } else {
      Logger.log("Erro ao enviar a mensagem para o WhatsApp via API: " + responseData.error);
    }
  } catch (e) {
    Logger.log("Erro ao enviar mensagem para o número: " + numero + ". Erro: " + e.message);
  }
}