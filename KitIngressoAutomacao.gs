/*******************************
 * Script: KitIngressoAutomacao.gs
 * 
 * Descrição: Automatiza processos na planilha "Kit Ingresso" 
 * com base em eventos de edição e rastreamento.
 *******************************/

/** 
 * Configuração das Mensagens
 * 
 * Defina aqui o nome do grupo 
 * que será exibido nas mensagens. 
 */

// Nome do Grupo
const NOME_DO_GRUPO = "NOVA ERA ONLINE"; // Substitua pelo nome do seu grupo

// IDs dos grupos de WhatsApp
const GRUPO_SERVICO = "SEU_ID_DO_GRUPO_DE_SERVICO@g.us"; // Substitua pelo ID do seu grupo de serviço
const GRUPO_KIT_INGRESSO = "SEU_ID_DO_GRUPO_KIT_INGRESSO@g.us"; // Substitua pelo ID do seu grupo de Kit Ingresso

/** 
 * Configuração das APIs
 * 
 * Armazene as URLs e tokens das APIs utilizadas.
 * Recomenda-se utilizar o Properties Service para armazenar credenciais de forma segura.
 */
const WHATSAPP_API_URL = "https://api.evolution.api/message/sendText/INSTANCIA"; // Substitua pela URL da sua API WhatsApp
const WHATSAPP_API_KEY = "seu_api_key_whatsapp"; // Substitua pelo seu API Key WhatsApp

const LINKETRACK_USER = "seu_usuario_linketrack"; // Substitua pelo seu usuário Linketrack
const LINKETRACK_TOKEN = "seu_token_linketrack"; // Substitua pelo seu token Linketrack

/** 
 * Índices das colunas na planilha (A=1, B=2, ...)
 * 
 * Atualizados conforme a ordem das colunas na sua planilha.
 */
const COLUNA_NOME = 1;                  // Coluna A: "NOME"
const COLUNA_DATA_INGRESSO = 2;        // Coluna B: "DATA INGRESSO"
const COLUNA_CONTATO = 3;               // Coluna C: "TELEFONE ZAP"
const COLUNA_STATUS = 4;                // Coluna D: "STATUS DO ENVIO"
const COLUNA_SERVIDORES = 5;            // Coluna E: "SERVIDORES"
const COLUNA_DATA_ATUALIZACAO = 6;      // Coluna F: "ÚLTIMA ATUALIZAÇÃO"
const COLUNA_OBSERVACOES = 7;           // Coluna G: "OBSERVAÇÕES"
const COLUNA_CEP = 8;                   // Coluna H: "CEP"
const COLUNA_ENDERECO = 9;              // Coluna I: "ENDEREÇO COMPLETO"
const COLUNA_CODIGO_RASTREIO = 10;      // Coluna J: "CÓDIGO DE RASTREIO"
const COLUNA_EVENTOS_RASTREIO = 11;     // Coluna K: "EVENTOS DO RASTREAMENTO"


/** 
 * Funções Principais 
 */


/**
 * Função acionada pelo gatilho instalável ao editar uma célula.
 * Preenche automaticamente o endereço com base no CEP inserido na coluna H.
 *
 * @param {Object} e - Objeto de evento que contém informações sobre a edição.
 */
function aoEditarCEP(e) {
  const sheet = e.source.getActiveSheet();
  
  // Verifica se a edição ocorreu na aba "Kit Ingresso"
  if (sheet.getName() !== "Kit Ingresso") return;
  
  const row = e.range.getRow();
  const column = e.range.getColumn();
  
  // Verifica se a edição foi na coluna H (CEP) e se há um valor inserido
  if (column === COLUNA_CEP && e.value) {
    const cepInput = e.value.trim();
    const cep = formataCep(cepInput);
    
    if (!cep) {
      sheet.getRange(row, COLUNA_ENDERECO).setValue("CEP inválido: O CEP deve conter exatamente 8 dígitos numéricos.");
      return;
    }
    
    const enderecoFormatado = buscaEndereco(cep);
    
    if (enderecoFormatado) {
      sheet.getRange(row, COLUNA_ENDERECO).setValue(enderecoFormatado);
    } else {
      sheet.getRange(row, COLUNA_ENDERECO).setValue("Endereço não encontrado para o CEP informado.");
    }
  }
}


/**
 * Função acionada pelo gatilho instalável ao editar uma célula.
 * Gerencia as ações baseadas na alteração do status na coluna D.
 *
 * @param {Object} e - Objeto de evento que contém informações sobre a edição.
 */
function aoEditarStatus(e) {
  const sheet = e.source.getActiveSheet();
  
  // Verifica se a edição ocorreu na aba "Kit Ingresso"
  if (sheet.getName() !== "Kit Ingresso") return;
  
  const row = e.range.getRow();
  const column = e.range.getColumn();
  const novoStatus = e.value ? e.value.toUpperCase() : "";
  
  // Verifica se a edição foi na coluna D (STATUS DO ENVIO) e se há um novo status
  if (column === COLUNA_STATUS && novoStatus) {
    switch (novoStatus) {
      case "FALTA ENDEREÇO":
        statusFaltaEndereco(sheet, row);
        break;
      case "FALTA ENVIAR":
        statusFaltaEnviar(sheet, row);
        break;
      case "ENVIADO":
        statusEnviado(sheet, row);
        break;
      case "RECEBIDO":
        statusRecebido(sheet, row);
        break;
      default:
        // Outros status não requerem ação
        break;
    }
  }
}


/**
 * Função chamada quando o status é alterado para "FALTA ENDEREÇO".
 * Envia uma mensagem ao ingressante solicitando o endereço completo.
 *
 * @param {Sheet} sheet - A aba da planilha onde a edição ocorreu.
 * @param {number} row - A linha que foi editada.
 */
function statusFaltaEndereco(sheet, row) {
  const nome = sheet.getRange(row, COLUNA_NOME).getValue();
  const contato = sheet.getRange(row, COLUNA_CONTATO).getValue();
  
  const mensagem = `Olá ${nome},\n\nPor favor, envie-nos seu endereço completo, incluindo CEP e pontos de referência, para que possamos enviar seu kit de ingresso.\n\nObrigado!`;
  
  enviarMensagemWhatsApp(mensagem, contato);
}


/**
 * Função chamada quando o status é alterado para "FALTA ENVIAR".
 * Envia uma mensagem ao ingressante e ao grupo do kit ingresso com os dados.
 *
 * @param {Sheet} sheet - A aba da planilha onde a edição ocorreu.
 * @param {number} row - A linha que foi editada.
 */
function statusFaltaEnviar(sheet, row) {
  const nome = sheet.getRange(row, COLUNA_NOME).getValue();
  const contato = sheet.getRange(row, COLUNA_CONTATO).getValue();
  const ingressoData = sheet.getRange(row, COLUNA_DATA_INGRESSO).getValue();
  const endereco = sheet.getRange(row, COLUNA_ENDERECO).getValue();
  const ingressoFormatado = formatarData(new Date(ingressoData));
  
  const mensagem = `*GRUPO ${NOME_DO_GRUPO}*\n\n*STATUS:* FALTA ENVIAR\n\n*Nome:* ${nome}\n*Contato:* ${contato}\n*Ingresso:* ${ingressoFormatado}\n*Endereço:*\n${endereco}`;
  
  // Envia a mensagem para o grupo do Kit Ingresso
  enviarMensagemWhatsApp(mensagem, GRUPO_KIT_INGRESSO);
  
  // Envia a mensagem para o contato do ingressante
  enviarMensagemWhatsApp(mensagem, contato);
  
  // Atualiza dados na planilha
  atualizarDados(sheet, row);
}


/**
 * Função chamada quando o status é alterado para "ENVIADO".
 * Envia uma mensagem ao ingressante e ao grupo do kit ingresso com o código de rastreio.
 *
 * @param {Sheet} sheet - A aba da planilha onde a edição ocorreu.
 * @param {number} row - A linha que foi editada.
 */
function statusEnviado(sheet, row) {
  const nome = sheet.getRange(row, COLUNA_NOME).getValue();
  const contato = sheet.getRange(row, COLUNA_CONTATO).getValue();
  const ingressoData = sheet.getRange(row, COLUNA_DATA_INGRESSO).getValue();
  const endereco = sheet.getRange(row, COLUNA_ENDERECO).getValue();
  const codigoRastreio = sheet.getRange(row, COLUNA_CODIGO_RASTREIO).getValue();
  const ingressoFormatado = formatarData(new Date(ingressoData));
  
  const mensagem = `*GRUPO ${NOME_DO_GRUPO}*\n\n*STATUS:* ENVIADO\n\n*Nome:* ${nome}\n*Contato:* ${contato}\n*Ingresso:* ${ingressoFormatado}\n*Endereço:*\n${endereco}\n\n*Código de Rastreamento:* ${codigoRastreio}\n*Acesse o site para rastrear:* https://rastreamento.correios.com.br`;
  
  // Envia a mensagem para o grupo do Kit Ingresso
  enviarMensagemWhatsApp(mensagem, GRUPO_KIT_INGRESSO);
  
  // Envia a mensagem para o contato do ingressante
  enviarMensagemWhatsApp(mensagem, contato);
  
  // Atualiza dados na planilha
  atualizarDados(sheet, row);
}


/**
 * Função chamada quando o status é alterado para "RECEBIDO".
 * Envia uma mensagem ao ingressante e ao grupo do kit ingresso confirmando o recebimento.
 *
 * @param {Sheet} sheet - A aba da planilha onde a edição ocorreu.
 * @param {number} row - A linha que foi editada.
 */
function statusRecebido(sheet, row) {
  const nome = sheet.getRange(row, COLUNA_NOME).getValue();
  const contato = sheet.getRange(row, COLUNA_CONTATO).getValue();
  
  const mensagem = `*GRUPO ${NOME_DO_GRUPO}*\n\n*STATUS:* RECEBIDO\n\n*Nome:* ${nome}\n*Contato:* ${contato}\n*Ingresso:* ${ingressoFormatado}\n\n*Código de Rastreamento:* ${codigoRastreio}\n*Acesse o site para rastrear:* https://rastreamento.correios.com.br \n\n*Confirme recebimento:* Por favor, se possível, compartilhe uma foto ou vídeo conosco para partilhar sua experiência com o grupo!`;
  
  // Envia a mensagem para o grupo do Kit Ingresso
  enviarMensagemWhatsApp(mensagem, GRUPO_KIT_INGRESSO);
  
  // Envia a mensagem para o contato do ingressante
  enviarMensagemWhatsApp(mensagem, contato);
  
  // Atualiza dados na planilha
  atualizarDados(sheet, row);
}


/**
 * Função que atualiza dados comuns na planilha após o envio de mensagens.
 *
 * @param {Sheet} sheet - A aba da planilha.
 * @param {number} row - A linha a ser atualizada.
 */
function atualizarDados(sheet, row) {
  const dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  
  // Atualiza a data da última atualização na coluna F
  sheet.getRange(row, COLUNA_DATA_ATUALIZACAO).setValue(dataAtual);
  
  // Atualiza a coluna E adicionando " e robô" se ainda não estiver presente
  let servidores = sheet.getRange(row, COLUNA_SERVIDORES).getValue();
  if (!servidores.toString().toLowerCase().includes("robô")) {
    servidores += " e robô";
    sheet.getRange(row, COLUNA_SERVIDORES).setValue(servidores);
  }
  
  // Atualiza a coluna G (Observações) substituindo a última ocorrência de "Atualizado pelo robô em [data]"
  let observacoes = sheet.getRange(row, COLUNA_OBSERVACOES).getValue();
  const novaObservacao = `Atualizado pelo robô em ${dataAtual}.`;
  
  if (observacoes.includes("Atualizado pelo robô em")) {
    // Remove a última ocorrência e adiciona a nova data
    observacoes = observacoes.replace(/Atualizado pelo robô em \d{2}\/\d{2}\/\d{4}\.$/, novaObservacao);
  } else {
    // Adiciona a nova observação
    observacoes += ` ${novaObservacao}`;
  }
  
  sheet.getRange(row, COLUNA_OBSERVACOES).setValue(observacoes.trim());
}


/**
 * Função que varre a planilha e consulta rastreamento na API Linketrack.
 * Após a atualização, envia mensagens via WhatsApp para o grupo de serviço
 * e também para o número de contato do ingressante.
 */
function atualizarRastreamento() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kit Ingresso");
  if (!sheet) return;
  
  const lastRow = sheet.getLastRow();
  
  for (let i = 2; i <= lastRow; i++) { // Começa na linha 2, assumindo que a linha 1 tem cabeçalhos
    try {
      const codigoRastreio = sheet.getRange(i, COLUNA_CODIGO_RASTREIO).getValue();
      const statusEnvio = sheet.getRange(i, COLUNA_STATUS).getValue();
      
      // Verifica se há um código de rastreamento e se o status não é "RECEBIDO"
      if (codigoRastreio && statusEnvio.toUpperCase() !== "RECEBIDO") {
        const resultado = consultarRastreamento(codigoRastreio);
        
        if (resultado && resultado.eventos && resultado.eventos.length > 0) {
          const eventosRastreamento = formatarEventos(resultado.eventos);
          const eventosAtuais = sheet.getRange(i, COLUNA_EVENTOS_RASTREIO).getValue();
          
          // Verifica se os eventos de rastreamento mudaram
          if (eventosAtuais !== eventosRastreamento) {
            const dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
            
            // Atualiza os eventos de rastreamento na coluna K
            sheet.getRange(i, COLUNA_EVENTOS_RASTREIO).setValue(eventosRastreamento);
            
            // Atualiza a data da última atualização na coluna F
            sheet.getRange(i, COLUNA_DATA_ATUALIZACAO).setValue(dataAtual);
            
            // Atualiza a coluna E adicionando " e robô" se ainda não estiver presente
            let servidores = sheet.getRange(i, COLUNA_SERVIDORES).getValue();
            if (!servidores.toString().toLowerCase().includes("robô")) {
              servidores += " e robô";
              sheet.getRange(i, COLUNA_SERVIDORES).setValue(servidores);
            }
            
            // Atualiza a coluna G (Observações)
            let observacoes = sheet.getRange(i, COLUNA_OBSERVACOES).getValue();
            if (resultado.eventos[0].status.toLowerCase().includes("entregue")) {
              const novaObservacao = `Confirmado recebimento pelo robô de rastreamento em ${dataAtual}.`;
              if (observacoes.includes("Confirmado recebimento pelo robô de rastreamento em")) {
                observacoes = observacoes.replace(/Confirmado recebimento pelo robô de rastreamento em \d{2}\/\d{2}\/\d{4}\.$/, novaObservacao);
              } else {
                observacoes += ` ${novaObservacao}`;
              }
              sheet.getRange(i, COLUNA_OBSERVACOES).setValue(observacoes.trim());
              
              // Atualiza o status para "RECEBIDO"
              sheet.getRange(i, COLUNA_STATUS).setValue("RECEBIDO");
              
              // Chama a função statusRecebido diretamente após alterar o status
              statusRecebido(sheet, i);
            } else {
              const novaObservacao = `Atualizado pelo robô em ${dataAtual}.`;
              if (observacoes.includes("Atualizado pelo robô em")) {
                observacoes = observacoes.replace(/Atualizado pelo robô em \d{2}\/\d{2}\/\d{4}\.$/, novaObservacao);
              } else {
                observacoes += ` ${novaObservacao}`;
              }
              sheet.getRange(i, COLUNA_OBSERVACOES).setValue(observacoes.trim());
            }
            
            // Coleta informações para a mensagem
            const nome = sheet.getRange(i, COLUNA_NOME).getValue();
            const ingressoData = sheet.getRange(i, COLUNA_DATA_INGRESSO).getValue();
            const ingressoFormatado = formatarData(new Date(ingressoData));
            const contato = sheet.getRange(i, COLUNA_CONTATO).getValue();
            const endereco = sheet.getRange(i, COLUNA_ENDERECO).getValue();
            const eventosMensagem = eventosRastreamento.replace(/;/g, "\n");
            
            const mensagem = `*GRUPO ${NOME_DO_GRUPO}*\n\n*Nome:* ${nome}\n*Contato:* ${contato}\n*Ingresso:* ${ingressoFormatado}\n*Código de Rastreamento:* ${codigoRastreio}\n*Rastreamento:*\n${eventosMensagem}\n*Endereço:*\n${endereco}\n\n*Acesse o site para rastrear:* https://rastreamento.correios.com.br`;
            
            // Envia mensagem para o grupo de serviço
            enviarMensagemWhatsApp(mensagem, GRUPO_SERVICO);
            
            // Envia mensagem para o ingressante
            enviarMensagemWhatsApp(mensagem, contato);
            
            // Aguarda 3 segundos para respeitar limites da API
            Utilities.sleep(3000);
          }
        }
      }
    } catch (error) {
      // Erros são ignorados para continuar o processamento das próximas linhas
      // Você pode implementar notificações ou logar erros conforme necessário
    }
  }
}



/** 
 * Funções Auxiliares 
 */

/**
 * Função que consulta a API de rastreamento Linketrack.
 *
 * @param {string} codigoRastreio - O código de rastreamento a ser consultado.
 * @return {Object|null} - Retorno da API de rastreamento com eventos e status ou null em caso de erro.
 */
function consultarRastreamento(codigoRastreio) {
  const url = `https://api.linketrack.com/track/json?user=${LINKETRACK_USER}&token=${LINKETRACK_TOKEN}&codigo=${codigoRastreio}`;
  
  try {
    const response = UrlFetchApp.fetch(url);
    const resultado = JSON.parse(response.getContentText());
    return resultado;
  } catch (e) {
    return null;
  }
}


/**
 * Função que formata os eventos retornados pela API para salvar na planilha.
 *
 * @param {Array} eventos - Array de eventos do objeto rastreado.
 * @return {string} - String formatada com os eventos do rastreamento separados por ponto e vírgula.
 */
function formatarEventos(eventos) {
  const eventosFormatados = eventos.map(evento => {
    return `${evento.data} ${evento.hora} ${evento.local} ${evento.status}`;
  }).reverse(); // Inverte a ordem dos eventos
  
  return eventosFormatados.join(";"); // Adiciona ponto e vírgula entre os eventos
}


/**
 * Função para formatar a data no formato "DiaSemana DD/MM/AAAA".
 *
 * @param {Date} data - Objeto Date a ser formatado.
 * @return {string} - Data formatada com o dia da semana.
 */
function formatarData(data) {
  const diasDaSemana = ["Domingo", "Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"];
  const diaSemana = diasDaSemana[data.getDay()];
  const dia = ("0" + data.getDate()).slice(-2);
  const mes = ("0" + (data.getMonth() + 1)).slice(-2);
  const ano = data.getFullYear();
  return `${diaSemana} ${dia}/${mes}/${ano}`;
}


/**
 * Função que envia uma mensagem via WhatsApp usando a API especificada.
 *
 * @param {string} mensagem - O texto da mensagem que será enviada.
 * @param {string} numero - O número de WhatsApp ou ID do grupo para enviar a mensagem.
 */
function enviarMensagemWhatsApp(mensagem, numero) {
  const payload = {
    "number": numero,
    "text": mensagem
  };
  
  const options = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "apikey": WHATSAPP_API_KEY,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // Permite capturar respostas mesmo com códigos de erro HTTP
  };
  
  try {
    UrlFetchApp.fetch(WHATSAPP_API_URL, options);
  } catch (e) {
    // Erros são ignorados para continuar o processamento das próximas mensagens
    // Você pode implementar notificações ou logar erros conforme necessário
  }
}


/**
 * Função para formatar o CEP garantindo que tenha exatamente 8 dígitos numéricos.
 *
 * @param {string} cep - O CEP inserido pelo usuário.
 * @return {string|null} - Retorna o CEP formatado ou null se inválido.
 */
function formataCep(cep) {
  cep = cep.replace(/\D/g, '');
  return cep.length === 8 ? cep : null;
}


/**
 * Função para buscar o endereço na API ViaCEP.
 *
 * @param {string} cep - O CEP formatado para consulta.
 * @return {string|null} - Retorna o endereço formatado ou null se não encontrado.
 */
function buscaEndereco(cep) {
  try {
    const url = `https://viacep.com.br/ws/${cep}/json`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const status = response.getResponseCode();
    const content = response.getContentText();
    
    if (status === 200) {
      const data = JSON.parse(content);
      if (data.erro) {
        return null;
      }
      return `CEP: ${data.cep}, ${data.logradouro}, Bairro: ${data.bairro}, ${data.localidade}/${data.uf}`;
    } else {
      return null;
    }
  } catch (error) {
    return null;
  }
}