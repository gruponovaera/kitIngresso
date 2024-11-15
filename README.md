# Kit Ingresso Automação

## Descrição
O `KitIngressoAutomacao` é um script automatizado para Google Sheets que gerencia o envio de mensagens via WhatsApp e o rastreamento de encomendas pelo correo para os ingressantes. Interage com APIs externas para enviar mensagens automáticas e atualizar o status de rastreamento, facilitando a comunicação dos servidores com os ingressantes.

## Funcionalidades
- **Atualização Automática de Endereço:** Preenche o endereço completo com base no CEP inserido.
- **Envio de Mensagens via WhatsApp:** Envia mensagens personalizadas para grupos de serviço e para os ingressantes.
- **Rastreamento de Encomendas:** Consulta o status de rastreamento através da API Linketrack e atualiza a planilha.
- **Atualização de Observações:** Mantém um histórico de atualizações realizadas pelo script.

## Instalação
1. **Clonar o Repositório:**
   ```bash
   git clone https://github.com/seu-usuario/KitIngressoAutomacao.git
   ```
   
2. Configure o Script no Google Sheets:
   - Abra a planilha "Kit Ingresso" no Google Sheets.
   - Vá em Extensões > Apps Script.
   - Apague qualquer código existente e cole o conteúdo do script fornecido.
   
3. Configure Credenciais e IDs:
   No script, substitua os seguintes placeholders pelas suas informações reais:
   - `NOME_DO_GRUPO`: Nome do seu grupo.
   - `SEU_ID_DO_GRUPO_DE_SERVICO@g.us`: ID do grupo de serviço no WhatsApp.
   - `SEU_ID_DO_GRUPO_KIT_INGRESSO@g.us`: ID do grupo de Kit Ingresso no WhatsApp.
   - `https://api.evolution.api/message/sendText/INSTANCIA`: URL da sua API WhatsApp.
   - `seu_api_key_whatsapp`: Seu token na API WhatsApp.
   - `seu_usuario_linketrack`: Seu usuário na API Linketrack.
   - `seu_token_linketrack`: Seu token na API Linketrack.
   
4. Configure os Gatilhos (Triggers):
   - No editor de Apps Script, clique no ícone de Relógio na barra lateral esquerda.
   - Crie os seguintes gatilhos:
   
     - Gatilho para `aoEditarCEP`:
       - Função a ser executada: `aoEditarCEP`
       - Tipo de evento de origem: Da planilha
       - Tipo de evento: Ao editar
	   
     - Gatilho para `aoEditarStatus`:
       - Função a ser executada: `aoEditarStatus`
       - Tipo de evento de origem: Da planilha
       - Tipo de evento: Ao editar
	   
     - Gatilho para `atualizarRastreamento`:
       - Função a ser executada: `atualizarRastreamento`
       - Tipo de evento de origem: Baseado em tempo
       - Tipo de gatilho de tempo: Escolha a frequência desejada (por exemplo, a cada hora, a cada dia).
	   
   - Salve e autorize os gatilhos conforme solicitado.

## Uso
- **Atualização de CEP:** Ao inserir ou editar um CEP na coluna H, o endereço será preenchido automaticamente.
- **Gestão de Status:** Alterar o status na coluna D para "FALTA ENDEREÇO", "FALTA ENVIAR", "ENVIADO" ou "RECEBIDO" executa ações correspondentes.
- **Rastreamento de Encomendas:** O gatilho `atualizarRastreamento` verifica periodicamente o status das encomendas e envia mensagens conforme necessário.

## Segurança
Proteja suas Credenciais: Nunca compartilhe tokens e chaves de API publicamente.

## Contribuição
Contribuições são bem-vindas! Abra uma issue ou envie um pull request para melhorias ou correções.

## Licença
Veja o arquivo LICENSE.md
