# Automatização do WhatsApp a partir do Excel

## Sumário
1. [Funcionalidades](#funcionalidades)
2. [Uso](#uso)

## Funcionalidades
O código de automação do WhatsApp a partir do Excel tem como objetivo contornar a necessidade imposta pelo WhatsApp na utilização da API (paga e complexa). Utilizei a biblioteca Selenium, juntamente com outros artifícios, para acessar a planilha do Excel.

- **Listar números com nomes**:
  
  ![Lista de Números com Nomes](https://github.com/gustavolace/Whatsapp-Automation-from-excel/assets/99536403/0cbb1d48-8fe8-495c-9967-53696ffade7a)

- **Enviar mensagens perfeitamente formatadas, incluindo quebra de linha**:
  
  Utilizei a palavra reservada `$cl` para substituir o nome do cliente no texto.

  ![Mensagem Formatada](https://github.com/gustavolace/Whatsapp-Automation-from-excel/assets/99536403/e751e543-82c6-4923-99d1-20d0472b570b)

- **Enviar imagens, junto com o texto**:
  
  Basta colocar a imagem sobre as células.

  ![Envio de Imagens](https://github.com/gustavolace/Whatsapp-Automation-from-excel/assets/99536403/d647be5c-1c46-4a04-b6c1-e17296b3a55e)

- **Planilha Completa**:
  
  ![Planilha Completa](https://github.com/gustavolace/Whatsapp-Automation-from-excel/assets/99536403/77730f35-fcbf-46f5-bd0b-1510af6a1210)
  ![image](https://github.com/gustavolace/Whatsapp-Automation-from-excel/assets/99536403/a62553a2-017d-441a-a715-1fd38dc71471)

## Uso
Antes de mais nada, se atente aos arquivos compactados, seria impossível enviar o arquivo inteiro ao GitHub, extraia ambos os main(s).

Selecione todos e aperte para extrair aqui <br>
Mova o arquivo extraído para a pasta raiz

![image](https://github.com/gustavolace/Excel-Whatsapp-MSG-Sender/assets/99536403/f292a377-95f5-40fd-adb3-4222a9ecca27)
![image](https://github.com/gustavolace/Excel-Whatsapp-MSG-Sender/assets/99536403/eacd20ec-0127-43c0-afb4-8bf78bb6f245)



O código conta com uma série de ações para adquirir as informações do Excel e após isso enviá-las pelo navegador. Para iniciar sua configuração, abra o arquivo `settings.exe`. Este abrirá o navegador; escaneie o código em até 30 segundos e, após isso, feche normalmente. Está tudo pronto para usar.

Aviso: os executáveis devem estar na mesma pasta da planilha para o funcionamento. Não faça alterações na estrutura geral da planilha, exceto a adição ou remoção de imagens, textos ou contatos.

O código contém dois executáveis: o `headless` executa o navegador offscreen, enquanto o `debug` mostra as mensagens. Por padrão, o ativado é o `headless`.

![image](https://github.com/gustavolace/Excel-Whatsapp-MSG-Sender/assets/99536403/d3e52f50-cd80-4957-8e7e-2bbfa389c93f)

Para fazer a alteração, entre no Excel em modo de administração, clique com o botão direito do mouse no botão "enviar", atribuir macro, `ExecutarExe`, editar.

![image](https://github.com/gustavolace/Excel-Whatsapp-MSG-Sender/assets/99536403/0f99dbd7-de31-4768-83b7-cb16fac9884a)

Faça a alteração necessária, substituindo `_headless` por `_debug`.

![image](https://github.com/gustavolace/Excel-Whatsapp-MSG-Sender/assets/99536403/1fb89815-814e-4f20-ab11-292e5bafb8e2)

Tendo isso em mente, apenas preencha os contatos e nomes na planilha. Certifique-se de botar um nome, mesmo que isso não seja relevante ao enviar a mensagem. Para possíveis erros, reinicie o computador ou entre em contato. Se estiver com medo de que as mensagens sejam enviadas de maneira indevida ou para pessoa errada, o máximo que pode acontecer é a mensagem de fato não ser enviada, por problemas de memória ou internet. Note que o máximo que o código espera antes que deixe de enviar a mensagem é 40 segundos.
