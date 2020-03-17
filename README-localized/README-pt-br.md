---
page_type: sample
products:
- office-outlook
- office-365
languages:
- html
- javascript
- ruby
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 8/11/2015 3:48:02 PM
---
# Outlook-Add-in-JavaScript-ViewYouTubeVideos

## Resumo
Esse suplemento do Outlook permite aos usuários exibir vídeos do YouTube no painel de suplemento do Outlook se a mensagem de e-mail ou compromisso selecionado contiver uma URL para um vídeo no YouTube. Ele também contém um script de configuração que implanta o suplemento em um servidor Web Ruby executando em um Mac. A figura a seguir é uma captura de tela do suplemento do YouTube ativado para uma mensagem no painel de leitura.
<br />
<br />
![Suplemento do Outlook executando um vídeo do YouTube no item de e-mail](/static/pic1.png)

## Pré-requisitos
* Mac OS X 10.10 ou posterior
* Bash
* Ruby 2.2.x+
* [Bundler](http://bundler.io/v1.5/gemfile.html)
* OpenSSL
* Um computador executando o Exchange 2013 com pelo menos uma conta de email ou uma conta do Office 365. Você pode [participar do Programa de Desenvolvedores do Office 365 e obter uma assinatura gratuita de 1 ano do Office 365](https://aka.ms/devprogramsignup).
* Qualquer navegador compatível com ECMAScript 5.1, HTML5 e CSS3, como o Chrome, o Firefox ou o Safari.
* Outlook 2016 para Mac

## Componentes principais do exemplo
* [```LICENSE.txt```](LICENSE.txt) os termos e condições de uso desse distribuídor
* [```config.ru```](config.ru) configuração do rack
* [```setup.sh```](setup.sh) Configure o script para gerar ```app. rb```, ```manifest. xml```e opcionalmente, um certificado
* [```cert/ss_certgen.sh```](cert/ss_certgen.sh) script de geração de certificado autoassinado
* [```public/res/js/strings_en-us.js```](public/res/js/strings_en-us.js) localização em inglês
* [```pública/res/js/strings_fr-fr. js```](public/res/js/strings_fr-fr.js) localização em francês

## Descrição do código

Os principais arquivos de código para esse suplemento são ```manifest. xml``` e ```YouTube. html```, juntamente com a biblioteca JavaScript e arquivos de cadeia de caracteres para suplementos do Office e um arquivo de imagem de logotipo. A seguir, um resumo de alto nível de como funciona o suplemento:

Esse suplemento de e-mail especifica no arquivo ```manifest.xml``` que requer um aplicativo host compatível com o recurso de caixa de correio:

```xml
<Capabilities>
    <Capability Name="Mailbox"/>
</Capabilities>
```

```xml
<DesktopSettings>
    <!-- Change the following line to specify the web server where the HTML file is hosted. -->
    <SourceLocation DefaultValue="https://webserver/YouTube/YouTube.htm"/>
    <RequestedHeight>216</RequestedHeight>
</DesktopSettings>
<TabletSettings>
    <!-- Change the following line to specify the web server where the HTML file is hosted. -->
    <SourceLocation DefaultValue="https://webserver/YouTube/YouTube.htm"/>
    <RequestedHeight>216</RequestedHeight>
</TabletSettings>
```
    
O suplemento também solicita a permissão ReadItem no arquivo de manifesto para que ele possa executar as expressões regulares, explicadas abaixo.

```xml
    <Permissions>ReadItem</Permissions>
```
    
O aplicativo host ativa esse suplemento quando a mensagem ou o compromisso selecionado contém uma URL para um vídeo do YouTube. Isso é feito primeiro ao ler o arquivo manifest.xml, que especifica uma regra de ativação que inclui uma expressão regular para procurar tal URL:

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" PropertyName="BodyAsPlaintext" RegExName="VideoURL" RegExValue="http://(((www\.)?youtube\.com/watch\?v=)|(youtu\.be/))[a-zA-Z0-9_-]{11}"/>
```
    
O suplemento define uma função Initialize que é um manipulador de eventos para o evento Initialize. Quando o ambiente de tempo de execução for carregado, o evento Initialize será disparado e a função Initialize chamará a função main do suplemento, `init`, conforme mostrado no código a seguir:

```javascript
Office.initialize = function () {
    init(Office.context.mailbox.item.getRegExMatches().VideoURL);
}
```

O método ```getRegExMatches``` do item selecionado retorna uma matriz de cadeias de caracteres que correspondem à expressão regular ```VideoURL```, especificada no arquivo ```manifest. xml```. Nesse caso, essa matriz contém URLs para vídeos no YouTube.

A função `init` assume como um parâmetro de entrada a matriz de URLs do YouTube e cria dinamicamente o HTML para exibir a miniatura e os detalhes correspondentes de cada vídeo.

Esse HTML dinamicamente criado exibe o primeiro vídeo em um player do YouTube Embedded, juntamente com detalhes sobre o vídeo. O painel de suplemento também exibe as miniaturas de todos os vídeos subsequentes. O usuário final pode escolher uma miniatura para exibir qualquer um dos vídeos no YouTube Embedded Player, sem sair do aplicativo host.

## Configurar
Este arquivo de exemplo é um ```setup.sh```- esse arquivo de configuração faz o seguinte:
* Verifica e instala [Dependências](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L23)
* [Gera```manifest.xml```](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L37)
* [Gera ```app.rb``` ](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L44)
* [Opcionalmente](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L34) gera um [certificado auto-assinado](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/cert/ss_certgen.sh#L49)

Para executar o script, digite no terminal compatível com POSIX:

    $ bash setup.sh
    
## Iniciar o servidor
Na raiz do projeto, execute:

    $ rackup

## Confiar em seu certificado auto-assinado
Como este exemplo usa um servidor local e [certificado autoassinado](https://en.wikipedia.org/wiki/Self-signed_certificate), primeiro é necessário estabelecer a "confiança" entre o localhost e o certificado autoassinado. Antes que o Outlook transmita dados potencialmente confidenciais para qualquer suplemento, seu Certificado SSL é verificado como confiável. Esse requisito ajuda a proteger a privacidade dos seus dados. Qualquer navegador moderno avisará o usuário sobre discrepâncias de certificados e vários deles também fornecerão um mecanismo para inspecionar e estabelecer a confiança. Depois de iniciar seu servidor local, abra seu navegador da Web preferido e navegue até a URL hospedada local especificada no seu arquivo manifest.xml. (Por padrão, o script setup.sh neste exemplo especifica essa URL como ```https://0.0.0.0:8443/youtube.html```.) Neste ponto, você poderá receber um aviso de certificado. Você precisa confiar neste certificado.

Abra o Safari |
:-:|
![Diloag de segurança do Safari para validar o certificado](/static/show_cert.png)|

Selecione "sempre confiar" no seu certificado auto-assinado |
:-:|
![Diloag de segurança do Safari para sempre confiar no certificado da Contoso](/static/add_trust.png)|

## Instale o suplemento do Office 365
A instalação deste suplemento de exemplo exige acesso ao Outlook na Web. A instalação pode ser realizada a partir de Configurações > Gerenciar aplicativos.

Selecione o menu "Configurações" e "Gerenciar aplicativos" | Instalar a partir do arquivo
:-:|:-:
![Lista suspensa para gerenciar aplicativos](/static/menu_loc.png)|![adicionar de uma página de propriedades de configurações de arquivo](/static/menu_opt.png)

Selecione o arquivo manifest.xml |
:-:|
![adicionar de uma página de propriedades do arquivo, configurando o nome do manifesto](/static/menu_chooser.png)|

Selecione "instalar" e, em seguida, continuar?
:-:|
![adicione de uma página de propriedades do arquivo, confirme a configuração para adicionar](/static/menu_warn.png)|

## Veja isso em ação
Para demonstrar a funcionalidade do suplemento, você precisará usar o cliente nativo do Office Outlook 2013.
* Abra o Outlook 2013
* Envie um link para um vídeo do YouTube por e-mail - Precisa de [ uma sugestão?](http://www.youtube.com/watch?v=oEx5lmbCKtY)
* Expanda o painel de suplemento para ver uma visualização

## Perguntas e comentários
* Se você tiver problemas para executar este exemplo, [relate um problema](https://github.com/OfficeDev/https://github.com/OfficeDev/Outlook-Add-in-Javascript-ViewYouTubeVideos/issues).
* Em geral, perguntas sobre o desenvolvimento de Suplementos do Office devem ser postadas no [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Não deixe de marcar as perguntas ou comentários com [office-addins].

## Recursos adicionais
* [Mais exemplos de Suplementos](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [Suplementos do Outlook](https://dev.office.com/code-samples#?filters=web,outlook)
* [API do Outlook](https://dev.outlook.com/)
* [Exemplos de código do Outlook no Centro de Desenvolvimento do Office](https://dev.office.com/code-samples#?filters=web,outlook)
* [Últimas novidades – centro de desenvolvimento do Office](http://dev.office.com/latestnews)
* [Treinamento – centro de desenvolvimento do Office](https://dev.office.com/training)

## Direitos autorais
Copyright © 2015 Microsoft. Todos os direitos reservados.


Este projeto adotou o [Código de Conduta do Código Aberto da Microsoft](https://opensource.microsoft.com/codeofconduct/). Para saber mais, confira [Perguntas frequentes sobre o Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou contate [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.
