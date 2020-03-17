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

## Resumen
Este complemento de Outlook permite a los usuarios ver vídeos de YouTube en el panel de complementos de Outlook si el mensaje de correo electrónico o la cita seleccionados contienen una URL a un vídeo en YouTube. También contiene un script de configuración que implementa el complemento en un servidor web de Ruby que se ejecute en un equipo Mac. La siguiente figura es una captura de pantalla del complemento de YouTube que se ha activado para un mensaje en el panel de lectura.
<br />
<br />
![Complemento de Outlook que ejecuta un vídeo de YouTube en el elemento de correo](/static/pic1.png)

## Requisitos previos
* Mac OS X 10.10 o versiones posteriores
* Bash
* Ruby 2.2.x+
* [Bundler](http://bundler.io/v1.5/gemfile.html)
* OpenSSL
* Un equipo que ejecute Exchange 2013 y, como mínimo, una cuenta de correo electrónico o una cuenta de Office 365. Puede [participar en el programa para desarrolladores Office 365 y obtener una suscripción gratuita durante 1 año a Office 365](https://aka.ms/devprogramsignup).
* Cualquier explorador que admita ECMAScript 5.1, HTML5 y CSS3, como Chrome, Firefox o Safari.
* Outlook 2016 para Mac

## Componentes clave del ejemplo
* [```LICENSE.txt```](LICENSE.txt) Los términos y condiciones de uso de este distribuible
* [```config.ru```](config.ru) Rack config
* [```setup.sh```](setup.sh) Script de configuración para generar ```app.rb```, ```manifest.xml```y, opcionalmente, un certificado
* [```cert/ss_certgen.sh```](cert/ss_certgen.sh) script generador de certificados autofirmados
* [```public/res/js/strings_en-us.js```](public/res/js/strings_en-us.js) Localización en inglés de Estados Unidos
* [```public/res/js/strings_fr-fr.js```](public/res/js/strings_fr-fr.js) Localización en francés

## Descripción del código

Los archivos de código principales para este complemento son ```manifest.xml``` y ```youtube.html```, junto con la biblioteca de JavaScript y los archivos de cadenas para complementos de Office y un archivo de imagen del logotipo. A continuación se muestra un resumen general del funcionamiento del complemento:

Este complemento de correo electrónico especifica en el archivo```manifest.xml``` que necesita una aplicación host que admita la funcionalidad mailbox:

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
    
El complemento también requiere el permiso ReadItem en el archivo de manifiesto para poder ejecutar expresiones regulares, lo que se explica a continuación.

```xml
    <Permissions>ReadItem</Permissions>
```
    
La aplicación host activa este complemento cuando el mensaje o la cita seleccionados contienen una URL a un vídeo de YouTube. Para ello, primero lee al inicio el archivo manifest.xml, que especifica una regla de activación que incluye una expresión regular para buscar esa URL:

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" PropertyName="BodyAsPlaintext" RegExName="VideoURL" RegExValue="http://(((www\.)?youtube\.com/watch\?v=)|(youtu\.be/))[a-zA-Z0-9_-]{11}"/>
```
    
El complemento define una función Initialize que es un controlador de eventos del evento Initialize. Cuando se carga el entorno de tiempo de ejecución, se activa el evento Initialize y, después, la función Initialize llama a la función principal del complemento, `init`, tal como se muestra en el siguiente código:

```javascript
Office.initialize = function () {
    init(Office.context.mailbox.item.getRegExMatches().VideoURL);
}
```

El método ```getRegExMatches``` del elemento seleccionado devuelve una matriz de cadenas que coinciden con la expresión regular ```VideoURL```, que se especifica en el archivo ```manifest.xml```. En este caso, la matriz contiene las direcciones URL a los vídeos de YouTube.

La función `init` toma como parámetro de entrada esa matriz de direcciones URL de YouTube y crea dinámicamente el código HTML para mostrar las miniaturas y los detalles correspondientes de cada vídeo.

Este código HTML generado dinámicamente muestra el primer vídeo en un reproductor incrustado de YouTube, junto con los detalles del vídeo. En el panel de complementos también se muestran las miniaturas de los vídeos posteriores. El usuario final puede elegir una miniatura para ver cualquiera de los vídeos de YouTube en el reproductor incrustado sin salir de la aplicación host.

## Configurar
Con este ejemplo se proporciona un archivo de configuración ```setup.sh``` que hace lo siguiente:
* Comprueba e instala las [dependencias](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L23)
* [Genera ```manifest.xml```](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L37)
* [Genera ```app.rb```](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L44)
* [Opcionalmente](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L34), genera un certificado [autofirmado](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/cert/ss_certgen.sh#L49)

Para ejecutar el script, escriba en el terminal POSIX-compliant:

    $ bash setup.sh
    
## Iniciar el servidor
Desde la raíz del proyecto, ejecute:

    $ rackup

## Confíe en el certificado autofirmado
Como en este ejemplo se usa un servidor local y un [certificado autofirmado](https://en.wikipedia.org/wiki/Self-signed_certificate), primero debe establecer la confianza entre el host local y el certificado autofirmado. Antes de que Outlook transmita datos que puedan ser confidenciales a cualquier complemento, se comprueba que su certificado SSL sea de confianza. Este requisito ayuda a proteger la privacidad de los datos. Cualquier explorador web moderno avisará al usuario si detecta discrepancias en los certificados y, además, muchos de ellos proporcionan un mecanismo para inspeccionar y establecer la confianza. Después de iniciar el servidor local, abra el explorador web que prefiera y vaya a la URL hospedada localmente que se especifica en el archivo manifest.xml. (De forma predeterminada, el script setup.sh de este ejemplo especifica esa URL como ```https://0.0.0.0:8443/youtube.html```). En este momento, es posible que se muestre una advertencia de certificado. Debe confiar en el certificado.

Abrir Safari|
:-:|
![Diálogo de seguridad de Safari para validar el certificado](/static/show_cert.png)|

Seleccione “Confiar siempre” en el certificado autofirmado|
:-:|
![Diálogo de seguridad de Safari para confiar siempre en el certificado de Contoso](/static/add_trust.png)|

## Instalar el complemento en Office 365
Para instalar este complemento de ejemplo es necesario tener acceso a Outlook en la Web. La instalación se puede hacer desde Configuración > Administrar aplicaciones.

Seleccione el menú “Configuración” y “Administrar aplicaciones”|Instalar desde archivo
:-:|:-:
![Lista desplegable de configuración para administrar aplicaciones](/static/menu_loc.png)|![Página de propiedad de configuración Agregar desde un archivo](/static/menu_opt.png)

Seleccione el archivo manifest.xml|
:-:|
![Página de propiedades Agregar desde un archivo que establece el nombre del manifiesto](/static/menu_chooser.png)|

Seleccione “Instalar” y después “Continuar”|
:-:|
![Página de propiedades Agregar desde un archivo que establece la confirmación para agregar](/static/menu_warn.png)|

## Ver el funcionamiento
Para demostrar la funcionalidad del complemento, tendrá que usar el cliente nativo de Office Outlook 2013.
* Abra Outlook 2013
* Envíese por correo electrónico un vínculo a un vídeo de YouTube. ¿Necesita alguna [sugerencia?](http://www.youtube.com/watch?v=oEx5lmbCKtY)
* Expanda el panel del complemento para ver una vista previa

## Preguntas y comentarios
* Si tiene algún problema para ejecutar este ejemplo, [registre un problema](https://github.com/OfficeDev/https://github.com/OfficeDev/Outlook-Add-in-Javascript-ViewYouTubeVideos/issues).
* Las preguntas sobre el desarrollo de complementos para Office en general deben enviarse a [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Asegúrese de que sus preguntas o comentarios se etiquetan con [office-addins].

## Recursos adicionales
* [Más complementos de ejemplo](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [Complementos de Outlook](https://dev.office.com/code-samples#?filters=web,outlook)
* [API de Outlook](https://dev.outlook.com/)
* [Ejemplos de código: Centro para desarrolladores de Office](https://dev.office.com/code-samples#?filters=web,outlook)
* [Últimas noticias: Centro para desarrolladores de Office](http://dev.office.com/latestnews)
* [Aprendizaje: Centro para desarrolladores de Office](https://dev.office.com/training)

## Derechos de autor
Copyright (c) 2015 Microsoft. Todos los derechos reservados.


Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, vea [Preguntas frecuentes sobre el código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.
