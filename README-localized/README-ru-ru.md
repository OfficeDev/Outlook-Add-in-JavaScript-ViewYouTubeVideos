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

## Описание
Благодаря этой надстройке Outlook видео с сайта YouTube можно просматривать в области надстроек в Outlook, если в выбранном сообщении электронной почты или встрече содержится URL-адрес видео на сайте YouTube. Кроме того, в этой настройке также есть сценарий настройки, который развертывает приложение на веб-сервере Ruby, запущенном на компьютере с Mac. На приведенном далее рисунке показан снимок экрана c надстройкой YouTube, активированной в сообщения в области чтения.
<br />
<br />
![Просмотр видео с сайта YouTube в надстройке Outlook в почтовом элементе](/static/pic1.png)

## Необходимые компоненты
* Mac OS X 10.10 или более поздняя версия
* Bash
* Ruby 2.2.x+
* [Bundler](http://bundler.io/v1.5/gemfile.html)
* OpenSSL
* Компьютер с запущенным Exchange 2013 с по крайней мере одной учетной записью электронной почты или учетной записью Office 365. Вы можете [присоединиться к Программе разработчика Office 365 и получить бесплатную годовую подписку на Office 365](https://aka.ms/devprogramsignup).
* Любой браузер с поддержкой ECMAScript 5.1, HTML5, и CSS3, например, Chrome, Firefox или Safari.
* Outlook 2016 для Mac

## Ключевые компоненты примера
* [```LICENSE.txt```](LICENSE.txt) Условия использования этого распространяемого приложения
* [```config.ru```](config.ru) Настройка Rack
* [```setup.sh```](setup.sh) Сценарий настройки для создания файлов ```app.rb```, ```manifest.xml``` и сертификата (при необходимости)
* [```CERT/ss_certgen. sh```](cert/ss_certgen.sh) Сценарий создания самозаверяющего сертификата
* [```public/res/js/strings_en-us.js```](public/res/js/strings_en-us.js) Локализация на английский язык (США)
* [```public/res/js/strings_fr-fr.js```](public/res/js/strings_fr-fr.js) Локализация на французский язык

## Описание кода

К основным файлам с кодом для этой надстройки относятся ```manifest.xml``` и ```youtube.html```, а также библиотека JavaScript, файлы со строками для надстроек Office и файл изображения логотипа. Ниже предлагается общее описание работы этой надстройки.

В файле```manifest.xml``` указывается, что для этой надстройки почты требуется ведущее приложение, которое поддерживает возможности почтового ящика

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
    
Кроме того, для этой надстройки в файле manifest также требуется указать разрешение ReadItem для запуска регулярных выражений. Подробнее об этом см. далее.

```xml
    <Permissions>ReadItem</Permissions>
```
    
Ведущее приложение активирует эту надстройку, если в выбранном сообщении или встрече содержится URL-адрес видео на сайте YouTube. Активация происходит благодаря тому, что при запуске сначала считывается файл manifest.xml, включающий правило активации, в котором содержится регулярное выражение для поиска такого URL-адреса:

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" PropertyName="BodyAsPlaintext" RegExName="VideoURL" RegExValue="http://(((www\.)?youtube\.com/watch\?v=)|(youtu\.be/))[a-zA-Z0-9_-]{11}"/>
```
    
Эта надстройка определяет функцию инициализации, которая выступает обработчиком событий для события инициализации. При загрузке среды выполнения запускается событие инициализации, после чего функция инициализации вызывает основную функцию надстройки,`init` в соответствии с приведенным далее кодом.

```javascript
Office.initialize = function () {
    init(Office.context.mailbox.item.getRegExMatches().VideoURL);
}
```

С помощью метода```getRegExMatches``` выбранного элемента создается массив возвращаемых строк, соответствующий регулярному выражению ```VideoURL```, которое содержится в файле```manifest.xml```. В этом конкретном случае в массиве представлены URL-адреса видео на сайте YouTube.

Функция `init` применяет этот массив URL-адресов на сайте YouTube в качестве входного параметра и динамически выстраивает HTML-текст, в котором отображаются соответствующие эскиз и подробности для каждого видео.

Благодаря этому динамически созданному HTML-тексту первое видео и подробное описание видео отображаются во внедренном проигрывателе YouTube. В области надстроек также отображаются эскизы последующих видеороликов при их наличии. Пользователь может выбрать эскиз для просмотра любого видео во внедренном проигрывателе YouTube без выхода из ведущего приложения.

## Настройка
К примеру прилагается файл ```setup.sh``` - Предназначение файла:
* Проверка и установка [зависимостей](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L23)
* [Создание ```manifest.xml```](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L37)
* [Создание ```app.rb```](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L44)
* [При необходимости](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L34) создание [самозаверяющего сертификата](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/cert/ss_certgen.sh#L49)

Для запуска сценария необходимо внести в терминал, соответствующий требованиям POSIX, следующий текст:

    $ bash setup.sh
    
## Запуск сервера
Запустите в корне проекта указанный ниже файл.

    $ rackup

## Добавление сaмозаверяющего сертификата в доверенные
Так как в этом примере используется локальный сервер и [самозаверяющий сертификат](https://en.wikipedia.org/wiki/Self-signed_certificate), нужно сначала установить "доверие" между локальным узлом и самозаверяющим сертификатом. До отправки любых потенциально конфиденциальных данных Outlook в любую надстройку SSL-сертификат этой надстройки проверяется на доверие. Такая проверка способствует защите конфиденциальности данных. Во всех современных браузерах пользователи получают предупреждение о несоответствии сертификатов, и во многих браузерах предлагается механизм проверки и установления доверия. После запуска локального сервера нужно открыть любой интернет-браузер и перейти по локально размещенному URL-адресу, указанному в файле manifest.xml. (В сценарии setup.sh в этом примере такой URL-адрес указан как```https://0.0.0.0:8443/youtube.html```). На этом этапе может появиться предупреждение о сертификате. Вам нужно добавить этот сертификат в доверенные сертификаты.

Откройте Safari|
:-:|
![Диалоговое окно безопасности Safari для проверки сертификата](/static/show_cert.png)|

Установите параметр "Всегда доверять" для своего самозаверяющего сертификата |
:-:|
![Всегда доверять сертификату Contoso в диалоговом окне безопасности Safari](/static/add_trust.png)|

## Установка надстройки в Office 365
Для установки этого примера надстройки требуется доступ к Outlook в Интернете. Установку можно выполнить через Параметры > Управление приложениями.

Выберите в меню пункт "Параметры"и пункт "Управление приложениями"|Установка из файла
:-:|:-:
![Раскрывающийся список параметров для управления приложениями](/static/menu_loc.png)|![добавить из параметров файла страницу свойства](/static/menu_opt.png)

Выберите файл manifest.xml|
:-:|
![добавить из свойств файла страницу с названием manifest](/static/menu_chooser.png)|

Выберите команду "Установить" и нажмите кнопку "Далее" |
:-:|
![добавить из параметров файла страницу с подтверждением на добавление](/static/menu_warn.png)|

## Наглядные примеры
Чтобы проверить функциональность надстройки необходимо использовать собственный клиент Office Outlook 2013.
* Откройте Outlook 2013
* Отправьте себе сообщение электронной почты со ссылкой на видео на сайте YouTube - Нужна [подсказка?](http://www.youtube.com/watch?v=oEx5lmbCKtY)
* Разверните панель надстроек для предварительного просмотра

## Вопросы и комментарии
* Если у вас возникли проблемы с запуском этого примера, [сообщите о неполадке](https://github.com/OfficeDev/https://github.com/OfficeDev/Outlook-Add-in-Javascript-ViewYouTubeVideos/issues).
* Вопросы о разработке надстроек Office в целом следует размещать в [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Обязательно помечайте свои вопросы и комментарии тегом [office-addins].

## Дополнительные ресурсы
* [Дополнительные примеры надстроек](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [Надстройки Outlook](https://dev.office.com/code-samples#?filters=web,outlook)
* [API Outlook](https://dev.outlook.com/)
* [Примеры кодов - Центр разработчиков Office](https://dev.office.com/code-samples#?filters=web,outlook)
* [Последние новости - Центр разработчиков Office](http://dev.office.com/latestnews)
* [Обучение - Центр разработчиков Office](https://dev.office.com/training)

## Авторские права
(c) Корпорация Майкрософт (Microsoft Corporation), 2015. Все права защищены.


Этот проект соответствует [Правилам поведения разработчиков открытого кода Майкрософт](https://opensource.microsoft.com/codeofconduct/). Дополнительные сведения см. в разделе [часто задаваемых вопросов о правилах поведения](https://opensource.microsoft.com/codeofconduct/faq/). Если у вас возникли вопросы или замечания, напишите нам по адресу [opencode@microsoft.com](mailto:opencode@microsoft.com).
