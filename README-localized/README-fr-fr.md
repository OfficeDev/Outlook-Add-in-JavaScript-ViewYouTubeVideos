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

## Résumé
Ce complément Outlook permet aux utilisateurs d’afficher des vidéos YouTube dans le volet de compléments d’Outlook si l’e-mail ou le rendez-vous sélectionné contient une URL vers une vidéo sur YouTube. Il contient également un script de configuration qui déploie le complément sur un serveur web Ruby exécuté sur un Mac. La figure suivante est une capture d’écran du complément YouTube activé pour un message dans le volet de lecture.
<br />
<br />
![Complément Outlook exécutant une vidéo YouTube dans l’élément e-mail](/static/pic1.png)

## Conditions préalables
* Mac OS X 10.10 ou version ultérieure
* Bash
* Ruby 2.2.x+
* [Bundler](http://bundler.io/v1.5/gemfile.html)
* OpenSSL
* Un ordinateur exécutant Exchange 2013 avec au moins un compte de messagerie ou un compte Office 365. Vous pouvez [participer au programme pour les développeurs Office 365 et obtenir un abonnement gratuit d’un an à Office 365](https://aka.ms/devprogramsignup).
* Tout navigateur prenant en charge ECMAScript 5.1, HTML5 et CSS3, tels que Chrome, Firefox ou Safari.
* Outlook 2016 pour Mac

## Composants clés de l’exemple
* [```LICENSE.txt```](LICENSE.txt) Les conditions d’utilisation de cette distribution
* [```config.ru```](config.ru) Configuration du rack
* [```setup.sh```](setup.sh) Script de configuration pour générer ```app.rb```, ```manifest.xml``` et éventuellement un certificat
* [```cert/ss_certgen.sh```](cert/ss_certgen.sh) script de génération de certificat auto-signé
* [```public/res/js/strings_en-us.js```](public/res/js/strings_en-us.js) Localisation en anglais des États-Unis
* [```public/res/js/strings_fr-fr.js```](public/res/js/strings_fr-fr.js) Localisation française

## Description du code

Les fichiers de codes principaux pour ce complément sont ```manifest.xml``` et ```youtube.html```, ainsi que la bibliothèque JavaScript et les fichiers de chaîne pour les compléments Office et un fichier image de logo. Voici un résumé général de la façon dont le complément fonctionne :

Ce complément courrier indique dans le fichier ```manifest.xml``` qu’il requiert une application hôte qui prend en charge la fonctionnalité de boîte aux lettres :

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
    
Le complément demande également l’autorisation ReadItem dans le fichier manifeste afin de pouvoir exécuter des expressions régulières, qui sont expliquées ci-dessous.

```xml
    <Permissions>ReadItem</Permissions>
```
    
L’application hôte active ce complément lorsque le message ou le rendez-vous sélectionné contient une URL vers une vidéo YouTube. Pour ce faire, il commence par lire le fichier manifest.xml, qui spécifie une règle d’activation qui inclut une expression régulière pour rechercher une telle URL :

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" PropertyName="BodyAsPlaintext" RegExName="VideoURL" RegExValue="http://(((www\.)?youtube\.com/watch\?v=)|(youtu\.be/))[a-zA-Z0-9_-]{11}"/>
```
    
Le complément définit une fonction d'initialisation qui est un gestionnaire d’événements pour l’événement d'initialisation. Lorsque l’environnement d’exécution est chargé, l’événement d'initialisation est déclenché, puis la fonction d'initialisation appelle la fonction principale du complément, `init`, comme illustré dans le code ci-dessous :

```javascript
Office.initialize = function () {
    init(Office.context.mailbox.item.getRegExMatches().VideoURL);
}
```

La méthode ```getRegExMatches``` de l’élément sélectionné renvoie une matrice de chaînes qui correspondent à l’expression régulière ```VideoURL```, qui est spécifié dans le fichier ```manifest.xml```. Dans ce cas, ce tableau contient des URL pour les vidéos sur YouTube.

La fonction `init` prend comme paramètre d’entrée des URL YouTube et construit de façon dynamique le code HTML pour afficher la miniature et les détails correspondants pour chaque vidéo.

Ce code HTML créé de façon dynamique affiche la première vidéo dans un lecteur YouTube incorporé, ainsi que des informations sur la vidéo. Le volet complément affiche également les miniatures des vidéos ultérieures. L’utilisateur final peut choisir une miniature pour afficher les vidéos du lecteur YouTube incorporé, sans quitter l’application hôte.

## Configurer
Fourni avec cet exemple est un ```setup.sh``` - ce fichier d’installation effectue les opérations suivantes :
* vérifie et installe les [dépendances](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L23)
* [génère la ```manifest.xml```](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L37)
* [génère le ```app.rb```](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L44)
* [éventuellement](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L34) génère un [certificat auto-signé](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/cert/ss_certgen.sh#L49)

Pour exécuter le script, tapez sur votre terminal compatible POSIX :

    $ bash setup.sh
    
## Lancez le serveur
À partir de la racine du projet, exécutez :

    $ rackup

## Approuver votre certificat auto-signé
Comme cet exemple utilise un serveur local et [certificat auto-signé](https://en.wikipedia.org/wiki/Self-signed_certificate), vous devez d’abord établir une approbation entre votre localhost et le certificat auto-signé. Avant qu'Outlook ne transmette des données potentiellement sensibles à un complément, son certificat SSL est vérifié pour approbation. Cette exigence permet de protéger la confidentialité de vos données. Les navigateurs web modernes avertissent l’utilisateur des divergences de certificats, et proposent un mécanisme permettant d’inspecter et d’établir une approbation. Une fois que vous avez lancé votre serveur local, ouvrez votre navigateur web de votre choix et accédez à l’URL hébergée localement spécifiée dans votre fichier manifest.xml. (Par défaut, le script setup.sh dans cet exemple spécifie cette URL comme ```https://0.0.0.0:8443/youtube.html```.) À ce stade, vous pouvez être confronté à un avertissement de certificat. Vous devez faire confiance à ce certificat.

Ouvrez Safari |
:-:|
![Safari Security diloag pour valider le certificat](/static/show_cert.png)|

Sélectionnez « toujours faire confiance » votre certificat auto-signé |
:-:|
![Safari Security diloag pour toujours faire confiance au certificat contoso](/static/add_trust.png)|

## Installer le complément dans Office 365
L’installation de cet exemple de complément nécessite l’accès à Outlook sur le web. L’installation peut être effectuée à partir des paramètres > Gérer des applications.

Sélectionnez le menu « Paramètres » et « Gérer des applications » | Installation à partir du fichier
:-:|:-:
![Liste déroulante des paramètres pour gérer les applications](/static/menu_loc.png)|![ajouter à partir d’une page de propriétés paramètres de fichier](/static/menu_opt.png)

Sélectionnez le fichier manifest.xml |
:-:|
![ajouter à partir d’une page de propriétés de fichier définissez le nom du manifeste](/static/menu_chooser.png)|

Sélectionnez « Installer », puis « Continuer » |
:-:|
![ajouter à partir d’une page de propriétés de fichier confirmation pour ajouter](/static/menu_warn.png)|

## Voir En action
Pour illustrer les fonctionnalités du complément, vous devez utiliser le client natif Office Outlook 2013.
* Ouvrez Outlook 2013
* Envoyez-vous un courrier électronique à partir d’une vidéo YouTube : avez-vous besoin d’une [suggestion ?](http://www.youtube.com/watch?v=oEx5lmbCKtY)
* développer le volet de compléments pour afficher un aperçu

## Questions et commentaires
* Si vous rencontrez des difficultés pour exécuter cet exemple, veuillez [consigner un problème](https://github.com/OfficeDev/https://github.com/OfficeDev/Outlook-Add-in-Javascript-ViewYouTubeVideos/issues).
* Si vous avez des questions générales sur le développement de compléments Office, envoyez-les sur [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Posez vos questions ou envoyez vos commentaires en incluant la balise [office-addins].

## Ressources supplémentaires
* [Autres exemples de compléments](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [Compléments Outlook](https://dev.office.com/code-samples#?filters=web,outlook)
* [API Outlook](https://dev.outlook.com/)
* [Exemples de code : Centre de développement Office](https://dev.office.com/code-samples#?filters=web,outlook)
* [Dernières actualités : Centre de développement Office](http://dev.office.com/latestnews)
* [Formation : Centre de développement Office](https://dev.office.com/training)

## Copyright
Copyright (c) 2015 Microsoft. Tous droits réservés.


Ce projet a adopté le [code de conduite Open Source de Microsoft](https://opensource.microsoft.com/codeofconduct/). Pour en savoir plus, reportez-vous à la [FAQ relative au code de conduite](https://opensource.microsoft.com/codeofconduct/faq/) ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.
