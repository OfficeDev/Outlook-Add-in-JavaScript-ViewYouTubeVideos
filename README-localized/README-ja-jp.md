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

## 概要
この Outlook アドインを使用すると、選択したメール メッセージまたは予定に YouTube ビデオへの URL が含まれている場合に、ユーザーは YouTube ビデオを Outlook の [アドイン] ウィンドウで表示することができます。この Outlook アドインには、Mac 上で実行される Ruby Web サーバーにアドインを展開する設定スクリプトも含まれています。次の図は、閲覧ウィンドウでメッセージに対して有効になっている YouTube アドインのスクリーンショットです。
<br />
<br />
![メール アイテム表示される、YouTube ビデオを実行する Outlook アドイン](/static/pic1.png)

## 前提条件
* Mac OS X 10.10 以降
* Bash
* Ruby 2.2.x 以降
* [Bundler](http://bundler.io/v1.5/gemfile.html)
* OpenSSL
* 少なくとも 1 つの電子メール アカウントまたは Office 365 アカウントがある Exchange 2013 を実行するコンピューター。[Office 365 Developer プログラムに参加すると、Office 365 の 1 年間無料のサブスクリプションを取得](https://aka.ms/devprogramsignup)できます。
* ECMAScript 5.1、HTML5、および CSS3 をサポートしている任意のブラウザー (Chrome、Firefox、Safari など)。
* Outlook 2016 for Mac

## このサンプルの主要なコンポーネント
* [```LICENSE.txt```](LICENSE.txt) この配布可能サンプルの使用条件
* [```config.ru```](config.ru) Rack の構成ファイル
* [```setup.sh```](setup.sh) ```app.rb```、```manifest.xml```、および証明書 (省略可) を生成する設定スクリプト　
* [```cert/ss_certgen.sh```](cert/ss_certgen.sh) スクリプトを生成する自己署名証明書
* [```public/res/js/strings_en-us.js```](public/res/js/strings_en-us.js) 英語 (米国) ローカリゼーション
* [```public/res/js/strings_fr-fr.js```](public/res/js/strings_fr-fr.js) フランス語 ローカリゼーション

## コードの説明

このアドインの主要なコード ファイルは、```manifest.xml``` と ```youtube.html``` の他、Office 用の JavaScript ライブラリと文字列ファイル、ロゴ画像ファイルになります。アドインの動作の大まかな概要を次に示します。

このメール アドインは、アドインにはメールボックス機能をサポートするホスト アプリケーションが必要であるということを ```manifest.xml.xml``` ファイルで指定します。

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
    
また、下で説明する正規表現をアドインで実行できるように、アドインはマニフェスト ファイルで ReadItem アクセス許可も要求します。

```xml
    <Permissions>ReadItem</Permissions>
```
    
選択したメッセージまたは予定に YouTube ビデオの URL が含まれていると、ホスト アプリケーションによりアドインが有効化されます。これは、manifest.xml ファイルを起動時に読み込むことにより実行されます。このファイルは、そのような URL を検索する正規表現が含まれる有効化ルールを指定しています。

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" PropertyName="BodyAsPlaintext" RegExName="VideoURL" RegExValue="http://(((www\.)?youtube\.com/watch\?v=)|(youtu\.be/))[a-zA-Z0-9_-]{11}"/>
```
    
アドインは、イベント初期化のイベント ハンドラーである初期化関数を定義します。次のコードに示すように、ランタイム環境が読み込まれると初期化イベントが発生し、初期化関数によりアドインのメインの関数である `init` が呼び出されます。

```javascript
Office.initialize = function () {
    init(Office.context.mailbox.item.getRegExMatches().VideoURL);
}
```

選択されたアイテムの ```getRegExMatches``` メソッドにより、```manifest.xml``` ファイルで指定される正規表現 ```VideoURL``` に一致する文字列配列が返されます。この場合、この配列には YouTube のビデオの URL が含まれています。

`init` 関数は、YouTube URL の配列を入力パラメーターとして受け取り、HTML を動的にビルドして各ビデオに対応するサムネイルと詳細を表示します。

動的に作成されたこの HTML は、最初のビデオとそのビデオの詳細を埋め込みの YouTube プレーヤーに表示します。アドイン ウィンドウには、後続のビデオのサムネイルも表示されます。エンド ユーザーは、サムネイルを選択することで、ホスト アプリケーションを離れずにどのビデオでも埋め込みの YouTube プレーヤーに表示できます。

## セットアップ
このサンプルには ```setup.sh``` が付属しています。このセットアップ ファイルにより次の操作が行われます:
* [依存関係](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L23)の検証とインストール
* [```manifest.xml``` の生成](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L37)
* [```app.rb``` の生成](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L44)
* [オプション](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L34)として、[自己署名証明書の生成](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/cert/ss_certgen.sh#L49)

スクリプトを実行するには、POSIX 準拠のターミナルに次のように入力します。

    $ bash setup.sh
    
## サーバーを起動する
プロジェクトのルートから、次のコマンドを実行します。

    $ rackup

## 自己署名証明書を信頼する
このサンプルではローカル サーバーと [自己署名証明書](https://en.wikipedia.org/wiki/Self-signed_certificate)を使用しているため、まずは localhost と自己署名証明書の間に "信頼" を築く必要があります。Outlook が機密の恐れがあるデータをアドインに送信するようになる前に、SSL 証明書の信頼性が検証されます。この要件は、データのプライバシーを保護するためのものです。すべての最新の Web ブラウザーでは、証明書の不一致があるとユーザーに警告が表示されます。信頼を検証して確立するための仕組みも多くのブラウザーで提供されています。ローカル サーバーを起動したら任意の Web ブラウザーを開き、manifest.xml ファイルが指定する、ローカルでホストされている URL に移動します。(既定では、このサンプルの setup.sh スクリプトでは、この URL として ```https://0.0.0.0:8443/youtube.html``` が指定されます。) このときに、証明書の警告が表示される可能性があります。この証明書を信頼する必要があります。

Safari を開きます|
:-:|
![証明書を検証する Safari セキュリティ ダイアログ](/static/show_cert.png)|

自己署名証明書を [常に信頼] を選択します|
:-:|
![Contoso 証明書を常に信頼する Safari セキュリティ ダイアログ](/static/add_trust.png)|

## Office 365 にアドインをインストールする
このサンプル アドインをインストールするには、Outlook on the web へのアクセスが必要です。インストールは、[設定] > [アプリの管理] から実行できます。

[設定] および [アプリの管理] メニュを選択します|ファイルからインストールします
:-:|:-:
![[設定] ドロップダウンの [アプリの管理]](/static/menu_loc.png)|![[ファイルから追加] 設定プロパティ ページ](/static/menu_opt.png)

manifest.xml ファイルを選択します|
:-:|
![[ファイルから追加] 設定プロパティ ページでのマニフェスト名](/static/menu_chooser.png)|

[インストール]、[続行] の順に選択します|
:-:|
![[ファイルから追加] 設定プロパティ ページでの追加の確認](/static/menu_warn.png)|

## 実際の動作を見る
アドインの機能をデモンストレーションするには、Office Outlook 2013 のネイティブ クライアントを使用する必要があります。
* Outlook 2013 を開きます
* YouTube ビデオへのリンクを自分自身にメールします。よろしければ[こちらのビデオ](http://www.youtube.com/watch?v=oEx5lmbCKtY)をお使いください。
* アドイン ウィンドウを展開して、プレビューを表示します

## 質問とコメント
* このサンプルの実行について問題がある場合は、[問題をログに記録](https://github.com/OfficeDev/https://github.com/OfficeDev/Outlook-Add-in-Javascript-ViewYouTubeVideos/issues)してください。
* Office アドイン開発全般の質問については、「[Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins)」に投稿してください。質問やコメントには、必ず "office-addins" のタグを付けてください。

## その他のリソース
* [その他のアドイン サンプル](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [Outlook アドイン](https://dev.office.com/code-samples#?filters=web,outlook)
* [Outlook API](https://dev.outlook.com/)
* [コード サンプル - Office デベロッパー センター](https://dev.office.com/code-samples#?filters=web,outlook)
* [最新ニュース - Office デベロッパー センター](http://dev.office.com/latestnews)
* [トレーニング - Office デベロッパー センター](https://dev.office.com/training)

## 著作権
Copyright (c) 2015 Microsoft.All rights reserved.


このプロジェクトでは、[Microsoft オープン ソース倫理規定](https://opensource.microsoft.com/codeofconduct/)が採用されています。詳細については、「[倫理規定の FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。
