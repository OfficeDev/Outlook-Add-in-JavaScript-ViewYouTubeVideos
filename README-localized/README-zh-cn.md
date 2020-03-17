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

## 摘要
如果所选电子邮件或约会包含 YouTube 视频的 URL，则此 Outlook 加载项允许用户在 Outlook 的加载项窗格中观看 YouTube 视频。它还包含一个安装脚本，该脚本可将加载项部署到在 Mac 上运行的 Ruby Web 服务器上。下图是在阅读窗格中为邮件激活的 YouTube 加载项的屏幕截图。
<br />
<br />
![在邮件项目中运行 YouTube 视频的 Outlook 加载项](/static/pic1.png)

## 先决条件
* Mac OS X 10.10 或更高版本
* Bash
* Ruby 2.2.x+
* [捆绑程序](http://bundler.io/v1.5/gemfile.html)
* OpenSSL
* 运行至少具有一个电子邮件帐户或 Office 365 帐户的 Exchange 2013 的计算机。你可以[加入 Office 365 开发人员计划并获得 Office 365 的 1 年免费订阅](https://aka.ms/devprogramsignup)。
* 任何支持 ECMAScript 5.1、HTML5 和 CSS3 的浏览器，例如 Chrome、Firefox 或 Safari。
* Outlook 2016 for Mac

## 示例的主要组件
* [```LICENSE.txt```](LICENSE.txt) 使用此可发行软件包的条款和条件
* [```config.ru```](config.ru) 机架配置
* [```setup.sh```](setup.sh) 用于生成 ```app.rb```、```manifest.xml``` 和（可选）证书的安装脚本
* [```cert/ss_certgen.sh```](cert/ss_certgen.sh) 生成脚本的自签名证书
* [```public/res/js/strings_en-us.js```](public/res/js/strings_en-us.js) 美国英语本地化
* [```public/res/js/strings_fr-fr.js```](public/res/js/strings_fr-fr.js) 法语本地化

## 代码说明

此加载项的主要代码文件包括 ```manifest.xml``` 和 ```youtube.html```，以及 Office 加载项的 JavaScript 库和字符串文件及徽标图像文件。以下是该加载项的工作方式的高级概述：

此邮件加载项在 ```manifest.xml``` 文件中指定它需要支持邮箱功能的主机应用程序：

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
    
此加载项还会请求清单文件中的 ReadItem 权限，以便运行如下所述的正则表达式。

```xml
    <Permissions>ReadItem</Permissions>
```
    
如果所选邮件或约会包含 YouTube 视频的 URL，则主机应用程序将激活此加载项。这是通过在启动时首先读取 manifest.xml 文件实现的，该文件指定了激活规则，其中包含用于查找此类 URL 的正则表达式：

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" PropertyName="BodyAsPlaintext" RegExName="VideoURL" RegExValue="http://(((www\.)?youtube\.com/watch\?v=)|(youtu\.be/))[a-zA-Z0-9_-]{11}"/>
```
    
该加载项定义 initialize 函数，它是 initialize 事件的事件处理程序。加载运行时环境时，将引发 initialize 事件，然后 initialize 函数会调用加载项的主要函数 `init`，如下面的代码所示：

```javascript
Office.initialize = function () {
    init(Office.context.mailbox.item.getRegExMatches().VideoURL);
}
```

所选项目的 ```getRegExMatches``` 方法将返回与在 ```manifest.xml``` 文件中指定的正则表达式 ```VideoURL``` 匹配的字符串数组。在本例中，该数组包含 YouTube 视频的 URL。

`init` 函数将 YouTube URL 数组作为输入参数，并动态构建 HTML 以显示每个视频的相应缩略图和详细信息。

此动态构建的 HTML 会在 YouTube 嵌入式播放器中显示第一个视频以及有关该视频的详细信息。加载项窗格还将显示任何后续视频的缩略图。最终用户可以选择缩略图以在 YouTube 嵌入式播放器中观看任何视频，而无需离开主机应用程序。

## 设置
此示例附带了 ```setup.sh``` - 此安装文件可执行以下操作：
* 验证并安装[依赖项](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L23)
* [生成 ```manifest.xml ```](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L37)
* [生成 ```app.rb```](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L44)
* [（可选）](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L34)生成[自签名证书](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/cert/ss_certgen.sh#L49)

要运行脚本，请在 POSIX 兼容终端上键入：

    $ bash setup.sh
    
## 启动服务器
从项目根目录运行：

    $ rackup

## 信任你的自签名证书
由于此示例使用本地服务器和[自签名证书](https://en.wikipedia.org/wiki/Self-signed_certificate)，因此必须先在本地主机和自签名证书之间建立“信任”。在 Outlook 将任何潜在敏感数据传输到任何加载项之前，必须先对其 SSL 证书进行信任验证。此项要求有助于保护数据的隐私。任何新式 Web 浏览器都会提醒用户注意证书的差异，并且许多浏览器还提供了检查和建立信任的机制。启动本地服务器后，打开你选择的 Web 浏览器，然后导航到在 manifest.xml 文件中指定的本地托管 URL。（默认情况下，此示例中的 setup.sh 脚本会将此 URL 指定为 ```https://0.0.0.0:8443/youtube.html```。） 此时，你可能会收到证书警告。你需要信任此证书。

打开 Safari|
:-:|
![用于验证证书的“Safari 安全性”对话框](/static/show_cert.png)|

选择“始终信任”你的自签名证书|
:-:|
![用于始终信任 Contoso 证书的“Safari 安全性”对话框](/static/add_trust.png)|

## 将加载项安装到 Office 365
安装此示例加载项需要访问 Outlook 网页版。可以从“设置”>“管理应用”执行安装。

依次选择“设置”和“管理应用”菜单|从文件安装
:-:|:-:
![用于管理应用的“设置”下拉菜单](/static/menu_loc.png)|![从文件设置属性页面添加](/static/menu_opt.png)

选择 manifest.xml 文件|
:-:|
![从文件属性页面设置中添加清单名称](/static/menu_chooser.png)|

依次选择“安装”和“继续”|
:-:|
![从文件属性页面设置中添加并确认添加](/static/menu_warn.png)|

## 查看实际操作
为了演示加载项的功能，你需要使用 Office Outlook 2013 本机客户端。
* 打开 Outlook 2013
* 通过电子邮件向自己发送 YouTube 视频链接 - 需要[建议？](http://www.youtube.com/watch?v=oEx5lmbCKtY)
*展开加载项窗格以查看预览

## 问题和意见
* 如果你在运行此示例时遇到任何问题，请[记录问题](https://github.com/OfficeDev/https://github.com/OfficeDev/Outlook-Add-in-Javascript-ViewYouTubeVideos/issues)。
* 与 Office 加载项开发相关的问题一般应发布到 [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins)。确保你的问题或意见标记有 [Office 加载项]。

## 其他资源
* [更多加载项示例](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [Outlook 加载项](https://dev.office.com/code-samples#?filters=web,outlook)
* [Outlook API](https://dev.outlook.com/)
* [代码示例 - Office 开发人员中心](https://dev.office.com/code-samples#?filters=web,outlook)
* [最新新闻 - Office 开发人员中心](http://dev.office.com/latestnews)
* [培训 - Office 开发人员中心](https://dev.office.com/training)

## 版权信息
版权所有 (c) 2015 Microsoft。保留所有权利。


此项目已采用 [Microsoft 开放源代码行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅[行为准则 FAQ](https://opensource.microsoft.com/codeofconduct/faq/)。如有其他任何问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。
