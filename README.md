# Outlook-Add-in-JavaScript-ViewYouTubeVideos

## Summary
This Outlook Add-in allows users to view YouTube videos in the add-in pane in Outlook if the selected email message or appointment contains a URL to a video on YouTube. It also contains a setup script that deploys the add-in to a Ruby web server running on a Mac. The following figure is a screen shot of the YouTube add-in activated for a message in the Reading Pane.
<br />
<br />
![Outlook Addin running a YouTube video in the mail item](/static/pic1.png)

## Prerequisites
* Mac OS X 10.10 or later
* Bash
* Ruby 2.2.x+
* [Bundler](http://bundler.io/v1.5/gemfile.html)
* OpenSSL
* A computer running Exchange 2013 with at least one email account, or an Office 365 account. You can [join the Office 365 Developer Program and get a free 1 year subscription to Office 365](https://aka.ms/devprogramsignup).
* Any browser that supports ECMAScript 5.1, HTML5, and CSS3, such as Chrome, Firefox, or Safari.
* Outlook 2016 for Mac

## Key components of the sample
* [```LICENSE.txt```](LICENSE.txt) The terms and conditions of using this distributable
* [```config.ru```](config.ru) Rack config
* [```setup.sh```](setup.sh) Setup script to generate ```app.rb```, ```manifest.xml```, and optionally, a certificate
* [```cert/ss_certgen.sh```](cert/ss_certgen.sh) Self-signed certificate generating script
* [```public/res/js/strings_en-us.js```](public/res/js/strings_en-us.js) US English localization
* [```public/res/js/strings_fr-fr.js```](public/res/js/strings_fr-fr.js) French localization

## Description of the code

The main code files for this add-in are ```manifest.xml``` and ```youtube.html```, along with the JavaScript library and string files for Office add-ins, and a logo image file. The following is a high-level summary of how the add-in works:

This mail add-in specifies in the ```manifest.xml``` file that it requires a host application that supports the mailbox capability:

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
    
The add-in also requests the ReadItem permission in the manifest file so that it can run regular expressions, which is explained below.

```xml
    <Permissions>ReadItem</Permissions>
```
    
The host application activates this add-in when the selected message or appointment contains a URL to a YouTube video. It does so by first reading on startup the manifest.xml file, which specifies an activation rule that includes a regular expression to look for such a URL:

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" PropertyName="BodyAsPlaintext" RegExName="VideoURL" RegExValue="http://(((www\.)?youtube\.com/watch\?v=)|(youtu\.be/))[a-zA-Z0-9_-]{11}"/>
```
    
The add-in defines an initialize function that is an event handler for the initialize event. When the run-time environment is loaded, the initialize event is fired, and then the initialize function calls the main function of the add-in, `init`, as shown in the code below:

```javascript
Office.initialize = function () {
    init(Office.context.mailbox.item.getRegExMatches().VideoURL);
}
```

The ```getRegExMatches``` method of the selected item returns an array of strings that match the regular expression ```VideoURL```, which is specified in the ```manifest.xml``` file. In this case, that array contains URLs to videos on YouTube.

The `init` function takes as an input parameter that array of YouTube URLs and dynamically builds the HTML to display the corresponding thumbnail and details for each video.

This dynamically built HTML displays the first video in a YouTube embedded player, together with details about the video. The add-in pane also displays the thumbnails of any subsequent videos. The end user can choose a thumbnail to view any of the videos in the YouTube embedded player, without leaving the host application.

## Set up
Shipped with this sample is a ```setup.sh``` - this setup file does the following:
* Verifies and installs [dependencies](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L23)
* [Generates the ```manifest.xml```](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L37)
* [Generates the ```app.rb```](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L44)
* [Optionally](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/setup.sh#L34) generates a [self-signed certificate](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ViewYouTubeVideos/blob/master/cert/ss_certgen.sh#L49)

To run the script, type at your POSIX-compliant terminal:

    $ bash setup.sh
    
## Start the server
From the project root, run:

    $ rackup

## Trust your self-signed certificate
Because this sample uses a local server and [self-signed certificate](https://en.wikipedia.org/wiki/Self-signed_certificate), you must first establish 'trust' between your localhost and the self-signed certificate. Before Outlook will transmit any potentially sensitive data to any add-in, its SSL Certificate is verified for trust.  This requirement helps protect the privacy of your data. Any modern web browser will alert the user to certificate discrepancies, and many also provide a mechanism for inspecting and establishing trust. After starting your local server, open your web browser of choice and navigate to the locally hosted URL specified in your manifest.xml file. (By default, the setup.sh script in this sample specifies this URL as ```https://0.0.0.0:8443/youtube.html```.) At this point you may be presented with a certificate warning. You need to trust this certificate.

Open Safari|
:-:|
![Safari security diloag to validate the certificate](/static/show_cert.png)|

Select 'Always trust' your self-signed certificate|
:-:|
![Safari security diloag to always trust the Contoso certificate](/static/add_trust.png)|

## Install the add-in to Office 365
Installation of this sample add-in requires access to Outlook on the web. Installation can be performed from Settings > Manage apps.

Select 'Settings' and 'Manage apps' menu|Install from file
:-:|:-:
![Settings dropdown for manage apps](/static/menu_loc.png)|![add from a file settings property page](/static/menu_opt.png)

Select the manifest.xml file|
:-:|
![add from a file properties page setting the manifest name](/static/menu_chooser.png)|

Select 'Install' and then 'Continue'|
:-:|
![add from a file properties page setting confirmation to add](/static/menu_warn.png)|

## See it in action
To demonstrate the functionality of the add-in, you'll need to use the Office Outlook 2013 native client.
* Open Outlook 2013
* Email yourself a link to a YouTube video - Need a [suggestion?](http://www.youtube.com/watch?v=oEx5lmbCKtY)
* Expand the add-in pane to see a preview

## Questions and comments
* If you have any trouble running this sample, please [log an issue](https://github.com/OfficeDev/https://github.com/OfficeDev/Outlook-Add-in-Javascript-ViewYouTubeVideos/issues).
* Questions about Office Add-in development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Make sure that your questions or comments are tagged with [office-addins].

## Additional resources
* [More Add-in samples](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
* [Outlook Add-ins](https://dev.office.com/code-samples#?filters=web,outlook)
* [Outlook API](https://dev.outlook.com/)
* [Code Samples - Office Dev Center](https://dev.office.com/code-samples#?filters=web,outlook)
* [Latest News - Office Dev Center](http://dev.office.com/latestnews)
* [Training - Office Dev Center](https://dev.office.com/training)

## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.
