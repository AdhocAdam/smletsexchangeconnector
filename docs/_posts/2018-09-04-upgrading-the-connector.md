---
title: "Upgrading the Connector"
excerpt_separator: "<!--more-->"
layout: post
author: Adam
---

So let's say the latest version of the SMLets Exchange Connector is out OR you prefer to live dangerously using an in development branch. What's true in both cases is you want it deployed. But upgrading this script could potentially lead to some sweaty palms as you cross check what settings you've enabled from what are otherwise set to a default of $false on the download. Wouldn't it be great if there was some way quick and easy way to compare your customized connector to the latest vanilla one you've just downloaded? If you are thinking we're just going to diff these files you're right, but what program are we going to use to do it? The same one the SMLets Exchange Connector is currently maintained in - [Visual Studio Code](https://code.visualstudio.com/).

<!--more-->

If you're editing PowerShell in the ISE, editing JS for your Cireson Portal or hacking up XML for your management packs in Notepad++ it's time to give [Visual Studio Code](https://code.visualstudio.com/) a try. Apart from it's ability to replace the Windows PowerShell ISE (which Microsoft has said will happen as a result of dropping the 'Windows' from PowerShell), it also serves the truly awesome purpose of being your defacto code editor for a host of languages. SQL? Query away. Markdown? Of course. HTML? Why not. Python? You know it. Plus, what doesn't ship out of box there is a free marketplace for built right into the editor to go get your language of choice for syntax highlighting. Almost forgot, VS Code integrates with Git _and_ TFS.

Alright enough about how much I enjoy [Visual Studio Code](https://code.visualstudio.com/). Let's run a diff of the connector that is running in our environment against the connector we just downloaded. Not only will you be able to quickly pinpoint what's different but you'll be able to see what's changed between versions. To start, let's open up the current running/live/in production connector and one recently downloaded from GitHub in Code.

![](/smletsexchangeconnector/images/UpgradingTheConnector/01.png)

On the left hand menu, top most button make sure it's expanded and you can clearly see which script is which. Right click your current in production script and choose "Select for Compare". Then right click the one you just downloaded from GitHub (i.e. the latest version) and choose "Compare with Selected". This will open a brand new compare window showing you the diff between both files. 

![](/smletsexchangeconnector/images/UpgradingTheConnector/02.png) ![](/smletsexchangeconnector/images/UpgradingTheConnector/03.png)

| ![](/smletsexchangeconnector/images/UpgradingTheConnector/04.png) | 
|:--:| 
| *Comparing v1.4.4 to an in development v1.4.5* |

With the magnified scroll window on the right of your screen you can quickly hop to see specifically what has changed signified by the small horizontal bars within the scroll. Green is addition. Red is deletion. This gives you the ability to see what new functionality has been added as well as what variables you'll want to update to keep parity with whatever features you've currently enabled. Not to mention, as you flip items from $false to $true you'll see these lines/sections drop their highlighting. Pretty nifty! But even _more_ useful is the ability to edit directly within this new compare Window as it will in turn update those individual windows/files as though you were editing them there.

Get those Service Manager Change and/or Release Requests ready, because today, you upgrade!