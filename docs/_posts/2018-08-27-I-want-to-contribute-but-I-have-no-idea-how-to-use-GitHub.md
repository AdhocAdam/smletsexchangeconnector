---
title: "I want to contribute, but I have no idea how to use GitHub"
excerpt_separator: "<!--more-->"
layout: post
author: Adam
---

This is something I've heard in passing. People are interested in contributing to the connector, but they are worried they'll make a mistake OR GitHub feels foreign if you've only ever been PowerShell scripting for yourself. Let's use this blogpost as the opportunity to address both of those things and how you can comfortably get started.

<!--more-->

**How do I _actually_ contribute code to this thing?**  
First you'll need a GitHub account which takes only a couple of seconds to obtain. Next, you'll need to head over to the SMLets Exchange Connector repo and Fork the repository. The Fork icon/action can be seen in the top right hand of the repo's page.

**What does it mean to Fork the repository?**  
Take the following picture from the Git home page as an example:
![alt text](https://git-scm.com/images/branching-illustration@2x.png "Logo Title Text 1")

The picture is used to illustrate the concept of branching but for the purposes of this explanation, I'm going to use it to explain forking. In the image above, imagine the red line and it's corresponding documents are every update I (Adam) made to the connector. This is my "master" branch of the repository. But then, you come along and want to add some functionality to the connector. When you Fork the repository, you go down the first green line. In this regard, you have created the first derivative piece of work.

**Does Forking alter the original repo?**  
No. When you Fork, you create a duplicate SMlets Exchange Connector under _your_ account. It's from here you can make edits to it, because it's contained underneath _your_ account. The only thing that establishes a link between the original work and yours, is my repo's Fork number has increased by one and your copy of it shows a "Forked from..." style URL. This is the only linkage that exists to show people who visit either my original repo or your derivative repo where the original code came from. Changes that you make will be held unto your account, so there isn't a risk of a loss of credit.

**Alright, I've made my edits on my connector. How do I get them back to you so folks using the connector benefit?**  
If the first green line was the mark of a derivative work (Fork), the second green line that merges those changes back into the original is a Pull Request. This is a formal request from you to the master SMLets Exchange Connector repo, that you wish you merge changes. This Pull Request is an opportunity to review what specifically has been changed between the original repo (mine) and the repo being requested for merge (yours) on a code line by code line basis. [Here's an example of a Pull Request from a previous version](https://github.com/AdhocAdam/smletsexchangeconnector/pull/46). Pull Requests have their own conversations, Commit counts, and most importantly shows which files have changed.

**Everything checked out and the merge was approved, what happens now?**  
Your...
- Forked copy of the connector remains under your account
- Account is shown as a contributor to connector
- Forked copy can optionally be deleted
- Service Manager peers have a chance of showering you with praise

**Do I need some kind of client software to work with GitHub?**  
No, this is optional. GitHub does offer a client for Windows that some may find useful when downloading, committing, and pushing back to GitHub. But make no mistake you don't have to use any of these pieces of software. If you want you can just download the *.ps1 for the connector, edit in VS Code, and commit back to your repo. This is all up to you. 

**Before I submitted the Pull Request, I accidently uploaded some sensitive information to my version of the connector but then erased it by uploading a new copy. Am I 'safe' ?**  
Assume you are not. The [GitHub policy on this matter](https://help.github.com/articles/removing-sensitive-data-from-a-repository/) dictates that
> Once you have pushed a commit to GitHub, you should consider any data it contains to be compromised. If you committed a password, change it! If you committed a key, generate a new one

The article also shares how to erase from Git history, but no matter what when you upload sensitive data you need to operate on the premise that it was compromised. Generate new passwords, change passwords, clear the Git history, Commit a new version and exercise caution in the future. We all make mistakes, but how we move forward from them is what matters most.

**I have another question, but raising an Issue on the repo feels unnecessary. Is there some way to contact you?**  
As long as you have a GitHub account just head over to my GitHub [profile's contact section](https://github.com/AdhocAdam) for info. But hopefully at this point you should have enough to get started on using GitHub and how to further extend the SMLets Exchange Connector to take Service Manager to new levels of awesome.
