---
title: "Preserving Customizations through Upgrades"
excerpt_separator: "<!--more-->"
layout: post
author: Adam
---

If you've joined myself and others on this open source Exchange Connector journey, something that may have come up is "Inserting custom logic on this thing is doable, but I could easily see myself forgetting to do it on an upgrade" or "Can't there be someway to preserve my own organization logic between upgrades?"

It's a great question and one whose answer deserves its own post.

<!--more-->

**Custom Events**  
The connector features the ability to call an external PowerShell script of Functions that is invoked during actions such as Before-ProcessEmail, After-ProcessEmail, Before-Take, After-Take, etc. etc. Given all of these events point to an external PowerShell script that we don't touch during upgrades, you're free to put as little or as much logic inside this as you'd like.

**Runbooks**  
While invoking Custom Events through PowerShell is one thing, perhaps you're looking to leverage more of your native System Center infrastructure without having to write PowerShell to do so. Which is why an alternative path forward could be creating a "Diagnostic" style runbook in System Center Orchestrator (SCO), attaching it to your Default Work Item template that the connector applies when Creating New Work Items from email, and then on every Work Item Creation that runbook is triggered. As such, this would be a runbook/event that occurs after the initial creation of a Work Item (the equivalent of After-ProcessEmail from Custom Events above) but through a more native System Center means _without_ slowing down connector processing in the slightest. But what kind of things could we put in said Diagnostic runbook?

- Get the Work Item Title, Description, and Affected User. Based on some internal matrix of keywords/individuals, re-assign this Work Item to a different Support Group
- Get the Work Item Affected User. If they are Person 1, 2, or 3 do X. Else, do Y.
- Struggling with vendors and their repsective ticketing systems? Use a Runbook to Regex parse your vendor's Ticket ID (Title/Description) into a custom class extension such as ExternalTicketID, then introduce matching logic directly on the connector thereby ensuring subsequent runs update your SINGLE Work Item in SCSM preventing duplication
- Based on keywords found within the Description, route to specific Analysts
- Extend your User class with a string (or relationship) such as "Specialities." Then based on words found within the Description triage to Analysts based on matching criteria

Or better yet - do some combination of all of these. Because your single Diagnostic runbook could just as easily invoke one or many child runbooks based on your orchestration logic. That way SCSM only has to call 1 runbook, while Orchestrator invokes 1 or Many thereby spreading the workload to the most appropriate tools within System Center.

Ultimatley what I hope you can see is that you have the freedom to implement as little or as much customization as you want, wherever and however you choose.