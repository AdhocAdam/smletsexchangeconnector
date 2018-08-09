---
title: "What Suggested Knowledge Articles/Request Offerings WOULD have been sent?"
excerpt_separator: "<!--more-->"
layout: post
author: Adam
---

You have the Cireson portal. You have the SMLets Exchange Connector and you are looking to put your SCSM deployment to the test by enabling the connector's ability to automatically Suggest Knowledge Articles and/or Request Offerings based on the content of the email. How useful will this feature actually be to your deployment if enabled? How much time do you stand to shave off per day? Not to mention you may even further be venturing down the Azure Cognitive Services route to see if you can improve the accuracy of the suggestions through picking out keywords instead of submitting the entire Subject/Body to the Cireson Web API. Just one problem - short of a stage environment that 100% mirrors your production environment, there isn't a good way to test this feature without just turning it on.

<!--more-->

Fret not, because once again it's time to do some function extraction from the connector and give you the ability to demo this feature against the your last month's worth of Work Items and see what the connector _would have_ suggested to the Affected User. We'll dump suggested Request Offerings and Knowledge Articles out to two seperate CSVs so you can freely analyze the data yourself and look for ways to improve the quality of your articles and/or request offerings.

To use the following script, you just need to change a few variables.
- The name of your management server if you're running this remotely
- Your Cireson portal web server
- The minimum number of words you want to match before a result is returned
- Point at Incidents or Service Requests
- (optional) use of Azure Cognitive Services Text Analytics API. Leave this blank if you don't use ACS.

If you are considering or just want to see what Azure Cognitive Services has to offer our Wiki offers a step by step on how to get setup with this here https://github.com/AdhocAdam/smletsexchangeconnector/wiki/Azure-Cognitive-Services-setup. In the event there is any confusion around how the connector leverages this feature or why it should be seriously considered for use take the following example. A new email comes in with the body of 

> I'd like to order a new laptop. How can I get this process started?

This _entire_ question will be submitted to the Cireson Web API to find Request Offering/Knowledge Articles matches. That means words such as "like" or "I'd" will be used in the search. With Azure Cognitive Services enabled, the message is first passed to ACS, and then the Cireson Web APIs. In this exact example, ACS returns 

> new laptop process

This results in improved accuracy and the incentive to set the minimum word match _lower_.

Anyway. The script! I suggest trying this on a small batch of work items first just to make sure everything is working the way you expect it to. You can do this by simply decreasing the Created Date range defined in the $cString variable near the beginning. The KA/RO functions seen here are used from the currently in development v1.4.5 of the connector which sorts results by number of words matched.

```powershell
#simulate RO suggestions through Azure Cognitive Services (ACS) as though the Exchange Connector was doing this
$mgmtServer = "localhost"
$azureCogSvcAPIKey = ""
$ciresonPortalServer = "https://webserver.domain.tld/"
$numberOfWordsToMatchFromEmailToRO = 1
$wiType = "incident" # "servicerequest"

$wiClass = get-scsmclass -name "system.workitem.$wiType$" -computername $mgmtServer
$affectedUserRelClass = get-scsmrelationshipclass -name "System.WorkItemAffectedUser$" -computername $mgmtServer
$sourceEmailEnum = (get-scsmenumeration -name "$wiType+SourceEnum.Email$" -ComputerName $mgmtServer).id
$cType = "Microsoft.EnterpriseManagement.Common.EnterpriseManagementObjectCriteria"
$cString = "Source = '$sourceEmailEnum' and CreatedDate > '7-1-2018' and CreatedDate < '7-31-2018'"
$crit = new-object $cType $cString, $wiClass
$emailWorkItems = Get-SCSMObject -criteria $crit -computername $mgmtServer

function Get-CiresonPortalUser ($username, $domain)
{
    $isAuthUserAPIurl = "api/V3/User/IsUserAuthorized?userName=$username&domain=$domain"
    $returnedUser = Invoke-WebRequest -Uri ($ciresonPortalServer+$isAuthUserAPIurl) -Method post -UseDefaultCredentials -SessionVariable userWebRequestSessionVar
    $ciresonPortalUserObject = $returnedUser.Content | ConvertFrom-Json
    return $ciresonPortalUserObject
}

function Search-AvailableCiresonPortalOfferings ($searchQuery, $ciresonPortalUser)
{
    $serviceCatalogAPIurl = "api/V3/ServiceCatalog/GetServiceCatalog?userId=$($ciresonPortalUser.id)&isScoped=$($ciresonPortalUser.Security.IsServiceCatalogScoped)"
    $serviceCatalogResults = Invoke-WebRequest -Uri ($ciresonPortalServer+$serviceCatalogAPIurl) -Method get -UseDefaultCredentials -SessionVariable ecWebSession
    $serviceCatalogResults = $serviceCatalogResults.Content | ConvertFrom-Json

    #### If the user has access to some Request Offerings, find which RO Titles/Description contain words from their original message ####
    if ($serviceCatalogResults)
    {
        #prepare the results by generating a URL array for the email
        $matchingRequestURLs = @()
        foreach ($serviceCatalogResult in $serviceCatalogResults)
        {
            $wordsMatched = ($searchQuery.Split() | ?{($serviceCatalogResult.RequestOfferingTitle -match "\b$_\b") -or ($serviceCatalogResult.RequestOfferingDescription -match "\b$_\b")}).count
            if ($wordsMatched -ge $numberOfWordsToMatchFromEmailToRO)
            {
                $ciresonPortalRequestURL = "`"" + $ciresonPortalServer + "SC/ServiceCatalog/RequestOffering/" + $serviceCatalogResult.RequestOfferingId + "," + $serviceCatalogResult.ServiceOfferingId + "`""
                $RequestOfferingURL = "<a href=$ciresonPortalRequestURL/>$($serviceCatalogResult.RequestOfferingTitle)</a><br />"
                $requestOfferingSuggestion = New-Object System.Object
                $requestOfferingSuggestion | Add-Member -type NoteProperty -name RequestOfferingURL -value $RequestOfferingURL
                $requestOfferingSuggestion | Add-Member -type NoteProperty -name WordsMatched -value $wordsMatched
                $matchingRequestURLs += $requestOfferingSuggestion
            }
        }
        $matchingRequestURLs = ($matchingRequestURLs | sort-object WordsMatched -Descending).RequestOfferingURL
        return $matchingRequestURLs
    }
}

function Search-CiresonKnowledgeBase ($ciresonPortalUser, $searchQuery)
{
    $results = Invoke-WebRequest -Uri ($ciresonPortalServer + "api/V3/KnowledgeBase/GetHTMLArticlesFullTextSearch?userId=$($ciresonPortalUser.ID)&searchValue=$searchQuery&isManager=$([bool]$ciresonPortalUser.KnowledgeManager)&userLanguageCode=$($ciresonPortalUser.LanguageCode)") -UseDefaultCredentials
    $kbResults = $results.Content | ConvertFrom-Json
    $kbResults =  $kbResults | ?{$_.endusercontent -ne ""} | select-object articleid, title

    if ($kbResults)
    {
	    #prepare the results by generating a URL array for the email
        $matchingKBURLs = @()
        foreach ($kbResult in $kbResults)
        {
            $wordsMatched = ($searchQuery.Split() | ?{($kbResult.title -match "\b$_\b")}).count
            if ($wordsMatched -ge $numberOfWordsToMatchFromEmailToKA)
            {
                $knowledgeSuggestion = New-Object System.Object
                $KnowledgeArticleURL = "<a href=$ciresonPortalServer" + "KnowledgeBase/View/$($kbResult.articleid)#/>$($kbResult.title)</a><br />"
                $knowledgeSuggestion | Add-Member -type NoteProperty -name KnowledgeArticleURL -value $KnowledgeArticleURL
                $knowledgeSuggestion | Add-Member -type NoteProperty -name WordsMatched -value $wordsMatched
                $matchingKBURLs += $knowledgeSuggestion
            }
        }
        $matchingKBURLs = ($matchingKBURLs | sort-object WordsMatched -Descending).KnowledgeArticleURL
        return $matchingKBURLs
    }
}

function Get-AzureEmailKeywords ($messageToEvaluate)
{
    #azure settings
    $azureRegion = "westus"

    #define cognitive services URLs
    $keyPhraseURI = "https://$azureRegion.api.cognitive.microsoft.com/text/analytics/v2.0/keyPhrases"

    #create the JSON request
    $documents = @()
    $requestHashtable = @{"language" = "en"; "id" = "1"; "text" = "$messageToEvaluate" };
    $documents += $requestHashtable
    $final = @{documents = $documents}
    $messagePayload = ConvertTo-Json $final

    #invoke the Text Analytics Keyword API
    $keywordResult = Invoke-RestMethod -Method Post -Uri $keyPhraseURI -Header @{ "Ocp-Apim-Subscription-Key" = $azureCogSvcAPIKey } -Body $messagePayload -ContentType "application/json"

    #return the keywords
    return $keywordResult.documents.keyPhrases
}

#create an empty array to store the resultant custom object in
$roResults = @()
$kaResults = @()

#loop through the email incidents to get Suggested Request Offerings
foreach ($emailWorkItem in $emailWorkItems)
{
    #get the Affected User's username, then get their Cireson portal user object/equivalent
    try{
        $affectedUser = get-scsmobject -id (get-scsmrelationshipobject -bysource $emailWorkItem -computername $mgmtServer | ?{$_.relationshipid -eq $affectedUserRelClass.id}).TargetObject.id -ComputerName $mgmtServer
    }
    catch{}

    if ($affectedUser)
    {
        #retrieve user properties
        $portalUser = Get-CiresonPortalUser -username $affectedUser.UserName -domain $affectedUser.Domain
        $analyst = [bool]$portalUser.Analyst
        $knowledgeManager = [bool]$portalUser.KnowledgeManager

        #build the search query & get the Work Item's related CIs if possible as another means to filter final CSV data set
        $relatedCIs = (get-scsmrelatedobject -smobject $emailWorkItem -relationship (get-scsmrelationshipclass -name "System.WorkItemAboutConfigItem$" -computername $mgmtServer) -computername $mgmtServer).displayname
        $textToEval = "$($emailWorkItem.title.trim()) $($emailWorkItem.description)" -join " "

        #take the Work Item Title and Description, ship to ACS, get Request Offering/Knowledge Article suggestions
        if ($azureCogSvcAPIKey)
        {
            $acsKeywords = Get-AzureEmailKeywords -messageToEvaluate $textToEval
            $suggestedRequstOfferings = Search-AvailableCiresonPortalOfferings -searchQuery $acsKeywords -ciresonPortalUser $portalUser
            $suggestedKnowledgeArticles = Search-CiresonKnowledgeBase -searchQuery $acsKeywords -ciresonPortalUser $portalUser
        }
        else
        {
            $suggestedRequstOfferings = Search-AvailableCiresonPortalOfferings -searchQuery $textToEval -ciresonPortalUser $portalUser
            $suggestedKnowledgeArticles = Search-CiresonKnowledgeBase -searchQuery $textToEval -ciresonPortalUser $portalUser
        }
        $suggestedRequstOfferings = $suggestedRequstOfferings -join ","
        $suggestedKnowledgeArticles = $suggestedKnowledgeArticles -join ","
        if ($azureCogSvcAPIKey) {$acsKeywords = $acsKeywords -join ","}
        $relatedCIs = $relatedCIs -join ","

        #create a custom powershell object to store all the Request Offering results per Work Item into
        $requestOfferingSuggestion = new-object system.object
        $requestOfferingSuggestion | Add-Member -type NoteProperty -name isAnalyst -value $analyst
        $requestOfferingSuggestion | Add-Member -type NoteProperty -name isKnowledgeManager -value $knowledgeManager
        $requestOfferingSuggestion | Add-Member -type NoteProperty -name AffectedUserDisplayName -value $portaluser.Name
        $requestOfferingSuggestion | Add-Member -type NoteProperty -name AffectedUserUserName -value $portaluser.UserName
        $requestOfferingSuggestion | Add-Member -type NoteProperty -name WorkItemTitle -value $emailWorkItem.Title
        $requestOfferingSuggestion | Add-Member -type NoteProperty -name WorkItemDescription -value $emailWorkItem.Description
        if ($azureCogSvcAPIKey) {$requestOfferingSuggestion | Add-Member -type NoteProperty -name ACSKeywords -value $acsKeywords}
        $requestOfferingSuggestion | Add-Member -type NoteProperty -name SuggestedROs -value $suggestedRequstOfferings
        $requestOfferingSuggestion | Add-Member -type NoteProperty -Name RelatedCIs -value $relatedCIs
        $roResults += $requestOfferingSuggestion

        #create a custom powershell object to store all the Knowledge Articles results per Work Item into
        $knowledgeSuggestion = new-object system.object
        $knowledgeSuggestion | Add-Member -type NoteProperty -name isAnalyst -value $analyst
        $knowledgeSuggestion | Add-Member -type NoteProperty -name isKnowledgeManager -value $knowledgeManager
        $knowledgeSuggestion | Add-Member -type NoteProperty -name AffectedUserDisplayName -value $portaluser.Name
        $knowledgeSuggestion | Add-Member -type NoteProperty -name AffectedUserUserName -value $portaluser.UserName
        $knowledgeSuggestion | Add-Member -type NoteProperty -name WorkItemTitle -value $emailIncident.Title
        $knowledgeSuggestion | Add-Member -type NoteProperty -name WorkItemDescription -value $emailIncident.Description
        if ($azureCogSvcAPIKey) {$knowledgeSuggestion | Add-Member -type NoteProperty -name ACSKeywords -value $acsKeywords}
        $knowledgeSuggestion | Add-Member -type NoteProperty -name SuggestedKAs -value $suggestedKnowledgeArticles
        $knowledgeSuggestion | Add-Member -type NoteProperty -Name RelatedCIs -value $relatedCIs
        $kaResults += $knowledgeSuggestion
    }
}

#export the array of custom objects to csv
$roResults | export-csv -path "c:\temp\SuggestedROResults.csv"
$kaResults | export-csv -path "c:\temp\SuggestedKAResults.csv"
```

Armed with this data you can now begin to refine the titles/descriptions of Request Offerings and/or the content of your Knowledge Articles featured on your respective Cireson portal. Perhaps all it takes is adding a few choice keywords or lowering the minimum word match count. In either case, I hope to have provided you with the knowledge to make a more informed decision and means to improve your Service Manager deployment with the SMLets Exchange Connector.
