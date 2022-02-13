using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

Install-Module "PNP.PowerShell" -AcceptLicense -Force -Verbose -Scope "CurrentUser"
Import-Module "PNP.PowerShell"

# Comments
# https://michalsacewicz.com/automatically-translate-news-on-multilingual-sharepoint-sites/
# https://www.youtube.com/watch?v=plS_1BsQAto&t=457s

# Sample
$req = @'
{
    "siteURL": "https://spjeff.sharepoint.com/sites/Portal",
    "language": "es",
    "pageTitle": "Contoso"
}
'@

# POST method: $req
#DEBUG $Request   = $req | ConvertFrom-Json
$clientId = "--GUID-HERE--"
$clientSecret = "--SECRET-HERE--"
 
# Interact with body of the request
$siteURL = $Request.Body.siteURL
$language = $Request.Body.language
$pageTitle = $Request.Body.pageTitle
 
# Translate function
function Start-Translation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$text,
        [Parameter(Mandatory = $true)]
        [string]$language
    )
 
    # Config
    $baseUri = "https://api.cognitive.microsofttranslator.com/translate?api-version=3.0"
    $headers = @{
        'Ocp-Apim-Subscription-Key'    = '--KEY-HERE--'
        'Ocp-Apim-Subscription-Region' = '--REGION-HERE--'
        'Content-Type'                 = 'application/json'
    }
 
    # Encoding clean
    $enc = [System.Text.Encoding]::UTF8.GetBytes($text)
    $text = [System.Text.Encoding]::ASCII.GetString($enc) | Out-String

    # Create JSON array with 1 object for request body
    $textJson = @{
        "Text" = "$text"
    } | ConvertTo-Json
    $body = "[$textJson]"
 
    # Uri for the request includes language code and text type, which is always html for SharePoint text web parts
    $uri = "$baseUri&from=en&to=$language&textType=html"
 
    # Send request for translation and extract translated text
    $results = Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $body
    $translatedText = $results[0].translations[0].text
    return $translatedText
}
 
#---START SCRIPT---#
Connect-PnPOnline $siteURL -ClientId $clientId -ClientSecret $clientSecret -WarningAction SilentlyContinue

$newPage = Get-PnPClientSidePage "$language/$pageTitle.aspx"
$newPage
$newPage.Controls | Ft -a
$newPage.Controls | select Type | Ft -a
$textControls = $newPage.Controls | Where-Object { $_.Type.Name -eq "ClientSideText" -or $_.Type.Name -eq "PageText" }
$textControls
 
Write-Host "Translating content..." -NoNewline

foreach ($textControl in $textControls) {
    $translatedControlText = Start-Translation -text $textControl.Text -language $language
    $translatedControlText 
    Set-PnPClientSideText -Page $newPage -InstanceId $textControl.InstanceId -Text $translatedControlText
}
    
Write-Host "Done!" -ForegroundColor "Green"

#DEBUG Out-File -Encoding "Ascii" -FilePath $res -inputObject "{'message':'Page $pageTitle has been translated to $language'}"
