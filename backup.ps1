
# Improvements
# - Work with classes: https://xainey.github.io/2016/powershell-classes-and-concepts/

# Config
[string]$userName = 'martin.studer'
[string[]]$confluenceProjectShortcuts = @("PROR", "RW", "QIDLUW")
[string[]]$jiraProjectShortcuts = @("PROR", "RW", "QUAL")

# Globals
[string]$path = Split-Path -Path $MyInvocation.MyCommand.Path
[string]$timestamp = Get-Date -Format o | foreach {$_ -replace ":", "."}


# Build headers for basic authentication for web requests
function Build-Headers($credential) {
    # https://stackoverflow.com/questions/27951561/use-invoke-webrequest-with-a-username-and-password-for-basic-authentication-on-t

    [string]$user = $credential.UserName
    [string]$password = $credential.GetNetworkCredential().Password

    $pair = "${user}:${password}"

    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)    $base64 = [System.Convert]::ToBase64String($bytes)

    $basicAuthValue = "Basic $base64"

    $headers = @{ Authorization = $basicAuthValue }

    return $headers
}


# Get download Uri for given confluence project with given shortcut
function Get-ConfluenceDownloadUri([string]$projectShortcut, $credential) {
    [string]$username = $credential.UserName
    [string]$password = $credential.GetNetworkCredential().Password

    $ie = New-Object -ComObject 'internetExplorer.Application'
    $ie.Visible= $true

    $ie.Navigate("https://sife-net.htwchur.ch/login.action")
    while ($ie.Busy -eq $true) {Start-Sleep -Milliseconds 100;}   

    # Login
    $usernamefield = $ie.Document.getElementByID('os_username')
    $usernamefield.value = $username

    $passwordfield = $ie.Document.getElementByID('os_password')
    $passwordfield.value = $password

    $Link=$ie.Document.getElementByID("loginButton")
    $Link.click()
    while ($ie.Busy -eq $true) {Start-Sleep -Milliseconds 100;}   
    
    # Export

    $ie.Navigate("https://sife-net.htwchur.ch/spaces/exportspacewelcome.action?key=$projectShortcut")
    while ($ie.Busy -eq $true) {Start-Sleep -Milliseconds 100;}   

    # Select radio buttons:
    # https://social.technet.microsoft.com/Forums/windowsserver/en-US/361fa844-3170-46ff-80b3-37ae4ae40f07/power-shell-script-how-to-selectclick-a-radio-button-?forum=winserverpowershell
    $myradios = $ie.Document.getElementsByTagName('input') | ? {$_.type -eq 'radio' -and $_.name -eq 'format'}
    $x = 0 #specific ridio button 
    $myradios[$x].setActive()
    $myradios[$x].click()

    $i=$ie.Document.getElementsByTagName("input")

    $j = $i | where-object {$_.Name -eq "confirm"}
    $j.click();
    while ($ie.Busy -eq $true) {Start-Sleep -Milliseconds 100;}   

    $i=$ie.Document.getElementsByTagName("input")
    $j = $i | where-object {$_.Name -eq "confirm"}
    $j.click();

    while ($ie.Busy -eq $true) {Start-Sleep -Milliseconds 100;}   

    Write-Host "Waiting ..."
    do {
        $i=$ie.Document.getElementsByTagName("a")
        $j = $i | where-object {$_.getAttributeNode('class').Value -eq 'space-export-download-path'}
        Start-Sleep -Milliseconds 100;
    } while (!($j))
    Write-Host "Done"
    
    $uri = $j.HREF

    $ie.quit()

    return $uri

}


# Save 
function Save-ConfluenceBackup([string]$projectShortcut, [string]$outputDir, $uri, $credential) {
    [string]$outputFilename = "$projectShortcut.zip"
    [string]$outputFilepath = Join-Path -Path $outputDir -ChildPath $outputFilename

    $headers = Build-Headers($credential)

    Invoke-WebRequest $uri -OutFile $outputFilepath -Headers $headers
}


function Get-JiraBackup([string]$projectShortcut, [string]$outputDir, $credential) {
    [string]$outputFilename = "$projectShortcut.doc"
    [string]$outputFilepath = Join-Path -Path $outputDir -ChildPath $outputFilename

    [string]$uri = "https://jira.htwchur.ch/sr/jira.issueviews:searchrequest-word/temp/SearchRequest.doc?jqlQuery=project+%3D+$projectShortcut+ORDER+BY+created+DESC&tempMax=1000"

    # Need headers with basic authentication
    # https://community.atlassian.com/t5/Jira-Software-questions/REST-error-quot-The-value-XXX-does-not-exist-for-the-field/qaq-p/654730
    # https://developer.atlassian.com/server/jira/platform/basic-authentication/
    $headers = Build-Headers($credential)

    Invoke-WebRequest $uri -OutFile $outputFilepath -Headers $headers
}

function Save-AllJiraBackups($credential, $projectShortcuts) {
    [string]$outputPath = Join-Path -Path $path -ChildPath 'Jira'
    [string]$outputTimestampPath = Join-Path -Path $outputPath -ChildPath $timestamp
    mkdir $outputTimestampPath

    foreach ($projectShortcut in $projectShortcuts) {
        JiraBackup -projectShortcut $projectShortcut -outputDir $outputTimestampPath -credential $credential
    }
}

function Save-AllConfluenceBackups($projectShortcuts, $credential) {
    [string]$outputPath = Join-Path -Path $path -ChildPath 'Confluence'
    [string]$outputPathTimestampPath = Join-Path -Path $outputPath -ChildPath $timestamp
    mkdir $outputPathTimestampPath
    foreach ($projectShortcut in $confluenceProjectShortcuts) {
        $uri = Get-ConfluenceDownloadUri -projectShortcut $projectShortcut -credential $confluenceCredential
        Save-ConfluenceBackup -projectShortcut $projectShortcut -outputDir $outputPathTimestampPath -uri $uri -credential $confluenceCredential
    }
}

$confluenceCredential = Get-Credential -Message 'Confluence' -UserName $userName
$jiraCredential = Get-Credential -Message 'Jira' -UserName $userName

Save-AllConfluenceBackups -projectShortcuts $projectShortcuts -credential $confluenceCredential
Save-AllJiraBackups -projectShortcuts $jiraProjectShortcuts -credential $jiraCredential