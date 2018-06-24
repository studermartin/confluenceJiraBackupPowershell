
# Naming Conventions
#
# Microsoft:
# - https://msdn.microsoft.com/en-us/library/ms714428(v=vs.85).aspx
# 
# https://github.com/PoshCode/PowerShellPracticeAndStyle
# - Capitalization guidelines: https://github.com/PoshCode/PowerShellPracticeAndStyle/issues/36

# Improvements
# - Use class ConfluenceScraper instead of Get-ConfluenceDownloadUri

class IE {
    
    $ie
    [int]$waitMilliseconds = 100

    IE() {
        $this.ie = New-Object -ComObject 'internetExplorer.Application'
        $this.ie.Visible= $true
    }

    NavigateAndWait([string]$url) {
        # https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752093(v=vs.85)
        $this.ie.Navigate($url)
        $this.Wait()
    }

    ClickAndWaitById([string]$id) {
        $element=$this.ie.Document.getElementByID($id)
        $element.Click()
        $this.Wait()
    }

    # Example: <input name=confirm>
    # ClickAndWait("input", "confirm")
    ClickAndWaitByTagNameAndName([string]$tagName, [string]$name) {
        $elements=$this.ie.Document.getElementsByTagName($tagName)
        $element = $elements | where-object {$_.Name -eq $name}
        $element.click()
        $this.Wait()
    }

    # Example: 
    # WaitByTagNameAndClass("a", 'space-export-download-path') waits for an element
    #   <a class="space-export-download-path">
    [object] WaitByTagNameAndClass([string]$tagName, [string]$classValue) {
        $element = $null
        do {
            $elements=$this.ie.Document.getElementsByTagName($tagName)

#TASK: The where clause fails if there is not attribute node with name 'class'. Check the existance of the attribute node first.

            $element = $elements | where-object {$_.getAttributeNode('class').Value -eq $classValue}
            Start-Sleep -Milliseconds $this.waitMilliseconds;
        } while (!($element))
        return $element
    }

    ClickRadio0ByTagNameAndName([string]$tagName, [string]$name) {
        # Select radio buttons:
        # https://social.technet.microsoft.com/Forums/windowsserver/en-US/361fa844-3170-46ff-80b3-37ae4ae40f07/power-shell-script-how-to-selectclick-a-radio-button-?forum=winserverpowershell
        $myradios = $this.ie.Document.getElementsByTagName($tagName) | ? {$_.type -eq 'radio' -and $_.name -eq $name}
        $x = 0 #specific ridio button 
        $myradios[$x].setActive()
        $myradios[$x].click()
    }

    [object] GetElementById([string]$id) {
        return $this.ie.Document.getElementByID($id)
    }

    Wait() {
        while ($this.ie.Busy -eq $true) {
            Start-Sleep -Milliseconds $this.waitMilliseconds;
        }   
    }

    Close() {
        $this.ie.quit();
    }
}

class ConfluenceScraper {
    [IE]$ie;

    ConfluenceScraper([IE]$ie) {
        $this.ie = $ie
    }

    LoginAndWait($credential) {
        [string]$username = $credential.UserName
        [string]$password = $credential.GetNetworkCredential().Password

        $this.ie.NavigateAndWait("https://sife-net.htwchur.ch/login.action")

        
        # Login
        $usernamefield = $this.ie.getElementByID('os_username')
        $usernamefield.value = $username

        $passwordfield = $this.ie.getElementByID('os_password')
        $passwordfield.value = $password

        $Link=$this.ie.ClickAndWaitById("loginButton")
    }

    [string]BackupUri([string]$uri) {
        # Export
        $this.ie.NavigateAndWait($uri)

        # Need some more time to build the radio buttons
        Start-Sleep -Milliseconds 500;
        $this.ie.ClickRadio0ByTagNameAndName('input', 'format')
        $this.ie.ClickAndWaitByTagNameAndName("input", "confirm")

        $this.ie.ClickAndWaitByTagNameAndName("input", "confirm")

        $element = $this.ie.WaitByTagNameAndClass("a", 'space-export-download-path')
        $uri = $element.HREF

        return $uri
    }
  
    Close() {
        $this.ie.Close()    
    }
}

# Config
[string]$userName = 'martin.studer'
[string[]]$confluenceProjectShortcuts = @("PROR", "RW", "QIDLUW", "RL")
[string[]]$jiraProjectShortcuts = @("PROR", "RW", "QUAL", "RL")

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
    if ($usernamefield) {
        $usernamefield.value = $username

        $passwordfield = $ie.Document.getElementByID('os_password')
        $passwordfield.value = $password

        $Link=$ie.Document.getElementByID("loginButton")
        $Link.click()
        while ($ie.Busy -eq $true) {Start-Sleep -Milliseconds 100;}   
    }
    
    # Export

    $ie.Navigate("https://sife-net.htwchur.ch/spaces/exportspacewelcome.action?key=$projectShortcut")
    while ($ie.Busy -eq $true) {Start-Sleep -Milliseconds 100;}   

    Start-Sleep -Milliseconds 200;

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

    do {
        $i=$ie.Document.getElementsByTagName("a")

#TASK: The where clause fails if there is not attribute node with name 'class'. Check the existance of the attribute node first.

        $j = $i | where-object {$_.getAttributeNode('class').Value -eq 'space-export-download-path'}
        Start-Sleep -Milliseconds 100;
    } while (!($j))
    
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

function Save-PrivateConfluenceSpace($credential, [string]$outputDir) {
    [string]$userName = $credential.UserName
    [string]$outputFilename = "$userName.zip"
    [string]$outputFilepath = Join-Path -Path $outputDir -ChildPath $outputFilename

    $headers = Build-Headers($credential)

    [IE]$ie = New-Object IE
    [ConfluenceScraper] $scraper = New-Object ConfluenceScraper($ie)
    $scraper.LoginAndWait($credential)
    [string]$start = "https://sife-net.htwchur.ch/spaces/exportspacewelcome.action?key=~$userName"
    [string]$uri = $scraper.BackupUri($start)
    $scraper.Close()

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
    Write-Host "Backup jira projects:"

    [string]$outputPath = Join-Path -Path $path -ChildPath 'Jira'
    [string]$outputTimestampPath = Join-Path -Path $outputPath -ChildPath $timestamp
    mkdir $outputTimestampPath # https://superuser.com/questions/1153961/powershell-silent-mkdir/1154277exi

    foreach ($projectShortcut in $projectShortcuts) {
        Write-Host "Backup $projectShortcut..."
        JiraBackup -projectShortcut $projectShortcut -outputDir $outputTimestampPath -credential $credential
        Write-Host "done."
    }
    Write-Host "Backup jira projects done."
}

function Save-AllConfluenceBackups($projectShortcuts, $credential) {
    Write-Host "Backup confluence spaces:"
    [string]$outputPath = Join-Path -Path $path -ChildPath 'Confluence'
    [string]$outputPathTimestampPath = Join-Path -Path $outputPath -ChildPath $timestamp
    mkdir $outputPathTimestampPath > $null # https://superuser.com/questions/1153961/powershell-silent-mkdir/1154277
    foreach ($projectShortcut in $confluenceProjectShortcuts) {
        $uri = Get-ConfluenceDownloadUri -projectShortcut $projectShortcut -credential $confluenceCredential
        Write-Host "Backup $projectShortcut..."
        Save-ConfluenceBackup -projectShortcut $projectShortcut -outputDir $outputPathTimestampPath -uri $uri -credential $confluenceCredential
        Write-Host "done."
    }
    Save-PrivateConfluenceSpace -credential $confluenceCredential -outputDir $outputPathTimestampPath 
    Write-Host "Backup confluence spaces done."
}

$confluenceCredential = Get-Credential -Message 'Confluence' -UserName $userName
$jiraCredential = Get-Credential -Message 'Jira' -UserName $userName

Save-AllConfluenceBackups -projectShortcuts $projectShortcuts -credential $confluenceCredential
Save-AllJiraBackups -projectShortcuts $jiraProjectShortcuts -credential $jiraCredential