# Version 1.0

# functions
function Initialize-ColorScheme
{
    $script:successColor = "Green"
    $script:infoColor = "DarkCyan"
    $script:warningColor = "Yellow"
    $script:failColor = "Red"    
}

function Show-Introduction
{
    Write-Host "This script exports all deleted contacts in FreshDesk into a CSV." -ForegroundColor $infoColor
    Read-Host "Press Enter to continue"
}

function Get-EncodedApiKey
{
    $secureString = Read-Host "Enter your API key" -AsSecureString
    $psCredential = Convert-SecureStringToPsCredential $secureString
    # Append :X because FreshDesk expects that. Could be X or anything else.
    return ConvertTo-Base64 ($psCredential.GetNetworkCredential().Password + ":X")    
}

function Convert-SecureStringToPsCredential($secureString)
{
    # Just passing "null" for username, because username will not be used.
    return New-Object System.Management.Automation.PSCredential("null", $secureString)
}

function ConvertTo-Base64($text)
{
    return [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($text))
}

function Test-APIConnection($freshDeskUrl, $encodedApiKey)
{
    $myProfileUrl = "$freshDeskUrl/api/v2/agents/me"
    $headers = @{
        Authorization = "Basic $encodedApiKey"      
    }

    try
    {
        Invoke-RestMethod -Method "Get" -Uri $myProfileUrl -Headers $headers -ErrorVariable "responseError" | Out-Null
        $connectionSuccess = $true
        Write-Host "Successfully connected to the FreshDesk API!" -ForegroundColor $successColor
    }
    catch
    {
        Write-Warning "API request for your profile returned an error:`n$($responseError[0].Message)"

        $responseCode = [int]$_.Exception.Response.StatusCode
        if (($responseCode -eq 401) -or ($responseCode -eq 403))
        {
            Write-Warning "API key invalid or lacks permissions."
        }
        else
        {
            # If the error can't be resolved by fixing the API key, we'll have to exit.
            exit
        }
        $connectionSuccess = $false
    }    
    return $connectionSuccess
}

function Get-DeletedContacts($freshDeskUrl, $encodedApiKey)
{
    Write-Host "Getting deleted contacts..." -ForegroundColor $infoColor
    
    $headers = @{
        Authorization = "Basic $encodedApiKey"
    }

    $results = New-Object -TypeName System.Collections.Generic.List[PSObject]
    $page = 1
    do
    {
        Write-Progress -Activity "Getting deleted contacts (100 contacts per page)..." -Status "$($page - 1) pages retrieved"
        $response = SafelyInvoke-WebRequest -Method "Get" -Uri "$freshDeskUrl/api/v2/contacts?state=deleted&per_page=100&page=$page" -Headers $headers
        $result = ConvertFrom-Json -InputObject $response.Content
        $results.Add($result)
        $page++
        Write-Progress -Activity "Getting deleted contacts..." -Status "$($page - 1) pages retrieved"
    }
    while ($response.Headers.Link) # When there is another page, there will be a link in the response headers to the next page.
    
    return $results
}

function SafelyInvoke-WebRequest($method, $uri, $headers, $body)
{
    try
    {
        $response = Invoke-WebRequest -Method $method -Uri $uri -Headers $headers -Body $body -ErrorVariable "responseError"
    }
    catch
    {
        Write-Host $responseError[0].Message -ForegroundColor $failColor
        exit
    }

    return $response
}

function Export-DeletedContacts($deletedContacts)
{
    Write-Host "Exporting contacts to CSV..." -ForegroundColor $infoColor

    $timeStamp = New-TimeStamp
    $path = "$PSScriptRoot\FreshDesk Deleted Contacts $timeStamp.csv"

    $contactsProcessed = 0
    foreach($array in $deletedContacts)
    {
        foreach($contact in $array)
        {
            Export-Csv -InputObject $contact -Path $path -Append -Force -NoTypeInformation
            $contactsProcessed++
            Write-Progress -Activity "Exporting contacts..." -Status "$contactsProcessed contacts processed"
        }
    }
    Write-Host "Finished exporting to $path" -ForegroundColor $successColor
}

function New-TimeStamp
{
    return (Get-Date -Format yyyy-MM-dd-hh-mm).ToString()
}

# main
Initialize-ColorScheme
Show-Introduction
$freshDeskUrl = "https://blueravensolar.freshdesk.com"
do
{
    $encodedApiKey = Get-EncodedApiKey
    $connectionSuccess = Test-APIConnection -FreshDeskUrl $freshDeskUrl -EncodedApiKey $encodedApiKey
}
while (-not($connectionSuccess))
$deletedContacts = Get-DeletedContacts -FreshDeskUrl $freshDeskUrl -EncodedApiKey $encodedApiKey
Export-DeletedContacts $deletedContacts
Read-Host "Press Enter to exit"