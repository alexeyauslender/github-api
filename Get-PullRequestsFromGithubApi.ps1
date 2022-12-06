param(
    [Parameter(Mandatory)]$To,
    [string]$Owner = "microsoft",
    [string]$RepoName = "vscode",
    [string]$HistoryInDays = 7,
    [string]$PAT
)

function Get-GitHubPullRequest
{
    param(
        [ValidateSet('Open', 'Closed', 'All')]
        [string] $State = 'All',
        [Parameter()]$Owner = "microsoft",
        [Parameter()]$RepoName = "vscode",
        [int]$HistoryInDays = 7,
        [int]$MaxPageNumberToFetch = 15,
        $PAT
    )
    $WebRequestArgs = @{ 
        Headers = @{ 
            Accept = "application/vnd.github.v3+json" 
        }
    }
    if ($PAT) 
    {
        $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$PAT"))
        $WebRequestArgs["Headers"].Add("Authorization", "Basic $base64AuthInfo") 
    }

    $nextLink = "https://api.github.com/repos/$Owner/$RepoName/pulls?state=$State&page=1"
    do
    {
        $result = Invoke-WebRequest -Uri $nextLink @WebRequestArgs
        $r += $result.Content | ConvertFrom-Json 
        if ($result.Headers.Count -gt 0)
        {
            $links = $result.Headers['Link'] -split ','
            foreach ($link in $links)
            {
                if ($link -match '<(.*page=(\d+)[^\d]*)>; rel="next"')
                {
                    $nextLink = $Matches[1]
                    $nextPageNumber = [int]$Matches[2]
                }
                elseif ($link -match '<.*page=(\d+)[^\d]+rel="last"')
                {
                    $numPages = [int]$Matches[1]
                    if ($numPages -gt $MaxPageNumberToFetch)
                    {
                        $numPages = $MaxPageNumberToFetch
                    }
                }
            }
        }
    } while ($nextPageNumber -lt $numPages)
    $r | Where-Object {
        $d = Get-Date -Date $_.updated_at 
        ([datetime]::UtcNow - $d).TotalDays -lt $historyInDays 
    }
}

function Invoke-SendReportFromLocalOutlookOrShowInConsole
{
    param(
        [Parameter(Mandatory)]$To,
        [Parameter(Mandatory)]$Body
    )
    try
    {
        $Outlook = New-Object -ComObject Outlook.Application
        $TlabEmail = $Outlook.CreateItem(0)
        $TlabEmail.To = $To
        $TlabEmail.Subject = "Your pull requests report from last 7 days."
        $TlabEmail.Body = $Body
        $TlabEmail.save()
        $inspector = $TlabEmail.GetInspector
        $inspector.Display() 
    }
    catch
    {
        Write-Host "Failed to open outlook with the following error $($_.Exception.Message). Falls back to the console."
        Write-Host $Body
    }  
}

function Invoke-ReportAsSingleString
{
    param(
        [array][Parameter(Mandatory)]$PullRquestStates,
        [string][Parameter(Mandatory)]$Owner,
        [string][Parameter(Mandatory)]$RepoName,
        [string]$PAT,
        [int]$HistoryInDays
    )
    $r = ""
    $PullRquestStates | ForEach-Object {
        $r += "-------------------------------------------------------`n"
        $r += "Your $_ PRs for the last 7 days `n"
        $r += "-------------------------------------------------------`n`n"
        $localCurrentItem = $_
        if ($localCurrentItem -eq "draft")
        {
            $localCurrentItem = "All"
        }
        $filteredByState = Get-GitHubPullRequest -State $localCurrentItem -Owner $Owner -RepoName $RepoName -PAT $PAT -HistoryInDays $HistoryInDays
        if ($localCurrentItem -eq "All")
        {
            $filteredByState = $filteredByState | Where-Object {
                $_.draft
            }
        }
        $r += $filteredByState | ForEach-Object {
            "Title: $($_.title)`n Author: $($_.user.login)`n Url: $($_.url) `n"
        }
        $r += "`n"
    }
    $r

}

$emailBody = Invoke-ReportAsSingleString -PullRquestStates @("open", "closed", "draft") -Owner $Owner -RepoName $RepoName -PAT $PAT -HistoryInDays $HistoryInDays

Invoke-SendReportFromLocalOutlookOrShowInConsole -To $To -Body $emailBody



