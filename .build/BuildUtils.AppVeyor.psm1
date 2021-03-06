Import-Module "$PSScriptRoot\GitReleaseInfo.psm1" -Verbose

function Read-GitReleaseInfoFromAppVeyor {
    [CmdletBinding()]
    $hasRepoTag = $env:APPVEYOR_REPO_TAG -eq "true"
    if ($hasRepoTag) {

        Write-Host "`nGit version tag detected: '$env:APPVEYOR_REPO_TAG_NAME'`n" -ForegroundColor Yellow
    
        if ($env:APPVEYOR_REPO_TAG_NAME.StartsWith("v")) {
            $version = $env:APPVEYOR_REPO_TAG_NAME.Substring(1)
        }
        else {
            $version = $env:APPVEYOR_REPO_TAG_NAME
        }
    
        $commitMessage = $env:APPVEYOR_REPO_COMMIT_MESSAGE
        if (-not [string]::IsNullOrWhiteSpace($env:APPVEYOR_REPO_COMMIT_MESSAGE_EXTENDED)) {
            $commitMessage = $commitMessage + "`n" + $env:APPVEYOR_REPO_COMMIT_MESSAGE_EXTENDED
        }

        Write-Verbose "`n`$env:APPVEYOR_REPO_COMMIT_MESSAGE ==> '$env:APPVEYOR_REPO_COMMIT_MESSAGE'"
        Write-Verbose "`n`$env:APPVEYOR_REPO_COMMIT_MESSAGE_EXTENDED ==> '$env:APPVEYOR_REPO_COMMIT_MESSAGE_EXTENDED'"

        $gitReleaseInfo = New-GitReleaseInfo -Version $version -CommitMessage $commitMessage
    }
    else {
        $gitReleaseInfo = New-EmptyGitReleaseInfo
    }

    return $gitReleaseInfo
}

Export-ModuleMember -Function Read-GitReleaseInfoFromAppVeyor