class GitReleaseInfo {
    hidden [string] $Version
    hidden [string] $CommitMessage
    hidden [bool] $HiddenHasReleaseInfo

    [bool] HasReleaseInfo() {
        return $this.HiddenHasReleaseInfo
    }
    
    [string] GetVersion() {
        return $this.Version
    }

    [string] GetCommitMessage() {
        return $this.CommitMessage
    }

    GitReleaseInfo([string] $Version, [string] $commitMessage) {
        if (-not $Version) {
            throw "Version (`$Version) is required for GitReleaseInfo object"
        }
        if (-not $commitMessage) {
            throw "Commit message (`$commitMessage) is required for GitReleaseInfo object"
        }
        $this.Version = $Version
        $this.CommitMessage = $commitMessage
        $this.HiddenHasReleaseInfo = $true;
    }

    GitReleaseInfo() {
        $this.HiddenHasReleaseInfo = $false;
    }
}

function New-GitReleaseInfo() {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string] $Version,

        [Parameter(Mandatory=$true)]
        [string] $CommitMessage
    )
    [GitReleaseInfo]::New($Version, $CommitMessage)
}

function New-EmptyGitReleaseInfo() {
    [CmdletBinding()]
    param()
    
    [GitReleaseInfo]::New()
}

Export-ModuleMember -Function New-GitReleaseInfo, New-EmptyGitReleaseInfo