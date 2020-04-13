class GitReleaseInfo {
    hidden [string] $PackageVersion
    hidden [string] $CommitMessage
    hidden [bool] $HiddenHasReleaseInfo

    [bool] HasReleaseInfo() {
        return $this.HiddenHasReleaseInfo
    }
    
    [string] GetPackageVersion() {
        return $this.PackageVersion
    }

    [string] GetCommitMessage() {
        return $this.CommitMessage
    }

    GitReleaseInfo([string] $packageVersion, [string] $commitMessage) {
        if (-not $packageVersion) {
            throw "Package version (`$packageVersion) is required for GitReleaseInfo object"
        }
        if (-not $commitMessage) {
            throw "Commit message (`$commitMessage) is required for GitReleaseInfo object"
        }
        $this.PackageVersion = $packageVersion
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
        [string] $PackageVersion,

        [Parameter(Mandatory=$true)]
        [string] $CommitMessage
    )
    [GitReleaseInfo]::New($PackageVersion, $CommitMessage)
}

function New-EmptyGitReleaseInfo() {
    [CmdletBinding()]
    param()
    
    [GitReleaseInfo]::New()
}

Export-ModuleMember -Function New-GitReleaseInfo, New-EmptyGitReleaseInfo