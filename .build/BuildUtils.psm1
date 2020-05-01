Import-Module "$PSScriptRoot\GitReleaseInfo.psm1" -Verbose

function Invoke-CommandWithLog {
    [CmdletBinding()]
    param([string] $Command, [string] $CommandName)
    Write-Host "`n-----------------------------------------"
    Write-Host "Starting $CommandName process"
    Write-Host "`n-----------------------------------------"
    Invoke-Expression $Command
    Write-Host "`n-----------------------------------------"
    Write-Host "$CommandName finished"
    Write-Host "`n-----------------------------------------"
}

function Invoke-Build {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object] $GitReleaseInfo,

        [Parameter(Mandatory = $true)]
        [string] $SolutionPath
    )

    if ($GitReleaseInfo.HasReleaseInfo()) {
        $packageVersion = $GitReleaseInfo.GetVersion()
        $packageVersionCommandArgument = " -p:Version=$packageVersion"
    }
    else {
        $packageVersionCommandArgument = ""
    }

    Invoke-CommandWithLog -Command "dotnet build $SolutionPath -c Release$packageVersionCommandArgument" -CommandName "build"
}

function Invoke-Tests {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $SolutionPath
    )

    Invoke-CommandWithLog -Command "dotnet test $SolutionPath -c Release --no-build" -CommandName "test"
}

function Publish-Package {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object] $GitReleaseInfo,

        [Parameter(Mandatory = $true)]
        [string] $ProjectPath,

        [Parameter(Mandatory = $true)]
        [string] $NugetSourceUrl,

        [Parameter(Mandatory = $true)]
        [string] $ApiKey,

        [Parameter()]
        [string] $ProjectName = $null
    )

    if ($GitReleaseInfo.HasReleaseInfo()) {

        $currentPath = Get-Location
        $version = $GitReleaseInfo.GetVersion()
        $commitMessage = $GitReleaseInfo.GetCommitMessage()
        Invoke-CommandWithLog -Command "dotnet pack $ProjectPath --no-build -c Release --include-symbols -o artifacts -p:PackageReleaseNotes=`"$commitMessage`" -p:PackageVersion=$version" -CommandName "pack"

        if (-not $ProjectName) {
            $csProj = Get-Item $ProjectPath
            if (-not $csProj) {
                throw "No file found for the given ProjectPath: '$ProjectPath'"
            }
            $ProjectName = $csProj.BaseName
        }

        $nugetPackageName = "$ProjectName.$version"
        Invoke-CommandWithLog -Command "dotnet nuget push $currentPath/artifacts/$nugetPackageName.nupkg -s $NugetSourceUrl -k $ApiKey" -CommandName "publish"
        Invoke-CommandWithLog -Command "dotnet nuget push $currentPath/artifacts/$nugetPackageName.symbols.nupkg -s $NugetSourceUrl -k $ApiKey" -CommandName "publish symbols"
    }
    else {
        Write-Host "Publishing step skipped. No release info provided" -ForegroundColor Yellow
    }
}

Export-ModuleMember -Function Invoke-Build, Invoke-Tests, Publish-Package