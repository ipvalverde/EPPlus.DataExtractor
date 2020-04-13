Import-Module "$PSScriptRoot\GitReleaseInfo.psm1" -Verbose

function Invoke-CommandWithLog {
    [CmdletBinding()]
    param([string] $Command, [string] $CommandName)
    Write-Output "`n-----------------------------------------"
    Write-Output "Starting $CommandName process"
    Write-Output "`n-----------------------------------------"
    Write-Verbose "Command to be executed: '$Command'"
    Invoke-Expression $Command
    Write-Output "`n-----------------------------------------"
    Write-Output "$CommandName finished"
    Write-Output "`n-----------------------------------------"
}

function Invoke-Build {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [object] $GitReleaseInfo,

        [Parameter(Mandatory=$true)]
        [string] $SolutionPath
    )

    if ($GitReleaseInfo.HasReleaseInfo()) {
        $packageVersion = $GitReleaseInfo.GetPackageVersion()
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
        [Parameter(Mandatory=$true)]
        [string] $SolutionPath
    )

    Invoke-CommandWithLog -Command "dotnet test $SolutionPath -c Release --no-build" -CommandName "test"
}

Export-ModuleMember -Function Invoke-Build, Invoke-Tests


# if (-not [string]::IsNullOrWhiteSpace($packageVersionCommandArgument)) {
#     Invoke-CommandWithLog -Command "dotnet pack $mainProjectPath --no-build -c Release --include-symbols -o artifacts -p:PackageReleaseNotes=`"$commitMessage`"$packageVersionCommandArgument" -CommandName "pack"

#     $nugetPackageName = "$projectName.$packageVersion"
#     Invoke-CommandWithLog -Command "dotnet nuget push $mainProjectDirectory/artifacts/$nugetPackageName.nupkg -s $nugetSourceUrl -k $env:MYGET_KEY" -CommandName "publish"
#     Invoke-CommandWithLog -Command "dotnet nuget push $mainProjectDirectory/artifacts/$nugetPackageName.symbols.nupkg -s $nugetSourceUrl -k $env:MYGET_KEY" -CommandName "publish symbols"
# }