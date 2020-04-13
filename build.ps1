Import-Module ".\.build\BuildUtils.psm1"
Import-Module ".\.build\BuildUtils.AppVeyor.psm1"

$rootFolder = (Get-Location).Path
$projectName = "EPPlus.DataExtractor"
$mainProjectPath = "src/EPPlus.DataExtractor/$projectName.csproj"
$nugetSourceUrl = "https://www.myget.org/F/pipelinenet/api/v2/package"
$mainProjectDirectory = "src/EPPlus.DataExtractor"
$solutionPath = ".\src\EPPlus.DataExtractor.sln"

$gitReleaseInfo = Read-GitReleaseInfoFromAppVeyor
Invoke-Build -GitReleaseInfo $gitReleaseInfo -SolutionPath $solutionPath -Verbose
Invoke-Tests -SolutionPath $solutionPath


# if (-not [string]::IsNullOrWhiteSpace($packageVersionCommandArgument)) {
#     Invoke-CommandWithLog -Command "dotnet pack $mainProjectPath --no-build -c Release --include-symbols -o artifacts -p:PackageReleaseNotes=`"$commitMessage`"$packageVersionCommandArgument" -CommandName "pack"

#     $nugetPackageName = "$projectName.$packageVersion"
#     Invoke-CommandWithLog -Command "dotnet nuget push $mainProjectDirectory/artifacts/$nugetPackageName.nupkg -s $nugetSourceUrl -k $env:MYGET_KEY" -CommandName "publish"
#     Invoke-CommandWithLog -Command "dotnet nuget push $mainProjectDirectory/artifacts/$nugetPackageName.symbols.nupkg -s $nugetSourceUrl -k $env:MYGET_KEY" -CommandName "publish symbols"
# }