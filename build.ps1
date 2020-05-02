Import-Module ".\.build\BuildUtils.psm1"
Import-Module ".\.build\BuildUtils.AppVeyor.psm1"

$projectName = "EPPlus.DataExtractor"
$mainProjectPath = "src/EPPlus.DataExtractor/$projectName.csproj"
$nugetSourceUrl = "https://www.myget.org/F/pipelinenet/api/v2/package"
$solutionPath = ".\src\EPPlus.DataExtractor.sln"

$gitReleaseInfo = Read-GitReleaseInfoFromAppVeyor
Invoke-Build -GitReleaseInfo $gitReleaseInfo -SolutionPath $solutionPath -Verbose
Invoke-Tests -SolutionPath $solutionPath

Publish-Package -GitReleaseInfo $gitReleaseInfo `
    -ProjectPath $mainProjectPath `
    -NugetSourceUrl $nugetSourceUrl `
    -ApiKey $env:MYGET_KEY