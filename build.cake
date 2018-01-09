#tool "nuget:?package=GitVersion.CommandLine"

//////////////////////////////////////////////////////////////////////
// ARGUMENTS
//////////////////////////////////////////////////////////////////////

var target = Argument("target", "Default");
var configuration = Argument("configuration", "Release");

//////////////////////////////////////////////////////////////////////
// PREPARATION
//////////////////////////////////////////////////////////////////////

const string directoryName = "EPPlus.DataExtractor";
const string solutionName = "EPPlus.DataExtractor.sln";

// Define directories.
var buildDir = Directory("./src/"+ directoryName + "/bin") + Directory(configuration);
var artifactsFolder = Directory("./artifacts");

//////////////////////////////////////////////////////////////////////
// TASKS
//////////////////////////////////////////////////////////////////////

Task("Clean")
    .Does(() =>
{
    CleanDirectory(buildDir);
    CleanDirectory(artifactsFolder);
});

Task("Restore-NuGet-Packages")
    .IsDependentOn("Clean")
    .Does(() =>
{
    NuGetRestore("./src/" + solutionName);
});

Task("Build")
    .IsDependentOn("Restore-NuGet-Packages")
    .Does(() =>
{
    if(IsRunningOnWindows())
    {
      // Use MSBuild
      MSBuild("./src/" + solutionName, settings =>
      {
        settings.SetConfiguration(configuration);
      });
    }
    else
    {
      // Use XBuild
      XBuild("./src/" + solutionName, settings =>
        settings.SetConfiguration(configuration));
    }
});

Task("Run-Unit-Tests")
    .IsDependentOn("Build")
    .Does(() =>
{
    MSTest("./src/**/bin/" + configuration + "/*.Tests.dll");
});


Task("Pack")
    .IsDependentOn("Run-Unit-Tests")
    .Does(() =>
{
    string nuspecPath = "./src/"+ directoryName + "/" + directoryName + ".nuspec";

    GitVersion gitVersion = GitVersion();

    Warning("Branch detected: " + gitVersion.BranchName);
    Warning("Nuget version detected: " + gitVersion.NuGetVersion);

    var packageSettings = new NuGetPackSettings
        {
            OutputDirectory = artifactsFolder,
            Version = gitVersion.NuGetVersion,
        };

    NuGetPack(nuspecPath, packageSettings);
});

//////////////////////////////////////////////////////////////////////
// TASK TARGETS
//////////////////////////////////////////////////////////////////////

Task("Default")
    .IsDependentOn("Pack");

//////////////////////////////////////////////////////////////////////
// EXECUTION
//////////////////////////////////////////////////////////////////////

RunTarget(target);
