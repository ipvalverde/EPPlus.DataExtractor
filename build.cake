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
const string testsDirectoryName = "EPPlus.DataExtractor.Tests";

// Define directories.
var buildDir = Directory("./src/"+ directoryName + "/bin") + Directory(configuration);
var artifactsFolder = Directory("./artifacts");

// Path to the csproj
string csProjPath = "./src/"+ directoryName + "/" + directoryName + ".csproj";

// Path to the unit tests project dlls
string testsCsProjPath = "./src/"+ testsDirectoryName + "/" + testsDirectoryName + "/bin/" + configuration;

GitVersion gitVersion = GitVersion();

//////////////////////////////////////////////////////////////////////
// TASKS
//////////////////////////////////////////////////////////////////////

Task("Clean")
    .Does(() =>
{
    CleanDirectory(buildDir);
    CleanDirectory(artifactsFolder);
});

Task("UpdateCsProjVersion")
    .IsDependentOn("Clean")
    .Does(() =>
{
    var file = File(csProjPath);
    XmlPoke(file, "/Project/PropertyGroup[@Label='MainGroup']/PackageVersion",  gitVersion.NuGetVersion);
});

Task("Restore-NuGet-Packages")
    .IsDependentOn("UpdateCsProjVersion")
    .Does(() =>
{
    NuGetRestore("./src/" + solutionName);
});

Task("Build")
    .IsDependentOn("Restore-NuGet-Packages")
    .Does(() =>
{
    // Use MSBuild
    MSBuild("./src/" + solutionName, settings =>
    {
        settings.SetConfiguration(configuration);
        settings.Targets.Clear();
        settings.Targets.Add("pack");
    });
});

Task("Run-Unit-Tests")
    .IsDependentOn("Build")
    .Does(() =>
{
    // ToDo: Fix unit tests
    // var xunit = new XUnit2Runner();
    // xunit.Run(new FilePath[] { new FilePath(testsCsProjPath) }, new XUnit2Settings());
});

//////////////////////////////////////////////////////////////////////
// TASK TARGETS
//////////////////////////////////////////////////////////////////////

Task("Default")
    .IsDependentOn("Run-Unit-Tests");

//////////////////////////////////////////////////////////////////////
// EXECUTION
//////////////////////////////////////////////////////////////////////

RunTarget(target);
