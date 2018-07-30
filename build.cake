#tool "nuget:?package=GitVersion.CommandLine"
#tool "nuget:?package=xunit.runner.console"

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

// Path to the csproj
string csProjPath = "./src/"+ directoryName + "/" + directoryName + ".csproj";
string nugetPackageSourcePath = "./src/"+ directoryName + "/bin/" + configuration + "/EPPlus.DataExtractor.*.nupkg";

// Path to tests project
string testsCsProjPath = "./src/"+ testsDirectoryName + "/" + testsDirectoryName + ".csproj";


GitVersion gitVersion = GitVersion();

//////////////////////////////////////////////////////////////////////
// TASKS
//////////////////////////////////////////////////////////////////////

var cleanTask = Task("Clean")
    .Does(() =>
{
    CleanDirectories("./src/**/bin/");
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
    MSBuild("./src/" + solutionName);
});

Task("Run-Unit-Tests")
    .IsDependentOn("Build")
    .Does(() =>
{
    DotNetCoreTest(testsCsProjPath, new DotNetCoreTestSettings
        {
            Configuration = "Release"
        });
});

Task("Pack")
    .IsDependentOn("Run-Unit-Tests")
    .Does(() =>
{
    cleanTask.Task.Execute(Context);

    // Build package
    MSBuild(csProjPath, settings =>
    {
        settings.SetConfiguration(configuration);
        settings.Targets.Clear();
        settings.Targets.Add("pack");
    });

    // Clean artifacts folder
    CleanDirectory("./artifacts/");

    Warning("Copying artifact: " + nugetPackageSourcePath);

    // Copy nuget package to artifacts folder
    CopyFiles(nugetPackageSourcePath, "./artifacts/");
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
