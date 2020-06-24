/*****************************************************************************
 *                                 Usings
 *****************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

/*****************************************************************************
 *                                Arguments
 *****************************************************************************/

var target = Argument("target", "Default");
var platform = Argument("platform", "AnyCPU");
var configuration = Argument("configuration", "Debug");

/*****************************************************************************
 *                              Configuration
 *****************************************************************************/

var ProjectName = "Serialize.OpenXml.CodeGen";
var MainRepo = "rmboggs/Serialize.OpenXml.CodeGen";
var MasterBranch = "master";
var ReleasePlatform = "Any CPU";
var ReleaseConfiguration = "Release";

/*****************************************************************************
 *                               Parameters
 *****************************************************************************/

var isPlatformAnyCPU = StringComparer.OrdinalIgnoreCase.Equals(platform, "Any CPU");
var isPlatformX86 = StringComparer.OrdinalIgnoreCase.Equals(platform, "x86");
var isPlatformX64 = StringComparer.OrdinalIgnoreCase.Equals(platform, "x64");
var isLocalBuild = BuildSystem.IsLocalBuild;
var isRunningOnUnix = IsRunningOnUnix();
var isRunningOnWindows = IsRunningOnWindows();
var isRunningOnAppVeyor = BuildSystem.AppVeyor.IsRunningOnAppVeyor;
var isPullRequest = BuildSystem.AppVeyor.Environment.PullRequest.IsPullRequest;
var isMainRepo = StringComparer.OrdinalIgnoreCase.Equals(MainRepo, BuildSystem.AppVeyor.Environment.Repository.Name);
var isMasterBranch = StringComparer.OrdinalIgnoreCase.Equals(MasterBranch, BuildSystem.AppVeyor.Environment.Repository.Branch);
var isTagged = BuildSystem.AppVeyor.Environment.Repository.Tag.IsTag 
               && !string.IsNullOrWhiteSpace(BuildSystem.AppVeyor.Environment.Repository.Tag.Name);
var isReleasable = StringComparer.OrdinalIgnoreCase.Equals(ReleasePlatform, platform) 
                   && StringComparer.OrdinalIgnoreCase.Equals(ReleaseConfiguration, configuration);
var isMyGetRelease = !isTagged && isReleasable;
var isNuGetRelease = isTagged && isReleasable;

/*****************************************************************************
 *                               Directories
 *****************************************************************************/

var artifactsDir = (DirectoryPath)Directory("./artifacts");
var zipRootDir = artifactsDir.Combine("zip");
var nugetRoot = artifactsDir.Combine("nuget");

var fileZipSuffix = ".zip";

var buildDirs = GetDirectories("./src/**/bin/**") + 
    GetDirectories("./src/**/obj/**") +     
    GetDirectories("./artifacts/**/zip/**") +
    GetDirectories(".artifacts/**/nuget/**");

var netCoreAppsRoot= "./src/";
var netCoreApp = String.Concat(netCoreAppsRoot, ProjectName);
var netCoreProjects = new Dictionary<string, string>();

netCoreProjects.Add(ProjectName, netCoreApp);

/*****************************************************************************
 *                                 Details
 *****************************************************************************/
var version = XmlPeek(String.Format("{0}/{1}.csproj", netCoreApp, ProjectName),
                        "//*[local-name()='Version']/text()");

Information("Building version {0} of Serialize.OpenXml.CodeGen ({1}, {2}, {3}) using version {4} of Cake.", 
    version,
    platform,
    configuration,
    target,
    typeof(ICakeContext).Assembly.GetName().Version.ToString());

if (isRunningOnAppVeyor)
{
    Information("Repository Name: " + BuildSystem.AppVeyor.Environment.Repository.Name);
    Information("Repository Branch: " + BuildSystem.AppVeyor.Environment.Repository.Branch);
}

Information("Target: " + target);
Information("Platform: " + platform);
Information("Configuration: " + configuration);
Information("IsLocalBuild: " + isLocalBuild);
Information("IsRunningOnUnix: " + isRunningOnUnix);
Information("IsRunningOnWindows: " + isRunningOnWindows);
Information("IsRunningOnAppVeyor: " + isRunningOnAppVeyor);
Information("IsPullRequest: " + isPullRequest);
Information("IsMainRepo: " + isMainRepo);
Information("IsMasterBranch: " + isMasterBranch);
Information("IsTagged: " + isTagged);
Information("IsReleasable: " + isReleasable);
Information("IsMyGetRelease: " + isMyGetRelease);
Information("IsNuGetRelease: " + isNuGetRelease);

/*****************************************************************************
 *                                 Tasks
 *****************************************************************************/

Task("Clean")
.Does(()=>{
    CleanDirectories(buildDirs);
    CleanDirectory(nugetRoot);
    CleanDirectory(zipRootDir);
});

Task("Restore-NetCore")
    .IsDependentOn("Clean")
    .Does(() =>
    {
        foreach (var project in netCoreProjects)
        {
            DotNetCoreRestore(project.Value);
        }
    });

Task("Build-NetCore")
    .IsDependentOn("Restore-NetCore")
    .Does(() =>
    {
        foreach (var project in netCoreProjects)
        {
            Information("Building: {0}", project.Key);
            DotNetCoreBuild(project.Value);
        }
    });

Task("Pack-NetCore")
    .IsDependentOn("Build-NetCore")
    .Does(() =>
    {
        foreach (var project in netCoreProjects)
        {
            Information("Packing: {0}", project.Key);
            DotNetCorePack(project.Value);
        }
    });

Task("Zip-NetCore")
    .IsDependentOn("Build-NetCore")
    .Does(() =>
    {
        foreach (var project in netCoreProjects)
        {
            Information("Zipping: {0}", project.Key);
            var netCoreAppBin = String.Concat(project.Value, "/bin/", configuration);
            var netCoreAppBinFiles = String.Concat(netCoreAppBin, "/**/*");
            //*
            Zip(netCoreAppBin, 
                zipRootDir.CombineWithFilePath(String.Concat(project.Key, ".", version, fileZipSuffix)),
                GetFiles(netCoreAppBinFiles));
            //*/
        }
    });

Task("Default")
  .IsDependentOn("Restore-NetCore")
  .IsDependentOn("Build-NetCore")
  .IsDependentOn("Pack-NetCore")
  .IsDependentOn("Zip-NetCore");
  
RunTarget(target);
