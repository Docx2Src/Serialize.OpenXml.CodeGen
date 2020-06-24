var target = Argument("target", "Default");

/*****************************************************************************
 *                               Directories
 *****************************************************************************/

var artifactsDir = (DirectoryPath)Directory("./artifacts");
var zipRootDir = artifactsDir.Combine("zip");
var nugetRoot = artifactsDir.Combine("nuget");

var fileZipSuffix = ".zip";

var buildDirs = GetDirectories("./src/**/bin/**") + 
    GetDirectories("./src/**/obj/**") +     
    GetDirectories("./artifacts/**/zip/**");

var netCoreAppsRoot= "./src/";
var netCoreApps = new string[] { "Serialize.OpenXml.CodeGen" };
var netCoreProjects = netCoreApps.Select(name => 
    new {
        Path = string.Format("{0}/{1}", netCoreAppsRoot, name),
        Name = name,
        Framework = XmlPeek(string.Format("{0}/{1}/{1}.csproj", netCoreAppsRoot, name), "//*[local-name()='TargetFramework']/text()")
    }).ToList();

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
        DotNetCoreRestore(project.Path);
    }
});

Task("Build-NetCore")
    .IsDependentOn("Restore-NetCore")
    .Does(() =>
{
    foreach (var project in netCoreProjects)
    {
        Information("Building: {0}", project.Name);

        /*
        var settings = new DotNetCoreBuildSettings {
            Configuration = configuration
        };

        if (!IsRunningOnWindows())
        {
            settings.Framework = "netcoreapp2.1";
        }
        */

        DotNetCoreBuild(project.Path);
    }
});
Task("Default")
  .IsDependentOn("Restore-NetCore")
  .IsDependentOn("Build-NetCore");
  
RunTarget(target);
