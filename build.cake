#tool "nuget:?package=NuGet.CommandLine&version=5.8.1"
#addin "nuget:?package=Cake.MinVer&version=0.2.0"

using System.Net.Http;

var target          = Argument<string>("target", "pack");
var buildVersion    = MinVer(s => s.WithTagPrefix("v").WithDefaultPreReleasePhase("preview"));

var excelDnaPackageId = "ExcelDna.AddIn";
var excelDnaPackageVersion = "1.1.1";
var excelDnaCopyrightYear = 2020;

var excelDnaPackageDirectoryPath = MakeAbsolute(new DirectoryPath($"./build/source/{excelDnaPackageId}.{excelDnaPackageVersion}"));
var excelDnaPackageFilePath = excelDnaPackageDirectoryPath.CombineWithFilePath($"{excelDnaPackageId}.{excelDnaPackageVersion}.nupkg");

Task("clean")
    .Does(() =>
{
    CleanDirectory("./build/source");
    CleanDirectory("./build/artifacts");
});

Task("download-nupkg")
    .IsDependentOn("clean")
    .Does(() =>
{
    if (FileExists(excelDnaPackageFilePath))
    {
        Information("NuGet package {0} v{1} already exists on:", excelDnaPackageFilePath, excelDnaPackageVersion);
        Information("{0}", excelDnaPackageFilePath);
        return;
    }

    Information("Downloading NuGet package {0} v{1}...", excelDnaPackageId, excelDnaPackageVersion);

    EnsureDirectoryExists(excelDnaPackageDirectoryPath);

    var packageIdLower = excelDnaPackageId.ToLowerInvariant();
    var packageVersionLower = excelDnaPackageVersion.ToLowerInvariant();
    var packageDownloadUrl = $"https://api.nuget.org/v3-flatcontainer/{packageIdLower}/{packageVersionLower}/{packageIdLower}.{packageVersionLower}.nupkg";

    Verbose("GET {0}", packageDownloadUrl);
    DownloadFile(packageDownloadUrl, excelDnaPackageFilePath);
});

Task("expand-nupkg")
    .IsDependentOn("download-nupkg")
    .Does(context =>
{
    var filesToExtract = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        { "ExcelDna.Integration.dll" },
        { "ExcelDnaPack.exe" },
        { "ExcelDnaPack.exe.config" },
    };

    using (var stream = System.IO.File.OpenRead(excelDnaPackageFilePath.FullPath))
    using (var zipStream = new System.IO.Compression.ZipArchive(stream))
    {
        var extractedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var entry in zipStream.Entries)
        {
            var packageFileName = excelDnaPackageFilePath.GetFilename();

            if (!entry.FullName.StartsWith("tools") || extractedFiles.Contains(entry.Name) || !filesToExtract.Contains(entry.Name))
            {
                Verbose("Skipping file {0} in {1}", entry.FullName, packageFileName);
                continue;
            }

            Information("Extracting file {0} from {1}", entry.FullName, packageFileName);

            var entryFilePath = excelDnaPackageDirectoryPath.CombineWithFilePath(entry.FullName);
            var entryDirectoryPath = entryFilePath.GetDirectory();

            EnsureDirectoryExists(entryDirectoryPath);

            using (var source = entry.Open())
            using (var target = context.FileSystem.GetFile(entryFilePath).OpenWrite())
            {
                source.CopyTo(target);
            }

            extractedFiles.Add(entry.Name);
        }
    }
});

Task("prepare-license")
    .IsDependentOn("expand-nupkg")
    .Does(() =>
{
    Information("Preparing LICENSE.txt (Copyright year {0})", excelDnaCopyrightYear);

    var licenseSrcFilePath = MakeAbsolute(new DirectoryPath("./src/Zlib-license-template.txt"));
    Verbose("License template: {0}", licenseSrcFilePath.FullPath);

    var licenseText = System.IO.File.ReadAllText(licenseSrcFilePath.FullPath)
        .Replace("{CopyrightYear}", excelDnaCopyrightYear.ToString());

    var licenseDestFilePath = excelDnaPackageDirectoryPath.CombineWithFilePath("LICENSE.txt");
    Verbose("NuGet package license: {0}", licenseDestFilePath.FullPath);

    System.IO.File.WriteAllText(licenseDestFilePath.FullPath, licenseText);
});

Task("prepare-nuspec")
    .IsDependentOn("prepare-license")
    .Does(context =>
{
    Information("Preparing NuGet package spec");

    Verbose("Calculating file hashes");
    var excelDnaPackExeFilePath = excelDnaPackageDirectoryPath.CombineWithFilePath("tools/ExcelDnaPack.exe");
    var excelDnaPackExeMd5Hash = context.CalculateFileHash(excelDnaPackExeFilePath, HashAlgorithm.MD5);
    var excelDnaPackExeSha256Hash = context.CalculateFileHash(excelDnaPackExeFilePath, HashAlgorithm.SHA256);
    var excelDnaPackExeSha512Hash = context.CalculateFileHash(excelDnaPackExeFilePath, HashAlgorithm.SHA512);

    var excelDnaIntegrationDllFilePath = excelDnaPackageDirectoryPath.CombineWithFilePath("tools/ExcelDna.Integration.dll");
    var excelDnaIntegrationDllMd5Hash = context.CalculateFileHash(excelDnaIntegrationDllFilePath, HashAlgorithm.MD5);
    var excelDnaIntegrationDllSha256Hash = context.CalculateFileHash(excelDnaIntegrationDllFilePath, HashAlgorithm.SHA256);
    var excelDnaIntegrationDllSha512Hash = context.CalculateFileHash(excelDnaIntegrationDllFilePath, HashAlgorithm.SHA512);

    var nuspecSrcFilePath = MakeAbsolute(new DirectoryPath("./src/ExcelDnaPack.nuspec"));
    Verbose("NuGet package spec template: {0}", nuspecSrcFilePath.FullPath);

    var nuspecText = System.IO.File.ReadAllText(nuspecSrcFilePath.FullPath)
        .Replace("{CopyrightYear}", excelDnaCopyrightYear.ToString())
        .Replace("{ExcelDnaPackVersion}", excelDnaPackageVersion)
        .Replace("{ExcelDnaPackExeMd5Hash}", excelDnaPackExeMd5Hash.ToHex())
        .Replace("{ExcelDnaPackExeSha256Hash}", excelDnaPackExeSha256Hash.ToHex())
        .Replace("{ExcelDnaPackExeSha512Hash}", excelDnaPackExeSha512Hash.ToHex())
        .Replace("{ExcelDnaIntegrationDllMd5Hash}", excelDnaIntegrationDllMd5Hash.ToHex())
        .Replace("{ExcelDnaIntegrationDllSha256Hash}", excelDnaIntegrationDllSha256Hash.ToHex())
        .Replace("{ExcelDnaIntegrationDllSha512Hash}", excelDnaIntegrationDllSha512Hash.ToHex())
    ;

    var nuspecDestFilePath = excelDnaPackageDirectoryPath.CombineWithFilePath("ExcelDnaPack.nuspec");
    Verbose("NuGet package spec: {0}", excelDnaPackageId, nuspecDestFilePath.FullPath);

    System.IO.File.WriteAllText(nuspecDestFilePath.FullPath, nuspecText);
});

Task("pack")
    .IsDependentOn("prepare-nuspec")
    .Does(() =>
{
    Information("Creating NuGet package");

    var nuspecDestFilePath = excelDnaPackageDirectoryPath.CombineWithFilePath("ExcelDnaPack.nuspec");

    NuGetPack(nuspecDestFilePath.FullPath, new NuGetPackSettings
    {
        Version = buildVersion.PackageVersion,
        OutputDirectory = "./build/artifacts",
    });
});

Task("publish")
    .IsDependentOn("pack")
    .Does(context =>
{
    Information("Publishing NuGet package");

    var url =  context.EnvironmentVariable("NUGET_URL");
    if (string.IsNullOrWhiteSpace(url))
    {
        context.Information("No NuGet URL specified. Skipping publishing of NuGet packages");
        return;
    }

    var apiKey =  context.EnvironmentVariable("NUGET_API_KEY");
    if (string.IsNullOrWhiteSpace(apiKey))
    {
        context.Information("No NuGet API key specified. Skipping publishing of NuGet packages");
        return;
    }

    var nugetPushSettings = new DotNetCoreNuGetPushSettings
    {
        Source = url,
        ApiKey = apiKey,
    };

    foreach (var nugetPackageFile in GetFiles("./build/artifacts/*.nupkg"))
    {
        DotNetCoreNuGetPush(nugetPackageFile.FullPath, nugetPushSettings);
    }
});

RunTarget(target);
