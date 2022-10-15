| README.md |
|:---|

<div align="center">

<img src="asset/exceldnapack.png" alt="ExcelDnaPack - NuGet package" />

</div>

<h1 align="center">ExcelDnaPack - NuGet package</h1>
<div align="center">

ExcelDnaPack distributed as a NuGet package for use as a tool in build scenarios.

[![NuGet Version](http://img.shields.io/nuget/v/ExcelDnaPack.svg?style=flat-square)](https://www.nuget.org/packages/ExcelDnaPack/) [![Stack Overflow](https://img.shields.io/badge/stack%20overflow-excel--dna-orange.svg?style=flat-square)](http://stackoverflow.com/questions/tagged/excel-dna)

</div>

## Give a Star! :star:

If you like or are using this project please give it a star. Thanks!

## Getting started :rocket:

To use the ExcelDnaPack tool in your build process, download the latest version of the [`ExcelDnaPack`](https://www.nuget.org/packages/ExcelDnaPack/) NuGet package from [nuget.org](https://www.nuget.org/packages/ExcelDnaPack/) or from the [Releases tab](https://github.com/augustoproiete/ExcelDnaPack-NuGet/releases) on GitHub, expand it to disk and run the executable under `tools/ExcelDnaPack.exe` with the arguments documented in the [usage](https://github.com/augustoproiete/ExcelDnaPack-NuGet#exceldnapack-usage) section.

## ExcelDnaPack usage

ExcelDnaPack is a command-line utility to pack ExcelDna add-ins into a single .xll file.

```
Usage: ExcelDnaPack.exe dnaPath [/O outputPath] [/Y]

  dnaPath           The path to the primary .dna file for the ExcelDna add-in.
  /Y                If the output .xll exists, overwrite without prompting.
  /NoCompression    no compress (LZMA) of resources
  /NoMultiThreading no multi threading to ensure deterministic packing order
  /O outPath        Output path - default is <dnaPath>-packed.xll.

Example: ExcelDnaPack.exe MyAddins\FirstAddin.dna
                 The packed add-in file will be created as MyAddins\FirstAddin-packed.xll.

The template add-in host file (the copy of ExcelDna.xll renamed to FirstAddin.xll) is
searched for in the same directory as FirstAddin.dna.

The Excel-Dna integration assembly (ExcelDna.Integration.dll) is searched for
  1. in the same directory as the .dna file, and if not found there,
  2. in the same directory as the ExcelDnaPack.exe file.

ExcelDnaPack will also pack the configuration file FirstAddin.xll.config if it is
found next to FirstAddin.dna.
Other assemblies are packed if marked with Pack="true" in the .dna file.
```

## Using ExcelDnaPack with Cake Build

Use the [Cake.ExcelDnaPack](https://github.com/cake-contrib/Cake.ExcelDnaPack/) addin for the [Cake](https://cakebuild.net) build automation system that enables you to use ExcelDnaPack for packing Excel-DNA addins into a single .xll file. [Cake.ExcelDnaPack](https://github.com/cake-contrib/Cake.ExcelDnaPack/) targets .NET 5.0, .NET Standard 2.0 and .NET Framework 4.6.1, and runs on Windows.

e.g.

```csharp
#tool "nuget:?package=ExcelDnaPack&version=1.5.1"
#addin "nuget:?package=Cake.ExcelDnaPack&version=2.0.0"

Task("Pack")
    .Does(context =>
{
    ExcelDnaPack(@"MyAddins\FirstAddin.dna");
});

RunTarget("Pack");
```

_Make sure the `&version=` attribute references the latest versions of [ExcelDnaPack](https://www.nuget.org/packages/ExcelDnaPack/) and [Cake.ExcelDnaPack](https://www.nuget.org/packages/Cake.ExcelDnaPack/) available on nuget.org_.

For more details on how `Cake.ExcelDnaPack` works, check its [documentation](https://github.com/cake-contrib/Cake.ExcelDnaPack).

## Release History

Click on the [Releases](https://github.com/augustoproiete/ExcelDnaPack-NuGet/releases) tab on GitHub.

## License

_Copyright &copy; 2022 Excel-DNA Contributors - Provided under the [Apache License, Version 2.0](LICENSE)._

---

_The ExcelDnaPack tool is Copyright &copy; 2005-2022 Govert van Drimmelen - Provided under the [zLib license](https://opensource.org/licenses/Zlib). See file LICENSE.txt included in the NuGet package._
