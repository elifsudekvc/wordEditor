# SautinSoft.HtmlToRtf

[![Nuget version](https://img.shields.io/nuget/v/sautinsoft.htmltortf?style=for-the-badge)](https://www.nuget.org/packages/sautinsoft.htmltortf/)[![Nuget downloads](https://img.shields.io/nuget/dt/sautinsoft.htmltortf?style=for-the-badge)](https://www.nuget.org/packages/sautinsoft.htmltortf/)

![htmltortf](https://camo.githubusercontent.com/6f21db7dfca6ea57f0993aa40beb9b838d2aff49274fced91b676590b63c4852/68747470733a2f2f7777772e73617574696e736f66742e636f6d2f6d656469612f6769746875622f682e706e67)

[SautinSoft.HtmlToRtf](https://sautinsoft.com/products/html-to-rtf/) is .NET assembly (SDK) which gives API to convert HTML to DOCX, RTF and Text. Merge and replace text in RTF documents.

## Quick links

+ [Developer Guide](https://sautinsoft.com/products/html-to-rtf/help/net/developer-guide/convert-html-to-docx-csharp-vb-net.php)
+ [API Reference](https://sautinsoft.com/products/html-to-rtf/help/net/api-reference/html/N_SautinSoft.htm)


## Top Features

+ [Convert HTML to DOCX file.](https://sautinsoft.com/products/html-to-rtf/help/net/developer-guide/convert-html-to-docx-file-csharp-vb-net.php)
+ [Convert HTML to RTF Stream.](https://sautinsoft.com/products/html-to-rtf/help/net/developer-guide/convert-html-to-rtf-stream-csharp-vb-net.php)
+ [Convert HTML to Text string.](https://sautinsoft.com/products/html-to-rtf/help/net/developer-guide/convert-html-to-text-string-csharp-vb-net.php)
+ [How to merge multiple RTF files.](https://sautinsoft.com/products/html-to-rtf/help/net/developer-guide/merge-multiple-rtf-files-csharp-vb-net.php)
+ [How to add a page header and footer.](https://sautinsoft.com/products/html-to-rtf/help/net/developer-guide/convert-html-rtf-docx-add-page-header-and-footer-csharp-vb-net.php)
+ [How to set a page size and margins.](https://sautinsoft.com/products/html-to-rtf/help/net/developer-guide/convert-html-rtf-docx-set-page-size-and-margins-csharp-vb-net.php)

## System Requirement

* .NET Framework 4.6.1 - 4.8.1
* .NET Core 2.0 - 3.1, .NET 5, 6, 7, 8
* .NET Standard 2.0
* Windows, Linux, macOS, Android, iOS.

## Getting Started with HTML to RTF .Net

Are you ready to give HTML to RTF .NET a try? Simply execute `Install-Package sautinsoft.htmltortf` from Package Manager Console in Visual Studio to fetch the NuGet package. If you already have HTML to RTF .NET and want to upgrade the version, please execute `Update-Package sautinsoft.htmltortf` to get the latest version.

## Convert HTML to WORD

```csharp
SautinSoft.HtmlToRtf h = new SautinSoft.HtmlToRtf();
string inputFile = @"sample.html";
// You want to save in DOCX.
string outputFile = @"result.docx";
HtmlConvertOptions opt = new HtmlConvertOptions();
opt.OutputFormat = HtmlToRtf.OutputFormat.Docx;
h.Convert(inputFile, outputFile, opt);
```
## Merge multiple RTF files

```csharp
SautinSoft.HtmlToRtf h = new SautinSoft.HtmlToRtf();
// Array with several RTF files.
string[] rtfFiles = new string[] { "footer.rtf", "footer.rtf", "footer.rtf" };
string singleRtf = String.Empty;

// Let's divide RTF documents using page break.
h.MergeOptions.PageBreakBetweenDocuments = true;

foreach (string rtfFile in rtfFiles)
	{
        string rtfFilePath = Path.Combine(htmlDir.FullName, rtfFile);

        // Copy 1st RTF to 'singleRtf'
        if (String.IsNullOrEmpty(singleRtf))
            singleRtf = File.ReadAllText(rtfFilePath);

        // Merge 2nd, 3rd ....
        else
            singleRtf = h.MergeRtfString(singleRtf, File.ReadAllText(rtfFilePath));
    }

// Save 'singleRtf' to a file only for demonstration purposes.
string singleRtfPath = Path.Combine(htmlDir.FullName, "single.rtf");
File.WriteAllText(singleRtfPath, singleRtf);
```

## Resources

+ **Website:** [www.sautinsoft.com](http://www.sautinsoft.com)
+ **Product Home:** [HTML to RTF .Net](https://sautinsoft.com/products/html-to-rtf/)
+ [Download SautinSoft.HTMLtoRTF](https://sautinsoft.com/products/html-to-rtf/download.php)
+ [Developer Guide](https://sautinsoft.com/products/html-to-rtf/help/net/developer-guide/convert-html-to-docx-csharp-vb-net.php)
+ [API Reference](https://sautinsoft.com/products/html-to-rtf/help/net/api-reference/html/N_SautinSoft.htm)
+ [Support Team](https://sautinsoft.com/support.php)
+ [License](https://sautinsoft.com/products/html-to-rtf/help/net/getting-started/agreement.php)
