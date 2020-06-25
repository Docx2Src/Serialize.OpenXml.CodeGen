Open XML Document Code Generator
================================

[![Build status](https://ci.appveyor.com/api/projects/status/rachpeoigx71q1ro/branch/master?svg=true)](https://ci.appveyor.com/project/rmboggs/serialize-openxml-codegen/branch/master)

The Open XML Document Code Generator is a dotnet standard library that contains processes that convert OpenXML documents (such as .docx, .pptx, and .xlsx) into [CodeCompileUnit](https://docs.microsoft.com/en-us/dotnet/api/system.codedom.codecompileunit?view=netcore-3.1) objects that can be transformed into source code using any class that inherits from the [CodeDomProvider](https://docs.microsoft.com/en-us/dotnet/api/system.codedom.compiler.codedomprovider?view=netcore-3.1) class.  This allows source code generation into other .NET languages, such as Visual Basic.net, for greater learning possibilities.

Please be aware that while this project is producing working code, I would consider this is in alpha status and not yet ready for production.  More testing should be done before it will be production ready.

## Examples

Generate the [CodeCompileUnit](https://docs.microsoft.com/en-us/dotnet/api/system.codedom.codecompileunit?view=netcore-3.1) to process manually:

```cs
using System;
using System.IO;
using System.Linq;
using System.Text;
using Serialize.OpenXml.CodeGen;
using System.CodeDom.Compiler;
using Microsoft.CSharp;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace CodeGenSample
{
    class Program
    {
        private static readonly CodeGeneratorOptions Cgo = new CodeGeneratorOptions()
        {
          BracingStyle = "C"
        };

        static void Main(string[] args)
        {
            var sourceFile = new FileInfo(@"C:\Temp\Sample1.xlsx");
            var targetFile = new FileInfo(@"C:\Temp\Sample1.cs");

            using (var source = sourceFile.Open(FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                using (var xlsx = SpreadsheetDocument.Open(source, false))
                {
                    if (xlsx != null)
                    {
                        var codeString = new StringBuilder();
                        var cs = new CSharpCodeProvider();

                        // This will build the CodeCompileUnit object containing all of
                        // the commands that would create the source code to rebuild Sample1.xlsx
                        var code = xlsx.GenerateSourceCode();

                        // This will convert the CodeCompileUnit into C# source code
                        using (var sw = new StringWriter(codeString))
                        {
                            cs.GenerateCodeFromCompileUnit(code, sw, Cgo);
                        }

                        // Save the source code to the target file
                        using (var target = targetFile.Open(FileMode.Create, FileAccess.ReadWrite))
                        {
                            using (var tw = new StreamWriter(target))
                            {
                                tw.Write(codeString.ToString().Trim());
                            }
                            target.Close();
                        }
                    }
                }
                source.Close();
            }
            Console.WriteLine("Press any key to quit");
            Console.ReadKey(true);
        }
    }
}
```

Generate the actual source code as a string value:

```cs
using System;
using System.IO;
using System.Linq;
using System.Text;
using Serialize.OpenXml.CodeGen;
using System.CodeDom.Compiler;
using Microsoft.VisualBasic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace CodeGenSample
{
    class Program
    {
        static void Main(string[] args)
        {
            var sourceFile = new FileInfo(@"./Sample1.xlsx");
            var targetFile = new FileInfo(@"./Sample1.vb");

            using (var source = sourceFile.Open(FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                using (var xlsx = SpreadsheetDocument.Open(source, false))
                {
                    if (xlsx != null)
                    {
                        // Generate VB.NET source code
                        var vb = new VBCodeProvider();

                        // Save the source code to the target file
                        using (var target = targetFile.Open(FileMode.Create, FileAccess.ReadWrite))
                        {
                            using (var tw = new StreamWriter(target))
                            {
                                // Providing the CodeDomProvider object as a parameter will
                                // cause the method to return the source code as a string
                                tw.Write(xlsx.GenerateSourceCode(vb).Trim());
                            }
                            target.Close();
                        }
                    }
                }
                source.Close();
            }
            Console.WriteLine("Press any key to quit");
            Console.ReadKey(true);
        }
    }
}
```
