Open XML Document Code Generator
================================

The Open XML Document Code Generator is a dotnet standard library that contains processes that convert OpenXML documents (such as .docx, .pptx, and .xlsx) into [CodeCompileUnit](https://docs.microsoft.com/en-us/dotnet/api/system.codedom.codecompileunit?view=netcore-3.1) objects that can be transformed into source code using any class that inherits from the [CodeDomProvider](https://docs.microsoft.com/en-us/dotnet/api/system.codedom.compiler.codedomprovider?view=netcore-3.1) class.  This allows source code generation into other .NET languages, such as Visual Basic.net, for greater learning possibilities.

Please be aware that while this project is producing working code, I would consider this is in alpha status and not yet ready for production.  More testing should be done before it will be production ready.

## Example

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
