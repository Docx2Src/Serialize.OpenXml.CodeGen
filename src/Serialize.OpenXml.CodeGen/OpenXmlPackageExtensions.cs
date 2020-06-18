/* MIT License

Copyright (c) 2020 Ryan Boggs

Permission is hereby granted, free of charge, to any person obtaining a copy of this
software and associated documentation files (the "Software"), to deal in the Software
without restriction, including without limitation the rights to use, copy, modify,
merge, publish, distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to the following
conditions:

The above copyright notice and this permission notice shall be included in all copies
or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
DEALINGS IN THE SOFTWARE.
*/

using DocumentFormat.OpenXml.Packaging;
using Serialize.OpenXml.CodeGen.Extentions;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.IO;

namespace Serialize.OpenXml.CodeGen
{
    /// <summary>
    /// Static class that converts <see cref="OpenXmlPackage">packages</see>
    /// into Code DOM objects.
    /// </summary>
    public static class OpenXmlPackageExtensions
    {
        #region Public Static Methods

        /// <summary>
        /// Converts an <see cref="OpenXmlPackage"/> into a CodeDom object that can be used
        /// to build code in a given .NET language to build the referenced <paramref name="pkg"/>.
        /// </summary>
        /// <param name="pkg">
        /// The <see cref="OpenXmlPackage"/> object to generate source code for.
        /// </param>
        /// <returns>
        /// A new <see cref="CodeCompileUnit"/> containing the instructions to build
        /// the referenced <see cref="OpenXmlPackage"/>.
        /// </returns>
        public static CodeCompileUnit GenerateSourceCode(this OpenXmlPackage pkg)
        {
            return pkg.GenerateSourceCode(NamespaceAliasOptions.Default);
        }

        /// <summary>
        /// Converts an <see cref="OpenXmlPackage"/> into a CodeDom object that can be used
        /// to build code in a given .NET language to build the referenced <paramref name="pkg"/>.
        /// </summary>
        /// <param name="pkg">
        /// The <see cref="OpenXmlPackage"/> object to generate source code for.
        /// </param>
        /// <param name="opts">
        /// The <see cref="NamespaceAliasOptions"/> to apply to the resulting source code.
        /// </param>
        /// <returns>
        /// A new <see cref="CodeCompileUnit"/> containing the instructions to build
        /// the referenced <see cref="OpenXmlPackage"/>.
        /// </returns>
        public static CodeCompileUnit GenerateSourceCode(this OpenXmlPackage pkg, NamespaceAliasOptions opts)
        {
            const string pkgVarName = "pkg";
            const string paramName = "stream";
            var result = new CodeCompileUnit();
            var pkgType = pkg.GetType();
            var pkgTypeName = pkgType.Name;
            var partTypeCounts = new Dictionary<string, int>();
            var namespaces = new SortedSet<string>();
            var mainNamespace = new CodeNamespace("OpenXmlSample");
            var bluePrints = new OpenXmlPartBluePrintCollection();
            CodeMemberMethod entryPoint;
            CodeMemberMethod createParts;
            CodeTypeDeclaration mainClass;
            CodeTryCatchFinallyStatement tryAndCatch;
            
            // Add all initial namespace names first
            namespaces.Add("System.IO");

            // Create the entry method
            entryPoint = new CodeMemberMethod()
            {
                Name = "CreatePackage",
                ReturnType = new CodeTypeReference(),
                Attributes = MemberAttributes.Public | MemberAttributes.Final
            };
            entryPoint.Parameters.Add(new CodeParameterDeclarationExpression(typeof(Stream).Name, paramName));
            
            // Create package declaration expression first
            entryPoint.Statements.Add(new CodeVariableDeclarationStatement(pkgTypeName, pkgVarName));
            
            // initialize package var
            var pkgCreateMethod = new CodeMethodReferenceExpression(
                new CodeTypeReferenceExpression(pkgTypeName), 
                "Create");
            var pkgCreateInvoke = new CodeMethodInvokeExpression(pkgCreateMethod, 
                new CodeVariableReferenceExpression(paramName));
            var initializePkg = new CodeAssignStatement(
                new CodeVariableReferenceExpression(pkgVarName),
                pkgCreateInvoke);

            // Call CreateParts method
            var partsCreateMethod = new CodeMethodReferenceExpression(
                new CodeThisReferenceExpression(),
                "CreateParts");
            var partsCreateInvoke = new CodeMethodInvokeExpression(
                partsCreateMethod,
                new CodeVariableReferenceExpression(pkgVarName));

            // Clean up pkg var
            var pkgDisposeMethod = new CodeMethodReferenceExpression(
                new CodeVariableReferenceExpression(pkgVarName),
                "Dispose");
            var pkgDisposeInvoke = new CodeMethodInvokeExpression(
                pkgDisposeMethod);

            // Setup the try/catch statement to properly initialize the pkg var
            tryAndCatch = new CodeTryCatchFinallyStatement();

            // Try statements
            tryAndCatch.TryStatements.Add(initializePkg);
            tryAndCatch.TryStatements.Add(new CodeSnippetStatement(String.Empty));
            tryAndCatch.TryStatements.Add(partsCreateInvoke);

            // Finally statements
            tryAndCatch.FinallyStatements.Add(pkgDisposeInvoke);
            entryPoint.Statements.Add(tryAndCatch);

            // Create the CreateParts method
            createParts = new CodeMemberMethod()
            {
                Name = "CreateParts",
                ReturnType = new CodeTypeReference(),
                Attributes = MemberAttributes.Private | MemberAttributes.Final
            };
            createParts.Parameters.Add(new CodeParameterDeclarationExpression(pkgTypeName, pkgVarName));

            // Add all of the child part references here
            if (pkg.Parts != null)
            {
                foreach (var pair in pkg.Parts)
                {
                    createParts.Statements.AddRange(
                        OpenXmlPartExtensions.BuildEntryMethodCodeStatements(
                            pair, opts, partTypeCounts, namespaces, bluePrints, pkgVarName));
                }
            }

            // Setup the main class next
            mainClass = new CodeTypeDeclaration($"{pkgTypeName}BuilderClass")
            {
                IsClass = true,
                Attributes = MemberAttributes.Public
            };

            // Setup the main class members
            mainClass.Members.Add(entryPoint);
            mainClass.Members.Add(createParts);
            mainClass.Members.AddRange(OpenXmlPartExtensions.BuildHelperMethods
                (bluePrints, opts, namespaces));

            // Setup the imports
            var codeNameSpaces = new List<CodeNamespaceImport>(namespaces.Count);
            foreach (var ns in namespaces)
            {
                codeNameSpaces.Add(ns.GetCodeNamespaceImport(opts));
            }
            codeNameSpaces.Sort(new CodeNamespaceImportComparer());

            mainNamespace.Imports.AddRange(codeNameSpaces.ToArray());
            mainNamespace.Types.Add(mainClass);

            // Finish up
            result.Namespaces.Add(mainNamespace);
            return result;
        }

        #endregion
    }
}