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
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace Serialize.OpenXml.CodeGen
{
    /// <summary>
    /// Static class that converts <see cref="OpenXmlPart"/> objects
    /// into Code DOM objects.
    /// </summary>
    public static class OpenXmlPartExtensions
    {
        #region Private Static Fields

        /// <summary>
        /// The default parameter name for an <see cref="OpenXmlPart"/> object.
        /// </summary>
        private const string methodParamName = "part";

        #endregion

        #region Public Static Methods

        /// <summary>
        /// Creates the appropriate code objects needed to create the entry method for the
        /// current request.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object and relationship id to build code for.
        /// </param>
        /// <param name="settings">
        /// The <see cref="ISerializeSettings"/> to use during the code generation
        /// process.
        /// </param>
        /// /// <param name="typeCounts">
        /// A lookup <see cref="IDictionary{TKey, TValue}"/> object containing the
        /// number of times a given type was referenced.  This is used for variable naming
        /// purposes.
        /// </param>
        /// <param name="namespaces">
        /// <see cref="IDictionary{TKey, TValue}"/> used to keep track of all openxml
        /// namespaces used during the process.
        /// </param>
        /// <param name="blueprints">
        /// The collection of <see cref="OpenXmlPartBluePrint"/> objects that have already been
        /// visited.
        /// </param>
        /// <param name="rootVar">
        /// The root variable name and <see cref="Type"/> to use when building code
        /// statements to create new <see cref="OpenXmlPart"/> objects.
        /// </param>
        /// <param name="token">
        /// Task cancellation token from the parent method.
        /// </param>
        /// <returns>
        /// A collection of code statements and expressions that could be used to generate
        /// a new <paramref name="part"/> object from code.
        /// </returns>
        public static CodeStatementCollection BuildEntryMethodCodeStatements(
            IdPartPair part,
            ISerializeSettings settings,
            IDictionary<string, int> typeCounts,
            IDictionary<string, string> namespaces,
            OpenXmlPartBluePrintCollection blueprints,
            KeyValuePair<string, Type> rootVar,
            CancellationToken token)
        {
            // Argument validation
            if (part is null) throw new ArgumentNullException(nameof(part));
            if (settings is null) throw new ArgumentNullException(nameof(settings));
            if (blueprints is null) throw new ArgumentNullException(nameof(blueprints));

            if (String.IsNullOrWhiteSpace(rootVar.Key))
            {
                throw new ArgumentNullException(nameof(rootVar.Key));
            }

            bool hasHandlers = settings?.Handlers != null;

            // Check to see if the task has been cancelled.
            if (token.IsCancellationRequested)
            {
                token.ThrowIfCancellationRequested();
            }

            // Use the custom handler methods if present and provide actual code
            if (hasHandlers && settings.Handlers.TryGetValue(part.OpenXmlPart.GetType(),
                out IOpenXmlHandler h))
            {
                if (h is IOpenXmlPartHandler partHandler)
                {
                    var customStatements = partHandler.BuildEntryMethodCodeStatements(
                        part, settings, typeCounts, namespaces, blueprints, rootVar, token);

                    if (customStatements != null) return customStatements;
                }
            }

            var result = new CodeStatementCollection();
            var partType = part.OpenXmlPart.GetType();

            CodeMethodReferenceExpression referenceExpression;
            CodeMethodInvokeExpression invokeExpression;
            CodeMethodReferenceExpression methodReference;

            // Check to see if the task has been cancelled.
            if (token.IsCancellationRequested)
            {
                result.Clear();
                token.ThrowIfCancellationRequested();
            }

            // Make sure that the namespace for the current part is captured
            if (!namespaces.ContainsKey(partType.Namespace))
            {
                // All OpenXmlPart objects are in the DocumentFormat.OpenXml.Packaging
                // namespace so there shouldn't be any collisions here.
                namespaces.Add(partType.Namespace, String.Empty);
            }

            // If the URI of the current part has already been included into
            // the blue prints collection, build the AddPart invocation
            // code statement and exit current method iteration.
            if (blueprints.TryGetValue(part.OpenXmlPart.Uri, out OpenXmlPartBluePrint bpTemp))
            {
                // Surround this snippet with blank lines to make it
                // stand out in the current section of code.
                result.AddBlankLine();
                referenceExpression = new CodeMethodReferenceExpression(
                    new CodeVariableReferenceExpression(rootVar.Key), "AddPart",
                    new CodeTypeReference(part.OpenXmlPart.GetType().Name));
                invokeExpression = new CodeMethodInvokeExpression(referenceExpression,
                    new CodeVariableReferenceExpression(bpTemp.VariableName),
                    new CodePrimitiveExpression(part.RelationshipId));
                result.Add(invokeExpression);
                result.AddBlankLine();
                return result;
            }

            var partTypeName = partType.Name;
            var partTypeFullName = partType.FullName;
            string varName = partType.Name.ToCamelCase();

            // Assign the appropriate variable name
            if (typeCounts.ContainsKey(partTypeFullName))
            {
                varName = String.Concat(varName, typeCounts[partTypeFullName]++);
            }
            else
            {
                typeCounts.Add(partTypeFullName, 1);
            }

            // Setup the blueprint
            bpTemp = new OpenXmlPartBluePrint(part.OpenXmlPart, varName);

            // Need to evaluate the current OpenXmlPart type first to make sure the
            // correct "Add" statement is used as not all Parts can be initialized
            // using the "AddNewPart" method
            var addNewPartExpressions = CreateAddNewPartMethod(part, rootVar);
            // referenceExpression = addNewPartExpressions.Item1;
            invokeExpression = addNewPartExpressions.Item2;

            result.Add(new CodeVariableDeclarationStatement(
                partTypeName, varName, invokeExpression));

            // Because the custom AddNewPart methods don't consistently take in a string relId
            // as a parameter, the id needs to be assigned after it is created.
            if (addNewPartExpressions.Item3)
            {
                methodReference = new CodeMethodReferenceExpression(
                    new CodeVariableReferenceExpression(rootVar.Key), "ChangeIdOfPart");
                result.Add(new CodeMethodInvokeExpression(methodReference,
                    new CodeVariableReferenceExpression(varName),
                    new CodePrimitiveExpression(part.RelationshipId)));
            }

            // Add the call to the method to populate the current OpenXmlPart object
            methodReference = new CodeMethodReferenceExpression(new CodeThisReferenceExpression(),
                bpTemp.MethodName);
            result.Add(new CodeMethodInvokeExpression(methodReference,
                new CodeDirectionExpression(FieldDirection.Ref,
                    new CodeVariableReferenceExpression(varName))));

            // Add the relationships, if applicable
            var relCodeStatements = part.OpenXmlPart.GenerateRelationshipCodeStatements(varName);

            if (relCodeStatements.Count > 0)
            {
                result.AddRange(relCodeStatements);
            }

            // put a line break before going through the child parts
            result.AddBlankLine();

            // Add the current blueprint to the collection
            blueprints.Add(bpTemp);

            // Now check to see if there are any child parts for the current OpenXmlPart object.
            if (bpTemp.Part.Parts != null)
            {
                OpenXmlPartBluePrint childBluePrint;

                foreach (var p in bpTemp.Part.Parts)
                {
                    // Check to see if the task has been cancelled.
                    if (token.IsCancellationRequested)
                    {
                        result.Clear();
                        token.ThrowIfCancellationRequested();
                    }

                    // If the current child object has already been created, simply add a reference
                    // to said object using the AddPart method.
                    if (blueprints.Contains(p.OpenXmlPart.Uri))
                    {
                        childBluePrint = blueprints[p.OpenXmlPart.Uri];

                        referenceExpression = new CodeMethodReferenceExpression(
                            new CodeVariableReferenceExpression(varName), "AddPart",
                            new CodeTypeReference(p.OpenXmlPart.GetType().Name));

                        invokeExpression = new CodeMethodInvokeExpression(referenceExpression,
                            new CodeVariableReferenceExpression(childBluePrint.VariableName),
                            new CodePrimitiveExpression(p.RelationshipId));

                        result.Add(invokeExpression);
                        continue;
                    }

                    // If this is a new part, call this method with the current part's details
                    result.AddRange(BuildEntryMethodCodeStatements(
                        p, settings, typeCounts, namespaces, blueprints,
                        new KeyValuePair<string, Type>(varName, partType), token));
                }
            }

            return result;
        }

        /// <summary>
        /// Creates the appropriate helper methods for all of the <see cref="OpenXmlPart"/> objects
        /// for the current request.
        /// </summary>
        /// <param name="bluePrints">
        /// The collection of <see cref="OpenXmlPartBluePrint"/> objects that have already been
        /// visited.
        /// </param>
        /// <param name="settings">
        /// The <see cref="ISerializeSettings"/> to use during the code generation
        /// process.
        /// </param>
        /// <param name="namespaces">
        /// <see cref="IDictionary{TKey, TValue}"/> collection used to keep track of all openxml
        /// namespaces used during the process.
        /// </param>
        /// <param name="token">
        /// Task cancellation token from the parent method.
        /// </param>
        /// <returns>
        /// A collection of code helper statements and expressions that could be used to generate a
        /// new <see cref="OpenXmlPart"/> object from code.
        /// </returns>
        public static CodeTypeMemberCollection BuildHelperMethods(
            OpenXmlPartBluePrintCollection bluePrints,
            ISerializeSettings settings,
            IDictionary<string, string> namespaces,
            CancellationToken token)
        {
            if (bluePrints == null) throw new ArgumentNullException(nameof(bluePrints));
            var result = new CodeTypeMemberCollection();
            var localTypes = new TypeMonitorCollection();
            Type rootElementType;
            CodeMemberMethod method;
            bool hasHandlers = settings?.Handlers != null;

            foreach (var bp in bluePrints)
            {
                // Check to see if the task has been cancelled.
                if (token.IsCancellationRequested)
                {
                    result.Clear();
                    token.ThrowIfCancellationRequested();
                }

                // Implement the custom helper if present
                if (hasHandlers && settings.Handlers.TryGetValue(bp.PartType, out IOpenXmlHandler h))
                {
                    if (h is IOpenXmlPartHandler partHandler)
                    {
                        method = partHandler.BuildHelperMethod(
                            bp.Part, bp.MethodName, settings, namespaces, token);

                        if (method != null)
                        {
                            result.Add(method);
                            continue;
                        }
                    }
                }

                // Setup the first method
                method = new CodeMemberMethod()
                {
                    Name = bp.MethodName,
                    Attributes = MemberAttributes.Private | MemberAttributes.Final,
                    ReturnType = new CodeTypeReference()
                };
                method.Parameters.Add(
                    new CodeParameterDeclarationExpression(bp.Part.GetType().Name, methodParamName)
                    { Direction = FieldDirection.Ref });

                // Code part elements next
                if (bp.Part.RootElement is null)
                {
                    // Put all of the pieces together
                    method.Statements.AddRange(bp.Part.BuildPartFeedData(namespaces));
                }
                else
                {
                    rootElementType = bp.Part.RootElement?.GetType();
                    localTypes.Clear();

                    // Build the element details of the requested part for the current method
                    method.Statements.AddRange(
                        bp.Part.RootElement.BuildCodeStatements(
                            settings, localTypes, namespaces, token, out string rootElementVar));

                    // Now finish up the current method by assigning the OpenXmlElement code
                    // statements back to the appropriate property of the part parameter
                    if (rootElementType != null && !String.IsNullOrWhiteSpace(rootElementVar))
                    {
                        foreach (var paramProp in bp.Part.GetType().GetProperties())
                        {
                            // Check to see if the task has been cancelled.
                            if (token.IsCancellationRequested)
                            {
                                result.Clear();
                                token.ThrowIfCancellationRequested();
                            }

                            if (paramProp.PropertyType == rootElementType)
                            {
                                var varRef = new CodeVariableReferenceExpression(rootElementVar);
                                var paramRef = new CodeArgumentReferenceExpression(methodParamName);
                                var propRef = new CodePropertyReferenceExpression(
                                    paramRef, paramProp.Name);

                                method.Statements.Add(new CodeAssignStatement(propRef, varRef));
                                break;
                            }
                        }
                    }
                }
                result.Add(method);
            }

            return result;
        }

        /// <summary>
        /// Builds code statements that will build <paramref name="part"/> using the
        /// <see cref="OpenXmlPart.FeedData(Stream)"/> method.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to build the source code for.
        /// </param>
        /// <param name="namespaces">
        /// <see cref="IDictionary{TKey, TValue}"/> values used to keep
        /// track of all openxml namespaces used during the process.
        /// </param>
        /// <returns>
        /// A <see cref="CodeStatementCollection">collection of code statements</see>
        /// that would regenerate <paramref name="part"/> using the
        /// <see cref="OpenXmlPart.FeedData(Stream)"/> method.
        /// </returns>
        public static CodeStatementCollection BuildPartFeedData(
            this OpenXmlPart part,
            IDictionary<string, string> namespaces)
        {
            // Make sure no null values were passed.
            if (part == null) throw new ArgumentNullException(nameof(part));
            if (namespaces == null) throw new ArgumentNullException(nameof(namespaces));

            // If the root element is not present (aka: null) then perform a simple feed
            // dump of the part in the current method
            const string memName = "mem";
            const string b64Name = "base64";

            var result = new CodeStatementCollection();

            // Add the necessary namespaces by hand to the namespace set
            if (!namespaces.ContainsKey("System"))
            {
                namespaces.Add("System", String.Empty);
            }
            if (!namespaces.ContainsKey("System.IO"))
            {
                namespaces.Add("System.IO", String.Empty);
            }

            using (var partStream = part.GetStream(FileMode.Open, FileAccess.Read))
            {
                using (var mem = new MemoryStream())
                {
                    partStream.CopyTo(mem);
                    result.Add(new CodeVariableDeclarationStatement(typeof(string), b64Name,
                        new CodePrimitiveExpression(Convert.ToBase64String(mem.ToArray()))));
                }
            }
            result.AddBlankLine();

            var fromBase64 = new CodeMethodReferenceExpression(
                new CodeVariableReferenceExpression("Convert"),
                "FromBase64String");
            var invokeFromBase64 = new CodeMethodInvokeExpression(fromBase64,
                new CodeVariableReferenceExpression("base64"));
            var createStream = new CodeObjectCreateExpression(new CodeTypeReference("MemoryStream"),
                invokeFromBase64, new CodePrimitiveExpression(false));
            var feedData = new CodeMethodReferenceExpression(
                new CodeArgumentReferenceExpression(methodParamName), "FeedData");
            var invokeFeedData = new CodeMethodInvokeExpression(
                feedData, new CodeVariableReferenceExpression(memName));
            var disposeMem = new CodeMethodReferenceExpression(
                new CodeVariableReferenceExpression(memName), "Dispose");
            var invokeDisposeMem = new CodeMethodInvokeExpression(disposeMem);

            // Setup the try statement
            var tryAndCatch = new CodeTryCatchFinallyStatement();
            tryAndCatch.TryStatements.Add(invokeFeedData);
            tryAndCatch.FinallyStatements.Add(invokeDisposeMem);

            // Put all of the pieces together
            result.Add(new CodeVariableDeclarationStatement("Stream", memName, createStream));
            result.Add(tryAndCatch);

            return result;
        }

        /// <summary>
        /// Converts an <see cref="OpenXmlPart"/> into a CodeDom object that can be used
        /// to build code in a given .NET language to build the referenced <see cref="OpenXmlPart"/>.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to generate source code for.
        /// </param>
        /// <returns>
        /// A new <see cref="CodeCompileUnit"/> containing the instructions to build
        /// the referenced <see cref="OpenXmlPart"/>.
        /// </returns>
        public static CodeCompileUnit GenerateSourceCode(this OpenXmlPart part)
        {
            return part.GenerateSourceCode(new DefaultSerializeSettings());
        }

        /// <summary>
        /// Converts an <see cref="OpenXmlPart"/> into a CodeDom object that can be used
        /// to build code in a given .NET language to build the referenced <see cref="OpenXmlPart"/>.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to generate source code for.
        /// </param>
        /// <param name="opts">
        /// The <see cref="NamespaceAliasOptions"/> to apply to the resulting source code.
        /// </param>
        /// <returns>
        /// A new <see cref="CodeCompileUnit"/> containing the instructions to build
        /// the referenced <see cref="OpenXmlPart"/>.
        /// </returns>
        public static CodeCompileUnit GenerateSourceCode(this OpenXmlPart part, NamespaceAliasOptions opts)
        {
            return part.GenerateSourceCode(new DefaultSerializeSettings(opts));
        }

        /// <summary>
        /// Converts an <see cref="OpenXmlPart"/> into a CodeDom object that can be used
        /// to build code in a given .NET language to build the referenced <see cref="OpenXmlPart"/>.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to generate source code for.
        /// </param>
        /// <param name="settings">
        /// The <see cref="ISerializeSettings"/> to use during the code generation
        /// process.
        /// </param>
        /// <returns>
        /// A new <see cref="CodeCompileUnit"/> containing the instructions to build
        /// the referenced <see cref="OpenXmlPart"/>.
        /// </returns>
        public static CodeCompileUnit GenerateSourceCode(
            this OpenXmlPart part, ISerializeSettings settings)
        {
            return DefaultSerializeSettings.TaskIndustry.StartNew(
                () => part.GenerateSourceCodeAsync(settings, CancellationToken.None))
                .Unwrap()
                .GetAwaiter()
                .GetResult();
        }

        /// <summary>
        /// Converts an <see cref="OpenXmlPart"/> into a <see cref="string"/> representation
        /// of dotnet source code that can be compiled to build <paramref name="part"/>.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to generate source code for.
        /// </param>
        /// <param name="provider">
        /// The <see cref="CodeDomProvider"/> object to create the resulting source code.
        /// </param>
        /// <returns>
        /// A <see cref="string"/> representation of the source code generated by
        /// <paramref name="provider"/> that could create <paramref name="part"/> when compiled.
        /// </returns>
        public static string GenerateSourceCode(this OpenXmlPart part, CodeDomProvider provider)
        {
            return part.GenerateSourceCode(new DefaultSerializeSettings(), provider);
        }

        /// <summary>
        /// Converts an <see cref="OpenXmlPart"/> into a <see cref="string"/> representation
        /// of dotnet source code that can be compiled to build <paramref name="part"/>.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to generate source code for.
        /// </param>
        /// <param name="opts">
        /// The <see cref="NamespaceAliasOptions"/> to apply to the resulting source code.
        /// </param>
        /// <param name="provider">
        /// The <see cref="CodeDomProvider"/> object to create the resulting source code.
        /// </param>
        /// <returns>
        /// A <see cref="string"/> representation of the source code generated by
        /// <paramref name="provider"/> that could create <paramref name="part"/> when compiled.
        /// </returns>
        public static string GenerateSourceCode(
            this OpenXmlPart part, NamespaceAliasOptions opts, CodeDomProvider provider)
        {
            return part.GenerateSourceCode(new DefaultSerializeSettings(opts), provider);
        }

        /// <summary>
        /// Converts an <see cref="OpenXmlPart"/> into a <see cref="string"/> representation
        /// of dotnet source code that can be compiled to build <paramref name="part"/>.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to generate source code for.
        /// </param>
        /// <param name="settings">
        /// The <see cref="ISerializeSettings"/> to use during the code generation
        /// process.
        /// </param>
        /// <param name="provider">
        /// The <see cref="CodeDomProvider"/> object to create the resulting source code.
        /// </param>
        /// <returns>
        /// A <see cref="string"/> representation of the source code generated by
        /// <paramref name="provider"/> that could create <paramref name="part"/> when compiled.
        /// </returns>
        public static string GenerateSourceCode(
            this OpenXmlPart part, ISerializeSettings settings, CodeDomProvider provider)
        {
            var codeString = new System.Text.StringBuilder();
            var code = part.GenerateSourceCode(settings);

            using (var sw = new StringWriter(codeString))
            {
                provider.GenerateCodeFromCompileUnit(code, sw,
                    new CodeGeneratorOptions() { BracingStyle = "C" });
            }
            return codeString.ToString().RemoveOutputHeaders().Trim();
        }

        /// <summary>
        /// Converts an <see cref="OpenXmlPart"/> into a CodeDom object that can be used
        /// to build code in a given .NET language to build the referenced <see cref="OpenXmlPart"/>.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to generate source code for.
        /// </param>
        /// <param name="token">
        /// Task cancellation token.
        /// </param>
        /// <returns>
        /// A new <see cref="CodeCompileUnit"/> containing the instructions to build
        /// the referenced <see cref="OpenXmlPart"/>.
        /// </returns>
        public static async Task<CodeCompileUnit> GenerateSourceCodeAsync(
            this OpenXmlPart part,
            CancellationToken token)
        {
            return await part.GenerateSourceCodeAsync(new DefaultSerializeSettings(), token);
        }

        /// <summary>
        /// Converts an <see cref="OpenXmlPart"/> into a CodeDom object that can be used
        /// to build code in a given .NET language to build the referenced <see cref="OpenXmlPart"/>.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to generate source code for.
        /// </param>
        /// <param name="opts">
        /// The <see cref="NamespaceAliasOptions"/> to apply to the resulting source code.
        /// </param>
        /// <param name="token">
        /// Task cancellation token.
        /// </param>
        /// <returns>
        /// A new <see cref="CodeCompileUnit"/> containing the instructions to build
        /// the referenced <see cref="OpenXmlPart"/>.
        /// </returns>
        public static async Task<CodeCompileUnit> GenerateSourceCodeAsync(
            this OpenXmlPart part,
            NamespaceAliasOptions opts,
            CancellationToken token)
        {
            return await part.GenerateSourceCodeAsync(new DefaultSerializeSettings(opts), token);
        }

        /// <summary>
        /// Converts an <see cref="OpenXmlPart"/> into a CodeDom object that can be used
        /// to build code in a given .NET language to build the referenced <see cref="OpenXmlPart"/>.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to generate source code for.
        /// </param>
        /// <param name="settings">
        /// The <see cref="ISerializeSettings"/> to use during the code generation
        /// process.
        /// </param>
        /// <param name="token">
        /// Task cancellation token.
        /// </param>
        /// <returns>
        /// A new <see cref="CodeCompileUnit"/> containing the instructions to build
        /// the referenced <see cref="OpenXmlPart"/>.
        /// </returns>
        public static async Task<CodeCompileUnit> GenerateSourceCodeAsync(
            this OpenXmlPart part,
            ISerializeSettings settings,
            CancellationToken token)
        {
            return await Task.Run(() =>
            {
                CodeMethodReferenceExpression methodRef;
                OpenXmlPartBluePrint mainBluePrint;
                var result = new CodeCompileUnit();
                var eType = part.GetType();
                var partTypeName = eType.Name;
                var partTypeFullName = eType.FullName;
                var varName = eType.Name.ToCamelCase();
                var partTypeCounts = new Dictionary<string, int>();
                var namespaces = new Dictionary<string, string>();
                var mainNamespace = new CodeNamespace(settings.NamespaceName);
                var bluePrints = new OpenXmlPartBluePrintCollection();

                // Set the var uniqueness indicator
                TypeMonitor.UseUniqueVariableNames = settings.UseUniqueVariableNames;

                // Assign the appropriate variable name
                if (partTypeCounts.ContainsKey(partTypeFullName))
                {
                    varName = String.Concat(varName, partTypeCounts[partTypeFullName]++);
                }
                else
                {
                    partTypeCounts.Add(partTypeFullName, 1);
                }

                // Generate a new blue print for the current part to help create the main
                // method reference then add it to the blue print collection
                mainBluePrint = new OpenXmlPartBluePrint(part, varName);
                bluePrints.Add(mainBluePrint);
                methodRef = new CodeMethodReferenceExpression(
                    new CodeThisReferenceExpression(), mainBluePrint.MethodName);

                // Build the entry method
                var entryMethod = new CodeMemberMethod()
                {
                    Name = $"Create{partTypeName}",
                    ReturnType = new CodeTypeReference(),
                    Attributes = MemberAttributes.Public | MemberAttributes.Final
                };
                entryMethod.Parameters.Add(
                    new CodeParameterDeclarationExpression(partTypeName, methodParamName)
                    { Direction = FieldDirection.Ref });

                // Check to see if the task has been cancelled.
                if (token.IsCancellationRequested)
                {
                    token.ThrowIfCancellationRequested();
                }

                var relCodeStatements = part.GenerateRelationshipCodeStatements(
                    new CodeThisReferenceExpression());
                if (relCodeStatements.Count > 0)
                {
                    entryMethod.Statements.AddRange(relCodeStatements);
                }

                // Add all of the child part references here
                if (part.Parts != null)
                {
                    var rootPartPair = new KeyValuePair<string, Type>(methodParamName, eType);
                    foreach (var pair in part.Parts)
                    {
                        // Check to see if the task has been cancelled.
                        if (token.IsCancellationRequested)
                        {
                            token.ThrowIfCancellationRequested();
                        }

                        entryMethod.Statements.AddRange(BuildEntryMethodCodeStatements(
                            pair,
                            settings,
                            partTypeCounts,
                            namespaces,
                            bluePrints,
                            rootPartPair,
                            token));
                    }
                }

                entryMethod.Statements.Add(new CodeMethodInvokeExpression(methodRef,
                    new CodeArgumentReferenceExpression(methodParamName)));

                // Setup the main class next
                var mainClass = new CodeTypeDeclaration($"{eType.Name}BuilderClass")
                {
                    IsClass = true,
                    Attributes = MemberAttributes.Public
                };
                mainClass.Members.Add(entryMethod);
                mainClass.Members.AddRange(BuildHelperMethods(
                    bluePrints, settings, namespaces, token));

                // Setup the imports
                var codeNameSpaces = new List<CodeNamespaceImport>(namespaces.Count);
                foreach (var ns in namespaces)
                {
                    // Check to see if the task has been cancelled.
                    if (token.IsCancellationRequested)
                    {
                        token.ThrowIfCancellationRequested();
                    }

                    if (!String.IsNullOrWhiteSpace(ns.Value))
                    {
                        codeNameSpaces.Add(settings.NamespaceAliasOptions.BuildNamespaceImport(
                            ns.Key, ns.Value));
                    }
                    else
                    {
                        codeNameSpaces.Add(new CodeNamespaceImport(ns.Key));
                    }
                }
                codeNameSpaces.Sort(new CodeNamespaceImportComparer(settings.NamespaceAliasOptions));

                mainNamespace.Imports.AddRange(codeNameSpaces.ToArray());
                mainNamespace.Types.Add(mainClass);

                // Finish up
                result.Namespaces.Add(mainNamespace);
                return result;
            }, token);
        }

        /// <summary>
        /// Converts an <see cref="OpenXmlPart"/> into a <see cref="string"/> representation
        /// of dotnet source code that can be compiled to build <paramref name="part"/>.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to generate source code for.
        /// </param>
        /// <param name="provider">
        /// The <see cref="CodeDomProvider"/> object to create the resulting source code.
        /// </param>
        /// <param name="token">
        /// Task cancellation token.
        /// </param>
        /// <returns>
        /// A <see cref="string"/> representation of the source code generated by
        /// <paramref name="provider"/> that could create <paramref name="part"/> when compiled.
        /// </returns>
        public static async Task<string> GenerateSourceCodeAsync(
            this OpenXmlPart part,
            CodeDomProvider provider,
            CancellationToken token)
        {
            return await part.GenerateSourceCodeAsync(
                new DefaultSerializeSettings(), provider, token);
        }

        /// <summary>
        /// Converts an <see cref="OpenXmlPart"/> into a <see cref="string"/> representation
        /// of dotnet source code that can be compiled to build <paramref name="part"/>.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to generate source code for.
        /// </param>
        /// <param name="opts">
        /// The <see cref="NamespaceAliasOptions"/> to apply to the resulting source code.
        /// </param>
        /// <param name="provider">
        /// The <see cref="CodeDomProvider"/> object to create the resulting source code.
        /// </param>
        /// <param name="token">
        /// Task cancellation token.
        /// </param>
        /// <returns>
        /// A <see cref="string"/> representation of the source code generated by
        /// <paramref name="provider"/> that could create <paramref name="part"/> when compiled.
        /// </returns>
        public static async Task<string> GenerateSourceCodeAsync(
            this OpenXmlPart part,
            NamespaceAliasOptions opts,
            CodeDomProvider provider,
            CancellationToken token)
        {
            return await part.GenerateSourceCodeAsync(
                new DefaultSerializeSettings(opts), provider, token);
        }

        /// <summary>
        /// Converts an <see cref="OpenXmlPart"/> into a <see cref="string"/> representation
        /// of dotnet source code that can be compiled to build <paramref name="part"/>.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to generate source code for.
        /// </param>
        /// <param name="settings">
        /// The <see cref="ISerializeSettings"/> to use during the code generation
        /// process.
        /// </param>
        /// <param name="provider">
        /// The <see cref="CodeDomProvider"/> object to create the resulting source code.
        /// </param>
        /// <param name="token">
        /// Task cancellation token.
        /// </param>
        /// <returns>
        /// A <see cref="string"/> representation of the source code generated by
        /// <paramref name="provider"/> that could create <paramref name="part"/> when compiled.
        /// </returns>
        public static async Task<string> GenerateSourceCodeAsync(
            this OpenXmlPart part,
            ISerializeSettings settings,
            CodeDomProvider provider,
            CancellationToken token)
        {
            var codeString = new System.Text.StringBuilder();
            var code = await part.GenerateSourceCodeAsync(settings, token);

            using (var sw = new StringWriter(codeString))
            {
                provider.GenerateCodeFromCompileUnit(code, sw,
                    new CodeGeneratorOptions() { BracingStyle = "C" });
            }
            return codeString.ToString().RemoveOutputHeaders().Trim();
        }

        #endregion

        #region Private Static Methods

        /// <summary>
        /// Constructs the appropriate <see cref="CodeMethodReferenceExpression"/> and
        /// <see cref="CodeMethodInvokeExpression"/> objects for adding new parts
        /// to an existing <see cref="OpenXmlPart"/> object.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object and relationship id to build
        /// the code expressions for.
        /// </param>
        /// <param name="rootVar">
        /// The root variable name and <see cref="Type"/> to use when building code
        /// statements to create new <see cref="OpenXmlPart"/> objects.
        /// </param>
        /// <returns>
        /// A tuple containing the AddNewPart code expressions.
        /// </returns>
        private static (CodeMethodReferenceExpression, CodeMethodInvokeExpression, bool)
            CreateAddNewPartMethod(
            IdPartPair part, KeyValuePair<string, Type> rootVar)
        {
            CodeMethodReferenceExpression item1 = null;
            CodeMethodInvokeExpression item2 = null;
            var pType = part.OpenXmlPart.GetType();
            var checkName = $"Add{pType.Name}";
            bool check(MethodInfo m) => m.Name.Equals(checkName, StringComparison.OrdinalIgnoreCase);
            var methods = rootVar.Value.GetMethods().Where(check).ToArray();

            // If no custom AddNewPart methods exist for the requested part, revert back to
            // the default AddNewPart method.
            if (methods.Length == 0)
            {
                item1 = new CodeMethodReferenceExpression(
                    new CodeVariableReferenceExpression(rootVar.Key), "AddNewPart",
                    new CodeTypeReference(pType.Name));
                item2 = new CodeMethodInvokeExpression(item1);
                item2.Parameters.Add(new CodePrimitiveExpression(part.RelationshipId));

                return (item1, item2, false);
            }

            // Initialize the custom AddNewPart method reference and invoke expresstions
            // before examining the method parameters
            item1 = new CodeMethodReferenceExpression(
                new CodeVariableReferenceExpression(rootVar.Key), checkName);
            item2 = new CodeMethodInvokeExpression(item1);

            // Now figure out what parameters are needed for this method
            ParameterInfo[] parameters;
            foreach (var m in methods)
            {
                // Please Note: To my knowledge, the custom AddNewPart methods have 2 main
                // types of method parameters common among all custom methods:
                // * Default (no parameters)
                // * string contentType
                // This loop should be looking for the second type primarily and reverting
                // to the first type if not found.
                parameters = m.GetParameters();

                // Methods with 1 string parameter should be the contentType parameter method
                if (parameters != null &&
                    parameters.Length == 1 &&
                    parameters[0].ParameterType == typeof(string))
                {
                    item2.Parameters.Add(
                        new CodePrimitiveExpression(part.OpenXmlPart.ContentType));
                    break;
                }
            }
            return (item1, item2, true);
        }

        #endregion
    }
}