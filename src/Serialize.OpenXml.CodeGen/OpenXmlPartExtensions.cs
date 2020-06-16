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
using System.Globalization;
using System.IO;
using System.Linq;

namespace Serialize.OpenXml.CodeGen
{
    /// <summary>
    /// Static class that converts <see cref="OpenXmlPart">OpenXmlParts</see>
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
        /// Converts an <see cref="OpenXmlPart"/> into a CodeDom object that can be used
        /// to build code in a given .NET language to build the referenced <see cref="OpenXmlPart"/>.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to generate source code for.
        /// </param>
        /// /// <returns>
        /// /// A new <see cref="CodeCompileUnit"/> containing the instructions to build
        /// the referenced <see cref="OpenXmlPart"/>.
        /// </returns>
        public static CodeCompileUnit GenerateSourceCode(this OpenXmlPart part)
        {
            return part.GenerateSourceCode(NamespaceAliasOptions.Default);
        }

        /// <summary>
        /// Converts an <see cref="OpenXmlPart"/> into a CodeDom object that can be used
        /// to build code in a given .NET language to build the referenced <see cref="OpenXmlPart"/>.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to generate source code for.
        /// </param>
        /// /// <param name="opts">
        /// The <see cref="NamespaceAliasOptions"/> to apply to the resulting source code.
        /// </param>
        /// <returns>
        /// A new <see cref="CodeCompileUnit"/> containing the instructions to build
        /// the referenced <see cref="OpenXmlPart"/>.
        /// </returns>
        public static CodeCompileUnit GenerateSourceCode(this OpenXmlPart part, NamespaceAliasOptions opts)
        {
            var result = new CodeCompileUnit();
            var eType = part.GetType();
            var partTypeName = eType.Name;
            var partTypeCounts = new Dictionary<string, int>();
            var uris = new HashSet<Uri>(new UriEqualityComparer());
            var namespaces = new SortedSet<string>();
            var mainNamespace = new CodeNamespace("OpenXmlSample");
            var helperMembers = BuildCodeStatements(part, opts, partTypeCounts, namespaces, uris, out string partName);
            var methodRef = new CodeMethodReferenceExpression(new CodeThisReferenceExpression(), partName);

            // Build the entry method
            var entryMethod = new CodeMemberMethod()
            {
                Name = $"Create{partTypeName}",
                ReturnType = new CodeTypeReference(),
                Attributes = MemberAttributes.Public | MemberAttributes.Final
            };
            entryMethod.Parameters.Add(new CodeParameterDeclarationExpression(partTypeName, methodParamName));
            entryMethod.Statements.Add(new CodeMethodInvokeExpression(methodRef, 
                new CodeVariableReferenceExpression(methodParamName)));

            // Setup the main class next
            var mainClass = new CodeTypeDeclaration($"{eType.Name}BuilderClass")
            {
                IsClass = true,
                Attributes = MemberAttributes.Public
            };
            mainClass.Members.Add(entryMethod);
            mainClass.Members.AddRange(helperMembers);
            
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

        #region Internal Static Methods

        /// <summary>
        /// Builds the appropriate code objects that would build the contents of an
        /// <see cref="OpenXmlPart"/> object.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to build code for.
        /// </param>
        /// /// <param name="opts">
        /// The <see cref="NamespaceAliasOptions"/> to use during the variable naming 
        /// process.
        /// </param>
        /// <param name="typeCounts">
        /// A lookup <see cref="IDictionary{TKey, TValue}"/> object containing the
        /// number of times a given type was referenced.  This is used for variable naming
        /// purposes.
        /// </param>
        /// <param name="namespaces">
        /// Collection <see cref="ISet{T}"/> used to keep track of all openxml namespaces
        /// used during the process.
        /// </param>
        /// <param name="targets">
        /// A collection of <see cref="System.Uri"/> objects of all of the child parts that
        /// are visited during this process.
        /// </param>
        /// <param name="partName">
        /// /// The variable name of the root <see cref="OpenXmlPart"/> object that was built
        /// from the <paramref name="part"/>.
        /// </param>
        /// /// <returns>
        /// A collection of code statements and expressions that could be used to generate
        /// a new <paramref name="e"/> object from code.
        /// </returns>
        internal static CodeTypeMemberCollection BuildCodeStatements(
            OpenXmlPart part,
            NamespaceAliasOptions opts,
            IDictionary<string, int> typeCounts,
            ISet<string> namespaces,
            ISet<Uri> targets,
            out string partName) 
        {
            // Argument validation
            if (part is null) throw new ArgumentNullException(nameof(part));
            if (opts is null) throw new ArgumentNullException(nameof(opts));
            if (typeCounts is null) throw new ArgumentNullException(nameof(typeCounts));
            if (namespaces is null) throw new ArgumentNullException(nameof(namespaces));
            if (targets is null) throw new ArgumentNullException(nameof(targets));

            var result = new CodeTypeMemberCollection();
            var partType = part.GetType();
            var partTypeName = partType.Name;
            var partTypeFullName = partType.FullName;
            var reType = part.RootElement?.GetType();
            var memberNamePrefix = $"Generate{partType.Name}";
            var parts = part.Parts != null ? new List<IdPartPair>(part.Parts) : new List<IdPartPair>();
            var localTypeCount = new Dictionary<Type, int>();
            var localTargets = new Dictionary<Uri, string>(new UriEqualityComparer());
            CodeMemberMethod method = null;
            CodeMethodReferenceExpression referenceExpression = null;
            CodeMethodInvokeExpression invokeExpression = null;
            Type tmpType = null;
            string partVarName = null;
            string subPartTypeName = null;

            // Make sure that the namespace for the current part is captured
            namespaces.Add(partType.Namespace);

            if (typeCounts.ContainsKey(partTypeFullName))
            {
                partName = String.Concat(memberNamePrefix, typeCounts[partTypeFullName]++);
            }
            else
            {
                partName = memberNamePrefix;
                typeCounts.Add(partTypeFullName, 1);
            }

            // Setup the first method
            method = new CodeMemberMethod()
            {
                Name = partName,
                Attributes = MemberAttributes.Private | MemberAttributes.Final,
                ReturnType = new CodeTypeReference()
            };
            method.Parameters.Add(new CodeParameterDeclarationExpression(partTypeName, methodParamName));

            // Add blank code line
            void addBlankLine() => method.Statements.Add(new CodeSnippetStatement(String.Empty));

            // Loop through and build all of the child parts of the current Part
            foreach (var p in parts)
            {
                // Check to see if the current part has been visited/created previously
                if (!targets.Add(p.OpenXmlPart.Uri))
                {
                    // TODO: If the current part has been previously visited, 
                    // TODO: the part's original parent needs to be found and referenced
                    // TODO: in the resulting code statements. For now, just continue the loop.
                    continue;
                }

                tmpType = p.OpenXmlPart.GetType();
                partVarName = tmpType.Name.ToCamelCase();
                subPartTypeName = tmpType.GetObjectTypeName(opts.Order);

                if (typeCounts.ContainsKey(tmpType.FullName))
                {
                    partVarName = String.Concat(partVarName, typeCounts[tmpType.FullName]++);
                }
                else
                {
                    typeCounts.Add(tmpType.FullName, 1);
                }

                // Initialize the part in OpenXml's own unique way
                referenceExpression = new CodeMethodReferenceExpression(
                    new CodeVariableReferenceExpression(methodParamName), "AddNewPart",
                    new CodeTypeReference(subPartTypeName));

                invokeExpression = new CodeMethodInvokeExpression(referenceExpression,
                    new CodePrimitiveExpression(p.RelationshipId));

                method.Statements.Add(new CodeVariableDeclarationStatement(subPartTypeName, partVarName, invokeExpression));

                // Create the part generation method
                result.AddRange(BuildCodeStatements(p.OpenXmlPart, opts, typeCounts, namespaces, targets, out string genMethod));
                
                if (!String.IsNullOrWhiteSpace(genMethod))
                {
                    referenceExpression = new CodeMethodReferenceExpression(
                        new CodeThisReferenceExpression(), genMethod);
                    invokeExpression = new CodeMethodInvokeExpression(referenceExpression,
                        new CodeVariableReferenceExpression(partVarName));
                    method.Statements.Add(invokeExpression);
                }
                addBlankLine();                
            }

            // Code part elements next
            if (part.RootElement is null)
            {
                // If the root element is not present (aka: null) then perform a simple feed
                // dump of the part in the current method
                const string memName = "mem";
                const string b64Name = "base64";

                // Add the necessary namespaces by hand to the namespace set
                namespaces.Add("System");
                namespaces.Add("System.IO");

                using (var partStream = part.GetStream(FileMode.Open, FileAccess.Read))
                {
                    using (var mem = new MemoryStream())
                    {
                        partStream.CopyTo(mem);
                        method.Statements.Add(new CodeVariableDeclarationStatement(typeof(string), b64Name,
                            new CodePrimitiveExpression(Convert.ToBase64String(mem.ToArray()))));
                    }
                }
                addBlankLine();

                var fromBase64 = new CodeMethodReferenceExpression(new CodeVariableReferenceExpression("Convert"),
                    "FromBase64String");
                var invokeFromBase64 = new CodeMethodInvokeExpression(fromBase64, new CodeVariableReferenceExpression("base64"));
                var createStream = new CodeObjectCreateExpression(new CodeTypeReference("MemoryStream"),
                    invokeFromBase64, new CodePrimitiveExpression(false));
                var feedData = new CodeMethodReferenceExpression(new CodeVariableReferenceExpression(methodParamName), "FeedData");
                var invokeFeedData = new CodeMethodInvokeExpression(feedData, new CodeVariableReferenceExpression(memName));
                var disposeMem = new CodeMethodReferenceExpression(new CodeVariableReferenceExpression(memName), "Dispose");
                var invokeDisposeMem = new CodeMethodInvokeExpression(feedData, new CodeVariableReferenceExpression(memName));

                // Setup the try statement
                var tryAndCatch = new CodeTryCatchFinallyStatement();
                tryAndCatch.TryStatements.Add(invokeFeedData);
                tryAndCatch.FinallyStatements.Add(invokeDisposeMem);

                // Put all of the pieces together
                method.Statements.Add(new CodeVariableDeclarationStatement("Stream", memName, createStream));
                method.Statements.Add(tryAndCatch);
            }
            else
            {
                var rootElementType = part.RootElement?.GetType();

                // Build the element details of the requested part for the current method
                method.Statements.AddRange(
                    OpenXmlElementExtensions.BuildCodeStatements(part.RootElement,
                    opts, localTypeCount, namespaces, out string rootElementVar));

                // Now finish up the current method by assigning the OpenXmlElement code statements
                // back to the appropriate property of the part parameter
                if (rootElementType != null && !String.IsNullOrWhiteSpace(rootElementVar))
                {
                    foreach (var paramProp in partType.GetProperties())
                    {
                        if (paramProp.PropertyType == rootElementType)
                        {
                            var varRef = new CodeVariableReferenceExpression(rootElementVar);
                            var paramRef = new CodeVariableReferenceExpression(methodParamName);
                            var propRef = new CodePropertyReferenceExpression(paramRef, paramProp.Name);
                            method.Statements.Add(new CodeAssignStatement(propRef, varRef));
                            break;
                        }
                    }
                }
            }

            // Add the appropriate code statements if the current part
            // contains any hyperlink relationships
            if (part.HyperlinkRelationships.Count() > 0)
            {
                // Add a line break first for easier reading
                addBlankLine();
                method.Statements.AddRange(
                    part.HyperlinkRelationships.BuildHyperlinkRelationshipStatements(partName));
            }

            // Add the appropriate code statements if the current part
            // contains any non-hyperlink external relationships
            if (part.ExternalRelationships.Count() > 0)
            {
                // Add a line break first for easier reading
                addBlankLine();
                method.Statements.AddRange(
                    part.ExternalRelationships.BuildExternalRelationshipStatements(partName));
            }

            result.Add(method);
            return result;
        }

        /// <summary>
        /// Creates a collection of code statements that describe how to add external relationships to 
        /// a <see cref="OpenXmlPartContainer"/> object.
        /// </summary>
        /// <param name="relationships">
        /// The collection of <see cref="ExternalRelationship"/> objects to build the code statements for.
        /// </param>
        /// <param name="parentName">
        /// The name of the <see cref="OpenXmlPartContainer"/> object that the external relationship
        /// assignments should be for.
        /// </param>
        /// <returns>
        /// A collection of code statements that could be used to generate and assign new
        /// <see cref="ExternalRelationship"/> objects to a <see cref="OpenXmlPartContainer"/> object.
        /// </returns>
        internal static CodeStatementCollection BuildExternalRelationshipStatements(
            this IEnumerable<ExternalRelationship> relationships, string parentName)
        {
            if (String.IsNullOrWhiteSpace(parentName)) throw new ArgumentNullException(nameof(parentName));

            var result = new CodeStatementCollection();

            // Return an empty code statement collection if the hyperlinks parameter is empty.
            if (relationships.Count() == 0) return result;

            CodeObjectCreateExpression createExpression;
            CodeMethodReferenceExpression methodReferenceExpression;
            CodeMethodInvokeExpression invokeExpression;
            CodePrimitiveExpression param;
            CodeTypeReference typeReference;

            foreach (var ex in relationships)
            {
                // Need special care to create the uri for the current object.
                typeReference = new CodeTypeReference(ex.Uri.GetType());
                param = new CodePrimitiveExpression(ex.Uri.ToString());
                createExpression = new CodeObjectCreateExpression(typeReference, param);

                // Create the AddHyperlinkRelationship statement
                methodReferenceExpression = new CodeMethodReferenceExpression(
                    new CodeVariableReferenceExpression(parentName),
                    "AddExternalRelationship");
                invokeExpression = new CodeMethodInvokeExpression(methodReferenceExpression,
                    createExpression,
                    new CodePrimitiveExpression(ex.IsExternal),
                    new CodePrimitiveExpression(ex.Id));
                result.Add(invokeExpression);
            }
            return result;
        }

        /// <summary>
        /// Creates a collection of code statements that describe how to add hyperlink relationships to 
        /// a <see cref="OpenXmlPartContainer"/> object.
        /// </summary>
        /// <param name="hyperlinks">
        /// The collection of <see cref="HyperlinkRelationship"/> objects to build the code statements for.
        /// </param>
        /// <param name="parentName">
        /// The name of the <see cref="OpenXmlPartContainer"/> object that the hyperlink relationship
        /// assignments should be for.
        /// </param>
        /// <returns>
        /// A collection of code statements that could be used to generate and assign new
        /// <see cref="HyperlinkRelationship"/> objects to a <see cref="OpenXmlPartContainer"/> object.
        /// </returns>
        internal static CodeStatementCollection BuildHyperlinkRelationshipStatements(
            this IEnumerable<HyperlinkRelationship> hyperlinks, string parentName)
        {
            if (String.IsNullOrWhiteSpace(parentName)) throw new ArgumentNullException(nameof(parentName));

            var result = new CodeStatementCollection();

            // Return an empty code statement collection if the hyperlinks parameter is empty.
            if (hyperlinks.Count() == 0) return result;

            CodeObjectCreateExpression createExpression;
            CodeMethodReferenceExpression methodReferenceExpression;
            CodeMethodInvokeExpression invokeExpression;
            CodePrimitiveExpression param;
            CodeTypeReference typeReference;

            foreach (var hl in hyperlinks)
            {
                // Need special care to create the uri for the current object.
                typeReference = new CodeTypeReference(hl.Uri.GetType());
                param = new CodePrimitiveExpression(hl.Uri.ToString());
                createExpression = new CodeObjectCreateExpression(typeReference, param);

                // Create the AddHyperlinkRelationship statement
                methodReferenceExpression = new CodeMethodReferenceExpression(
                    new CodeVariableReferenceExpression(parentName),
                    "AddHyperlinkRelationship");
                invokeExpression = new CodeMethodInvokeExpression(methodReferenceExpression,
                    createExpression,
                    new CodePrimitiveExpression(hl.IsExternal),
                    new CodePrimitiveExpression(hl.Id));
                result.Add(invokeExpression);
            }
            return result;
        }

        #endregion

        /// <summary>
        /// Equality comparer class used to ensure that circular part references are avoided when trying
        /// to build source code for <see cref="OpenXmlPart"/> objects.
        /// </summary>
        private sealed class UriEqualityComparer : EqualityComparer<Uri>
        {
            #region Public Instance Methods 
            
            /// <inheritdoc/>
            public override bool Equals(Uri x, Uri y)
            {
                if (x == null || y == null) return false;
                return x == y;
            }

            /// <inheritdoc/>
            public override int GetHashCode(Uri obj)
            {
                return obj.ToString().GetHashCode();
            }

            #endregion
        }
    }
}