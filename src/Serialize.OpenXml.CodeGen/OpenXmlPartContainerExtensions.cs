/* MIT License

Copyright (c) 2021 Ryan Boggs

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
using System.Linq;

namespace Serialize.OpenXml.CodeGen
{
    /// <summary>
    /// Static class that is responsible for generating <see cref="CodeStatement"/>
    /// objects for <see cref="OpenXmlPartContainer"/> derived objects.
    /// </summary>
    public static class OpenXmlPartContainerExtensions
    {
        #region Public Static Methods

        /// <summary>
        /// Creates a collection of code statements that describe how to add external relationships
        /// to a <see cref="OpenXmlPartContainer"/> object.
        /// </summary>
        /// <param name="relationships">
        /// The collection of <see cref="ExternalRelationship"/> objects to build the code
        /// statements for.
        /// </param>
        /// <param name="parentName">
        /// The name of the <see cref="OpenXmlPartContainer"/> object that the external
        /// relationship assignments should be for.
        /// </param>
        /// <returns>
        /// A collection of code statements that could be used to generate and assign new
        /// <see cref="ExternalRelationship"/> objects to a <see cref="OpenXmlPartContainer"/>
        /// object.
        /// </returns>
        public static CodeStatementCollection BuildExternalRelationshipStatements(
            this IEnumerable<ExternalRelationship> relationships, string parentName)
        {
            if (String.IsNullOrWhiteSpace(parentName))
                throw new ArgumentNullException(nameof(parentName));

            return relationships.BuildExternalRelationshipStatements(
                new CodeVariableReferenceExpression(parentName));
        }

        /// <summary>
        /// Creates a collection of code statements that describe how to add external relationships
        /// to a <see cref="OpenXmlPartContainer"/> object.
        /// </summary>
        /// <param name="relationships">
        /// The collection of <see cref="ExternalRelationship"/> objects to build the code
        /// statements for.
        /// </param>
        /// <param name="parent">
        /// The <see cref="CodeExpression"/> object that the generated code statements will
        /// reference to build the external relationship assignments.
        /// </param>
        /// <returns>
        /// A collection of code statements that could be used to generate and assign new
        /// <see cref="ExternalRelationship"/> objects to a <see cref="OpenXmlPartContainer"/>
        /// object.
        /// </returns>
        public static CodeStatementCollection BuildExternalRelationshipStatements(
            this IEnumerable<ExternalRelationship> relationships, CodeExpression parent)
        {
            if (parent is null) throw new ArgumentNullException(nameof(parent));

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
                methodReferenceExpression = new CodeMethodReferenceExpression(parent,
                    "AddExternalRelationship");
                invokeExpression = new CodeMethodInvokeExpression(methodReferenceExpression,
                    new CodePrimitiveExpression(ex.RelationshipType),
                    createExpression,
                    new CodePrimitiveExpression(ex.Id));
                result.Add(invokeExpression);
            }
            return result;
        }

        /// <summary>
        /// Creates a collection of code statements that describe how to add hyperlink
        /// relationships to a <see cref="OpenXmlPartContainer"/> object.
        /// </summary>
        /// <param name="hyperlinks">
        /// The collection of <see cref="HyperlinkRelationship"/> objects to build the code
        /// statements for.
        /// </param>
        /// <param name="parentName">
        /// The name of the <see cref="OpenXmlPartContainer"/> object that the hyperlink
        /// relationship assignments should be for.
        /// </param>
        /// <returns>
        /// A collection of code statements that could be used to generate and assign new
        /// <see cref="HyperlinkRelationship"/> objects to a <see cref="OpenXmlPartContainer"/>
        /// object.
        /// </returns>
        public static CodeStatementCollection BuildHyperlinkRelationshipStatements(
            this IEnumerable<HyperlinkRelationship> hyperlinks, string parentName)
        {
            if (String.IsNullOrWhiteSpace(parentName))
            {
                throw new ArgumentNullException(nameof(parentName));
            }

            return hyperlinks.BuildHyperlinkRelationshipStatements(
                new CodeVariableReferenceExpression(parentName));
        }

        /// <summary>
        /// Creates a collection of code statements that describe how to add hyperlink
        /// relationships to a <see cref="OpenXmlPartContainer"/> object.
        /// </summary>
        /// <param name="hyperlinks">
        /// The collection of <see cref="HyperlinkRelationship"/> objects to build the code
        /// statements for.
        /// </param>
        /// <param name="parent">
        /// The <see cref="CodeExpression"/> object that the generated code statements will
        /// reference to build the hyperlink relationship assignments.
        /// </param>
        /// <returns>
        /// A collection of code statements that could be used to generate and assign new
        /// <see cref="HyperlinkRelationship"/> objects to a <see cref="OpenXmlPartContainer"/>
        /// object.
        /// </returns>
        public static CodeStatementCollection BuildHyperlinkRelationshipStatements(
            this IEnumerable<HyperlinkRelationship> hyperlinks, CodeExpression parent)
        {
            if (parent is null) throw new ArgumentNullException(nameof(parent));

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
                methodReferenceExpression = new CodeMethodReferenceExpression(parent,
                    "AddHyperlinkRelationship");
                invokeExpression = new CodeMethodInvokeExpression(methodReferenceExpression,
                    createExpression,
                    new CodePrimitiveExpression(hl.IsExternal),
                    new CodePrimitiveExpression(hl.Id));
                result.Add(invokeExpression);
            }
            return result;
        }

        /// <summary>
        /// Locates any relationship elements associated with an <see cref="OpenXmlPartContainer"/>
        /// object and generates codes statements for them.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPartContainer"/>
        /// </param>
        /// <param name="varName">
        /// The variable name that the generated code will reference.
        /// </param>
        /// <returns>
        /// A new <see cref="CodeStatementCollection"/> containing all of the code statements
        /// needed to create the relationships associated with <paramref name="part"/>.
        /// </returns>
        /// <remarks>
        /// If <paramref name="part"/> doesn't contain any relationships, then an empty
        /// <see cref="CodeStatementCollection"/> is returned.
        /// </remarks>
        public static CodeStatementCollection GenerateRelationshipCodeStatements(
            this OpenXmlPartContainer part, string varName)
        {
            if (String.IsNullOrEmpty(varName)) throw new ArgumentNullException(nameof(varName));

            return part.GenerateRelationshipCodeStatements(
                new CodeVariableReferenceExpression(varName));
        }

        /// <summary>
        /// Locates any relationship elements associated with an <see cref="OpenXmlPartContainer"/>
        /// object and generates codes statements for them.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPartContainer"/>
        /// </param>
        /// <param name="variable">
        /// The <see cref="CodeExpression"/> variable that the generated code will reference.
        /// </param>
        /// <returns>
        /// A new <see cref="CodeStatementCollection"/> containing all of the code statements
        /// needed to create the relationships associated with <paramref name="part"/>.
        /// </returns>
        /// <remarks>
        /// If <paramref name="part"/> doesn't contain any relationships, then an empty
        /// <see cref="CodeStatementCollection"/> is returned.
        /// </remarks>
        public static CodeStatementCollection GenerateRelationshipCodeStatements(
            this OpenXmlPartContainer part, CodeExpression variable)
        {
            if (variable is null) throw new ArgumentNullException(nameof(variable));

            var result = new CodeStatementCollection();

            // Generate statements for the hyperlink relationships if applicable.
            if (part.HyperlinkRelationships.Count() > 0)
            {
                // Add a line break first for easier reading
                result.AddBlankLine();
                result.AddRange(
                    part.HyperlinkRelationships.BuildHyperlinkRelationshipStatements(variable));
            }

            // Generate statement for the non-hyperlink/external relationships if applicable
            if (part.ExternalRelationships.Count() > 0)
            {
                // Add a line break first for easier reading
                result.AddBlankLine();
                result.AddRange(
                    part.ExternalRelationships.BuildExternalRelationshipStatements(variable));
            }
            return result;
        }

        #endregion
    }
}
