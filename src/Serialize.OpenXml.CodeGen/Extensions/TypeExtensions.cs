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

using DocumentFormat.OpenXml;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Serialize.OpenXml.CodeGen.Extentions
{
    /// <summary>
    /// Collection of extension methods for the <see cref="Type"/> class
    /// specific to the generating code dom representations of
    /// OpenXml objects.
    /// </summary>
    public static class TypeExtensions
    {
        #region Static Constructors

        /// <summary>
        /// Static constructor.
        /// </summary>
        static TypeExtensions()
        {
            // Now setup the simple type collection.
            var simpleTypes = new Type[]
            {
                typeof(StringValue),
                typeof(OpenXmlSimpleValue<uint>),
                typeof(OpenXmlSimpleValue<int>),
                typeof(OpenXmlSimpleValue<byte>),
                typeof(OpenXmlSimpleValue<sbyte>),
                typeof(OpenXmlSimpleValue<short>),
                typeof(OpenXmlSimpleValue<long>),
                typeof(OpenXmlSimpleValue<ushort>),
                typeof(OpenXmlSimpleValue<ulong>),
                typeof(OpenXmlSimpleValue<float>),
                typeof(OpenXmlSimpleValue<double>),
                typeof(OpenXmlSimpleValue<decimal>),
                typeof(OpenXmlSimpleValue<bool>),
                typeof(OpenXmlSimpleValue<DateTime>)
            };

            SimpleValueTypes = simpleTypes.ToList();
        }

        #endregion

        #region Public Static Properties

        /// <summary>
        /// Gets a collection of <see cref="OpenXmlSimpleValue{T}"/> based types.
        /// </summary>
        /// <remarks>
        /// This is used to identify property types of <see cref="OpenXmlElement"/> objects that
        /// can be initialized with simple values of their base type counterparts.
        /// </remarks>
        public static IReadOnlyList<Type> SimpleValueTypes { get; private set; }

        #endregion

        #region Public Static methods

        /// <summary>
        /// Checks to see if a <see cref="Type"/> exists in a different namespace.
        /// </summary>
        /// <param name="t">
        /// The <see cref="Type"/> with the name to search for.
        /// </param>
        /// <param name="namespaces">
        /// The <see cref="IDictionary{TKey, TValue}"/> of collected namespaces
        /// to search in.
        /// </param>
        /// <returns>
        /// <see langword="true"/> if <paramref name="t"/> exists in a different
        /// namespace; otherwise, <see langword="false"/>.
        /// </returns>
        public static bool ExistsInDifferentNamespace(this Type t,
            IDictionary<string, string> namespaces)
        {
            if (namespaces is null) throw new ArgumentNullException(nameof(namespaces));
            if (namespaces.Count == 0) return false;

            Type tmp;
            bool foundElsewhere = false;
            Type[] classes;
            string tmpName;

            bool getTypesWhere(Type c) =>
                !String.IsNullOrEmpty(c.Namespace) &&
                c.Namespace.Equals(t.Namespace, StringComparison.Ordinal);

            foreach (var ns in namespaces)
            {
                // First scan for the actual type name in other existing
                // namespaces
                tmpName = $"{ns.Key}.{t.Name}";
                tmp = t.Assembly.GetType(tmpName) ?? Type.GetType(tmpName);

                if (tmp != null)
                {
                    foundElsewhere = true;
                    break;
                }
            }

            if (!foundElsewhere)
            {
                // Next, try to scan for other classes in the type's namespace
                // in other namespaces
                classes = Assembly.GetAssembly(t).GetTypes()
                    .Where(getTypesWhere)
                    .ToArray();

                foreach (var ns in namespaces)
                {
                    foreach (var cl in classes)
                    {
                        tmpName = $"{ns.Key}.{cl.Name}";
                        tmp = t.Assembly.GetType(tmpName) ?? Type.GetType(tmpName);

                        if (tmp != null)
                        {
                            foundElsewhere = true;
                            break;
                        }
                    }
                    if (foundElsewhere) break;
                }
            }

            return foundElsewhere;
        }

        /// <summary>
        /// Checks to see if a <see cref="Type"/> exists in a different namespace.
        /// </summary>
        /// <param name="t">
        /// The <see cref="Type"/> with the name to search for.
        /// </param>
        /// <param name="namespaces">
        /// The <see cref="IDictionary{TKey, TValue}"/> of collected namespaces
        /// to search in.
        /// </param>
        /// <param name="alias">
        /// The namespace alias to use, if necessary.
        /// </param>
        /// <returns>
        /// <see langword="true"/> if <paramref name="t"/> exists in a different
        /// namespace; otherwise, <see langword="false"/>.
        /// </returns>
        public static bool ExistsInDifferentNamespace(this Type t,
            IDictionary<string, string> namespaces, out string alias)
        {
            if (namespaces is null) throw new ArgumentNullException(nameof(namespaces));
            if (namespaces.Count == 0)
            {
                alias = String.Empty;
                return false;
            }

            Type tmp;
            bool foundElsewhere = false;
            Type[] classes;
            string tmpName;
            string al = String.Empty;

            bool getTypesWhere(Type c) =>
                !String.IsNullOrEmpty(c.Namespace) &&
                c.Namespace.Equals(t.Namespace, StringComparison.Ordinal);

            // Tries to find an alias to use for the specified type
            string findAlias(Type c)
            {
                foundElsewhere = true;
                string result = null;

                // If the current type is a subclass of OpenXmlElement
                // try to initialize it with a default constructor to
                // get its prefix for the namespace alias
                if (c.IsSubclassOf(typeof(OpenXmlElement)))
                {
                    var ctor = c.GetConstructor(Type.EmptyTypes);

                    if (ctor != null)
                    {
                        var element = c.Assembly.CreateInstance(c.FullName) as OpenXmlElement;
                        result = element.Prefix.ToUpperInvariant();
                    }
                }

                // Create a new alias if the element prefix could not be located.
                if (String.IsNullOrEmpty(result))
                {
                    var sb = new StringBuilder();
                    var ns = c.Namespace;

                    // Use only upper case or numeric characters from the
                    // type's namespace as the new alias.
                    for (int i = 0; i < ns.Length; i++)
                    {
                        if (Char.IsUpper(ns[i]) || Char.IsDigit(ns[i]))
                        {
                            sb.Append(ns[i]);
                        }
                    }
                    result = sb.ToString();
                }
                return result;
            }

            foreach (var ns in namespaces)
            {
                // First scan for the actual type name in other existing
                // namespaces
                tmpName = $"{ns.Key}.{t.Name}";
                tmp = t.Assembly.GetType(tmpName) ?? Type.GetType(tmpName);

                if (tmp != null)
                {
                    al = findAlias(t);
                    break;
                }
            }

            if (!foundElsewhere)
            {
                // Next, try to scan for other classes in the type's namespace
                // in other namespaces
                classes = Assembly.GetAssembly(t).GetTypes()
                    .Where(getTypesWhere)
                    .ToArray();

                foreach (var ns in namespaces)
                {
                    foreach (var cl in classes)
                    {
                        tmpName = $"{ns.Key}.{cl.Name}";
                        tmp = t.Assembly.GetType(tmpName) ?? Type.GetType(tmpName);


                        if (tmp != null)
                        {
                            al = findAlias(cl);
                            break;
                        }
                    }
                    if (foundElsewhere) break;
                }
            }

            alias = al;
            return foundElsewhere;
        }

        /// <summary>
        /// Generates a variable name to use when generating the appropriate
        /// CodeDom objects for a given <see cref="Type"/>.
        /// </summary>
        /// <param name="t">
        /// The <see cref="Type"/> to generate the variable name for.
        /// </param>
        /// <param name="tries">
        /// The number of variables that have been already created for
        /// <paramref name="t"/>.
        /// </param>
        /// <param name="namespaces">
        /// Collection <see cref="IDictionary{TKey, TValue}"/> used to keep
        /// track of all openxml namespaces used during the process.
        /// </param>
        /// <returns>
        /// A new variable name to use to represent <paramref name="t"/>.
        /// </returns>
        public static string GenerateVariableName(
            this Type t,
            int tries,
            IDictionary<string, string> namespaces)
        {
            if (namespaces is null) throw new ArgumentNullException(nameof(namespaces));

            string tmp;  // Hold the generated name
            string nsPrefix = String.Empty;

            // Include the namespace alias as part of the variable name
            if (namespaces.ContainsKey(t.Namespace) &&
                !String.IsNullOrWhiteSpace(namespaces[t.Namespace]))
            {
                nsPrefix = namespaces[t.Namespace].ToLowerInvariant();
            }

            // Simply return the generated name if the current
            // type is not considered generic.
            if (!t.IsGenericType)
            {
                tmp = String.Concat(nsPrefix, t.Name).ToCamelCase();
                if (tries > 0)
                {
                    return String.Concat(tmp, tries);
                }
                return tmp;
            }

            // Include the generic types as part of the var name.
            var sb = new StringBuilder();
            foreach (var item in t.GenericTypeArguments)
            {
                sb.Append(item.Name.RetrieveUpperCaseChars().ToTitleCase());
            }
            tmp = t.Name;

            if (tries > 0)
            {
                return String.Concat(nsPrefix,
                    tmp.Substring(0, tmp.IndexOf("`")),
                    sb.ToString(),
                    tries).ToCamelCase();
            }

            return String.Concat(nsPrefix,
                tmp.Substring(0, tmp.IndexOf("`")),
                sb.ToString()).ToCamelCase();
        }

        /// <summary>
        /// Creates a class name <see cref="String"/> to use when generating
        /// source code.
        /// </summary>
        /// <param name="t">
        /// The <see cref="Type"/> object containing the class name to evaluate.
        /// </param>
        /// <param name="namespaces">
        /// Collection <see cref="IDictionary{TKey, TValue}"/> used to keep track of a
        /// ll openxml namespaces used during the process.
        /// </param>
        /// <param name="order">
        /// The <see cref="NamespaceAliasOrder"/> value to evaluate when building the
        /// appropriate class name.
        /// </param>
        /// <returns>
        /// The class name to use when building a new <see cref="CodeObjectCreateExpression"/>
        /// object.
        /// </returns>
        public static string GetObjectTypeName(this Type t,
            IDictionary<string, string> namespaces, NamespaceAliasOrder order)
        {
            if (!namespaces.ContainsKey(t.Namespace) || order == NamespaceAliasOrder.None)
            {
                return t.FullName;
            }
            if (!String.IsNullOrWhiteSpace(namespaces[t.Namespace]))
            {
                return $"{namespaces[t.Namespace]}.{t.Name}";
            }
            return t.Name;
        }

        /// <summary>
        /// Gets all of the <see cref="PropertyInfo"/> objects that inherit from
        /// the <see cref="OpenXmlSimpleType"/> class in <paramref name="t"/>.
        /// </summary>
        /// <param name="t">
        /// The <see cref="Type"/> object to retrieve the <see cref="PropertyInfo"/>
        /// objects from.
        /// </param>
        /// <returns>
        /// A collection of <see cref="PropertyInfo"/> objects that inherit from
        /// the <see cref="OpenXmlSimpleType"/> class.
        /// </returns>
        /// <remarks>
        /// All necessary OpenXml object properties inherit from the
        /// <see cref="OpenXmlSimpleType"/> class.  This makes it easier to
        /// iterate through.
        /// </remarks>
        public static IReadOnlyList<PropertyInfo> GetOpenXmlSimpleTypeProperties(this Type t)
        {
            return t.GetOpenXmlSimpleTypeProperties(true);
        }

        /// <summary>
        /// Gets all of the <see cref="PropertyInfo"/> objects that inherit from
        /// the <see cref="OpenXmlSimpleType"/> class in <paramref name="t"/>.
        /// </summary>
        /// <param name="t">
        /// The <see cref="Type"/> object to retrieve the <see cref="PropertyInfo"/>
        /// objects from.
        /// </param>
        /// <param name="includeSimpleValueTypes">
        /// Include properties with types that inherit from <see cref="OpenXmlSimpleValue{T}"/>
        /// type.
        /// </param>
        /// <returns>
        /// A collection of <see cref="PropertyInfo"/> objects that inherit from
        /// the <see cref="OpenXmlSimpleType"/> class.
        /// </returns>
        /// <remarks>
        /// All necessary OpenXml object properties inherit from the
        /// <see cref="OpenXmlSimpleType"/> class.  This makes it easier to
        /// iterate through.
        /// </remarks>
        public static IReadOnlyList<PropertyInfo> GetOpenXmlSimpleTypeProperties(
            this Type t, bool includeSimpleValueTypes)
        {
            var props = t.GetProperties();
            var result = new List<PropertyInfo>();

            // Collect all properties that are of type or subclass of
            // OpenXmlSimpleType
            foreach (var p in props)
            {
                if (p.PropertyType.Equals(typeof(OpenXmlSimpleType)) ||
                    p.PropertyType.IsSubclassOf(typeof(OpenXmlSimpleType)))
                {
                    if (includeSimpleValueTypes || !p.PropertyType.IsSimpleValueType())
                    {
                        result.Add(p);
                    }
                }
            }
            // Return the collected
            return result;
        }

        /// <summary>
        /// Gets all of the <see cref="PropertyInfo"/> objects that inherit from
        /// the <see cref="OpenXmlSimpleValue{T}"/> class in <paramref name="t"/>.
        /// </summary>
        /// <param name="t">
        /// The <see cref="Type"/> object to retrieve the <see cref="PropertyInfo"/>
        /// objects from.
        /// </param>
        /// <returns>
        /// A collection of <see cref="PropertyInfo"/> objects that inherit from
        /// the <see cref="OpenXmlSimpleValue{T}"/> class.
        /// </returns>
        /// <remarks>
        /// All necessary OpenXml object properties inherit from the
        /// <see cref="OpenXmlSimpleValue{T}"/> class.  This makes it easier to
        /// iterate through.
        /// </remarks>
        public static IReadOnlyList<PropertyInfo> GetOpenXmlSimpleValuesProperties(this Type t)
        {
            var props = t.GetProperties();
            var result = new List<PropertyInfo>();

            // Collect all properties that are of type or subclass of
            // OpenXmlSimpleType
            foreach (var p in props)
            {
                foreach (var item in SimpleValueTypes)
                {
                    if (p.PropertyType.Equals(item) ||
                        p.PropertyType.IsSubclassOf(item))
                    {
                        result.Add(p);
                        break;
                    }
                }
            }
            // Return the collected properties
            return result;
        }

        /// <summary>
        /// Gets all <see cref="PropertyInfo"/> objects from <paramref name="t"/>
        /// of type <see cref="StringValue"/>.
        /// </summary>
        /// <param name="t">
        /// The <see cref="Type"/> to retrieve the <see cref="PropertyInfo"/> objects
        /// from.
        /// </param>
        /// <returns>
        /// A read only collection of <see cref="PropertyInfo"/> objects with a
        /// <see cref="StringValue"/> property type.
        /// </returns>
        public static IReadOnlyList<PropertyInfo> GetStringValueProperties(this Type t)
            => t.GetProperties().Where(s => s.PropertyType == typeof(StringValue)).ToList();

        /// <summary>
        /// Checks to see if <paramref name="t"/> is EnumValue`1.
        /// </summary>
        /// <param name="t">
        /// The <see cref="PropertyInfo"/> to evaluate.
        /// </param>
        /// <returns>
        /// <see langword="true"/> if the type of <paramref name="t"/>
        /// is EnumValue`1; otherwise, <see langword="false"/>.
        /// </returns>
        public static bool IsEnumValueType(this Type t) =>
          t.Name.Equals("EnumValue`1", StringComparison.Ordinal);

        /// <summary>
        /// Indicates whether or not <paramref name="t"/> is considered
        /// a <see cref="OpenXmlSimpleValue{T}"/> type.
        /// </summary>
        /// <param name="t">
        /// The <see cref="Type"/> to check.
        /// </param>
        /// <returns>
        /// <see langword="true"/> if <paramref name="t"/> is derived
        /// from an OpenXmlSimpleValue type; otherwise, <see langword="false"/>.
        /// </returns>
        public static bool IsSimpleValueType(this Type t)
        {
            foreach (var item in SimpleValueTypes)
            {
                if (t.Equals(item) || t.IsSubclassOf(item))
                {
                    return true;
                }
            }
            return false;
        }

        #endregion
    }
}