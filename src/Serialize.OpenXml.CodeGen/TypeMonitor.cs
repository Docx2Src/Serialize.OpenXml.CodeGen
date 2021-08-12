﻿/* MIT License

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

using Serialize.OpenXml.CodeGen.Extentions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace Serialize.OpenXml.CodeGen
{
    /// <summary>
    /// Class designed to monitor the variable names generated by a single type.
    /// </summary>
    /// <remarks>
    /// This will be used only for code generation of
    /// <see cref="DocumentFormat.OpenXml.OpenXmlElement"/> objects.
    /// </remarks>
    public class TypeMonitor : IReadOnlyDictionary<string, bool>
    {
        #region Private Instance Fields

        /// <summary>
        /// Holds the <see cref="Type"/> that the current instance will
        /// represent for its lifetime.
        /// </summary>
        private readonly Type _type;

        /// <summary>
        /// Holds the counter used when unique variable names are generated.
        /// </summary>
        private int _uniqueCount = 0;

        /// <summary>
        /// The object <see cref="Dictionary{TKey, TValue}"/> containing all of the
        /// variable names that have already been generated and their 'consumed'
        /// indicators.
        /// </summary>
#pragma warning disable IDE0044 // Add readonly modifier
        private Dictionary<string, bool> _values = new Dictionary<string, bool>();
#pragma warning restore IDE0044 // Add readonly modifier

        #endregion

        #region Public Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TypeMonitor"/> class with
        /// the type that it will represent.
        /// </summary>
        /// <param name="t">
        /// The <see cref="Type"/> that this new object will represent.
        /// </param>
        public TypeMonitor(Type t)
        {
            _type = t ?? throw new ArgumentNullException(nameof(t));
        }

        #endregion

        #region Internal Static Properties

        /// <summary>
        /// Indicates whether or not to use unique/numbered variable names
        /// for the current request.
        /// </summary>
        internal static bool UseUniqueVariableNames { get; set; } = false;

        #endregion

        #region Public Instance Properties

        /// <inheritdoc/>
        public bool this[string key]
        {
            get => _values[key];
            set => _values[key] = value;
        }

        /// <inheritdoc/>
        public IEnumerable<string> Keys => _values.Keys;

        /// <inheritdoc/>
        public IEnumerable<bool> Values => _values.Values;

        /// <inheritdoc/>
        public int Count => UseUniqueVariableNames ? _uniqueCount : _values.Count;

        /// <summary>
        /// Gets the <see cref="Type"/> that the current object represents.
        /// </summary>
        public Type Type { get => _type; }

        #endregion

        #region Protected Instance Properties

        /// <summary>
        /// Gets the underlying <see cref="IDictionary{TKey, TValue}"/> object
        /// for the current instance.
        /// </summary>
        protected IDictionary<string, bool> Dictionary { get { return _values; } }

        #endregion

        #region Public Instance Methods

        /// <inheritdoc/>
        public bool ContainsKey(string key) => _values.ContainsKey(key);

        /// <summary>
        /// Retrieves an existing or new variable name, depending on what is
        /// available at the time, to generate new code statements with.
        /// </summary>
        /// <param name="namespaces">
        /// Collection <see cref="IDictionary{TKey, TValue}"/> used to keep
        /// track of all openxml namespaces used during the process.
        /// </param>
        /// <param name="varName">
        /// The variable name to use.
        /// </param>
        /// <returns>
        /// <see langword="true"/> if <paramref name="varName"/> already exists
        /// in existing generated code with constructor statements; otherwise
        /// <see langword="false"/> is returned, indicating that no existing
        /// generated code yet exists for <paramref name="varName"/>.
        /// </returns>
        public bool GetVariableName(IDictionary<string, string> namespaces, out string varName)
        {
            if (UseUniqueVariableNames)
            {

                varName = Type.GenerateVariableName(_uniqueCount++, namespaces);
                return false;
            }
            else
            {
                return CreateAndTrackVariableName(namespaces, out varName);
            }
        }

        /// <inheritdoc/>
        public bool TryGetValue(string key, out bool value)
        {
            return _values.TryGetValue(key, out value);
        }

        #endregion

        #region Protected Instance Methods

        /// <summary>
        /// Adds a new key/value pair to this collection.
        /// </summary>
        /// <param name="key">The key of the new pair.</param>
        /// <param name="value">The value of the new pair.</param>
        protected void Add(string key, bool value) => _values.Add(key, value);

        /// <summary>
        /// Responible for retrieving an existing or new variable name,
        /// depending on what is available at the time, to generate new
        /// code statements with.
        /// </summary>
        /// <param name="namespaces">
        /// Collection <see cref="IDictionary{TKey, TValue}"/> used to keep
        /// track of all openxml namespaces used during the process.
        /// </param>
        /// <param name="varName">
        /// The variable name to use.
        /// </param>
        /// <returns>
        /// <see langword="true"/> if <paramref name="varName"/> already exists
        /// in existing generated code with constructor statements; otherwise
        /// <see langword="false"/> is returned, indicating that no existing
        /// generated code yet exists for <paramref name="varName"/>.
        /// </returns>
        protected virtual bool CreateAndTrackVariableName(
            IDictionary<string, string> namespaces, out string varName)
        {
            if (Count > 0 && this.Values.Any(v => v))
            {
                varName = this.LastOrDefault(v => v.Value).Key;
                this[varName] = false;
                return true;
            }

            int tries = _values.Count;
            varName = Type.GenerateVariableName(tries, namespaces);
            Add(varName, false);
            return false;
        }

        #endregion

        #region Private Instance Methods;

        /// <inheritdoc/>
        IEnumerator<KeyValuePair<string, bool>> IEnumerable<KeyValuePair<string, bool>>.GetEnumerator()
        {
            return ((IEnumerable<KeyValuePair<string, bool>>)_values).GetEnumerator();
        }

        /// <inheritdoc/>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return _values.GetEnumerator();
        }

        #endregion
    }
}
