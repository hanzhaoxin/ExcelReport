using System;
using System.Collections;
using System.Collections.Generic;

namespace ExcelReport.Meta
{
    public class Container<Element> : IEnumerable<Element> where Element : INamed, new()
    {
        private readonly Dictionary<string, Element> _pairs = new Dictionary<string, Element>();

        public Element this[string name]
        {
            get
            {
                if (!_pairs.ContainsKey(name))
                {
                    var element = Activator.CreateInstance<Element>();
                    element.Name = name;
                    _pairs.Add(element.Name, element);
                }
                return _pairs[name];
            }
        }

        public IEnumerator<Element> GetEnumerator()
        {
            return _pairs.Values.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _pairs.Values.GetEnumerator();
        }
    }
}