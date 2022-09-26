using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ReportsAPISite.Extensions
{
    public static class DictionaryExtensions
    {
        public static Dictionary<string, object> TrimStrings(this Dictionary<string, object> source)
        {
            var changes = new Dictionary<string, object>();

            foreach (var arg in source.Where(a => a.Value != null))
            {
                var type = arg.Value.GetType();
                if (IsEnumerable(type))
                {
                    changes[arg.Key] = TrimEnumerable((IEnumerable)arg.Value);
                }
                else if (IsComplexObject(type))
                {
                    changes[arg.Key] = TrimObject(arg.Value);
                }
            }
            foreach (var change in changes)
            {
                source[change.Key] = change.Value;
            }

            return source;
        }

        private static IEnumerable TrimEnumerable(IEnumerable value)
        {
            var enumerable = value as object[] ?? value.Cast<object>().ToArray();
            return enumerable.OfType<string>().Any() ?
                enumerable.Cast<string>().Select(s => s == null
                    ? null
                    : s.Trim())
                : enumerable.Select(TrimObject);
        }

        private static bool IsEnumerable(Type t)
        {
            return t.IsAssignableFrom(typeof(IEnumerable));
        }

        private static bool IsComplexObject(Type value)
        {
            return value.IsClass && !value.IsArray;
        }

        private static object TrimObject(object argValue)
        {
            if (argValue == null) return null;
            var argType = argValue.GetType();
            if (IsEnumerable(argType))
            {
                TrimEnumerable((IEnumerable)argValue);
            }
            var s = argValue as string;
            if (s != null)
            {
                return s.Trim();
            }
            if (!IsComplexObject(argType))
            {
                return argValue;
            }
            var props = argType
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(prop => prop.PropertyType == typeof(string))
                .Where(prop => prop.GetIndexParameters().Length == 0)
                .Where(prop => prop.CanWrite && prop.CanRead);

            foreach (var prop in props)
            {
                var value = (string)prop.GetValue(argValue, null);
                if (value != null)
                {
                    value = value.Trim();
                    prop.SetValue(argValue, value, null);
                }
            }
            return argValue;
        }
    }
}