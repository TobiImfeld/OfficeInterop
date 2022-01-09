using System;

namespace WordVba
{
    internal class TypeCompat
    {
        public static bool IsPrimitive(object v)
        {
            return v.GetType().IsPrimitive;
        }
        public static bool IsSubclassOf(Type t, Type c)
        {
            return t.IsSubclassOf(c);
        }

        internal static bool IsGenericType(Type t)
        {
            return t.IsGenericType;
        }

        public static object GetPropertyValue(object v, string name)
        {
            return v.GetType().GetProperty(name).GetValue(v, null);
        }
    }
}
