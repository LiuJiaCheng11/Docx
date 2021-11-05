using Microsoft.Extensions.Primitives;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Docx.Core
{
    public static class ObjectExtensions
    {
        public static bool HasProperty<T>(this T @this, string propertyName)
        {
            return @this.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic) != (PropertyInfo)null;
        }

        public static bool IsNullOrEmpty<T>(this IEnumerable<T> source)
        {
            if (source != null)
                return !source.Any<T>();
            return true;
        }

        public static bool IsNullOrWhiteSpace(this string str)
        {
            return string.IsNullOrWhiteSpace(str);
        }

        /// <summary>转为字符串或空白字符串</summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public static string TryToStringOrEmpty(this object source)
        {
            if (source == null)
                return string.Empty;
            if (source is string str)
                return str;
            if (source is StringValues)
                return source.ToString();
            return string.Empty;
        }

        public static object Getter(this object data, string fieldName)
        {
            object obj = (object)null;
            if (data == null)
                return (object)null;
            if (fieldName.IsNullOrWhiteSpace())
                return (object)null;
            string[] strArray = fieldName.Split('.');
            if (strArray.Length != 1 && !data.HasProperty<object>(fieldName))
            {
                IDictionary<string, object> dictionary = data as IDictionary<string, object>;
                if (dictionary == null || !dictionary.ContainsKey(fieldName))
                {
                    string pattern = "\\[\\d+\\]";
                    for (int index1 = 0; index1 < strArray.Length; ++index1)
                    {
                        string str = Regex.Replace(strArray[index1], pattern, string.Empty);
                        IEnumerable source1 = obj as IEnumerable;
                        if (source1 != null && source1 != null)
                        {
                            Type type = obj.GetType();
                            if (type != typeof(string) && !typeof(IDictionary<string, object>).IsAssignableFrom(type) && (!typeof(IDictionary).IsAssignableFrom(type) && !typeof(JToken).IsAssignableFrom(type)) || typeof(JArray).IsAssignableFrom(type))
                            {
                                IEnumerable<object> source2 = source1.Cast<object>();
                                if (!source2.IsNullOrEmpty<object>())
                                {
                                    int index2 = 0;
                                    string input = strArray[index1 - 1];
                                    if (Regex.IsMatch(input, pattern))
                                        index2 = Convert.ToInt32(Regex.Match(input, pattern).Value.TrimStart('[').TrimEnd(']'));
                                    if (index2 != 0 && source2.Count<object>() <= index2)
                                    {
                                        obj = (object)null;
                                        break;
                                    }
                                    data = source2.ElementAt(index2);
                                }
                            }
                        }
                        obj = ObjectExtensions.GetFieldNameValue(data, str);
                        if (obj == null && index1 > 0 && data.GetType() == typeof(string))
                        {
                            JObject jobject = (JObject)null;
                            try
                            {
                                jobject = JsonConvert.DeserializeObject<JObject>(data.TryToStringOrEmpty(), new JsonSerializerSettings()
                                {
                                    Formatting = Newtonsoft.Json.Formatting.None,
                                    PreserveReferencesHandling = PreserveReferencesHandling.None,
                                    ReferenceLoopHandling = ReferenceLoopHandling.Ignore,
                                    ContractResolver = (IContractResolver)new DefaultContractResolver(),
                                    NullValueHandling = NullValueHandling.Ignore,
                                    MissingMemberHandling = MissingMemberHandling.Ignore,
                                    DateFormatHandling = DateFormatHandling.IsoDateFormat,
                                    DateFormatString = "yyyy-MM-dd HH:mm:ss"
                                });
                            }
                            catch
                            {
                            }
                            if (jobject != null)
                                obj = (object)jobject.SelectToken(str);
                        }
                        if (obj != null)
                        {
                            if (index1 < strArray.Length - 1)
                            {
                                data = obj;
                            }
                            else
                            {
                                JValue jvalue = obj as JValue;
                                if (jvalue != null)
                                    obj = jvalue.Value;
                            }
                        }
                        else
                            break;
                    }
                    goto label_29;
                }
            }
            obj = ObjectExtensions.GetFieldNameValue(data, fieldName);
        label_29:
            return obj;
        }

        /// <summary>A T extension method that gets property value.</summary>
        /// <typeparam name="T">Generic type parameter.</typeparam>
        /// <param name="this">The @this to act on.</param>
        /// <param name="propertyName">Name of the property.</param>
        /// <returns>The property value.</returns>
        public static object GetPropertyValue<T>(this T @this, string propertyName)
        {
            PropertyInfo property = @this.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            if (property != (PropertyInfo)null)
                return property.GetValue((object)@this, (object[])null);
            return (object)null;
        }

        private static object GetFieldNameValue(object data, string fieldName)
        {
            IDictionary<string, object> dictionary1 = data as IDictionary<string, object>;
            if (dictionary1 != null)
            {
                object obj;
                if (dictionary1.TryGetValue(fieldName, out obj))
                    return obj;
                return (object)null;
            }
            IDictionary dictionary2 = data as IDictionary;
            if (dictionary2 != null)
            {
                if (dictionary2.Contains((object)fieldName))
                    return dictionary2[(object)fieldName];
                return (object)null;
            }
            JObject jobject = data as JObject;
            if (jobject != null)
                return (object)jobject.SelectToken(fieldName);
            JArray source = data as JArray;
            if (source != null)
                return (object)(source.ElementAt(0) as JObject).SelectToken(fieldName);
            JProperty jproperty = data as JProperty;
            if (jproperty != null)
                return (object)jproperty.Value;
            if (data != null)
                return data.GetPropertyValue<object>(fieldName);
            return (object)null;
        }

        /// <summary>Clone Object</summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source"></param>
        /// <returns></returns>
        public static T Clone<T>(this T source) => JsonConvert.DeserializeObject<T>(JsonConvert.SerializeObject((object)source));

        /// <summary>
        /// object带格式的ToString
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="format"></param>
        /// <param name="provider"></param>
        /// <returns></returns>
        public static string ToString(this object obj, string format, IFormatProvider provider = null)
        {
            try
            {
                if (obj is IFormattable)
                {
                    var temp = "{0:" + format + "}";
                    var result = string.Format(provider ?? new DefaultFormatter(), temp, obj);
                    return result;
                }

                return obj != null ? obj.ToString() : string.Empty;
            }
            catch (FormatException)
            {
                return format;
            }
        }
    }
}
