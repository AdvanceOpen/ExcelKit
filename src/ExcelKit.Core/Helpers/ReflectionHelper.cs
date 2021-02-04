using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ExcelKit.Core.Attributes;
using ExcelKit.Core.Helpers;
using ExcelKit.Core.Infrastructure.Exceptions;

namespace ExcelKit.Core.Infrastructure.Reflections
{
	/// <summary>
	/// 类型反射辅助类
	/// </summary>
	internal class ReflectionHelper
	{
		#region Fields

		/// <summary>
		/// 类型字段缓存(Key:类型的Type的FullName  Value：缓存的字段信息)
		/// </summary>
		static ConcurrentDictionary<string, List<PropertyInfo>> _typeProps = new ConcurrentDictionary<string, List<PropertyInfo>>();

		/// <summary>
		/// 定义了ExcelKit特性的字段信息
		/// </summary>
		static ConcurrentDictionary<string, List<(float sort, PropertyInfo prop, ExcelKitAttribute attribute)>> _excelKitPropInfo =
		new ConcurrentDictionary<string, List<(float sort, PropertyInfo prop, ExcelKitAttribute attribute)>>();

		/// <summary>
		/// 新实例
		/// </summary>
		public static ReflectionHelper NewInstance => new ReflectionHelper();

		#endregion

		#region Construct

		/// <summary>
		/// 构造函数
		/// </summary>
		private ReflectionHelper()
		{

		}

		#endregion

		#region 是否是可空类型

		/// <summary>
		/// 是否是可空类型
		/// </summary>
		/// <param name="type"></param>
		/// <returns></returns>
		bool IsNullableType(Type type)
		{
			return type.IsGenericType && type.GetGenericTypeDefinition().Equals(typeof(Nullable<>));
		}

		#endregion

		#region 获取对象缓存的字段信息

		/// <summary>
		/// 获取对象缓存的字段信息，对象
		/// </summary>
		/// <param name="obj">要获取的对象</param>
		public List<PropertyInfo> GetCachePropertyInfo(object obj)
		{
			Inspector.NotNull(obj, "object can not be null");

			var fullName = obj.GetType().FullName;
			if (_typeProps.ContainsKey(fullName))
			{
				return _typeProps[fullName];
			}

			return _typeProps[fullName] = obj.GetType().GetProperties().ToList();
		}

		#endregion

		#region 获取对象缓存的字段信息

		/// <summary>
		/// 获取对象缓存的字段信息,泛型
		/// </summary>
		public List<PropertyInfo> GetCachePropertyInfo<T>() where T : class, new()
		{
			var fullName = typeof(T).FullName;
			if (_typeProps.ContainsKey(fullName))
			{
				return _typeProps[fullName];
			}

			return _typeProps[fullName] = typeof(T).GetProperties().ToList();
		}

		#endregion

		#region 获取已排序且定义了特性的字段

		/// <summary>
		/// 获取排序后且定义了特性的字段
		/// </summary>
		/// <typeparam name="T">泛型</typeparam>
		/// <returns></returns>
		public List<(float sort, PropertyInfo prop, ExcelKitAttribute attr)> GetSortedExcelKitProps<T>() where T : class, new()
		{
			var fullName = typeof(T).FullName;
			if (_excelKitPropInfo.ContainsKey(fullName))
			{
				return _excelKitPropInfo[fullName];
			}

			//调整字段排序
			var propsInfo = new List<(float sort, PropertyInfo prop, ExcelKitAttribute attribute)>();
			foreach (var prop in this.GetCachePropertyInfo<T>())
			{
				var attribute = prop.GetCustomAttribute<ExcelKitAttribute>();
				if (attribute == null || attribute.IsIgnore) { continue; }
				attribute.Code = attribute.Code ?? prop.Name;

				propsInfo.Add((attribute.Sort, prop, attribute));
			}

			return _excelKitPropInfo[fullName] = propsInfo.OrderBy(t => t.sort).ToList();
		}

		/// <summary>
		/// 获取排序后且定义了特性的字段(去除指定了IsOnlyWriteIgnore的)
		/// </summary>
		/// <typeparam name="T">泛型</typeparam>
		/// <returns></returns>
		public List<(float sort, PropertyInfo prop, ExcelKitAttribute attr)> GetSortedExportProps<T>() where T : class, new()
		{
			var sortedProps = GetSortedExcelKitProps<T>();
			sortedProps.RemoveAll(t => t.attr.IsOnlyIgnoreWrite);
			return sortedProps;
		}

		/// <summary>
		/// 获取排序后且定义了特性的字段(去除指定了IsOnlyReadIgnore的)
		/// </summary>
		/// <typeparam name="T">泛型</typeparam>
		/// <returns></returns>
		public List<(float sort, PropertyInfo prop, ExcelKitAttribute attr)> GetSortedReadProps<T>() where T : class, new()
		{
			var sortedProps = GetSortedExcelKitProps<T>();
			sortedProps.RemoveAll(t => t.attr.IsOnlyIgnoreRead);
			return sortedProps;
		}

		#endregion

		#region 设置泛型对象的属性值

		public void SetValue<T>(PropertyInfo prop, ref T obj, string value, bool allowNull) where T : class, new()
		{
			Inspector.NotNull(obj, "SetValue的泛型对象不能为空");
			Inspector.NotNull(prop, "SetValue的PropertyInfo对象不能为空");

			if (allowNull && string.IsNullOrWhiteSpace(value))
			{
				return;
			}

			//可空类型获取
			bool isNullableType = IsNullableType(prop.PropertyType);
			var thisType = isNullableType ? Nullable.GetUnderlyingType(prop.PropertyType) : prop.PropertyType;

			//枚举类型
			if (prop.PropertyType.IsEnum)
			{
				var _data = value;
				//是否为数字
				if (StringHelper.IsNumber(_data))
				{
					prop.SetValue(obj, Enum.Parse(prop.PropertyType, _data, true), null);
				}
				else
				{
					var enums = EnumHelper.GetEnumInfo(prop.PropertyType);
					if (!enums.Exists(t => t.EnumDesc == _data))
					{
						throw new ExcelKitException($"该字段为枚举数据项，数据源中的项【{_data}】无效");
					}
					prop.SetValue(obj, Enum.Parse(prop.PropertyType, enums.FirstOrDefault(t => t.EnumDesc == _data).EnumName, true), null);
				}
			}
			//时间类型
			else if (prop.PropertyType == typeof(DateTime))
			{
				var status = DateTime.TryParse(value, out DateTime dateTime);
				var convertedTime = status ? dateTime : DateTime.FromOADate(System.Convert.ToDouble(value));
				prop.SetValue(obj, convertedTime, null);
			}
			//数值类型优先处理
			else if (thisType == typeof(System.Single))
			{
				prop.SetValue(obj, Convert.ToSingle(value), null);
			}
			//数值类型优先处理
			else if (thisType == typeof(System.Double))
			{
				prop.SetValue(obj, Convert.ToDouble(Convert.ToDouble(value)), null);
			}
			//decimal类型比下面的值类型优先处理（这种处理是为了兼容读取出来的是含E科学计数法，如：0.0145647787897897，读取出来就是）
			else if (thisType == typeof(System.Decimal))
			{
				prop.SetValue(obj, Convert.ToDecimal(Convert.ToDouble(value)), null);
			}
			//值类型(一定要用thisType，因为可能是可空类型)
			else if (prop.PropertyType.IsValueType)
			{
				var methodInfo = thisType.GetMethod("Parse", new Type[] { typeof(string) });
				if (methodInfo != null)
				{
					object convertValue = methodInfo.Invoke(null, new object[] { value.Replace("'", "") });
					prop.SetValue(obj, convertValue, null);
				}
			}
			else
			{
				prop.SetValue(obj, value);
			}
		}

		#endregion
	}
}
