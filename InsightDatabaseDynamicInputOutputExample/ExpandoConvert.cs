using System.Linq;
using Insight.Database;
using System;
using System.Collections.Concurrent;
using System.Reflection;
using System.Reflection.Emit;

namespace InsightDatabaseDynamicInputOutputExample
{
	/// <summary>
	/// Generates a dynamic method to convert an object to a FastExpando.
	/// </summary>
	internal static class ExpandoConvert
	{
		/// <summary>
		/// The cache for the methods.
		/// </summary>
		private static ConcurrentDictionary<Type, Func<FastExpando, object>> _converters = new ConcurrentDictionary<Type, Func<FastExpando, object>>();

		/// <summary>
		/// The default constructor for type T.
		/// </summary>
		private static readonly ConstructorInfo _constructor = typeof(FastExpando).GetConstructor(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic, null, Type.EmptyTypes, null);

		/// <summary>
		/// The FastExpando[] method.
		/// </summary>
		private static readonly MethodInfo _fastExpandoGetValue = typeof(FastExpando).GetMethod("get_Item", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic, null, new Type[] { typeof(string) }, null);

		/// <summary>
		/// Changes the <c>originExpando</c> into the <c>targetType</c>.
		/// 
		/// Limitations: 
		/// 
		/// 1. The <c>targetType</c> must have a public default parameterless constructor.
		/// 
		/// 2. Only the properties with public set accessors on <c>targetType</c> will be mapped from <c>originExpando</c>
		/// </summary>
		/// <param name="originExpando">The <see cref="FastExpando"/> to use for the origin data.</param>
		/// <returns>A new instance of the <c>targetType</c> initialized from the <c>originExpando</c>.</returns>
		public static object ChangeType(FastExpando originExpando, Type targetType)
		{
			if (originExpando == null)
				throw new ArgumentNullException(nameof(originExpando));

			// get the converter for the type
			var converter = _converters.GetOrAdd(targetType, CreateConverter);
			
			return converter(originExpando);
		}

		/// <summary>
		/// Uses IL to create an instance of <c>type</c> using the default constructor
		/// then assign any public properties of the <c>type</c> from <see cref="FastExpando"/>
		/// to the new instance.
		/// </summary>
		/// <param name="type">The type of object to be able to convert.</param>
		/// <returns>A function that can convert that type of <see cref="FastExpando"/> to a <c>type</c>.</returns>
		private static Func<FastExpando, object> CreateConverter(Type type)
		{
			// create a dynamic method
			var dm = new DynamicMethod($"ExpandoGeneratorTo-{type.FullName}", typeof(object), new[] { typeof(FastExpando) }, typeof(ExpandoConvert), true);
			
			// build the il
			var il = dm.GetILGenerator();
			
			// define the output local to store the new type
			var outputLocal = il.DeclareLocal(type);

			// new instance of output type
			il.Emit(OpCodes.Newobj, type.GetConstructor(BindingFlags.Instance | BindingFlags.Public, null, Type.EmptyTypes, null));
			il.Emit(OpCodes.Stloc, outputLocal);

			// for each public field or method, get the value
			// Note: Limitation of only looking at public properties with the ability to be set.
			foreach (var setProperty in type.GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty))
			{
				// set the first argument of the call to the name of the property
				il.Emit(OpCodes.Ldloc, outputLocal);
				il.Emit(OpCodes.Ldarg_0);
				il.Emit(OpCodes.Ldstr, setProperty.Name);

				// call the get value
				il.Emit(OpCodes.Callvirt, _fastExpandoGetValue);

				// cast or unbox the value 
				// depending on whether its a value type or not
				if (setProperty.PropertyType.IsValueType)
				{
					// value types have to be unboxed

					// unbox the type
					il.Emit(OpCodes.Unbox_Any, setProperty.PropertyType);
				}
				else
				{
					// class types have to be casted

					// cast the value to the correct type
					il.Emit(OpCodes.Castclass, setProperty.PropertyType);
				}

				// set the property
				il.Emit(OpCodes.Callvirt, setProperty.SetMethod);
			}

			// return the initialized instance
			il.Emit(OpCodes.Ldloc, outputLocal);
			il.Emit(OpCodes.Ret);
			
			return (Func<FastExpando, object>)dm.CreateDelegate(typeof(Func<FastExpando, object>));
		}
	}
}
