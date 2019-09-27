using System;
using System.Linq;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;

namespace InsightDatabaseDynamicInputOutputExample
{
	/// <summary>
	/// Creates a wrapper for <see cref="IDictionary{String, Object}"/> that exposes the keys as true properties.
	/// </summary>
	/// <remarks>
	/// <para>The basic problem solved by this class is creating a runtime <see cref="Type"/> that has get accessors for 
	/// each property dynamically.</para>
	/// <para>The underlying implementation creates dynamic Types at runtime based on a <see cref="Dictionary{string, Type}"/> definition.</para>
	/// </remarks>
	public static class RuntimeDictionaryWrapper
	{
		private static readonly Type DictionaryType = typeof(IDictionary<string, object>);
		private static readonly Type[] OneDictionaryType = new Type[] { DictionaryType };
		private static readonly MethodInfo DictionaryGetItemMethodInfo = DictionaryType.GetMethod("get_Item", BindingFlags.Public | BindingFlags.Instance);
		private static readonly MethodInfo DictionarySetItemMethodInfo = DictionaryType.GetMethod("set_Item", BindingFlags.Public | BindingFlags.Instance);
		private static readonly ConstructorInfo DefaultObjectConstructorInfo = typeof(object).GetConstructor(new Type[0]);
		private static readonly ConstructorInfo ArgumentNullExceptionStringConstructorInfo = typeof(System.ArgumentNullException).GetConstructor(new Type[] { typeof(string) });

		/// <summary>
		/// The cached types created by their name. We use ConcurrentDictionary to prevent threaded issues of creating more than one type with the same name.
		/// </summary>
		private static readonly ConcurrentDictionary<string, Type> WrapperTypeByName = new ConcurrentDictionary<string, Type>();

		/// <summary>
		/// The cached converters created by their type.
		/// </summary>
		private static readonly ConcurrentDictionary<Type, Func<IDictionary<string, object>, object>> ConverterByType = new ConcurrentDictionary<Type, Func<IDictionary<string, object>, object>>();

		/// <summary>
		/// Thread Save. This is the thread safe singleton of the <see cref="ModuleBuilder"/>. It prevents more than one from being created.
		/// </summary>
		private readonly static Lazy<ModuleBuilder> ModuleBuilder = new Lazy<ModuleBuilder>(CreateModuleBuilder);

		/// <summary>
		/// Creates a <see cref="ModuleBuilder"/> to be used to store all the dynamic types created by this class.
		/// </summary>
		/// <returns>The <see cref="ModuleBuilder"/>.</returns>
		private static ModuleBuilder CreateModuleBuilder()
		{
			var an = new AssemblyName($"DictionaryConvert-{Guid.NewGuid()}");
			var assemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(an, AssemblyBuilderAccess.Run);
			return assemblyBuilder.DefineDynamicModule("DictionaryConvertModule");
		}
		
		/// <summary>
		/// Creats an MD5 from the <c>s</c>.
		/// </summary>
		/// <param name="s"></param>
		/// <returns></returns>
		private static string ConvertToMD5(string s)
		{
			if (s is null)
			{
				throw new ArgumentNullException(nameof(s));
			}

			using (var provider = System.Security.Cryptography.MD5.Create())
			{
				StringBuilder builder = new StringBuilder();

				foreach (byte b in provider.ComputeHash(Encoding.UTF8.GetBytes(s)))
					builder.Append(b.ToString("x2").ToLower());

				return builder.ToString();
			}
		}

		/// <summary>
		/// Thread Safe. Gets the <see cref="Type"/> which wraps a <see cref="IDictionary{String, Object}"/> where the new type 
		/// exposes the keys and values as .NET properties. Any value types will be exposed as <see cref="Nullable{T}"/>. The type 
		/// has one constructor which takes in the <see cref="IDictionary{String, Object}"/> instance.
		/// </summary>
		/// <param name="properties">The properties to expose.</param>
		/// <returns>The <see cref="Type"/> that exposes the keys/values of a <see cref="IDictionary{String, Object}"/> instance as <c>properties</c>.</returns>
		public static Type GetOrCreateWrapperType(Dictionary<string, Type> properties)
		{
			if (properties is null)
			{
				throw new ArgumentNullException(nameof(properties));
			}

			// create a unique name based on the inputs
			var uniqueName = string.Concat(properties.Select(kv => $"{kv.Key}_{kv.Value.AssemblyQualifiedName}").OrderBy(s => s));

			// convert it to md5 to create a unqiue valid type name
			var validTypeName = ConvertToMD5(uniqueName);

			// get or add the wrapper type
			return WrapperTypeByName.GetOrAdd(validTypeName, (s) => CreateWrapperType(properties, s));
		}

		/// <summary>
		/// Creates the actual type. This should be called only once for each <c>typeName</c> as they are unique within the <see cref="ModuleBuilder"/>.
		/// </summary>
		/// <param name="properties">The properties</param>
		/// <param name="typeName">The name of the type to create.</param>
		/// <returns>The created type.</returns>
		private static Type CreateWrapperType(Dictionary<string, Type> properties, string typeName)
		{
			if (typeName is null)
			{
				throw new ArgumentNullException(nameof(typeName));
			}

			if (properties is null)
			{
				throw new ArgumentNullException(nameof(properties));
			}

			// create the type builder
			var typeBuilder = ModuleBuilder.Value.DefineType(typeName,
					TypeAttributes.Public |
					TypeAttributes.Class |
					TypeAttributes.AutoClass |
					TypeAttributes.AnsiClass |
					TypeAttributes.BeforeFieldInit |
					TypeAttributes.AutoLayout,
					null);

			// create the private _wrapper field
			FieldBuilder dictionaryField = typeBuilder.DefineField("_dictionary", typeof(System.Collections.Generic.Dictionary<string, object>), FieldAttributes.Private | FieldAttributes.InitOnly);

			// create the constructor which accepts one param of System.Collections.Generic.Dictionary<string, object>
			DefineDictionaryWrapperConstructor(typeBuilder, dictionaryField);

			// define all the wrapper properties
			foreach (var field in properties)
				DefineDictionaryWrapperProperty(typeBuilder, field.Key, field.Value, dictionaryField);

			// create the type info based on what we've defined so far
			TypeInfo objectTypeInfo = typeBuilder.CreateTypeInfo();
			return objectTypeInfo;
		}

		/// <summary>
		/// Define a public constructor that takes a <see cref="IDictionary{String, Object}"/> as its parameter.
		/// </summary>
		/// <param name="typeBuilder">The type builder.</param>
		/// <param name="dictionaryField">The field that stores the wrapped <see cref="IDictionary{String, Object}"/>.</param>
		private static void DefineDictionaryWrapperConstructor(TypeBuilder typeBuilder, FieldBuilder dictionaryField)
		{
			// define the name of the parameter for the constructor
			const string ParamName = "dictionary";

			if (typeBuilder is null)
			{
				throw new ArgumentNullException(nameof(typeBuilder));
			}

			if (dictionaryField is null)
			{
				throw new ArgumentNullException(nameof(dictionaryField));
			}

			// define the ctor
			var ctor = typeBuilder.DefineConstructor(MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.RTSpecialName | MethodAttributes.HideBySig, CallingConventions.Any, OneDictionaryType);

			// define the incoming param
			var param1 = ctor.DefineParameter(0, ParameterAttributes.In, ParamName);

			// start generating il
			var il = ctor.GetILGenerator();

			// create the obj by calling the default ctor on object
			il.Emit(OpCodes.Ldarg_0);
			il.Emit(OpCodes.Call, DefaultObjectConstructorInfo);

			// null check on param
			il.Emit(OpCodes.Ldarg_1);
			var endIfLabel = il.DefineLabel();
			il.Emit(OpCodes.Brtrue_S, endIfLabel);
			il.Emit(OpCodes.Ldstr, ParamName);
			il.Emit(OpCodes.Newobj, ArgumentNullExceptionStringConstructorInfo);
			il.Emit(OpCodes.Throw);
			il.MarkLabel(endIfLabel);

			// store the param into our private field
			il.Emit(OpCodes.Ldarg_0);
			il.Emit(OpCodes.Ldarg_1);
			il.Emit(OpCodes.Stfld, dictionaryField);
			il.Emit(OpCodes.Ret);
		}

		/// <summary>
		/// Define a public property that exposes the value of the <see cref="IDictionary{String, Object}"/> with the same <c>propertyName</c>.
		/// </summary>
		/// <param name="typeBuilder">The type builder.</param>
		/// <param name="propertyName">The name of the property and also the same name of the key in the dictionary to expose.</param>
		/// <param name="propertyType">The type of the property to be exposed. If it's a value type, it'll be exposed as a Nullable type.</param>
		/// <param name="dictionaryField">The field that stores the wrapped <see cref="IDictionary{String, Object}"/>.</param>
		private static void DefineDictionaryWrapperProperty(TypeBuilder typeBuilder, string propertyName, Type propertyType, FieldBuilder dictionaryField)
		{
			if (typeBuilder is null)
			{
				throw new ArgumentNullException(nameof(typeBuilder));
			}

			if (propertyName is null)
			{
				throw new ArgumentNullException(nameof(propertyName));
			}

			if (propertyType is null)
			{
				throw new ArgumentNullException(nameof(propertyType));
			}

			if (dictionaryField is null)
			{
				throw new ArgumentNullException(nameof(dictionaryField));
			}
					   
			// convert all incoming value types to their nullable version
			if (propertyType.IsValueType)
			{
				propertyType = typeof(Nullable<>).MakeGenericType(propertyType);
			}
			
			PropertyBuilder propertyBuilder = typeBuilder.DefineProperty(propertyName, PropertyAttributes.HasDefault, propertyType, null);

			// define the get accessor
			MethodBuilder getPropMthdBldr = typeBuilder.DefineMethod("get_" + propertyName, MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.HideBySig, propertyType, Type.EmptyTypes);
			ILGenerator getIl = getPropMthdBldr.GetILGenerator();
			
			getIl.Emit(OpCodes.Ldarg_0);
			getIl.Emit(OpCodes.Ldfld, dictionaryField);
			getIl.Emit(OpCodes.Ldstr, propertyName);
			getIl.Emit(OpCodes.Callvirt, DictionaryGetItemMethodInfo);
			if (propertyType.IsValueType)
			{				
				getIl.Emit(OpCodes.Unbox_Any, propertyType);
			}
			else
			{
				getIl.Emit(OpCodes.Castclass, propertyType);
			}			
			getIl.Emit(OpCodes.Ret);

			propertyBuilder.SetGetMethod(getPropMthdBldr);

			MethodBuilder setPropMthdBldr =
				typeBuilder.DefineMethod("set_" + propertyName,
				  MethodAttributes.Public |
				  MethodAttributes.SpecialName |
				  MethodAttributes.HideBySig,
				  null, new[] { propertyType });

			ILGenerator setIl = setPropMthdBldr.GetILGenerator();
			setIl.Emit(OpCodes.Ldarg_0);
			setIl.Emit(OpCodes.Ldfld, dictionaryField);
			setIl.Emit(OpCodes.Ldstr, propertyName);
			setIl.Emit(OpCodes.Ldarg_1);

			if (propertyType.IsValueType)
			{
				setIl.Emit(OpCodes.Box, propertyType);
			}

			setIl.Emit(OpCodes.Callvirt, DictionarySetItemMethodInfo);
			setIl.Emit(OpCodes.Ret);

			propertyBuilder.SetSetMethod(setPropMthdBldr);
		}

		/// <summary>
		/// Uses IL to create an instance of <c>type</c> using the default constructor
		/// then assign any public properties of the <c>type</c> from <see cref="FastExpando"/>
		/// to the new instance.
		/// </summary>
		/// <param name="wrapperType">The type of object to be able to convert.</param>
		/// <returns>A function that can convert that type of <see cref="FastExpando"/> to a <c>type</c>.</returns>
		public static Func<IDictionary<string, object>, object> GetOrCreateConverter(Type wrapperType) => ConverterByType.GetOrAdd(wrapperType, CreateConverter);

		/// <summary>
		/// Uses IL to create an instance of <c>type</c> using the default constructor
		/// then assign any public properties of the <c>type</c> from <see cref="FastExpando"/>
		/// to the new instance.
		/// </summary>
		/// <param name="wrapperType">The type of object to be able to convert.</param>
		/// <returns>A function that can convert that type of <see cref="FastExpando"/> to a <c>type</c>.</returns>
		private static Func<IDictionary<string, object>, object> CreateConverter(Type wrapperType)
		{
			if (wrapperType is null)
			{
				throw new ArgumentNullException(nameof(wrapperType));
			}

			// create a dynamic method
			var dm = new DynamicMethod($"DictionaryConvert-{wrapperType.FullName}", typeof(object), OneDictionaryType, typeof(RuntimeDictionaryWrapper), true);

			// build the il
			var il = dm.GetILGenerator();

			// create a new instance of the wrapper type passing in the argument 0 as the param 0
			il.Emit(OpCodes.Ldarg_0);
			il.Emit(OpCodes.Newobj, wrapperType.GetConstructor(OneDictionaryType));
			il.Emit(OpCodes.Ret);

			return (Func<IDictionary<string, object>, object>)dm.CreateDelegate(typeof(Func<IDictionary<string, object>, object>));
		}
	}
}
