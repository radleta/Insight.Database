using System;
using System.Collections.Generic;
using System.Reflection.Emit;
using System.Reflection;

namespace InsightDatabaseDynamicInputOutputExample
{
	/// <summary>
	/// This class provides the ability to build a System.Type at runtime.
	/// </summary>
    public static class RuntimeTypeBuilder
    {
		/// <summary>
		/// Create a <see cref="TypeInfo"/> at runtime of a class with public properties of <c>props</c>.
		/// </summary>
		/// <param name="assemblyName">The name of the assembly at store the type. It's dynamically created so the numbe needs to not conflict with existing assembly names.</param>
		/// <param name="moduleName">The module within the <c>assemblyName</c> assembly.</param>
		/// <param name="typeName">The name of the <see cref="System.Type"/> to be stored in the <c>moduleName</c> module.</param>
		/// <param name="props">The properties to create on the outputted type. They will have a public get and set accessor.</param>
		/// <returns></returns>
        public static TypeInfo CompileResultTypeInfo(string assemblyName, string moduleName, string typeName, Dictionary<string, Type> props)
        {
            TypeBuilder tb = GetTypeBuilder(assemblyName, moduleName, typeName);
            ConstructorBuilder constructor = tb.DefineDefaultConstructor(MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.RTSpecialName);
			
            foreach (var field in props)
                CreateProperty(tb, field.Key, field.Value);

            TypeInfo objectTypeInfo = tb.CreateTypeInfo();
            return objectTypeInfo;
        }

        private static TypeBuilder GetTypeBuilder(string assemblyName, string moduleName, string typeName)
        {
            var an = new AssemblyName(assemblyName);
            var assemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(an, AssemblyBuilderAccess.Run);
            ModuleBuilder moduleBuilder = assemblyBuilder.DefineDynamicModule(moduleName);
            TypeBuilder tb = moduleBuilder.DefineType(typeName,
                    TypeAttributes.Public |
                    TypeAttributes.Class |
                    TypeAttributes.AutoClass |
                    TypeAttributes.AnsiClass |
                    TypeAttributes.BeforeFieldInit |
                    TypeAttributes.AutoLayout,
                    null);
            return tb;
        }

        private static void CreateProperty(TypeBuilder tb, string propertyName, Type propertyType)
        {
            FieldBuilder fieldBuilder = tb.DefineField("_" + propertyName, propertyType, FieldAttributes.Private);

            PropertyBuilder propertyBuilder = tb.DefineProperty(propertyName, PropertyAttributes.HasDefault, propertyType, null);
            MethodBuilder getPropMthdBldr = tb.DefineMethod("get_" + propertyName, MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.HideBySig, propertyType, Type.EmptyTypes);
            ILGenerator getIl = getPropMthdBldr.GetILGenerator();

            getIl.Emit(OpCodes.Ldarg_0);
            getIl.Emit(OpCodes.Ldfld, fieldBuilder);
            getIl.Emit(OpCodes.Ret);

            MethodBuilder setPropMthdBldr =
                tb.DefineMethod("set_" + propertyName,
                  MethodAttributes.Public |
                  MethodAttributes.SpecialName |
                  MethodAttributes.HideBySig,
                  null, new[] { propertyType });

            ILGenerator setIl = setPropMthdBldr.GetILGenerator();
            Label modifyProperty = setIl.DefineLabel();
            Label exitSet = setIl.DefineLabel();

            setIl.MarkLabel(modifyProperty);
            setIl.Emit(OpCodes.Ldarg_0);
            setIl.Emit(OpCodes.Ldarg_1);
            setIl.Emit(OpCodes.Stfld, fieldBuilder);

            setIl.Emit(OpCodes.Nop);
            setIl.MarkLabel(exitSet);
            setIl.Emit(OpCodes.Ret);

            propertyBuilder.SetGetMethod(getPropMthdBldr);
            propertyBuilder.SetSetMethod(setPropMthdBldr);
        }


    }
}
