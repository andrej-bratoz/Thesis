using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ILDisassembler;

namespace MetaInspector.Logic
{
    public class AssemblyLoader
    {
        public List<NamespaceDTO> Namespaces { get; set; } 
        private readonly string path;
        public AssemblyLoader(string path)
        {
            this.path = path;
        }
        
        public List<ClassInfoDTO> GetAssemblyInfo()
        {
            var list = new List<ClassInfoDTO>();
            var assembly = Assembly.LoadFrom(path);
            assembly.GetTypes().ToList().ForEach(type =>
            {
                if (!type.Name.StartsWith("<"))
                {
                    var classInfo = new ClassInfoDTO();
                    if (!type.IsGenericType)
                    {
                        classInfo.ClassName = String.Format("{0}" , type.Name);
                        classInfo.Namespace = type.Namespace;
                    }
                    else
                    {
                        classInfo.ClassName = String.Format("{0}<{1}>", type.Name.Split('`').First(),
                            GetGenericArguments(type));
                        classInfo.Namespace = type.Namespace;
                    }
                    classInfo.Fields = GetFieldInfo(type);
                    classInfo.Properties = GetPropertyInfo(type);
                    classInfo.Methods = GetMethodInfo(type);
                    classInfo.Events = GetEventInfo(type);
                    list.Add(classInfo);
                }
            });
            list = list.OrderBy(x => x.Namespace).ToList();
            var namespaces = new List<NamespaceDTO>();
            list.Select(x => x.Namespace).Distinct().ToList().ForEach(x =>
            {
                var ns = new NamespaceDTO {ClassInfo = new List<ClassInfoDTO>()};
                list.Where(y => y.Namespace == x).ToList().ForEach(z => ns.ClassInfo.Add(z));
                namespaces.Add(ns);
            });
            Namespaces = namespaces;
            return list;
        }

        public List<string> GetFieldInfo(Type classType)
        {
            var list = new List<string>();
            var fieldInfos = classType.GetFields();
            foreach (var info in fieldInfos)
            {
                if (!info.FieldType.IsGenericType)
                {
                    var fieldInfo = string.Format("{0} (Type : {1})",
                        info.Name, info.FieldType.Name);
                    list.Add(fieldInfo);
                }
                else
                {
                    var fieldInfo = string.Format("{0} (Type : {1}<{2}>)",
                        info.Name, info.FieldType.GetGenericTypeDefinition().Name,
                        GetGenericArguments(info.FieldType));
                    list.Add(fieldInfo);
                }
            }
            return list;
        }

        public List<string> GetPropertyInfo(Type classType)
        {
            var list = new List<string>();
            var allProps = classType.GetProperties();
            foreach(var propInfo in allProps)
            {
                if (!propInfo.PropertyType.IsGenericType)
                {
                    var fieldInfo = string.Format("{1} {0} {{ {2} {3} }}",
                        propInfo.Name.Split('`').First(), propInfo.PropertyType.Name, 
                        propInfo.CanRead ? "get;" : "",
                        propInfo.CanWrite ? "set;" : "");
                    list.Add(fieldInfo);
                }
                else
                {
                    var fieldInfo = string.Format("{1}<{2}> {0} {{ {3} {4} }}",
                        propInfo.Name, propInfo.PropertyType.GetGenericTypeDefinition().Name.Split('`').First(),
                        GetGenericArguments(propInfo.PropertyType), propInfo.CanRead ? "get;" : "",
                        propInfo.CanWrite ? "set;" : "");
                    list.Add(fieldInfo);
                }
            }
            return list;
        }

        public List<string> GetEventInfo(Type classType)
        {
            var list = new List<string>();
            var eventInfos = classType.GetEvents();
            foreach (var info in eventInfos)
            {
                if (info.EventHandlerType.IsGenericType)
                {
                    var fieldInfo = string.Format("{0} event {1}<{2}> {3}",
                   "public", info.EventHandlerType.Name.Split('`').First(), GetGenericArguments(info.EventHandlerType), info.Name);
                    list.Add(fieldInfo);
                }
                else
                {
                    var fieldInfo = string.Format("{0} event {1} {2}",
                   "public", info.EventHandlerType.Name, info.Name);
                    list.Add(fieldInfo);
                }
               
            }
            return list;
        }

        public List<MethodInfoDto> GetMethodInfo(System.Type classType)
        {
            var methodInfos = classType.GetMethods();
            var list = new List<MethodInfoDto>();
            methodInfos.ToList().ForEach(info =>
            {
                var item = new MethodInfoDto();
                item.MethodName =
                    string.Format("{0} {1} ({2})",
                        GetMethodAccessModifiers(info),
                        info.Name, GetMethodParameters(info));
                item.MethodBody = GetMethodILBody(info);
                list.Add(item);
            });
            return list;
        }
        //return retvals;
        
        public string GetGenericArguments(Type t)
        {
            var list = new List<string>();
            t.GetGenericArguments().ToList().ForEach(x => list.Add(x.Name));
            return string.Join(",", list.ToArray());
        }

        public string GetMethodAccessModifiers(MethodInfo info)
        {
            var mod = "";
            if (info.IsPrivate) mod += "private";
            if (info.IsPublic) mod += "public";
            if (info.IsStatic) mod += " static";
            return mod;
        }
       
        public string GetMethodParameters(MethodInfo info)
        {
            var parameters = info.GetParameters();
            var list = new List<string>();
            parameters.ToList().ForEach(x =>
            {
                if (!x.ParameterType.IsGenericType)
                {
                    list.Add(String.Format("{0} {1}", x.ParameterType.Name, x.Name));
                }
                else
                {
                    list.Add(String.Format("{0} {1}", GetGenericArguments(x.ParameterType), x.Name));
                }
            });
            return string.Join(", ", list.ToArray());
        }

        public string GetMethodILBody(MethodInfo info)
        {
            Disassembler dis = new Disassembler();
            return dis.DisassembleMethod(info);

        }

    }
}
