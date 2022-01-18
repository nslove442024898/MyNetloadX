using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MyNetloadX
{
    /// <summary>
    /// cad的dll运行时拓展
    /// </summary>
    public static class AcRuntimeEx
    {
        /// <summary>
        /// 获取dll含有的特性的全部方法
        /// </summary>
        /// <param name="cadDllAss"></param>
        /// <returns>通过反射提取带特性的public的方法</returns>
        public static List<MethodInfo> GetWithAttributeMethods(this Assembly cadDllAss)
        {
            //提起全部的公共类
            var cls = cadDllAss.GetTypes().Where(c => c.IsClass && c.IsPublic).ToList();
            //提取全部含有特性的方法
            var methods = new List<MethodInfo>();
            foreach (var item in cls)
            {
                var mis = item.GetMethods().Where(c => c.IsPublic && c.GetCustomAttributes(true).Length > 0).ToList();
                methods.AddRange(mis);
            }
            return methods;
        }
        /// <summary>
        /// cad注册命令的的信息
        /// </summary>
        /// <param name="mi">MethodInfo</param>
        /// <returns>cad注册命令的的信息</returns>
        public static AcadCustomCmdinfor GetAcadCmdInfor(this MethodInfo mi)=> mi.IsCadCustomCmdMethod() ? new AcadCustomCmdinfor(mi) : null;

        /// <summary>
        /// 判断是否是含有cad的特性的方法
        /// </summary>
        /// <param name="mi"></param>
        /// <returns></returns>
        public static bool IsCadCustomCmdMethod(this MethodInfo mi)=> mi.GetCustomAttribute(typeof(CommandMethodAttribute)) != null;
    }
    /// <summary>
    /// cad 命令的类
    /// </summary>
    public class AcadCustomCmdinfor
    {
        public string CmdName { get; set; }
        public string MethodName { get; set; }
        public string ClassName { get; set; }
        public string NameSpaceName { get; set; }
        public string AssemblyName { get; set; }
        public AcadCustomCmdinfor(MethodInfo mi)
        {
            var att = (CommandMethodAttribute) mi.GetCustomAttribute(typeof(CommandMethodAttribute));
            if (att != null) this.CmdName = att.GlobalName;
            this.MethodName = mi.Name;
            this.ClassName = mi.DeclaringType?.Name;
            this.NameSpaceName = mi.DeclaringType?.Namespace;
            this.AssemblyName = mi.DeclaringType?.Assembly.FullName;
        }

        public override string ToString()
        {
            var sb = new StringBuilder();

            var props = this.GetType().GetProperties().Where(p => p.PropertyType.Namespace == "System").ToList();

            foreach (var t in props)
            {
                var porpVal = t.GetValue(this);
                sb.Append(t.Name+"===>"+porpVal.ToString()+"\r\t"+Environment.NewLine);
            }
            return sb.ToString();
        }
    }
}
