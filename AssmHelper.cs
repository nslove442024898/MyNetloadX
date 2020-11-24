using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace MyNetloadX
{
    public class FileAndDependentAssemblyList
    {
        /// <summary>
        /// dll引用来源
        /// </summary>
        public StringBuilder SourceFile;
        /// <summary>
        /// dll文件路径
        /// </summary>
        public string File;
        /// <summary>
        /// dll内的程序集
        /// </summary>
        public List<string> Dependent;

        /// <summary>
        /// 储存Dll的依赖结构
        /// </summary>
        public FileAndDependentAssemblyList(string file = null, string sourceFile = null)
        {
            if (sourceFile != null)
            {
                SourceFile = new StringBuilder(sourceFile);
            }
            File = file;
            Dependent = new List<string>();
        }
    }

    public class AssemblyDependent
    {
        string _dllFile;
        /// <summary>
        /// cad程序域依赖_内存区(不可以卸载)
        /// </summary>
        private Assembly[] _cadAs;

        /// <summary>
        /// cad程序域依赖_映射区(不可以卸载)
        /// </summary>
        private Assembly[] _cadAsRef;

        /// <summary>
        /// 当前域加载事件
        /// </summary>
        public event ResolveEventHandler CurrentDomainAssemblyResolveEvent;

        /// <summary>
        /// 加载dll的和相关的依赖
        /// </summary>
        /// <param name="dllFile"></param>
        public AssemblyDependent(string dllFile)
        {
            _dllFile = dllFile;

            //cad程序集的依赖
            _cadAs = AppDomain.CurrentDomain.GetAssemblies();

            //映射区
            _cadAsRef = AppDomain.CurrentDomain.ReflectionOnlyGetAssemblies();

            //运行时出错的话,就靠这个事件来解决
            if (CurrentDomainAssemblyResolveEvent != null)
            {
                AppDomain.CurrentDomain.AssemblyResolve += CurrentDomainAssemblyResolveEvent;
            }
            else
            {
                AppDomain.CurrentDomain.AssemblyResolve += RunTimeCurrentDomain.DefaultAssemblyResolve;
            }
        }



        /// <summary>
        /// 返回的类型,描述加载的错误
        /// </summary>
        public class LoadDllMessage
        {
            public string Path;
            public bool LoadYes;

            public LoadDllMessage(string path, bool loadYes)
            {
                Path = path;
                LoadYes = loadYes;
            }
        }

        /// <summary>
        /// 字节加载
        /// </summary>
        /// <param name="_dllFile"></param>
        /// <returns></returns>
        public LoadDllMessage[] Load()
        {
            var loadYesList = new List<LoadDllMessage>();
            if (!File.Exists(_dllFile))
            {
                return loadYesList.ToArray();
            }

            //查询加载链之后再逆向加载,确保前面不丢失
            var allRefs = GetAllRefPaths(_dllFile);
            allRefs.Reverse();

            foreach (var path in allRefs)
            {
                //路径转程序集名
                string assName = AssemblyName.GetAssemblyName(path).FullName;
                //路径转程序集名
                Assembly assembly = _cadAs.FirstOrDefault((Assembly a) => a.FullName == assName);
                if (assembly == null)
                {

                    //为了实现debug时候出现断点,见链接
                    // https://www.cnblogs.com/DasonKwok/p/10510218.html
                    // https://www.cnblogs.com/DasonKwok/p/10523279.html

                    //实现字节加载 
                    var buffer = File.ReadAllBytes(path);
#if DEBUG
                    var dir = Path.GetDirectoryName(path);
                    var pdbName = Path.GetFileNameWithoutExtension(path) + ".pdb";
                    var pdbFullName = Path.Combine(dir, pdbName);
                    if (File.Exists(pdbFullName))
                    {
                        var pdbbuffer = File.ReadAllBytes(pdbFullName);
                        Assembly.Load(buffer, pdbbuffer);//就是这句会占用vs生成,可能这个问题是net strandard
                    }
                    else
                    {
                        Assembly.Load(buffer);
                    }
#else
                    Assembly.Load(buffer);
#endif   
                    loadYesList.Add(new LoadDllMessage(path, true));//加载成功
                }
                else
                {
                    loadYesList.Add(new LoadDllMessage(path, false));//版本号没变不加载
                }
            }
            return loadYesList.ToArray();
        }


        /// <summary>
        /// 获取加载链
        /// </summary>
        /// <param name="dll"></param>
        /// <param name="dlls"></param>
        /// <returns></returns>
        List<string> GetAllRefPaths(string dll, List<string> dlls = null)
        {
            dlls = dlls ?? new List<string>();
            //如果含有 || 不存在文件
            if (dlls.Contains(dll) || !File.Exists(dll))
            {
                return dlls;
            }
            dlls.Add(dll);

            //路径转程序集名
            string assName = AssemblyName.GetAssemblyName(dll).FullName;

            //在当前程序域的assemblyAs内存区和assemblyAsRef映射区找这个程序集名
            Assembly assemblyAs = _cadAs.FirstOrDefault((Assembly a) => a.FullName == assName);
            Assembly assemblyAsRef;

            //内存区有表示加载过
            //映射区有表示查找过但没有加载(一般来说不存在.只是debug会注释掉Assembly.Load的时候用来测试)
            if (assemblyAs != null)
            {
                assemblyAsRef = assemblyAs;
            }
            else
            {
                assemblyAsRef = _cadAsRef.FirstOrDefault((Assembly a) => a.FullName == assName);

                //内存区和映射区都没有的话就把dll加载到映射区,用来找依赖表
                assemblyAsRef = assemblyAsRef ?? Assembly.ReflectionOnlyLoad(File.ReadAllBytes(dll));
            }

            //遍历依赖,如果存在dll拖拉加载目录就加入dlls集合
            foreach (var assemblyName in assemblyAsRef.GetReferencedAssemblies())
            {
                //dll拖拉加载路径-搜索路径(可以增加到这个dll下面的所有文件夹?)
                string directoryName = Path.GetDirectoryName(dll);

                var path = directoryName + "\\" + assemblyName.Name;
                var paths = new string[]
                {
                    path + ".dll",
                    path + ".exe"
                };
                foreach (var patha in paths)
                {
                    GetAllRefPaths(patha, dlls);
                }
            }
            return dlls;
        }
    }

    public class MyHelper
    {
        public static Version GetFileVersionCe(string fileName)
        {
            int handle = 0;
            int length = GetFileVersionInfoSize(fileName, ref handle);
            Version v = null;
            if (length > 0)
            {
                IntPtr buffer = System.Runtime.InteropServices.Marshal.AllocHGlobal(length);
                if (GetFileVersionInfo(fileName, handle, length, buffer))
                {
                    IntPtr fixedbuffer = IntPtr.Zero;
                    int fixedlen = 0;
                    if (VerQueryValue(buffer, "\\\\", ref fixedbuffer, ref fixedlen))
                    {
                        byte[] fixedversioninfo = new byte[fixedlen];
                        System.Runtime.InteropServices.Marshal.Copy(fixedbuffer, fixedversioninfo, 0, fixedlen);
                        v = new Version(
                            BitConverter.ToInt16(fixedversioninfo, 10),
                            BitConverter.ToInt16(fixedversioninfo, 8),
                            BitConverter.ToInt16(fixedversioninfo, 14),
                            BitConverter.ToInt16(fixedversioninfo, 12));
                    }
                }
                Marshal.FreeHGlobal(buffer);
            }
            return v;
        }

        [DllImport("version.dll", EntryPoint = "GetFileVersionInfo", SetLastError = true)]
        private static extern bool GetFileVersionInfo(string filename, int handle, int len, IntPtr buffer);
        [DllImport("version.dll", EntryPoint = "GetFileVersionInfoSize", SetLastError = true)]
        private static extern int GetFileVersionInfoSize(string filename, ref int handle);
        [DllImport("version.dll", EntryPoint = "VerQueryValue", SetLastError = true)]
        private static extern bool VerQueryValue(IntPtr buffer, string subblock, ref IntPtr blockbuffer, ref int len);

    }
    public static class RunTimeCurrentDomain
    {
        #region  程序域运行事件 
        // 动态编译要注意所有的引用外的dll的加载顺序
        // cad2008若没有这个事件,会使动态命令执行时候无法引用当前的程序集函数
        // 跨程序集反射
        // 动态加载时,dll的地址会在系统的动态目录里,而它所处的程序集(运行域)是在动态目录里.
        // netload会把所处的运行域给改到cad自己的,而动态编译不通过netload,所以要自己去改.
        // 这相当于是dll注入的意思,只是动态编译的这个"dll"不存在实体,只是一段内存.

        /// <summary>
        /// 程序域运行事件
        /// </summary>   
        public static Assembly DefaultAssemblyResolve(object sender, ResolveEventArgs args)
        {
            var cad = AppDomain.CurrentDomain.GetAssemblies();

#if false
            /*获取名称一致,但是版本号不同的,调用最开始的版本*/
            //获取执行程序集的参数
             var ag = args.Name.Split(',')[0];
            //获取 匹配符合条件的第一个或者默认的那个
            // var load = cad.FirstOrDefault(a => a.GetName().FullName.Split(',')[0] == ag); 
#endif

            /*获取名称和版本号都一致的,调用它*/
            Assembly load = null;
            load = cad.FirstOrDefault(a => a.GetName().FullName == args.Name);
            if (load == null)
            {
                /*获取名称一致,但是版本号不同的,调用最后的可用版本*/
                var ag = args.Name.Split(',')[0];
                //获取 最后一个符合条件的,
                //否则a.dll引用b.dll函数的时候,b.dll修改重生成之后,加载进去会调用第一个版本的b.dll            
                foreach (var item in cad)
                {
                    if (item.GetName().FullName.Split(',')[0] == ag)
                    {
                        //为什么加载的程序版本号最后要是*
                        //因为vs会帮你迭代这个版本号,所以最后的可用就是循环到最后的.
                        load = item;
                    }
                }
            }

            return load;
        }
        #endregion
    }
}