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

    public class DependentAssembly
    {
        static string _location;
        static string _loadPath;
        static Assembly[] _cadAppAssemblie;
        static string _dllFile;

        /// <summary>
        /// 加载dll的和相关的依赖
        /// </summary>
        /// <param name="dllFile"></param>
        public DependentAssembly(string dllFile)
        {
            _dllFile = dllFile;
            //cad程序集的依赖
            _cadAppAssemblie = AppDomain.CurrentDomain.GetAssemblies();

            //这里是有关CLR的Assembly的搜索路径的过程,
            //插件目录 "G:\\K01.惊惊连盒\\net35"
            _location = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            //拖拉文件的路径 "D:\\桌面\\若海的\\test2\\bin\\Debug"
            _loadPath = Path.GetDirectoryName(_dllFile);

            //运行时靠事件来解决问题,见博客--记得反注释
            //AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
        }

        /// <summary>
        /// 依赖的dll列表
        /// </summary>
        public FileAndDependentAssemblyList[] AssemblyList()
        {
            BuildDependentAssemblyList(_dllFile, null);
            return _fadaList.ToArray();//数组0后面跟的就是数组0的引用
        }

        //dll引用表,数组[1]就是[0]的引用
        static readonly List<FileAndDependentAssemblyList> _fadaList = new List<FileAndDependentAssemblyList>();
        // 储存dll的引用
        string _sourceFile;
        // 每次递归的临时对象
        static AssemblyName _assemblyName;

        /// <summary>
        /// 链式查找依赖的dll
        /// </summary>
        /// <param name="pa">dll文件路径</param>
        /// <param name="dlls">首次为null</param>
        /// <returns>返回在_fadaList集合上</returns>
        void BuildDependentAssemblyList(string pa, FileAndDependentAssemblyList dlls)
        {
            if (dlls == null)
            {
                dlls = new FileAndDependentAssemblyList();
            }

            // 是否已经包含这个路径的程序了
            if (dlls.Dependent.Contains(pa))
                return;

            bool reflectionOnlyLoad = true;
            bool addlies = true;
            Assembly asm = null;

            // 路径 || 程序名
            if ((pa.IndexOf(Path.DirectorySeparatorChar, 0, pa.Length) != -1) ||
                (pa.IndexOf(Path.AltDirectorySeparatorChar, 0, pa.Length) != -1))
            {
                dlls = new FileAndDependentAssemblyList(pa, _sourceFile);
                _sourceFile = pa;
                // 从这个路径加载程序集,路径
                asm = Assembly.ReflectionOnlyLoadFrom(pa);
            }
            else
            {
                //判断cad程序集中是否存在,dll的依赖,不存在则搜索拖拉路径的dll,
                //如果已经存在cad程序集上面就直接加载就可以了 
                //(cad程序集的依赖)和(加载的dll程序集的依赖)不一样,就加载,并加入集合,==0直接加载
                if (_cadAppAssemblie.Count(c => c.FullName == pa) == 0)
                {
                    var asdl = _assemblyName.Name + ".dll";
                    string path1 = Path.Combine(_location, asdl);
                    string path2 = Path.Combine(_loadPath, asdl);
                    //如果插件目录和拖拉目录都有,那么拖拉目录会覆盖前面
                    reflectionOnlyLoad = LoadFrom(path1, ref dlls, ref asm, ref addlies);
                    reflectionOnlyLoad = LoadFrom(path2, ref dlls, ref asm, ref addlies);
                }
                if (reflectionOnlyLoad)
                {
                    addlies = false;
                    dlls.Dependent.Add(pa);
                    // 反射加载上下文,但不能执行程序集
                    asm = Assembly.ReflectionOnlyLoad(pa);
                }
            }

            if (asm == null)
            {
                return;
            }

            //如果包含的话,表示要插入到这个dll下的 
            if (addlies)
            {
                _fadaList.Add(dlls);
            }

            // 遍历依赖(所有的引用)，并进行递归
            var asms = asm.GetReferencedAssemblies();
            foreach (AssemblyName item in asms)
            {
                _assemblyName = item;
                BuildDependentAssemblyList(item.FullName, dlls);
            }
        }


        /// <summary>
        /// 反射加载上下文,执行程序集
        /// </summary> 
        bool LoadFrom(string path,
            ref FileAndDependentAssemblyList dlls,
            ref Assembly asm,
            ref bool addlies)
        {
            if (File.Exists(path))
            {
                dlls = new FileAndDependentAssemblyList(path, _sourceFile);
                _sourceFile = path;
                // 反射加载上下文,执行程序集
                asm = Assembly.ReflectionOnlyLoadFrom(path);
                addlies = true;
                return false;
            }
            return true;
        }

        /// <summary>
        /// 加载依赖的dll
        /// </summary>
        public void Load()
        {
            LoadDependentAssemblyList(_dllFile);
        }

        /// <summary>
        /// 加载依赖的dll
        /// </summary>
        /// <param name="path"></param>
        /// <param name="assemblies"></param>
        /// <returns></returns>
        void LoadDependentAssemblyList(string path)
        {
            Assembly asm = null;
            // 路径 || 程序名
            if ((path.IndexOf(Path.DirectorySeparatorChar, 0, path.Length) != -1) ||
                (path.IndexOf(Path.AltDirectorySeparatorChar, 0, path.Length) != -1))
            {
                // 从这个路径加载程序集,路径 
                asm = Assembly.Load(File.ReadAllBytes(path));
            }
            else if (_cadAppAssemblie.Count(c => c.FullName == path) == 0)
            {
                //判断cad程序集中是否存在,dll的依赖,不存在则搜索拖拉路径的dll,
                //如果已经存在cad程序集上面就直接加载就可以了
                //(cad程序集的依赖)和(加载的dll程序集的依赖)不一样,就加载,并加入集合,==0直接加载
                var asdl = _assemblyName.Name + ".dll";
                string path1 = Path.Combine(_location, asdl);
                string path2 = Path.Combine(_loadPath, asdl);
                //如果插件目录和拖拉目录都有,那么拖拉目录会覆盖前面
                if (File.Exists(path2))
                {
                    // 从这个路径加载程序集,路径
                    asm = Assembly.Load(File.ReadAllBytes(path2));
                }
                else if (File.Exists(path1))
                {
                    // 从这个路径加载程序集,路径 
                    asm = Assembly.Load(File.ReadAllBytes(path1));
                }
            }

            if (asm == null)
            {
                return;
            }

            // 遍历依赖(所有的引用)，并进行递归
            var asms = asm.GetReferencedAssemblies();
            foreach (AssemblyName item in asms)
            {
                _assemblyName = item;
                LoadDependentAssemblyList(item.FullName);
            }
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
}