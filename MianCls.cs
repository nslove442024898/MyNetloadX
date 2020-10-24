
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace MyNetloadX
{
    public class AutoCADNetLoader : IExtensionApplication
    {
        public string currentDllPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
        public void Initialize()//初始化程序。
        {
            helper.Register2HKCR();
            string path = Assembly.GetExecutingAssembly().Location;
            helper.AddCmdtoMenuBar(helper.GetDllCmds(path));
            Application.SetSystemVariable("SECURELOAD", 0);
        }

        public void Terminate()
        {
            throw new NotImplementedException();
        }
    }

    public class Tools
    {
        static List<string[]> listDllInfo = new List<string[]>();

        [CommandMethod("MyNetLoader")]
        public void Dll加载到内存并卸载()
        {
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
            Autodesk.AutoCAD.Windows.OpenFileDialog ofd = new Autodesk.AutoCAD.Windows.OpenFileDialog("Dll加载到内存并卸载", "Dll加载到内存并卸载", "dll", "Dll加载到内存并卸载", Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.AllowMultiple);
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (var item in ofd.GetFilenames())
                {
                    var _cadAppAssemblie = AppDomain.CurrentDomain.GetAssemblies();
                    var verCurLoad = MyHelper.GetFileVersionCe(item);
                    var alreadyLoad = listDllInfo.Where(c => c[0] == Path.GetFileNameWithoutExtension(item)).ToList();
                    DependentAssembly dependentAssembly = new DependentAssembly(item);
                    if (alreadyLoad.Count == 0)
                    {
                        dependentAssembly.Load();
                        Application.ShowAlertDialog($"首次加载{Path.GetFileNameWithoutExtension(item)} ====> File Version {verCurLoad.ToString()} 成功！！");
                        listDllInfo.Add(new string[] { Path.GetFileNameWithoutExtension(item), verCurLoad.ToString() });
                    }
                    else
                    {
                        StringBuilder sb = new StringBuilder();
                        bool flag = false;
                        for (int i = 0; i < alreadyLoad.Count; i++)
                        {
                            if (alreadyLoad[i][1] == verCurLoad.ToString())
                            {
                                Application.ShowAlertDialog($"{alreadyLoad[i][0]},{alreadyLoad[i][1]} 已经在程序中，如需覆盖加载请修改dll的 Assembly FileVersion！！");
                                flag = true;
                            }
                            else sb.AppendLine(alreadyLoad[i][1]);
                        }
                        if (flag == false)
                        {
                            dependentAssembly.Load();
                            Application.ShowAlertDialog($"覆盖加载 {Path.GetFileNameWithoutExtension(item)} 成功！！加载的版本为{verCurLoad.ToString()},以下版本不再可用了\n" + sb.ToString());
                            listDllInfo.Add(new string[] { Path.GetFileNameWithoutExtension(item), verCurLoad.ToString() });
                        }
                    }

                }
            }
            else Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("请选择一个或者多个dll文件！");
        }
        private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            return (sender as AppDomain).GetAssemblies().FirstOrDefault(a => a.GetName().FullName.Split(',')[0] == args.Name.Split(',')[0]);
        }

    }
    public static class helper


    {/// <summary>
     ///提取所有的命令
     /// </summary>
     /// <param name="dllFiles">dll的路径</param>
     /// <returns></returns>
        public static List<gcadDllcmd> GetDllCmds(params string[] dllFiles)
        {
            List<gcadDllcmd> res = new List<gcadDllcmd>();
            List<gcadCmds> cmds = new List<gcadCmds>();
            #region 提取所以的命令
            for (int i = 0; i < dllFiles.Length; i++)
            {
                Assembly ass = Assembly.LoadFile(dllFiles[i]);//反射加载dll程序集
                var clsCollection = ass.GetTypes().Where(t => t.IsClass && t.IsPublic).ToList();
                if (clsCollection.Count > 0)
                {
                    foreach (var cls in clsCollection)
                    {
                        var methods = cls.GetMethods().Where(m => m.IsPublic && m.GetCustomAttributes(true).Length > 0).ToList();
                        if (methods.Count > 0)
                        {
                            foreach (MethodInfo mi in methods)
                            {
                                var atts = mi.GetCustomAttributes(true).Where(c => c is CommandMethodAttribute).ToList();
                                if (atts.Count == 1)
                                {
                                    gcadCmds cmd = new gcadCmds(cls.Name, mi.Name, (atts[0] as CommandMethodAttribute).GlobalName, ass.ManifestModule.Name.Substring(0, ass.ManifestModule.Name.Length - 4));
                                    cmds.Add(cmd);
                                }
                            }
                        }
                    }
                }

            }
            #endregion
            if (cmds.Count > 0)
            {
                List<string> dllName = new List<string>();
                foreach (var item in cmds)
                {
                    if (!dllName.Contains(item.dllName)) dllName.Add(item.dllName);
                }
                foreach (var item in dllName) res.Add(new gcadDllcmd(item, cmds));
            }
            return res;
            //
        }
        public static void AddCmdtoMenuBar(List<gcadDllcmd> cmds)
        {
            var dllName = Assembly.GetExecutingAssembly().ManifestModule.Name.Substring(0, Assembly.GetExecutingAssembly().ManifestModule.Name.Length - 4);
            dynamic gcadApp = Application.AcadApplication;
            dynamic mg = null;
            dynamic count = gcadApp.MenuGroups.Count;
            for (int i = 0; i < count; i++) if (gcadApp.MenuGroups.Item(i).Name == "ACAD") mg = gcadApp.MenuGroups.Item(i);

            for (int i = 0; i < mg.Menus.Count; i++) if (mg.Menus.Item(i).Name == dllName) mg.Menus.Item(i).RemoveFromMenuBar();
            dynamic popMenu = mg.Menus.Add(dllName);
            for (int i = 0; i < cmds.Count; i++)
            {
                var dllPopMenu = popMenu.AddSubMenu(popMenu.Count + 1, cmds[i].DllName);
                for (int j = 0; j < cmds[i].clsCmds.Count; j++)
                {
                    var clsPopMenu = dllPopMenu.AddSubMenu(dllPopMenu.Count + 1, cmds[i].clsCmds[j].clsName);
                    for (int k = 0; k < cmds[i].clsCmds[j].curClscmds.Count; k++)
                    {
                        var methodPopMenu = clsPopMenu.AddMenuItem(clsPopMenu.Count + 1, cmds[i].clsCmds[j].curClscmds[k].cmdName, cmds[i].clsCmds[j].curClscmds[k].cmdMacro + " ");
                    }
                }
            }
            popMenu.InsertInMenuBar(mg.Menus.Count + 1);
        }

        /// <summary>
        /// 将菜单加载到AutoCAD
        /// </summary>
        public static void Register2HKCR()
        {
            string hkcrKey = HostApplicationServices.Current.UserRegistryProductRootKey;
            var assName = Assembly.GetExecutingAssembly().Location;
            var apps_Acad = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(Path.Combine(hkcrKey, "Applications"));
            if (apps_Acad.GetSubKeyNames().Count(c => c == Path.GetFileNameWithoutExtension(assName)) == 0)
            {
                var myNetLoader = apps_Acad.CreateSubKey(Path.GetFileNameWithoutExtension(assName), RegistryKeyPermissionCheck.Default);
                myNetLoader.SetValue("DESCRIPTION", "加载自定义dll文件", Microsoft.Win32.RegistryValueKind.String);
                myNetLoader.SetValue("LOADCTRLS", 2, Microsoft.Win32.RegistryValueKind.DWord);
                //b.注册表键值"LOADCTRLS"控制说明，控制ARX程序的加载方式
                //0x01：Load the application upon detection of proxy object.
                //当代理对像被控知时另载相应ARX程序.
                //0x02：Load the application upon AutoCAD startup.

                //当AutoCAD启动时加载相应ARX程序.
                //0x04：Load the application upon invocation of a command.

                //当输入命令时加载相应ARX程序.
                //0x08：Load the application upon request by the user or another application.
                //当有用户或别的程序请求时加载相应ARX程序.
                //0x10：Do not load the application.
                //从不加载该应用程序.
                //0x20：Load the application transparently.

                //显式加载该应该程序.(不知该项译法是否有误)
                myNetLoader.SetValue("LOADER", assName, Microsoft.Win32.RegistryValueKind.String);
                myNetLoader.SetValue("MANAGED", 1, Microsoft.Win32.RegistryValueKind.DWord);
                Application.ShowAlertDialog(Path.GetFileNameWithoutExtension(assName) + "程序自动加载完成，重启AutoCAD 生效！");
            }
            else Application.ShowAlertDialog(Path.GetFileNameWithoutExtension(assName) + "程序已经启动！欢迎使用");

        }

        public static bool CheckFileReadOnly(string fileName)
        {
            bool inUse = true;
            FileStream fs = null;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.None);
                inUse = false;
            }
            catch { }
            return inUse;//true表示正在使用,false没有使用  
        }


        public static void SaveToFile(MemoryStream ms, string fileName)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                byte[] data = ms.ToArray();

                fs.Write(data, 0, data.Length);
                fs.Flush();

                data = null;
            }
        }
        public static void OpenFile(string fileName)
        {
            System.Diagnostics.Process.Start(fileName);
            System.Diagnostics.Process pro = new System.Diagnostics.Process();
            pro.EnableRaisingEvents = false;
            pro.StartInfo.FileName = "rundll32.exe";
            pro.StartInfo.Arguments = "shell32,OpenAs_RunDLL" + fileName;
            pro.Start();
        }


    }
    /// <summary>
    /// 储存自定义的cad命令的信息的类
    /// </summary>
    public class gcadCmds
    {
        public string clsName { get; set; }
        public string cmdName { get; set; }
        public string cmdMacro { get; set; }
        public string dllName { get; set; }

        public gcadCmds(string _clsName, string _cmdName, string _macro, string _dllName)
        {
            this.dllName = _dllName;
            this.clsName = _clsName;
            this.cmdMacro = _macro;
            this.cmdName = _cmdName;
        }

    }
    /// <summary>
    /// 储存包含自定命令的类
    /// </summary>
    public class gcadClscmd
    {
        public string clsName { get; set; }

        public string dllName { get; set; }

        public bool HasGcadcmds { get; set; }

        public List<gcadCmds> curClscmds { get; set; }

        public gcadClscmd(string _clsName, List<gcadCmds> cmds)
        {
            this.clsName = _clsName;
            this.dllName = cmds.First().dllName;
            var clsCmds = cmds.Where(c => c.clsName == this.clsName).ToList();
            if (clsCmds.Count > 0)
            {
                this.HasGcadcmds = true;
                this.curClscmds = new List<gcadCmds>();
                foreach (var item in clsCmds)
                {
                    if (item.clsName == this.clsName) this.curClscmds.Add(item);
                }

            }
            else this.HasGcadcmds = false;
        }
    }
    /// <summary>
    /// 储存每个dll类的
    /// </summary>
    public class gcadDllcmd
    {
        public string DllName { get; set; }
        public bool HasGcadcls { get; set; }
        public List<gcadClscmd> clsCmds { get; set; }
        public List<gcadCmds> curDllcmds { get; set; }
        public gcadDllcmd(string _dllname, List<gcadCmds> cmds)
        {
            this.DllName = _dllname;
            var curDllcmds = cmds.Where(c => c.dllName == this.DllName).ToList();
            if (curDllcmds.Count > 0)
            {
                this.HasGcadcls = true;
                this.curDllcmds = curDllcmds;
                List<string> listClsName = new List<string>();
                foreach (gcadCmds item in this.curDllcmds)
                {
                    if (!listClsName.Contains(item.clsName)) listClsName.Add(item.clsName);
                }
                this.clsCmds = new List<gcadClscmd>();
                foreach (var item in listClsName)
                {
                    gcadClscmd clsCmds = new gcadClscmd(item, this.curDllcmds.Where(c => c.clsName == item).ToList());
                    this.clsCmds.Add(clsCmds);
                }


            }
            else this.HasGcadcls = false;
        }


    }
}
