
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ViewModel.PointCloudManager;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MyNetloadX
{
    public class AutoCADNetLoader : IExtensionApplication
    {
        public string currentDllPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
        /// <summary>
        /// 写入程序到注册表，实现加载一次后，自动加载到AutoCAD
        /// </summary>
        /// <param name="reg_Key">加入到那个注册表，HKLM?,HKCU?</param>
        public void AcadNetDllAutoLoader(RegistryKey reg_Key)
        {
            string regPath = ""; bool flag_currentApp = true;
            if (reg_Key.ToString() == Registry.LocalMachine.ToString())
            {
                regPath = Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.Current.MachineRegistryProductRootKey;
            }
            else regPath = Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.Current.UserRegistryProductRootKey;
            string assemblyFileName = Assembly.GetExecutingAssembly().CodeBase;
            RegistryKey acad_key = reg_Key.OpenSubKey(Path.Combine(regPath, "Applications"), false);
            foreach (var item in acad_key.GetSubKeyNames()) if (item == Path.GetFileNameWithoutExtension(assemblyFileName)) flag_currentApp = false;
            if (flag_currentApp)
            {
                acad_key = reg_Key.OpenSubKey(Path.Combine(regPath, "Applications"), true);
                RegistryKey myAppkey = acad_key.CreateSubKey(Path.GetFileNameWithoutExtension(assemblyFileName), Microsoft.Win32.RegistryKeyPermissionCheck.Default);
                myAppkey.SetValue("DESCRIPTION", "加载自定义Dll");
                myAppkey.SetValue("LOADCTRLS", 0x02, Microsoft.Win32.RegistryValueKind.DWord);
                myAppkey.SetValue("LOADER", assemblyFileName, Microsoft.Win32.RegistryValueKind.String);
                myAppkey.SetValue("MANAGED", 0x01, Microsoft.Win32.RegistryValueKind.DWord);
                Application.ShowAlertDialog($"{Path.GetFileNameWithoutExtension(assemblyFileName)} 程序加载完成，重启CAD生效！");
            }
            else Application.ShowAlertDialog($"欢迎使用 {Path.GetFileNameWithoutExtension(assemblyFileName)} 程序！");
        }

        public void Initialize()//初始化程序。
        {
            AcadNetDllAutoLoader(Registry.CurrentUser);                //写入注册表
            Application.SetSystemVariable("SECURELOAD", 0);
        }

        public void Terminate()
        {
            throw new NotImplementedException();
        }
    }

    public class Tools
    {
        [CommandMethod("MyNetLoader")]
        public void Dll加载到内存并卸载()
        {
            Autodesk.AutoCAD.Windows.OpenFileDialog ofd = new Autodesk.AutoCAD.Windows.OpenFileDialog("Dll加载到内存并卸载", "Dll加载到内存并卸载", "dll", "Dll加载到内存并卸载", Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.AllowMultiple);
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (var item in ofd.GetFilenames())
                {
                    var assBytes = System.IO.File.ReadAllBytes(item);
                    Assembly assembly = Assembly.Load(assBytes);
                }
            }
            else Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("请选择一个或者多个dll文件！");
        }
    }


    // Assembly没有Unload的功能，但可以使用AppDomain来解决这个问题。基本思路是，创建一个新的AppDomain，在这个新建的AppDomain中装载assembly，调用其中的方法，然后将获得的结果返回。在完成所有操作以后，调用AppDomain.Unload方法卸载这个新建的AppDomain，这样也同时卸载了assembly。注意：你无法将装载的assembly直接返回到当前应用程序域（AppDomain）。

    //首先，创建一个RemoteLoader，这个RemoteLoader用于在新建的AppDomain中装载assembly，并向外公布一个属性，以便外界能够获得assembly的FullName。RemoteLoader需要继承于MarshalByRefObject。代码如下：
    public class RemoteLoader : MarshalByRefObject
    {
        private Assembly assembly;
        public void LoadAssembly(string fullName)
        {
            var assBytes = System.IO.File.ReadAllBytes(fullName);
            assembly = Assembly.Load(assBytes);
            //assembly = Assembly.LoadFrom(fullName);
        }
        public string FullName
        {
            get { return assembly.FullName; }
        }
    }

    //其次，创建一个LocalLoader。LocalLoader的功能是创建新的AppDomain，然后在这个新的AppDomain中调用RemoteLoader，以便通过RemoteLoader来创建assembly并获得assembly的相关信息。此时被调用的assembly自然被装载于新的AppDomain中。最后，LocalLoader还需要提供一个新的方法，就是AppDomain的卸载。代码如下：
    public class LocalLoader
    {
        private AppDomain appDomain;
        private RemoteLoader remoteLoader;

        public LocalLoader()
        {
            AppDomainSetup setup = new AppDomainSetup();
            setup.ApplicationName = "Test";
            setup.ApplicationBase = AppDomain.CurrentDomain.BaseDirectory;
            setup.PrivateBinPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "private");
            setup.CachePath = setup.ApplicationBase;
            setup.ShadowCopyFiles = "true";
            setup.ShadowCopyDirectories = setup.ApplicationBase;

            appDomain = AppDomain.CreateDomain("TestDomain", null, setup);
            string name = Assembly.GetExecutingAssembly().GetName().FullName;
            remoteLoader = (RemoteLoader)appDomain.CreateInstanceAndUnwrap(
                name,
                typeof(RemoteLoader).FullName);
        }

        public void LoadAssembly(string fullName)
        {
            remoteLoader.LoadAssembly(fullName);
        }

        public void Unload()
        {
            AppDomain.Unload(appDomain);
            appDomain = null;
        }

        public string FullName
        {
            get
            {
                return remoteLoader.FullName;
            }
        }
    }

}
