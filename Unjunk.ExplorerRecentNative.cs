using System;
using System.Runtime.InteropServices;

namespace Unjunk {
    // Documented Windows API: clear usage for one file without the context menu's
    // persistent hide entry. The AppUserModelID is deliberately fixed to Explorer files.
    public static class ExplorerRecentNative {
        [ComImport, Guid("12337d35-94c6-48a0-bce7-6a9c69d4d600"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IApplicationDestinations {
            void SetAppID([MarshalAs(UnmanagedType.LPWStr)] string appId);
            void RemoveDestination([MarshalAs(UnmanagedType.IUnknown)] object item);
            void RemoveAllDestinations();
        }
        [ComImport, Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IShellItem {
            void BindToHandler(IntPtr context, ref Guid handler, ref Guid iid, out IntPtr result);
            void GetParent(out IShellItem parent);
            void GetDisplayName(uint type, out IntPtr name);
            void GetAttributes(uint mask, out uint attributes);
            void Compare(IShellItem other, uint hints, out int order);
        }
        [DllImport("shell32.dll", CharSet=CharSet.Unicode, PreserveSig=false)]
        static extern void SHCreateItemFromParsingName(string path, IntPtr context, ref Guid iid,
            [MarshalAs(UnmanagedType.Interface)] out IShellItem item);

        public static void RemoveFileHistory(string path) {
            if(String.IsNullOrWhiteSpace(path)) throw new ArgumentException("Recent-file identity is unavailable.");
            IShellItem item=null;
            IApplicationDestinations list=null;
            try {
                Guid iid=typeof(IShellItem).GUID;
                SHCreateItemFromParsingName(path,IntPtr.Zero,ref iid,out item);
                list=(IApplicationDestinations)Activator.CreateInstance(Type.GetTypeFromCLSID(
                    new Guid("86c14003-4d6b-4ef3-a7b4-0506663b2e68")));
                list.SetAppID("Microsoft.Windows.Explorer_RecentFiles");
                list.RemoveDestination(item);
            } finally {
                if(list!=null) Marshal.FinalReleaseComObject(list);
                if(item!=null) Marshal.FinalReleaseComObject(item);
            }
        }
    }
}
