using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace MAPI.NET
{
    using Ole = System.Runtime.InteropServices.ComTypes;
    class MAPINative
    {
        [DllImport("MAPI32.dll")]
        internal static extern HRESULT MAPIInitialize(IntPtr lpMapiInit);

        [DllImport("MAPI32.dll")]
        internal static extern void MAPIUninitialize();

        [DllImport("MAPI32.dll")]
        internal static extern int MAPILogonEx(uint ulUIParam, [MarshalAs(UnmanagedType.LPWStr)] string lpszProfileName,
                [MarshalAs(UnmanagedType.LPWStr)] string lpszPassword, uint flFlags, out IntPtr lppSession);

        [DllImport("MAPI32.dll")]
        internal static extern HRESULT MAPIAllocateBuffer(uint size, out IntPtr lpBuffer);

        [DllImport("MAPI32.dll")]
        internal static extern HRESULT MAPIAllocateMore(uint size, IntPtr pObject, out IntPtr lpBuffer);

        [DllImport("MAPI32.dll")]
        internal static extern HRESULT MAPIFreeBuffer(IntPtr lpBuffer);

        [DllImport("MAPI32.DLL", CharSet = CharSet.Ansi, EntryPoint = "HrAllocAdviseSink@12")]
        internal static extern void HrAllocAdviseSink32(OnAdviseCallbackHandler lpfnCallback, IntPtr lpvContext, out IntPtr lppAdviseSink);

        [DllImport("MAPI32.DLL", CharSet = CharSet.Ansi, EntryPoint = "MAPIGetDefaultMalloc@0")]
        static extern IMalloc MAPIGetDefaultMalloc32();
        [DllImport("MAPI32.DLL", CharSet = CharSet.Ansi, EntryPoint = "MAPIGetDefaultMalloc")]
        static extern IMalloc MAPIGetDefaultMalloc64();

        [DllImport("MAPI32.DLL", CharSet = CharSet.Ansi, EntryPoint = "OpenIMsgSession@12")]
        static extern HRESULT OpenIMsgSession32(IMalloc lpMalloc, uint ulFlags, out IntPtr lppMsgSess);
        [DllImport("MAPI32.DLL", CharSet = CharSet.Ansi, EntryPoint = "OpenIMsgSession")]
        static extern HRESULT OpenIMsgSession64(IMalloc lpMalloc, uint ulFlags, out IntPtr lppMsgSess);


        [DllImport("MAPI32.DLL", CharSet = CharSet.Ansi, EntryPoint = "OpenIMsgOnIStg@44")]
        static extern HRESULT OpenIMsgOnIStg32(IntPtr lpMsgSess, AllocateBuffer lpAllocateBuffer, AllocateMore lpAllocateMore, FreeBuffer lpFreeBuffer, IMalloc lpmalloc, IntPtr lpMapiSup, IStorage lpStg, IntPtr lpfMsgCallRelease, uint ulCallerData, uint ulFlags, out IMessage lppMsg);
        [DllImport("MAPI32.DLL", CharSet = CharSet.Ansi, EntryPoint = "OpenIMsgOnIStg")]
        static extern HRESULT OpenIMsgOnIStg64(IntPtr lpMsgSess, AllocateBuffer lpAllocateBuffer, AllocateMore lpAllocateMore, FreeBuffer lpFreeBuffer, IMalloc lpmalloc, IntPtr lpMapiSup, IStorage lpStg, IntPtr lpfMsgCallRelease, uint ulCallerData, uint ulFlags, out IMessage lppMsg);


        [DllImport("MAPI32.dll", CharSet = CharSet.Ansi, EntryPoint = "CloseIMsgSession@4")]
        static extern void CloseIMsgSession32(IntPtr pMsgSess);
        [DllImport("MAPI32.dll", CharSet = CharSet.Ansi, EntryPoint = "CloseIMsgSession")]
        static extern void CloseIMsgSession64(IntPtr pMsgSess);

        [DllImport("MAPI32.dll", CharSet = CharSet.Ansi, EntryPoint = "WrapCompressedRTFStream")]
        internal static extern HRESULT WrapCompressedRTFStream(Ole.IStream compressedStream, uint flags, out Ole.IStream unCompressedStream);

        internal static IMalloc MAPIGetDefaultMalloc()
        {
            switch (IntPtr.Size)
            {
                case 8:
                    return MAPIGetDefaultMalloc64();
                default:
                    return MAPIGetDefaultMalloc32();
            }
        }

        internal static HRESULT OpenIMsgSession(IMalloc lpMalloc, uint ulFlags, out IntPtr lppMsgSess)
        {
            switch (IntPtr.Size)
            {
                case 8:
                    return OpenIMsgSession64(lpMalloc, ulFlags, out lppMsgSess);
                default:
                    return OpenIMsgSession32(lpMalloc, ulFlags, out lppMsgSess);
            }
        }

        internal static HRESULT OpenIMsgOnIStg(IntPtr lpMsgSess, AllocateBuffer lpAllocateBuffer, AllocateMore lpAllocateMore, FreeBuffer lpFreeBuffer, IMalloc lpmalloc, IntPtr lpMapiSup, IStorage lpStg, IntPtr lpfMsgCallRelease, uint ulCallerData, uint ulFlags, out IMessage lppMsg)
        {
            switch (IntPtr.Size)
            {
                case 8:
                    return OpenIMsgOnIStg64(lpMsgSess, lpAllocateBuffer, lpAllocateMore, lpFreeBuffer, lpmalloc, lpMapiSup, lpStg, lpfMsgCallRelease, ulCallerData, ulFlags, out lppMsg);
                default:
                    return OpenIMsgOnIStg32(lpMsgSess, lpAllocateBuffer, lpAllocateMore, lpFreeBuffer, lpmalloc, lpMapiSup, lpStg, lpfMsgCallRelease, ulCallerData, ulFlags, out lppMsg);
            }
        }

        internal static void CloseIMsgSession(IntPtr pMsgSess)
        {
            switch (IntPtr.Size)
            {
                case 8:
                    CloseIMsgSession64(pMsgSess);
                    break;
                default:
                    CloseIMsgSession32(pMsgSess);
                    break;
            }
        }
    }
}
