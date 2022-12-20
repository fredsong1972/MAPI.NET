using System;
using System.Runtime.InteropServices;

namespace MAPI.NET
{
    using  Ole = System.Runtime.InteropServices.ComTypes;
    /// <summary>
    /// 
    /// </summary>
    internal enum STGM : uint
    {
        DIRECT = 0x00000000,
        TRANSACTED = 0x00010000,
        SIMPLE = 0x08000000,
        READ = 0x00000000,
        WRITE = 0x00000001,
        READWRITE = 0x00000002,
        SHARE_DENY_NONE = 0x00000040,
        SHARE_DENY_READ = 0x00000030,
        SHARE_DENY_WRITE = 0x00000020,
        SHARE_EXCLUSIVE = 0x00000010,
        PRIORITY = 0x00040000,
        DELETEONRELEASE = 0x04000000,
        NOSCRATCH = 0x00100000,
        CREATE = 0x00001000,
        CONVERT = 0x00020000,
        FAILIFTHERE = 0x00000000,
        NOSNAPSHOT = 0x00200000,
        DIRECT_SWMR = 0x00400000,
    }

    internal enum STGC : uint
    {
        DEFAULT = 0,
        OVERWRITE = 1,
        ONLYIFCURRENT = 2,
        DANGEROUSLYCOMMITMERELYTODISKCACHE = 4,
        CONSOLIDATE = 8
    }

    /// <summary>
    /// C# definition of the IMalloc interface.
    /// </summary>
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("00000002-0000-0000-C000-000000000046")]
    internal interface IMalloc
    {
        /// <summary>
        /// Allocate a block of memory
        /// </summary>
        /// <param name="cb">Size, in bytes, of the memory block to be allocated.</param>
        /// <returns>a pointer to the allocated memory block.</returns>
        [PreserveSig]
        IntPtr Alloc( [In] UInt32 cb);

        /// <summary>
        /// Changes the size of a previously allocated memory block.
        /// </summary>
        /// <param name="pv">Pointer to the memory block to be reallocated</param>
        /// <param name="cb">Size of the memory block, in bytes, to be reallocated.</param>
        /// <returns>reallocated memory block</returns>
        [PreserveSig]
        IntPtr Realloc( [In] IntPtr pv, [In] UInt32 cb);

        /// <summary>
        /// Free a previously allocated block of memory.
        /// </summary>
        /// <param name="pv">Pointer to the memory block to be freed.</param>
        [PreserveSig]
        void Free([In] IntPtr pv);

        /// <summary>
        /// This method returns the size, in bytes, of a memory block previously allocated with IMalloc::Alloc or IMalloc::Realloc.
        /// </summary>
        /// <param name="pv">Pointer to the memory block for which the size is requested</param>
        /// <returns>The size of the allocated memory block in bytes.</returns>
        [PreserveSig]
        UInt32 GetSize([In] IntPtr pv);

        /// <summary>
        /// This method determines whether this allocator was used to allocate the specified block of memory.
        /// </summary>
        /// <param name="pv">Pointer to the memory block</param>
        /// <returns>
        /// 1 - allocated 
        /// 0 - not allocated by this IMalloc Instance.
        /// -1 if DidAlloc is unable to determine whether or not it allocated the memory block.
        /// </returns>
        [PreserveSig]
        Int16 DidAlloc([In] IntPtr pv);

        /// <summary>
        /// Minimizes the heap by releasing unused memory to the operating system.
        /// </summary>
        [PreserveSig]
        void HeapMinimize();
    }

    [ComImport]
    [Guid("0000000d-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IEnumSTATSTG
    {
        // The user needs to allocate an STATSTG array whose size is celt.
        [PreserveSig]
        uint Next(uint celt, [MarshalAs(UnmanagedType.LPArray), Out] Ole.STATSTG[] rgelt, out uint pceltFetched );

        void Skip(uint celt);

        void Reset();

        [return: MarshalAs(UnmanagedType.Interface)]
        IEnumSTATSTG Clone();
    }

    [ComImport]
    [Guid("0000000b-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IStorage
    {
        void CreateStream(
            /* [string][in] */ string pwcsName,
            /* [in] */ uint grfMode,
            /* [in] */ uint reserved1,
            /* [in] */ uint reserved2,
            /* [out] */ out Ole.IStream ppstm);

        void OpenStream(
            /* [string][in] */ string pwcsName,
            /* [unique][in] */ IntPtr reserved1,
            /* [in] */ uint grfMode,
            /* [in] */ uint reserved2,
            /* [out] */ out Ole.IStream ppstm);

        void CreateStorage(
            /* [string][in] */ string pwcsName,
            /* [in] */ uint grfMode,
            /* [in] */ uint reserved1,
            /* [in] */ uint reserved2,
            /* [out] */ out IStorage ppstg);

        void OpenStorage(
            /* [string][unique][in] */ string pwcsName,
            /* [unique][in] */ IStorage pstgPriority,
            /* [in] */ uint grfMode,
            /* [unique][in] */ IntPtr snbExclude,
            /* [in] */ uint reserved,
            /* [out] */ out IStorage ppstg);

        void CopyTo(
            /* [in] */ uint ciidExclude,
            /* [size_is][unique][in] */ Guid rgiidExclude, // should this be an array?
            /* [unique][in] */ IntPtr snbExclude,
            /* [unique][in] */ IStorage pstgDest);

        void MoveElementTo(
            /* [string][in] */ string pwcsName,
            /* [unique][in] */ IStorage pstgDest,
            /* [string][in] */ string pwcsNewName,
            /* [in] */ uint grfFlags);

        void Commit(
            /* [in] */ uint grfCommitFlags);

        void Revert();

        void EnumElements(
            /* [in] */ uint reserved1,
            /* [size_is][unique][in] */ IntPtr reserved2,
            /* [in] */ uint reserved3,
            /* [out] */ out IEnumSTATSTG ppenum);

        void DestroyElement(
            /* [string][in] */ string pwcsName);

        void RenameElement(
            /* [string][in] */ string pwcsOldName,
            /* [string][in] */ string pwcsNewName);

        void SetElementTimes(
            /* [string][unique][in] */ string pwcsName,
            /* [unique][in] */ Ole.FILETIME pctime,
            /* [unique][in] */ Ole.FILETIME patime,
            /* [unique][in] */ Ole.FILETIME pmtime);

        void SetClass(
            /* [in] */ Guid clsid);

        void SetStateBits(
            /* [in] */ uint grfStateBits,
            /* [in] */ uint grfMask);

        void Stat(
            /* [out] */ out Ole.STATSTG pstatstg,
            /* [in] */ uint grfStatFlag);

    }

    class Storage
    {
        [DllImport("ole32.dll")]
        internal static extern HRESULT StgCreateDocfile([MarshalAs(UnmanagedType.LPWStr)]  string pwcsName, STGM grfMode, uint reserved, out IStorage ppstgOpen);

        [DllImport("ole32.dll", CharSet = CharSet.Ansi)]
        internal static extern HRESULT WriteClassStg(IStorage pStg, [In, MarshalAs(UnmanagedType.LPStruct)] Guid rclsid);

        public static Guid IID_Stream = new Guid("0000000c-0000-0000-C000-000000000046");

        public static int Read(Ole.IStream strm,  byte[] buffer)
        {
            IntPtr pRead = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(int)));
            strm.Read(buffer, buffer.Length, pRead);
            int read = Marshal.ReadInt32(pRead);
            Marshal.FreeCoTaskMem(pRead);
            return read;
        }


        public static int Write(Ole.IStream strm, byte[] buffer)
        {
            IntPtr pWritten = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(int)));
            strm.Write(buffer, buffer.Length, pWritten);
            int written = Marshal.ReadInt32(pWritten);
            Marshal.FreeCoTaskMem(pWritten);
            return written;
        }
    }
}
