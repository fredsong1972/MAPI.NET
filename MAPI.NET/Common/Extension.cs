using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace MAPI.NET
{
    /// <summary>
    /// Extension
    /// </summary>
    static class Extension
    {
        public static IntPtr MarshalToAdrListPtr(this ADRENTRY[] entries)
        {
            if (entries.Length == 0)
                return IntPtr.Zero;
            int adrListSize = Marshal.SizeOf(typeof(ADRLIST));
            IntPtr adrListPtr = Marshal.AllocHGlobal(adrListSize + (entries.Length -1) * Marshal.SizeOf(typeof(ADRENTRY)));
            ADRLIST adrList = new ADRLIST();
            adrList.cEntries = (uint)entries.Length;
            adrList.aEntries = entries[0];
            Marshal.StructureToPtr(adrList, adrListPtr, true);
            for (int i = 1; i < entries.Length; i++)
            {
                Marshal.StructureToPtr(entries[i], adrListPtr + adrListSize + (i -1)* Marshal.SizeOf(typeof(ADRENTRY)), true);
            }
            return adrListPtr;
        }

        public static IntPtr MarshalToIntPtr(this pSPropValue[] propValues)
        {
            if (propValues.Length == 0)
                return IntPtr.Zero;
            IntPtr pValues = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(pSPropValue)) * propValues.Length);
            for (int i = 0; i < propValues.Length; i++)
            {
                pSPropValue val = propValues[i];
                Marshal.StructureToPtr(val, pValues + i * Marshal.SizeOf(typeof(pSPropValue)), true);
            }
            return pValues;
         }

        public static void MarshalFreeIntPtr(this pSPropValue[] propValues, IntPtr p)
        {
            for (int i = 0; i < propValues.Length; i++)
            {
                pSPropValue val = propValues[i];
                switch ((PT)((uint)val.ulPropTag & 0xFFFF))
                {
                    case PT.PT_TSTRING:
                        if (val.Value.lpszW != IntPtr.Zero)
                            Marshal.FreeHGlobal(val.Value.lpszW);
                        break;

                    case PT.PT_BINARY:
                        if (val.Value.bin.lpb != IntPtr.Zero)
                            SBinary.SBinaryRelease(ref val.Value.bin);
                        break;

                }
            }
            if (p != null && p != IntPtr.Zero)
                Marshal.FreeHGlobal(p);
        }
    }
}
