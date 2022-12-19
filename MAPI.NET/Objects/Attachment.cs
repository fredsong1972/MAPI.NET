using System;
using System.Runtime.InteropServices;

namespace MAPI.NET
{
    using Ole = System.Runtime.InteropServices.ComTypes;

    /// <summary>
    /// The IAttach interface is used to maintain and provide access to the properties of attachments in messages.
    /// </summary>
    [ComVisible(false)]
    [ComImport()]
    [Guid("00020308-0000-0000-C000-000000000046")]
    [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAttach : IMAPIProp
    {
        /// <exclude/>
        new HRESULT GetLastError(int hResult, uint ulFlags, out IntPtr lppMAPIError);
        /// <exclude/>
        new HRESULT SaveChanges(uint ulFlags);
        /// <exclude/>
        new HRESULT GetProps([In, MarshalAs(UnmanagedType.LPArray)] uint[] lpPropTagArray, uint ulFlags, out uint lpcValues, ref IntPtr lppPropArray);
        /// <exclude/>
        new HRESULT GetPropList(uint ulFlags, out IntPtr lppPropTagArray);
        /// <exclude/>
        new HRESULT OpenProperty(uint ulPropTag, ref Guid lpiid, uint ulInterfaceOptions, uint ulFlags, out IntPtr lppUnk);
        /// <exclude/>
        new HRESULT SetProps(uint cValues, IntPtr lpPropArray, out IntPtr lppProblems);
        /// <exclude/>
        new HRESULT DeleteProps(IntPtr lpPropTagArray, out IntPtr lppProblems);
        /// <exclude/>
        new HRESULT CopyTo(uint ciidExclude, ref Guid rgiidExclude, [In, MarshalAs(UnmanagedType.LPArray)] uint[] lpExcludeProps, IntPtr ulUIParam,
            IntPtr lpProgress, ref Guid lpInterface, IntPtr lpDestObj, uint ulFlags, IntPtr lppProblems);
        /// <exclude/>
        new HRESULT CopyProps(IntPtr lpIncludeProps, uint ulUIParam, IntPtr lpProgress, ref Guid lpInterface,
            IntPtr lpDestObj, uint ulFlags, out IntPtr lppProblems);
        /// <exclude/>
        new HRESULT GetNamesFromIDs(out IntPtr lppPropTags, ref Guid lpPropSetGuid, uint ulFlags,
            out uint lpcPropNames, out IntPtr lpppPropNames);
        /// <exclude/>
        new HRESULT GetIDsFromNames(uint cPropNames, ref IntPtr lppPropNames, uint ulFlags, out IntPtr lppPropTags);
    }

    /// <summary>
    /// AttachMehod contains a MAPI-defined constant representing the way the contents of an attachment can be accessed. 
    /// </summary>
    public enum AttachMethod : uint
    {
        /// <summary>
        /// No attachment
        /// </summary>
        NO_ATTACHMENT = 0,
        /// <summary>
        /// By value
        /// </summary>
        BY_VALUE = 1,
        /// <summary>
        /// By reference
        /// </summary>
        BY_REFERENCE = 2,
        /// <summary>
        /// By resolved reference
        /// </summary>
        BY_REF_RESOLVE = 3,
        /// <summary>
        /// By reference only
        /// </summary>
        BY_REF_ONLY = 4,
        /// <summary>
        /// Embedded message
        /// </summary>
        EMBEDDED_MSG = 5,
        /// <summary>
        /// OLE
        /// </summary>
        OLE	= 6
    }

    /// <summary>
    /// Represents an attachment to an e-mail. 
    /// </summary>
    public class Attachment : MAPIObject
    {
        /// <summary>
        /// IAttach interface Guid
        /// </summary>
        Guid IID_IAttach = new Guid("00020308-0000-0000-C000-000000000046");

        const int BufSize = 4096;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="attach"></param>
        public Attachment(IAttach attach)
            : base(attach, null)
        {
        }

        #region Public Properties

        /// <summary>
        /// IAttach interface ID
        /// </summary>
        public override Guid InterfaceID
        {
            get
            {
                return IID_IAttach;
            }
        }

        #endregion


        #region Public Methods
        /// <summary>
        /// Creates a mail attachment using the content from the specified file, and the specified name.
        /// </summary>
        /// <param name="filePath">file path</param>
        /// <param name="name">attachment name</param>
        /// <returns>true if successful; otherwise, false</returns>
        public bool Attach(string filePath, string name)
        {
            HRESULT hResult = HRESULT.E_FAIL;
            if (SetPropertyValues(new IPropValue[] 
            { 
                new MAPIProp((uint)PropTags.PR_ATTACH_METHOD, (int)AttachMethod.BY_VALUE),
                new MAPIProp((uint)PropTags.PR_RENDERING_POSITION, -1),
                new MAPIProp((uint)PropTags.PR_ATTACH_LONG_FILENAME, name),
                new MAPIProp((uint)PropTags.PR_ATTACH_FILENAME, name)
            }))
            {
                int attachSize = 0;
                IntPtr pObject = OpenProperty((uint)PropTags.PR_ATTACH_DATA_BIN, Storage.IID_Stream, OpenPropertyMode.CREATE | OpenPropertyMode.MODIFY);
                if (pObject != IntPtr.Zero)
                {
                    Ole.IStream stream = Marshal.GetObjectForIUnknown(pObject) as Ole.IStream;
                    using (System.IO.Stream fs = System.IO.File.OpenRead(filePath))
                    {
                        byte[] bytes = new byte[BufSize];
                        int bytesRead;
                        while ((bytesRead = fs.Read(bytes, 0, bytes.Length)) > 0)
                        {
                            IntPtr pcbWritten = IntPtr.Zero;
                            stream.Write(bytes, bytesRead, pcbWritten);
                            attachSize += bytesRead;
                        }
                     }
                    stream.Commit((int)STGC.DEFAULT);
                    Marshal.ReleaseComObject(stream);
                    if (SetPropertyValue((uint)PropTags.PR_ATTACH_SIZE, attachSize))
                    {
                        Save(SaveFlags.KEEP_OPEN_READONLY);
                        hResult = HRESULT.S_OK;
                    }
                 }
            }
            return hResult == HRESULT.S_OK;
        }

        /// <summary>
        /// Creates a mail attachment using the embedded MAPI message, and the specified name.
        /// </summary>
        /// <param name="message">embedded MAPIMessage object</param>
        /// <param name="name">attachment name</param>
        /// <returns>true if successful; otherwise, false</returns>
        public bool Attach(MAPIMessage message, string name)
        {
            if (SetPropertyValues(new IPropValue[] 
            { 
                new MAPIProp((uint)PropTags.PR_ATTACH_METHOD, (int)AttachMethod.EMBEDDED_MSG),
                new MAPIProp((uint)PropTags.PR_DISPLAY_NAME, name),
                new MAPIProp((uint)PropTags.PR_ATTACH_FILENAME, name)
            }))
            {
                IntPtr pObject = OpenProperty((uint)PropTags.PR_ATTACH_DATA_OBJ, message.InterfaceID, OpenPropertyMode.CREATE | OpenPropertyMode.MODIFY);
                if (pObject != IntPtr.Zero)
                {
                    IMessage msg = Marshal.GetObjectForIUnknown(pObject) as IMessage;
                    if (message.CopyTo(
                     new uint[] { 
                (uint)PropTags.PR_ACCESS, (uint)PropTags.PR_BODY, (uint)PropTags.PR_RTF_SYNC_BODY_COUNT, (uint)PropTags.PR_RTF_SYNC_BODY_CRC, (uint)PropTags.PR_RTF_SYNC_BODY_TAG, (uint)PropTags.PR_RTF_SYNC_PREFIX_COUNT, (uint)PropTags.PR_RTF_SYNC_TRAILING_COUNT},
                     msg
                     ))
                    {
                        msg.SaveChanges((uint)SaveFlags.Close);
                        Marshal.ReleaseComObject(msg);
                        Save(SaveFlags.KEEP_OPEN_READONLY);
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Save attachment to a specified file.
        /// </summary>
        /// <param name="filePath">file path</param>
        /// <returns>true if successful; otherwise, false</returns>
        public bool SaveToDisk(string filePath)
        {
            IntPtr pObject = OpenProperty((uint)PropTags.PR_ATTACH_DATA_BIN, Storage.IID_Stream, (uint)STGM.READ, OpenPropertyMode.READ);
            if (pObject != IntPtr.Zero)
            {
                Ole.IStream stream = Marshal.GetObjectForIUnknown(pObject) as Ole.IStream;
                IntPtr newPosition = IntPtr.Zero;
                using (System.IO.Stream fs = System.IO.File.Open(filePath, System.IO.FileMode.Create))
                {
                    byte[] bytes = new byte[BufSize];
                    int readBytes = 0;
                    while ((readBytes = Storage.Read(stream, bytes)) > 0)
                    {
                        fs.Write(bytes, 0, readBytes);
                    }
                    fs.Close();
                }
                Marshal.ReleaseComObject(stream);
                return true;
            }
            return false;
        }

        #endregion

    }
}
