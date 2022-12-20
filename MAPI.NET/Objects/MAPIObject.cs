using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace MAPI.NET
{
    /// <summary>
    /// Enables clients, service providers, and MAPI to work with properties. All objects that support properties implement this interface.
    /// </summary>
    [
       ComImport, ComVisible(false),
       InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
       Guid("00020303-0000-0000-C000-000000000046")
   ]
    public interface IMAPIProp
    {
        /// <summary>
        /// Returns a MAPIERROR structure that contains information about the previous error.
        /// </summary>
        /// <param name="hResult">A handle to the error code generated in the previous method call.</param>
        /// <param name="ulFlags">A bitmask of flags that indicates the format for the text returned in the MAPIERROR structure pointed to by lppMAPIError.</param>
        /// <param name="lppMAPIError">A pointer to a pointer to the MAPIERROR structure that contains version, component, and context information for the error. The lppMAPIError parameter can be set to NULL if there is no error information to return.</param>
        /// <returns>S_OK, if the error information was returned;otherwise failed.</returns>
        [PreserveSig]
        HRESULT GetLastError(int hResult, uint ulFlags, out IntPtr lppMAPIError);
        /// <summary>
        /// Makes permanent any changes that were made to an object since the last save operation.
        /// </summary>
        /// <param name="ulFlags"> A bitmask of flags that controls what happens to the object when the IMAPIProp::SaveChanges method is called.</param>
        /// <returns>
        /// S_OK
        ///     The commitment of changes was successful.
        /// MAPI_E_NO_ACCESS
        ///     SaveChanges cannot keep the object open for read-only permission if KEEP_OPEN_READONLY is set, or read/write permission if KEEP_OPEN_READWRITE is set. No changes are committed.
        /// MAPI_E_OBJECT_CHANGED
        ///     The object has changed since it was opened.
        /// MAPI_E_OBJECT_DELETED
        ///     The object has been deleted since it was opened.
        /// </returns>
        [PreserveSig]
        HRESULT SaveChanges(uint ulFlags);
        /// <summary>
        /// Retrieves the property value of one or more properties of an object.
        /// </summary>
        /// <param name="lpPropTagArray"> A pointer to an array of property tags that identify the properties whose values are to be retrieved. The lpPropTagArray parameter must be either NULL, indicating that values for all properties of the object should be returned, or point to an SPropTagArray structure that contains one or more property tags.</param>
        /// <param name="ulFlags">A bitmask of flags that indicates the format for properties that have the type PT_UNSPECIFIED. </param>
        /// <param name="lpcValues">A pointer to a count of property values pointed to by the lppPropArray parameter. If lppPropArray is NULL, the content of the lpcValues parameter is zero.</param>
        /// <param name="lppPropArray">A pointer to the retrieved property values.</param>
        /// <returns>
        /// S_OK
        ///     The property values were successfully retrieved.
        /// MAPI_W_ERRORS_RETURNED
        ///     The call succeeded overall, but one or more properties could not be accessed. The aulPropTag member of the property value for each unavailable property has a property type of PT_ERROR and an identifier of zero. When this warning is returned, the call should be handled as successful. 
        ///  MAPI_E_INVALID_PARAMETER
        ///     Zero was passed in the cValues member of the SPropTagArray structure pointed to by lpPropTagArray.
        /// </returns>
        HRESULT GetProps([In, MarshalAs(UnmanagedType.LPArray)] uint[] lpPropTagArray, uint ulFlags, out uint lpcValues, ref IntPtr lppPropArray);
        /// <summary>
        /// Returns property tags for all properties.
        /// </summary>
        /// <param name="ulFlags">A bitmask of flags that controls the format for the strings in the returned property tags.</param>
        /// <param name="lppPropTagArray">A pointer to a pointer to the property tag array that contains tags for all of the object's properties.</param>
        /// <returns>
        /// S_OK
        ///     All of the property tags were returned successfully.
        /// MAPI_E_BAD_CHARWIDTH
        ///     Either the MAPI_UNICODE flag was set and the implementation does not support Unicode, or MAPI_UNICODE was not set and the implementation supports only Unicode.
        /// </returns>
        [PreserveSig]
        HRESULT GetPropList(uint ulFlags, out IntPtr lppPropTagArray);
        /// <summary>
        /// Returns a pointer to an interface that can be used to access a property.
        /// </summary>
        /// <param name="ulPropTag">The property tag for the property to be accessed. Both the identifier and the type must be included in the property tag</param>
        /// <param name="lpiid"> A pointer to the identifier for the interface to be used to access the property. The lpiid parameter must not be null.</param>
        /// <param name="ulInterfaceOptions">Data that relates to the interface identified by the lpiid parameter.</param>
        /// <param name="ulFlags"> A bitmask of flags that controls access to the property.</param>
        /// <param name="lppUnk">A pointer to the requested interface to be used for property access.</param>
        /// <returns>S_OK, if the requested interface pointer was successfully returned; otherwise, failed.</returns>
        HRESULT OpenProperty(uint ulPropTag, ref Guid lpiid, uint ulInterfaceOptions, uint ulFlags, out IntPtr lppUnk);
        /// <summary>
        /// Updates one or more properties.
        /// </summary>
        /// <param name="cValues">The count of property values pointed to by the lpPropArray parameter. The cValues parameter must not be 0.</param>
        /// <param name="lpPropArray">A pointer to an array of SPropValue structures that contain property values to be updated.</param>
        /// <param name="lppProblems">On input, a pointer to a pointer to an SPropProblemArray structure; otherwise, NULL, indicating no need for error information. If lppProblems is a valid pointer on input, SetProps returns detailed information about errors in updating one or more properties.</param>
        /// <returns>S_OK, if the properties were successfully updated.; otherwise, failed.</returns>
        [PreserveSig]
        HRESULT SetProps(uint cValues, IntPtr lpPropArray, out IntPtr lppProblems);
        /// <summary>
        /// Deletes one or more properties from an object.
        /// </summary>
        /// <param name="lpPropTagArray">A pointer to an array of property tags that indicate the properties to delete. The cValues member of the SPropTagArray structure pointed to by lpPropTagArray must not be zero, and the lpPropTagArray parameter itself must not be NULL.</param>
        /// <param name="lppProblems">On input, a pointer to a pointer to an SPropProblemArray structure; otherwise, NULL, which indicates that there is no need for error information. If lppProblems is a valid pointer on input, DeleteProps returns detailed information about errors in deleting one or more properties.</param>
        /// <returns>S_OK, if properties were successfully deleted;otherwise, failed.</returns>
        [PreserveSig]
        HRESULT DeleteProps(IntPtr lpPropTagArray, out IntPtr lppProblems);
        /// <summary>
        /// Copies or moves all properties, except for specifically excluded properties.
        /// </summary>
        /// <param name="ciidExclude">The count of interfaces to exclude when properties are copied or moved.</param>
        /// <param name="rgiidExclude">An array of interface identifiers (IIDs) that specify interfaces that should not be used when supplemental information is copied or moved to the destination object.</param>
        /// <param name="lpExcludeProps"> A pointer to a property tag array that identifies the property tags that should be excluded from the copy or move operation. Passing null in the lpExcludeProps parameter indicates that all of the object's properties should be copied or moved. CopyTo returns MAPI_E_INVALID_PARAMETER when the cValues member of the SPropProblemArray structure pointed to by lpExcludeProps is set to 0.</param>
        /// <param name="ulUIParam"> A handle to the parent window of the progress indicator.</param>
        /// <param name="lpProgress">A pointer to a progress indicator implementation. If null is passed in the lpProgress parameter, MAPI provides the progress implementation. The lpProgress parameter is ignored unless the MAPI_DIALOG flag is set in the ulFlags parameter.</param>
        /// <param name="lpInterface">A pointer to the interface identifier (IID) that represents the interface to be used to access the object pointed to by the lpDestObj parameter. The lpInterface parameter must not be null.</param>
        /// <param name="lpDestObj">A pointer to the object to receive the copied or moved properties.</param>
        /// <param name="ulFlags"> A bitmask of flags that controls the copy or move operation. </param>
        /// <param name="lppProblems">On input, a pointer to a pointer to an SPropProblemArray structure; otherwise, null, indicating no need for error information. If lppProblems is a valid pointer on input, CopyTo returns detailed information about errors in copying one or more properties.</param>
        /// <returns>S_OK, if the properties have been successfully copied or moved;otherwise failed.</returns>
        [PreserveSig]
        HRESULT CopyTo(uint ciidExclude, ref Guid rgiidExclude, [In, MarshalAs(UnmanagedType.LPArray)] uint[] lpExcludeProps, IntPtr ulUIParam,
            IntPtr lpProgress, ref Guid lpInterface, IntPtr lpDestObj, uint ulFlags, IntPtr lppProblems);
        /// <summary>
        /// Copies or moves selected properties.
        /// </summary>
        /// <param name="lpIncludeProps">A pointer to a property tag array that specifies the properties to copy or move. PR_NULL (PidTagNull) cannot be included in the array. The lpIncludeProps parameter cannot be null.</param>
        /// <param name="ulUIParam"> A handle to the parent window of the progress indicator.</param>
        /// <param name="lpProgress"> A pointer to an implementation of a progress indicator. If null is passed in the lpProgress parameter, the progress indicator is displayed by using the MAPI implementation. The lpProgress parameter is ignored unless the MAPI_DIALOG flag is set in the ulFlags parameter.</param>
        /// <param name="lpInterface">A pointer to the interface identifier (IID) that represents the interface that must be used to access the object pointed to by the lpDestObj parameter. The lpInterface parameter must not be null.</param>
        /// <param name="lpDestObj"> A pointer to the object to receive the copied or moved properties.</param>
        /// <param name="ulFlags">A bitmask of flags that controls the copy or move operation.</param>
        /// <param name="lppProblems">On input, a pointer to a pointer to an SPropProblemArray structure; otherwise, null, indicating that there is no need for error information. If lppProblems is a valid pointer on input, CopyProps returns detailed information about errors in copying one or more properties.</param>
        /// <returns>S_OK, if properties have been successfully copied or moved;otherwise failed.</returns>
        [PreserveSig]
        HRESULT CopyProps(IntPtr lpIncludeProps, uint ulUIParam, IntPtr lpProgress, ref Guid lpInterface,
            IntPtr lpDestObj, uint ulFlags, out IntPtr lppProblems);
        /// <summary>
        /// Provides the property names that correspond to one or more property identifiers.
        /// </summary>
        /// <param name="lppPropTags">On input, a pointer to an SPropTagArray structure that contains an array of property tags; otherwise, NULL, indicating that all names should be returned. The cValues member for the property tag array cannot be 0. If lppPropTags is a valid pointer on input, GetNamesFromIDs returns names for each property identifier included in the array.</param>
        /// <param name="lpPropSetGuid"> A pointer to a GUID, or GUID structure, that identifies a property set. The lpPropSetGuid parameter can point to a valid property set or can be NULL.</param>
        /// <param name="ulFlags">A bitmask of flags that indicates the type of names to be returned.</param>
        /// <param name="lpcPropNames">A pointer to a count of the property name pointers in the array pointed to by the lppPropNames parameter.</param>
        /// <param name="lpppPropNames"> A pointer to an array of pointers to MAPINAMEID structures that contains property names.</param>
        /// <returns>S_OK, if the property names were successfully returned;otherwise failed.</returns>
        [PreserveSig]
        HRESULT GetNamesFromIDs(out IntPtr lppPropTags, ref Guid lpPropSetGuid, uint ulFlags,
            out uint lpcPropNames, out IntPtr lpppPropNames);
        /// <summary>
        /// Provides the property identifiers that correspond to one or more property names.
        /// </summary>
        /// <param name="cPropNames">The count of property names pointed to by the lppPropNames parameter. If lppPropNames is NULL, the cPropNames parameter must be 0.</param>
        /// <param name="lppPropNames"> A pointer to an array of property names, or NULL. Passing NULL requests property identifiers for all property names in all property sets about which the object has information. The lppPropNames parameter must not be NULL if the MAPI_CREATE flag is set in the ulFlags parameter.</param>
        /// <param name="ulFlags"> A bitmask of flags that indicates how the property identifiers should be returned. </param>
        /// <param name="lppPropTags">A pointer to a pointer to an array of property tags that contains existing or newly assigned property identifiers.</param>
        /// <returns>S_OK, if the identifiers for the specified property names were successfully returned;otherwise failed.</returns>
        [PreserveSig]
        HRESULT GetIDsFromNames(uint cPropNames, ref IntPtr lppPropNames, uint ulFlags, out IntPtr lppPropTags);
    }

    /// <summary>
    /// A bitmask of flags that controls access to the property.
    /// </summary>
    public enum OpenPropertyMode
    {
        /// <summary>
        /// Read
        /// </summary>
        READ = 0,
        /// <summary>
        /// Modify
        /// </summary>
        MODIFY = 1,
        /// <summary>
        /// If the property does not exist, it should be created. If the property does exist, the current value of the property should be discarded. When a caller sets the MAPI_CREATE flag, it should also set the MAPI_MODIFY flag.
        /// </summary>
        CREATE = 2,
        /// <summary>
        /// Append
        /// </summary>
        APPEND = 4,
    }

    /// <summary>
    /// IMAPIProp .Net Wrapper object
    /// </summary>
    public class MAPIObject : IDisposable
    {
        /// <exclude/>
        public const string MailItem = "IPM.Note";
        /// <exclude/>
        public const string ContactItem = "IPM.Contact";
        /// <exclude/>
        public const string AppointmentItem = "IPM.Appointment";
        /// <exclude/>
        public const string MeetingItem = "IPM.Schedule.Meeting";
        /// <exclude/>
        public const string MeetingRequestItem = "IPM.Schedule.Meeting.Request";
        /// <exclude/>
        public const string MeetingCanceledItem = "IPM.Schedule.Meeting.Canceled";
        /// <exclude/>
        public const string MeetingResponseItem = "IPM.Schedule.Meeting.Resp";

        static Guid IID_IMAPIProp = new Guid("00020303-0000-0000-C000-000000000046");
        enum MAPINameIDKind : int
        {
            MNID_ID = 0,
            MNID_STRING = 1
        }
        internal IMAPIProp mapiObj_;
        /// <summary>
        /// EntryID of the object
        /// </summary>
        protected EntryID Id_;

        /// <summary>
        /// Initializes a new instance of the MAPIObject class. 
        /// </summary>
        /// <param name="session">MAPISession object</param>
        /// <param name="id">EntryID of MAPIObject</param>
        public MAPIObject(MAPISession session, EntryID id)
        {
            mapiObj_ = session.OpenEntry(id);
        }
        /// <summary>
        /// Initializes a new instance of the MAPIObject class.
        /// </summary>
        /// <param name="obj">IMAPIProp object</param>
        /// <param name="id">EntryID of MAPIObject</param>
        public MAPIObject(IMAPIProp obj, EntryID id)
        {
            mapiObj_ = obj;
            Id_ = id;
        }
        /// <summary>
        /// Initializes a new instance of the MAPIObject class.
        /// </summary>
        /// <param name="obj">IMAPIProp object</param>
        public MAPIObject(IMAPIProp obj)
        {
            mapiObj_ = obj;
            IPropValue entry = GetProperty(PropTags.PR_ENTRYID);
            if (entry != null && entry.AsBinary != null)
            {
                Id_ = new EntryID(entry.AsBinary);
            }
        }

        #region Public Properties
        /// <summary>
        /// IMAPIProp Interface Guid
        /// </summary>
        public virtual Guid InterfaceID
        {
            get { return IID_IMAPIProp; }
        }
        /// <summary>
        /// EntryID of MAPIObject
        /// </summary>
        public EntryID EntryID { get { return Id_; } }


        #endregion

        #region Public Methods
        /// <summary>
        /// Retrieves the property value of one or more properties of an object.
        /// </summary>
        /// <param name="tags">Array of PropTags</param>
        /// <returns>A dictionary of the retrieved property values.</returns>
        public Dictionary<PropTags, IPropValue> GetProperties(PropTags[] tags)
        {
            List<uint> utags = new List<uint>();
            foreach (PropTags tag in tags)
                utags.Add((uint)tag);
            Dictionary<PropTags, IPropValue> props = new Dictionary<PropTags, IPropValue>();
            Dictionary<uint, IPropValue> uprops = GetProperties(utags.ToArray());
            foreach (uint tag in uprops.Keys)
            {
                props[(PropTags)tag] = uprops[tag];
            }
            return props;
        }

        /// <summary>
        /// Retrieves the property value of one or more properties of an object.
        /// </summary>
        /// <param name="tags">Array of uint tags</param>
        /// <returns>A dictionary of the retrieved property values.</returns>
        public Dictionary<uint, IPropValue> GetProperties(uint[] tags)
        {
            Dictionary<uint, IPropValue> props = new Dictionary<uint, IPropValue>();
            IntPtr propVals = IntPtr.Zero;
            uint ivalues = 0;
            uint iprops = (uint)tags.Length;
            uint[] p = new uint[tags.Length + 1];

            p[0] = (uint)tags.Length;
            for (int i = 0; i < tags.Length; i++)
                p[i + 1] = tags[i];

            HRESULT hResult = HRESULT.E_FAIL;
            try
            {
                hResult = mapiObj_.GetProps(p, (uint)CharacterSet.UNICODE, out ivalues, ref propVals);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.Message);
            }

            if (propVals != IntPtr.Zero)
            {

                for (int i = 0; i < ivalues; i++)
                {
                    int size = Marshal.SizeOf(typeof(pSPropValue));
                    pSPropValue lpProp = (pSPropValue)Marshal.PtrToStructure(propVals + i * Marshal.SizeOf(typeof(pSPropValue)), typeof(pSPropValue));
                    if ((lpProp.ulPropTag & 0xFFFF) == (uint)PT.PT_ERROR)
                        continue;
                    if ((lpProp.ulPropTag == tags[i]) || ((lpProp.ulPropTag & 0xFFFF0000) == tags[i]))
                    {
                        MAPIProp prop = new MAPIProp(lpProp);
                        props[tags[i]] = prop;
                    }
                }
                MAPINative.MAPIFreeBuffer(propVals);
            }
            return props;
        }
        /// <summary>
        /// Retrieves one property value of an object.
        /// </summary>
        /// <param name="tag">Prop tag</param>
        /// <returns>The retrieved property value</returns>
        public IPropValue GetProperty(PropTags tag)
        {
            return GetProperty((uint)tag);
        }

        /// <summary>
        /// Retrieves one property value of an object.
        /// </summary>
        /// <param name="tag">uint tag</param>
        /// <returns>The retrieved property value</returns>
        public IPropValue GetProperty(uint tag)
        {
            Dictionary<uint, IPropValue> props = GetProperties(new uint[] { tag });
            return props.ContainsKey(tag) ? props[tag] : null;
        }
        /// <summary>
        /// Updates one property of the object.
        /// </summary>
        /// <param name="tag">Prop tag</param>
        /// <param name="value">Property value to be updated</param>
        /// <returns>true if successful; otherwise, false</returns>
        public bool SetPropertyValue(PropTags tag, object value)
        {
            return SetPropertyValue((uint)tag, value);
        }
        /// <summary>
        /// Updates one property of the object.
        /// </summary>
        /// <param name="tag">uint tag</param>
        /// <param name="value">Property value to be updated</param>
        /// <returns>true if successful; otherwise, false</returns>
        public bool SetPropertyValue(uint tag, object value)
        {
            // Create MAPIProp
            MAPIProp prop = new MAPIProp(tag, value);
            if (prop.Type == null)
                return false;
            //If the field is NOT a fetch properties
            return SetPropertyValue(prop);
        }
        /// <summary>
        /// Makes permanent any changes that were made to an object since the last save operation,and keep the object open.
        /// </summary>
        public void Save()
        {
            Save(SaveFlags.KEEP_OPEN_READWRITE);
        }
        /// <summary>
        /// Makes permanent any changes that were made to an object since the last save operation,and control the object per the flag.
        /// </summary>
        /// <param name="flags">A bitmask of flags that controls what happens to the object</param>
        public void Save(SaveFlags flags)
        {
            mapiObj_.SaveChanges((uint)flags);
        }
        /// <summary>
        /// Makes permanent any changes that were made to an object since the last save operation,and close the object.
        /// </summary>
        public void Close()
        {
            Save(SaveFlags.Close);
            Dispose();
        }

        /// <summary>
        /// Copies or moves all properties, except for specifically excluded properties.
        /// </summary>
        /// <param name="excludeProps">>A property tag array that identifies the property tags that should be excluded from the copy or move operation</param>
        /// <param name="mapiObject">A mapi prop object to receive the copied or moved properties</param>
        /// <returns>true if successful; otherwise, false</returns>
        public bool CopyTo(uint[] excludeProps, MAPIObject mapiObject)
        {
            return CopyTo(excludeProps, mapiObject.mapiObj_);
        }

        /// <summary>
        /// Copies or moves all properties, except for specifically excluded properties.
        /// </summary>
        /// <param name="excludeProps">A property tag array that identifies the property tags that should be excluded from the copy or move operation.</param>
        /// <param name="mapiObject">A mapi prop object to receive the copied or moved properties.</param>
        /// <returns>true if successful; otherwise, false</returns>
        public bool CopyTo(uint[] excludeProps, IMAPIProp mapiObject)
        {
            IntPtr pObject = Marshal.GetIUnknownForObject(mapiObject);
            uint[] p = new uint[excludeProps.Length + 1];
            p[0] = (uint)excludeProps.Length;
            for (int i = 0; i < excludeProps.Length; i++)
                p[i + 1] = excludeProps[i];
            HRESULT hResult = HRESULT.E_FAIL;
            Guid excludeIid = Guid.Empty;
            Guid iid = InterfaceID;
            hResult = mapiObj_.CopyTo(0, ref excludeIid, p, IntPtr.Zero, IntPtr.Zero, ref iid, pObject, 0, IntPtr.Zero);
            return hResult == HRESULT.S_OK;
        }
        /// <summary>
        /// Returns a pointer to an interface that can be used to access a property.
        /// </summary>
        /// <param name="tag">The property tag for the property to be accessed</param>
        /// <param name="iid">the Guid for the interface to be used to access the property</param>
        /// <param name="mode">A bitmask of flags that controls access to the property</param>
        /// <returns>the requested interface pointer</returns>
        public IntPtr OpenProperty(uint tag, Guid iid, OpenPropertyMode mode)
        {
            return OpenProperty(tag, iid, 0, mode);
        }
        /// <summary>
        /// Returns a pointer to an interface that can be used to access a property.
        /// </summary>
        /// <param name="tag">The property tag for the property to be accessed</param>
        /// <param name="iid">the Guid for the interface to be used to access the property</param>
        /// <param name="option">Data that relates to the interface identified by the Guid.</param>
        /// <param name="mode">A bitmask of flags that controls access to the property</param>
        /// <returns>the requested interface pointer</returns>
        public IntPtr OpenProperty(uint tag, Guid iid, uint option, OpenPropertyMode mode)
        {
            IntPtr pObject;
            try
            {
                HRESULT hResult = mapiObj_.OpenProperty(tag, ref iid, option, (uint)mode, out pObject);
                if (hResult == HRESULT.S_OK)
                    return pObject;
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.Message);
            }
            return IntPtr.Zero;
        }

        /// <summary>
        /// Get outlook properties
        /// </summary>
        /// <param name="data">data</param>
        /// <param name="tags">array of tags</param>
        /// <returns>dictionary of properties</returns>
        public Dictionary<uint, IPropValue> GetOutlookProperties(int data, uint[] tags)
        {
            uint[] outlookTags = GetOutlookPropTags(data, tags, false);
            if (outlookTags.Length == 0)
                return new Dictionary<uint, IPropValue>();
            return GetProperties(outlookTags);
        }

        /// <summary>
        /// Get outlook properties
        /// </summary>
        /// <param name="data">data</param>
        /// <param name="tag">tag</param>
        /// <returns>property</returns>
        public IPropValue GetOutlookProperty(int data, uint tag)
        {
            uint[] outlookTags = GetOutlookPropTags(data, new uint[] { tag }, false);
            if (outlookTags.Length == 0)
                return null;
            return GetProperty(outlookTags[0]);
        }

        /// <exclude/>
        public static explicit operator IntPtr (MAPIObject mapiObject)
        {
            IntPtr p = IntPtr.Zero;
            if (mapiObject.mapiObj_ != null)
                p = Marshal.GetIUnknownForObject(mapiObject.mapiObj_);
            return p;
        }

        #endregion


        #region Private Methods/Events
        /// <summary>
        /// Updates one property of the object.
        /// </summary>
        /// <param name="prop">A IPropValue object that contain property values to be updated.</param>
        /// <returns>true if successful; otherwise, false</returns>
        protected bool SetPropertyValue(IPropValue prop)
        {
            return SetPropertyValues(new IPropValue[] { prop });
        }
        /// <summary>
        /// Updates one or more properties of the object.
        /// </summary>
        /// <param name="props">An array of IPropValue that contain property values to be updated.</param>
        /// <returns>true if successful; otherwise, false</returns>
        protected bool SetPropertyValues(IPropValue[] props)
        {
            int num = props.Length;
            List<pSPropValue> values = new List<pSPropValue>();
            for (int i = 0; i < props.Length; i++)
            {
                IPropValue prop = props[i];
                pSPropValue val = new pSPropValue();
                val.ulPropTag = (uint)prop.Tag;
                switch ((PT)(val.ulPropTag & 0xFFFF))
                {
                    case PT.PT_BINARY:
                        val.Value.bin = SBinary.SBinaryCreate(prop.AsBinary);
                        break;
                    case PT.PT_TSTRING:
                        val.Value.lpszW = Marshal.StringToHGlobalUni(prop.AsString);
                        break;
                    case PT.PT_I2:
                    case PT.PT_LONG:
                    case PT.PT_BOOLEAN:
                        val.Value.ul = (uint)prop.AsInt32;
                        break;
                    case PT.PT_SYSTIME:
                    case PT.PT_I8:
                        val.Value.li = (ulong)prop.AsUInt64;
                        break;
                }
                values.Add(val);
            }
            IntPtr pProblems;
            IntPtr array = values.ToArray().MarshalToIntPtr();
            HRESULT hResult = mapiObj_.SetProps((uint)num, array, out pProblems);
            values.ToArray().MarshalFreeIntPtr(array);
            return hResult == HRESULT.S_OK;
        }

        /// <summary>
        /// Get outlook property tags
        /// </summary>
        /// <param name="data">data</param>
        /// <param name="properties">properties</param>
        /// <param name="bCreate">create tag if not exist</param>
        /// <returns>int array</returns>
        protected uint[] GetOutlookPropTags(int data, uint[] properties, bool bCreate)
        {
            Guid guidOutlook = new Guid(data, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
            List<MAPINAMEID> names = new List<MAPINAMEID>();
            int int32Size = Marshal.SizeOf(typeof(Int32));
            foreach (int p in properties)
            {
                MAPINAMEID nameID = new MAPINAMEID();
                IntPtr pGuid = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(Guid)));
                Marshal.StructureToPtr(guidOutlook, pGuid, true);
                nameID.pGuid = pGuid;
                nameID.ulKind = 0;
                nameID.Kind.lID = p;
                names.Add(nameID);
            }
            List<uint> tags = new List<uint>();
            IntPtr pTags = IntPtr.Zero;
            try
            {
                IntPtr pNames = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(MAPINAMEID)) * names.Count);
                for (int i = 0; i < names.Count; i++)
                    Marshal.StructureToPtr(names[i], pNames + i * Marshal.SizeOf(typeof(MAPINAMEID)), true);
                HRESULT hResult = mapiObj_.GetIDsFromNames(1, pNames, bCreate ? (uint)OpenPropertyMode.CREATE : 0, out pTags);
                if (hResult == HRESULT.S_OK)
                {
                    int count = Marshal.ReadInt32(pTags);
                    for (int i = 0; i < count; i++)
                    {
                        tags.Add((uint)Marshal.ReadInt32(pTags + (i + 1) * int32Size));
                    }
                }
                Marshal.FreeHGlobal(pNames);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex);
            }
            foreach (MAPINAMEID n in names)
            {
                if (n.pGuid != null)
                    Marshal.FreeHGlobal(n.pGuid);
            }
            return tags.ToArray();
        }

        /// <summary>
        /// Get one outlook property tag
        /// </summary>
        /// <param name="data">data</param>
        /// <param name="property">property</param>
        /// <param name="bCreate">create tag if not exists</param>
        /// <returns>create tag if not exist</returns>
        /// <exception cref="Exception"></exception>
        protected uint GetOutlookPropTag(int data, uint property, bool bCreate)
        {
            Guid guidOutlook = new Guid(data, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
            MAPINAMEID nameID = new MAPINAMEID(); ;
            IntPtr pGuid = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(Guid)));
            Marshal.StructureToPtr(guidOutlook, pGuid, true);
            nameID.pGuid = pGuid;
            nameID.ulKind = (int)MAPINameIDKind.MNID_ID;
            nameID.Kind.lID = (int)property;
            IntPtr pNames = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(MAPINAMEID)));
            Marshal.StructureToPtr(nameID, pNames, true);
            IntPtr pTags = IntPtr.Zero;
            int int32Size = Marshal.SizeOf(typeof(Int32));
            HRESULT hr = mapiObj_.GetIDsFromNames(1, pNames, bCreate ? (uint)OpenPropertyMode.CREATE : 0, out pTags);
            Marshal.FreeHGlobal(nameID.pGuid);
            Marshal.FreeHGlobal(pNames);
            if (hr == HRESULT.S_OK)
            {
                int count = Marshal.ReadInt32(pTags);
                if (count > 0)
                {
                    return (uint)Marshal.ReadInt32(pTags + int32Size);
                }
            }
            throw new Exception("The outlook tag is not existed");
        }
        #endregion

        #region IDisposable Interface

        /// <summary>
        /// Dispose MAPI object.
        /// </summary>
        public virtual void Dispose()
        {
            if (mapiObj_ != null)
            {
                Marshal.ReleaseComObject(mapiObj_);
                mapiObj_ = null;
            }
        }

        #endregion
    }
}
