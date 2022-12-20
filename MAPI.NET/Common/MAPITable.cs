using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace MAPI.NET
{
    /// <summary>
    /// Provides a read-only view of a table. IMAPITable is used by clients and service providers to manipulate the way a table appears. 
    /// </summary>
    [Guid("00020301-0000-0000-c000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMAPITable
    {
        /// <summary>
        /// Returns a MAPIERROR structure containing information about the previous error on the table.
        /// </summary>
        /// <param name="hResult">HRESULT containing the error generated in the previous method call.</param>
        /// <param name="ulFlags">Bitmask of flags that controls the type of the returned strings. </param>
        /// <param name="lppMAPIError">Pointer to a pointer to the returned MAPIERROR structure containing version, component, and context information for the error.</param>
        /// <returns>S_OK, if the call succeeded and has returned the expected value or values; otherwise, failed.</returns>
        HRESULT GetLastError(int hResult, uint ulFlags, out IntPtr lppMAPIError);
        /// <summary>
        /// Registers an advise sink object to receive notification of specified events affecting the table.
        /// </summary>
        /// <param name="ulEventMask">Value indicating the type of event that will generate the notification.</param>
        /// <param name="lpAdviseSink">Pointer to an advise sink object to receive the subsequent notifications. This advise sink object must have been already allocated.</param>
        /// <param name="lpulConnection">Pointer to a nonzero value that represents the successful notification registration.</param>
        /// <returns>S_OK, if the notification registration successfully completed; otherwise, failed.</returns>
        HRESULT Advise(uint ulEventMask, IntPtr lpAdviseSink, IntPtr lpulConnection);
        /// <summary>
        /// Cancels the sending of notifications previously set up with a call to the IMAPITable::Advise method.
        /// </summary>
        /// <param name="ulConnection">The number of the registration connection returned by a call to IMAPITable::Advise.</param>
        /// <returns>S_OK, if the call succeeded; otherwise, failed.</returns>
        HRESULT Unadvise(uint ulConnection);
        /// <summary>
        /// Returns the table's status and type.
        /// </summary>
        /// <param name="lpulTableStatus">Pointer to a value indicating the status of the table.</param>
        /// <param name="lpulTableType">Pointer to a value that indicates the table's type.</param>
        /// <returns>S_OK, if the table's status was successfully returned; otherwise, failed.</returns>
        HRESULT GetStatus(IntPtr lpulTableStatus, IntPtr lpulTableType);
        /// <summary>
        /// Defines the particular properties and order of properties to appear as columns in the table.
        /// </summary>
        /// <param name="lpPropTagArray">Pointer to an array of property tags identifying properties to be included as columns in the table. </param>
        /// <param name="ulFlags">Bitmask of flags that controls the return of an asynchronous call to SetColumns.</param>
        /// <returns>S_OK, if the column setting operation was successful; otherwise, failed.</returns>
        HRESULT SetColumns([MarshalAs(UnmanagedType.LPArray)] uint[] lpPropTagArray, uint ulFlags);
        /// <summary>
        /// Returns a list of columns for the table.
        /// </summary>
        /// <param name="ulFlags">Bitmask of flags that indicates which column set should be returned.</param>
        /// <param name="lpPropTagArray">Pointer to an SPropTagArray structure containing the property tags for the column set.</param>
        /// <returns>S_OK, if the column set was successfully returned; otherwise, failed.</returns>
        HRESULT QueryColumns(uint ulFlags, IntPtr lpPropTagArray);
        /// <summary>
        /// Returns the total number of rows in the table. 
        /// </summary>
        /// <param name="ulFlags">Reserved; must be zero.</param>
        /// <param name="lpulCount">Pointer to the number of rows in the table.</param>
        /// <returns>S_OK, if the row count was successfully returned; otherwise, failed.</returns>
        HRESULT GetRowCount(uint ulFlags, out uint lpulCount);
        /// <summary>
        /// Moves the cursor to a specific position in the table.
        /// </summary>
        /// <param name="bkOrigin">The bookmark identifying the starting position for the seek operation.</param>
        /// <param name="lRowCount">The signed count of the number of rows to move, starting from the bookmark identified by the bkOrigin parameter.</param>
        /// <param name="lplRowsSought">If lRowCount is a valid pointer on input, lplRowsSought points to the number of rows that were processed in the seek operation, the sign of which indicates the direction of search, forward or backward. If lRowCount is negative, then lplRowsSought is negative.</param>
        /// <returns>S_OK, if the seek operation was successful; otherwise, failed.</returns>
        HRESULT SeekRow(int bkOrigin, int lRowCount, out IntPtr lplRowsSought);
        /// <summary>
        /// Moves the cursor to an approximate fractional position in the table. 
        /// </summary>
        /// <param name="ulNumerator">The numerator of the fraction representing the table position</param>
        /// <param name="ulDenominator">The denominator of the fraction representing the table position</param>
        /// <returns>S_OK, if the seek operation was successful; otherwise, failed.</returns>
        HRESULT SeekRowApprox(uint ulNumerator, uint ulDenominator);
        /// <summary>
        /// Retrieves the current table row position of the cursor, based on a fractional value.
        /// </summary>
        /// <param name="lpulRow">Pointer to the number of the current row.</param>
        /// <param name="lpulNumerator">Pointer to the numerator for the fraction identifying the table position.</param>
        /// <param name="lpulDenominator">Pointer to the denominator for the fraction identifying the table position.</param>
        /// <returns>S_OK, if the method returned valid values in lpulRow, lpulNumerator, and lpulDenominator; otherwise, failed.</returns>
        HRESULT QueryPosition(IntPtr lpulRow, IntPtr lpulNumerator, IntPtr lpulDenominator);
        /// <summary>
        /// Finds the next row in a table that matches specific search criteria and moves the cursor to that row.
        /// </summary>
        /// <param name="lpRestriction">A pointer to an SRestriction structure that describes the search criteria.</param>
        /// <param name="BkOrigin">A bookmark identifying the row where FindRow should begin its search.</param>
        /// <param name="ulFlags">A bitmask of flags that controls the direction of the search.</param>
        /// <returns>S_OK, if the find operation was successful; otherwise, failed.</returns>
        HRESULT FindRow(out IntPtr lpRestriction, uint BkOrigin, uint ulFlags);
        /// <summary>
        /// Applies a filter to a table, reducing the row set to only those rows matching the specified criteria.
        /// </summary>
        /// <param name="lpRestriction">Pointer to an SRestriction structure defining the conditions of the filter. Passing NULL in the lpRestriction parameter removes the current filter.</param>
        /// <param name="ulFlags">Bitmask of flags that controls the timing of the restriction operation.</param>
        /// <returns>S_OK, if the filter was successfully applied; otherwise, failed.</returns>
        HRESULT Restrict(out IntPtr lpRestriction, uint ulFlags);
        /// <summary>
        /// Creates a bookmark at the table's current position.
        /// </summary>
        /// <param name="lpbkPosition">Pointer to the returned 32-bit bookmark value. This bookmark can later be passed in a call to the IMAPITable::SeekRow method</param>
        /// <returns>S_OK, if the call succeeded and has returned the expected value or values; otherwise, failed.</returns>
        HRESULT CreateBookmark(IntPtr lpbkPosition);
        /// <summary>
        /// Releases the memory associated with a bookmark.
        /// </summary>
        /// <param name="bkPosition">The bookmark to be freed, created by calling the IMAPITable::CreateBookmark method.</param>
        /// <returns>S_OK, if the bookmark was successfully freed; otherwise, failed.</returns>
        HRESULT FreeBookmark(IntPtr bkPosition);
        /// <summary>
        /// Orders the rows of the table, depending on sort criteria.
        /// </summary>
        /// <param name="lpSortCriteria">Pointer to an SSortOrderSet structure that contains the sort criteria to apply.</param>
        /// <param name="ulFlags">Bitmask of flags that controls the timing of the IMAPITable::SortTable operation.</param>
        /// <returns>S_OK, if the sort operation was successful; otherwise, failed.</returns>
        HRESULT SortTable(IntPtr lpSortCriteria, int ulFlags);
        /// <summary>
        /// Retrieves the current sort order for a table.
        /// </summary>
        /// <param name="lppSortCriteria">Pointer to a pointer to the SSortOrderSet structure holding the current sort order.</param>
        /// <returns>S_OK, if the current sort order was successfully returned; otherwise, failed.</returns>
        HRESULT QuerySortOrder(IntPtr lppSortCriteria);
        /// <summary>
        /// Returns one or more rows from a table, beginning at the current cursor position.
        /// </summary>
        /// <param name="lRowCount">Maximum number of rows to be returned.</param>
        /// <param name="ulFlags">Bitmask of flags that control how rows are returned.</param>
        /// <param name="lppRows">Pointer to a pointer to an SRowSet structure holding the table rows.</param>
        /// <returns>S_OK, if the rows were successfully returned; otherwise, failed.</returns>
        HRESULT QueryRows(int lRowCount, uint ulFlags, out IntPtr lppRows);
        /// <summary>
        /// Stops any asynchronous operations currently in progress for the table.
        /// </summary>
        /// <returns>S_OK, if one or more asynchronous operations have been stopped; otherwise, failed.</returns>
        HRESULT Abort();
        /// <summary>
        /// Expands a collapsed table category, adding the leaf or lower-level heading rows belonging to the category to the table view.
        /// </summary>
        /// <param name="cbInstanceKey">The count of bytes in the PR_INSTANCE_KEY property pointed to by the pbInstanceKey parameter.</param>
        /// <param name="pbInstanceKey">A pointer to the PR_INSTANCE_KEY property that identifies the heading row for the category.</param>
        /// <param name="ulRowCount">The maximum number of rows to return in the lppRows parameter. </param>
        /// <param name="ulFlags">Reserved; must be zero.</param>
        /// <param name="lppRows">A pointer to an SRowSet structure receiving the first (up to ulRowCount) rows that have been inserted into the table view as a result of the expansion.</param>
        /// <param name="lpulMoreRows">A pointer to the total number of rows that were added to the table view.</param>
        /// <returns>S_OK, if the category was expanded successfully; otherwise, failed.</returns>
        HRESULT ExpandRow(uint cbInstanceKey, IntPtr pbInstanceKey, uint ulRowCount, uint ulFlags, IntPtr lppRows, IntPtr lpulMoreRows);
        /// <summary>
        /// Collapses an expanded table category, removing any lower-level headings and leaf rows belonging to the category from the table view.
        /// </summary>
        /// <param name="cbInstanceKey">The count of bytes in the PR_INSTANCE_KEY property pointed to by the pbInstanceKey parameter.</param>
        /// <param name="pbInstanceKey">A pointer to the PR_INSTANCE_KEY property that identifies the heading row for the category. </param>
        /// <param name="ulFlags">Reserved; must be zero.</param>
        /// <param name="lpulRowCount">A pointer to the total number of rows that are being removed from the table view.</param>
        /// <returns>S_OK, if the collapse operation has succeeded; otherwise, failed.</returns>
        HRESULT CollapseRow(uint cbInstanceKey, IntPtr pbInstanceKey, uint ulFlags, IntPtr lpulRowCount);
        /// <summary>
        /// Suspends processing until one or more asynchronous operations in progress on the table have completed.
        /// </summary>
        /// <param name="ulFlags">Reserved; must be zero.</param>
        /// <param name="ulTimeout">Maximum number of milliseconds to wait for the asynchronous operation or operations to complete.</param>
        /// <param name="lpulTableStatus">On input, either a valid pointer or NULL. On output, if lpulTableStatus is a valid pointer, it points to the most recent status of the table. </param>
        /// <returns>S_OK, if the wait operation was successful; otherwise, failed.</returns>
        HRESULT WaitForCompletion(uint ulFlags, uint ulTimeout, IntPtr lpulTableStatus);
        /// <summary>
        /// Returns the data that is needed to rebuild the current collapsed or expanded state of a categorized table.
        /// </summary>
        /// <param name="ulFlags">Reserved; must be zero.</param>
        /// <param name="cbInstanceKey">The count of bytes in the instance key pointed to by the lpbInstanceKey parameter.</param>
        /// <param name="lpbInstanceKey">A pointer to the PR_INSTANCE_KEY property of the row at which the current collapsed or expanded state should be rebuilt. </param>
        /// <param name="lpcbCollapseState">A pointer to the count of structures pointed to by the lppbCollapseState parameter.</param>
        /// <param name="lppbCollapseState">A pointer to a pointer to structures that contain data that describes the current table view.</param>
        /// <returns>S_OK, if the state for the categorized table was successfully saved; otherwise, failed.</returns>
        HRESULT GetCollapseState(uint ulFlags, uint cbInstanceKey, IntPtr lpbInstanceKey, IntPtr lpcbCollapseState, IntPtr lppbCollapseState);
        /// <summary>
        /// Rebuilds the current expanded or collapsed state of a categorized table using data that was saved by a prior call to the IMAPITable::GetCollapseState method.
        /// </summary>
        /// <param name="ulFlags">Reserved; must be zero.</param>
        /// <param name="cbCollapseState">Count of bytes in the structure pointed to by the pbCollapseState parameter.</param>
        /// <param name="pbCollapseState">Pointer to the structures containing the data needed to rebuild the table view.</param>
        /// <param name="lpbkLocation">Pointer to a bookmark identifying the row in the table at which the collapsed or expanded state should be rebuilt. </param>
        /// <returns>S_OK, if the state of the categorized table was successfully rebuilt; otherwise, failed.</returns>
        HRESULT SetCollapseState(uint ulFlags, uint cbCollapseState, IntPtr pbCollapseState, IntPtr lpbkLocation);

    }

    /// <summary>
    /// .Net wrapper over IMAPITable interface.
    /// </summary>
    public class MAPITable : IDisposable
    {
        Guid IID_IMAPITable = new Guid("00020301-0000-0000-c000-000000000046");

        private IMAPITable tb_;

        /// <summary>
        /// Initializes a new instance of the MAPITable class. 
        /// </summary>
        /// <param name="mapiTable">mapi table</param>
        public MAPITable(IMAPITable mapiTable)
        {
            tb_ = mapiTable;
        }
        /// <summary>
        /// Defines the particular properties and order of properties to appear as columns in the table.
        /// </summary>
        /// <param name="tags">An array of property tags identifying properties to be included as columns in the table</param>
        /// <returns>true, if successful; otherwise, failed.</returns>
        public bool SetColumns(PropTags[] tags)
        {
            uint[] t = new uint[tags.Length + 1];
            t[0] = (uint)tags.Length;
            for (int i = 0; i < tags.Length; i++)
                t[i + 1] = (uint)tags[i];
            return tb_.SetColumns(t, 0) == HRESULT.S_OK;
        }
        /// <summary>
        /// Gets the total number of rows in the table. 
        /// </summary>
        /// <returns>the total number of rows in the table</returns>
        public uint GetRowCount()
        {
            uint count = 0;
            HRESULT hr = tb_.GetRowCount(0, out count);
            if (hr == HRESULT.S_OK)
                return count;
            return 0;
        }

        /// <summary>
        /// Returns one or more rows from a table, beginning at the current cursor position.
        /// </summary>
        /// <param name="lRowCount">Maximum number of rows to be returned.</param>
        /// <param name="sRows">an SRow array holding the table rows.</param>
        /// <returns>true, if successful; otherwise, failed.</returns>
        public bool QueryRows(int lRowCount, out SRow[] sRows)
        {
            IntPtr pRowSet = IntPtr.Zero;
            HRESULT hr = tb_.QueryRows(lRowCount, 0, out pRowSet);
            if (hr != HRESULT.S_OK)
            {
                MAPINative.MAPIFreeBuffer(pRowSet);
            }

            uint cRows = (uint)Marshal.ReadInt32(pRowSet);
            sRows = new SRow[cRows];

            if (cRows < 1)
            {
                MAPINative.MAPIFreeBuffer(pRowSet);
                return false;
            }

            int pIntSize = IntPtr.Size, intSize = Marshal.SizeOf(typeof(Int32));
            int sizeOfSRow = 2 * intSize + pIntSize;
            IntPtr rows = pRowSet + intSize;
            for (int i = 0; i < cRows; i++)
            {
                IntPtr pRowOffset = rows + i * sizeOfSRow;
                uint cValues = (uint)Marshal.ReadInt32(pRowOffset + pIntSize);
                IntPtr pProps = Marshal.ReadIntPtr(pRowOffset + pIntSize + intSize);

                IPropValue[] lpProps = new IPropValue[cValues];
                for (int j = 0; j < cValues; j++) // each column
                {
                    pSPropValue lpProp = (pSPropValue)Marshal.PtrToStructure(pProps + j * Marshal.SizeOf(typeof(pSPropValue)), typeof(pSPropValue));
                    lpProps[j] = new MAPIProp(lpProp);
                }
                sRows[i].propVals = lpProps;
            }
            MAPINative.MAPIFreeBuffer(pRowSet);
            return true;
        }
        /// <summary>
        /// Orders the rows of the table, depending on sort criteria.
        /// </summary>
        /// <param name="tag">Property tag identifying the sort key.</param>
        /// <param name="order">The order in which the data is to be sorted.</param>
        /// <returns>true, if the sort operation was successful; otherwise, failed.</returns>
        public bool SortRows(PropTags tag, TableSortOrder order)
        {
            int sizeS = Marshal.SizeOf(typeof(SSortOrder));
            IntPtr sortArray = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(uint)) * 3 + sizeS);
            Marshal.WriteInt32(sortArray, 1);
            Marshal.WriteInt64(sortArray, Marshal.SizeOf(typeof(uint)), 0);
            SSortOrder s;
            s.ulOrder = order;
            s.ulPropTag = tag;
            Marshal.StructureToPtr(s, sortArray + Marshal.SizeOf(typeof(uint)) * 3, false);
            HRESULT hr = tb_.SortTable(sortArray, 0);
            Marshal.FreeHGlobal(sortArray);
            return hr == HRESULT.S_OK;
        }
        /// <summary>
        /// Moves the cursor to a specific position in the table.
        /// </summary>
        /// <param name="bookMark">The bookmark identifying the starting position for the seek operation.</param>
        /// <param name="rowCount">The signed count of the number of rows to move, starting from the bookmark.</param>
        /// <returns>true, if successful; otherwise, failed</returns>
        public bool SeekRow(BookMark bookMark, int rowCount)
        {
            IntPtr pRowsSought;
            HRESULT hResult = tb_.SeekRow((int)bookMark, rowCount, out pRowsSought);
            return hResult == HRESULT.S_OK;
        }

        #region IDisposable Interface
        /// <summary>
        /// Dispose MAPITable object.
        /// </summary>
        public void Dispose()
        {
            if (tb_ != null)
            {
                Marshal.ReleaseComObject(tb_);
                tb_ = null;
            }
        }

        #endregion
    }
}
