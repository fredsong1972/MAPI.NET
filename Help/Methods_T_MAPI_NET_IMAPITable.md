# IMAPITable Methods




## Methods
<table>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_Abort.md">Abort</a></td>
<td>Stops any asynchronous operations currently in progress for the table.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_Advise.md">Advise</a></td>
<td>Registers an advise sink object to receive notification of specified events affecting the table.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_CollapseRow.md">CollapseRow</a></td>
<td>Collapses an expanded table category, removing any lower-level headings and leaf rows belonging to the category from the table view.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_CreateBookmark.md">CreateBookmark</a></td>
<td>Creates a bookmark at the table's current position.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_ExpandRow.md">ExpandRow</a></td>
<td>Expands a collapsed table category, adding the leaf or lower-level heading rows belonging to the category to the table view.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_FindRow.md">FindRow</a></td>
<td>Finds the next row in a table that matches specific search criteria and moves the cursor to that row.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_FreeBookmark.md">FreeBookmark</a></td>
<td>Releases the memory associated with a bookmark.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_GetCollapseState.md">GetCollapseState</a></td>
<td>Returns the data that is needed to rebuild the current collapsed or expanded state of a categorized table.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_GetLastError.md">GetLastError</a></td>
<td>Returns a MAPIERROR structure containing information about the previous error on the table.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_GetRowCount.md">GetRowCount</a></td>
<td>Returns the total number of rows in the table.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_GetStatus.md">GetStatus</a></td>
<td>Returns the table's status and type.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_QueryColumns.md">QueryColumns</a></td>
<td>Returns a list of columns for the table.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_QueryPosition.md">QueryPosition</a></td>
<td>Retrieves the current table row position of the cursor, based on a fractional value.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_QueryRows.md">QueryRows</a></td>
<td>Returns one or more rows from a table, beginning at the current cursor position.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_QuerySortOrder.md">QuerySortOrder</a></td>
<td>Retrieves the current sort order for a table.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_Restrict.md">Restrict</a></td>
<td>Applies a filter to a table, reducing the row set to only those rows matching the specified criteria.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_SeekRow.md">SeekRow</a></td>
<td>Moves the cursor to a specific position in the table.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_SeekRowApprox.md">SeekRowApprox</a></td>
<td>Moves the cursor to an approximate fractional position in the table.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_SetCollapseState.md">SetCollapseState</a></td>
<td>Rebuilds the current expanded or collapsed state of a categorized table using data that was saved by a prior call to the IMAPITable::GetCollapseState method.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_SetColumns.md">SetColumns</a></td>
<td>Defines the particular properties and order of properties to appear as columns in the table.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_SortTable.md">SortTable</a></td>
<td>Orders the rows of the table, depending on sort criteria.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_Unadvise.md">Unadvise</a></td>
<td>Cancels the sending of notifications previously set up with a call to the IMAPITable::Advise method.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPITable_WaitForCompletion.md">WaitForCompletion</a></td>
<td>Suspends processing until one or more asynchronous operations in progress on the table have completed.</td></tr>
</table>

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPITable.md">IMAPITable Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
