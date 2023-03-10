# TableSortOrder Enumeration


Defines how to sort the rows of a table, what column to use as the sort key, and the direction of the sort.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public enum TableSortOrder
```
**VB**
``` VB
Public Enumeration TableSortOrder
```
**C++**
``` C++
public enum class TableSortOrder
```



## Members
<table>
<tr>
<td>TABLE_SORT_ASCEND</td>
<td>0</td>
<td>The table should be sorted in ascending order.</td></tr>
<tr>
<td>TABLE_SORT_DESCEND</td>
<td>1</td>
<td>The table should be sorted in descending order.</td></tr>
<tr>
<td>TABLE_SORT_COMBINE</td>
<td>2</td>
<td>The sort operation should create a category that combines the property identified as the sort key column in the ulPropTag member with the sort key column specified in the previous SSortOrder structure.</td></tr>
</table>

## See Also


#### Reference
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
