# SeekRow Method


Moves the cursor to a specific position in the table.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public bool SeekRow(
	BookMark bookMark,
	int rowCount
)
```
**VB**
``` VB
Public Function SeekRow ( 
	bookMark As BookMark,
	rowCount As Integer
) As Boolean
```
**C++**
``` C++
public:
bool SeekRow(
	BookMark bookMark, 
	int rowCount
)
```



#### Parameters
<dl><dt>  <a href="T_MAPI_NET_BookMark.md">BookMark</a></dt><dd>The bookmark identifying the starting position for the seek operation.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.int32" target="_blank" rel="noopener noreferrer">Int32</a></dt><dd>The signed count of the number of rows to move, starting from the bookmark.</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a>  
true, if successful; otherwise, failed

## See Also


#### Reference
<a href="T_MAPI_NET_MAPITable.md">MAPITable Class</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
