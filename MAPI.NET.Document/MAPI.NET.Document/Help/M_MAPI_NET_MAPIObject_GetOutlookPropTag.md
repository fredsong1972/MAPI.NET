# GetOutlookPropTag Method


Get one outlook property tag



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
protected uint GetOutlookPropTag(
	int data,
	uint property,
	bool bCreate
)
```
**VB**
``` VB
Protected Function GetOutlookPropTag ( 
	data As Integer,
	property As UInteger,
	bCreate As Boolean
) As UInteger
```
**C++**
``` C++
protected:
unsigned int GetOutlookPropTag(
	int data, 
	unsigned int property, 
	bool bCreate
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.int32" target="_blank" rel="noopener noreferrer">Int32</a></dt><dd>data</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>property</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a></dt><dd>create tag if not exists</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a>  
create tag if not exist

## Exceptions
<table>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.exception" target="_blank" rel="noopener noreferrer">Exception</a></td>
<td /></tr>
</table>

## See Also


#### Reference
<a href="T_MAPI_NET_MAPIObject.md">MAPIObject Class</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
