# GetOutlookPropTags Method


Get outlook property tags



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
protected uint[] GetOutlookPropTags(
	int data,
	uint[] properties,
	bool bCreate
)
```
**VB**
``` VB
Protected Function GetOutlookPropTags ( 
	data As Integer,
	properties As UInteger(),
	bCreate As Boolean
) As UInteger()
```
**C++**
``` C++
protected:
array<unsigned int>^ GetOutlookPropTags(
	int data, 
	array<unsigned int>^ properties, 
	bool bCreate
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.int32" target="_blank" rel="noopener noreferrer">Int32</a></dt><dd>data</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a>[]</dt><dd>properties</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a></dt><dd>create tag if not exist</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a>[]  
int array

## See Also


#### Reference
<a href="T_MAPI_NET_MAPIObject.md">MAPIObject Class</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
