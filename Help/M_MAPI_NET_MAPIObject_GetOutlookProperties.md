# GetOutlookProperties Method


Get outlook properties



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public Dictionary<uint, IPropValue> GetOutlookProperties(
	int data,
	uint[] tags
)
```
**VB**
``` VB
Public Function GetOutlookProperties ( 
	data As Integer,
	tags As UInteger()
) As Dictionary(Of UInteger, IPropValue)
```
**C++**
``` C++
public:
Dictionary<unsigned int, IPropValue^>^ GetOutlookProperties(
	int data, 
	array<unsigned int>^ tags
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.int32" target="_blank" rel="noopener noreferrer">Int32</a></dt><dd>data</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a>[]</dt><dd>array of tags</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.collections.generic.dictionary-2" target="_blank" rel="noopener noreferrer">Dictionary</a>(<a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a>, <a href="T_MAPI_NET_IPropValue.md">IPropValue</a>)  
dictionary of properties

## See Also


#### Reference
<a href="T_MAPI_NET_MAPIObject.md">MAPIObject Class</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
