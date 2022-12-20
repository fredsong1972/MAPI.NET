# SetColumns Method


Defines the particular properties and order of properties to appear as columns in the table.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public bool SetColumns(
	PropTags[] tags
)
```
**VB**
``` VB
Public Function SetColumns ( 
	tags As PropTags()
) As Boolean
```
**C++**
``` C++
public:
bool SetColumns(
	array<PropTags>^ tags
)
```



#### Parameters
<dl><dt>  <a href="T_MAPI_NET_PropTags.md">PropTags</a>[]</dt><dd>An array of property tags identifying properties to be included as columns in the table</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a>  
true, if successful; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_MAPITable.md">MAPITable Class</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
