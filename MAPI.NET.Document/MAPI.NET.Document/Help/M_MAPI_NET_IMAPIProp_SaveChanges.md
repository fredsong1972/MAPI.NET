# SaveChanges Method


Makes permanent any changes that were made to an object since the last save operation.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT SaveChanges(
	uint ulFlags
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function SaveChanges ( 
	ulFlags As UInteger
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT SaveChanges(
	unsigned int ulFlags
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls what happens to the object when the IMAPIProp::SaveChanges method is called.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK The commitment of changes was successful. MAPI_E_NO_ACCESS SaveChanges cannot keep the object open for read-only permission if KEEP_OPEN_READONLY is set, or read/write permission if KEEP_OPEN_READWRITE is set. No changes are committed. MAPI_E_OBJECT_CHANGED The object has changed since it was opened. MAPI_E_OBJECT_DELETED The object has been deleted since it was opened.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
