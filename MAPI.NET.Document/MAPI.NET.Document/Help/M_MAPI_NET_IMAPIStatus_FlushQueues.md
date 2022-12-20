# FlushQueues Method


Forces all messages waiting to be sent or received to be immediately uploaded or downloaded.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT FlushQueues(
	IntPtr ulUIParam,
	uint cbTargetTransport,
	IntPtr lpTargetTransport,
	uint ulFlags
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function FlushQueues ( 
	ulUIParam As IntPtr,
	cbTargetTransport As UInteger,
	lpTargetTransport As IntPtr,
	ulFlags As UInteger
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT FlushQueues(
	IntPtr ulUIParam, 
	unsigned int cbTargetTransport, 
	IntPtr lpTargetTransport, 
	unsigned int ulFlags
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A handle to the parent window of any dialog boxes or windows that this method displays.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The byte count in the entry identifier pointed to by the lpTargetTransport parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the entry identifier of the transport provider that is to flush its message queues.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls the flush operation.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the flush operation was successful; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIStatus.md">IMAPIStatus Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
