# SetReadFlag Method


Sets or clears the MSGFLAG_READ flag in the PR_MESSAGE_FLAGS property of the message and manages the sending of read reports.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT SetReadFlag(
	uint ulFlags
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function SetReadFlag ( 
	ulFlags As UInteger
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT SetReadFlag(
	unsigned int ulFlags
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Bitmask of flags that controls the setting of a message's read flag that is, the message's MSGFLAG_READ flag in its PR_MESSAGE_FLAGS property and the processing of read reports.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the message's read flag has been successfully set or cleared; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMessage.md">IMessage Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
