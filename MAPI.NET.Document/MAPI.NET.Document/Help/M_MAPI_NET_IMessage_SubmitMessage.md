# SubmitMessage Method


Saves all of the message's properties and marks the message as ready to be sent.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT SubmitMessage(
	uint ulFlags
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function SubmitMessage ( 
	ulFlags As UInteger
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT SubmitMessage(
	unsigned int ulFlags
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Bitmask of flags used to control how a message is submitted</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the call succeeded and has returned the expected value or values; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMessage.md">IMessage Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
