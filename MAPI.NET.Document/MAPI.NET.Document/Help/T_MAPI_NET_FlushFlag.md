# FlushFlag Enumeration


Bitmask of flags that controls the flush operation.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public enum FlushFlag
```
**VB**
``` VB
Public Enumeration FlushFlag
```
**C++**
``` C++
public enum class FlushFlag
```



## Members
<table>
<tr>
<td>UPLOAD</td>
<td>2</td>
<td>The outbound message queues should be flushed.</td></tr>
<tr>
<td>DOWNLOAD</td>
<td>4</td>
<td>The inbound message queues should be flushed.</td></tr>
<tr>
<td>FORCE</td>
<td>8</td>
<td>The flush operation should occur regardless, in spite of the possibility of performance degradation. This flag must be set when targeting an asynchronous transport provider.</td></tr>
<tr>
<td>NO_UI</td>
<td>16</td>
<td>The status object should not display a progress indicator. This flag is used only by the MAPI spooler; providers ignore this flag.</td></tr>
<tr>
<td>ASYNC_OK</td>
<td>32</td>
<td>The flush operation can occur asynchronously. This flag only applies to the MAPI spooler's status object.</td></tr>
</table>

## See Also


#### Reference
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
