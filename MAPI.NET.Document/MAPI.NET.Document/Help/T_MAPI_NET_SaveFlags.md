# SaveFlags Enumeration


A bitmask of flags that controls what happens to the object when the IMAPIProp::SaveChanges method is called.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public enum SaveFlags
```
**VB**
``` VB
Public Enumeration SaveFlags
```
**C++**
``` C++
public enum class SaveFlags
```



## Members
<table>
<tr>
<td>Close</td>
<td>0</td>
<td>Close the object</td></tr>
<tr>
<td>KEEP_OPEN_READONLY</td>
<td>1</td>
<td>Keep the object open as readonly</td></tr>
<tr>
<td>KEEP_OPEN_READWRITE</td>
<td>2</td>
<td>Keep the object open as readwrite</td></tr>
<tr>
<td>FORCE_SAVE</td>
<td>4</td>
<td>Force to save</td></tr>
</table>

## See Also


#### Reference
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
