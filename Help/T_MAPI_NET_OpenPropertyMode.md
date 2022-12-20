# OpenPropertyMode Enumeration


A bitmask of flags that controls access to the property.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public enum OpenPropertyMode
```
**VB**
``` VB
Public Enumeration OpenPropertyMode
```
**C++**
``` C++
public enum class OpenPropertyMode
```



## Members
<table>
<tr>
<td>READ</td>
<td>0</td>
<td>Read</td></tr>
<tr>
<td>MODIFY</td>
<td>1</td>
<td>Modify</td></tr>
<tr>
<td>CREATE</td>
<td>2</td>
<td>If the property does not exist, it should be created. If the property does exist, the current value of the property should be discarded. When a caller sets the MAPI_CREATE flag, it should also set the MAPI_MODIFY flag.</td></tr>
<tr>
<td>APPEND</td>
<td>4</td>
<td>Append</td></tr>
</table>

## See Also


#### Reference
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
