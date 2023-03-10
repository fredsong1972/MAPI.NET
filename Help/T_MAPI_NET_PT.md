# PT Enumeration


Property Type



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public enum PT
```
**VB**
``` VB
Public Enumeration PT
```
**C++**
``` C++
public enum class PT
```



## Members
<table>
<tr>
<td>PT_UNSPECIFIED</td>
<td>0</td>
<td>(Reserved for interface use) type doesn't matter to caller</td></tr>
<tr>
<td>PT_NULL</td>
<td>1</td>
<td>NULL property value</td></tr>
<tr>
<td>PT_I2</td>
<td>2</td>
<td>Signed 16-bit value</td></tr>
<tr>
<td>PT_LONG</td>
<td>3</td>
<td>Signed 32-bit value</td></tr>
<tr>
<td>PT_R4</td>
<td>4</td>
<td>4-byte floating point</td></tr>
<tr>
<td>PT_DOUBLE</td>
<td>5</td>
<td>Floating point double</td></tr>
<tr>
<td>PT_CURRENCY</td>
<td>6</td>
<td>Signed 64-bit int (decimal w/ 4 digits right of decimal pt)</td></tr>
<tr>
<td>PT_APPTIME</td>
<td>7</td>
<td>Application time</td></tr>
<tr>
<td>PT_ERROR</td>
<td>10</td>
<td>32-bit error value</td></tr>
<tr>
<td>PT_BOOLEAN</td>
<td>11</td>
<td>16-bit boolean (non-zero true,zero false)</td></tr>
<tr>
<td>PT_BOOLEAN_DESKTOP</td>
<td>11</td>
<td>16-bit boolean (non-zero true)</td></tr>
<tr>
<td>PT_OBJECT</td>
<td>13</td>
<td>Embedded object in a property</td></tr>
<tr>
<td>PT_I8</td>
<td>20</td>
<td>8-byte signed integer</td></tr>
<tr>
<td>PT_STRING8</td>
<td>30</td>
<td>Null terminated 8-bit character string</td></tr>
<tr>
<td>PT_UNICODE</td>
<td>31</td>
<td>Null terminated Unicode string</td></tr>
<tr>
<td>PT_TSTRING</td>
<td>31</td>
<td>IF MAPI Unicode, PT_TString is Unicode string; otheriwse 8-bit character string</td></tr>
<tr>
<td>PT_SYSTIME</td>
<td>64</td>
<td>FILETIME 64-bit int w/ number of 100ns periods since Jan 1,1601</td></tr>
<tr>
<td>PT_CLSID</td>
<td>72</td>
<td>OLE GUID</td></tr>
<tr>
<td>PT_BINARY</td>
<td>258</td>
<td>Uninterpreted (counted byte array)</td></tr>
</table>

## See Also


#### Reference
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
