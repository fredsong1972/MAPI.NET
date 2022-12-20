# MAPIObject Class


IMAPIProp .Net Wrapper object



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public class MAPIObject : IDisposable
```
**VB**
``` VB
Public Class MAPIObject
	Implements IDisposable
```
**C++**
``` C++
public ref class MAPIObject : IDisposable
```

<table><tr><td><strong>Inheritance</strong></td><td><a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>  →  MAPIObject</td></tr>
<tr><td><strong>Derived</strong></td><td><a href="T_MAPI_NET_Attachment.md">MAPI.NET.Attachment</a><br /><a href="T_MAPI_NET_MAPIFolder.md">MAPI.NET.MAPIFolder</a><br /><a href="T_MAPI_NET_MAPIMessage.md">MAPI.NET.MAPIMessage</a><br /><a href="T_MAPI_NET_MAPIStatus.md">MAPI.NET.MAPIStatus</a><br /><a href="T_MAPI_NET_MsgStore.md">MAPI.NET.MsgStore</a></td></tr>
<tr><td><strong>Implements</strong></td><td><a href="https://learn.microsoft.com/dotnet/api/system.idisposable" target="_blank" rel="noopener noreferrer">IDisposable</a></td></tr>
</table>



## Constructors
<table>
<tr>
<td><a href="M_MAPI_NET_MAPIObject__ctor.md">MAPIObject(IMAPIProp)</a></td>
<td>Initializes a new instance of the MAPIObject class.</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject__ctor_1.md">MAPIObject(IMAPIProp, EntryID)</a></td>
<td>Initializes a new instance of the MAPIObject class.</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject__ctor_2.md">MAPIObject(MAPISession, EntryID)</a></td>
<td>Initializes a new instance of the MAPIObject class.</td></tr>
</table>

## Properties
<table>
<tr>
<td><a href="P_MAPI_NET_MAPIObject_EntryID.md">EntryID</a></td>
<td>EntryID of MAPIObject</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPIObject_InterfaceID.md">InterfaceID</a></td>
<td>IMAPIProp Interface Guid</td></tr>
</table>

## Methods
<table>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_Close.md">Close</a></td>
<td>Makes permanent any changes that were made to an object since the last save operation,and close the object.</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_CopyTo.md">CopyTo(UInt32[], IMAPIProp)</a></td>
<td>Copies or moves all properties, except for specifically excluded properties.</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_CopyTo_1.md">CopyTo(UInt32[], MAPIObject)</a></td>
<td>Copies or moves all properties, except for specifically excluded properties.</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_Dispose.md">Dispose</a></td>
<td>Dispose MAPI object.</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.equals#system-object-equals(system-object)" target="_blank" rel="noopener noreferrer">Equals</a></td>
<td>Determines whether the specified object is equal to the current object.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.finalize#system-object-finalize" target="_blank" rel="noopener noreferrer">Finalize</a></td>
<td>Allows an object to try to free resources and perform other cleanup operations before it is reclaimed by garbage collection.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.gethashcode#system-object-gethashcode" target="_blank" rel="noopener noreferrer">GetHashCode</a></td>
<td>Serves as the default hash function.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_GetOutlookProperties.md">GetOutlookProperties</a></td>
<td>Get outlook properties</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_GetOutlookProperty.md">GetOutlookProperty</a></td>
<td>Get outlook properties</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_GetOutlookPropTag.md">GetOutlookPropTag</a></td>
<td>Get one outlook property tag</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_GetOutlookPropTags.md">GetOutlookPropTags</a></td>
<td>Get outlook property tags</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_GetProperties.md">GetProperties(PropTags[])</a></td>
<td>Retrieves the property value of one or more properties of an object.</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_GetProperties_1.md">GetProperties(UInt32[])</a></td>
<td>Retrieves the property value of one or more properties of an object.</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_GetProperty.md">GetProperty(PropTags)</a></td>
<td>Retrieves one property value of an object.</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_GetProperty_1.md">GetProperty(UInt32)</a></td>
<td>Retrieves one property value of an object.</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.gettype#system-object-gettype" target="_blank" rel="noopener noreferrer">GetType</a></td>
<td>Gets the <a href="https://learn.microsoft.com/dotnet/api/system.type" target="_blank" rel="noopener noreferrer">Type</a> of the current instance.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.memberwiseclone#system-object-memberwiseclone" target="_blank" rel="noopener noreferrer">MemberwiseClone</a></td>
<td>Creates a shallow copy of the current <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_OpenProperty.md">OpenProperty(UInt32, Guid, OpenPropertyMode)</a></td>
<td>Returns a pointer to an interface that can be used to access a property.</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_OpenProperty_1.md">OpenProperty(UInt32, Guid, UInt32, OpenPropertyMode)</a></td>
<td>Returns a pointer to an interface that can be used to access a property.</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_Save.md">Save()</a></td>
<td>Makes permanent any changes that were made to an object since the last save operation,and keep the object open.</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_Save_1.md">Save(SaveFlags)</a></td>
<td>Makes permanent any changes that were made to an object since the last save operation,and control the object per the flag.</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_SetPropertyValue.md">SetPropertyValue(IPropValue)</a></td>
<td>Updates one property of the object.</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_SetPropertyValue_1.md">SetPropertyValue(PropTags, Object)</a></td>
<td>Updates one property of the object.</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_SetPropertyValue_2.md">SetPropertyValue(UInt32, Object)</a></td>
<td>Updates one property of the object.</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_SetPropertyValues.md">SetPropertyValues</a></td>
<td>Updates one or more properties of the object.</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.tostring#system-object-tostring" target="_blank" rel="noopener noreferrer">ToString</a></td>
<td>Returns a string that represents the current object.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
</table>

## Fields
<table>
<tr>
<td><a href="F_MAPI_NET_MAPIObject_Id_.md">Id_</a></td>
<td>EntryID of the object</td></tr>
</table>

## See Also


#### Reference
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
