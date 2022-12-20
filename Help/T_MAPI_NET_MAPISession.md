# MAPISession Class


IMAPISession .Net wrapper object



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public class MAPISession : Component
```
**VB**
``` VB
Public Class MAPISession
	Inherits Component
```
**C++**
``` C++
public ref class MAPISession : public Component
```

<table><tr><td><strong>Inheritance</strong></td><td><a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>  →  <a href="https://learn.microsoft.com/dotnet/api/system.marshalbyrefobject" target="_blank" rel="noopener noreferrer">MarshalByRefObject</a>  →  <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component" target="_blank" rel="noopener noreferrer">Component</a>  →  MAPISession</td></tr>
</table>



## Constructors
<table>
<tr>
<td><a href="M_MAPI_NET_MAPISession__ctor.md">MAPISession()</a></td>
<td>Initializes a new instance of the MAPISession class.</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPISession__ctor_1.md">MAPISession(IContainer)</a></td>
<td>Initializes a new instance of the MAPISession class.</td></tr>
</table>

## Properties
<table>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component.canraiseevents#system-componentmodel-component-canraiseevents" target="_blank" rel="noopener noreferrer">CanRaiseEvents</a></td>
<td>Gets a value indicating whether the component can raise an event.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component" target="_blank" rel="noopener noreferrer">Component</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component.container#system-componentmodel-component-container" target="_blank" rel="noopener noreferrer">Container</a></td>
<td>Gets the <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.icontainer" target="_blank" rel="noopener noreferrer">IContainer</a> that contains the <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component" target="_blank" rel="noopener noreferrer">Component</a>.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component" target="_blank" rel="noopener noreferrer">Component</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component.designmode#system-componentmodel-component-designmode" target="_blank" rel="noopener noreferrer">DesignMode</a></td>
<td>Gets a value that indicates whether the <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component" target="_blank" rel="noopener noreferrer">Component</a> is currently in design mode.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component" target="_blank" rel="noopener noreferrer">Component</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component.events#system-componentmodel-component-events" target="_blank" rel="noopener noreferrer">Events</a></td>
<td>Gets the list of event handlers that are attached to this <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component" target="_blank" rel="noopener noreferrer">Component</a>.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component" target="_blank" rel="noopener noreferrer">Component</a>)</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPISession_MsgStore.md">MsgStore</a></td>
<td>Gets current message store.</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPISession_PrimaryIdentity.md">PrimaryIdentity</a></td>
<td>Get primary identity</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPISession_ProfileEmailAddress.md">ProfileEmailAddress</a></td>
<td>Get profile email address</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component.site#system-componentmodel-component-site" target="_blank" rel="noopener noreferrer">Site</a></td>
<td>Gets or sets the <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.isite" target="_blank" rel="noopener noreferrer">ISite</a> of the <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component" target="_blank" rel="noopener noreferrer">Component</a>.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component" target="_blank" rel="noopener noreferrer">Component</a>)</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPISession_Spooler.md">Spooler</a></td>
<td>Gets current spooler.</td></tr>
</table>

## Methods
<table>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component.dispose#system-componentmodel-component-dispose" target="_blank" rel="noopener noreferrer">Dispose()</a></td>
<td>Releases all resources used by the <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component" target="_blank" rel="noopener noreferrer">Component</a>.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component" target="_blank" rel="noopener noreferrer">Component</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPISession_Dispose.md">Dispose(Boolean)</a></td>
<td>Clean up any resources being used.<br />(Overrides <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component.dispose#system-componentmodel-component-dispose(system-boolean)" target="_blank" rel="noopener noreferrer">Component.Dispose(Boolean)</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.equals#system-object-equals(system-object)" target="_blank" rel="noopener noreferrer">Equals</a></td>
<td>Determines whether the specified object is equal to the current object.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component.finalize#system-componentmodel-component-finalize" target="_blank" rel="noopener noreferrer">Finalize</a></td>
<td>Releases unmanaged resources and performs other cleanup operations before the <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component" target="_blank" rel="noopener noreferrer">Component</a> is reclaimed by garbage collection.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component" target="_blank" rel="noopener noreferrer">Component</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.gethashcode#system-object-gethashcode" target="_blank" rel="noopener noreferrer">GetHashCode</a></td>
<td>Serves as the default hash function.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.marshalbyrefobject.getlifetimeservice#system-marshalbyrefobject-getlifetimeservice" target="_blank" rel="noopener noreferrer">GetLifetimeService</a></td>
<td>Retrieves the current lifetime service object that controls the lifetime policy for this instance.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.marshalbyrefobject" target="_blank" rel="noopener noreferrer">MarshalByRefObject</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component.getservice#system-componentmodel-component-getservice(system-type)" target="_blank" rel="noopener noreferrer">GetService</a></td>
<td>Returns an object that represents a service provided by the <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component" target="_blank" rel="noopener noreferrer">Component</a> or by its <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.container" target="_blank" rel="noopener noreferrer">Container</a>.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component" target="_blank" rel="noopener noreferrer">Component</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.gettype#system-object-gettype" target="_blank" rel="noopener noreferrer">GetType</a></td>
<td>Gets the <a href="https://learn.microsoft.com/dotnet/api/system.type" target="_blank" rel="noopener noreferrer">Type</a> of the current instance.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPISession_Initialize.md">Initialize</a></td>
<td>MAPI session intialization and logon.</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.marshalbyrefobject.initializelifetimeservice#system-marshalbyrefobject-initializelifetimeservice" target="_blank" rel="noopener noreferrer">InitializeLifetimeService</a></td>
<td>Obtains a lifetime service object to control the lifetime policy for this instance.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.marshalbyrefobject" target="_blank" rel="noopener noreferrer">MarshalByRefObject</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.memberwiseclone#system-object-memberwiseclone" target="_blank" rel="noopener noreferrer">MemberwiseClone()</a></td>
<td>Creates a shallow copy of the current <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.marshalbyrefobject.memberwiseclone#system-marshalbyrefobject-memberwiseclone(system-boolean)" target="_blank" rel="noopener noreferrer">MemberwiseClone(Boolean)</a></td>
<td>Creates a shallow copy of the current <a href="https://learn.microsoft.com/dotnet/api/system.marshalbyrefobject" target="_blank" rel="noopener noreferrer">MarshalByRefObject</a> object.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.marshalbyrefobject" target="_blank" rel="noopener noreferrer">MarshalByRefObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPISession_OpendAddressBook.md">OpendAddressBook</a></td>
<td>Opens the MAPI integrated address book.</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPISession_OpenEntry.md">OpenEntry</a></td>
<td>Opens an object and returns an interface pointer for additional access.</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPISession_OpenMessageStore.md">OpenMessageStore</a></td>
<td>Opens a message store.</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPISession_SendAll.md">SendAll</a></td>
<td>Flush all outgoing messages.</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component.tostring#system-componentmodel-component-tostring" target="_blank" rel="noopener noreferrer">ToString</a></td>
<td>Returns a <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a> containing the name of the <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component" target="_blank" rel="noopener noreferrer">Component</a>, if any. This method should not be overridden.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component" target="_blank" rel="noopener noreferrer">Component</a>)</td></tr>
</table>

## Events
<table>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component.disposed" target="_blank" rel="noopener noreferrer">Disposed</a></td>
<td>Occurs when the component is disposed by a call to the <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component.dispose#system-componentmodel-component-dispose" target="_blank" rel="noopener noreferrer">Dispose()</a> method.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.componentmodel.component" target="_blank" rel="noopener noreferrer">Component</a>)</td></tr>
</table>

## See Also


#### Reference
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
