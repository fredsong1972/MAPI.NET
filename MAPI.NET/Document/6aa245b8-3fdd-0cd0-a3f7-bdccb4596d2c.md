# MAPIObject Class


IMAPIProp .Net Wrapper object



## Definition
**Namespace:** <a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET</a>  
**Assembly:** AllNetMAPI (in AllNetMAPI.dll) Version: 1.0.0.0 (1.0.0.0)

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
**F#**
``` F#
type MAPIObject = 
    class
        interface IDisposable
    end
```

<table><tr><td><strong>Inheritance</strong></td><td><a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>  →  MAPIObject</td></tr>
<tr><td><strong>Derived</strong></td><td><a href="de627363-1dfa-9d37-618f-123210bd71ef.md">MAPI.NET.Attachment</a><br /><a href="f0f65788-8462-2019-0156-d17cd0205fa2.md">MAPI.NET.MAPIFolder</a><br /><a href="29b8d96c-1ec2-828d-35a5-fae12d8802c8.md">MAPI.NET.MAPIMessage</a><br /><a href="284425d5-5386-92cf-e310-cb18bc291055.md">MAPI.NET.MAPIStatus</a><br /><a href="6f2a2863-4894-51bc-e286-04b5a90167ef.md">MAPI.NET.MsgStore</a></td></tr>
<tr><td><strong>Implements</strong></td><td><a href="https://learn.microsoft.com/dotnet/api/system.idisposable" target="_blank" rel="noopener noreferrer">IDisposable</a></td></tr>
</table>



## Constructors
<table>
<tr>
<td><a href="1d6a30d2-e162-dce2-08b0-66d4e2efd51f.md">MAPIObject(IMAPIProp)</a></td>
<td>Initializes a new instance of the MAPIObject class.</td></tr>
<tr>
<td><a href="7d009418-727d-d308-a94c-eee0beda569a.md">MAPIObject(IMAPIProp, EntryID)</a></td>
<td>Initializes a new instance of the MAPIObject class.</td></tr>
<tr>
<td><a href="20d6e313-58f7-e871-099a-bdefd1765510.md">MAPIObject(MAPISession, EntryID)</a></td>
<td>Initializes a new instance of the MAPIObject class.</td></tr>
</table>

## Properties
<table>
<tr>
<td><a href="361b1fae-bad7-de5a-f54e-55df88c08a15.md">EntryID</a></td>
<td>EntryID of MAPIObject</td></tr>
<tr>
<td><a href="760157ae-77d7-574f-57ee-ff447325863b.md">InterfaceID</a></td>
<td>IMAPIProp Interface Guid</td></tr>
</table>

## Methods
<table>
<tr>
<td><a href="b1604318-dca4-e638-24ff-96c115bcb92d.md">Close</a></td>
<td>Makes permanent any changes that were made to an object since the last save operation,and close the object.</td></tr>
<tr>
<td><a href="f90c1bb1-8f48-8cfd-7471-94b3d7c93ddb.md">CopyTo(UInt32[], IMAPIProp)</a></td>
<td>Copies or moves all properties, except for specifically excluded properties.</td></tr>
<tr>
<td><a href="aefe3b61-4ef7-0557-d875-ac15ca0ca7da.md">CopyTo(UInt32[], MAPIObject)</a></td>
<td> </td></tr>
<tr>
<td><a href="30bbca25-2433-aec6-4a4f-081540f03dd4.md">Dispose</a></td>
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
<td><a href="c8497ef8-5ce6-7bfb-37ac-3f62b8e67081.md">GetOutlookProperties</a></td>
<td> </td></tr>
<tr>
<td><a href="446193b7-c3bd-12a8-ba52-2a082f145431.md">GetOutlookProperty</a></td>
<td> </td></tr>
<tr>
<td><a href="368f345a-10a0-293f-f0f8-c15b1b5b756f.md">GetOutlookPropTag</a></td>
<td> </td></tr>
<tr>
<td><a href="0b3c6f7c-b5e3-b2e3-b39e-4f44c5f3be14.md">GetOutlookPropTags</a></td>
<td> </td></tr>
<tr>
<td><a href="127b1def-bf2a-8712-68af-ba6681d8691e.md">GetProperties(PropTags[])</a></td>
<td>Retrieves the property value of one or more properties of an object.</td></tr>
<tr>
<td><a href="04afec4b-6454-d08f-abc2-8c208393caf1.md">GetProperties(UInt32[])</a></td>
<td>Retrieves the property value of one or more properties of an object.</td></tr>
<tr>
<td><a href="0d817cf0-fed1-ddb0-84ad-7bba034d9b5a.md">GetProperty(PropTags)</a></td>
<td>Retrieves one property value of an object.</td></tr>
<tr>
<td><a href="5bdc244a-b327-1fcb-6248-63efd0baf6b8.md">GetProperty(UInt32)</a></td>
<td>Retrieves one property value of an object.</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.gettype#system-object-gettype" target="_blank" rel="noopener noreferrer">GetType</a></td>
<td>Gets the <a href="https://learn.microsoft.com/dotnet/api/system.type" target="_blank" rel="noopener noreferrer">Type</a> of the current instance.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.memberwiseclone#system-object-memberwiseclone" target="_blank" rel="noopener noreferrer">MemberwiseClone</a></td>
<td>Creates a shallow copy of the current <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
<tr>
<td><a href="3b4e18d1-557e-1b4a-8b40-887c1d12a896.md">OpenProperty(UInt32, Guid, OpenPropertyMode)</a></td>
<td>Returns a pointer to an interface that can be used to access a property.</td></tr>
<tr>
<td><a href="fadcb7a0-6c25-07e1-9829-78b5befe3332.md">OpenProperty(UInt32, Guid, UInt32, OpenPropertyMode)</a></td>
<td>Returns a pointer to an interface that can be used to access a property.</td></tr>
<tr>
<td><a href="35ac712c-f619-848b-c083-49e9caba63d3.md">Save()</a></td>
<td>Makes permanent any changes that were made to an object since the last save operation,and keep the object open.</td></tr>
<tr>
<td><a href="881bfeda-e0da-7af8-44a5-01a7ae472761.md">Save(SaveFlags)</a></td>
<td>Makes permanent any changes that were made to an object since the last save operation,and control the object per the flag.</td></tr>
<tr>
<td><a href="8676076d-7624-8b70-6965-26b95249236c.md">SetPropertyValue(IPropValue)</a></td>
<td>Updates one property of the object.</td></tr>
<tr>
<td><a href="73313aee-42c6-5528-6c07-eb16297b5558.md">SetPropertyValue(PropTags, Object)</a></td>
<td>Updates one property of the object.</td></tr>
<tr>
<td><a href="a20ed580-2449-2a36-e5a3-95d803f8c0c7.md">SetPropertyValue(UInt32, Object)</a></td>
<td>Updates one property of the object.</td></tr>
<tr>
<td><a href="0d738f8a-e192-3393-0e75-298ad7c1c0d3.md">SetPropertyValues</a></td>
<td>Updates one or more properties of the object.</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.tostring#system-object-tostring" target="_blank" rel="noopener noreferrer">ToString</a></td>
<td>Returns a string that represents the current object.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
</table>

## Operators
<table>
<tr>
<td><a href="cb3ce7a7-58b3-2619-9c19-f64eab425389.md">Explicit(MAPIObject to IntPtr)</a></td>
<td> </td></tr>
</table>

## Fields
<table>
<tr>
<td><a href="3507b1fb-f2e3-fa88-5ab4-6d26f8426d1a.md">AppointmentItem</a></td>
<td> </td></tr>
<tr>
<td><a href="34d07db5-46d9-7929-f822-a3a57fb7920e.md">ContactItem</a></td>
<td> </td></tr>
<tr>
<td><a href="919fe468-c687-6efa-8c56-6ba5c82d0bc4.md">Id_</a></td>
<td>EntryID of the object</td></tr>
<tr>
<td><a href="7277df44-418a-ddd2-fa04-46a593193353.md">MailItem</a></td>
<td> </td></tr>
<tr>
<td><a href="a476991e-8077-7463-4103-757919cd40ce.md">MeetingCanceledItem</a></td>
<td> </td></tr>
<tr>
<td><a href="9ff70a0d-75aa-60b7-0203-6cf3394bef1f.md">MeetingItem</a></td>
<td> </td></tr>
<tr>
<td><a href="1ad070c6-1ee1-d626-c1df-0195ae3a1a95.md">MeetingRequestItem</a></td>
<td> </td></tr>
<tr>
<td><a href="ee381c38-b575-f081-0835-735360f683c6.md">MeetingResponseItem</a></td>
<td> </td></tr>
</table>

## See Also


#### Reference
<a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET Namespace</a>  
