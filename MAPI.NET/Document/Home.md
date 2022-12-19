# MAPI.NET


The MAPI.NET contains interfaces and classes that define MAPI wrapper objects , which allow users to use MAPI functions from .Net level.



## Classes
<table>
<tr>
<td><a href="13ed75e1-5dd4-0ede-0e85-b151cb2a9a73.md">Appointment</a></td>
<td> </td></tr>
<tr>
<td><a href="de627363-1dfa-9d37-618f-123210bd71ef.md">Attachment</a></td>
<td>Represents an attachment to an e-mail.</td></tr>
<tr>
<td><a href="15d9a756-dc0b-8a38-6c7c-2733a049e18c.md">Contact</a></td>
<td> </td></tr>
<tr>
<td><a href="db2ff999-cb6d-b06d-47cc-55b8797d7482.md">EntryID</a></td>
<td>Represents an entry identifier for a MAPI object.</td></tr>
<tr>
<td><a href="039f2a40-3232-755a-8642-c2f615c80c69.md">MAPIAddressBook</a></td>
<td>IAddrBook .Net Wrapper object</td></tr>
<tr>
<td><a href="f0f65788-8462-2019-0156-d17cd0205fa2.md">MAPIFolder</a></td>
<td>IMAPIFolder .Net Wrapper object</td></tr>
<tr>
<td><a href="29b8d96c-1ec2-828d-35a5-fae12d8802c8.md">MAPIMessage</a></td>
<td>IMessage .Net Wrapper object</td></tr>
<tr>
<td><a href="6aa245b8-3fdd-0cd0-a3f7-bdccb4596d2c.md">MAPIObject</a></td>
<td>IMAPIProp .Net Wrapper object</td></tr>
<tr>
<td><a href="04791c9c-49a6-3b6d-99fa-53509df4be95.md">MAPIProp</a></td>
<td>Represents a property of various MAPI objects.</td></tr>
<tr>
<td><a href="565716dd-6368-0783-4ced-5771b200faf1.md">MAPISession</a></td>
<td>IMAPISession .Net wrapper object</td></tr>
<tr>
<td><a href="284425d5-5386-92cf-e310-cb18bc291055.md">MAPIStatus</a></td>
<td>IMAPIStatus .Net wrapper object</td></tr>
<tr>
<td><a href="fa40f65f-c468-2f4f-aefc-ab5a19ba58ba.md">MAPITable</a></td>
<td>.Net wrapper over IMAPITable interface.</td></tr>
<tr>
<td><a href="6f2a2863-4894-51bc-e286-04b5a90167ef.md">MsgStore</a></td>
<td>IMsgStore .Net Wrapper object</td></tr>
<tr>
<td><a href="030314f7-15ca-df6e-358f-6deb46b3381b.md">MsgStoreNewMailEventArgs</a></td>
<td>Provides data for the arrival of a new message.</td></tr>
<tr>
<td><a href="6d88cbf2-403c-24bb-f59d-466e86328fd4.md">MsgStoreObjectEventArgs</a></td>
<td>Provides data for an object that has undergone a change, such as being copied or modified.</td></tr>
<tr>
<td><a href="661b1e87-cef6-6469-0805-eb273bffec6d.md">Recipient</a></td>
<td>Defines a Recipient for MAPI messages.</td></tr>
<tr>
<td><a href="15ea5a8a-d1a8-a96f-fbfb-337247707bc3.md">RtfHtmlConverter</a></td>
<td> </td></tr>
</table>

## Structures
<table>
<tr>
<td><a href="cc3d16dd-0463-6646-eb2d-dc20ff4eaa4c.md">ADRENTRY</a></td>
<td>Describes zero or more properties that belong to a recipient.</td></tr>
<tr>
<td><a href="ebc03677-6a1a-b71d-8501-83bacf5af4d3.md">ADRLIST</a></td>
<td>An ADRLIST structure contains one or more ADRENTRY structures, each describing the properties of a recipient.</td></tr>
<tr>
<td><a href="0d5a90ba-cc29-8f93-38bb-6ae91a4c028d.md">NEWMAIL_NOTIFICATION</a></td>
<td>Describes information that relate to the arrival of a new message.</td></tr>
<tr>
<td><a href="3bd32534-061c-3006-0ac9-bea37bc973cf.md">OBJECT_NOTIFICATION</a></td>
<td>Contains information about an object that has undergone a change, such as being copied or modified.</td></tr>
<tr>
<td><a href="eacc7e7c-7e19-0ee3-9cb2-16700d317824.md">SRow</a></td>
<td>Describes a row from a table containing selected properties for a specific object.</td></tr>
<tr>
<td><a href="6cc775d7-842b-3fa0-ca6b-61f67dc4c98b.md">SSortOrder</a></td>
<td>Defines how to sort rows of a table, describing the column to use as the sort key and the direction of the sort.</td></tr>
</table>

## Interfaces
<table>
<tr>
<td><a href="3e0ae0ab-2ec1-3cb4-6c4f-5d6faee00a6e.md">IAddrBook</a></td>
<td>Supports access to the MAPI address book and includes operations such as displaying common dialog boxes; opening containers, messaging users, and distribution lists; and performing name resolution.</td></tr>
<tr>
<td><a href="ce25a38b-9434-ec81-c314-5444e5b10bd9.md">IAttach</a></td>
<td>The IAttach interface is used to maintain and provide access to the properties of attachments in messages.</td></tr>
<tr>
<td><a href="c97c2b5a-4844-a7b2-caa5-d1278d87cf97.md">IMAPIAdviseSink</a></td>
<td>The IMAPIAdviseSink interface is used to implement an Advise Sink object for handling notifications.</td></tr>
<tr>
<td><a href="d9a68088-6545-338f-9dc8-439874dbd7a1.md">IMAPIContainer</a></td>
<td>Manages high-level operations on container objects such as address books, distribution lists, and folders.</td></tr>
<tr>
<td><a href="a5eb5918-6571-0710-67c7-a210d1ad706f.md">IMAPIFolder</a></td>
<td>PerPerforms operations on the messages and subfolders in a folder.forms operations on the messages and subfolders in a folder.</td></tr>
<tr>
<td><a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp</a></td>
<td>Enables clients, service providers, and MAPI to work with properties. All objects that support properties implement this interface.</td></tr>
<tr>
<td><a href="d28ec202-b730-fb1f-42ac-5545b0b43d47.md">IMAPISession</a></td>
<td>Manages objects associated with a MAPI logon session.</td></tr>
<tr>
<td><a href="e0749ad9-46d7-9716-4d9d-030334fc0ed3.md">IMAPIStatus</a></td>
<td>Provides status information about the MAPI subsystem, the integrated address book, and the MAPI spooler.</td></tr>
<tr>
<td><a href="06a9b727-f5d6-e992-c936-a2712197dcee.md">IMAPITable</a></td>
<td>Provides a read-only view of a table. IMAPITable is used by clients and service providers to manipulate the way a table appears.</td></tr>
<tr>
<td><a href="f542b7a9-d1ab-fed6-c2df-7c20b044fccc.md">IMessage</a></td>
<td>Manages messages, attachments, and recipients.</td></tr>
<tr>
<td><a href="74ee1853-dea0-4e58-cb66-c6c8017d5a04.md">IMsgStore</a></td>
<td>Provides access to message store information and to messages and folders.</td></tr>
<tr>
<td><a href="2a268271-39cd-b9bd-d434-1bd1ce5d3066.md">IPropValue</a></td>
<td>Provides an interface to manage MAPI property value</td></tr>
</table>

## Enumerations
<table>
<tr>
<td><a href="549f17d5-0e76-b912-7cc0-521750417dfa.md">AddressType</a></td>
<td>Recipient address type enumeration</td></tr>
<tr>
<td><a href="14e2f584-f92c-9419-6aba-fba6f3d22327.md">AttachMethod</a></td>
<td>AttachMehod contains a MAPI-defined constant representing the way the contents of an attachment can be accessed.</td></tr>
<tr>
<td><a href="b9417f7b-9fe9-5616-bc23-0dea95fc592f.md">BookMark</a></td>
<td>The bookmark identifying the starting position for the seek operation.</td></tr>
<tr>
<td><a href="dadf865e-e172-1376-6484-03f12b8e455a.md">CharacterSet</a></td>
<td>Bitmask of flags that controls the mapi character set.</td></tr>
<tr>
<td><a href="5a15b17a-4117-5c7c-d72d-e89c6cb03fe4.md">EEventMask</a></td>
<td>A mask of values that indicate the types of notification events that the caller is interested in and should be included in the registration. There is a corresponding NOTIFICATION structure associated with each type of event that holds information about the event.</td></tr>
<tr>
<td><a href="a8aa90da-9176-c9f6-efb7-249af2c5a049.md">FlushFlag</a></td>
<td>Bitmask of flags that controls the flush operation.</td></tr>
<tr>
<td><a href="50596607-a328-ef10-6ea9-0448fbb7d197.md">HRESULT</a></td>
<td>The HRESULT data type is a 32-bit value is used to describe an error or warning.</td></tr>
<tr>
<td><a href="dd17ca09-28c9-081a-d0ae-952c371256eb.md">Importance</a></td>
<td> </td></tr>
<tr>
<td><a href="7220477e-4546-a1f5-bc16-da59253841a8.md">MAPIFlag</a></td>
<td>Generic MAPI Bitmask</td></tr>
<tr>
<td><a href="0975897b-3f97-989d-67c8-b122390b4252.md">MessageFormat</a></td>
<td> </td></tr>
<tr>
<td><a href="482631ca-d80c-485d-f070-1f3aeaa04ecd.md">ModifyRecipientFlag</a></td>
<td>A bitmask of flags that controls recipient</td></tr>
<tr>
<td><a href="cf188769-ec48-5722-68da-fd8239b3601b.md">ObjectType</a></td>
<td>Type of MAPI object</td></tr>
<tr>
<td><a href="3437a9c9-1746-4adf-e9be-22a29a6f431c.md">OpenPropertyMode</a></td>
<td>A bitmask of flags that controls access to the property.</td></tr>
<tr>
<td><a href="1ae9a3cd-e604-b415-e46a-a883db158f2a.md">PropTags</a></td>
<td>MAPI Property Tag</td></tr>
<tr>
<td><a href="dd8f0d98-0ec4-9e22-d22c-9e6cd7e32e9c.md">PT</a></td>
<td>Property Type</td></tr>
<tr>
<td><a href="14320c7c-e367-59b1-9f4f-88100fa32543.md">RecipientType</a></td>
<td>Contains the recipient type for a message recipient.</td></tr>
<tr>
<td><a href="54e85136-cce0-088d-fb78-fe89a2bc8e60.md">SaveFlags</a></td>
<td>A bitmask of flags that controls what happens to the object when the IMAPIProp::SaveChanges method is called.</td></tr>
<tr>
<td><a href="2aef4a23-80c8-a80c-c45f-dd5fa7514c18.md">Sensitivity</a></td>
<td> </td></tr>
<tr>
<td><a href="46600767-8fde-69ef-002c-64f0976b4e69.md">TableSortOrder</a></td>
<td>Defines how to sort the rows of a table, what column to use as the sort key, and the direction of the sort.</td></tr>
</table>