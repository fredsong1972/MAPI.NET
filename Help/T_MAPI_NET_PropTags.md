# PropTags Enumeration


MAPI Property Tag



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public enum PropTags
```
**VB**
``` VB
Public Enumeration PropTags
```
**C++**
``` C++
public enum class PropTags
```



## Members
<table>
<tr>
<td>PR_NULL</td>
<td>1</td>
<td>Represents a null value or setting of a property or reserves array space. <br />Data Type: PT_NULL</td></tr>
<tr>
<td>PR_ERROR</td>
<td>10</td>
<td>Error</td></tr>
<tr>
<td>PR_ACKNOWLEDGEMENT_MODE</td>
<td>65,539</td>
<td>Contains the identifier of the mode for message acknowledgment. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_ALTERNATE_RECIPIENT_ALLOWED</td>
<td>131,083</td>
<td>Contains TRUE if the sender permits autoforwarding of this message. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_AUTHORIZING_USERS</td>
<td>196,866</td>
<td>Contains a list of entry identifiers for users who have authorized the sending of a message. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_AUTO_FORWARD_COMMENT_A</td>
<td>262,174</td>
<td>Contains a comment added by the autoforwarding agent. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_AUTO_FORWARD_COMMENT</td>
<td>262,175</td>
<td>Contains a comment added by the autoforwarding agent. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_AUTO_FORWARD_COMMENT_W</td>
<td>262,175</td>
<td>Contains a comment added by the autoforwarding agent. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_AUTO_FORWARDED</td>
<td>327,691</td>
<td>Contains TRUE if an automatic agent has forwarded a message. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_CONTENT_CONFIDENTIALITY_ALGORITHM_ID</td>
<td>393,474</td>
<td>Contains an identifier for the algorithm used to confirm message content confidentiality. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_CONTENT_CORRELATOR</td>
<td>459,010</td>
<td>Contains a value the message sender can use to match a report with the original message. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_CONTENT_IDENTIFIER_A</td>
<td>524,318</td>
<td>Contains a key value that enables the message recipient to identify its content. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_CONTENT_IDENTIFIER</td>
<td>524,319</td>
<td>Contains a key value that enables the message recipient to identify its content. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_CONTENT_IDENTIFIER_W</td>
<td>524,319</td>
<td>Contains a key value that enables the message recipient to identify its content. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_CONTENT_LENGTH</td>
<td>589,827</td>
<td>Contains a message length, in bytes, passed to a client application or service provider to determine if a message of that length can be delivered. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_CONTENT_RETURN_REQUESTED</td>
<td>655,371</td>
<td>Contains TRUE if a message should be returned with a nondelivery report. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_CONVERSION_EITS</td>
<td>786,690</td>
<td>Contains the encoded information types (EITs) that are applied to a message in transit to describe conversions. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_CONVERSION_WITH_LOSS_PROHIBITED</td>
<td>851,979</td>
<td>Contains TRUE if a message transfer agent (MTA) is prohibited from making message text conversions that lose information. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_CONVERTED_EITS</td>
<td>917,762</td>
<td>Contains an identifier for the types of text in a message after conversion. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_DEFERRED_DELIVERY_TIME</td>
<td>983,104</td>
<td>Contains the date and time at which a message sender wants a message delivered. <br />Data Type: PT_SYSTIME</td></tr>
<tr>
<td>PR_DELIVER_TIME</td>
<td>1,048,640</td>
<td>Contains the date and time at which the original message was delivered. <br />Data Type: PT_SYSTIME</td></tr>
<tr>
<td>PR_DISCARD_REASON</td>
<td>1,114,115</td>
<td>Contains a reason why a message transfer agent (MTA) has discarded a message. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_DISCLOSURE_OF_RECIPIENTS</td>
<td>1,179,659</td>
<td>Contains TRUE if disclosure of recipients is allowed. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_DL_EXPANSION_HISTORY</td>
<td>1,245,442</td>
<td>Contains a history showing how a distribution list has been expanded during message transmission. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_DL_EXPANSION_PROHIBITED</td>
<td>1,310,731</td>
<td>Contains TRUE if a message transfer agent (MTA) is prohibited from expanding distribution lists. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_EXPIRY_TIME</td>
<td>1,376,320</td>
<td>Contains the date and time at which the messaging system can invalidate the content of a message. <br />Data Type: PT_SYSTIME</td></tr>
<tr>
<td>PR_IMPLICIT_CONVERSION_PROHIBITED</td>
<td>1,441,803</td>
<td>Contains TRUE if a message transfer agent (MTA) is prohibited from making implicit message text conversions. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_IMPORTANCE</td>
<td>1,507,331</td>
<td>Contains a value indicating the message sender's opinion of the importance of a message. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_LATEST_DELIVERY_TIME</td>
<td>1,638,464</td>
<td>Contains the latest date and time when a message transfer agent (MTA) should deliver a message. <br />Data Type: PT_SYSTIME</td></tr>
<tr>
<td>PR_MESSAGE_CLASS_A</td>
<td>1,703,966</td>
<td>Contains a text string that identifies the sender-defined message class, such as IPM.Note. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_MESSAGE_CLASS</td>
<td>1,703,967</td>
<td>Contains a text string that identifies the sender-defined message class, such as IPM.Note. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_MESSAGE_CLASS_W</td>
<td>1,703,967</td>
<td>Contains a text string that identifies the sender-defined message class, such as IPM.Note. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_MESSAGE_DELIVERY_ID</td>
<td>1,769,730</td>
<td>Contains a message transfer system (MTS) identifier for a message delivered to a client application. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_MESSAGE_SECURITY_LABEL</td>
<td>1,966,338</td>
<td>Contains a security label for a message. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_OBSOLETED_IPMS</td>
<td>2,031,874</td>
<td>Contains the identifiers of messages that this message supersedes <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_ORIGINALLY_INTENDED_RECIPIENT_NAME</td>
<td>2,097,410</td>
<td>Contains the encoded name of the originally intended recipient of an autoforwarded message. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_ORIGINAL_EITS</td>
<td>2,162,946</td>
<td>Contains a copy of the original encoded information types (EITs) for message text. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_ORIGINATOR_CERTIFICATE</td>
<td>2,228,482</td>
<td>Contains an ASN.1 certificate for the message originator. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_ORIGINATOR_DELIVERY_REPORT_REQUESTED</td>
<td>2,293,771</td>
<td>Contains TRUE if a message sender requests a delivery report for a particular recipient from the messaging system before the message is placed in the message store. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_ORIGINATOR_RETURN_ADDRESS</td>
<td>2,359,554</td>
<td>Contains the binary-encoded return address of the message originator. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_PRIORITY</td>
<td>2,490,371</td>
<td>Contains the relative priority of a message. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_ORIGIN_CHECK</td>
<td>2,556,162</td>
<td>Contains a binary verification value enabling a delivery report recipient to verify the origin of the original message. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_PROOF_OF_SUBMISSION_REQUESTED</td>
<td>2,621,451</td>
<td>Contains TRUE if a message sender requests proof that the message transfer system has submitted a message for delivery to the originally intended recipient. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_READ_RECEIPT_REQUESTED</td>
<td>2,686,987</td>
<td>Contains TRUE if a message sender wants the messaging system to generate a read report when the recipient has read a message. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_RECEIPT_TIME</td>
<td>2,752,576</td>
<td>Contains the date and time a delivery report is generated. <br />Data Type: PT_SYSTIME</td></tr>
<tr>
<td>PR_RECIPIENT_REASSIGNMENT_PROHIBITED</td>
<td>2,818,059</td>
<td>Contains TRUE if recipient reassignment is prohibited. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_REDIRECTION_HISTORY</td>
<td>2,883,842</td>
<td>Contains information about the route covered by a delivered message. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_RELATED_IPMS</td>
<td>2,949,378</td>
<td>Contains a list of identifiers for messages to which a message is related. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_ORIGINAL_SENSITIVITY</td>
<td>3,014,659</td>
<td>Contains the sensitivity value assigned by the sender of the first version of a message — that is, the message before being forwarded or replied to. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_LANGUAGES_A</td>
<td>3,080,222</td>
<td>Contains an ASCII list of the languages incorporated in a message. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_LANGUAGES</td>
<td>3,080,223</td>
<td>Contains an ASCII list of the languages incorporated in a message. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_LANGUAGES_W</td>
<td>3,080,223</td>
<td>Contains an ASCII list of the languages incorporated in a message. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_REPLY_TIME</td>
<td>3,145,792</td>
<td>Contains the date and time by which a reply is expected for a message. <br />Data Type: PT_SYSTIME</td></tr>
<tr>
<td>PR_REPORT_TAG</td>
<td>3,211,522</td>
<td>Contains a binary tag value that the messaging system should copy to any report generated for the message. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_REPORT_TIME</td>
<td>3,276,864</td>
<td>Contains the date and time when the messaging system generated a report. <br />Data Type: PT_SYSTIME</td></tr>
<tr>
<td>PR_RETURNED_IPM</td>
<td>3,342,347</td>
<td>Contains TRUE if the original message is being returned with a nonread report. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_SECURITY</td>
<td>3,407,875</td>
<td>Contains a flag that indicates the security level of a message. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_INCOMPLETE_COPY</td>
<td>3,473,419</td>
<td>Contains TRUE if this message is an incomplete copy of another message. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_SENSITIVITY</td>
<td>3,538,947</td>
<td>Contains a value indicating the message sender's opinion of the sensitivity of a message. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_SUBJECT_A</td>
<td>3,604,510</td>
<td>Contains the full subject of a message. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_SUBJECT</td>
<td>3,604,511</td>
<td>Contains the full subject of a message. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_SUBJECT_W</td>
<td>3,604,511</td>
<td>Contains the full subject of a message. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_SUBJECT_IPM</td>
<td>3,670,274</td>
<td>Contains a binary value that is copied from the message for which a report is being generated. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_CLIENT_SUBMIT_TIME</td>
<td>3,735,616</td>
<td>Contains the date and time the message sender submitted a message. <br />Data Type: PT_SYSTIME</td></tr>
<tr>
<td>PR_REPORT_NAME_A</td>
<td>3,801,118</td>
<td>Contains the display name for the recipient that should get reports for this message. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_REPORT_NAME</td>
<td>3,801,119</td>
<td>Contains the display name for the recipient that should get reports for this message. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_REPORT_NAME_W</td>
<td>3,801,119</td>
<td>Contains the display name for the recipient that should get reports for this message. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_SENT_REPRESENTING_SEARCH_KEY</td>
<td>3,866,882</td>
<td>Contains the search key for the messaging user represented by the sender. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_X400_CONTENT_TYPE</td>
<td>3,932,418</td>
<td>Contains the content type for a submitted message. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_SUBJECT_PREFIX_A</td>
<td>3,997,726</td>
<td>Contains a subject prefix that typically indicates some action on a message. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_SUBJECT_PREFIX</td>
<td>3,997,727</td>
<td>Contains a subject prefix that typically indicates some action on a message. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_SUBJECT_PREFIX_W</td>
<td>3,997,727</td>
<td>Contains a subject prefix that typically indicates some action on a message. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_NON_RECEIPT_REASON</td>
<td>4,063,235</td>
<td>dContains reasons why a message was not received that forms part of a nondelivery report. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_RECEIVED_BY_ENTRYID</td>
<td>4,129,026</td>
<td>Contains the entry identifier of the messaging user that actually receives the message. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_RECEIVED_BY_NAME_A</td>
<td>4,194,334</td>
<td>Contains the display name of the messaging user that actually receives the message. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_RECEIVED_BY_NAME</td>
<td>4,194,335</td>
<td>Contains the display name of the messaging user that actually receives the message. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_RECEIVED_BY_NAME_W</td>
<td>4,194,335</td>
<td>Contains the display name of the messaging user that actually receives the message. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_SENT_REPRESENTING_ENTRYID</td>
<td>4,260,098</td>
<td>Contains the entry identifier for the messaging user represented by the sender. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_SENT_REPRESENTING_NAME_A</td>
<td>4,325,406</td>
<td>Contains the display name for the messaging user represented by the sender. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_SENT_REPRESENTING_NAME</td>
<td>4,325,407</td>
<td>Contains the display name for the messaging user represented by the sender. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_SENT_REPRESENTING_NAME_W</td>
<td>4,325,407</td>
<td>Contains the display name for the messaging user represented by the sender. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_RCVD_REPRESENTING_ENTRYID</td>
<td>4,391,170</td>
<td>Contains the entry identifier for the messaging user represented by the receiving user. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_RCVD_REPRESENTING_NAME_A</td>
<td>4,456,478</td>
<td>Contains the display name for the messaging user represented by the receiving user. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_RCVD_REPRESENTING_NAME</td>
<td>4,456,479</td>
<td>Contains the display name for the messaging user represented by the receiving user. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_RCVD_REPRESENTING_NAME_W</td>
<td>4,456,479</td>
<td>Contains the display name for the messaging user represented by the receiving user. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_REPORT_ENTRYID</td>
<td>4,522,242</td>
<td>Contains the entry identifier for the recipient that should get reports for this message. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_READ_RECEIPT_ENTRYID</td>
<td>4,587,778</td>
<td>Contains an entry identifier for the messaging user to which the messaging system should direct a read report for this message. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_MESSAGE_SUBMISSION_ID</td>
<td>4,653,314</td>
<td>Contains a message transfer system (MTS) identifier for the message transfer agent (MTA). <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_PROVIDER_SUBMIT_TIME</td>
<td>4,718,656</td>
<td>Contains the date and time a transport provider passed a message to its underlying messaging system. <br />Data Type: PT_SYSTIME</td></tr>
<tr>
<td>PR_ORIGINAL_SUBJECT_A</td>
<td>4,784,158</td>
<td>Contains the subject of an original message for use in a report about the message. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_ORIGINAL_SUBJECT</td>
<td>4,784,159</td>
<td>Contains the subject of an original message for use in a report about the message. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ORIGINAL_SUBJECT_W</td>
<td>4,784,159</td>
<td>Contains the subject of an original message for use in a report about the message. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_ORIG_MESSAGE_CLASS_A</td>
<td>4,915,230</td>
<td>Contains the class of the original message for use in a report. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_ORIG_MESSAGE_CLASS</td>
<td>4,915,231</td>
<td>Contains the class of the original message for use in a report. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ORIG_MESSAGE_CLASS_W</td>
<td>4,915,231</td>
<td>Contains the class of the original message for use in a report. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_ORIGINAL_AUTHOR_ENTRYID</td>
<td>4,980,994</td>
<td>Contains the entry identifier of the author of the first version of a message, that is, the message before being forwarded or replied to. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_ORIGINAL_AUTHOR_NAME_A</td>
<td>5,046,302</td>
<td>Contains the display name of the author of the first version of a message, that is, the message before being forwarded or replied to. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_ORIGINAL_AUTHOR_NAME</td>
<td>5,046,303</td>
<td>Contains the display name of the author of the first version of a message, that is, the message before being forwarded or replied to. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ORIGINAL_AUTHOR_NAME_W</td>
<td>5,046,303</td>
<td>Contains the display name of the author of the first version of a message, that is, the message before being forwarded or replied to. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_ORIGINAL_SUBMIT_TIME</td>
<td>5,111,872</td>
<td>Contains the original submission date and time of the message in the report. <br />Data Type: PT_SYSTIME</td></tr>
<tr>
<td>PR_REPLY_RECIPIENT_ENTRIES</td>
<td>5,177,602</td>
<td>Contains a sized array of entry identifiers for recipients that are to get a reply. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_REPLY_RECIPIENT_NAMES_A</td>
<td>5,242,910</td>
<td>Contains a list of display names for recipients that are to get a reply. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_REPLY_RECIPIENT_NAMES</td>
<td>5,242,911</td>
<td>Contains a list of display names for recipients that are to get a reply. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_REPLY_RECIPIENT_NAMES_W</td>
<td>5,242,911</td>
<td>Contains a list of display names for recipients that are to get a reply. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_RECEIVED_BY_SEARCH_KEY</td>
<td>5,308,674</td>
<td>Contains the search key of the messaging user that actually receives the message. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_RCVD_REPRESENTING_SEARCH_KEY</td>
<td>5,374,210</td>
<td>Contains the search key for the messaging user represented by the receiving user. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_READ_RECEIPT_SEARCH_KEY</td>
<td>5,439,746</td>
<td>Contains a search key for the messaging user to which the messaging system should direct a read report for a message. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_REPORT_SEARCH_KEY</td>
<td>5,505,282</td>
<td>Contains the search key for the recipient that should get reports for this message. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_ORIGINAL_DELIVERY_TIME</td>
<td>5,570,624</td>
<td>Contains a copy of the original message's delivery date and time in a thread. <br />Data Type: PT_SYSTIME</td></tr>
<tr>
<td>PR_ORIGINAL_AUTHOR_SEARCH_KEY</td>
<td>5,636,354</td>
<td>Contains the search key of the author of the first version of a message, that is, the message before being forwarded or replied to. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_MESSAGE_TO_ME</td>
<td>5,701,643</td>
<td>Contains TRUE if this messaging user is specifically named as a primary (To) recipient of this message and is not part of a distribution list. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_MESSAGE_CC_ME</td>
<td>5,767,179</td>
<td>Contains TRUE if this messaging user is specifically named as a carbon copy (CC) recipient of this message and is not part of a distribution list. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_MESSAGE_RECIP_ME</td>
<td>5,832,715</td>
<td>Contains TRUE if this messaging user is specifically named as a primary (To), carbon copy (CC), or blind carbon copy (BCC) recipient of this message and is not part of a distribution list. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_ORIGINAL_SENDER_NAME_A</td>
<td>5,898,270</td>
<td>Contains the display name of the sender of the first version of a message, that is, the message before being forwarded or replied to. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_ORIGINAL_SENDER_NAME</td>
<td>5,898,271</td>
<td>Contains the display name of the sender of the first version of a message, that is, the message before being forwarded or replied to. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ORIGINAL_SENDER_NAME_W</td>
<td>5,898,271</td>
<td>Contains the display name of the sender of the first version of a message, that is, the message before being forwarded or replied to. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_ORIGINAL_SENDER_ENTRYID</td>
<td>5,964,034</td>
<td>Contains the entry identifier of the sender of the first version of a message, that is, the message before being forwarded or replied to. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_ORIGINAL_SENDER_SEARCH_KEY</td>
<td>6,029,570</td>
<td>Contains the search key for the sender of the first version of a message, that is, the message before being forwarded or replied to. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_ORIGINAL_SENT_REPRESENTING_NAME_A</td>
<td>6,094,878</td>
<td>Contains the display name of the messaging user on whose behalf the original message was sent. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_ORIGINAL_SENT_REPRESENTING_NAME</td>
<td>6,094,879</td>
<td>Contains the display name of the messaging user on whose behalf the original message was sent. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ORIGINAL_SENT_REPRESENTING_NAME_W</td>
<td>6,094,879</td>
<td>Contains the display name of the messaging user on whose behalf the original message was sent. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_ORIGINAL_SENT_REPRESENTING_ENTRYID</td>
<td>6,160,642</td>
<td>Contains the entry identifier of the messaging user on whose behalf the original message was sent. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_ORIGINAL_SENT_REPRESENTING_SEARCH_KEY</td>
<td>6,226,178</td>
<td>Contains the search key of the messaging user on whose behalf the original message was sent. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_START_DATE</td>
<td>6,291,520</td>
<td>Contains the starting date and time of an appointment as managed by a scheduling application. <br />Data Type: PT_SYSTIME</td></tr>
<tr>
<td>PR_END_DATE</td>
<td>6,357,056</td>
<td>Contains the ending date and time of an appointment as managed by a scheduling application. <br />Data Type: PT_SYSTIME</td></tr>
<tr>
<td>PR_OWNER_APPT_ID</td>
<td>6,422,531</td>
<td>Contains an identifier for an appointment in the owner's schedule. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_RESPONSE_REQUESTED</td>
<td>6,488,075</td>
<td>Contains TRUE if the message sender wants a response to a meeting request. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_SENT_REPRESENTING_ADDRTYPE_A</td>
<td>6,553,630</td>
<td>Contains the address type for the messaging user represented by the sender. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_SENT_REPRESENTING_ADDRTYPE</td>
<td>6,553,631</td>
<td>Contains the address type for the messaging user represented by the sender. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_SENT_REPRESENTING_ADDRTYPE_W</td>
<td>6,553,631</td>
<td>Contains the address type for the messaging user represented by the sender. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_SENT_REPRESENTING_EMAIL_ADDRESS_A</td>
<td>6,619,166</td>
<td>Contains the e-mail address for the messaging user represented by the sender. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_SENT_REPRESENTING_EMAIL_ADDRESS</td>
<td>6,619,167</td>
<td>Contains the e-mail address for the messaging user represented by the sender. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_SENT_REPRESENTING_EMAIL_ADDRESS_W</td>
<td>6,619,167</td>
<td>Contains the e-mail address for the messaging user represented by the sender. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_ORIGINAL_SENDER_ADDRTYPE_A</td>
<td>6,684,702</td>
<td>Contains the address type of the sender of the first version of a message, that is, the message before being forwarded or replied to. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_ORIGINAL_SENDER_ADDRTYPE</td>
<td>6,684,703</td>
<td>Contains the address type of the sender of the first version of a message, that is, the message before being forwarded or replied to. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ORIGINAL_SENDER_ADDRTYPE_W</td>
<td>6,684,703</td>
<td>Contains the address type of the sender of the first version of a message, that is, the message before being forwarded or replied to. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_ORIGINAL_SENDER_EMAIL_ADDRESS_A</td>
<td>6,750,238</td>
<td>Contains the e-mail address of the sender of the first version of a message, that is, the message before being forwarded or replied to. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_ORIGINAL_SENDER_EMAIL_ADDRESS</td>
<td>6,750,239</td>
<td>Contains the e-mail address of the sender of the first version of a message, that is, the message before being forwarded or replied to. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ORIGINAL_SENDER_EMAIL_ADDRESS_W</td>
<td>6,750,239</td>
<td>Contains the e-mail address of the sender of the first version of a message, that is, the message before being forwarded or replied to. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_ORIGINAL_SENT_REPRESENTING_ADDRTYPE_A</td>
<td>6,815,774</td>
<td>Contains the address type of the messaging user on whose behalf the original message was sent. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_ORIGINAL_SENT_REPRESENTING_ADDRTYPE</td>
<td>6,815,775</td>
<td>Contains the address type of the messaging user on whose behalf the original message was sent. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ORIGINAL_SENT_REPRESENTING_ADDRTYPE_W</td>
<td>6,815,775</td>
<td>Contains the address type of the messaging user on whose behalf the original message was sent. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS_A</td>
<td>6,881,310</td>
<td>Contains the e-mail address of the messaging user on whose behalf the original message was sent. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS</td>
<td>6,881,311</td>
<td>Contains the e-mail address of the messaging user on whose behalf the original message was sent. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS_W</td>
<td>6,881,311</td>
<td>Contains the e-mail address of the messaging user on whose behalf the original message was sent. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_CONVERSATION_TOPIC_A</td>
<td>7,340,062</td>
<td>Contains the topic of the first message in a conversation thread. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_CONVERSATION_TOPIC</td>
<td>7,340,063</td>
<td>Contains the topic of the first message in a conversation thread. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_CONVERSATION_TOPIC_W</td>
<td>7,340,063</td>
<td>Contains the topic of the first message in a conversation thread. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_CONVERSATION_INDEX</td>
<td>7,405,826</td>
<td>Contains a binary value that indicates the relative position of this message within a conversation thread. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_ORIGINAL_DISPLAY_BCC_A</td>
<td>7,471,134</td>
<td>Contains the display names of any blind carbon copy (BCC) recipients of the original message. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_ORIGINAL_DISPLAY_BCC</td>
<td>7,471,135</td>
<td>Contains the display names of any blind carbon copy (BCC) recipients of the original message. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ORIGINAL_DISPLAY_BCC_W</td>
<td>7,471,135</td>
<td>Contains the display names of any blind carbon copy (BCC) recipients of the original message. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_ORIGINAL_DISPLAY_CC_A</td>
<td>7,536,670</td>
<td>Contains the display names of any carbon copy (CC) recipients of the original message. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_ORIGINAL_DISPLAY_CC</td>
<td>7,536,671</td>
<td>Contains the display names of any carbon copy (CC) recipients of the original message. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ORIGINAL_DISPLAY_CC_W</td>
<td>7,536,671</td>
<td>Contains the display names of any carbon copy (CC) recipients of the original message. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_ORIGINAL_DISPLAY_TO_A</td>
<td>7,602,206</td>
<td>Contains the display names of the primary (To) recipients of the original message. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_ORIGINAL_DISPLAY_TO</td>
<td>7,602,207</td>
<td>Contains the display names of the primary (To) recipients of the original message. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ORIGINAL_DISPLAY_TO_W</td>
<td>7,602,207</td>
<td>Contains the display names of the primary (To) recipients of the original message. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_RECEIVED_BY_ADDRTYPE_A</td>
<td>7,667,742</td>
<td>Contains the e-mail address type, such as SMTP, for the messaging user that actually receives the message. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_RECEIVED_BY_ADDRTYPE</td>
<td>7,667,743</td>
<td>Contains the e-mail address type, such as SMTP, for the messaging user that actually receives the message. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_RECEIVED_BY_ADDRTYPE_W</td>
<td>7,667,743</td>
<td>Contains the e-mail address type, such as SMTP, for the messaging user that actually receives the message. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_RECEIVED_BY_EMAIL_ADDRESS_A</td>
<td>7,733,278</td>
<td>Contains the e-mail address for the messaging user that actually receives the message. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_RECEIVED_BY_EMAIL_ADDRESS</td>
<td>7,733,279</td>
<td>Contains the e-mail address for the messaging user that actually receives the message. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_RECEIVED_BY_EMAIL_ADDRESS_W</td>
<td>7,733,279</td>
<td>Contains the e-mail address for the messaging user that actually receives the message. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_RCVD_REPRESENTING_ADDRTYPE_A</td>
<td>7,798,814</td>
<td>Contains the address type for the messaging user represented by the user actually receiving the message. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_RCVD_REPRESENTING_ADDRTYPE</td>
<td>7,798,815</td>
<td>Contains the address type for the messaging user represented by the user actually receiving the message. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_RCVD_REPRESENTING_ADDRTYPE_W</td>
<td>7,798,815</td>
<td>Contains the address type for the messaging user represented by the user actually receiving the message. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_RCVD_REPRESENTING_EMAIL_ADDRESS_A</td>
<td>7,864,350</td>
<td>Contains the e-mail address for the messaging user represented by the receiving user. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_RCVD_REPRESENTING_EMAIL_ADDRESS</td>
<td>7,864,351</td>
<td>Contains the e-mail address for the messaging user represented by the receiving user. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_RCVD_REPRESENTING_EMAIL_ADDRESS_W</td>
<td>7,864,351</td>
<td>Contains the e-mail address for the messaging user represented by the receiving user. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_ORIGINAL_AUTHOR_ADDRTYPE_A</td>
<td>7,929,886</td>
<td>Contains the address type of the author of the first version of a message. That is — the message before being forwarded or replied to. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_ORIGINAL_AUTHOR_ADDRTYPE</td>
<td>7,929,887</td>
<td>Contains the address type of the author of the first version of a message. That is — the message before being forwarded or replied to. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ORIGINAL_AUTHOR_ADDRTYPE_W</td>
<td>7,929,887</td>
<td>Contains the address type of the author of the first version of a message. That is — the message before being forwarded or replied to. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_ORIGINAL_AUTHOR_EMAIL_ADDRESS_A</td>
<td>7,995,422</td>
<td>Contains the e-mail address of the author of the first version of a message. That is — the message before being forwarded or replied to. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_ORIGINAL_AUTHOR_EMAIL_ADDRESS</td>
<td>7,995,423</td>
<td>Contains the e-mail address of the author of the first version of a message. That is — the message before being forwarded or replied to. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ORIGINAL_AUTHOR_EMAIL_ADDRESS_W</td>
<td>7,995,423</td>
<td>Contains the e-mail address of the author of the first version of a message. That is — the message before being forwarded or replied to. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_ORIGINALLY_INTENDED_RECIP_ADDRTYPE_A</td>
<td>8,060,958</td>
<td>Contains the address type of the originally intended recipient of an autoforwarded message. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_ORIGINALLY_INTENDED_RECIP_ADDRTYPE</td>
<td>8,060,959</td>
<td>Contains the address type of the originally intended recipient of an autoforwarded message. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ORIGINALLY_INTENDED_RECIP_ADDRTYPE_W</td>
<td>8,060,959</td>
<td>Contains the address type of the originally intended recipient of an autoforwarded message. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS_A</td>
<td>8,126,494</td>
<td>Contains the e-mail address of the originally intended recipient of an autoforwarded message. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS</td>
<td>8,126,495</td>
<td>Contains the e-mail address of the originally intended recipient of an autoforwarded message. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS_W</td>
<td>8,126,495</td>
<td>Contains the e-mail address of the originally intended recipient of an autoforwarded message. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_TRANSPORT_MESSAGE_HEADERS_A</td>
<td>8,192,030</td>
<td>Contains transport-specific message envelope information. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_TRANSPORT_MESSAGE_HEADERS</td>
<td>8,192,031</td>
<td>Contains transport-specific message envelope information <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_TRANSPORT_MESSAGE_HEADERS_W</td>
<td>8,192,031</td>
<td>Contains transport-specific message envelope information. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_DELEGATION</td>
<td>8,257,794</td>
<td>Contains the converted value of the attDelegate workgroup property. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_TNEF_CORRELATION_KEY</td>
<td>8,323,330</td>
<td>Contains a value used to correlate a Transport Neutral Encapsulation Format (TNEF) attachment with a message. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_CONTENT_INTEGRITY_CHECK</td>
<td>201,326,850</td>
<td>Contains an ASN.1 content integrity check value that allows a message sender to help protect message content from disclosure to unauthorized recipients. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_EXPLICIT_CONVERSION</td>
<td>201,392,131</td>
<td>Contains a value that indicates that a message sender has requested a message content conversion for a particular recipient. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_IPM_RETURN_REQUESTED</td>
<td>201,457,675</td>
<td>Contains TRUE if this message should be returned with a report. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_MESSAGE_TOKEN</td>
<td>201,523,458</td>
<td>Contains an ASN.1 security token for a message. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_NDR_REASON_CODE</td>
<td>201,588,739</td>
<td>Contains an encoded reason for nondelivery that forms part of a nondelivery report. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_NDR_DIAG_CODE</td>
<td>201,654,275</td>
<td>Contains a diagnostic code that forms part of a nondelivery report. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_NON_RECEIPT_NOTIFICATION_REQUESTED</td>
<td>201,719,819</td>
<td>Contains TRUE if a message sender wants notification of non-receipt for a specified recipient. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_DELIVERY_POINT</td>
<td>201,785,347</td>
<td>Contains a value that specifies the nature of the functional entity by means of which a message was or would have been delivered to the recipient. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_ORIGINATOR_NON_DELIVERY_REPORT_REQUESTED</td>
<td>201,850,891</td>
<td>Contains TRUE if a message sender requests a nondelivery report for a particular recipient. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_ORIGINATOR_REQUESTED_ALTERNATE_RECIPIENT</td>
<td>201,916,674</td>
<td>Contains an entry identifier for an alternate recipient designated by the sender. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_PHYSICAL_DELIVERY_BUREAU_FAX_DELIVERY</td>
<td>201,981,963</td>
<td>Contains TRUE if the messaging system should use a fax bureau for physical delivery of this message. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_PHYSICAL_DELIVERY_MODE</td>
<td>202,047,491</td>
<td>Contains a bitmask of flags defining the physical delivery mode (for example, special delivery) for a message designated for a specific recipient. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_PHYSICAL_DELIVERY_REPORT_REQUEST</td>
<td>202,113,027</td>
<td>Contains the mode of a report to be delivered to a particular message recipient upon completion of physical message delivery or delivery by the message handling system. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_PHYSICAL_FORWARDING_ADDRESS</td>
<td>202,178,818</td>
<td>Contains the physical forwarding address of a message recipient and is used only with message reports. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_PHYSICAL_FORWARDING_ADDRESS_REQUESTED</td>
<td>202,244,107</td>
<td>Contains TRUE if a message sender requests the message transfer agent to attach a physical forwarding address for a message recipient. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_PHYSICAL_FORWARDING_PROHIBITED</td>
<td>202,309,643</td>
<td>Contains TRUE if a message sender prohibits physical message forwarding for a specific recipient. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_PHYSICAL_RENDITION_ATTRIBUTES</td>
<td>202,375,426</td>
<td>Contains an ASN.1 object identifier used for rendering message attachments. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_PROOF_OF_DELIVERY</td>
<td>202,440,962</td>
<td>Contains an ASN.1 proof of delivery value. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_PROOF_OF_DELIVERY_REQUESTED</td>
<td>202,506,251</td>
<td>Contains TRUE if a message sender requests proof of delivery for a particular recipient. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_RECIPIENT_CERTIFICATE</td>
<td>202,572,034</td>
<td>Contains a message recipient's ASN.1 certificate for use in a report. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_RECIPIENT_NUMBER_FOR_ADVICE_A</td>
<td>202,637,342</td>
<td>Contains a message recipient's voice telephone number to call to advise of the physical delivery of a message. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_RECIPIENT_NUMBER_FOR_ADVICE</td>
<td>202,637,343</td>
<td>Contains a message recipient's voice telephone number to call to advise of the physical delivery of a message. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_RECIPIENT_NUMBER_FOR_ADVICE_W</td>
<td>202,637,343</td>
<td>Contains a message recipient's voice telephone number to call to advise of the physical delivery of a message. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_RECIPIENT_TYPE</td>
<td>202,702,851</td>
<td>Contains the recipient type for a message recipient. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_REGISTERED_MAIL_TYPE</td>
<td>202,768,387</td>
<td>Contains the type of registration used for physical delivery of a message. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_REPLY_REQUESTED</td>
<td>202,833,931</td>
<td>Contains TRUE if a message sender requests a reply from a recipient. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_REQUESTED_DELIVERY_METHOD</td>
<td>202,899,459</td>
<td>Contains a binary array of delivery methods (service providers), in order of a message sender's preference. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_SENDER_ENTRYID</td>
<td>202,965,250</td>
<td>Contains the message sender's entry identifier. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_SENDER_NAME_A</td>
<td>203,030,558</td>
<td>Contains the message sender's display name. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_SENDER_NAME</td>
<td>203,030,559</td>
<td>Contains the message sender's display name. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_SENDER_NAME_W</td>
<td>203,030,559</td>
<td>Contains the message sender's display name. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_SUPPLEMENTARY_INFO_A</td>
<td>203,096,094</td>
<td>Contains additional information for use in a report. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_SUPPLEMENTARY_INFO</td>
<td>203,096,095</td>
<td>Contains additional information for use in a report. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_SUPPLEMENTARY_INFO_W</td>
<td>203,096,095</td>
<td>Contains additional information for use in a report. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_TYPE_OF_MTS_USER</td>
<td>203,161,603</td>
<td>Contains the type of a message recipient, for use in a report. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_SENDER_SEARCH_KEY</td>
<td>203,227,394</td>
<td>Contains the message sender's search key. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_SENDER_ADDRTYPE_A</td>
<td>203,292,702</td>
<td>Contains the message sender's e-mail address type. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_SENDER_ADDRTYPE</td>
<td>203,292,703</td>
<td>Contains the message sender's e-mail address type. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_SENDER_ADDRTYPE_W</td>
<td>203,292,703</td>
<td>Contains the message sender's e-mail address type. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_SENDER_EMAIL_ADDRESS_A</td>
<td>203,358,238</td>
<td>Contains the message sender's e-mail address. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_SENDER_EMAIL_ADDRESS</td>
<td>203,358,239</td>
<td>Contains the message sender's e-mail address. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_SENDER_EMAIL_ADDRESS_W</td>
<td>203,358,239</td>
<td>Contains the message sender's e-mail address. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_CURRENT_VERSION</td>
<td>234,881,044</td>
<td>Was originally meant to contain the current version of a message store. No longer supported. <br />Data Type: PT_I8</td></tr>
<tr>
<td>PR_DELETE_AFTER_SUBMIT</td>
<td>234,946,571</td>
<td>Contains TRUE if a client application wants MAPI to delete the associated message after submission. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_DISPLAY_BCC_A</td>
<td>235,012,126</td>
<td>Contains an ASCII list of the display names of any blind carbon copy (BCC) message recipients, separated by semicolons (;). Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_DISPLAY_BCC</td>
<td>235,012,127</td>
<td>Contains an ASCII list of the display names of any blind carbon copy (BCC) message recipients, separated by semicolons (;). <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_DISPLAY_BCC_W</td>
<td>235,012,127</td>
<td>Contains an ASCII list of the display names of any blind carbon copy (BCC) message recipients, separated by semicolons (;). UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_DISPLAY_CC_A</td>
<td>235,077,662</td>
<td>Contains an ASCII list of the display names of any carbon copy (CC) message recipients, separated by semicolons (;). Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_DISPLAY_CC</td>
<td>235,077,663</td>
<td>Contains an ASCII list of the display names of any carbon copy (CC) message recipients, separated by semicolons (;). <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_DISPLAY_CC_W</td>
<td>235,077,663</td>
<td>Contains an ASCII list of the display names of any carbon copy (CC) message recipients, separated by semicolons (;). UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_DISPLAY_TO_A</td>
<td>235,143,198</td>
<td>Contains an ASCII list of the display names of the primary (To) message recipients, separated by semicolons (;). Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_DISPLAY_TO</td>
<td>235,143,199</td>
<td>Contains an ASCII list of the display names of the primary (To) message recipients, separated by semicolons (;). <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_DISPLAY_TO_W</td>
<td>235,143,199</td>
<td>Contains an ASCII list of the display names of the primary (To) message recipients, separated by semicolons (;). UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_PARENT_DISPLAY_A</td>
<td>235,208,734</td>
<td>Contains the display name of the folder in which a message was found during a search. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_PARENT_DISPLAY</td>
<td>235,208,735</td>
<td>Contains the display name of the folder in which a message was found during a search. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_PARENT_DISPLAY_W</td>
<td>235,208,735</td>
<td>Contains the display name of the folder in which a message was found during a search. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_MESSAGE_DELIVERY_TIME</td>
<td>235,274,304</td>
<td>Contains the date and time a message was delivered. <br />Data Type: PT_SYSTIME</td></tr>
<tr>
<td>PR_MESSAGE_FLAGS</td>
<td>235,339,779</td>
<td>Contains a bitmask of flags indicating the origin and current state of a message. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_MESSAGE_SIZE</td>
<td>235,405,315</td>
<td>Contains the sum, in bytes, of the sizes of all properties on a message object. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_PARENT_ENTRYID</td>
<td>235,471,106</td>
<td>Contains the entry identifier of the folder containing a folder or message. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_SENTMAIL_ENTRYID</td>
<td>235,536,642</td>
<td>Contains the entry identifier of the folder where the message should be moved after submission. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_CORRELATE</td>
<td>235,667,467</td>
<td>Contains TRUE if the sender of a message requests the correlation feature of the messaging system. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_CORRELATE_MTSID</td>
<td>235,733,250</td>
<td>Contains the message transfer system (MTS) identifier used in correlating reports with sent messages. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_DISCRETE_VALUES</td>
<td>235,798,539</td>
<td>Contains TRUE if a nondelivery report applies only to discrete members of a distribution list rather than the entire list. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_RESPONSIBILITY</td>
<td>235,864,075</td>
<td>Contains TRUE if some transport provider has already accepted responsibility for delivering the message to this recipient, and FALSE if the MAPI spooler considers that this transport provider should accept responsibility. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_SPOOLER_STATUS</td>
<td>235,929,603</td>
<td>Contains the status of the message based on information available to the MAPI spooler. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_MESSAGE_RECIPIENTS</td>
<td>236,060,685</td>
<td>Contains a table of restrictions that can be applied to a contents table to find all messages that contain recipient subobjects meeting the restrictions. <br />Data Type: PT_OBJECT</td></tr>
<tr>
<td>PR_MESSAGE_ATTACHMENTS</td>
<td>236,126,221</td>
<td>Contains a table of restrictions that can be applied to a contents table to find all messages that contain attachment subobjects meeting the restrictions. <br />Data Type: PT_OBJECT</td></tr>
<tr>
<td>PR_SUBMIT_FLAGS</td>
<td>236,191,747</td>
<td>Contains a bitmask of flags indicating details about a message submission. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_RECIPIENT_STATUS</td>
<td>236,257,283</td>
<td>Contains a value used by the MAPI spooler in assigning delivery responsibility among transport providers. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_TRANSPORT_KEY</td>
<td>236,322,819</td>
<td>Contains a value used by the MAPI spooler to track the progress of an outbound message through the outgoing transport providers. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_MSG_STATUS</td>
<td>236,388,355</td>
<td>Contains a bitmask of property tags that define the status of a message. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_MESSAGE_DOWNLOAD_TIME</td>
<td>236,453,891</td>
<td>Contains the estimated time to download a message from a remote server to a local message store. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_HASATTACH</td>
<td>236,650,507</td>
<td>Contains at least one attachment. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_BODY_CRC</td>
<td>236,716,035</td>
<td>Contains a circular redundancy check (CRC) value on the message text. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_NORMALIZED_SUBJECT_A</td>
<td>236,781,598</td>
<td>Contains the message subject with any prefix removed. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_NORMALIZED_SUBJECT</td>
<td>236,781,599</td>
<td>ontains the message subject with any prefix removed. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_NORMALIZED_SUBJECT_W</td>
<td>236,781,599</td>
<td>Contains the message subject with any prefix removed. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_RTF_IN_SYNC</td>
<td>236,912,651</td>
<td>Contains TRUE if PR_RTF_COMPRESSED has the same text content as PR_BODY for this message. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_ATTACH_SIZE</td>
<td>236,978,179</td>
<td>Contains the sum, in bytes, of the sizes of all properties on an attachment. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_ATTACH_NUM</td>
<td>237,043,715</td>
<td>Contains a number that uniquely identifies the attachment within its parent message. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_PREPROCESS</td>
<td>237,109,259</td>
<td>Contains an ASN.1 proof of submission value. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_ACCESS</td>
<td>267,649,027</td>
<td>Contains a bitmask of flags indicating the operations a client application can perform on the open object. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_ROW_TYPE</td>
<td>267,714,563</td>
<td>Contains a value that indicates the type of a row in a table. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_INSTANCE_KEY</td>
<td>267,780,354</td>
<td>Contains a value that uniquely identifies a row in a table. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_ACCESS_LEVEL</td>
<td>267,845,635</td>
<td>Contains a bitmask of flags indicating the level at which a client application can access the open object. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_MAPPING_SIGNATURE</td>
<td>267,911,426</td>
<td>Contains the mapping signature for named properties of a particular MAPI object. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_RECORD_KEY</td>
<td>267,976,962</td>
<td>Contains a unique binary-comparable identifier for a specific object. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_STORE_RECORD_KEY</td>
<td>268,042,498</td>
<td>Contains the unique binary-comparable identifier (record key) of the message store in which an object resides. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_STORE_ENTRYID</td>
<td>268,108,034</td>
<td>Contains the unique entry identifier of the message store in which an object resides. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_MINI_ICON</td>
<td>268,173,570</td>
<td>Contains a bitmap of a half-size icon for a form. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_ICON</td>
<td>268,239,106</td>
<td>Contains a bitmap of a full size icon for a form. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_OBJECT_TYPE</td>
<td>268,304,387</td>
<td>Provides file type information for a non-Windows attachment. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_ENTRYID</td>
<td>268,370,178</td>
<td>Contains a MAPI entry identifier used to open and edit properties of a particular MAPI object. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_BODY_A</td>
<td>268,435,486</td>
<td>Contains the message text. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_BODY</td>
<td>268,435,487</td>
<td>Contains the message text. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_BODY_W</td>
<td>268,435,487</td>
<td>Contains the message text. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_REPORT_TEXT_A</td>
<td>268,501,022</td>
<td>Contains optional text for a report generated by the messaging system. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_REPORT_TEXT</td>
<td>268,501,023</td>
<td>Contains optional text for a report generated by the messaging system. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_REPORT_TEXT_W</td>
<td>268,501,023</td>
<td>Contains optional text for a report generated by the messaging system. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_ORIGINATOR_AND_DL_EXPANSION_HISTORY</td>
<td>268,566,786</td>
<td>Contains information about a message originator and a distribution list expansion history. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_REPORTING_DL_NAME</td>
<td>268,632,322</td>
<td>Contains the display name of a distribution list for which the messaging system is delivering a report. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_REPORTING_MTA_CERTIFICATE</td>
<td>268,697,858</td>
<td>Contains an identifier for the message transfer agent that generated a report. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_RTF_SYNC_BODY_CRC</td>
<td>268,828,675</td>
<td>Contains the cyclical redundancy check (CRC) computed for the message text. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_RTF_SYNC_BODY_COUNT</td>
<td>268,894,211</td>
<td>Contains a count of the significant characters of the message text. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_RTF_SYNC_BODY_TAG</td>
<td>268,959,775</td>
<td>Contains significant characters that appear at the beginning of the message text. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_RTF_COMPRESSED</td>
<td>269,025,538</td>
<td> </td></tr>
<tr>
<td>PR_RTF_SYNC_PREFIX_COUNT</td>
<td>269,484,035</td>
<td>Contains a count of the ignorable characters that appear before the significant characters of the message. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_RTF_SYNC_TRAILING_COUNT</td>
<td>269,549,571</td>
<td>Contains a count of the ignorable characters that appear after the significant characters of the message. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_BODY_HTML</td>
<td>269,680,671</td>
<td> </td></tr>
<tr>
<td>PR_ROWID</td>
<td>805,306,371</td>
<td>Contains a unique identifier for a recipient in a recipient table or status table. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_DISPLAY_NAME_A</td>
<td>805,371,934</td>
<td>Contains the display name of the message store. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_DISPLAY_NAME</td>
<td>805,371,935</td>
<td>Contains the display name of the message store. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_DISPLAY_NAME_W</td>
<td>805,371,935</td>
<td>Contains the display name of the message store. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_ADDRTYPE_A</td>
<td>805,437,470</td>
<td>Contains the messaging user's e-mail address type, such as Simple Mail Transfer Protocol (SMTP). Non-UNICODE compilation. <br />Data Type:PT.PT_STRING8</td></tr>
<tr>
<td>PR_ADDRTYPE</td>
<td>805,437,471</td>
<td>Contains the messaging user's e-mail address type, such as Simple Mail Transfer Protocol (SMTP). <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ADDRTYPE_W</td>
<td>805,437,471</td>
<td>Contains the messaging user's e-mail address type, such as Simple Mail Transfer Protocol (SMTP). UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_EMAIL_ADDRESS_A</td>
<td>805,503,006</td>
<td>Contains the messaging user's e-mail address. Non-UNICODE compilation. <br />Data Type: PT.PT_STRING8</td></tr>
<tr>
<td>PR_EMAIL_ADDRESS</td>
<td>805,503,007</td>
<td>Contains the messaging user's e-mail address. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_EMAIL_ADDRESS_W</td>
<td>805,503,007</td>
<td>Contains the messaging user's e-mail address. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_COMMENT_A</td>
<td>805,568,542</td>
<td>Contains a comment about the purpose or content of an object. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_COMMENT</td>
<td>805,568,543</td>
<td>Contains a comment about the purpose or content of an object. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_COMMENT_W</td>
<td>805,568,543</td>
<td>Contains a comment about the purpose or content of an object. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_DEPTH</td>
<td>805,634,051</td>
<td>Represents the relative level of indentation, or depth, of an object in a hierarchy table. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_PROVIDER_DISPLAY_A</td>
<td>805,699,614</td>
<td>Contains the vendor-defined display name for a service provider. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_PROVIDER_DISPLAY</td>
<td>805,699,615</td>
<td>Contains the vendor-defined display name for a service provider. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_PROVIDER_DISPLAY_W</td>
<td>805,699,615</td>
<td>Contains the vendor-defined display name for a service provider. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_CREATION_TIME</td>
<td>805,765,184</td>
<td>Contains the creation date and time for a message. <br />Data Type: PT_SYSTIME</td></tr>
<tr>
<td>PR_LAST_MODIFICATION_TIME</td>
<td>805,830,720</td>
<td>Contains the date and time the object or subobject was last modified. <br />Data Type: PT_SYSTIME</td></tr>
<tr>
<td>PR_RESOURCE_FLAGS</td>
<td>805,896,195</td>
<td>Contains a bitmask of flags for message services and providers. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_PROVIDER_DLL_NAME_A</td>
<td>805,961,758</td>
<td>Contains the base filename of the MAPI service provider DLL. Non-UNICODE compilation. <br />Data Type: PT_STRING8</td></tr>
<tr>
<td>PR_PROVIDER_DLL_NAME</td>
<td>805,961,759</td>
<td>Contains the base filename of the MAPI service provider DLL. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_PROVIDER_DLL_NAME_W</td>
<td>805,961,759</td>
<td>Contains the base filename of the MAPI service provider DLL. UNICODE compilation. <br />Data Type: PT_UNICODE</td></tr>
<tr>
<td>PR_SEARCH_KEY</td>
<td>806,027,522</td>
<td>Contains a binary-comparable key that identifies correlated objects for a search. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_PROVIDER_UID</td>
<td>806,093,058</td>
<td>Contains a MAPIUID structure of the service provider that is handling a message. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_PROVIDER_ORDINAL</td>
<td>806,158,339</td>
<td>Contains the zero-based index of a service provider's position in the provider table. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_DEFAULT_STORE</td>
<td>872,415,243</td>
<td>Contains TRUE if a message store is the default message store in the message store table. <br />Data Type: PT_BOOLEAN</td></tr>
<tr>
<td>PR_STORE_SUPPORT_MASK</td>
<td>873,267,203</td>
<td>Contains a bitmask of flags that client applications should query to determine the characteristics of a message store. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_STORE_STATE</td>
<td>873,332,739</td>
<td>Contains a flag that describes the state of the message store. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_IPM_SUBTREE_SEARCH_KEY</td>
<td>873,464,066</td>
<td>Was originally meant to contain the search key of the interpersonal message (IPM) root folder. No longer supported <br /> Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_IPM_OUTBOX_SEARCH_KEY</td>
<td>873,529,602</td>
<td>Was originally meant to contain the search key of the standard Outbox folder. No longer supported. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_IPM_WASTEBASKET_SEARCH_KEY</td>
<td>873,595,138</td>
<td>Was originally meant to contain the search key of the standard Deleted Items folder. No longer supported. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_IPM_SENTMAIL_SEARCH_KEY</td>
<td>873,660,674</td>
<td>Was originally meant to contain the search key of the standard Sent Items folder. No longer supported. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_MDB_PROVIDER</td>
<td>873,726,210</td>
<td>Contains a provider-defined MAPIUID structure that indicates the type of the message store. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_RECEIVE_FOLDER_SETTINGS</td>
<td>873,791,501</td>
<td>Contains a table of a message store's receive folder settings. <br />Data Type: PT_OBJECT</td></tr>
<tr>
<td>PR_VALID_FOLDER_MASK</td>
<td>903,806,979</td>
<td>Contains a bitmask of flags that indicate the validity of the entry identifiers of the folders in a message store. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_IPM_SUBTREE_ENTRYID</td>
<td>903,872,770</td>
<td>Contains the entry identifier of the root of the IPM folder subtree in the message store's folder tree. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_IPM_OUTBOX_ENTRYID</td>
<td>904,003,842</td>
<td>Contains the entry identifier of the standard interpersonal message (IPM) Outbox folder. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_IPM_WASTEBASKET_ENTRYID</td>
<td>904,069,378</td>
<td>Contains the entry identifier of the standard IPM Deleted Items folder. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_IPM_SENTMAIL_ENTRYID</td>
<td>904,134,914</td>
<td>Contains the entry identifier of the standard IPM Sent Items folder. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_VIEWS_ENTRYID</td>
<td>904,200,450</td>
<td>Contains the entry identifier of the user-defined Views folder. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_COMMON_VIEWS_ENTRYID</td>
<td>904,265,986</td>
<td>Contains the entry identifier of the predefined common view folder. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_FINDER_ENTRYID</td>
<td>904,331,522</td>
<td>Contains the entry identifier for the folder in which search results are typically created. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_CONTAINER_CLASS</td>
<td>907,214,879</td>
<td> </td></tr>
<tr>
<td>PR_IPM_APPOINTMENT_ENTRYID</td>
<td>919,601,410</td>
<td> </td></tr>
<tr>
<td>PR_IPM_CONTACT_ENTRYID</td>
<td>919,666,946</td>
<td> </td></tr>
<tr>
<td>PR_IPM_JOURNAL_ENTRYID</td>
<td>919,732,482</td>
<td> </td></tr>
<tr>
<td>PR_IPM_NOTE_ENTRYID</td>
<td>919,798,018</td>
<td> </td></tr>
<tr>
<td>PR_IPM_TASK_ENTRYID</td>
<td>919,863,554</td>
<td> </td></tr>
<tr>
<td>PR_IPM_DRAFTS_ENTRYID</td>
<td>920,060,162</td>
<td>Contains the EntryID of the Outlook Drafts folder. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_ATTACH_DATA_OBJ</td>
<td>922,812,429</td>
<td>Contains an attachment object typically accessed through the OLE IStorage:IUnknown interface. <br />Data Type: PT_OBJECT</td></tr>
<tr>
<td>PR_ATTACH_DATA_BIN</td>
<td>922,812,674</td>
<td>Contains binary attachment data typically accessed through the COM IStream:IUnknown interface. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_ATTACH_ENCODING</td>
<td>922,878,210</td>
<td>Contains an ASN.1 object identifier specifying the encoding for an attachment. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_ATTACH_EXTENSION</td>
<td>922,943,519</td>
<td>Contains a filename extension that indicates the document type of an attachment. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ATTACH_FILENAME</td>
<td>923,009,055</td>
<td>Contains an attachment's base filename and extension, excluding path. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ATTACH_METHOD</td>
<td>923,074,563</td>
<td>Contains a MAPI-defined constant representing the way the contents of an attachment can be accessed. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_ATTACH_LONG_FILENAME</td>
<td>923,205,663</td>
<td>Contains an attachment's long filename and extension, excluding path. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ATTACH_PATHNAME</td>
<td>923,271,199</td>
<td>Contains an attachment's fully qualified path and filename. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ATTACH_RENDERING</td>
<td>923,336,962</td>
<td>Contains a Microsoft Windows metafile with rendering information for an attachment. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_ATTACH_TAG</td>
<td>923,402,498</td>
<td>Contains an ASN.1 object identifier specifying the application that supplied an attachment. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_RENDERING_POSITION</td>
<td>923,467,779</td>
<td>Contains an offset, in characters, to use in rendering an attachment within the main message text. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_ATTACH_TRANSPORT_NAME</td>
<td>923,533,343</td>
<td>Contains the name of an attachment file modified so that it can be correlated with TNEF messages. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ATTACH_LONG_PATHNAME</td>
<td>923,598,879</td>
<td>Contains an attachment's fully qualified long path and filename. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ATTACH_MIME_TAG</td>
<td>923,664,415</td>
<td>Contains formatting information about a Multipurpose Internet Mail Extensions (MIME) attachment. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ATTACH_ADDITIONAL_INFO</td>
<td>923,730,178</td>
<td>Provides file type information for a non-Windows attachment. <br />Data Type: PT_BINARY</td></tr>
<tr>
<td>PR_SMTP_ADDRESS</td>
<td>972,947,487</td>
<td>Contains the SMTP address for the address book object. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_GENERATION</td>
<td>973,406,239</td>
<td>Contains a generational abbreviation that follows the full name of the recipient. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_GIVEN_NAME</td>
<td>973,471,775</td>
<td>Contains the display name for the messaging user represented by the sender. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_BUSINESS_TELEPHONE_NUMBER</td>
<td>973,602,847</td>
<td>Contains the primary telephone number of the recipient's place of business. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_HOME_TELEPHONE_NUMBER</td>
<td>973,668,383</td>
<td>Contains the primary telephone number of the recipient's home. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_SURNAME</td>
<td>974,192,671</td>
<td>Contains the last name of the messaging user. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_POSTAL_ADDRESS</td>
<td>974,454,815</td>
<td>Contains the postal address of the messaging user. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_COMPANY_NAME</td>
<td>974,520,351</td>
<td>Contains the recipient's company name. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_TITLE</td>
<td>974,585,887</td>
<td>Contains the job title of the messaging user. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_DEPARTMENT_NAME</td>
<td>974,651,423</td>
<td>Contains a name for the department in which the recipient works. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_OFFICE_LOCATION</td>
<td>974,716,959</td>
<td>Contains the recipient's office location. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_PRIMARY_TELEPHONE_NUMBER</td>
<td>974,782,495</td>
<td>Contains the recipient's primary telephone number. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_BUSINESS2_TELEPHONE_NUMBER</td>
<td>974,848,031</td>
<td>Contains a secondary telephone number at the recipient's place of business. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_MOBILE_TELEPHONE_NUMBER</td>
<td>974,913,567</td>
<td>Contains the recipient's mobile telephone number. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_PRIMARY_FAX_NUMBER</td>
<td>975,372,319</td>
<td>Contains the telephone number of the recipient's primary fax machine. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_BUSINESS_FAX_NUMBER</td>
<td>975,437,855</td>
<td>Contains the telephone number of the recipient's business fax machine. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_HOME_FAX_NUMBER</td>
<td>975,503,391</td>
<td>Contains the telephone number of the recipient's home fax machine. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_HOME2_TELEPHONE_NUMBER</td>
<td>976,158,751</td>
<td>Contains a secondary telephone number at the recipient's home. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_ASSISTANT</td>
<td>976,224,287</td>
<td>Contains the name of the recipient's administrative assistant. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_WEDDING_ANNIVERSARY</td>
<td>977,338,432</td>
<td>Contains the date of the recipient's wedding anniversary. <br />Data Type: PT_SYSTIME</td></tr>
<tr>
<td>PR_BIRTHDAY</td>
<td>977,403,968</td>
<td>Contains the date of the recipient's birthday. <br />Data Type: PT_SYSTIME</td></tr>
<tr>
<td>PR_MIDDLE_NAME</td>
<td>977,535,007</td>
<td>Contains the middle name of the messaging user. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_DISPLAY_NAME_PREFIX</td>
<td>977,600,543</td>
<td>Contains the display name prefix for a given MAPI object. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_PROFESSION</td>
<td>977,666,079</td>
<td>Contains the name of the messaging user's line of business <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_SPOUSE_NAME</td>
<td>977,797,151</td>
<td>Contains the name of the recipient's spouse/partner. <br />

Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_MANAGER_NAME</td>
<td>978,190,367</td>
<td>Contains the name of the recipient's manager. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_NICKNAME</td>
<td>978,255,903</td>
<td>Contains the recipient's nickname. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_PERSONAL_HOME_PAGE</td>
<td>978,321,439</td>
<td>Contains the Web address, also known as the Uniform Resource Locator (URL), of the messaging user's personal home page. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_BUSINESS_HOME_PAGE</td>
<td>978,386,975</td>
<td>Contains the Web address, also know as the Uniform Resource Locator (URL), of the business home page of the messaging user. <br />Data Type: PT_TSTRING</td></tr>
<tr>
<td>PR_RESOURCE_TYPE</td>
<td>1,040,384,003</td>
<td>Contains a value indicating the service provider type. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_STATUS_CODE</td>
<td>1,040,449,539</td>
<td>Contains a bitmask of flags indicating the current status of a session resource. All service providers set status codes as does MAPI to report on the status of the subsystem, the MAPI spooler, and the integrated address book. <br />Data Type: PT_LONG</td></tr>
<tr>
<td>PR_MSG_EDITOR_FORMAT</td>
<td>1,493,762,051</td>
<td>Specifies the format for an editor to use to display a message. <br />Data Type: PT_LONG</td></tr>
</table>

## See Also


#### Reference
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
