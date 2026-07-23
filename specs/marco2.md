[MS-OFFMACRO2]:
Office Macro-Enabled File Format Version 2
Intellectual Property Rights Notice for Open Specifications Documentation
 Technical Documentation. Microsoft publishes Open Specifications documentation (“this
documentation”) for protocols, file formats, data portability, computer languages, and standards
support. Additionally, overview documents cover inter-protocol relationships and interactions.
 Copyrights. This documentation is covered by Microsoft copyrights. Regardless of any other
terms that are contained in the terms of use for the Microsoft website that hosts this
documentation, you can make copies of it in order to develop implementations of the technologies
that are described in this documentation and can distribute portions of it in your implementations
that use these technologies or in your documentation as necessary to properly document the
implementation. You can also distribute in your implementation, with or without modification, any
schemas, IDLs, or code samples that are included in the documentation. This permission also
applies to any documents that are referenced in the Open Specifications documentation.
 No Trade Secrets. Microsoft does not claim any trade secret rights in this documentation.
 Patents. Microsoft has patents that might cover your implementations of the technologies
described in the Open Specifications documentation. Neither this notice nor Microsoft's delivery of
this documentation grants any licenses under those patents or any other Microsoft patents.
However, a given Open Specifications document might be covered by the Microsoft Open
Specifications Promise or the Microsoft Community Promise. If you would prefer a written license,
or if the technologies described in this documentation are not covered by the Open Specifications
Promise or Community Promise, as applicable, patent licenses are available by contacting
iplg@microsoft.com.
 License Programs. To see all of the protocols in scope under a specific license program and the
associated patents, visit the Patent Map.
 Trademarks. The names of companies and products contained in this documentation might be
covered by trademarks or similar intellectual property rights. This notice does not grant any
licenses under those rights. For a list of Microsoft trademarks, visit
www.microsoft.com/trademarks.
 Fictitious Names. The example companies, organizations, products, domain names, email
addresses, logos, people, places, and events that are depicted in this documentation are fictitious.
No association with any real company, organization, product, domain name, email address, logo,
person, place, or event is intended or should be inferred.
Reservation of Rights. All other rights are reserved, and this notice does not grant any rights other
than as specifically described above, whether by implication, estoppel, or otherwise.
Tools. The Open Specifications documentation does not require the use of Microsoft programming
tools or programming environments in order for you to develop an implementation. If you have access
to Microsoft programming tools and environments, you are free to take advantage of them. Certain
Open Specifications documents are intended for use in conjunction with publicly available standards
specifications and network programming art and, as such, assume that the reader either is familiar
with the aforementioned material or has immediate access to it.
Support. For questions and support, please contact dochelp@microsoft.com.
1 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

Revision Summary
|            | Revision  | Revision  |                       |
| ---------- | --------- | --------- | --------------------- |
| Date       | History   | Class     | Comments              |
| 7/13/2009  | 0.1       | Major     | Initial Availability  |
8/28/2009  0.2  Editorial  Revised and edited the technical content
11/6/2009  0.3  Editorial  Revised and edited the technical content
| 2/19/2010  | 1.0  | Minor  | Updated the technical content  |
| ---------- | ---- | ------ | ------------------------------ |
3/31/2010  1.01  Editorial  Revised and edited the technical content
4/30/2010  1.02  Editorial  Revised and edited the technical content
6/7/2010  1.03  Editorial  Revised and edited the technical content
6/29/2010  1.04  Editorial  Changed language and formatting in the technical content.
No changes to the meaning, language, or formatting of the
| 7/23/2010  | 1.04  | None  |     |
| ---------- | ----- | ----- | --- |
technical content.
No changes to the meaning, language, or formatting of the
| 9/27/2010  | 1.04  | None  |     |
| ---------- | ----- | ----- | --- |
technical content.
No changes to the meaning, language, or formatting of the
| 11/15/2010  | 1.04  | None  | technical content.  |
| ----------- | ----- | ----- | ------------------- |
No changes to the meaning, language, or formatting of the
| 12/17/2010  | 1.04  | None  |     |
| ----------- | ----- | ----- | --- |
technical content.
No changes to the meaning, language, or formatting of the
| 3/18/2011  | 1.04  | None  |     |
| ---------- | ----- | ----- | --- |
technical content.
No changes to the meaning, language, or formatting of the
| 6/10/2011  | 1.04  | None  |     |
| ---------- | ----- | ----- | --- |
technical content.
1/20/2012  1.5  Minor  Clarified the meaning of the technical content.
No changes to the meaning, language, or formatting of the
| 4/11/2012  | 1.5  | None  |     |
| ---------- | ---- | ----- | --- |
technical content.
7/16/2012  1.6  Minor  Clarified the meaning of the technical content.
10/8/2012  1.7  Minor  Clarified the meaning of the technical content.
No changes to the meaning, language, or formatting of the
| 2/11/2013  | 1.7  | None  |     |
| ---------- | ---- | ----- | --- |
technical content.
No changes to the meaning, language, or formatting of the
| 7/30/2013  | 1.7  | None  |     |
| ---------- | ---- | ----- | --- |
technical content.
No changes to the meaning, language, or formatting of the
| 11/18/2013  | 1.7  | None  |     |
| ----------- | ---- | ----- | --- |
technical content.
No changes to the meaning, language, or formatting of the
| 2/10/2014  | 1.7  | None  |     |
| ---------- | ---- | ----- | --- |
technical content.
4/30/2014  1.8  Minor  Clarified the meaning of the technical content.
No changes to the meaning, language, or formatting of the
| 7/31/2014  | 1.8  | None  |     |
| ---------- | ---- | ----- | --- |
technical content.
2 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

Revision Revision
Date History Class Comments
No changes to the meaning, language, or formatting of the
10/30/2014 1.8 None
technical content.
3/16/2015 2.0 Major Significantly changed the technical content.
9/4/2015 2.1 Minor Clarified the meaning of the technical content.
7/15/2016 2.2 Minor Clarified the meaning of the technical content.
No changes to the meaning, language, or formatting of the
9/14/2016 2.2 None
technical content.
12/15/2016 2.3 Minor Clarified the meaning of the technical content.
12/12/2017 2.4 Minor Clarified the meaning of the technical content.
4/27/2018 3.0 Major Significantly changed the technical content.
8/28/2018 4.0 Major Significantly changed the technical content.
4/22/2021 5.0 Major Significantly changed the technical content.
8/17/2021 6.0 Major Significantly changed the technical content.
4/16/2024 7.0 Major Significantly changed the technical content.
8/20/2024 8.0 Major Significantly changed the technical content.
3 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

Table of Contents
1 Introduction ............................................................................................................ 6
1.1 Glossary ........................................................................................................... 6
1.2 References ........................................................................................................ 7
1.2.1 Normative References ................................................................................... 7
1.2.2 Informative References ................................................................................. 8
1.3 Structure Overview ............................................................................................ 8
1.4 Relationship to Protocols and Other Structures ...................................................... 8
1.5 Applicability Statement ....................................................................................... 8
1.6 Versioning and Localization ................................................................................. 9
1.7 Vendor-Extensible Fields ..................................................................................... 9
2 Structures ............................................................................................................. 10
2.1 Macro-enabled WordprocessingML Structures ...................................................... 10
2.1.1 Part Enumeration ........................................................................................ 10
2.1.1.1 Main Document ..................................................................................... 10
2.1.1.2 VBA Project .......................................................................................... 10
2.1.1.3 VBA Supplemental Data ......................................................................... 11
2.1.2 Extensions ................................................................................................. 11
2.1.3 Global Elements ......................................................................................... 11
2.1.3.1 vbaSuppData ........................................................................................ 11
2.1.4 Global Attributes......................................................................................... 11
2.1.5 Complex Types ........................................................................................... 11
2.1.5.1 CT_DocEvents ...................................................................................... 11
2.1.5.2 CT_Mcd ............................................................................................... 12
2.1.5.3 CT_Mcds .............................................................................................. 13
2.1.5.4 CT_VbaSuppData .................................................................................. 13
2.1.6 Simple Types ............................................................................................. 14
2.2 Macro-enabled SpreadsheetML structures ........................................................... 14
2.2.1 Part Enumeration ........................................................................................ 14
2.2.1.1 International Macro Sheet ...................................................................... 14
2.2.1.2 Macro Sheet ......................................................................................... 15
2.2.1.3 VBA Project .......................................................................................... 15
2.2.1.4 Workbook ............................................................................................ 16
2.2.2 Excel Extensions ......................................................................................... 16
2.2.2.1 Extensions by Part ................................................................................ 16
2.2.2.1.1 extLst ............................................................................................ 16
2.2.3 Global Elements ......................................................................................... 16
2.2.3.1 macrosheet .......................................................................................... 16
2.2.3.2 definedNames ...................................................................................... 17
2.2.4 Global Attributes......................................................................................... 17
2.2.5 Complex Types ........................................................................................... 17
2.2.5.1 CT_Worksheet ...................................................................................... 17
2.2.5.2 CT_DefinedNames ................................................................................. 20
2.2.5.3 CT_DefinedName .................................................................................. 20
2.2.5.4 CT_DefinedNameArgumentDescriptions ................................................... 21
2.2.5.5 CT_DefinedNameArgumentDescription ..................................................... 21
2.2.6 Simple Types ............................................................................................. 22
2.2.7 Formulas ................................................................................................... 22
2.3 Macro-Enabled PresentationML Structures ........................................................... 22
2.3.1 Part Enumeration ........................................................................................ 22
2.3.1.1 Presentation ......................................................................................... 22
2.3.1.2 VBA Project .......................................................................................... 22
2.3.2 Extensions ................................................................................................. 23
2.3.3 Global Elements ......................................................................................... 23
2.3.4 Global Attributes......................................................................................... 23
4 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

2.3.5 Complex Types ........................................................................................... 23
2.3.5.1 CT_SlideMasterIdListEntry ..................................................................... 23
2.3.6 Simple Types ............................................................................................. 23
3 Structure Examples ............................................................................................... 24
3.1 VBA Supplemental Data Part ............................................................................. 24
3.2 Macro Sheet Part ............................................................................................. 24
4 Security Considerations ......................................................................................... 25
4.1 Security Considerations for Implementers ........................................................... 25
4.2 Index of Security Fields .................................................................................... 25
5 Appendix A: Full XML Schemas .............................................................................. 26
5.1 http://schemas.microsoft.com/office/word/2006/wordml Schema .......................... 26
5.2 http://schemas.microsoft.com/office/excel/2006/main Schema ............................. 26
5.3 http://schemas.microsoft.com/office/excel/2008/2/main Schema .......................... 27
6 Appendix B: Product Behavior ............................................................................... 29
7 Change Tracking .................................................................................................... 30
8 Index ..................................................................................................................... 31
5 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

1 Introduction
This document specifies the Office Macro-Enabled File Format Version 2. This file format specifies the
structures of macro-enabled WordprocessingML, SpreadsheetML, and PresentationML file formats that
are either not present in or modified from non-macro-enabled WordprocessingML, SpreadsheetML, and
PresentationML file formats described in [ISO/IEC29500-1:2016]. The macro-enabled file formats
consist of the non-macro-enabled file formats described in [ISO/IEC29500-1:2016] and the XML
structures, XML, and non-XML parts described in this document.
Sections 1.7 and 2 of this specification are normative. All other sections and examples in this
specification are informative.
1.1 Glossary
This document uses the following terms:
AutoFilter: A mechanism that can be used to filter tabular data based on user-defined criteria
such as values, strings, and formatting.
conditional formatting: A mechanism that changes the appearance of a user interface element
based on the evaluation of a rule or expression.
data consolidation: The process of combining tabular data from various worksheets into a single
list.
document: An object in a content database such as a file, folder, list, or site. Each object is
identified by a URI.
embedded object: An object that is created by using one application and is hosted in a document
that was created by using another application. Embedding an object, rather than inserting or
pasting it, ensures that the object retains its original format. Users can double-click an
embedded object and edit it with the toolbars and menus from the application that was used to
create it. See also Object Linking and Embedding (OLE).
Hypertext Markup Language (HTML): An application of the Standard Generalized Markup
Language (SGML) that uses tags to mark elements in a document, as described in [HTML].
locale: A collection of rules and data that are specific to a language and a geographical area. A
locale can include information about sorting rules, date and time formatting, numeric and
monetary conventions, and character classification.
macro: A set of instructions that are recorded or written, and then typically saved to a file. When a
macro is run, all of the instructions are performed automatically.
macro sheet: A single, logical container that is used to store and run Microsoft Excel 4.0 for
Windows macro formulas.
page break: A divider that breaks a worksheet into separate pages for printing. Page breaks are
inserted automatically based on the paper size, margin settings, scaling options, and the
positions of any page breaks that are inserted manually.
page margin: The white space at the top, bottom, and sides of a document when a document is
printed.
phonetic information: A series of characters that appear above text in a cell and provide
information that helps users pronounce the text.
presentation: A collection of slides that are intended to be viewed by an audience.
6 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

protection: A mechanism that helps restrict users from making unwanted changes to the data or
structure of a workbook.
range: An addressable region that is in a workbook. A range typically consists of zero or more cells
and represents a single, contiguous rectangle of cells on a single sheet.
sheet: A worksheet. The term sheet frequently refers to a worksheet because worksheets are the
most common type of sheet.
sheet view: A collection of display settings, such as which cells are shown, and the zoom level for
a sheet window.
sort condition: A condition that determines how to sort cells in a range.
table: A list that is defined in a workbook.
Vector Markup Language (VML): A system of marking up or tagging two-dimensional vector
graphics for publication on the World Wide Web. VML graphics are scalable and editable, and
typically require less disk space and less time to download.
Visual Basic for Applications (VBA): A macro-based programming language that derives from
Microsoft Visual Basic and can be used to customize and extend an application. Unlike Visual
Basic, Microsoft Visual Basic for Applications (VBA) code and macros can be run only from within
a host application that supports VBA.
workbook: A container for a collection of sheets.
worksheet: A single logical container for a set of tabular data and other objects in a workbook.
XML: The Extensible Markup Language, as described in [XML1.0].
MAY, SHOULD, MUST, SHOULD NOT, MUST NOT: These terms (in all caps) are used as defined
in [RFC2119]. All statements of optional behavior use either MAY, SHOULD, or SHOULD NOT.
1.2 References
Links to a document in the Microsoft Open Specifications library point to the correct section in the
most recently published version of the referenced document. However, because individual documents
in the library are not updated at the same time, the section numbers in the documents may not
match. You can confirm the correct section numbering by checking the Errata.
1.2.1 Normative References
We conduct frequent surveys of the normative references to assure their continued availability. If you
have any issue with finding a normative reference, please contact dochelp@microsoft.com. We will
assist you in finding the relevant information.
[ISO/IEC29500-1:2016] ISO/IEC, "Information technology -- Document description and processing
languages -- Office Open XML File Formats -- Part 1: Fundamentals and Markup Language Reference",
ISO/IEC 29500-1:2016, https://www.iso.org/standard/71691.html
[ISO/IEC29500-3:2015] ISO/IEC, "Information technology -- Document description and processing
languages -- Office Open XML File Formats -- Part 3: Markup Compatibility and Extensibility",
https://www.iso.org/standard/65533.html
[ISO/IEC29500-4:2016] ISO/IEC, "Information technology -- Document description and processing
languages -- Office Open XML File Formats -- Part 4: Transitional Migration Features",
https://www.iso.org/standard/71692.html
7 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

[MS-OVBA] Microsoft Corporation, "Office VBA File Format Structure".
[MS-XLSB] Microsoft Corporation, "Excel (.xlsb) Binary File Format".
[MS-XLSX] Microsoft Corporation, "Excel (.xlsx) Extensions to the Office Open XML SpreadsheetML File
Format".
[RFC2119] Bradner, S., "Key words for use in RFCs to Indicate Requirement Levels", BCP 14, RFC
2119, March 1997, https://www.rfc-editor.org/info/rfc2119
[XMLSCHEMA1/2] Thompson, H., Beech, D., Maloney, M., and Mendelsohn, N., Eds., "XML Schema
Part 1: Structures Second Edition", W3C Recommendation, October 2004,
https://www.w3.org/TR/2004/REC-xmlschema-1-20041028/
[XMLSCHEMA2/2] Biron, P., and Malhotra, A., Eds., "XML Schema Part 2: Datatypes Second Edition",
W3C Recommendation, October 2004, https://www.w3.org/TR/2004/REC-xmlschema-2-20041028/
1.2.2 Informative References
[ISO/IEC29500-2:2012] ISO/IEC, "Information technology -- Document description and processing
languages -- Office Open XML File Formats -- Part 2: Open Packaging Conventions", ISO/IEC 29500-
2:2012, http://www.iso.org/iso/home/store/catalogue_ics/catalogue_detail_ics.htm?csnumber=61796
1.3 Structure Overview
This file format specifies structures of macro-enabled WordprocessingML, SpreadsheetML, and
PresentationML file formats that are either not present in or modified from non-macro-enabled
WordprocessingML, SpreadsheetML, and PresentationML file formats described in [ISO/IEC29500-
1:2016]. Document, Workbook, and Presentation data are each contained in a compressed file
(.zip) package conforming to the Open Packaging Conventions as described in [ISO/IEC29500-
2:2012]. Individual files stored in a ZIP package, called parts ([ISO/IEC29500-1:2016]), contain
information about the content and structure of a document, workbook, or presentation. Some parts
store information using XML and other parts store information by using binary data.
1.4 Relationship to Protocols and Other Structures
The structures specified in section 2 are dependent on the structures and concepts defined in the
following references:
 [ISO/IEC29500-1:2016] for baseline WordprocessingML, SpreadsheetML, and PresentationML
persistence formats.
 [ISO/IEC29500-2:2012] for open packaging conventions.
 [ISO/IEC29500-3:2015] for markup compatibility and extensibility.
 [ISO/IEC29500-4:2016] for backward compatibility considerations.
1.5 Applicability Statement
This document specifies a persistence format for structures either not present in or modified from the
structures described in [ISO/IEC29500-1:2016] for WordprocessingML, SpreadsheetML, and
PresentationML documents. The structures specified in this document are not applicable as a stand-
alone file format. Each structure specified in this document is integrated with [ISO/IEC29500-1:2016]
WordprocessingML, SpreadsheetML, or PresentationML documents in a particular way, as specified in
the section for that structure.
8 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

1.6 Versioning and Localization
None.
1.7 Vendor-Extensible Fields
None.
9 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

2 Structures
All global elements, global attributes, complex types, and simple types not defined in [ISO/IEC29500-
1:2016] and not defined in the following target namespaces are extensions that use conventions
specified in [ISO/IEC29500-3:2015]:
 http://schemas.microsoft.com/office/word/2006/wordml
 http://schemas.microsoft.com/office/excel/2006/main
 http://schemas.openxmlformats.org/presentationml/2006/main
 http://schemas.microsoft.com/office/excel/2008/2/main
2.1 Macro-enabled WordprocessingML Structures
2.1.1 Part Enumeration
This section specifies the parts ([ISO/IEC29500-1:2016]) of the macro-enabled WordprocessingML file
format that are different from the parts of the non-macro-enabled WordprocessingML file format
specified in [ISO/IEC29500-1:2016] section 11.3. These parts and their relationships are summarized
in the following table.
Part Relationship target
Main Document (section 2.1.1.1) WordprocessingML Package ([ISO/IEC29500-1:2016]
section 11.2)
VBA Project (section 2.1.1.2) Main Document (section 2.1.1.1)
VBA Supplemental Data (section 2.1.1.3) VBA Project (section 2.1.1.2)
2.1.1.1 Main Document
Content Types:
 application/vnd.ms-word.document.macroEnabled.main+xml
 application/vnd.ms-word.template.macroEnabledTemplate.main+xml
Root Namespace: http://schemas.openxmlformats.org/wordprocessingml/2006/main
Source Relationship:
http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument
The part is modified from the Main Document specified in [ISO/IEC29500-1:2016], section 11.3.10.
A Main Document part is permitted to have an implicit relationship to the VBA Project part.
2.1.1.2 VBA Project
Content Type: application/vnd.ms-office.vbaProject
Root Namespace: Not applicable
Source Relationship: http://schemas.microsoft.com/office/2006/relationships/vbaProject
10 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

This part is specified in [MS-OVBA].
A package MUST contain at most one VBA Project part, which MUST be the target of an implicit
relationship from the Main Document part (section 2.1.1.1).
A VBA Project part MUST have an implicit relationship to the VBA Supplemental Data part (section
2.1.1.3).
A VBA Project part MUST NOT have implicit or explicit relationships to any other part specified by this
document or any part specified in [ISO/IEC29500-1:2016].
2.1.1.3 VBA Supplemental Data
Content Type: application/vnd.ms-word.vbaData+xml
Root Namespace: http://schemas.microsoft.com/office/word/2006/wordml
Source Relationship: http://schemas.microsoft.com/office/2006/relationships/wordVbaData
An instance of this part type specifies Visual Basic for Applications (VBA) supplemental Data.
A VBA Supplemental Data part MUST be the target of an implicit relationship from the VBA Project
part (section 2.1.1.2).
A VBA Supplemental Data part MUST NOT have implicit or explicit relationships to any other part
specified by this document or any part specified in [ISO/IEC29500-1:2016].
The root element for a part of this content type MUST be vbaSuppData (section 2.1.3.1).
2.1.2 Extensions
None.
2.1.3 Global Elements
2.1.3.1 vbaSuppData
Target namespace: http://schemas.microsoft.com/office/word/2006/wordml
A CT_VbaSuppData element (section 2.1.5.4) that specifies supplemental Visual Basic for
Applications (VBA) data.
The following W3C XML Schema ([XMLSCHEMA1/2] section 2.1) fragment defines the contents of this
element.
<xsd:element name="vbaSuppData" type="CT_VbaSuppData"/>
2.1.4 Global Attributes
None.
2.1.5 Complex Types
2.1.5.1 CT_DocEvents
Target namespace: http://schemas.microsoft.com/office/word/2006/wordml
11 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

Referenced by: CT_VbaSuppData
A complex type that specifies a container of document events that are active for this document.
Child Elements:
eventDocNew: An element that specifies that the New event is active.
eventDocOpen: An element that specifies that the Open event is active.
eventDocClose: An element that specifies that the Close event is active.
eventDocSync: An element that specifies that the Sync event is active.
eventDocXmlAfterInsert: An element that specifies that the XmlAfterInsert event is active.
eventDocXmlBeforeDelete: An element that specifies that the XmlBeforeDelete event is active.
eventDocContentControlAfterInsert: An element that specifies that the
ContentControlAfterInsert event is active.
eventDocContentControlBeforeDelete: An element that specifies that the
ContentControlBeforeDelete event is active.
eventDocContentControlOnExit: An element that specifies that the ContentControlOnExit event
is active.
eventDocContentControlOnEnter: An element that specifies that the ContentControlOnEnter
event is active.
eventDocStoreUpdate: An element that specifies that the StoreUpdate event is active.
eventDocContentControlContentUpdate: An element that specifies that the
ContentControlContentUpdate event is active.
eventDocBuildingBlockAfterInsert: An element that specifies that the BuildingBlockAfterInsert
event is active.
The following W3C XML Schema ([XMLSCHEMA1/2] section 2.1) fragment defines the contents of this
complex type.
<xsd:complexType name="CT_DocEvents">
<xsd:sequence>
<xsd:element name="eventDocNew" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocOpen" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocClose" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocSync" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocXmlAfterInsert" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocXmlBeforeDelete" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocContentControlAfterInsert" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocContentControlBeforeDelete" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocContentControlOnExit" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocContentControlOnEnter" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocStoreUpdate" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocContentControlContentUpdate" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocBuildingBlockAfterInsert" minOccurs="0" maxOccurs="1"/>
</xsd:sequence>
</xsd:complexType>
2.1.5.2 CT_Mcd
Target namespace: http://schemas.microsoft.com/office/word/2006/wordml
12 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

Referenced by: CT_Mcds
A complex type that specifies a macro.
Attributes:
macroName: An ST_String attribute ([ISO/IEC29500-1:2016] section 22.9.2.13) that MUST equal
the name attribute with every character set to uppercase.
name: An ST_String attribute ([ISO/IEC29500-1:2016] section 22.9.2.13) that specifies the name of
the macro. This attribute MUST NOT exceed 255 characters.
menuHelp: An ST_String attribute ([ISO/IEC29500-1:2016] section 22.9.2.13) that MUST be
ignored.
bEncrypt: An ST_UcharHexNumber attribute ([ISO/IEC29500-1:2016] section 17.18.98) that
MUST be 0.
cmg: An ST_UcharHexNumber attribute ([ISO/IEC29500-1:2016] section 17.18.98) that MUST be
56.
The following W3C XML Schema ([XMLSCHEMA1/2] section 2.1) fragment defines the contents of this
complex type.
<xsd:complexType name="CT_Mcd">
<xsd:attribute name="macroName" type="w:ST_String" use="optional"/>
<xsd:attribute name="name" type="w:ST_String" use="optional"/>
<xsd:attribute name="menuHelp" type="w:ST_String" use="optional"/>
<xsd:attribute name="bEncrypt" type="w:ST_UcharHexNumber"/>
<xsd:attribute name="cmg" type="w:ST_UcharHexNumber"/>
</xsd:complexType>
2.1.5.3 CT_Mcds
Target namespace: http://schemas.microsoft.com/office/word/2006/wordml
Referenced by: CT_VbaSuppData
A complex type that specifies a collection of macros.
Child Elements:
mcd: A CT_Mcd element (section 2.1.5.2) that specifies a macro.
The following W3C XML Schema ([XMLSCHEMA1/2] section 2.1) fragment defines the contents of this
complex type.
<xsd:complexType name="CT_Mcds">
<xsd:sequence>
<xsd:element name="mcd" type="CT_Mcd" minOccurs="0" maxOccurs="unbounded"/>
</xsd:sequence>
</xsd:complexType>
2.1.5.4 CT_VbaSuppData
Target namespace: http://schemas.microsoft.com/office/word/2006/wordml
Referenced by: vbaSuppData
13 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

A complex type that specifies document events and macro information for the VBA project.
Child Elements:
docEvents: A CT_DocEvents element (section 2.1.5.1) that specifies a container of document
events that are active for this document.
mcds: A CT_Mcds element (section 2.1.5.3) that specifies a collection of macros stored in the VBA
project.
The following W3C XML Schema ([XMLSCHEMA1/2] section 2.1) fragment defines the contents of this
complex type.
<xsd:complexType name="CT_VbaSuppData">
<xsd:sequence>
<xsd:element name="docEvents" type="CT_DocEvents" minOccurs="0" maxOccurs="1"/>
<xsd:element name="mcds" type="CT_Mcds" minOccurs="0" maxOccurs="1"/>
</xsd:sequence>
</xsd:complexType>
2.1.6 Simple Types
None.
2.2 Macro-enabled SpreadsheetML structures
2.2.1 Part Enumeration
This section specifies the parts ([ISO/IEC29500-1:2016]) of the macro-enabled SpreadsheetML file
format that are different from the parts of the non-macro-enabled SpreadsheetML file format specified
in [ISO/IEC29500-1:2016] section 12.3. These parts and their relationships are summarized in the
following table.
Part Relationship target
International Macro Sheet (section 2.2.1.1) Workbook (section 2.2.1.4)
Macro Sheet (section 2.2.1.2) Workbook (section 2.2.1.4)
VBA Project (section 2.2.1.3) Workbook (section 2.2.1.4)
Workbook (section 2.2.1.4) SpreadsheetML Package ([ISO/IEC29500-1:2016]
section 12.2)
2.2.1.1 International Macro Sheet
Content Type: application/vnd.ms-excel.intlmacrosheet+xml
Root Namespace: http://schemas.microsoft.com/office/excel/2006/main
Source Relationship: http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet
An instance of this part type specifies the same information as a Macro Sheet part type (section
2.2.1.2). When this part type is present, the associated macro sheet will display and operate with the
US English locale, regardless of what the product locale, user locale, or system locale is. This will
maximize compatibility when running under various locales.
14 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

2.2.1.2 Macro Sheet
Content Type: application/vnd.ms-excel.macrosheet+xml
Root Namespace: http://schemas.microsoft.com/office/excel/2006/main
Source Relationship: http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet
An instance of this part type specifies a macro sheet. References to the macro sheet part in this
document include international macro sheets (section 2.2.1.1) as well.
A macro sheet part MUST be the target of an explicit relationship from the Workbook part (section
2.2.1.4).
The root element for a part of this content type MUST be macrosheet (section 2.2.3.1).
A macro sheet part is permitted to have implicit relationships to the following parts specified in
[ISO/IEC29500-1:2016]:
 Comments
A macro sheet part is permitted to have implicit relationships to the following parts specified in [MS-
XLSB]:
 Macro Sheet Binary Index
 Sort Map
A macro sheet part is permitted to have explicit relationships to the following parts specified in
[ISO/IEC29500-1:2016]:
 Custom Property
 Drawings
 OLE Object
 OLE Package
 Images
 Printer Settings
 VML Drawing
A macro sheet part MUST NOT have implicit or explicit relationships to any other part specified in
[ISO/IEC29500-1:2016].
A formula ([ISO/IEC29500-1:2016] section 18.17) in a macro sheet is specified by Formulas (section
2.2.7), which is an extension of the definition in [MS-XLSX] section 2.2.2.
2.2.1.3 VBA Project
Content Type: application/vnd.ms-office.vbaProject
Root Namespace: Not applicable
Source Relationship: http://schemas.microsoft.com/office/2006/relationships/vbaProject
This part is specified in [MS-OVBA].
15 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

A package MUST contain at most one VBA Project part, which MUST be the target of an implicit
relationship from the Workbook part (section 2.2.1.4).
A VBA Project part MUST NOT have implicit or explicit relationships to any other part specified by this
document or any other part specified in [ISO/IEC29500-1:2016].
2.2.1.4 Workbook
Content Types:
 application/vnd.ms-excel.sheet.macroEnabled.main+xml
 application/vnd.ms-excel.template.macroEnabled.main+xml
Root Namespace: http://schemas.openxmlformats.org/spreadsheetml/2006/main
Source Relationship:
http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument
This part is modified from the part specified in [ISO/IEC29500-1:2016] Part1, section 12.3.23.
A Workbook part is permitted to have an implicit relationship to the VBA Project (section 2.2.1.3),
Macro Sheet (section 2.2.1.2), and International Macro Sheet (section 2.2.1.1) parts.
2.2.2 Excel Extensions
Extensions are specified by using conventions provided by [ISO/IEC29500-3:2015].
2.2.2.1 Extensions by Part
This section specifies the elements from [ISO/IEC29500-1:2016] that are extended by this format.
Either the Ignorable attribute ([ISO/IEC29500-3:2015] section 7.2), AlternateContent element
([ISO/IEC29500-3:2015] section 7.5), or extLst element ([ISO/IEC29500-1:2016] section 18.2.10)
MUST be used to maintain compatibility with ISO/IEC 29500 implementations.
2.2.2.1.1 extLst
The extLst ([ISO/IEC29500-1:2016] section 18.2.10) child element of the workbook ([ISO/IEC29500-
1:2016] section 18.2.27) element is extended by the addition of new child ext ([ISO/IEC29500-
1:2016] section 18.2.7) elements, whose structure is specified in the following table.
Ext uri Child element
{8193D73E-4604-4355-AFF4-8A2333D13BE4} definedNames (section 2.2.3.2)
For more information about how extension lists are used, see [ISO/IEC29500-1:2016] section
18.2.10.
2.2.3 Global Elements
2.2.3.1 macrosheet
Target namespace: http://schemas.microsoft.com/office/excel/2006/main
A CT_Worksheet element (section 2.2.5.1) that specifies a macro sheet.
16 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

The following W3C XML Schema ([XMLSCHEMA1/2] section 2.1) fragment defines the contents of this
element.
<xsd:element name="macrosheet" type="CT_Worksheet"/>
2.2.3.2 definedNames
Target namespace: http://schemas.microsoft.com/office/excel/2008/2/main
A CT_DefinedNames element (section 2.2.5.2) that specifies additional properties for macros.
The following W3C XML Schema ([XMLSCHEMA1/2] section 2.1) fragment defines the contents of this
element.
<xsd:element name="definedNames" type="CT_DefinedNames"/>
2.2.4 Global Attributes
None.
2.2.5 Complex Types
2.2.5.1 CT_Worksheet
Target namespace: http://schemas.microsoft.com/office/excel/2006/main
Referenced by: macrosheet
A complex type that specifies properties for a macro sheet.
Child Elements:
sheetPr: A CT_SheetPr element ([ISO/IEC29500-1:2016] section A.2) that specifies sheet-level
properties.
dimension: A CT_SheetDimension element ([ISO/IEC29500-1:2016] section A.2) that specifies the
used range of the worksheet.
sheetViews: A CT_SheetViews element ([ISO/IEC29500-1:2016] section A.2) that specifies sheet
view information.
sheetFormatPr: A CT_SheetFormatPr element ([ISO/IEC29500-1:2016] section A.2) that specifies
sheet formatting properties.
cols: A CT_Cols element ([ISO/IEC29500-1:2016] section A.2) that specifies information about whole
columns of the worksheet.
sheetData: A CT_SheetData element ([ISO/IEC29500-1:2016] section A.2) that specifies the
Worksheet ([ISO/IEC29500-1:2016] section 18.3) cell table.
sheetCalcPr: This element MUST be ignored in a macro sheet.
sheetProtection: A CT_SheetProtection element ([ISO/IEC29500-1:2016] section A.2) that
specifies sheet protection options.
protectedRanges: This element MUST be ignored in a macro sheet.
17 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

scenarios: This element MUST be ignored in a macro sheet.
autoFilter: A CT_AutoFilter element ([ISO/IEC29500-1:2016] section A.2) that specifies AutoFilter
settings.
sortState: A CT_SortState element ([ISO/IEC29500-1:2016] section A.2) that specifies the different
sort conditions that apply to a range.
dataConsolidate: A CT_DataConsolidate element ([ISO/IEC29500-1:2016] section A.2) that
specifies data consolidation settings.
customSheetViews: A CT_CustomSheetViews element ([ISO/IEC29500-1:2016] section A.2) that
specifies custom sheet views.
mergeCells: This element MUST be ignored in a macro sheet.
phoneticPr: A CT_PhoneticPr element ([ISO/IEC29500-1:2016] section A.2) that specifies default
phonetic information.
conditionalFormatting: A CT_ConditionalFormatting element ([ISO/IEC29500-1:2016] section
A.2) that specifies conditional formatting rules.
dataValidations: This element MUST be ignored in a macro sheet.
hyperlinks: This element MUST be ignored in a macro sheet.
printOptions: A CT_PrintOptions element ([ISO/IEC29500-1:2016] section A.2) that specifies print
options for the sheet.
pageMargins: A CT_PageMargins element ([ISO/IEC29500-1:2016] section A.2) that specifies
page margins for the sheet.
pageSetup: A CT_PageSetup element ([ISO/IEC29500-1:2016] section A.2) that specifies page
setup settings for the sheet.
headerFooter: A CT_HeaderFooter element ([ISO/IEC29500-1:2016] section A.2) that specifies
header and footer settings.
rowBreaks: A CT_PageBreak element ([ISO/IEC29500-1:2016] section A.2) that specifies
horizontal page break information.
colBreaks: A CT_PageBreak element ([ISO/IEC29500-1:2016] section A.2) that specifies vertical
page break information.
customProperties: A CT_CustomProperties element ([ISO/IEC29500-1:2016] section A.2) that
specifies references to binary parts containing arbitrary user-defined data.
cellWatches: This element MUST be ignored in a macro sheet.
ignoredErrors: This element MUST be ignored in a macro sheet.
smartTags: This element MUST be ignored in a macro sheet.
drawing: A CT_Drawing element ([ISO/IEC29500-1:2016] section A.2) that specifies a reference to
a drawing part ([ISO/IEC29500-1:2016] section 12.3.8) containing one or more drawing
([ISO/IEC29500-1:2016] section 20.5) elements.
picture: A CT_SheetBackgroundPicture element ([ISO/IEC29500-1:2016] section A.2) that
specifies a background image of a sheet.
oleObjects: A CT_OleObjects element ([ISO/IEC29500-1:2016] section A.2) that specifies
embedded objects in the sheet.
18 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

controls: This element MUST be ignored in a macro sheet.
webPublishItems: A CT_WebPublishItems element ([ISO/IEC29500-1:2016] section A.2) that
specifies individual objects that have been published to HTML.
tableParts: A CT_TableParts element ([ISO/IEC29500-1:2016] section A.2) that specifies
references to parts specifying every table on the sheet.
extLst: A CT_ExtensionList element ([ISO/IEC29500-1:2016] section A.2) that specifies future
extensions to this format.
The following W3C XML Schema ([XMLSCHEMA1/2] section 2.1) fragment defines the contents of this
complex type.
<xsd:complexType name="CT_Worksheet">
<xsd:sequence>
<xsd:element name="sheetPr" type="x:CT_SheetPr" minOccurs="0" maxOccurs="1"/>
<xsd:element name="dimension" type="x:CT_SheetDimension" minOccurs="0" maxOccurs="1"/>
<xsd:element name="sheetViews" type="x:CT_SheetViews" minOccurs="0" maxOccurs="1"/>
<xsd:element name="sheetFormatPr" type="x:CT_SheetFormatPr" minOccurs="0" maxOccurs="1"/>
<xsd:element name="cols" type="x:CT_Cols" minOccurs="0" maxOccurs="unbounded"/>
<xsd:element name="sheetData" type="x:CT_SheetData" minOccurs="1" maxOccurs="1"/>
<xsd:element name="sheetCalcPr" type="x:CT_SheetCalcPr" minOccurs="0" maxOccurs="1"/>
<xsd:element name="sheetProtection" type="x:CT_SheetProtection" minOccurs="0"
maxOccurs="1"/>
<xsd:element name="protectedRanges" type="x:CT_ProtectedRanges" minOccurs="0"
maxOccurs="1"/>
<xsd:element name="scenarios" type="x:CT_Scenarios" minOccurs="0" maxOccurs="1"/>
<xsd:element name="autoFilter" type="x:CT_AutoFilter" minOccurs="0" maxOccurs="1"/>
<xsd:element name="sortState" type="x:CT_SortState" minOccurs="0" maxOccurs="1"/>
<xsd:element name="dataConsolidate" type="x:CT_DataConsolidate" minOccurs="0"
maxOccurs="1"/>
<xsd:element name="customSheetViews" type="x:CT_CustomSheetViews" minOccurs="0"
maxOccurs="1"/>
<xsd:element name="mergeCells" type="x:CT_MergeCells" minOccurs="0" maxOccurs="1"/>
<xsd:element name="phoneticPr" type="x:CT_PhoneticPr" minOccurs="0" maxOccurs="1"/>
<xsd:element name="conditionalFormatting" type="x:CT_ConditionalFormatting" minOccurs="0"
maxOccurs="unbounded"/>
<xsd:element name="dataValidations" type="x:CT_DataValidations" minOccurs="0"
maxOccurs="1"/>
<xsd:element name="hyperlinks" type="x:CT_Hyperlinks" minOccurs="0" maxOccurs="1"/>
<xsd:element name="printOptions" type="x:CT_PrintOptions" minOccurs="0" maxOccurs="1"/>
<xsd:element name="pageMargins" type="x:CT_PageMargins" minOccurs="0" maxOccurs="1"/>
<xsd:element name="pageSetup" type="x:CT_PageSetup" minOccurs="0" maxOccurs="1"/>
<xsd:element name="headerFooter" type="x:CT_HeaderFooter" minOccurs="0" maxOccurs="1"/>
<xsd:element name="rowBreaks" type="x:CT_PageBreak" minOccurs="0" maxOccurs="1"/>
<xsd:element name="colBreaks" type="x:CT_PageBreak" minOccurs="0" maxOccurs="1"/>
<xsd:element name="customProperties" type="x:CT_CustomProperties" minOccurs="0"
maxOccurs="1"/>
<xsd:element name="cellWatches" type="x:CT_CellWatches" minOccurs="0" maxOccurs="1"/>
<xsd:element name="ignoredErrors" type="x:CT_IgnoredErrors" minOccurs="0" maxOccurs="1"/>
<xsd:element name="smartTags" type="x:CT_SmartTags" minOccurs="0" maxOccurs="1"/>
<xsd:element name="drawing" type="x:CT_Drawing" minOccurs="0" maxOccurs="1"/>
<xsd:element name="picture" type="x:CT_SheetBackgroundPicture" minOccurs="0"
maxOccurs="1"/>
<xsd:element name="oleObjects" type="x:CT_OleObjects" minOccurs="0" maxOccurs="1"/>
<xsd:element name="controls" type="x:CT_Controls" minOccurs="0" maxOccurs="1"/>
<xsd:element name="webPublishItems" type="x:CT_WebPublishItems" minOccurs="0"
maxOccurs="1"/>
<xsd:element name="tableParts" type="x:CT_TableParts" minOccurs="0" maxOccurs="1"/>
<xsd:element name="extLst" type="x:CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
</xsd:sequence>
</xsd:complexType>
19 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

2.2.5.2 CT_DefinedNames
Target namespace: http://schemas.microsoft.com/office/excel/2008/2/main
Referenced by: definedNames
A complex type that specifies a group of additional properties for macros.
Child Elements:
definedName: A CT_DefinedName element (section 2.2.5.3) that specifies additional properties for
a particular macro.
The following W3C XML Schema ([XMLSCHEMA1/2] section 2.1) fragment defines the contents of this
complex type.
<xsd:complexType name="CT_DefinedNames">
<xsd:sequence>
<xsd:element name="definedName" type="CT_DefinedName" minOccurs="1"
maxOccurs="unbounded"/>
</xsd:sequence>
</xsd:complexType>
2.2.5.3 CT_DefinedName
Target namespace: http://schemas.microsoft.com/office/excel/2008/2/main
Referenced by: CT_DefinedNames
A complex type that specifies additional properties for a macro named name.
Child Elements:
argumentDescriptions: A CT_DefinedNameArgumentDescriptions element (section 2.2.5.4)
that specifies a group of descriptions for the macro arguments if the macro has arguments.
Attributes:
name: An ST_Xstring attribute ([ISO/IEC29500-1:2016] section 22.9.2.19) that specifies the name
of a macro. The name attribute value MUST be unique for each CT_DefinedName element inside the
CT_DefinedNames element (section 2.2.5.2). At least one of the following conditions MUST be true:
There exists a CT_DefinedName ([ISO/IEC29500-1:2016] section A.2) element with a name attribute
equal to this value and either the vbProcedure attribute or the xlm attribute is equal to 1, or there
exists a Visual Basic for Applications (VBA) macro having a name equal to this value.
The string MUST adhere to the name grammar as specified in [MS-XLSX] section 2.2.2.
The following W3C XML Schema ([XMLSCHEMA1/2] section 2.1) fragment defines the contents of this
complex type.
<xsd:complexType name="CT_DefinedName">
<xsd:sequence>
<xsd:element name="argumentDescriptions" type="CT_DefinedNameArgumentDescriptions"
minOccurs="0" maxOccurs="1"/>
</xsd:sequence>
<xsd:attribute name="name" type="x:ST_Xstring" use="required"/>
</xsd:complexType>
20 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

2.2.5.4 CT_DefinedNameArgumentDescriptions
Target namespace: http://schemas.microsoft.com/office/excel/2008/2/main
Referenced by: CT_DefinedName
A complex type that specifies a group of argument descriptions for a particular macro.
Child Elements:
argumentDescription: A CT_DefinedNameArgumentDescription element (section 2.2.5.5) that
specifies an argument description for a particular macro. Each index attribute of these child elements
MUST be unique for all CT_DefinedNameArgumentDescription elements (section 2.2.5.5) within
this CT_DefinedNameArgumentDescriptions element.
Attributes:
count: An unsignedInt attribute ([XMLSCHEMA2/2] section 3.3.22) that specifies the number of
CT_DefinedNameArgumentDescription child elements (section 2.2.5.5) of this element and
specifies the number of argument descriptions of the associated macro. MUST be greater than or
equal to 1 and MUST be less than or equal to 60.
The following W3C XML Schema ([XMLSCHEMA1/2] section 2.1) fragment defines the contents of this
complex type.
<xsd:complexType name="CT_DefinedNameArgumentDescriptions">
<xsd:sequence>
<xsd:element name="argumentDescription" type="CT_DefinedNameArgumentDescription"
minOccurs="1" maxOccurs="unbounded"/>
</xsd:sequence>
<xsd:attribute name="count" type="xsd:unsignedInt" use="optional"/>
</xsd:complexType>
2.2.5.5 CT_DefinedNameArgumentDescription
Target namespace: http://schemas.microsoft.com/office/excel/2008/2/main
Referenced by: CT_DefinedNameArgumentDescriptions
A complex type that specifies an argument description for a particular argument of a macro. The
length of the ST_Xstring ([ISO/IEC29500-1:2016] section 22.9.2.19) MUST be less than or equal to
255.
Attributes:
index: An unsignedInt attribute ([XMLSCHEMA2/2] section 3.3.22) that specifies the zero-based
index referring to the argument at position index of the associated macro. MUST be less than the
count attribute value in the parent CT_DefinedNameArgumentDescription element (section
2.2.5.5).
The following W3C XML Schema ([XMLSCHEMA1/2] section 2.1) fragment defines the contents of this
complex type.
<xsd:complexType name="CT_DefinedNameArgumentDescription">
<xsd:simpleContent>
<xsd:extension base="x:ST_Xstring">
<xsd:attribute name="index" type="xsd:unsignedInt" use="required"/>
</xsd:extension>
</xsd:simpleContent>
</xsd:complexType>
21 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

2.2.6 Simple Types
None.
2.2.7 Formulas
function-call =/ (macro-function-call / command-function-call)
2.3 Macro-Enabled PresentationML Structures
2.3.1 Part Enumeration
This section specifies the parts ([ISO/IEC29500-1:2016]) of the macro-enabled PresentationML file
format that are different from the parts of the non-macro-enabled PresentationML file format specified
in [ISO/IEC29500-1:2016] section 13.3. These parts and their relationships are summarized in the
following table.
Part Relationship target
Presentation (section 2.3.1.1) PresentationML Package ([ISO/IEC29500-1:2016]
section 13.2)
VBA Project (section 2.3.1.2) Presentation (section 2.3.1.1)
2.3.1.1 Presentation
Content Types:
 application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml
 application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml
 application/vnd.ms-powerpoint.template.macroEnabled.main+xml
Root Namespace: http://schemas.openxmlformats.org/presentationml/2006/main
Source Relationship:
http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument
This part is modified from the presentation part specified in [ISO/IEC29500-1:2016] Part 1, section
13.3.6.presentation part is permitted to have an implicit relationship to the VBA Project part.
2.3.1.2 VBA Project
Content Type: application/vnd.ms-office.vbaProject
Root Namespace: Not applicable
Source Relationship: http://schemas.microsoft.com/office/2006/relationships/vbaProject
This part is specified in [MS-OVBA].
22 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

A package MUST contain at most one VBA Project part, which MUST be the target of an implicit
relationship from the presentation part (section 2.3.1.1).
A VBA Project part MUST NOT have implicit or explicit relationships to any other part specified by this
document or any part specified in [ISO/IEC29500-1:2016].
2.3.2 Extensions
None.
2.3.3 Global Elements
None.
2.3.4 Global Attributes
None.
2.3.5 Complex Types
2.3.5.1 CT_SlideMasterIdListEntry
Target namespace: http://schemas.openxmlformats.org/presentationml/2006/main
This element specifies a slide master that is available within the corresponding presentation. A slide
master is a slide that is specifically designed to be a template for all related child layout slides.
This complex type is modified from the sldMasterId specified in [ISO/IEC29500-1:2016], section
19.2.1.36.
The id attribute of this part MUST be present in the macro-enabled PresentationML format.
2.3.6 Simple Types
None.
23 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

3 Structure Examples
3.1 VBA Supplemental Data Part
This example shows a usage of a Word VBA Supplemental Data part (section 2.1.1.3). Consider the
following XML, showing the complete contents of the VBA Supplemental Data part (section 2.1.1.3)
of a word-processing document.
<wne:vbaSuppData
xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
xmlns:w10="urn:schemas-microsoft-com:office:word"
xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">
<wne:docEvents>
<wne:eventDocOpen/>
<wne:eventDocXmlAfterInsert/>
</wne:docEvents>
<wne:mcds>
<wne:mcd wne:macroName="PROJECT.NEWMACROS.MACRO1" wne:name="Project.NewMacros.Macro1"
wne:bEncrypt="00" wne:cmg="56"/>
</wne:mcds>
</wne:vbaSuppData>
The docEvents child of the vbaSuppData element (section 2.1.3.1) specifies that both the Open
event and the XMLAfterInsert event are active for this document. The mcds element of the
vbaSuppData element (section 2.1.3.1) contains an mcd element that specifies macro information
for the "Project.NewMacros.Macro1" macro.
3.2 Macro Sheet Part
This example shows a usage of an Excel macro sheet part (section 2.2.1.2). Consider the following
XML, showing the complete contents of a macro sheet part (section 2.2.1.2) of a workbook.
<xm:macrosheet
xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="A1"/>
<sheetViews>
<sheetView showFormulas="1" tabSelected="1" workbookViewId="0">
<selection activeCell="A2" sqref="A2"/>
</sheetView>
</sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
<sheetData>
<row r="1" spans="1:1">
<c r="A1" t="b">
<f>ACTIVATE()</f>
<v>0</v>
</c>
</row>
</sheetData>
<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</xm:macrosheet>
The macrosheet element (section 2.2.3.1) specifies the macro sheet. A formula following the
grammar specified in section 2.2.7 is specified by the f element ([ISO/IEC29500-1:2016] section
18.17).
24 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

4 Security Considerations
4.1 Security Considerations for Implementers
None.
4.2 Index of Security Fields
None.
25 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

5 Appendix A: Full XML Schemas
For ease of implementation, this section provides the full W3C XML Schemas for the new elements,
attributes, complex types, and simple types specified in the preceding sections. Any schema
references to namespaces included in [ISO/IEC29500-1:2016] refer specifically to the transitional
schemas as specified in [ISO/IEC29500-4:2016].
5.1 http://schemas.microsoft.com/office/word/2006/wordml Schema
<xsd:schema elementFormDefault="qualified"
xmlns:xsd="http://www.w3.org/2001/XMLSchema"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
xmlns="http://schemas.microsoft.com/office/word/2006/wordml"
targetNamespace="http://schemas.microsoft.com/office/word/2006/wordml">
<xsd:import namespace="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
schemaLocation="word12.xsd" />
<xsd:import id="rel"
namespace="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
schemaLocation="orel.xsd" />
<xsd:element name="vbaSuppData" type="CT_VbaSuppData"/>
<xsd:complexType name="CT_DocEvents">
<xsd:sequence>
<xsd:element name="eventDocNew" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocOpen" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocClose" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocSync" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocXmlAfterInsert" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocXmlBeforeDelete" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocContentControlAfterInsert" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocContentControlBeforeDelete" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocContentControlOnExit" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocContentControlOnEnter" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocStoreUpdate" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocContentControlContentUpdate" minOccurs="0" maxOccurs="1"/>
<xsd:element name="eventDocBuildingBlockAfterInsert" minOccurs="0" maxOccurs="1"/>
</xsd:sequence>
</xsd:complexType>
<xsd:complexType name="CT_Mcd">
<xsd:attribute name="macroName" type="w:ST_String" use="optional"/>
<xsd:attribute name="name" type="w:ST_String" use="optional"/>
<xsd:attribute name="menuHelp" type="w:ST_String" use="optional"/>
<xsd:attribute name="bEncrypt" type="w:ST_UcharHexNumber"/>
<xsd:attribute name="cmg" type="w:ST_UcharHexNumber"/>
</xsd:complexType>
<xsd:complexType name="CT_Mcds">
<xsd:sequence>
<xsd:element name="mcd" type="CT_Mcd" minOccurs="0" maxOccurs="unbounded"/>
</xsd:sequence>
</xsd:complexType>
<xsd:complexType name="CT_VbaSuppData">
<xsd:sequence>
<xsd:element name="docEvents" type="CT_DocEvents" minOccurs="0" maxOccurs="1"/>
<xsd:element name="mcds" type="CT_Mcds" minOccurs="0" maxOccurs="1"/>
</xsd:sequence>
</xsd:complexType>
</xsd:schema>
5.2 http://schemas.microsoft.com/office/excel/2006/main Schema
<xsd:schema elementFormDefault="qualified" xmlns:xsd="http://www.w3.org/2001/XMLSchema"
xmlns="http://schemas.microsoft.com/office/excel/2006/main"
xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
26 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

targetNamespace="http://schemas.microsoft.com/office/excel/2006/main"
<xsd:import namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
schemaLocation="xlsheet.xsd" />
<xsd:import
namespace="http://schemas.microsoft.com/office/excel/2006/main"
schemaLocation="xlmacrosheet.xsd" />
<xsd:element name="macrosheet" type="CT_Worksheet"/>
<xsd:complexType name="CT_Worksheet">
<xsd:sequence>
<xsd:element name="sheetPr" type="x:CT_SheetPr" minOccurs="0" maxOccurs="1"/>
<xsd:element name="dimension" type="x:CT_SheetDimension" minOccurs="0" maxOccurs="1"/>
<xsd:element name="sheetViews" type="x:CT_SheetViews" minOccurs="0" maxOccurs="1"/>
<xsd:element name="sheetFormatPr" type="x:CT_SheetFormatPr" minOccurs="0"
maxOccurs="1"/>
<xsd:element name="cols" type="x:CT_Cols" minOccurs="0" maxOccurs="unbounded"/>
<xsd:element name="sheetData" type="x:CT_SheetData" minOccurs="1" maxOccurs="1"/>
<xsd:element name="sheetCalcPr" type="x:CT_SheetCalcPr" minOccurs="0" maxOccurs="1"/>
<xsd:element name="sheetProtection" type="x:CT_SheetProtection" minOccurs="0"
maxOccurs="1"/>
<xsd:element name="protectedRanges" type="x:CT_ProtectedRanges" minOccurs="0"
maxOccurs="1"/>
<xsd:element name="scenarios" type="x:CT_Scenarios" minOccurs="0" maxOccurs="1"/>
<xsd:element name="autoFilter" type="x:CT_AutoFilter" minOccurs="0" maxOccurs="1"/>
<xsd:element name="sortState" type="x:CT_SortState" minOccurs="0" maxOccurs="1"/>
<xsd:element name="dataConsolidate" type="x:CT_DataConsolidate" minOccurs="0"
maxOccurs="1"/>
<xsd:element name="customSheetViews" type="x:CT_CustomSheetViews" minOccurs="0"
maxOccurs="1"/>
<xsd:element name="mergeCells" type="x:CT_MergeCells" minOccurs="0" maxOccurs="1"/>
<xsd:element name="phoneticPr" type="x:CT_PhoneticPr" minOccurs="0" maxOccurs="1"/>
<xsd:element name="conditionalFormatting" type="x:CT_ConditionalFormatting"
minOccurs="0" maxOccurs="unbounded"/>
<xsd:element name="dataValidations" type="x:CT_DataValidations" minOccurs="0"
maxOccurs="1"/>
<xsd:element name="hyperlinks" type="x:CT_Hyperlinks" minOccurs="0" maxOccurs="1"/>
<xsd:element name="printOptions" type="x:CT_PrintOptions" minOccurs="0" maxOccurs="1"/>
<xsd:element name="pageMargins" type="x:CT_PageMargins" minOccurs="0" maxOccurs="1"/>
<xsd:element name="pageSetup" type="x:CT_PageSetup" minOccurs="0" maxOccurs="1"/>
<xsd:element name="headerFooter" type="x:CT_HeaderFooter" minOccurs="0" maxOccurs="1"/>
<xsd:element name="rowBreaks" type="x:CT_PageBreak" minOccurs="0" maxOccurs="1"/>
<xsd:element name="colBreaks" type="x:CT_PageBreak" minOccurs="0" maxOccurs="1"/>
<xsd:element name="customProperties" type="x:CT_CustomProperties" minOccurs="0"
maxOccurs="1"/>
<xsd:element name="cellWatches" type="x:CT_CellWatches" minOccurs="0" maxOccurs="1"/>
<xsd:element name="ignoredErrors" type="x:CT_IgnoredErrors" minOccurs="0"
maxOccurs="1"/>
<xsd:element name="smartTags" type="x:CT_SmartTags" minOccurs="0" maxOccurs="1"/>
<xsd:element name="drawing" type="x:CT_Drawing" minOccurs="0" maxOccurs="1"/>
<xsd:element name="picture" type="x:CT_SheetBackgroundPicture" minOccurs="0"
maxOccurs="1"/>
<xsd:element name="oleObjects" type="x:CT_OleObjects" minOccurs="0" maxOccurs="1"/>
<xsd:element name="controls" type="x:CT_Controls" minOccurs="0" maxOccurs="1"/>
<xsd:element name="webPublishItems" type="x:CT_WebPublishItems" minOccurs="0"
maxOccurs="1"/>
<xsd:element name="tableParts" type="x:CT_TableParts" minOccurs="0" maxOccurs="1"/>
<xsd:element name="extLst" type="x:CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
</xsd:sequence>
</xsd:complexType>
</xsd:schema>
5.3 http://schemas.microsoft.com/office/excel/2008/2/main Schema
<xsd:schema elementFormDefault="qualified"
xmlns:xsd="http://www.w3.org/2001/XMLSchema"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns="http://schemas.microsoft.com/office/excel/2008/2/main"
xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
27 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"
targetNamespace="http://schemas.microsoft.com/office/excel/2008/2/main">
<xsd:import namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
schemaLocation="xlbasictypes.xsd" />
<xsd:import namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
schemaLocation="xlsheet.xsd" />
<xsd:import namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
schemaLocation="xlpivot.xsd" />
<xsd:import namespace="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
schemaLocation="orel.xsd" />
<xsd:import namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
schemaLocation="xlworkbook.xsd" />
<xsd:import namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
schemaLocation="xlsupbook.xsd" />
<xsd:element name="macrosheet" type="CT_Worksheet"/>
<xsd:element name="definedNames" type="CT_DefinedNames"/>
<xsd:complexType name="CT_DefinedName">
<xsd:sequence>
<xsd:element name="argumentDescriptions" type="CT_DefinedNameArgumentDescriptions"
minOccurs="0" maxOccurs="1"/>
</xsd:sequence>
<xsd:attribute name="name" type="x:ST_Xstring" use="required"/>
</xsd:complexType>
<xsd:complexType name="CT_DefinedNameArgumentDescription">
<xsd:simpleContent>
<xsd:extension base="x:ST_Xstring">
<xsd:attribute name="index" type="xsd:unsignedInt" use="required"/>
</xsd:extension>
</xsd:simpleContent>
</xsd:complexType>
<xsd:complexType name="CT_DefinedNameArgumentDescriptions">
<xsd:sequence>
<xsd:element name="argumentDescription" type="CT_DefinedNameArgumentDescription"
minOccurs="1" maxOccurs="unbounded"/>
</xsd:sequence>
<xsd:attribute name="count" type="xsd:unsignedInt" use="optional"/>
</xsd:complexType>
<xsd:complexType name="CT_DefinedNames">
<xsd:sequence>
<xsd:element name="definedName" type="CT_DefinedName" minOccurs="1"
maxOccurs="unbounded"/>
</xsd:sequence>
</xsd:complexType>
</xsd:schema>
28 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

6 Appendix B: Product Behavior
The information in this specification is applicable to the following Microsoft products or supplemental
software. References to product versions include updates to those products.
 Microsoft Office 2010 suites
 Microsoft Office 2013
 Microsoft Visual Basic for Applications 6.0 (VBA 6.0)
 Microsoft Visual Basic for Applications 7.0 (VBA 7.0)
 Microsoft Office 2016
 Microsoft Office 2019
 Microsoft Office 2021
 Microsoft Office LTSC 2024
Exceptions, if any, are noted in this section. If an update version, service pack or Knowledge Base
(KB) number appears with a product name, the behavior changed in that update. The new behavior
also applies to subsequent updates unless otherwise specified. If a product edition appears with the
product version, behavior is different in that product edition.
Unless otherwise specified, any statement of optional behavior in this specification that is prescribed
using the terms "SHOULD" or "SHOULD NOT" implies product behavior in accordance with the
SHOULD or SHOULD NOT prescription. Unless otherwise specified, the term "MAY" implies that the
product does not follow the prescription.
29 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

7 Change Tracking
This section identifies changes that were made to this document since the last release. Changes are
classified as Major, Minor, or None.
The revision class Major means that the technical content in the document was significantly revised.
Major changes affect protocol interoperability or implementation. Examples of major changes are:
 A document revision that incorporates changes to interoperability requirements.
 A document revision that captures changes to protocol functionality.
The revision class Minor means that the meaning of the technical content was clarified. Minor changes
do not affect protocol interoperability or implementation. Examples of minor changes are updates to
clarify ambiguity at the sentence, paragraph, or table level.
The revision class None means that no new technical changes were introduced. Minor editorial and
formatting changes may have been made, but the relevant technical content is identical to the last
released version.
The changes made to this document are listed in the following table. For more information, please
contact dochelp@microsoft.com.
Section Description Revision class
6 Appendix B: Product Behavior Updated list of supported products. Major
30 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

8 Index
A global attributes - macro-enabled SpreadsheetML
structures 17
Applicability 8 global attributes - macro-enabled
WordprocessingML structures 11
C global elements - macro-enabled PresentationML
structures 23
Change tracking 30 International Macro Sheet - part enumeration 14
Common data types and fields 10 Macro Sheet - part enumeration 15
Complex types macrosheet – global element 16
CT_DefinedName 20 main document - part enumeration 10
CT_DefinedNameArgumentDescription 21 part enumeration - macro-enabled PresentationML
CT_DefinedNameArgumentDescriptions 21 structures 22
CT_DefinedNames 20 part enumeration - macro-enabled SpreadsheetML
CT_DocEvents 11 structures 14
CT_Mcd 12 part enumeration - macro-enabled
CT_Mcds 13 WordprocessingML structures 10
CT_SlideMasterIdListEntry 23 Presentation - part enumeration 22
CT_VbaSuppData 13 simple types - macro-enabled PresentationML
CT_Worksheet 17 structures 23
CT_DefinedName – complex type 20 simple types - macro-enabled SpreadsheetML
CT_DefinedNameArgumentDescription – complex structures 22
type 21 simple types - macro-enabled WordprocessingML
CT_DefinedNameArgumentDescriptions – complex structures 14
type 21 VBA Project - part enumeration (section 2.1.1.2
CT_DefinedNames – complex type 20 10, section 2.2.1.3 15, section 2.3.1.2 22)
CT_DocEvents – complex type 11 VBA Supplemental Data - part enumeration 11
CT_Mcd – complex type 12 vbaSuppData – global element 11
CT_Mcds – complex type 13 Workbook - part enumeration 16
CT_SlideMasterIdListEntry – complex type 23
CT_VbaSuppData – complex type 13 E
CT_Worksheet – complex type 17
Examples
D Macro Sheet Part 24
VBA Supplemental Data Part 24
Data types and fields - common 10 Excel extensions
definedNames – global element 17 extensions by part 16
Details Excel extensions – macro-enabled SpreadsheetML
common data types and fields 10 structures 16
CT_DefinedName – complex type 20 Extensions – macro-enabled PresentationML
CT_DefinedNameArgumentDescription – complex structures 23
type 21 Extensions – macro-enabled WordprocessingML
CT_DefinedNameArgumentDescriptions – complex structures 11
type 21 Extensions by part - Excel extensions 16
CT_DefinedNames – complex type 20
CT_DocEvents – complex type 11 F
CT_Mcd – complex type 12
CT_Mcds – complex type 13 Fields - vendor-extensible 9
CT_SlideMasterIdListEntry – complex type 23 Full XML schema 26
CT_VbaSuppData – complex type 13
CT_Worksheet – complex type 17 G
definedNames – global element 17
Excel extensions - macro-enabled SpreadsheetML Global attributes – macro-enabled PresentationML
structures 16 structures 23
extensions - macro-enabled PresentationML Global attributes – macro-enabled SpreadsheetML
structures 23 structures 17
extensions - macro-enabled WordprocessingML Global attributes – macro-enabled WordprocessingML
structures 11 structures 11
extensions by part – Excel extensions 16 Global elements
global attributes - macro-enabled PresentationML definedNames 17
structures 23 macrosheet 16
vbaSuppData 11
31 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

Global elements – macro-enabled PresentationML Simple types – macro-enabled SpreadsheetML
structures 23 structures 22
Glossary 6 Simple types – macro-enabled WordprocessingML
structures 14
I Structures
overview 10
Implementer - security considerations 25 Structures – macro-enabled PresentationML
Informative references 8 CT_SlideMasterIdListEntry – complex type 23
International Macro Sheet - part enumeration 14 extensions 23
Introduction 6 global attributes 23
global elements 23
L part enumeration 22
Presentation - part enumeration 22
Localization 9 simple types 23
VBA Project - part enumeration 22
M Structures – macro-enabled SpreadsheetML
CT_DefinedName – complex type 20
CT_DefinedNameArgumentDescription – complex
Macro Sheet - part enumeration 15
type 21
Macro Sheet Part example 24
CT_DefinedNameArgumentDescriptions – complex
macrosheet – global element 16
type 21
Main document - part enumeration 10
CT_DefinedNames – complex type 20
CT_Worksheet – complex type 17
N
definedNames – global element 17
Excel extensions 16
Normative references 7
extensions by part - Excel extensions 16
global attributes 17
O
International Macro Sheet - part enumeration 14
Macro Sheet - part enumeration 15
Overview (synopsis) 8
macrosheet – global element 16
part enumeration 14
P
simple types 22
VBA Project - part enumeration 15
Part enumeration Workbook - part enumeration 16
International Macro Sheet 14 Structures – macro-enabled WordprocessingML
Macro Sheet 15 CT_DocEvents – complex type 11
main document 10 CT_Mcd – complex type 12
Presentation 22 CT_Mcds – complex type 13
VBA Project (section 2.1.1.2 10, section 2.2.1.3 CT_VbaSuppData – complex type 13
15, section 2.3.1.2 22) extensions 11
VBA Supplemental Data 11 global attributes 11
Workbook 16 main document - part enumeration 10
Part enumeration – macro-enabled PresentationML part enumeration 10
structures 22 simple types 14
Part enumeration – macro-enabled SpreadsheetML VBA Project - part enumeration 10
structures 14 VBA Supplemental Data - part enumeration 11
Part enumeration – macro-enabled vbaSuppData – global element 11
WordprocessingML structures 10
Presentation - part enumeration 22 T
Product behavior 29
Tracking changes 30
R
V
References 7
informative 8
VBA Project - part enumeration (section 2.1.1.2 10,
normative 7
section 2.2.1.3 15, section 2.3.1.2 22)
Relationship to protocols and other structures 8
VBA Supplemental Data - part enumeration 11
VBA Supplemental Data Part example 24
S
vbaSuppData – global element 11
Vendor-extensible fields 9
Security Versioning 9
implementer considerations 25
index of security fields 25 W
Simple types – macro-enabled PresentationML
structures 23
Workbook - part enumeration 16
32 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024

X
XML schema 26
33 / 33
[MS-OFFMACRO2] - v20240820
Office Macro-Enabled File Format Version 2
Copyright © 2024 Microsoft Corporation
Release: August 20, 2024