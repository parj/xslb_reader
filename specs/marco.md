[MS-OFFMACRO]:
Office Macro-Enabled File Format
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
1 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

Revision Summary
|            | Revision  | Revision  |                       |
| ---------- | --------- | --------- | --------------------- |
| Date       | History   | Class     | Comments              |
| 7/13/2009  | 0.1       | Major     | Initial Availability  |
8/28/2009  0.2  Editorial  Revised and edited the technical content
11/6/2009  0.3  Editorial  Revised and edited the technical content
2/19/2010  1.0  Editorial  Revised and edited the technical content
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
No changes to the meaning, language, or formatting of the
| 7/16/2012  | 1.5  | None  |     |
| ---------- | ---- | ----- | --- |
technical content.
No changes to the meaning, language, or formatting of the
| 10/8/2012  | 1.5  | None  |     |
| ---------- | ---- | ----- | --- |
technical content.
No changes to the meaning, language, or formatting of the
| 2/11/2013  | 1.5  | None  |     |
| ---------- | ---- | ----- | --- |
technical content.
7/30/2013  1.5  None  No changes to the meaning, language, or formatting of the
technical content.
No changes to the meaning, language, or formatting of the
| 11/18/2013  | 1.5  | None  |     |
| ----------- | ---- | ----- | --- |
technical content.
No changes to the meaning, language, or formatting of the
| 2/10/2014  | 1.5  | None  |     |
| ---------- | ---- | ----- | --- |
technical content.
No changes to the meaning, language, or formatting of the
| 4/30/2014  | 1.5  | None  |     |
| ---------- | ---- | ----- | --- |
technical content.
2 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

|       | Revision  | Revision  |           |
| ----- | --------- | --------- | --------- |
| Date  | History   | Class     | Comments  |
No changes to the meaning, language, or formatting of the
| 7/31/2014  | 1.5  | None  |     |
| ---------- | ---- | ----- | --- |
technical content.
No changes to the meaning, language, or formatting of the
| 10/30/2014  | 1.5  | None  |     |
| ----------- | ---- | ----- | --- |
technical content.
No changes to the meaning, language, or formatting of the
| 6/23/2016  | 1.5  | None  |     |
| ---------- | ---- | ----- | --- |
technical content.
No changes to the meaning, language, or formatting of the
| 9/14/2016  | 1.5  | None  |     |
| ---------- | ---- | ----- | --- |
technical content.
3 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

Table of Contents
1 Introduction ............................................................................................................ 6
1.1 Glossary ........................................................................................................... 6
1.2 References ........................................................................................................ 7
1.2.1 Normative References ................................................................................... 7
1.2.2 Informative References ................................................................................. 8
1.3 Structure Overview (Synopsis) ............................................................................ 8
1.3.1 Organization of This Document ....................................................................... 8
1.4 Relationship to Protocols and Other Structures ...................................................... 8
1.5 Applicability Statement ....................................................................................... 9
1.6 Versioning and Localization ................................................................................. 9
1.7 Vendor-Extensible Fields ..................................................................................... 9
2 Structures ............................................................................................................. 10
2.1 Macro-Enabled WordprocessingML Structures ...................................................... 10
2.1.1 Part Enumeration ........................................................................................ 10
2.1.1.1 Main Document ..................................................................................... 10
2.1.1.2 VBA Project .......................................................................................... 10
2.1.1.3 VBA Supplemental Data ......................................................................... 11
2.1.2 Global Elements ......................................................................................... 11
2.1.2.1 vbaSuppData ........................................................................................ 11
2.1.3 Global Attributes......................................................................................... 11
2.1.4 Complex Types ........................................................................................... 11
2.1.4.1 CT_Mcd ............................................................................................... 11
2.1.4.2 CT_Mcds .............................................................................................. 12
2.1.4.3 CT_DocEvents ...................................................................................... 12
2.1.4.4 CT_VbaSuppData .................................................................................. 13
2.1.5 Simple Types ............................................................................................. 14
2.2 Macro-Enabled SpreadsheetML Structures ........................................................... 14
2.2.1 Part Enumeration ........................................................................................ 14
2.2.1.1 Workbook ............................................................................................ 14
2.2.1.2 VBA Project .......................................................................................... 14
2.2.1.3 Macro Sheet ......................................................................................... 15
2.2.1.4 International Macro Sheet ...................................................................... 15
2.2.2 Global Elements ......................................................................................... 16
2.2.2.1 macrosheet .......................................................................................... 16
2.2.3 Global Attributes......................................................................................... 16
2.2.4 Complex Types ........................................................................................... 16
2.2.4.1 CT_Worksheet ...................................................................................... 16
2.2.5 Simple Types ............................................................................................. 19
2.2.6 Formulas ................................................................................................... 19
2.3 Macro-Enabled PresentationML Structures ........................................................... 34
2.3.1 Part Enumeration ........................................................................................ 34
2.3.1.1 Presentation ......................................................................................... 34
2.3.1.2 VBA Project .......................................................................................... 34
2.3.2 Global Elements ......................................................................................... 34
2.3.3 Global Attributes......................................................................................... 34
2.3.4 Complex Types ........................................................................................... 35
2.3.4.1 CT_SlideMasterIdListEntry ..................................................................... 35
2.3.5 Simple Types ............................................................................................. 35
3 Structure Examples ............................................................................................... 36
3.1 VBA Supplemental Data Part of a Word-Processing Document................................ 36
3.2 Macro Sheet Part of a Workbook ........................................................................ 36
4 Security Considerations ......................................................................................... 37
4 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

5 Appendix A: Full XML Schemas .............................................................................. 38
5.1 http://schemas.microsoft.com/office/word/2006/wordml ...................................... 38
5.2 http://schemas.microsoft.com/office/excel/2006/main ......................................... 38
6 Appendix B: Product Behavior ............................................................................... 40
7 Change Tracking .................................................................................................... 41
8 Index ..................................................................................................................... 42
5 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

1 Introduction
This document specifies the Office Macro-Enabled File Format for WordprocessingML, SpreadsheetML,
and PresentationML file formats that are either not present in or modified from non-macro-enabled
WordprocessingML, SpreadsheetML, and PresentationML file formats described in [ECMA-376]. The
macro-enabled file formats consist of the non-macro-enabled file formats described in [ECMA-376]
and the XML structures and the XML and non-XML parts specified in this document.
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
document: An object in a content database such as a file, folder, list (1), or site (2). Each object is
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
6 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

protection: A mechanism that helps restrict users from making unwanted changes to the data or
structure of a workbook.
range: An addressable region that is in a workbook. A range typically consists of zero or more cells
and represents a single, contiguous rectangle of cells on a single sheet.
sheet: A part of an Excel workbook. There are four types of sheets: worksheet, macro sheet,
dialog sheet, and chart sheet. Multiple sheets are stored together within a workbook.
sheet view: A collection of display settings, such as which cells are shown, and the zoom level for
a sheet window.
sort condition: A condition that determines how to sort cells in a range.
table: A list (2) that is defined in a workbook.
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
[ECMA-376] ECMA International, "Office Open XML File Formats", 1st Edition, ECMA-376, December
2006, http://www.ecma-international.org/publications/standards/Ecma-376.htm
[MS-OVBA] Microsoft Corporation, "Office VBA File Format Structure".
[MS-XLSX] Microsoft Corporation, "Excel (.xlsx) Extensions to the Office Open XML SpreadsheetML File
Format".
[RFC2119] Bradner, S., "Key words for use in RFCs to Indicate Requirement Levels", BCP 14, RFC
2119, March 1997, http://www.rfc-editor.org/rfc/rfc2119.txt
7 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

[XMLSCHEMA1] Thompson, H., Beech, D., Maloney, M., and Mendelsohn, N., Eds., "XML Schema Part
1: Structures", W3C Recommendation, May 2001, http://www.w3.org/TR/2001/REC-xmlschema-1-
20010502/
1.2.2 Informative References
[ISO/IEC-29500-4] International Organization for Standardization, "Information technology --
Document description and processing languages -- Office Open XML File Formats -- Part 4:
Transitional Migration Features", ISO/IEC 29500-4:2008,
http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51462
1.3 Structure Overview (Synopsis)
This document specifies structures of macro-enabled WordprocessingML, SpreadsheetML, and
PresentationML file formats that are either not present in or modified from non-macro-enabled
WordprocessingML, SpreadsheetML, and PresentationML file formats described in [ECMA-376].
Document, Workbook, and Presentation data are each contained in a in a compressed file (.zip)
package conforming to the Open Packaging Conventions as described in [ECMA-376] section 2.
Individual files stored in a ZIP package, called parts ([ECMA-376]), contain information about the
content and structure of a document, workbook, or presentation. Some parts ([ECMA-376]) store
information using XML and other parts ([ECMA-376]) store information by using binary data.
1.3.1 Organization of This Document
Section 2 of this documentation is organized by application.
Section 2.1 covers the structures in the macro-enabled WordprocessingML file format that are either
not present in or modified from the structures in the non-macro-enabled WordprocessingML file format
described in [ECMA-376].
Section 2.2 covers the structures in the macro-enabled SpreadsheetML file format that are either not
present in or modified from the structures in the non-macro-enabled SpreadsheetML file format
described in [ECMA-376].
Section 2.3 covers the structures in the macro-enabled PresentationML file format that are either not
present in or modified from the structures in the non-macro-enabled PresentationML file format
described in [ECMA-376].
Within sections 2.1, 2.2, and 2.3, subsection 1 specifies the parts within the package of a macro-
enabled file that are either not present in or modified from a non-macro-enabled file described in
[ECMA-376]. Sections 2, 3, 4, and 5 described the details of the structures that are either not present
in or are modified from the structures in the non-macro-enabled file formats specified by [ECMA-376].
Section 2 describes global elements, section 3 describes global attributes, section 4 describes complex
types, and section 5 describes simple types.
Section 3 provides specific examples intended to illustrate the structures specified in section 2.
Section 4 discusses security issues relating to the structures specified by this document.
Section 5 is a list of version-specific behaviors. It is not intended to be read alone, but rather to be
understood in the context of specifications in section 2. Specifications in section 2 provide links to the
relevant items in section 5.
1.4 Relationship to Protocols and Other Structures
The structures described in section 2 are dependent on the structures and concepts described in the
following references:
8 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

 [ECMA-376] sections 1, 3, and 4 for baseline WordprocessingML, SpreadsheetML, and
PresentationML persistence formats.
 [ECMA-376] section 2 for open packaging conventions.
 [ECMA-376] section 5 for markup compatibility and extensibility.
1.5 Applicability Statement
This document specifies a persistence format for structures not present in or modified from the
structures described in [ECMA-376] for WordprocessingML, SpreadsheetML, and PresentationML
documents. The structures specified in this document are not applicable as a stand-alone file format.
Each structure described in this document is integrated with [ECMA-376] WordprocessingML,
SpreadsheetML, or PresentationML documents in a particular way, as described in the section for that
structure.
1.6 Versioning and Localization
None.
1.7 Vendor-Extensible Fields
None.
9 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

2 Structures
2.1 Macro-Enabled WordprocessingML Structures
2.1.1 Part Enumeration
This section specifies the parts ([ECMA-376]) of the macro-enabled WordprocessingML file format that
are different from the parts of the non-macro-enabled WordprocessingML file format specified in
[ECMA-376] Part 1, Section 11.3. These parts and their relationships are summarized in the following
table.
Part Relationship target
Main Document (section 2.1.1.1) WordprocessingML Package ([ECMA-376] Part 1, Section 11.2)
VBA Project (section 2.1.1.2) Main Document
VBA Supplemental Data (section
VBA Project
2.1.1.3)
2.1.1.1 Main Document
Content Type:
 application/vnd.ms-word.document.macroEnabled.main+xml
 application/vnd.ms-word.template.macroEnabledTemplate.main+xml
Root Namespace: http://schemas.openxmlformats.org/wordprocessingml/2006/main
Source
Relationship:http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocumen
t
The part is specified in [ECMA-376] Part 1, Section 11.3.10.
2.1.1.2 VBA Project
Content Type: application/vnd.ms-office.vbaProject
Root Namespace: Not applicable
Source Relationship: http://schemas.microsoft.com/office/2006/relationships/vbaProject
This part is specified in [MS-OVBA].
A package MUST contain at most one VBA Project part, which MUST be the target of an implicit
relationship from the main document part (section 2.1.1.1).
A VBA Project part MUST have an implicit relationship to the VBA Supplemental Data part (section
2.1.1.3).
10 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

A VBA Project part MUST NOT have implicit or explicit relationships to any other part specified by this
document or any part specified in [ECMA-376].
2.1.1.3 VBA Supplemental Data
Content Type: application/vnd.ms-word.vbaData+xml
Root Namespace: http://schemas.microsoft.com/office/word/2006/wordml
Source Relationship: http://schemas.microsoft.com/office/2006/relationships/wordVbaData
An instance of this part type specifies Microsoft Visual Basic for Applications (VBA) supplemental
Data.
A VBA Supplemental Data part MUST be the target of an implicit relationship from the VBA Project
part (section 2.1.1.2).
A VBA Supplemental Data part MUST NOT have implicit or explicit relationships to any other part
specified by this document or any part specified in [ECMA-376].
The root element for a part of this content type MUST be vbaSuppData (section 2.1.2.1).
2.1.2 Global Elements
2.1.2.1 vbaSuppData
Target namespace: http://schemas.microsoft.com/office/word/2006/wordml
A CT_VbaSuppData element (section 2.1.4.4) that specifies supplemental Microsoft Visual Basic for
Applications (VBA) data.
The following W3C XML Schema ([XMLSCHEMA1] section 2.1) fragment defines the contents of this
element.
<xsd:element name="vbaSuppData" type="CT_VbaSuppData"/>
2.1.3 Global Attributes
None.
2.1.4 Complex Types
2.1.4.1 CT_Mcd
Target namespace: http://schemas.microsoft.com/office/word/2006/wordml
Referenced by: CT_Mcds
A complex type that specifies a macro.
Attributes:
macroName : An ST_String attribute ([ECMA-376] Part 4, Section 2.18.89) that MUST equal the
name attribute with every character set to uppercase.
11 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

name : An ST_String attribute ([ECMA-376] Part 4, Section 2.18.89) that specifies the name of the
macro. name MUST NOT exceed 255 characters.
menuHelp : An ST_String attribute ([ECMA-376] Part 4, Section 2.18.89) that MUST be ignored.
bEncrypt : An ST_UcharHexNumber attribute ([ECMA-376] Part 4, Section 2.18.106) that MUST be
0.
cmg : An ST_UcharHexNumber attribute ([ECMA-376] Part 4, Section 2.18.106) that MUST be 56.
The following W3C XML Schema ([XMLSCHEMA1] section 2.1) fragment defines the contents of this
complex type.
<xsd:complexType name="CT_Mcd">
<xsd:attribute name="macroName" type="w:ST_String" use="optional"/>
<xsd:attribute name="name" type="w:ST_String" use="optional"/>
<xsd:attribute name="menuHelp" type="w:ST_String" use="optional"/>
<xsd:attribute name="bEncrypt" type="w:ST_UcharHexNumber"/>
<xsd:attribute name="cmg" type="w:ST_UcharHexNumber"/>
</xsd:complexType>
2.1.4.2 CT_Mcds
Target namespace: http://schemas.microsoft.com/office/word/2006/wordml
Referenced by: CT_VbaSuppData
A complex type that specifies a collection of macros.
Child Elements:
mcd : A CT_Mcd element (section 2.1.4.1) that specifies a macro.
The following W3C XML Schema ([XMLSCHEMA1] section 2.1) fragment defines the contents of this
complex type.
<xsd:complexType name="CT_Mcds">
<xsd:sequence>
<xsd:element name="mcd" type="CT_Mcd" minOccurs="0" maxOccurs="unbounded"/>
</xsd:sequence>
</xsd:complexType>
2.1.4.3 CT_DocEvents
Target namespace: http://schemas.microsoft.com/office/word/2006/wordml
Referenced by: CT_VbaSuppData
A complex type that specifies a container of document events that are active for this document.
Child Elements:
eventDocNew : An element that specifies that the New event is active.
eventDocOpen : An element that specifies that the Open event is active.
eventDocClose : An element that specifies that the Close event is active.
eventDocSync : An element that specifies that the Sync event is active.
12 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

eventDocXmlAfterInsert : An element that specifies that the XmlAfterInsert event is active.
eventDocXmlBeforeDelete : An element that specifies that the XmlBeforeDelete event is active.
eventDocContentControlAfterInsert : An element that specifies that the
ContentControlAfterInsert event is active.
eventDocContentControlBeforeDelete : An element that specifies that the
ContentControlBeforeDelete event is active.
eventDocContentControlOnExit : An element that specifies that the ContentControlOnExit event
is active.
eventDocContentControlOnEnter : An element that specifies that the ContentControlOnEnter
event is active.
eventDocStoreUpdate : An element that specifies that the StoreUpdate event is active.
eventDocContentControlContentUpdate : An element that specifies that the
ContentControlContentUpdate event is active.
eventDocBuildingBlockAfterInsert : An element that specifies that the BuildingBlockAfterInsert
event is active.
The following W3C XML Schema ([XMLSCHEMA1] section 2.1) fragment defines the contents of this
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
2.1.4.4 CT_VbaSuppData
Target namespace: http://schemas.microsoft.com/office/word/2006/wordml
Referenced by: vbaSuppData
A complex type that specifies document events and macro information for the VBA project.
Child Elements:
docEvents : A CT_DocEvents element (section 2.1.4.3) that specifies a container of document
events that are active for this document.
mcds : A CT_Mcds element (section 2.1.4.2) that specifies a collection of macros stored in the VBA
project.
13 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

The following W3C XML Schema ([XMLSCHEMA1] section 2.1) fragment defines the contents of this
complex type.
<xsd:complexType name="CT_VbaSuppData">
<xsd:sequence>
<xsd:element name="docEvents" type="CT_DocEvents" minOccurs="0" maxOccurs="1"/>
<xsd:element name="mcds" type="CT_Mcds" minOccurs="0" maxOccurs="1"/>
</xsd:sequence>
</xsd:complexType>
2.1.5 Simple Types
None.
2.2 Macro-Enabled SpreadsheetML Structures
2.2.1 Part Enumeration
This section specifies the parts ([ECMA-376]) of the macro-enabled SpreadsheetML file format that are
different from the parts of the non-macro-enabled SpreadsheetML file format specified in [ECMA-376]
Part 1, Section 12.3. These parts and their relationships are summarized in the following table.
Part Relationship target
Workbook (section 2.2.1.1) SpreadsheetML Package ([ECMA-376] Part 1, Section 12.2)
VBA Project (section 2.2.1.2) Workbook
Macro Sheet (section 2.2.1.3) Workbook
International Macro Sheet (section
Workbook
2.2.1.4)
2.2.1.1 Workbook
Content Type(s):
 application/vnd.ms-excel.sheet.macroEnabled.main+xml
 application/vnd.ms-excel.template.macroEnabled.main+xml
Root Namespace: http://schemas.openxmlformats.org/spreadsheetml/2006/main
Source Relationship:
http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument
This part is specified in [ECMA-376] Part 1, Section 12.3.23.
2.2.1.2 VBA Project
Content Type: application/vnd.ms-office.vbaProject
Root Namespace: Not applicable.
14 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

Source Relationship: http://schemas.microsoft.com/office/2006/relationships/vbaProject
This part is specified in [MS-OVBA].
A package MUST contain at most one VBA Project part, which MUST be the target of an implicit
relationship from the workbook part (section 2.2.1.1).
A VBA Project part MUST NOT have implicit or explicit relationships to any other part specified by this
document or any other part specified in [ECMA-376].
2.2.1.3 Macro Sheet
Content Type: application/vnd.ms-excel.macrosheet+xml
Root Namespace: http://schemas.microsoft.com/office/excel/2006/main
Source Relationship: http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet
An instance of this part type specifies a macro sheet. References to the Macro Sheet part in this
document include international macro sheets (section 2.2.1.4) as well.
A Macro Sheet part MUST be the target of an explicit relationship from the workbook part (section
2.2.1.1).
The root element for a part of this content type MUST be a Macro Sheet (section 2.2.1.3).
A Macro Sheet part is permitted to have implicit relationships to the following parts specified in
[ECMA-376]:
 Comments
 Macro Sheet Binary Index
 Sort Map
A Macro Sheet part is permitted to have explicit relationships to the following parts specified in
[ECMA-376]:
 Custom Property
 Drawings
 OLE Object
 OLE Package
 Images
 Printer Settings
 VML Drawing
A Macro Sheet part MUST NOT have implicit or explicit relationships to any other part specified in
[ECMA-376].
A formula ([ECMA-376] Part 4, Section 3.17) in a macro sheet is specified in section 2.2.6, which is an
extension of the definition in [MS-XLSX] section 2.2.2.
2.2.1.4 International Macro Sheet
Content Type: application/vnd.ms-excel.intlmacrosheet+xml
15 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

Root Namespace: http://schemas.microsoft.com/office/excel/2006/main
Source Relationship: http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet
An instance of this part type specifies the same information as a macrosheet part type (section
2.2.2.1). When this part type is present, the associated macro sheet will display and operate with the
US English locale, regardless of what the product locale, user locale, or system locale is. This will
maximize compatibility when running under various locales.
2.2.2 Global Elements
2.2.2.1 macrosheet
Target namespace: http://schemas.microsoft.com/office/excel/2006/main
A CT_Worksheet element (section 2.2.4.1) that specifies a macro sheet.
The following W3C XML Schema ([XMLSCHEMA1] section 2.1) fragment defines the contents of this
element.
<xsd:element name="macrosheet" type="CT_Worksheet"/>
2.2.3 Global Attributes
None.
2.2.4 Complex Types
2.2.4.1 CT_Worksheet
Target namespace: http://schemas.microsoft.com/office/excel/2006/main
Referenced by: macrosheet
A complex type that specifies properties for a macro sheet.
Child Elements:
sheetPr : A CT_SheetPr element ([ECMA-376] Part 4, Annex A) that specifies sheet-level
properties.
dimension : A CT_SheetDimension element ([ECMA-376] Part 4, Annex A) that specifies the used
range of the worksheet.
sheetViews : A CT_SheetViews element ([ECMA-376] Part 4, Annex A) that specifies sheet view
information.
sheetFormatPr : A CT_SheetFormatPr element ([ECMA-376] Part 4, Annex A) that specifies sheet
formatting properties.
cols : A CT_Cols element ([ECMA-376] Part 4, Annex A) that specifies information about whole
columns of the worksheet.
sheetData : A CT_SheetData element ([ECMA-376] Part 4, Annex A) that specifies the Worksheet
([ECMA-376] Part 4, Section 3.3) cell table.
sheetCalcPr : This element MUST be ignored in a macro sheet.
16 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

sheetProtection : A CT_SheetProtection element ([ECMA-376] Part 4, Annex A) that specifies
sheet protection options.
protectedRanges : This element MUST be ignored in a macro sheet.
scenarios : This element MUST be ignored in a macro sheet.
autoFilter : A CT_AutoFilter element ([ECMA-376] Part 4, Annex A) that specifies AutoFilter
settings.
sortState : A CT_SortState element ([ECMA-376] Part 4, Annex A) that specifies the different sort
conditions that apply to a range.
dataConsolidate : A CT_DataConsolidate element ([ECMA-376] Part 4, Annex A) that specifies
data consolidation settings.
customSheetViews : A CT_CustomSheetViews element ([ECMA-376] Part 4, Annex A) that
specifies custom sheet views.
mergeCells : This element MUST be ignored in a macro sheet.
phoneticPr : A CT_PhoneticPr element ([ECMA-376] Part 4, Annex A) that specifies default
phonetic information.
conditionalFormatting : A CT_ConditionalFormatting element ([ECMA-376] Part 4, Annex A) that
specifies conditional formatting rules.
dataValidations : This element MUST be ignored in a macro sheet.
hyperlinks : This element MUST be ignored in a macro sheet.
printOptions : A CT_PrintOptions element ([ECMA-376] Part 4, Annex A) that specifies print
options for the sheet.
pageMargins : A CT_PageMargins element ([ECMA-376] Part 4, Annex A) that specifies page
margins for the sheet.
pageSetup : A CT_PageSetup element ([ECMA-376] Part 4, Annex A) that specifies page setup
settings for the sheet.
headerFooter : A CT_HeaderFooter element ([ECMA-376] Part 4, Annex A) that specifies header
and footer settings.
rowBreaks : A CT_PageBreak element ([ECMA-376] Part 4, Annex A) that specifies horizontal page
break information.
colBreaks : A CT_PageBreak element ([ECMA-376] Part 4, Annex A) that specifies vertical page
break information.
customProperties : A CT_CustomProperties element ([ECMA-376] Part 4, Annex A) that specifies
references to binary parts containing arbitrary user-defined data.
cellWatches : This element MUST be ignored in a macro sheet.
ignoredErrors : This element MUST be ignored in a macro sheet.
smartTags : This element MUST be ignored in a macro sheet.
drawing : A CT_Drawing element ([ECMA-376] Part 4, Annex A) that specifies a reference to a
drawing part ([ECMA-376] Part 1, Section 12.3.8) containing one or more drawing ([ECMA-376] Part
4, Section 5.6) elements.
17 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

picture : A CT_SheetBackgroundPicture element ([ECMA-376] Part 4, Annex A) that specifies a
background image of a sheet.
oleObjects : A CT_OleObjects element ([ECMA-376] Part 4, Annex A) that specifies embedded
objects in the sheet.
controls : This element MUST be ignored in a macro sheet.
webPublishItems : A CT_WebPublishItems element ([ECMA-376] Part 4, Annex A) that specifies
individual objects that have been published to HTML.
tableParts : A CT_TableParts element ([ECMA-376] Part 4, Annex A) that specifies references to
parts specifying every table on the sheet.
extLst : A CT_ExtensionList element ([ECMA-376] Part 4, Annex A) that specifies future extensions
to this format.
The following W3C XML Schema ([XMLSCHEMA1] section 2.1) fragment defines the contents of this
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
18 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

<xsd:element name="extLst" type="x:CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
</xsd:sequence>
</xsd:complexType>
2.2.5 Simple Types
None.
2.2.6 Formulas
The syntax of a formula ([ECMA-376] Part 4, Section 3.17) in a macro sheet is specified by the
grammar in [MS-XLSX] section 2.2.2, along with the following extension to that grammar.
function-call =/ (macro-function-call / command-function-call)
ref-macro-function-call = "ABSREF" absref-params
ref-macro-function-call =/ "ACTIVE.CELL" active-cell-params
ref-macro-function-call =/ "CALL" call-params
ref-macro-function-call =/ "CALLER" caller-params
ref-macro-function-call =/ "EVALUATE" evaluate-params
ref-macro-function-call =/ "GET.DOCUMENT" get-document-params
ref-macro-function-call =/ "INPUT" input-params
ref-macro-function-call =/ "LAST.ERROR" last-error-params
ref-macro-function-call =/ "SCENARIO.GET" scenario-get-params
ref-macro-function-call =/ "SELECTION" selection-params
ref-macro-function-call =/ "TEXTREF" textref-params
ref-macro-function-call =/ "VIEW.GET" view-get-params
macro-function-call = ref-macro-function-call
macro-function-call =/ "ADD.BAR" add-bar-params
macro-function-call =/ "ADD.COMMAND" add-command-params
macro-function-call =/ "ADD.MENU" add-menu-params
macro-function-call =/ "ADD.TOOLBAR" add-toolbar-params
macro-function-call =/ "APP.TITLE" app-title-params
macro-function-call =/ "ARGUMENT" argument-params
macro-function-call =/ "BREAK" break-params
macro-function-call =/ "CANCEL.KEY" cancel-key-params
macro-function-call =/ "CHECK.COMMAND" check-command-params
macro-function-call =/ "CREATE.OBJECT" create-object-params
macro-function-call =/ "CUSTOM.REPEAT" custom-repeat-params
macro-function-call =/ "CUSTOM.UNDO" custom-undo-params
macro-function-call =/ "DELETE.BAR" delete-bar-params
macro-function-call =/ "DELETE.COMMAND" delete-command-params
macro-function-call =/ "DELETE.MENU" delete-menu-params
macro-function-call =/ "DELETE.TOOLBAR" delete-toolbar-params
macro-function-call =/ "DEREF" deref-params
macro-function-call =/ "DIALOG.BOX" dialog-box-params
macro-function-call =/ "DIRECTORY" directory-params
macro-function-call =/ "DOCUMENTS" documents-params
macro-function-call =/ "ECHO" echo-params
macro-function-call =/ "ELSE" else-params
macro-function-call =/ "ELSE.IF" else-if-params
macro-function-call =/ "ENABLE.COMMAND" enable-command-params
macro-function-call =/ "ENABLE.TOOL" enable-tool-params
macro-function-call =/ "END.IF" end-if-params
macro-function-call =/ "ERROR" error-params
macro-function-call =/ "EXEC" exec-params
macro-function-call =/ "EXECUTE" execute-params
macro-function-call =/ "FCLOSE" fclose-params
macro-function-call =/ "FILES" files-params
macro-function-call =/ "FOPEN" fopen-params
macro-function-call =/ "FOR" for-params
macro-function-call =/ "FOR.CELL" for-cell-params
macro-function-call =/ "FORMULA.CONVERT" formula-convert-params
macro-function-call =/ "FPOS" fpos-params
macro-function-call =/ "FREAD" fread-params
19 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

macro-function-call =/ "FREADLN" freadln-params
macro-function-call =/ "FSIZE" fsize-params
macro-function-call =/ "FWRITE" fwrite-params
macro-function-call =/ "FWRITELN" fwriteln-params
macro-function-call =/ "GET.BAR" get-bar-params
macro-function-call =/ "GET.CELL" get-cell-params
macro-function-call =/ "GET.CHART.ITEM" get-chart-item-params
macro-function-call =/ "GET.DEF" get-def-params
macro-function-call =/ "GET.FORMULA" get-formula-params
macro-function-call =/ "GET.LINK.INFO" get-link-info-params
macro-function-call =/ "GET.MOVIE" get-movie-params
macro-function-call =/ "GET.NAME" get-name-params
macro-function-call =/ "GET.NOTE" get-note-params
macro-function-call =/ "GET.OBJECT" get-object-params
macro-function-call =/ "GET.TOOL" get-tool-params
macro-function-call =/ "GET.TOOLBAR" get-toolbar-params
macro-function-call =/ "GET.WINDOW" get-window-params
macro-function-call =/ "GET.WORKBOOK" get-workbook-params
macro-function-call =/ "GET.WORKSPACE" get-workspace-params
macro-function-call =/ "GOTO" goto-params
macro-function-call =/ "GROUP" group-params
macro-function-call =/ "HALT" halt-params
macro-function-call =/ "HELP" help-params
macro-function-call =/ "INITIATE" initiate-params
macro-function-call =/ "LINKS" links-params
macro-function-call =/ "MOVIE.COMMAND" movie-command-params
macro-function-call =/ "NAMES" names-params
macro-function-call =/ "NEXT" next-params
macro-function-call =/ "NOTE" note-params
macro-function-call =/ "OPEN.DIALOG" open-dialog-params
macro-function-call =/ "OPTIONS.LISTS.GET" options-lists-get-params
macro-function-call =/ "PAUSE" pause-params
macro-function-call =/ "POKE" poke-params
macro-function-call =/ "PRESS.TOOL" press-tool-params
macro-function-call =/ "REFTEXT" reftext-params
macro-function-call =/ "REGISTER" register-params
macro-function-call =/ "REGISTER.ID" register-id-params
macro-function-call =/ "RELREF" relref-params
macro-function-call =/ "RENAME.COMMAND" rename-command-params
macro-function-call =/ "REQUEST" request-params
macro-function-call =/ "RESET.TOOLBAR" reset-toolbar-params
macro-function-call =/ "RESTART" restart-params
macro-function-call =/ "RESULT" result-params
macro-function-call =/ "RESUME" resume-params
macro-function-call =/ "RETURN" return-params
macro-function-call =/ "SAVE.DIALOG" save-dialog-params
macro-function-call =/ "SAVE.TOOLBAR" save-toolbar-params
macro-function-call =/ "SET.NAME" set-name-params
macro-function-call =/ "SET.VALUE" set-value-params
macro-function-call =/ "SHOW.BAR" show-bar-params
macro-function-call =/ "SPELLING.CHECK" spelling-check-params
macro-function-call =/ "STEP" step-params
macro-function-call =/ "TERMINATE" terminate-params
macro-function-call =/ "TEXT.BOX" text-box-params
macro-function-call =/ "UNREGISTER" unregister-params
macro-function-call =/ "VOLATILE" volatile-params
macro-function-call =/ "WHILE" while-params
macro-function-call =/ "WINDOW.TITLE" window-title-params
macro-function-call =/ "WINDOWS" windows-params
absref-params = "(" argument "," ref-expression ")"
active-cell-params = "(" *space ")"
add-bar-params = "(" argument ")"
add-command-params = "(" argument "," argument "," argument ["," argument ["," argument]] ")"
add-menu-params = "(" argument "," argument ["," argument ["," argument]] ")"
add-toolbar-params = "(" argument ["," argument] ")"
app-title-params = "(" argument ")"
argument-params = "(" argument ["," argument ["," ref-expression]] ")"
break-params = "(" *space ")"
20 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

call-params = "(" (expression / (argument "," argument *253("," argument))) ")"
caller-params = "(" *space ")"
cancel-key-params = "(" argument ["," ref-expression] ")"
check-command-params = "(" argument "," argument "," argument "," argument ["," argument] ")"
create-object-params = "(" argument "," argument ["," argument ["," argument ["," argument
["," argument ["," argument ["," argument ["," argument ["," argument ["," argument]]]]]]]]]
")"
custom-repeat-params = "(" argument ["," argument ["," argument]] ")"
custom-undo-params = "(" argument ["," argument] ")"
delete-bar-params = "(" expression ")"
delete-command-params = "(" argument "," argument "," argument ["," argument] ")"
delete-menu-params = "(" argument "," argument ["," argument] ")"
delete-toolbar-params = "(" expression ")"
deref-params = "(" ref-expression ")"
dialog-box-params = "(" expression ")"
directory-params = "(" argument ")"
documents-params = "(" argument ["," argument] ")"
echo-params = "(" argument ")"
else-params = "(" *space ")"
else-if-params = "(" expression ")"
enable-command-params = "(" argument "," argument "," argument "," argument ["," argument]
")"
enable-tool-params = "(" argument "," argument "," argument ")"
end-if-params = "(" *space ")"
error-params = "(" argument ["," argument] ")"
evaluate-params = "(" expression ")"
exec-params = "(" (expression / (argument "," argument ["," argument ["," argument]])) ")"
execute-params = "(" argument "," argument ")"
fclose-params = "(" expression ")"
files-params = "(" argument ["," argument] ")"
fopen-params = "(" (expression / (argument "," argument)) ")"
for-params = "(" argument "," argument "," argument ["," argument] ")"
for-cell-params = "(" (expression / (argument "," argument ["," argument])) ")"
formula-convert-params = "(" argument "," argument ["," argument ["," argument [","
argument]]] ")"
fpos-params = "(" (expression / (argument "," argument)) ")"
fread-params = "(" argument "," argument ")"
freadln-params = "(" expression ")"
fsize-params = "(" expression ")"
fwrite-params = "(" argument "," argument ")"
fwriteln-params = "(" argument "," argument ")"
get-bar-params = "(" argument ["," argument ["," argument ["," argument]]] ")"
get-cell-params = "(" (expression / (argument "," ref-expression)) ")"
get-chart-item-params = "(" (expression / (argument "," argument ["," argument])) ")"
get-def-params = "(" (expression / (argument "," argument ["," argument])) ")"
get-document-params = "(" (expression / (argument "," argument)) ")"
get-formula-params = "(" expression ")"
get-link-info-params = "(" argument "," argument ["," argument ["," argument]] ")"
get-movie-params = "(" argument "," argument ["," argument] ")"
get-name-params = "(" (expression / (argument "," argument)) ")"
get-note-params = "(" argument ["," argument ["," argument]] ")"
get-object-params = "(" (expression / (argument "," argument ["," argument ["," argument [","
argument]]])) ")"
get-tool-params = "(" (expression / (argument "," argument ["," argument])) ")"
get-toolbar-params = "(" (expression / (argument "," argument)) ")"
get-window-params = "(" (expression / (argument "," argument)) ")"
get-workbook-params = "(" (expression / (argument "," argument)) ")"
get-workspace-params = "(" expression ")"
goto-params = "(" ref-expression ")"
group-params = "(" *space ")"
halt-params = "(" argument ")"
help-params = "(" argument ")"
initiate-params = "(" argument "," argument ")"
input-params = "(" (expression / (argument "," argument ["," argument ["," argument [","
argument ["," argument ["," argument]]]]])) ")"
last-error-params = "(" *space ")"
links-params = "(" argument ["," argument] ")"
movie-command-params = "(" argument "," argument "," argument ["," argument] ")"
names-params = "(" argument ["," argument ["," argument]] ")"
21 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

next-params = "(" *space ")"
note-params = "(" argument ["," argument ["," argument ["," argument]]] ")"
open-dialog-params = "(" argument ["," argument ["," argument ["," argument]]] ")"
options-lists-get-params = "(" expression ")"
pause-params = "(" argument ")"
poke-params = "(" argument "," argument "," argument ")"
press-tool-params = "(" argument "," argument "," argument ")"
reftext-params = "(" ref-expression "," argument ")"
register-params = "(" (expression / (argument "," argument ["," argument ["," argument [","
argument ["," argument ["," argument ["," argument ["," argument ["," argument *245(","
argument)]]]]]]]])) ")"
register-id-params = "(" argument "," argument ["," argument] ")"
relref-params = "(" ref-expression "," ref-expression ")"
rename-command-params = "(" argument "," argument "," argument "," argument ["," argument]
")"
request-params = "(" argument "," argument ")"
reset-toolbar-params = "(" expression ")"
restart-params = "(" argument ")"
result-params = "(" argument ")"
resume-params = "(" argument ")"
return-params = "(" argument ")"
save-dialog-params = "(" argument ["," argument ["," argument ["," argument ["," argument]]]]
")"
save-toolbar-params = "(" argument ["," argument] ")"
scenario-get-params = "(" (expression / (argument "," argument)) ")"
selection-params = "(" *space ")"
set-name-params = "(" (expression / (argument "," argument)) ")"
set-value-params = "(" ref-expression "," argument ")"
show-bar-params = "(" argument ")"
spelling-check-params = "(" (expression / (argument "," argument ["," argument])) ")"
step-params = "(" *space ")"
terminate-params = "(" expression ")"
text-box-params = "(" (expression / (argument "," argument ["," argument ["," argument]]))
")"
textref-params = "(" (expression / (argument "," argument)) ")"
unregister-params = "(" expression ")"
view-get-params = "(" (expression / (argument "," argument)) ")"
volatile-params = "(" argument ")"
while-params = "(" expression ")"
window-title-params = "(" argument ")"
windows-params = "(" argument ["," argument] ")"
command-function-call = "A1.R1C1" ["?"] a1-r1c1-params
command-function-call =/ "ACTIVATE" ["?"] activate-params
command-function-call =/ "ACTIVATE.NEXT" ["?"] activate-next-params
command-function-call =/ "ACTIVATE.NOTES" ["?"] activate-notes-params
command-function-call =/ "ACTIVATE.PREV" ["?"] activate-prev-params
command-function-call =/ "ACTIVE.CELL.FONT" ["?"] active-cell-font-params
command-function-call =/ "ADD.ARROW" ["?"] add-arrow-params
command-function-call =/ "ADD.CHART.AUTOFORMAT" ["?"] add-chart-autoformat-params
command-function-call =/ "ADD.LIST.ITEM" ["?"] add-list-item-params
command-function-call =/ "ADD.OVERLAY" ["?"] add-overlay-params
command-function-call =/ "ADD.PRINT.AREA" ["?"] add-print-area-params
command-function-call =/ "ADD.TOOL" ["?"] add-tool-params
command-function-call =/ "ADDIN.MANAGER" ["?"] addin-manager-params
command-function-call =/ "ALERT" ["?"] alert-params
command-function-call =/ "ALIGNMENT" ["?"] alignment-params
command-function-call =/ "APP.ACTIVATE" ["?"] app-activate-params
command-function-call =/ "APP.ACTIVATE.MICROSOFT" ["?"] app-activate-microsoft-params
command-function-call =/ "APP.MAXIMIZE" ["?"] app-maximize-params
command-function-call =/ "APP.MINIMIZE" ["?"] app-minimize-params
command-function-call =/ "APP.MOVE" ["?"] app-move-params
command-function-call =/ "APP.RESTORE" ["?"] app-restore-params
command-function-call =/ "APP.SIZE" ["?"] app-size-params
command-function-call =/ "APPLY.NAMES" ["?"] apply-names-params
command-function-call =/ "APPLY.STYLE" ["?"] apply-style-params
command-function-call =/ "ARRANGE.ALL" ["?"] arrange-all-params
command-function-call =/ "ASSIGN.TO.OBJECT" ["?"] assign-to-object-params
command-function-call =/ "ASSIGN.TO.TOOL" ["?"] assign-to-tool-params
22 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

command-function-call =/ "ATTACH.TEXT" ["?"] attach-text-params
command-function-call =/ "ATTACH.TOOLBARS" ["?"] attach-toolbars-params
command-function-call =/ "ATTRIBUTES" ["?"] attributes-params
command-function-call =/ "AUTO.OUTLINE" ["?"] auto-outline-params
command-function-call =/ "AUTOCORRECT" ["?"] autocorrect-params
command-function-call =/ "AXES" ["?"] axes-params
command-function-call =/ "BEEP" ["?"] beep-params
command-function-call =/ "BORDER" ["?"] border-params
command-function-call =/ "BRING.TO.FRONT" ["?"] bring-to-front-params
command-function-call =/ "CALCULATE.DOCUMENT" ["?"] calculate-document-params
command-function-call =/ "CALCULATE.NOW" ["?"] calculate-now-params
command-function-call =/ "CALCULATION" ["?"] calculation-params
command-function-call =/ "CANCEL.COPY" ["?"] cancel-copy-params
command-function-call =/ "CELL.PROTECTION" ["?"] cell-protection-params
command-function-call =/ "CHANGE.LINK" ["?"] change-link-params
command-function-call =/ "CHART.ADD.DATA" ["?"] chart-add-data-params
command-function-call =/ "CHART.TREND" ["?"] chart-trend-params
command-function-call =/ "CHART.WIZARD" ["?"] chart-wizard-params
command-function-call =/ "CHECKBOX.PROPERTIES" ["?"] checkbox-properties-params
command-function-call =/ "CLEAR" ["?"] clear-params
command-function-call =/ "CLEAR.OUTLINE" ["?"] clear-outline-params
command-function-call =/ "CLEAR.PRINT.AREA" ["?"] clear-print-area-params
command-function-call =/ "CLEAR.ROUTING.SLIP" ["?"] clear-routing-slip-params
command-function-call =/ "CLOSE" ["?"] close-params
command-function-call =/ "CLOSE.ALL" ["?"] close-all-params
command-function-call =/ "COLOR.PALETTE" ["?"] color-palette-params
command-function-call =/ "COLUMN.WIDTH" ["?"] column-width-params
command-function-call =/ "COMBINATION" ["?"] combination-params
command-function-call =/ "CONSOLIDATE" ["?"] consolidate-params
command-function-call =/ "CONSTRAIN.NUMERIC" ["?"] constrain-numeric-params
command-function-call =/ "COPY" ["?"] copy-params
command-function-call =/ "COPY.CHART" ["?"] copy-chart-params
command-function-call =/ "COPY.PICTURE" ["?"] copy-picture-params
command-function-call =/ "COPY.TOOL" ["?"] copy-tool-params
command-function-call =/ "CREATE.NAMES" ["?"] create-names-params
command-function-call =/ "CREATE.PUBLISHER" ["?"] create-publisher-params
command-function-call =/ "CUSTOMIZE.TOOLBAR" ["?"] customize-toolbar-params
command-function-call =/ "CUT" ["?"] cut-params
command-function-call =/ "DATA.DELETE" ["?"] data-delete-params
command-function-call =/ "DATA.FIND" ["?"] data-find-params
command-function-call =/ "DATA.FIND.NEXT" ["?"] data-find-next-params
command-function-call =/ "DATA.FIND.PREV" ["?"] data-find-prev-params
command-function-call =/ "DATA.FORM" ["?"] data-form-params
command-function-call =/ "DATA.LABEL" ["?"] data-label-params
command-function-call =/ "DATA.SERIES" ["?"] data-series-params
command-function-call =/ "DEFINE.NAME" ["?"] define-name-params
command-function-call =/ "DEFINE.STYLE" ["?"] define-style-params
command-function-call =/ "DELETE.ARROW" ["?"] delete-arrow-params
command-function-call =/ "DELETE.CHART.AUTOFORMAT" ["?"] delete-chart-autoformat-params
command-function-call =/ "DELETE.FORMAT" ["?"] delete-format-params
command-function-call =/ "DELETE.NAME" ["?"] delete-name-params
command-function-call =/ "DELETE.NOTE" ["?"] delete-note-params
command-function-call =/ "DELETE.OVERLAY" ["?"] delete-overlay-params
command-function-call =/ "DELETE.STYLE" ["?"] delete-style-params
command-function-call =/ "DELETE.TOOL" ["?"] delete-tool-params
command-function-call =/ "DEMOTE" ["?"] demote-params
command-function-call =/ "DISABLE.INPUT" ["?"] disable-input-params
command-function-call =/ "DISPLAY" ["?"] display-params
command-function-call =/ "DUPLICATE" ["?"] duplicate-params
command-function-call =/ "EDIT.COLOR" ["?"] edit-color-params
command-function-call =/ "EDIT.DELETE" ["?"] edit-delete-params
command-function-call =/ "EDIT.OBJECT" ["?"] edit-object-params
command-function-call =/ "EDIT.REPEAT" ["?"] edit-repeat-params
command-function-call =/ "EDIT.SERIES" ["?"] edit-series-params
command-function-call =/ "EDIT.TOOL" ["?"] edit-tool-params
command-function-call =/ "EDITBOX.PROPERTIES" ["?"] editbox-properties-params
command-function-call =/ "EDITION.OPTIONS" ["?"] edition-options-params
command-function-call =/ "ENABLE.OBJECT" ["?"] enable-object-params
command-function-call =/ "ENABLE.TIPWIZARD" ["?"] enable-tipwizard-params
23 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

command-function-call =/ "ENTER.DATA" ["?"] enter-data-params
command-function-call =/ "ERRORBAR.X" ["?"] errorbar-x-params
command-function-call =/ "ERRORBAR.Y" ["?"] errorbar-y-params
command-function-call =/ "EXTEND.POLYGON" ["?"] extend-polygon-params
command-function-call =/ "EXTRACT" ["?"] extract-params
command-function-call =/ "FILE.CLOSE" ["?"] file-close-params
command-function-call =/ "FILE.DELETE" ["?"] file-delete-params
command-function-call =/ "FILL.AUTO" ["?"] fill-auto-params
command-function-call =/ "FILL.DOWN" ["?"] fill-down-params
command-function-call =/ "FILL.GROUP" ["?"] fill-group-params
command-function-call =/ "FILL.LEFT" ["?"] fill-left-params
command-function-call =/ "FILL.RIGHT" ["?"] fill-right-params
command-function-call =/ "FILL.UP" ["?"] fill-up-params
command-function-call =/ "FILTER" ["?"] filter-params
command-function-call =/ "FILTER.ADVANCED" ["?"] filter-advanced-params
command-function-call =/ "FILTER.SHOW.ALL" ["?"] filter-show-all-params
command-function-call =/ "FIND.FILE" ["?"] find-file-params
command-function-call =/ "FONT" ["?"] font-params
command-function-call =/ "FONT.PROPERTIES" ["?"] font-properties-params
command-function-call =/ "FORMAT.AUTO" ["?"] format-auto-params
command-function-call =/ "FORMAT.CHART" ["?"] format-chart-params
command-function-call =/ "FORMAT.CHARTTYPE" ["?"] format-charttype-params
command-function-call =/ "FORMAT.FONT" ["?"] format-font-params
command-function-call =/ "FORMAT.LEGEND" ["?"] format-legend-params
command-function-call =/ "FORMAT.MAIN" ["?"] format-main-params
command-function-call =/ "FORMAT.MOVE" ["?"] format-move-params
command-function-call =/ "FORMAT.NUMBER" ["?"] format-number-params
command-function-call =/ "FORMAT.OVERLAY" ["?"] format-overlay-params
command-function-call =/ "FORMAT.SHAPE" ["?"] format-shape-params
command-function-call =/ "FORMAT.SIZE" ["?"] format-size-params
command-function-call =/ "FORMAT.TEXT" ["?"] format-text-params
command-function-call =/ "FORMULA" ["?"] formula-params
command-function-call =/ "FORMULA.ARRAY" ["?"] formula-array-params
command-function-call =/ "FORMULA.FILL" ["?"] formula-fill-params
command-function-call =/ "FORMULA.FIND" ["?"] formula-find-params
command-function-call =/ "FORMULA.FIND.NEXT" ["?"] formula-find-next-params
command-function-call =/ "FORMULA.FIND.PREV" ["?"] formula-find-prev-params
command-function-call =/ "FORMULA.GOTO" ["?"] formula-goto-params
command-function-call =/ "FORMULA.REPLACE" ["?"] formula-replace-params
command-function-call =/ "FREEZE.PANES" ["?"] freeze-panes-params
command-function-call =/ "FULL" ["?"] full-params
command-function-call =/ "FULL.SCREEN" ["?"] full-screen-params
command-function-call =/ "FUNCTION.WIZARD" ["?"] function-wizard-params
command-function-call =/ "GALLERY.3D.AREA" ["?"] gallery-3d-area-params
command-function-call =/ "GALLERY.3D.BAR" ["?"] gallery-3d-bar-params
command-function-call =/ "GALLERY.3D.COLUMN" ["?"] gallery-3d-column-params
command-function-call =/ "GALLERY.3D.LINE" ["?"] gallery-3d-line-params
command-function-call =/ "GALLERY.3D.PIE" ["?"] gallery-3d-pie-params
command-function-call =/ "GALLERY.3D.SURFACE" ["?"] gallery-3d-surface-params
command-function-call =/ "GALLERY.AREA" ["?"] gallery-area-params
command-function-call =/ "GALLERY.BAR" ["?"] gallery-bar-params
command-function-call =/ "GALLERY.COLUMN" ["?"] gallery-column-params
command-function-call =/ "GALLERY.CUSTOM" ["?"] gallery-custom-params
command-function-call =/ "GALLERY.DOUGHNUT" ["?"] gallery-doughnut-params
command-function-call =/ "GALLERY.LINE" ["?"] gallery-line-params
command-function-call =/ "GALLERY.PIE" ["?"] gallery-pie-params
command-function-call =/ "GALLERY.RADAR" ["?"] gallery-radar-params
command-function-call =/ "GALLERY.SCATTER" ["?"] gallery-scatter-params
command-function-call =/ "GOAL.SEEK" ["?"] goal-seek-params
command-function-call =/ "GRIDLINES" ["?"] gridlines-params
command-function-call =/ "HIDE" ["?"] hide-params
command-function-call =/ "HIDE.DIALOG" ["?"] hide-dialog-params
command-function-call =/ "HIDE.OBJECT" ["?"] hide-object-params
command-function-call =/ "HIDEALL.INKANNOTS" ["?"] hideall-inkannots-params
command-function-call =/ "HIDEALL.NOTES" ["?"] hideall-notes-params
command-function-call =/ "HIDECURR.NOTE" ["?"] hidecurr-note-params
command-function-call =/ "HLINE" ["?"] hline-params
command-function-call =/ "HPAGE" ["?"] hpage-params
command-function-call =/ "HSCROLL" ["?"] hscroll-params
24 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

command-function-call =/ "INSERT" ["?"] insert-params
command-function-call =/ "INSERT.MAP.OBJECT" ["?"] insert-map-object-params
command-function-call =/ "INSERT.OBJECT" ["?"] insert-object-params
command-function-call =/ "INSERT.PICTURE" ["?"] insert-picture-params
command-function-call =/ "INSERT.TITLE" ["?"] insert-title-params
command-function-call =/ "INSERTDATATABLE" ["?"] insertdatatable-params
command-function-call =/ "JUSTIFY" ["?"] justify-params
command-function-call =/ "LABEL.PROPERTIES" ["?"] label-properties-params
command-function-call =/ "LAYOUT" ["?"] layout-params
command-function-call =/ "LEGEND" ["?"] legend-params
command-function-call =/ "LINE.PRINT" ["?"] line-print-params
command-function-call =/ "LINK.COMBO" ["?"] link-combo-params
command-function-call =/ "LINK.FORMAT" ["?"] link-format-params
command-function-call =/ "LIST.NAMES" ["?"] list-names-params
command-function-call =/ "LISTBOX.PROPERTIES" ["?"] listbox-properties-params
command-function-call =/ "MACRO.OPTIONS" ["?"] macro-options-params
command-function-call =/ "MAIL.ADD.MAILER" ["?"] mail-add-mailer-params
command-function-call =/ "MAIL.DELETE.MAILER" ["?"] mail-delete-mailer-params
command-function-call =/ "MAIL.EDIT.MAILER" ["?"] mail-edit-mailer-params
command-function-call =/ "MAIL.FORWARD" ["?"] mail-forward-params
command-function-call =/ "MAIL.LOGOFF" ["?"] mail-logoff-params
command-function-call =/ "MAIL.LOGON" ["?"] mail-logon-params
command-function-call =/ "MAIL.NEXT.LETTER" ["?"] mail-next-letter-params
command-function-call =/ "MAIL.REPLY" ["?"] mail-reply-params
command-function-call =/ "MAIL.REPLY.ALL" ["?"] mail-reply-all-params
command-function-call =/ "MAIL.SEND.MAILER" ["?"] mail-send-mailer-params
command-function-call =/ "MAIN.CHART" ["?"] main-chart-params
command-function-call =/ "MAIN.CHART.TYPE" ["?"] main-chart-type-params
command-function-call =/ "MENU.EDITOR" ["?"] menu-editor-params
command-function-call =/ "MERGE.STYLES" ["?"] merge-styles-params
command-function-call =/ "MESSAGE" ["?"] message-params
command-function-call =/ "MOVE.BRK" ["?"] move-brk-params
command-function-call =/ "MOVE.TOOL" ["?"] move-tool-params
command-function-call =/ "NEW" ["?"] new-params
command-function-call =/ "NEW.WINDOW" ["?"] new-window-params
command-function-call =/ "NEWWEBQUERY" ["?"] newwebquery-params
command-function-call =/ "NORMAL" ["?"] normal-params
command-function-call =/ "OBJECT.PROPERTIES" ["?"] object-properties-params
command-function-call =/ "OBJECT.PROTECTION" ["?"] object-protection-params
command-function-call =/ "ON.DATA" ["?"] on-data-params
command-function-call =/ "ON.DOUBLECLICK" ["?"] on-doubleclick-params
command-function-call =/ "ON.ENTRY" ["?"] on-entry-params
command-function-call =/ "ON.KEY" ["?"] on-key-params
command-function-call =/ "ON.RECALC" ["?"] on-recalc-params
command-function-call =/ "ON.SHEET" ["?"] on-sheet-params
command-function-call =/ "ON.TIME" ["?"] on-time-params
command-function-call =/ "ON.WINDOW" ["?"] on-window-params
command-function-call =/ "OPEN" ["?"] open-params
command-function-call =/ "OPEN.LINKS" ["?"] open-links-params
command-function-call =/ "OPEN.MAIL" ["?"] open-mail-params
command-function-call =/ "OPEN.TEXT" ["?"] open-text-params
command-function-call =/ "OPTIONS.CALCULATION" ["?"] options-calculation-params
command-function-call =/ "OPTIONS.CHART" ["?"] options-chart-params
command-function-call =/ "OPTIONS.EDIT" ["?"] options-edit-params
command-function-call =/ "OPTIONS.GENERAL" ["?"] options-general-params
command-function-call =/ "OPTIONS.LISTS.ADD" ["?"] options-lists-add-params
command-function-call =/ "OPTIONS.LISTS.DELETE" ["?"] options-lists-delete-params
command-function-call =/ "OPTIONS.ME" ["?"] options-me-params
command-function-call =/ "OPTIONS.MENONO" ["?"] options-menono-params
command-function-call =/ "OPTIONS.SAVE" ["?"] options-save-params
command-function-call =/ "OPTIONS.SPELL" ["?"] options-spell-params
command-function-call =/ "OPTIONS.TRANSITION" ["?"] options-transition-params
command-function-call =/ "OPTIONS.VIEW" ["?"] options-view-params
command-function-call =/ "OUTLINE" ["?"] outline-params
command-function-call =/ "OVERLAY" ["?"] overlay-params
command-function-call =/ "OVERLAY.CHART.TYPE" ["?"] overlay-chart-type-params
command-function-call =/ "PAGE.SETUP" ["?"] page-setup-params
command-function-call =/ "PARSE" ["?"] parse-params
command-function-call =/ "PASTE" ["?"] paste-params
25 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

command-function-call =/ "PASTE.LINK" ["?"] paste-link-params
command-function-call =/ "PASTE.PICTURE" ["?"] paste-picture-params
command-function-call =/ "PASTE.PICTURE.LINK" ["?"] paste-picture-link-params
command-function-call =/ "PASTE.SPECIAL" ["?"] paste-special-params
command-function-call =/ "PASTE.TOOL" ["?"] paste-tool-params
command-function-call =/ "PATTERNS" ["?"] patterns-params
command-function-call =/ "PICKLIST" ["?"] picklist-params
command-function-call =/ "PIVOT.ADD.FIELDS" ["?"] pivot-add-fields-params
command-function-call =/ "PIVOT.FIELD" ["?"] pivot-field-params
command-function-call =/ "PIVOT.FIELD.GROUP" ["?"] pivot-field-group-params
command-function-call =/ "PIVOT.FIELD.PROPERTIES" ["?"] pivot-field-properties-params
command-function-call =/ "PIVOT.FIELD.UNGROUP" ["?"] pivot-field-ungroup-params
command-function-call =/ "PIVOT.ITEM" ["?"] pivot-item-params
command-function-call =/ "PIVOT.ITEM.PROPERTIES" ["?"] pivot-item-properties-params
command-function-call =/ "PIVOT.REFRESH" ["?"] pivot-refresh-params
command-function-call =/ "PIVOT.SHOW.PAGES" ["?"] pivot-show-pages-params
command-function-call =/ "PIVOT.TABLE.CHART" ["?"] pivot-table-chart-params
command-function-call =/ "PIVOT.TABLE.WIZARD" ["?"] pivot-table-wizard-params
command-function-call =/ "POST.DOCUMENT" ["?"] post-document-params
command-function-call =/ "PRECISION" ["?"] precision-params
command-function-call =/ "PREFERRED" ["?"] preferred-params
command-function-call =/ "PRINT" ["?"] print-params
command-function-call =/ "PRINT.PREVIEW" ["?"] print-preview-params
command-function-call =/ "PRINTER.SETUP" ["?"] printer-setup-params
command-function-call =/ "PROMOTE" ["?"] promote-params
command-function-call =/ "PROTECT.DOCUMENT" ["?"] protect-document-params
command-function-call =/ "PROTECT.REVISIONS" ["?"] protect-revisions-params
command-function-call =/ "PUSHBUTTON.PROPERTIES" ["?"] pushbutton-properties-params
command-function-call =/ "QUIT" ["?"] quit-params
command-function-call =/ "REMOVE.LIST.ITEM" ["?"] remove-list-item-params
command-function-call =/ "REMOVE.PAGE.BREAK" ["?"] remove-page-break-params
command-function-call =/ "RENAME.OBJECT" ["?"] rename-object-params
command-function-call =/ "REPLACE.FONT" ["?"] replace-font-params
command-function-call =/ "RESET.TOOL" ["?"] reset-tool-params
command-function-call =/ "RM.PRINT.AREA" ["?"] rm-print-area-params
command-function-call =/ "ROUTE.DOCUMENT" ["?"] route-document-params
command-function-call =/ "ROUTING.SLIP" ["?"] routing-slip-params
command-function-call =/ "ROW.HEIGHT" ["?"] row-height-params
command-function-call =/ "RUN" ["?"] run-params
command-function-call =/ "SAVE" ["?"] save-params
command-function-call =/ "SAVE.AS" ["?"] save-as-params
command-function-call =/ "SAVE.COPY.AS" ["?"] save-copy-as-params
command-function-call =/ "SAVE.NEW.OBJECT" ["?"] save-new-object-params
command-function-call =/ "SAVE.WORKBOOK" ["?"] save-workbook-params
command-function-call =/ "SAVE.WORKSPACE" ["?"] save-workspace-params
command-function-call =/ "SCALE" ["?"] scale-params
command-function-call =/ "SCENARIO.ADD" ["?"] scenario-add-params
command-function-call =/ "SCENARIO.CELLS" ["?"] scenario-cells-params
command-function-call =/ "SCENARIO.DELETE" ["?"] scenario-delete-params
command-function-call =/ "SCENARIO.EDIT" ["?"] scenario-edit-params
command-function-call =/ "SCENARIO.MERGE" ["?"] scenario-merge-params
command-function-call =/ "SCENARIO.SHOW" ["?"] scenario-show-params
command-function-call =/ "SCENARIO.SHOW.NEXT" ["?"] scenario-show-next-params
command-function-call =/ "SCENARIO.SUMMARY" ["?"] scenario-summary-params
command-function-call =/ "SCROLLBAR.PROPERTIES" ["?"] scrollbar-properties-params
command-function-call =/ "SELECT" ["?"] select-params
command-function-call =/ "SELECT.ALL" ["?"] select-all-params
command-function-call =/ "SELECT.CHART" ["?"] select-chart-params
command-function-call =/ "SELECT.END" ["?"] select-end-params
command-function-call =/ "SELECT.LAST.CELL" ["?"] select-last-cell-params
command-function-call =/ "SELECT.LIST.ITEM" ["?"] select-list-item-params
command-function-call =/ "SELECT.PLOT.AREA" ["?"] select-plot-area-params
command-function-call =/ "SELECT.SPECIAL" ["?"] select-special-params
command-function-call =/ "SEND.KEYS" ["?"] send-keys-params
command-function-call =/ "SEND.MAIL" ["?"] send-mail-params
command-function-call =/ "SEND.TO.BACK" ["?"] send-to-back-params
command-function-call =/ "SERIES.AXES" ["?"] series-axes-params
command-function-call =/ "SERIES.ORDER" ["?"] series-order-params
command-function-call =/ "SERIES.X" ["?"] series-x-params
26 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

command-function-call =/ "SERIES.Y" ["?"] series-y-params
command-function-call =/ "SET.CONTROL.VALUE" ["?"] set-control-value-params
command-function-call =/ "SET.CRITERIA" ["?"] set-criteria-params
command-function-call =/ "SET.DATABASE" ["?"] set-database-params
command-function-call =/ "SET.DIALOG.DEFAULT" ["?"] set-dialog-default-params
command-function-call =/ "SET.DIALOG.FOCUS" ["?"] set-dialog-focus-params
command-function-call =/ "SET.EXTRACT" ["?"] set-extract-params
command-function-call =/ "SET.LIST.ITEM" ["?"] set-list-item-params
command-function-call =/ "SET.PAGE.BREAK" ["?"] set-page-break-params
command-function-call =/ "SET.PREFERRED" ["?"] set-preferred-params
command-function-call =/ "SET.PRINT.AREA" ["?"] set-print-area-params
command-function-call =/ "SET.PRINT.TITLES" ["?"] set-print-titles-params
command-function-call =/ "SET.UPDATE.STATUS" ["?"] set-update-status-params
command-function-call =/ "SHARE" ["?"] share-params
command-function-call =/ "SHARE.NAME" ["?"] share-name-params
command-function-call =/ "SHEET.BACKGROUND" ["?"] sheet-background-params
command-function-call =/ "SHORT.MENUS" ["?"] short-menus-params
command-function-call =/ "SHOW.ACTIVE.CELL" ["?"] show-active-cell-params
command-function-call =/ "SHOW.CLIPBOARD" ["?"] show-clipboard-params
command-function-call =/ "SHOW.DETAIL" ["?"] show-detail-params
command-function-call =/ "SHOW.DIALOG" ["?"] show-dialog-params
command-function-call =/ "SHOW.INFO" ["?"] show-info-params
command-function-call =/ "SHOW.LEVELS" ["?"] show-levels-params
command-function-call =/ "SHOW.TOOLBAR" ["?"] show-toolbar-params
command-function-call =/ "SORT" ["?"] sort-params
command-function-call =/ "SORT.SPECIAL" ["?"] sort-special-params
command-function-call =/ "SOUND.NOTE" ["?"] sound-note-params
command-function-call =/ "SOUND.PLAY" ["?"] sound-play-params
command-function-call =/ "SPELLING" ["?"] spelling-params
command-function-call =/ "SPLIT" ["?"] split-params
command-function-call =/ "STANDARD.FONT" ["?"] standard-font-params
command-function-call =/ "STANDARD.WIDTH" ["?"] standard-width-params
command-function-call =/ "STYLE" ["?"] style-params
command-function-call =/ "SUBSCRIBE.TO" ["?"] subscribe-to-params
command-function-call =/ "SUBTOTAL.CREATE" ["?"] subtotal-create-params
command-function-call =/ "SUBTOTAL.REMOVE" ["?"] subtotal-remove-params
command-function-call =/ "SUMMARY.INFO" ["?"] summary-info-params
command-function-call =/ "TAB.ORDER" ["?"] tab-order-params
command-function-call =/ "TABLE" ["?"] table-params
command-function-call =/ "TEXT.TO.COLUMNS" ["?"] text-to-columns-params
command-function-call =/ "TRACER.CLEAR" ["?"] tracer-clear-params
command-function-call =/ "TRACER.DISPLAY" ["?"] tracer-display-params
command-function-call =/ "TRACER.ERROR" ["?"] tracer-error-params
command-function-call =/ "TRACER.NAVIGATE" ["?"] tracer-navigate-params
command-function-call =/ "TRAVERSE.NOTES" ["?"] traverse-notes-params
command-function-call =/ "UNDO" ["?"] undo-params
command-function-call =/ "UNGROUP" ["?"] ungroup-params
command-function-call =/ "UNGROUP.SHEETS" ["?"] ungroup-sheets-params
command-function-call =/ "UNHIDE" ["?"] unhide-params
command-function-call =/ "UNLOCKED.NEXT" ["?"] unlocked-next-params
command-function-call =/ "UNLOCKED.PREV" ["?"] unlocked-prev-params
command-function-call =/ "UNPROTECT.REVISIONS" ["?"] unprotect-revisions-params
command-function-call =/ "UPDATE.LINK" ["?"] update-link-params
command-function-call =/ "VBA.INSERT.FILE" ["?"] vba-insert-file-params
command-function-call =/ "VBA.MAKE.ADDIN" ["?"] vba-make-addin-params
command-function-call =/ "VBA.PROCEDURE.DEFINITION" ["?"] vba-procedure-definition-params
command-function-call =/ "VBAActivate" ["?"] vbaactivate-params
command-function-call =/ "VIEW.3D" ["?"] view-3d-params
command-function-call =/ "VIEW.DEFINE" ["?"] view-define-params
command-function-call =/ "VIEW.DELETE" ["?"] view-delete-params
command-function-call =/ "VIEW.SHOW" ["?"] view-show-params
command-function-call =/ "VLINE" ["?"] vline-params
command-function-call =/ "VPAGE" ["?"] vpage-params
command-function-call =/ "VSCROLL" ["?"] vscroll-params
command-function-call =/ "WAIT" ["?"] wait-params
command-function-call =/ "WEB.PUBLISH" ["?"] web-publish-params
command-function-call =/ "WINDOW.MAXIMIZE" ["?"] window-maximize-params
command-function-call =/ "WINDOW.MINIMIZE" ["?"] window-minimize-params
command-function-call =/ "WINDOW.MOVE" ["?"] window-move-params
27 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

command-function-call =/ "WINDOW.RESTORE" ["?"] window-restore-params
command-function-call =/ "WINDOW.SIZE" ["?"] window-size-params
command-function-call =/ "WORKBOOK.ACTIVATE" ["?"] workbook-activate-params
command-function-call =/ "WORKBOOK.ADD" ["?"] workbook-add-params
command-function-call =/ "WORKBOOK.COPY" ["?"] workbook-copy-params
command-function-call =/ "WORKBOOK.DELETE" ["?"] workbook-delete-params
command-function-call =/ "WORKBOOK.HIDE" ["?"] workbook-hide-params
command-function-call =/ "WORKBOOK.INSERT" ["?"] workbook-insert-params
command-function-call =/ "WORKBOOK.MOVE" ["?"] workbook-move-params
command-function-call =/ "WORKBOOK.NAME" ["?"] workbook-name-params
command-function-call =/ "WORKBOOK.NEW" ["?"] workbook-new-params
command-function-call =/ "WORKBOOK.NEXT" ["?"] workbook-next-params
command-function-call =/ "WORKBOOK.OPTIONS" ["?"] workbook-options-params
command-function-call =/ "WORKBOOK.PREV" ["?"] workbook-prev-params
command-function-call =/ "WORKBOOK.PROTECT" ["?"] workbook-protect-params
command-function-call =/ "WORKBOOK.SCROLL" ["?"] workbook-scroll-params
command-function-call =/ "WORKBOOK.SELECT" ["?"] workbook-select-params
command-function-call =/ "WORKBOOK.TAB.SPLIT" ["?"] workbook-tab-split-params
command-function-call =/ "WORKBOOK.UNHIDE" ["?"] workbook-unhide-params
command-function-call =/ "WORKGROUP" ["?"] workgroup-params
command-function-call =/ "WORKGROUP.OPTIONS" ["?"] workgroup-options-params
command-function-call =/ "WORKSPACE" ["?"] workspace-params
command-function-call =/ "ZOOM" ["?"] zoom-params
a1-r1c1-params = "(" argument ")"
activate-params = "(" argument ["," argument] ")"
activate-next-params = "(" argument ")"
activate-notes-params = "(" argument ["," argument] ")"
activate-prev-params = "(" argument ")"
active-cell-font-params = "(" argument *13("," argument) ")"
add-arrow-params = "(" *space ")"
add-chart-autoformat-params = "(" argument ["," argument] ")"
add-list-item-params = "(" argument ["," argument] ")"
add-overlay-params = "(" *space ")"
add-print-area-params = "(" *space ")"
add-tool-params = "(" argument *2("," argument) ")"
addin-manager-params = "(" argument *2("," argument) ")"
alert-params = "(" argument *2("," argument) ")"
alignment-params = "(" argument *9("," argument) ")"
app-activate-params = "(" argument ["," argument] ")"
app-activate-microsoft-params = "(" argument ")"
app-maximize-params = "(" *space ")"
app-minimize-params = "(" *space ")"
app-move-params = "(" argument ["," argument] ")"
app-restore-params = "(" *space ")"
app-size-params = "(" argument ["," argument] ")"
apply-names-params = "(" argument *6("," argument) ")"
apply-style-params = "(" argument ")"
arrange-all-params = "(" argument *3("," argument) ")"
assign-to-object-params = "(" argument ")"
assign-to-tool-params = "(" argument *2("," argument) ")"
attach-text-params = "(" argument *2("," argument) ")"
attach-toolbars-params = "(" *space ")"
attributes-params = "(" argument ["," argument] ")"
auto-outline-params = "(" *space ")"
autocorrect-params = "(" argument ["," argument] ")"
axes-params = "(" argument *5("," argument) ")"
beep-params = "(" argument ")"
border-params = "(" argument *26("," argument) ")"
bring-to-front-params = "(" *space ")"
calculate-document-params = "(" *space ")"
calculate-now-params = "(" *space ")"
calculation-params = "(" argument *10("," argument) ")"
cancel-copy-params = "(" argument ")"
cell-protection-params = "(" argument ["," argument] ")"
change-link-params = "(" argument *2("," argument) ")"
chart-add-data-params = "(" argument *5("," argument) ")"
chart-trend-params = "(" argument *7("," argument) ")"
chart-wizard-params = "(" argument *13("," argument) ")"
28 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

checkbox-properties-params = "(" argument *4("," argument) ")"
clear-params = "(" argument ")"
clear-outline-params = "(" *space ")"
clear-print-area-params = "(" *space ")"
clear-routing-slip-params = "(" argument ")"
close-params = "(" argument ["," argument] ")"
close-all-params = "(" *space ")"
color-palette-params = "(" argument ")"
column-width-params = "(" argument *4("," argument) ")"
combination-params = "(" argument ")"
consolidate-params = "(" argument *4("," argument) ")"
constrain-numeric-params = "(" argument ")"
copy-params = "(" argument ["," argument] ")"
copy-chart-params = "(" argument ")"
copy-picture-params = "(" argument *2("," argument) ")"
copy-tool-params = "(" argument ["," argument] ")"
create-names-params = "(" argument *3("," argument) ")"
create-publisher-params = "(" argument *3("," argument) ")"
customize-toolbar-params = "(" argument ")"
cut-params = "(" argument ["," argument] ")"
data-delete-params = "(" *space ")"
data-find-params = "(" argument ")"
data-find-next-params = "(" *space ")"
data-find-prev-params = "(" *space ")"
data-form-params = "(" *space ")"
data-label-params = "(" argument *9("," argument) ")"
data-series-params = "(" argument *5("," argument) ")"
define-name-params = "(" argument *6("," argument) ")"
define-style-params = "(" argument *13("," argument) ")"
delete-arrow-params = "(" *space ")"
delete-chart-autoformat-params = "(" argument ")"
delete-format-params = "(" argument ")"
delete-name-params = "(" argument ")"
delete-note-params = "(" argument ")"
delete-overlay-params = "(" *space ")"
delete-style-params = "(" argument ")"
delete-tool-params = "(" argument ["," argument] ")"
demote-params = "(" argument ")"
disable-input-params = "(" argument ")"
display-params = "(" argument *8("," argument) ")"
duplicate-params = "(" *space ")"
edit-color-params = "(" argument *3("," argument) ")"
edit-delete-params = "(" argument ")"
edit-object-params = "(" argument ")"
edit-repeat-params = "(" *space ")"
edit-series-params = "(" argument *6("," argument) ")"
edit-tool-params = "(" argument ["," argument] ")"
editbox-properties-params = "(" argument *3("," argument) ")"
edition-options-params = "(" argument *6("," argument) ")"
enable-object-params = "(" argument ["," argument] ")"
enable-tipwizard-params = "(" argument ")"
enter-data-params = "(" argument ")"
errorbar-x-params = "(" argument *3("," argument) ")"
errorbar-y-params = "(" argument *3("," argument) ")"
extend-polygon-params = "(" argument ")"
extract-params = "(" argument ")"
file-close-params = "(" argument ["," argument] ")"
file-delete-params = "(" argument ")"
fill-auto-params = "(" argument ["," argument] ")"
fill-down-params = "(" *space ")"
fill-group-params = "(" argument ")"
fill-left-params = "(" *space ")"
fill-right-params = "(" *space ")"
fill-up-params = "(" *space ")"
filter-params = "(" argument *5("," argument) ")"
filter-advanced-params = "(" argument *4("," argument) ")"
filter-show-all-params = "(" *space ")"
find-file-params = "(" *space ")"
font-params = "(" argument ["," argument] ")"
29 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

font-properties-params = "(" argument *13("," argument) ")"
format-auto-params = "(" argument *6("," argument) ")"
format-chart-params = "(" argument *17("," argument) ")"
format-charttype-params = "(" argument *3("," argument) ")"
format-font-params = "(" argument *14("," argument) ")"
format-legend-params = "(" argument ")"
format-main-params = "(" argument *13("," argument) ")"
format-move-params = "(" argument *2("," argument) ")"
format-number-params = "(" argument ")"
format-overlay-params = "(" argument *13("," argument) ")"
format-shape-params = "(" argument *4("," argument) ")"
format-size-params = "(" argument *2("," argument) ")"
format-text-params = "(" argument *10("," argument) ")"
formula-params = "(" argument ["," argument] ")"
formula-array-params = "(" argument ["," argument] ")"
formula-fill-params = "(" argument ["," argument] ")"
formula-find-params = "(" argument *11("," argument) ")"
formula-find-next-params = "(" *space ")"
formula-find-prev-params = "(" *space ")"
formula-goto-params = "(" argument ["," argument] ")"
formula-replace-params = "(" argument *10("," argument) ")"
freeze-panes-params = "(" argument *2("," argument) ")"
full-params = "(" argument ")"
full-screen-params = "(" argument ")"
function-wizard-params = "(" argument ")"
gallery-3d-area-params = "(" argument ")"
gallery-3d-bar-params = "(" argument ")"
gallery-3d-column-params = "(" argument ")"
gallery-3d-line-params = "(" argument ")"
gallery-3d-pie-params = "(" argument ")"
gallery-3d-surface-params = "(" argument ")"
gallery-area-params = "(" argument ["," argument] ")"
gallery-bar-params = "(" argument ["," argument] ")"
gallery-column-params = "(" argument ["," argument] ")"
gallery-custom-params = "(" argument ")"
gallery-doughnut-params = "(" argument ["," argument] ")"
gallery-line-params = "(" argument ["," argument] ")"
gallery-pie-params = "(" argument ["," argument] ")"
gallery-radar-params = "(" argument ["," argument] ")"
gallery-scatter-params = "(" argument ["," argument] ")"
goal-seek-params = "(" argument *2("," argument) ")"
gridlines-params = "(" argument *6("," argument) ")"
hide-params = "(" *space ")"
hide-dialog-params = "(" argument ")"
hide-object-params = "(" argument ["," argument] ")"
hideall-inkannots-params = "(" argument ")"
hideall-notes-params = "(" argument ")"
hidecurr-note-params = "(" argument ["," argument] ")"
hline-params = "(" argument ")"
hpage-params = "(" argument ")"
hscroll-params = "(" argument ["," argument] ")"
insert-params = "(" argument ["," argument] ")"
insert-map-object-params = "(" *space ")"
insert-object-params = "(" argument *12("," argument) ")"
insert-picture-params = "(" argument ["," argument] ")"
insert-title-params = "(" argument *4("," argument) ")"
insertdatatable-params = "(" argument ")"
justify-params = "(" *space ")"
label-properties-params = "(" argument *2("," argument) ")"
layout-params = "(" *space ")"
legend-params = "(" argument ")"
line-print-params = "(" argument *10("," argument) ")"
link-combo-params = "(" argument ")"
link-format-params = "(" *space ")"
list-names-params = "(" *space ")"
listbox-properties-params = "(" argument *4("," argument) ")"
macro-options-params = "(" argument *9("," argument) ")"
mail-add-mailer-params = "(" *space ")"
mail-delete-mailer-params = "(" *space ")"
30 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

mail-edit-mailer-params = "(" argument *5("," argument) ")"
mail-forward-params = "(" *space ")"
mail-logoff-params = "(" *space ")"
mail-logon-params = "(" argument *2("," argument) ")"
mail-next-letter-params = "(" *space ")"
mail-reply-params = "(" *space ")"
mail-reply-all-params = "(" *space ")"
mail-send-mailer-params = "(" argument ["," argument] ")"
main-chart-params = "(" argument *9("," argument) ")"
main-chart-type-params = "(" argument ")"
menu-editor-params = "(" *space ")"
merge-styles-params = "(" argument ")"
message-params = "(" argument ["," argument] ")"
move-brk-params = "(" argument *3("," argument) ")"
move-tool-params = "(" argument *5("," argument) ")"
new-params = "(" argument *2("," argument) ")"
new-window-params = "(" *space ")"
newwebquery-params = "(" argument ")"
normal-params = "(" *space ")"
object-properties-params = "(" argument ["," argument] ")"
object-protection-params = "(" argument ["," argument] ")"
on-data-params = "(" argument ["," argument] ")"
on-doubleclick-params = "(" argument ["," argument] ")"
on-entry-params = "(" argument ["," argument] ")"
on-key-params = "(" argument ["," argument] ")"
on-recalc-params = "(" argument ["," argument] ")"
on-sheet-params = "(" argument *2("," argument) ")"
on-time-params = "(" argument *3("," argument) ")"
on-window-params = "(" argument ["," argument] ")"
open-params = "(" argument *16("," argument) ")"
open-links-params = "(" argument *14("," argument) ")"
open-mail-params = "(" argument ["," argument] ")"
open-text-params = "(" argument *16("," argument) ")"
options-calculation-params = "(" argument *9("," argument) ")"
options-chart-params = "(" argument *2("," argument) ")"
options-edit-params = "(" argument *10("," argument) ")"
options-general-params = "(" argument *13("," argument) ")"
options-lists-add-params = "(" argument ["," argument] ")"
options-lists-delete-params = "(" argument ")"
options-me-params = "(" argument *8("," argument) ")"
options-menono-params = "(" argument *4("," argument) ")"
options-save-params = "(" argument *3("," argument) ")"
options-spell-params = "(" argument *11("," argument) ")"
options-transition-params = "(" argument *4("," argument) ")"
options-view-params = "(" argument *17("," argument) ")"
outline-params = "(" argument *3("," argument) ")"
overlay-params = "(" argument *11("," argument) ")"
overlay-chart-type-params = "(" argument ")"
page-setup-params = "(" argument *29("," argument) ")"
parse-params = "(" argument ["," argument] ")"
paste-params = "(" argument ")"
paste-link-params = "(" *space ")"
paste-picture-params = "(" *space ")"
paste-picture-link-params = "(" *space ")"
paste-special-params = "(" argument *6("," argument) ")"
paste-tool-params = "(" argument ["," argument] ")"
patterns-params = "(" argument *12("," argument) ")"
picklist-params = "(" *space ")"
pivot-add-fields-params = "(" argument *4("," argument) ")"
pivot-field-params = "(" argument *3("," argument) ")"
pivot-field-group-params = "(" argument *3("," argument) ")"
pivot-field-properties-params = "(" argument *6("," argument) ")"
pivot-field-ungroup-params = "(" *space ")"
pivot-item-params = "(" argument *3("," argument) ")"
pivot-item-properties-params = "(" argument *6("," argument) ")"
pivot-refresh-params = "(" argument ")"
pivot-show-pages-params = "(" argument ["," argument] ")"
pivot-table-chart-params = "(" argument *15("," argument) ")"
pivot-table-wizard-params = "(" argument *15("," argument) ")"
31 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

post-document-params = "(" argument ")"
precision-params = "(" argument ")"
preferred-params = "(" *space ")"
print-params = "(" argument *16("," argument) ")"
print-preview-params = "(" argument ")"
printer-setup-params = "(" argument ")"
promote-params = "(" argument ")"
protect-document-params = "(" argument *6("," argument) ")"
protect-revisions-params = "(" *space ")"
pushbutton-properties-params = "(" argument *5("," argument) ")"
quit-params = "(" *space ")"
remove-list-item-params = "(" argument ["," argument] ")"
remove-page-break-params = "(" argument ["," argument] ")"
rename-object-params = "(" argument ")"
replace-font-params = "(" argument *9("," argument) ")"
reset-tool-params = "(" argument ["," argument] ")"
rm-print-area-params = "(" argument ")"
route-document-params = "(" *space ")"
routing-slip-params = "(" argument *5("," argument) ")"
row-height-params = "(" argument *3("," argument) ")"
run-params = "(" argument ["," argument] ")"
save-params = "(" *space ")"
save-as-params = "(" argument *6("," argument) ")"
save-copy-as-params = "(" argument ")"
save-new-object-params = "(" argument ")"
save-workbook-params = "(" argument *5("," argument) ")"
save-workspace-params = "(" argument ")"
scale-params = "(" argument *9("," argument) ")"
scenario-add-params = "(" argument *5("," argument) ")"
scenario-cells-params = "(" argument ")"
scenario-delete-params = "(" argument ")"
scenario-edit-params = "(" argument *6("," argument) ")"
scenario-merge-params = "(" argument ")"
scenario-show-params = "(" argument ")"
scenario-show-next-params = "(" *space ")"
scenario-summary-params = "(" argument ["," argument] ")"
scrollbar-properties-params = "(" argument *6("," argument) ")"
select-params = "(" argument ["," argument] ")"
select-all-params = "(" *space ")"
select-chart-params = "(" *space ")"
select-end-params = "(" argument ")"
select-last-cell-params = "(" *space ")"
select-list-item-params = "(" argument ["," argument] ")"
select-plot-area-params = "(" *space ")"
select-special-params = "(" argument *2("," argument) ")"
send-keys-params = "(" argument ["," argument] ")"
send-mail-params = "(" argument *2("," argument) ")"
send-to-back-params = "(" *space ")"
series-axes-params = "(" argument ")"
series-order-params = "(" argument *2("," argument) ")"
series-x-params = "(" argument ")"
series-y-params = "(" argument ["," argument] ")"
set-control-value-params = "(" argument ")"
set-criteria-params = "(" *space ")"
set-database-params = "(" *space ")"
set-dialog-default-params = "(" argument ")"
set-dialog-focus-params = "(" argument ")"
set-extract-params = "(" *space ")"
set-list-item-params = "(" argument ["," argument] ")"
set-page-break-params = "(" *space ")"
set-preferred-params = "(" argument ")"
set-print-area-params = "(" argument ")"
set-print-titles-params = "(" argument ["," argument] ")"
set-update-status-params = "(" argument *2("," argument) ")"
share-params = "(" *space ")"
share-name-params = "(" argument ")"
sheet-background-params = "(" argument ["," argument] ")"
short-menus-params = "(" argument ")"
show-active-cell-params = "(" *space ")"
32 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

show-clipboard-params = "(" *space ")"
show-detail-params = "(" argument *3("," argument) ")"
show-dialog-params = "(" argument ")"
show-info-params = "(" argument ")"
show-levels-params = "(" argument ["," argument] ")"
show-toolbar-params = "(" argument *9("," argument) ")"
sort-params = "(" argument *16("," argument) ")"
sort-special-params = "(" argument *13("," argument) ")"
sound-note-params = "(" argument *2("," argument) ")"
sound-play-params = "(" argument *2("," argument) ")"
spelling-params = "(" argument *5("," argument) ")"
split-params = "(" argument ["," argument] ")"
standard-font-params = "(" argument *8("," argument) ")"
standard-width-params = "(" argument ")"
style-params = "(" argument ["," argument] ")"
subscribe-to-params = "(" argument ["," argument] ")"
subtotal-create-params = "(" argument *5("," argument) ")"
subtotal-remove-params = "(" *space ")"
summary-info-params = "(" argument *4("," argument) ")"
tab-order-params = "(" *space ")"
table-params = "(" argument ["," argument] ")"
text-to-columns-params = "(" argument *13("," argument) ")"
tracer-clear-params = "(" *space ")"
tracer-display-params = "(" argument ["," argument] ")"
tracer-error-params = "(" *space ")"
tracer-navigate-params = "(" argument *2("," argument) ")"
traverse-notes-params = "(" argument ["," argument] ")"
undo-params = "(" *space ")"
ungroup-params = "(" *space ")"
ungroup-sheets-params = "(" *space ")"
unhide-params = "(" argument ")"
unlocked-next-params = "(" *space ")"
unlocked-prev-params = "(" *space ")"
unprotect-revisions-params = "(" *space ")"
update-link-params = "(" argument ["," argument] ")"
vba-insert-file-params = "(" argument ")"
vba-make-addin-params = "(" argument ")"
vba-procedure-definition-params = "(" *space ")"
vbaactivate-params = "(" argument ["," argument] ")"
view-3d-params = "(" argument *5("," argument) ")"
view-define-params = "(" argument *2("," argument) ")"
view-delete-params = "(" argument ")"
view-show-params = "(" argument ")"
vline-params = "(" argument ")"
vpage-params = "(" argument ")"
vscroll-params = "(" argument ["," argument] ")"
wait-params = "(" argument ")"
web-publish-params = "(" argument *8("," argument) ")"
window-maximize-params = "(" argument ")"
window-minimize-params = "(" argument ")"
window-move-params = "(" argument *2("," argument) ")"
window-restore-params = "(" argument ")"
window-size-params = "(" argument *2("," argument) ")"
workbook-activate-params = "(" argument ["," argument] ")"
workbook-add-params = "(" argument *2("," argument) ")"
workbook-copy-params = "(" argument *2("," argument) ")"
workbook-delete-params = "(" argument ")"
workbook-hide-params = "(" argument ["," argument] ")"
workbook-insert-params = "(" argument ")"
workbook-move-params = "(" argument *2("," argument) ")"
workbook-name-params = "(" argument ["," argument] ")"
workbook-new-params = "(" argument *2("," argument) ")"
workbook-next-params = "(" *space ")"
workbook-options-params = "(" argument *2("," argument) ")"
workbook-prev-params = "(" *space ")"
workbook-protect-params = "(" argument *2("," argument) ")"
workbook-scroll-params = "(" argument ["," argument] ")"
workbook-select-params = "(" argument *2("," argument) ")"
workbook-tab-split-params = "(" argument ")"
33 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

workbook-unhide-params = "(" argument ")"
workgroup-params = "(" argument ")"
workgroup-options-params = "(" *space ")"
workspace-params = "(" argument *15("," argument) ")"
zoom-params = "(" argument ")"
2.3 Macro-Enabled PresentationML Structures
2.3.1 Part Enumeration
This section specifies the parts ([ECMA-376]) of the macro-enabled PresentationML file format that are
different from the parts of the non-macro-enabled PresentationML file format specified in [ECMA-376]
Part 1, Section 13.3. These parts and their relationships are summarized in the following table.
Part Relationship target
Presentation (section 2.3.1.1) PresentationML Package ([ECMA-376] Part 1, Section 13.2)
VBA Project (section 2.3.1.2) Presentation
2.3.1.1 Presentation
Content Type(s):
 application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml
 application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml
 application/vnd.ms-powerpoint.template.macroEnabled.main+xml
Root Namespace: http://schemas.openxmlformats.org/presentationml/2006/main
Source Relationship:
http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument
This part is specified in [ECMA-376] Part 1, Section 13.3.6.
2.3.1.2 VBA Project
Content Type: application/vnd.ms-office.vbaProject
Root Namespace: Not applicable.
Source Relationship: http://schemas.microsoft.com/office/2006/relationships/vbaProject
This part is specified in [MS-OVBA].
A package MUST contain at most one VBA Project part, which MUST be the target of an implicit
relationship from the presentation part (section 2.3.1.1).
A VBA Project part MUST NOT have implicit or explicit relationships to any other part specified by this
document or any part specified in [ECMA-376].
34 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

2.3.2 Global Elements
None.
2.3.3 Global Attributes
None.
2.3.4 Complex Types
2.3.4.1 CT_SlideMasterIdListEntry
Target namespace: http://schemas.openxmlformats.org/presentationml/2006/main
This element specifies a slide master that is available within the corresponding presentation. A slide
master is a slide that is specifically designed to be a template for all related child layout slides.
This complex type is modified from the CT_SlideMasterIdListEntry specified in [ECMA-376] Part 3,
section 4.1.3.1.
The id attribute of this part MUST be present in the macro-enabled PresentationML format.
2.3.5 Simple Types
None.
35 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

3 Structure Examples
3.1 VBA Supplemental Data Part of a Word-Processing Document
This example shows a usage of a Word VBA Supplemental Data part (section 2.1.1.3). See section
2.1.1.3 for more information. Consider the following XML, showing the complete contents of the VBA
Supplemental Data part of a word-processing document.
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
The docEvents child of the vbaSuppData element (section 2.1.2.1) specifies that both the Open
event and the XMLAfterInsert event are active for this document. The mcds element of the
vbaSuppData element contains an mcd element that specifies macro information for the
"Project.NewMacros.Macro1" macro.
3.2 Macro Sheet Part of a Workbook
This example shows a usage of an Excel Macro Sheet part. See section 2.2.1.3 for more information.
Consider the following XML, showing the complete contents of a Macro Sheet part of a workbook.
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
The macrosheet element (section 2.2.2.1) specifies the macro sheet. A formula following the
grammar specified in section 2.2.6 is specified by the f element ([ECMA-376] Part 4, Section 3.17).
36 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

4 Security Considerations
None.
37 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

5 Appendix A: Full XML Schemas
For ease of implementation, this section provides the full W3C XML Schemas for the new elements,
attributes, complex types, and simple types specified in the preceding sections. Any schema
references to namespaces included in ISO/IEC-29500:2008 refer specifically to the transitional
schemas as described in [ISO/IEC-29500-4].
5.1 http://schemas.microsoft.com/office/word/2006/wordml
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
<xsd:element name="eventDocContentControlContentUpdate" minOccurs="0" maxOccs="1"/>
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
5.2 http://schemas.microsoft.com/office/excel/2006/main
<xsd:schema elementFormDefault="qualified" xmlns:xsd="http://www.w3.org/2001/XMLSchema"
xmlns="http://schemas.microsoft.com/office/excel/2006/main"
xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
38 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

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
39 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

6 Appendix B: Product Behavior
The information in this specification is applicable to the following Microsoft products or supplemental
software. References to product versions include released service packs.
 The 2007 Microsoft Office system
 Microsoft Visual Basic for Applications 6.0 (VBA 6.0)
Exceptions, if any, are noted below. If a service pack or Quick Fix Engineering (QFE) number appears
with the product version, behavior changed in that service pack or QFE. The new behavior also applies
to subsequent service packs of the product unless otherwise specified. If a product edition appears
with the product version, behavior is different in that product edition.
Unless otherwise specified, any statement of optional behavior in this specification that is prescribed
using the terms SHOULD or SHOULD NOT implies product behavior in accordance with the SHOULD or
SHOULD NOT prescription. Unless otherwise specified, the term MAY implies that the product does not
follow the prescription.
40 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

7 Change Tracking
No table of changes is available. The document is either new or has had no changes since its last
release.
41 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

8 Index
A
E
Applicability 9
Examples
C Macro Sheet Part of a Workbook 36
VBA Supplemental Data Part of a Word-Processing
Change tracking 41 Document 36
Complex types
CT_DocEvents 12 F
CT_Mcd 11
CT_Mcds 12 Fields - vendor-extensible 9
CT_SlideMasterIdListEntry 35 Formulas – macro-enabled SpreadsheetML structures
CT_VbaSuppData 13 19
CT_Worksheet 16 Full XML schema 38
CT_DocEvents – complex type 12
CT_Mcd – complex type 11 http://schemas.microsoft.com/office/excel/2006
CT_Mcds – complex type 12 /main 38
CT_SlideMasterIdListEntry – complex type 35
CT_VbaSuppData – complex type 13 http://schemas.microsoft.com/office/word/2006
CT_Worksheet – complex type 16 /wordml 38
D G
Details Global attributes – macro-enabled PresentationML
CT_DocEvents – complex type 12 structures 34
CT_Mcd – complex type 11 Global attributes – macro-enabled SpreadsheetML
CT_Mcds – complex type 12 structures 16
CT_SlideMasterIdListEntry – complex type 35 Global attributes – macro-enabled WordprocessingML
CT_VbaSuppData – complex type 13 structures 11
CT_Worksheet – complex type 16 Global elements
formulas - macro-enabled SpreadsheetML macrosheet 16
structures 19 vbaSuppData 11
global attributes - macro-enabled PresentationML Global elements – macro-enabled PresentationML
structures 34 structures 34
global attributes - macro-enabled SpreadsheetML Glossary 6
structures 16
global attributes - macro-enabled I
WordprocessingML structures 11
global elements - macro-enabled PresentationML Implementer - security considerations 37
structures 34 Informative references 8
International Macro Sheet - part enumeration 15 International Macro Sheet - part enumeration 15
Macro Sheet - part enumeration 15 Introduction 6
macrosheet – global element 16
main document - part enumeration 10 L
part enumeration - macro-enabled PresentationML
structures 34 Localization 9
part enumeration - macro-enabled SpreadsheetML
structures 14 M
part enumeration - macro-enabled
WordprocessingML structures 10
Macro Sheet - part enumeration 15
Presentation - part enumeration 34
Macro Sheet Part of a Workbook example 36
simple types - macro-enabled PresentationML
macrosheet – global element 16
structures 35
Main document - part enumeration 10
simple types - macro-enabled SpreadsheetML
structures 19
N
simple types - macro-enabled WordprocessingML
structures 14
Normative references 7
VBA Project - part enumeration (section 2.1.1.2
10, section 2.2.1.2 14, section 2.3.1.2 34)
O
VBA Supplemental Data - part enumeration 11
vbaSuppData – global element 11
Overview (synopsis) 8
Workbook - part enumeration 14
42 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016

Overview (synopsis) – organization of this document part enumeration 10
8 simple types 14
VBA Project - part enumeration 10
P VBA Supplemental Data - part enumeration 11
vbaSuppData – global element 11
Part enumeration
International Macro Sheet 15 T
Macro Sheet 15
main document 10 Tracking changes 41
Presentation 34
VBA Project (section 2.1.1.2 10, section 2.2.1.2 V
14, section 2.3.1.2 34)
VBA Supplemental Data 11 VBA Project - part enumeration (section 2.1.1.2 10,
Workbook 14 section 2.2.1.2 14, section 2.3.1.2 34)
Part enumeration – macro-enabled PresentationML VBA Supplemental Data - part enumeration 11
structures 34 VBA Supplemental Data Part of a Word-Processing
Part enumeration – macro-enabled SpreadsheetML Document example 36
structures 14 vbaSuppData – global element 11
Part enumeration – macro-enabled Vendor-extensible fields 9
WordprocessingML structures 10 Versioning 9
Presentation - part enumeration 34
Product behavior 40 W
R Workbook - part enumeration 14
References 7 X
informative 8
normative 7 XML schema 38
Relationship to protocols and other structures 8
S
Security - implementer considerations 37
Simple types – macro-enabled PresentationML
structures 35
Simple types – macro-enabled SpreadsheetML
structures 19
Simple types – macro-enabled WordprocessingML
structures 14
Structures – macro-enabled PresentationML
CT_SlideMasterIdListEntry – complex type 35
global attributes 34
global elements 34
part enumeration 34
Presentation - part enumeration 34
simple types 35
VBA Project - part enumeration 34
Structures – macro-enabled SpreadsheetML
CT_Worksheet – complex type 16
formulas 19
global attributes 16
International Macro Sheet - part enumeration 15
Macro Sheet - part enumeration 15
macrosheet – global element 16
part enumeration 14
simple types 19
VBA Project - part enumeration 14
Workbook - part enumeration 14
Structures – macro-enabled WordprocessingML
CT_DocEvents – complex type 12
CT_Mcd – complex type 11
CT_Mcds – complex type 12
CT_VbaSuppData – complex type 13
global attributes 11
main document - part enumeration 10
43 / 43
[MS-OFFMACRO] - v20160914
Office Macro-Enabled File Format
Copyright © 2016 Microsoft Corporation
Release: September 14, 2016