﻿
Option Explicit On
Option Strict On

Namespace Contensive.Addons.aoCollectionExporter
    Public Module constants
        Public Const FormIDSelectCollection As Integer = 0
        Public Const FormIDDisplayResults As Integer = 1
        '
        Public Const RequestNameButton As String = "button"
        Public Const RequestNameFormID As String = "formid"
        Public Const RequestnameExecutableFile As String = "executablefile"
        Public Const RequestNameCollectionID As String = "collectionid"
        '
        '
        '-----------------------------------------------------------------------
        ' ----- Field type Definitions
        '       Field Types are numeric values that describe how to treat values
        '       stored as ContentFieldDefinitionType (FieldType property of FieldType Type.. ;)
        '-----------------------------------------------------------------------
        '
        Public Const FieldTypeInteger As Integer = 1       ' An long number
        Public Const FieldTypeText As Integer = 2          ' A text field (up to 255 characters)
        Public Const FieldTypeLongText As Integer = 3      ' A memo field (up to 8000 characters)
        Public Const FieldTypeBoolean As Integer = 4       ' A yes/no field
        Public Const FieldTypeDate As Integer = 5          ' A date field
        Public Const FieldTypeFile As Integer = 6          ' A filename of a file in the files directory.
        Public Const FieldTypeLookup As Integer = 7        ' A lookup is a FieldTypeInteger that indexes into another table
        Public Const FieldTypeRedirect As Integer = 8      ' creates a link to another section
        Public Const FieldTypeCurrency As Integer = 9      ' A Float that prints in dollars
        Public Const FieldTypeTextFile As Integer = 10     ' Text saved in a file in the files area.
        Public Const FieldTypeImage As Integer = 11        ' A filename of a file in the files directory.
        Public Const FieldTypeFloat As Integer = 12        ' A float number
        Public Const FieldTypeAutoIncrement As Integer = 13 'long that automatically increments with the new record
        Public Const FieldTypeManyToMany As Integer = 14    ' no database field - sets up a relationship through a Rule table to another table
        Public Const FieldTypeMemberSelect As Integer = 15 ' This ID is a ccMembers record in a group defined by the MemberSelectGroupID field
        Public Const FieldTypeCSSFile As Integer = 16      ' A filename of a CSS compatible file
        Public Const FieldTypeXMLFile As Integer = 17      ' the filename of an XML compatible file
        Public Const FieldTypeJavascriptFile As Integer = 18 ' the filename of a javascript compatible file
        Public Const FieldTypeLink As Integer = 19           ' Links used in href tags -- can go to pages or resources
        Public Const FieldTypeResourceLink As Integer = 20   ' Links used in resources, link <img or <object. Should not be pages
        Public Const FieldTypeHTML As Integer = 21           ' LongText field that expects HTML content
        Public Const FieldTypeHTMLFile As Integer = 22       ' TextFile field that expects HTML content
        Public Const FieldTypeMax As Integer = 22
        '
        '
        Public Const FieldDescriptorInteger = "Integer"
        Public Const FieldDescriptorText = "Text"
        Public Const FieldDescriptorLongText = "LongText"
        Public Const FieldDescriptorBoolean = "Boolean"
        Public Const FieldDescriptorDate = "Date"
        Public Const FieldDescriptorFile = "File"
        Public Const FieldDescriptorLookup = "Lookup"
        Public Const FieldDescriptorRedirect = "Redirect"
        Public Const FieldDescriptorCurrency = "Currency"
        Public Const FieldDescriptorImage = "Image"
        Public Const FieldDescriptorFloat = "Float"
        Public Const FieldDescriptorManyToMany = "ManyToMany"
        Public Const FieldDescriptorTextFile = "TextFile"
        Public Const FieldDescriptorCSSFile = "CSSFile"
        Public Const FieldDescriptorXMLFile = "XMLFile"
        Public Const FieldDescriptorJavascriptFile = "JavascriptFile"
        Public Const FieldDescriptorLink = "Link"
        Public Const FieldDescriptorResourceLink = "ResourceLink"
        Public Const FieldDescriptorMemberSelect = "MemberSelect"
        Public Const FieldDescriptorHTML = "HTML"
        Public Const FieldDescriptorHTMLFile = "HTMLFile"
        '
        Public Const FieldDescriptorLcaseInteger = "integer"
        Public Const FieldDescriptorLcaseText = "text"
        Public Const FieldDescriptorLcaseLongText = "longtext"
        Public Const FieldDescriptorLcaseBoolean = "boolean"
        Public Const FieldDescriptorLcaseDate = "date"
        Public Const FieldDescriptorLcaseFile = "file"
        Public Const FieldDescriptorLcaseLookup = "lookup"
        Public Const FieldDescriptorLcaseRedirect = "redirect"
        Public Const FieldDescriptorLcaseCurrency = "currency"
        Public Const FieldDescriptorLcaseImage = "image"
        Public Const FieldDescriptorLcaseFloat = "float"
        Public Const FieldDescriptorLcaseManyToMany = "manytomany"
        Public Const FieldDescriptorLcaseTextFile = "textfile"
        Public Const FieldDescriptorLcaseCSSFile = "cssfile"
        Public Const FieldDescriptorLcaseXMLFile = "xmlfile"
        Public Const FieldDescriptorLcaseJavascriptFile = "javascriptfile"
        Public Const FieldDescriptorLcaseLink = "link"
        Public Const FieldDescriptorLcaseResourceLink = "resourcelink"
        Public Const FieldDescriptorLcaseMemberSelect = "memberselect"
        Public Const FieldDescriptorLcaseHTML = "html"
        Public Const FieldDescriptorLcaseHTMLFile = "htmlfile"

        '
        ' ----- This should match the Lookup List in the NavIconType field in the Navigator Entry content definition
        '
        Public Const navTypeIDList = "Add-on,Report,Setting,Tool"
        Public Const NavTypeIDAddon = 1
        Public Const NavTypeIDReport = 2
        Public Const NavTypeIDSetting = 3
        Public Const NavTypeIDTool = 4
        '
        Public Const NavIconTypeList = "Custom,Advanced,Content,Folder,Email,User,Report,Setting,Tool,Record,Addon,help"
        Public Const NavIconTypeCustom = 1
        Public Const NavIconTypeAdvanced = 2
        Public Const NavIconTypeContent = 3
        Public Const NavIconTypeFolder = 4
        Public Const NavIconTypeEmail = 5
        Public Const NavIconTypeUser = 6
        Public Const NavIconTypeReport = 7
        Public Const NavIconTypeSetting = 8
        Public Const NavIconTypeTool = 9
        Public Const NavIconTypeRecord = 10
        Public Const NavIconTypeAddon = 11
        Public Const NavIconTypeHelp = 12
    End Module
End Namespace

