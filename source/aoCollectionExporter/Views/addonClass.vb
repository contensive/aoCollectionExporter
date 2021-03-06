
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses
Imports Contensive.Addons.aoCollectionExporter
Imports ICSharpCode.SharpZipLib
Imports System.IO

Namespace Contensive.Addons
    Public Class aoCollectionExporterClass
        Inherits AddonBaseClass
        '
        '====================================================================================================
        ' instance store
        '
        Private cp As CPBaseClass
        '
        '====================================================================================================
        ''' <summary>
        ''' addon api
        ''' </summary>
        ''' <param name="CP"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String = ""
            Try
                Me.cp = CP
                Dim siteContext As New siteContextClass(CP)
                Dim Button As String
                Dim FormID As Integer
                Dim CollectionID As Integer
                Dim CollectionName As String
                Dim CollectionFilename As String = ""
                '
                ' Every form returns a button and a formid
                '
                Button = CP.Doc.GetText("button")
                FormID = CP.Doc.GetInteger("formid")
                '
                ' Process the current form submission
                '
                If Button <> "" Then
                    Select Case FormID
                        Case FormIDDisplayResults
                            '
                            ' nothing to process
                            '
                        Case Else
                            '
                            ' process the Select Collection Form button
                            CollectionID = CP.Doc.GetInteger(RequestNameCollectionID)
                            CollectionName = CP.Content.GetRecordName("Add-on Collections", CollectionID)
                            If CollectionName = "" Then
                                Call CP.UserError.Add("The collection file you selected could not be found. Please select another.")
                            Else
                                CollectionFilename = GetCollectionZipPathFilename(CollectionID)
                            End If
                            If Not CP.UserError.OK() Then
                                FormID = FormIDSelectCollection
                            Else

                                FormID = FormIDDisplayResults
                            End If
                    End Select
                End If
                'hint = hint & ",400"
                'Call Main.testpoint("hint=" & hint)
                '
                ' Reply with the next form
                '
                Select Case FormID
                    Case FormIDDisplayResults
                        '
                        ' Diplay the results page
                        '
                        'hint = hint & ",500"
                        'Call Main.testpoint("hint=" & hint)
                        returnHtml = CP.UserError.GetList() _
                                & vbCrLf & vbTab & "<div class=""responseForm"">" _
                                & vbCrLf & vbTab & vbTab & "<p>Click <a href=""" & CP.Site.FilePath & Replace(CollectionFilename, "\", "/") & """>here</a> to download the collection file</p>" _
                                & vbCrLf & vbTab & "</div>"
                    Case Else
                        '
                        ' ask them to select a collectioin to export
                        '
                        returnHtml = "" _
                                & vbCrLf & vbTab & "<div class=""mainForm"">" _
                                & vbCrLf & vbTab & vbTab & CP.UserError.GetList() _
                                & vbCrLf & vbTab & vbTab & vbTab & "<p>Select a collection to be exported. If the project is being developed and you need to add an executable resource that is not installed as an add-on on this site, use the file upload.</p>" _
                                & vbCrLf & vbTab & vbTab & vbTab & "<p>" & CP.Html.SelectContent(RequestNameCollectionID, "0", "Add-on Collections") & "<br>The collection to export</p>" _
                                & vbCrLf & vbTab & vbTab & vbTab & "<p>" & CP.Html.Button("button", "Export Collection") & "</p>" _
                                & vbCrLf & vbTab & vbTab & "</form>" _
                                & vbCrLf & vbTab & "</div>"
                End Select
                '
                returnHtml = "" _
                        & vbCrLf & vbTab & "<div class=""collectionExport"">" _
                        & CP.Html.Form(returnHtml) _
                        & vbCrLf & vbTab & "</div>"
            Catch ex As Exception
                errorReport(CP, ex, "execute")
                returnHtml = ""
            End Try
            Return returnHtml
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' error handler
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="ex"></param>
        ''' <param name="method"></param>
        ''' <remarks></remarks>
        Private Sub errorReport(ByVal cp As CPBaseClass, ByVal ex As Exception, ByVal method As String)
            Try
                cp.Site.ErrorReport(ex, "Unexpected error in aoEollectionExportClass." & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                '
            End Try
        End Sub
        '
        '====================================================================================================
        ''' <summary>
        ''' Create the collection file and return a filepath to the zip
        ''' </summary>
        ''' <param name="CollectionID"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetCollectionZipPathFilename(CollectionID As Integer) As String
            Dim collectionZipPathFilename As String = ""
            Try
                Dim IncludeSharedStyleGuidList As String
                Dim isUpdatable As Boolean
                Dim fieldLookupListValue As String = ""
                Dim FieldValueInteger As Integer
                Dim FieldLookupContentName As String
                Dim fieldPtr As Integer
                Dim fieldCnt As Integer
                Dim fieldNames() As String = {}
                Dim fieldTypes() As Integer = {}
                Dim fieldLookupContent() As String = {}
                Dim fieldLookupList() As String = {}
                Dim FieldLookupContentID As Integer
                Dim Criteria As String
                Dim supportsGuid As Boolean
                Dim reload As Boolean
                Dim ContentID As Integer
                Dim FieldValue As String
                Dim FieldTypeNumber As Integer
                Dim DataRecordList As String
                Dim DataRecords() As String
                Dim DataRecord As String
                Dim DataSplit() As String
                Dim DataContentName As String
                Dim DataContentId As Integer
                Dim DataRecordGuid As String
                Dim DataRecordName As String
                Dim TestString As String
                Dim FieldName As String
                Dim FieldNodes As String
                Dim RecordNode As String
                Dim Modules() As String
                Dim ModuleGuid As String
                Dim Code As String
                Dim ManualFilename As String = ""
                Dim ResourceCnt As Integer
                Dim ContentName As String
                Dim FileList As String
                Dim Files() As String
                Dim Ptr As Integer
                Dim PathFilename As String
                Dim Filename As String
                Dim Path As String
                Dim Pos As Integer
                Dim collectionXml As String
                Dim Node As String
                Dim CollectionGuid As String
                Dim Guid As String
                Dim ArchiveFilename As String
                Dim ArchivePath As String
                Dim InstallFilename As String
                Dim CollectionName As String
                Dim AddFilename As String
                Dim PhysicalWWWPath As String
                Dim CollectionPath As String = ""
                Dim LastChangeDate As Date
                Dim AddonPath As String
                Dim AddFileList As New List(Of String)
                Dim IncludeModuleGuidList As String = ""
                Dim ExecFileListNode As String = ""
                Dim blockNavigatorNode As Boolean
                Dim CSData As CPCSBaseClass = cp.CSNew()
                Dim CS As CPCSBaseClass = cp.CSNew()
                Dim CS2 As CPCSBaseClass = cp.CSNew()
                Dim CS3 As CPCSBaseClass = cp.CSNew()
                Dim CSlookup As CPCSBaseClass = cp.CSNew()
                '
                IncludeSharedStyleGuidList = ""
                '
                CS.OpenRecord("Add-on Collections", CollectionID)
                If Not CS.OK() Then
                    Call cp.UserError.Add("The collection you selected could not be found")
                Else
                    CollectionGuid = CS.GetText("ccGuid")
                    If CollectionGuid = "" Then
                        CollectionGuid = cp.Utils.CreateGuid()
                        Call CS.SetField("ccGuid", CollectionGuid)
                    End If
                    CollectionName = CS.GetText("name")
                    isUpdatable = CS.GetBoolean("updatable")
                    blockNavigatorNode = CS.GetBoolean("blockNavigatorNode")
                    '
                    ' Archive Filenames
                    '
                    'Call Main.testpoint("getCollection, 200")
                    ArchivePath = cp.Site.PhysicalFilePath & "CollectionExport\"
                    'Call Main.testpoint("getCollection, 201")
                    InstallFilename = encodeFilename(CollectionName & ".xml")
                    'Call Main.testpoint("getCollection, 202")
                    InstallFilename = ArchivePath & InstallFilename
                    'Call Main.testpoint("getCollection, 203")
                    ArchiveFilename = encodeFilename(CollectionName & ".zip")
                    'Call Main.testpoint("getCollection, 204")
                    ArchiveFilename = ArchivePath & ArchiveFilename
                    'Call Main.testpoint("getCollection, 205")
                    collectionZipPathFilename = "CollectionExport\" & encodeFilename(CollectionName & ".zip")
                    'Call Main.testpoint("getCollection, 207")
                    '
                    ' Delete old archive file
                    '
                    Call cp.File.Delete(ArchiveFilename)
                    '
                    collectionXml = "" _
                            & "<?xml version=""1.0"" encoding=""windows-1252""?>" _
                            & vbCrLf & "<Collection name=""" & cp.Utils.EncodeHTML(CollectionName) & """ guid=""" & CollectionGuid & """ system=""" & kmaGetYesNo(CS.GetBoolean("system")) & """ updatable=""" & kmaGetYesNo(isUpdatable) & """ blockNavigatorNode=""" & kmaGetYesNo(blockNavigatorNode) & """>"
                    If File.Exists(InstallFilename) Then
                        File.Delete(InstallFilename)
                    End If
                    Using sw As StreamWriter = File.CreateText(InstallFilename)
                        sw.WriteLine(collectionXml)
                        collectionXml = ""
                    End Using
                    '
                    ' Build executable file list Resource Node so executables can be added to addons for Version40compatibility
                    '   but save it for the end, executableFileList
                    '
                    'Call Main.testpoint("getCollection, 400")
                    AddonPath = cp.Site.PhysicalInstallPath & "\addons\"
                    FileList = CS.GetText("execFileList")
                    If FileList <> "" Then
                        '
                        ' There are executable files to include in the collection
                        '   If installed, source path is collectionpath, if not installed, collectionpath will be empty
                        '   and file will be sourced right from addon path
                        '
                        Call GetLocalCollectionArgs(CollectionGuid, CollectionPath, LastChangeDate)
                        If CollectionPath <> "" Then
                            CollectionPath = CollectionPath & "\"
                        End If
                        Files = Split(FileList, vbCrLf)
                        For Ptr = 0 To UBound(Files)
                            PathFilename = Files(Ptr)
                            If PathFilename <> "" Then
                                PathFilename = Replace(PathFilename, "\", "/")
                                Path = ""
                                Filename = PathFilename
                                Pos = InStrRev(PathFilename, "/")
                                If Pos > 0 Then
                                    Filename = Mid(PathFilename, Pos + 1)
                                    Path = Mid(PathFilename, 1, Pos - 1)
                                End If
                                If LCase(Filename) <> LCase(ManualFilename) Then
                                    AddFilename = AddonPath & CollectionPath & Filename
                                    If Not AddFileList.Contains(AddFilename) Then
                                        AddFileList.Add(AddFilename)
                                        ExecFileListNode = ExecFileListNode & vbCrLf & vbTab & "<Resource name=""" & cp.Utils.EncodeHTML(Filename) & """ type=""executable"" path=""" & cp.Utils.EncodeHTML(Path) & """ />"
                                    End If
                                End If
                                ResourceCnt = ResourceCnt + 1
                            End If
                        Next
                    End If
                    'Call Main.testpoint("getCollection, 500")
                    If (ResourceCnt = 0) And (CollectionPath <> "") Then
                        '
                        ' If no resources were in the collection record, this might be an old installation
                        ' Add all .dll files in the CollectionPath
                        '
                        ExecFileListNode = ExecFileListNode & AddCompatibilityResources(AddonPath & CollectionPath, ArchiveFilename, "")
                    End If
                    '
                    ' helpLink
                    '
                    Dim HelpLink As String = CS.GetText("HelpLink")
                    collectionXml = collectionXml & vbCrLf & vbTab & "<HelpLink>" & cp.Utils.EncodeHTML(HelpLink) & "</HelpLink>"
                    collectionXml = collectionXml & vbCrLf & vbTab & "<Help>" & cp.Utils.EncodeHTML(CS.GetText("Help")) & "</Help>"
                    Using sw As StreamWriter = File.AppendText(InstallFilename)
                        sw.WriteLine(collectionXml)
                        collectionXml = ""
                    End Using
                    '
                    ' Addons
                    '
                    CS2.Open("Add-ons", "collectionid=" & CollectionID, , , "ccguid")
                    Do While CS2.OK()
                        collectionXml = collectionXml & GetAddonNode(CS2.GetInteger("id"), IncludeModuleGuidList, IncludeSharedStyleGuidList)
                        Call CS2.GoNext()
                    Loop
                    Using sw As StreamWriter = File.AppendText(InstallFilename)
                        sw.WriteLine(collectionXml)
                        collectionXml = ""
                    End Using
                    '
                    ' Data Records
                    '
                    'Call Main.testpoint("getCollection, 600")
                    DataRecordList = CS.GetText("DataRecordList")
                    If DataRecordList <> "" Then
                        DataRecords = Split(DataRecordList, vbCrLf)
                        RecordNode = ""
                        collectionXml = "<data>"
                        Using sw As StreamWriter = File.AppendText(InstallFilename)
                            sw.WriteLine(collectionXml)
                            collectionXml = ""
                        End Using
                        For Ptr = 0 To UBound(DataRecords)
                            FieldNodes = ""
                            DataRecordName = ""
                            DataRecordGuid = ""
                            DataRecord = DataRecords(Ptr)
                            If DataRecord <> "" Then
                                DataSplit = Split(DataRecord, ",")
                                If UBound(DataSplit) >= 0 Then
                                    DataContentName = Trim(DataSplit(0))
                                    DataContentId = cp.Content.GetID(DataContentName)
                                    If DataContentId <= 0 Then
                                        RecordNode = "" _
                                                & RecordNode _
                                                & vbCrLf & vbTab & "<!-- data missing, content not found during export, content=""" & DataContentName & """ guid=""" & DataRecordGuid & """ name=""" & DataRecordName & """ -->"
                                    Else
                                        supportsGuid = cp.Content.IsField(DataContentName, "ccguid")
                                        If UBound(DataSplit) = 0 Then
                                            Criteria = ""
                                        Else
                                            TestString = Trim(DataSplit(1))
                                            If TestString = "" Then
                                                '
                                                ' blank is a select all
                                                '
                                                Criteria = ""
                                                DataRecordName = ""
                                                DataRecordGuid = ""
                                            ElseIf Not supportsGuid Then
                                                '
                                                ' if no guid, this is name
                                                '
                                                DataRecordName = TestString
                                                DataRecordGuid = ""
                                                Criteria = "name=" & cp.Db.EncodeSQLText(DataRecordName)
                                            ElseIf (Len(TestString) = 38) And (Left(TestString, 1) = "{") And (Right(TestString, 1) = "}") Then
                                                '
                                                ' guid {726ED098-5A9E-49A9-8840-767A74F41D01} format
                                                '
                                                DataRecordGuid = TestString
                                                DataRecordName = ""
                                                Criteria = "ccguid=" & cp.Db.EncodeSQLText(DataRecordGuid)
                                            ElseIf (Len(TestString) = 36) And (Mid(TestString, 9, 1) = "-") Then
                                                '
                                                ' guid 726ED098-5A9E-49A9-8840-767A74F41D01 format
                                                '
                                                DataRecordGuid = TestString
                                                DataRecordName = ""
                                                Criteria = "ccguid=" & cp.Db.EncodeSQLText(DataRecordGuid)
                                            ElseIf (Len(TestString) = 32) And (InStr(1, TestString, " ") = 0) Then
                                                '
                                                ' guid 726ED0985A9E49A98840767A74F41D01 format
                                                '
                                                DataRecordGuid = TestString
                                                DataRecordName = ""
                                                Criteria = "ccguid=" & cp.Db.EncodeSQLText(DataRecordGuid)
                                            Else
                                                '
                                                ' use name
                                                '
                                                DataRecordName = TestString
                                                DataRecordGuid = ""
                                                Criteria = "name=" & cp.Db.EncodeSQLText(DataRecordName)
                                            End If
                                        End If
                                        If Not CSData.Open(DataContentName, Criteria, "id") Then
                                            RecordNode = "" _
                                                    & RecordNode _
                                                    & vbCrLf & vbTab & "<!-- data missing, record not found during export, content=""" & DataContentName & """ guid=""" & DataRecordGuid & """ name=""" & DataRecordName & """ -->"
                                        Else
                                            '
                                            ' determine all valid fields
                                            '
                                            fieldCnt = 0
                                            Dim Sql As String = "select * from ccFields where contentid=" & DataContentId
                                            Dim csFields As CPCSBaseClass = cp.CSNew()
                                            If csFields.Open("content fields", "contentid=" & DataContentId) Then
                                                Do
                                                    FieldName = csFields.GetText("name")
                                                    If FieldName <> "" Then
                                                        FieldLookupContentID = 0
                                                        FieldLookupContentName = ""
                                                        FieldTypeNumber = csFields.GetInteger("type")
                                                        Select Case LCase(FieldName)
                                                            Case "ccguid", "name", "id", "dateadded", "createdby", "modifiedby", "modifieddate", "createkey", "contentcontrolid", "editsourceid", "editarchive", "editblank", "contentcategoryid"
                                                            Case Else
                                                                If FieldTypeNumber = 7 Then
                                                                    FieldLookupContentID = csFields.GetInteger("Lookupcontentid")
                                                                    fieldLookupListValue = csFields.GetText("LookupList")
                                                                    If FieldLookupContentID <> 0 Then
                                                                        FieldLookupContentName = cp.Content.GetRecordName("content", FieldLookupContentID)
                                                                    End If
                                                                End If
                                                                Select Case FieldTypeNumber
                                                                    Case FieldTypeLookup, FieldTypeBoolean, FieldTypeCSSFile, FieldTypeJavascriptFile, FieldTypeTextFile, FieldTypeXMLFile, FieldTypeCurrency, FieldTypeFloat, FieldTypeInteger, FieldTypeDate, FieldTypeLink, FieldTypeLongText, FieldTypeResourceLink, FieldTypeText, FieldTypeHTML, FieldTypeHTMLFile
                                                                        '
                                                                        ' this is a keeper
                                                                        '
                                                                        ReDim Preserve fieldNames(fieldCnt)
                                                                        ReDim Preserve fieldTypes(fieldCnt)
                                                                        ReDim Preserve fieldLookupContent(fieldCnt)
                                                                        ReDim Preserve fieldLookupList(fieldCnt)
                                                                        'fieldLookupContent
                                                                        fieldNames(fieldCnt) = FieldName
                                                                        fieldTypes(fieldCnt) = FieldTypeNumber
                                                                        fieldLookupContent(fieldCnt) = FieldLookupContentName
                                                                        fieldLookupList(fieldCnt) = fieldLookupListValue
                                                                        fieldCnt = fieldCnt + 1
                                                                        'end case
                                                                End Select
                                                                'end case
                                                        End Select
                                                    End If
                                                    Call csFields.GoNext()
                                                Loop While csFields.OK()
                                            End If
                                            Call csFields.Close()
                                            '
                                            ' output records
                                            '
                                            DataRecordGuid = ""
                                            Do While CSData.OK()
                                                FieldNodes = ""
                                                DataRecordName = CSData.GetText("name")
                                                If supportsGuid Then
                                                    DataRecordGuid = CSData.GetText("ccguid")
                                                    If DataRecordGuid = "" Then
                                                        DataRecordGuid = cp.Utils.CreateGuid()
                                                        Call CSData.SetField("ccGuid", DataRecordGuid)
                                                    End If
                                                End If
                                                For fieldPtr = 0 To fieldCnt - 1
                                                    FieldName = fieldNames(fieldPtr)
                                                    FieldTypeNumber = cp.Utils.EncodeInteger(fieldTypes(fieldPtr))
                                                    Select Case FieldTypeNumber
                                                        Case FieldTypeBoolean
                                                            '
                                                            ' true/false
                                                            '
                                                            FieldValue = CSData.GetBoolean(FieldName).ToString()
                                                        Case FieldTypeCSSFile, FieldTypeJavascriptFile, FieldTypeTextFile, FieldTypeXMLFile
                                                            '
                                                            ' text files
                                                            '
                                                            FieldValue = CSData.GetText(FieldName)
                                                            FieldValue = EncodeCData(FieldValue)
                                                        Case FieldTypeInteger
                                                            '
                                                            ' integer
                                                            '
                                                            FieldValue = CSData.GetInteger(FieldName).ToString()
                                                        Case FieldTypeCurrency, FieldTypeFloat
                                                            '
                                                            ' numbers
                                                            '
                                                            FieldValue = CSData.GetNumber(FieldName).ToString()
                                                        Case FieldTypeDate
                                                            '
                                                            ' date
                                                            '
                                                            FieldValue = CSData.GetDate(FieldName).ToString()
                                                        Case FieldTypeLookup
                                                            '
                                                            ' lookup
                                                            '
                                                            FieldValue = ""
                                                            FieldValueInteger = CSData.GetInteger(FieldName)
                                                            If (FieldValueInteger <> 0) Then
                                                                FieldLookupContentName = fieldLookupContent(fieldPtr)
                                                                fieldLookupListValue = fieldLookupList(fieldPtr)
                                                                If (FieldLookupContentName <> "") Then
                                                                    '
                                                                    ' content lookup
                                                                    '
                                                                    If cp.Content.IsField(FieldLookupContentName, "ccguid") Then
                                                                        Call CSlookup.OpenRecord(FieldLookupContentName, FieldValueInteger)
                                                                        If CSlookup.OK() Then
                                                                            FieldValue = CSlookup.GetText("ccguid")
                                                                            If FieldValue = "" Then
                                                                                FieldValue = cp.Utils.CreateGuid()
                                                                                Call CSlookup.SetField("ccGuid", FieldValue)
                                                                            End If
                                                                        End If
                                                                        Call CSlookup.Close()
                                                                    End If
                                                                ElseIf fieldLookupListValue <> "" Then
                                                                    '
                                                                    ' list lookup, ok to save integer
                                                                    '
                                                                    FieldValue = FieldValueInteger.ToString()
                                                                End If
                                                            End If
                                                        Case Else
                                                            '
                                                            ' text types
                                                            '
                                                            FieldValue = CSData.GetText(FieldName)
                                                            FieldValue = EncodeCData(FieldValue)
                                                    End Select
                                                    FieldNodes = FieldNodes & vbCrLf & vbTab & "<field name=""" & cp.Utils.EncodeHTML(FieldName) & """>" & FieldValue & "</field>"
                                                Next
                                                RecordNode = "" _
                                                        & vbCrLf & vbTab & "<record content=""" & cp.Utils.EncodeHTML(DataContentName) & """ guid=""" & DataRecordGuid & """ name=""" & cp.Utils.EncodeHTML(DataRecordName) & """>" _
                                                        & tabIndent(FieldNodes) _
                                                        & vbCrLf & vbTab & "</record>"
                                                Using sw As StreamWriter = File.AppendText(InstallFilename)
                                                    sw.WriteLine(RecordNode)
                                                    RecordNode = ""
                                                End Using
                                                Call CSData.GoNext()
                                            Loop
                                        End If
                                        Call CSData.Close()
                                    End If
                                End If
                            End If
                        Next
                        collectionXml = "</data>"
                        Using sw As StreamWriter = File.AppendText(InstallFilename)
                            sw.WriteLine(collectionXml)
                            collectionXml = ""
                        End Using
                    End If
                    '
                    ' CDef
                    '
                    'Call Main.testpoint("getCollection, 700")
                    CS2.Open("Add-on Collection CDef Rules", "CollectionID=" & CollectionID)
                    Do While CS2.OK()
                        ContentName = ""
                        ContentID = CS2.GetInteger("contentid")
                        '
                        ' get name and make sure there is a guid
                        '
                        reload = False
                        If CS3.OpenRecord("content", ContentID) Then
                            ContentName = CS3.GetText("name")
                            If CS3.GetText("ccguid") = "" Then
                                Call CS3.SetField("ccGuid", cp.Utils.CreateGuid())
                                reload = True
                            End If
                        End If
                        Call CS3.Close()
                        '
                        If Not String.IsNullOrEmpty(ContentName) Then
                            Dim xmlTool As New xmlToolsClass(cp)
                            Node = xmlTool.GetXMLContentDefinition3(ContentName)
                            '
                            ' remove the <collection> top node
                            '
                            Pos = InStr(1, Node, "<cdef", vbTextCompare)
                            If Pos > 0 Then
                                Node = Mid(Node, Pos)
                                Pos = InStr(1, Node, "</cdef>", vbTextCompare)
                                If Pos > 0 Then
                                    Node = Mid(Node, 1, Pos + 6)
                                    collectionXml = collectionXml & vbCrLf & vbTab & Node
                                    Using sw As StreamWriter = File.AppendText(InstallFilename)
                                        sw.WriteLine(collectionXml)
                                        collectionXml = ""
                                    End Using
                                End If
                            End If
                        End If
                        Call CS2.GoNext()
                    Loop
                    Call CS2.Close()
                    '
                    ' Scripting Modules
                    '
                    'Call Main.testpoint("getCollection, 800")

                    If IncludeModuleGuidList <> "" Then
                        Modules = Split(IncludeModuleGuidList, vbCrLf)
                        For Ptr = 0 To UBound(Modules)
                            ModuleGuid = Modules(Ptr)
                            If ModuleGuid <> "" Then
                                CS2.Open("Scripting Modules", "ccguid=" & cp.Db.EncodeSQLText(ModuleGuid))
                                If CS2.OK() Then
                                    Code = Trim(CS2.GetText("code"))
                                    Code = EncodeCData(Code)
                                    collectionXml = collectionXml & vbCrLf & vbTab & "<ScriptingModule Name=""" & cp.Utils.EncodeHTML(CS2.GetText("name")) & """ guid=""" & ModuleGuid & """>" & Code & "</ScriptingModule>"
                                    Using sw As StreamWriter = File.AppendText(InstallFilename)
                                        sw.WriteLine(collectionXml)
                                        collectionXml = ""
                                    End Using
                                End If
                                Call CS2.Close()
                            End If
                        Next
                    End If
                    '
                    ' shared styles
                    '
                    If (Controllers.genericController.buildVersion(cp) < "5") Then
                        Dim recordGuids() As String
                        Dim recordGuid As String
                        If (IncludeSharedStyleGuidList <> "") Then
                            recordGuids = Split(IncludeSharedStyleGuidList, vbCrLf)
                            For Ptr = 0 To UBound(recordGuids)
                                recordGuid = recordGuids(Ptr)
                                If recordGuid <> "" Then
                                    CS2.Open("Shared Styles", "ccguid=" & cp.Db.EncodeSQLText(recordGuid))
                                    If CS2.OK() Then
                                        collectionXml = collectionXml & vbCrLf & vbTab & "<SharedStyle" _
                                                & " Name=""" & cp.Utils.EncodeHTML(CS2.GetText("name")) & """" _
                                                & " guid=""" & recordGuid & """" _
                                                & " alwaysInclude=""" & CS2.GetBoolean("alwaysInclude") & """" _
                                                & " prefix=""" & cp.Utils.EncodeHTML(CS2.GetText("prefix")) & """" _
                                                & " suffix=""" & cp.Utils.EncodeHTML(CS2.GetText("suffix")) & """" _
                                                & " sortOrder=""" & cp.Utils.EncodeHTML(CS2.GetText("sortOrder")) & """" _
                                                & ">" _
                                                & EncodeCData(Trim(CS2.GetText("styleFilename"))) _
                                                & "</SharedStyle>"
                                        Using sw As StreamWriter = File.AppendText(InstallFilename)
                                            sw.WriteLine(collectionXml)
                                            collectionXml = ""
                                        End Using
                                    End If
                                    Call CS2.Close()
                                End If
                            Next
                        End If
                    End If
                    '
                    ' Import Collections
                    '
                    Node = ""
                    If CS3.Open("Add-on Collection Parent Rules", "parentid=" & CollectionID) Then
                        Do
                            CS2.OpenRecord("Add-on Collections", CS3.GetInteger("childid"))
                            If CS2.OK() Then
                                Guid = CS2.GetText("ccGuid")
                                If Guid = "" Then
                                    Guid = cp.Utils.CreateGuid()
                                    Call CS2.SetField("ccGuid", Guid)
                                End If
                                Node = Node & vbCrLf & vbTab & "<ImportCollection name=""" & cp.Utils.EncodeHTML(CS2.GetText("name")) & """>" & Guid & "</ImportCollection>"
                            End If
                            Call CS2.Close()
                            Call CS3.GoNext()
                        Loop While CS3.OK()
                    End If
                    Call CS3.Close()
                    collectionXml = collectionXml & Node
                    Using sw As StreamWriter = File.AppendText(InstallFilename)
                        sw.WriteLine(collectionXml)
                        collectionXml = ""
                    End Using
                    '
                    ' wwwFileList
                    '
                    ResourceCnt = 0
                    FileList = CS.GetText("wwwFileList")
                    If FileList <> "" Then
                        PhysicalWWWPath = cp.Site.PhysicalWWWPath
                        If Right(PhysicalWWWPath, 1) <> "\" Then
                            PhysicalWWWPath = PhysicalWWWPath & "\"
                        End If
                        Files = Split(FileList, vbCrLf)
                        For Ptr = 0 To UBound(Files)
                            PathFilename = Files(Ptr)
                            If PathFilename <> "" Then
                                PathFilename = Replace(PathFilename, "\", "/")
                                Path = ""
                                Filename = PathFilename
                                Pos = InStrRev(PathFilename, "/")
                                If Pos > 0 Then
                                    Filename = Mid(PathFilename, Pos + 1)
                                    Path = Mid(PathFilename, 1, Pos - 1)
                                End If
                                If LCase(Filename) = "collection.hlp" Then
                                    '
                                    ' legacy file, remove it
                                    '
                                Else
                                    PathFilename = Replace(PathFilename, "/", "\")
                                    AddFilename = PhysicalWWWPath & PathFilename
                                    If AddFileList.Contains(AddFilename) Then
                                        Call cp.UserError.Add("There was an error exporting this collection because there were multiple files with the same filename [" & Filename & "]")
                                    Else
                                        AddFileList.Add(AddFilename)
                                        collectionXml = collectionXml & vbCrLf & vbTab & "<Resource name=""" & cp.Utils.EncodeHTML(Filename) & """ type=""www"" path=""" & cp.Utils.EncodeHTML(Path) & """ />"
                                        Using sw As StreamWriter = File.AppendText(InstallFilename)
                                            sw.WriteLine(collectionXml)
                                            collectionXml = ""
                                        End Using
                                    End If
                                    ResourceCnt = ResourceCnt + 1
                                End If
                            End If
                        Next
                    End If
                    '
                    ' ContentFileList
                    '
                    FileList = CS.GetText("ContentFileList")
                    If FileList <> "" Then
                        Files = Split(FileList, vbCrLf)
                        For Ptr = 0 To UBound(Files)
                            PathFilename = Files(Ptr)
                            If PathFilename <> "" Then
                                PathFilename = Replace(PathFilename, "\", "/")
                                Path = ""
                                Filename = PathFilename
                                Pos = InStrRev(PathFilename, "/")
                                If Pos > 0 Then
                                    Filename = Mid(PathFilename, Pos + 1)
                                    Path = Mid(PathFilename, 1, Pos - 1)
                                End If
                                PathFilename = Replace(PathFilename, "/", "\")
                                If Left(PathFilename, 1) = "\" Then
                                    PathFilename = Mid(PathFilename, 2)
                                End If
                                AddFilename = cp.Site.PhysicalFilePath & PathFilename
                                If AddFileList.Contains(AddFilename) Then
                                    Call cp.UserError.Add("There was an error exporting this collection because there were multiple files with the same filename [" & Filename & "]")
                                Else
                                    AddFileList.Add(AddFilename)
                                    collectionXml = collectionXml & vbCrLf & vbTab & "<Resource name=""" & cp.Utils.EncodeHTML(Filename) & """ type=""content"" path=""" & cp.Utils.EncodeHTML(Path) & """ />"
                                    Using sw As StreamWriter = File.AppendText(InstallFilename)
                                        sw.WriteLine(collectionXml)
                                        collectionXml = ""
                                    End Using
                                End If
                                ResourceCnt = ResourceCnt + 1
                            End If
                        Next
                    End If
                    '
                    ' ExecFileListNode
                    '
                    collectionXml = collectionXml & ExecFileListNode
                    Using sw As StreamWriter = File.AppendText(InstallFilename)
                        sw.WriteLine(collectionXml)
                        collectionXml = ""
                    End Using
                    '
                    ' Other XML
                    '
                    Dim OtherXML As String
                    OtherXML = CS.GetText("otherxml")
                    If Trim(OtherXML) <> "" Then
                        collectionXml = collectionXml & vbCrLf & OtherXML
                        Using sw As StreamWriter = File.AppendText(InstallFilename)
                            sw.WriteLine(collectionXml)
                            collectionXml = ""
                        End Using
                    End If
                    '
                    ' -- v5 templates

                    '
                    ' -- done, close collection
                    collectionXml = collectionXml & vbCrLf & "</Collection>"
                    Using sw As StreamWriter = File.AppendText(InstallFilename)
                        sw.WriteLine(collectionXml)
                        collectionXml = ""
                    End Using
                    Call CS.Close()
                    '
                    ' Save the installation file and add it to the archive
                    '
                    'Call cp.File.Save(InstallFilename, collectionXml)
                    If Not AddFileList.Contains(InstallFilename) Then
                        AddFileList.Add(InstallFilename)
                    End If
                    Call zipFile(ArchiveFilename, AddFileList)
                    'Call runAtServer("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable("@" & AddFileListFilename))
                End If
            Catch ex As Exception
                errorReport(cp, ex, "GetCollection")
            End Try
            Return collectionZipPathFilename
        End Function
        '
        '====================================================================================================

        Private Function GetAddonNode(addonid As Integer, ByRef Return_IncludeModuleGuidList As String, ByRef Return_IncludeSharedStyleGuidList As String) As String
            GetAddonNode = ""
            Dim s As String = ""
            Try
                '
                Dim styleId As Integer
                Dim fieldType As String
                Dim fieldTypeID As Integer
                Dim TriggerContentID As Integer
                Dim StylesTest As String
                Dim NavType As String
                Dim Styles As String
                Dim NodeInnerText As String
                Dim IncludedAddonID As Integer
                Dim ScriptingModuleID As Integer
                Dim Guid As String
                Dim addonName As String
                Dim CS As CPCSBaseClass = cp.CSNew()
                Dim CS2 As CPCSBaseClass = cp.CSNew()
                Dim CS3 As CPCSBaseClass = cp.CSNew()
                '
                If CS.OpenRecord("Add-ons", addonid) Then
                    addonName = CS.GetText("name")
                    '
                    ' ActiveX DLL node is being deprecated. This should be in the collection resource section
                    '
                    s = s & GetNodeText("Copy", CS, "Copy")
                    s = s & GetNodeText("CopyText", CS, "CopyText")
                    '
                    ' DLL
                    '

                    s = s & GetNodeText("ActiveXProgramID", CS, "objectprogramid")
                    s = s & GetNodeText("DotNetClass", CS, "DotNetClass")
                    '
                    ' Features
                    '
                    s = s & GetNodeText("ArgumentList", CS, "ArgumentList")
                    s = s & GetNodeBoolean("AsAjax", CS, "AsAjax")
                    s = s & GetNodeBoolean("Filter", CS, "Filter")
                    s = s & GetNodeText("Help", CS, "Help")
                    s = s & GetNodeText("HelpLink", CS, "HelpLink")
                    s = s & vbCrLf & vbTab & "<Icon Link=""" & CS.GetText("iconfilename") & """ width=""" & CS.GetInteger("iconWidth") & """ height=""" & CS.GetInteger("iconHeight") & """ sprites=""" & CS.GetInteger("iconSprites") & """ />"
                    s = s & GetNodeBoolean("InIframe", CS, "InFrame")
                    s = s & GetNodeBoolean("BlockEditTools", CS, "BlockEditTools")
                    '
                    ' Form XML
                    '
                    s = s & GetNodeText("FormXML", CS, "FormXML")
                    '
                    NodeInnerText = ""
                    CS2.Open("Add-on Include Rules", "addonid=" & addonid)
                    Do While CS2.OK()
                        IncludedAddonID = CS2.GetInteger("IncludedAddonID")
                        CS3.Open("Add-ons", "ID=" & IncludedAddonID)
                        If CS3.OK() Then
                            Guid = CS3.GetText("ccGuid")
                            If Guid = "" Then
                                Guid = cp.Utils.CreateGuid()
                                Call CS3.SetField("ccGuid", Guid)
                            End If
                            s = s & vbCrLf & vbTab & "<IncludeAddon name=""" & cp.Utils.EncodeHTML(CS3.GetText("name")) & """ guid=""" & Guid & """/>"
                        End If
                        Call CS3.Close()
                        Call CS2.GoNext()
                    Loop
                    Call CS2.Close()
                    '
                    s = s & GetNodeBoolean("IsInline", CS, "IsInline")
                    '
                    ' -- javascript
                    If (Controllers.genericController.buildVersion(cp) < "5") Then
                        '
                        ' -- version 4.1 
                        s = s & GetNodeText("JavascriptOnLoad", CS, "JavascriptOnLoad")
                        s = s & GetNodeText("JavascriptInHead", CS, "JSFilename")
                        s = s & GetNodeText("JavascriptBodyEnd", CS, "JavascriptBodyEnd")
                    Else
                        '
                        ' -- version 5.0+
                        s = s & GetNodeText("JSHeadScriptSrc", CS, "JSHeadScriptSrc")
                        s = s & GetNodeText("JavascriptInHead", CS, "JSFilename")
                        s = s & GetNodeText("JSBodyScriptSrc", CS, "JSBodyScriptSrc")
                        s = s & GetNodeText("JavascriptBodyEnd", CS, "JavascriptBodyEnd")
                    End If
                    s = s & GetNodeText("MetaDescription", CS, ("MetaDescription"))
                    s = s & GetNodeText("OtherHeadTags", CS, "OtherHeadTags")
                    '
                    ' Placements
                    '
                    s = s & GetNodeBoolean("Content", CS, ("Content"))
                    s = s & GetNodeBoolean("Template", CS, ("Template"))
                    s = s & GetNodeBoolean("Email", CS, ("Email"))
                    s = s & GetNodeBoolean("Admin", CS, ("Admin"))
                    s = s & GetNodeBoolean("OnPageEndEvent", CS, ("OnPageEndEvent"))
                    s = s & GetNodeBoolean("OnPageStartEvent", CS, ("OnPageStartEvent"))
                    s = s & GetNodeBoolean("OnBodyStart", CS, ("OnBodyStart"))
                    s = s & GetNodeBoolean("OnBodyEnd", CS, ("OnBodyEnd"))
                    s = s & GetNodeBoolean("RemoteMethod", CS, ("RemoteMethod"))
                    's = s & GetNodeBoolean("OnNewVisitEvent", cs, ( "OnNewVisitEvent"))
                    '
                    ' Process
                    '
                    If ((LCase(addonName) = "oninstall") Or (LCase(addonName) = "_oninstall")) Then
                        s = s & GetNodeBoolean("ProcessRunOnce", True)
                    End If
                    s = s & GetNodeInteger("ProcessInterval", CS.GetInteger("ProcessInterval"))
                    '
                    ' Meta
                    '
                    s = s & GetNodeText("PageTitle", CS, ("PageTitle"))
                    s = s & GetNodeText("RemoteAssetLink", CS, ("RemoteAssetLink"))
                    '
                    ' Styles
                    '
                    s = s & GetNodeText("Styles", CS, "StylesFilename")
                    s = s & GetNodeText("StylesLinkHref", CS, "StylesLinkHref")
                    '
                    If (Controllers.genericController.buildVersion(cp) > "5") Then
                        '
                        ' -- v5 does not support styles in block or custom styles 
                        Styles = Trim(CS.GetText("StylesFilename"))
                    Else
                        '
                        ' -- v4
                        Styles = ""
                        If Not CS.GetBoolean("BlockDefaultStyles") Then
                            Styles = Trim(CS.GetText("StylesFilename"))
                        End If
                        StylesTest = Trim(CS.GetText("CustomStylesFilename"))
                        If StylesTest <> "" Then
                            If Styles <> "" Then
                                Styles = Styles & vbCrLf & StylesTest
                            Else
                                Styles = StylesTest
                            End If
                        End If
                        s = s & GetNodeText("Styles", Styles)
                    End If
                    '
                    ' Scripting
                    '
                    If (Controllers.genericController.buildVersion(cp) > "5") Then
                        '
                        ' -- version 5+
                        s = s & vbCrLf & vbTab & "<Scripting Language=""" & CS.GetText("ScriptingLanguageID") & """ EntryPoint=""" & CS.GetText("ScriptingEntryPoint") & """ Timeout=""" & CS.GetText("ScriptingTimeout") & """/>"
                    Else
                        '
                        ' -- version 4
                        NodeInnerText = Trim(CS.GetText("ScriptingCode"))
                        If NodeInnerText <> "" Then
                            NodeInnerText = vbCrLf & vbTab & vbTab & "<Code>" & EncodeCData(NodeInnerText) & "</Code>"
                        End If
                        CS2.Open("Add-on Scripting Module Rules", "addonid=" & addonid)
                        Do While CS2.OK()
                            ScriptingModuleID = CS2.GetInteger("ScriptingModuleID")
                            CS3.Open("Scripting Modules", "ID=" & ScriptingModuleID)
                            If CS3.OK() Then
                                Guid = CS3.GetText("ccGuid")
                                If Guid = "" Then
                                    Guid = cp.Utils.CreateGuid()
                                    Call CS3.SetField("ccGuid", Guid)
                                End If
                                Return_IncludeModuleGuidList = Return_IncludeModuleGuidList & vbCrLf & Guid
                                NodeInnerText = NodeInnerText & vbCrLf & vbTab & vbTab & "<IncludeModule name=""" & cp.Utils.EncodeHTML(CS3.GetText("name")) & """ guid=""" & Guid & """/>"
                            End If
                            Call CS3.Close()
                            Call CS2.GoNext()
                        Loop
                        Call CS2.Close()
                        If NodeInnerText = "" Then
                            s = s & vbCrLf & vbTab & "<Scripting Language=""" & CS.GetText("ScriptingLanguageID") & """ EntryPoint=""" & CS.GetText("ScriptingEntryPoint") & """ Timeout=""" & CS.GetText("ScriptingTimeout") & """/>"
                        Else
                            s = s & vbCrLf & vbTab & "<Scripting Language=""" & CS.GetText("ScriptingLanguageID") & """ EntryPoint=""" & CS.GetText("ScriptingEntryPoint") & """ Timeout=""" & CS.GetText("ScriptingTimeout") & """>" & NodeInnerText & vbCrLf & vbTab & "</Scripting>"
                        End If
                    End If
                    '
                    ' Shared Styles
                    '
                    If (Controllers.genericController.buildVersion(cp) > "5") Then
                        '
                        ' -- not supported in version 5
                    Else
                        '
                        ' -- v4 only
                        CS2.Open("Shared Styles Add-on Rules", "addonid=" & addonid)
                        Do While CS2.OK()
                            styleId = CS2.GetInteger("styleId")
                            CS3.Open("shared styles", "ID=" & styleId)
                            If CS3.OK() Then
                                Guid = CS3.GetText("ccGuid")
                                If Guid = "" Then
                                    Guid = cp.Utils.CreateGuid()
                                    Call CS3.SetField("ccGuid", Guid)
                                End If
                                Return_IncludeSharedStyleGuidList = Return_IncludeSharedStyleGuidList & vbCrLf & Guid
                                s = s & vbCrLf & vbTab & "<IncludeSharedStyle name=""" & cp.Utils.EncodeHTML(CS3.GetText("name")) & """ guid=""" & Guid & """/>"
                            End If
                            Call CS3.Close()
                            Call CS2.GoNext()
                        Loop
                        Call CS2.Close()
                    End If
                    '
                    ' Process Triggers
                    '
                    NodeInnerText = ""
                    CS2.Open("Add-on Content Trigger Rules", "addonid=" & addonid)
                    Do While CS2.OK()
                        TriggerContentID = CS2.GetInteger("ContentID")
                        CS3.Open("content", "ID=" & TriggerContentID)
                        If CS3.OK() Then
                            Guid = CS3.GetText("ccGuid")
                            If Guid = "" Then
                                Guid = cp.Utils.CreateGuid()
                                Call CS3.SetField("ccGuid", Guid)
                            End If
                            NodeInnerText = NodeInnerText & vbCrLf & vbTab & vbTab & "<ContentChange name=""" & cp.Utils.EncodeHTML(CS3.GetText("name")) & """ guid=""" & Guid & """/>"
                        End If
                        Call CS3.Close()
                        Call CS2.GoNext()
                    Loop
                    Call CS2.Close()
                    If NodeInnerText <> "" Then
                        s = s & vbCrLf & vbTab & "<ProcessTriggers>" & NodeInnerText & vbCrLf & vbTab & "</ProcessTriggers>"
                    End If
                    '
                    ' Editors
                    '
                    If cp.Content.IsField("Add-on Content Field Type Rules", "id") Then
                        NodeInnerText = ""
                        CS2.Open("Add-on Content Field Type Rules", "addonid=" & addonid)
                        Do While CS2.OK()
                            fieldTypeID = CS2.GetInteger("contentFieldTypeID")
                            fieldType = cp.Content.GetRecordName("Content Field Types", fieldTypeID)
                            If fieldType <> "" Then
                                NodeInnerText = NodeInnerText & vbCrLf & vbTab & vbTab & "<type>" & fieldType & "</type>"
                            End If
                            Call CS2.GoNext()
                        Loop
                        Call CS2.Close()
                        If NodeInnerText <> "" Then
                            s = s & vbCrLf & vbTab & "<Editors>" & NodeInnerText & vbCrLf & vbTab & "</Editors>"
                        End If
                    End If
                    '
                    '
                    '
                    Guid = CS.GetText("ccGuid")
                    If Guid = "" Then
                        Guid = cp.Utils.CreateGuid()
                        Call CS.SetField("ccGuid", Guid)
                    End If
                    NavType = CS.GetText("NavTypeID")
                    If NavType = "" Then
                        NavType = "Add-on"
                    End If
                    s = "" _
                        & vbCrLf & vbTab & "<Addon name=""" & cp.Utils.EncodeHTML(addonName) & """ guid=""" & Guid & """ type=""" & NavType & """>" _
                        & tabIndent(s) _
                        & vbCrLf & vbTab & "</Addon>"
                End If
                Call CS.Close()
                '
                GetAddonNode = s
            Catch ex As Exception
                errorReport(cp, ex, "GetAddonNode")
            End Try
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' create a simple text node with a name and content
        ''' </summary>
        ''' <param name="NodeName"></param>
        ''' <param name="fieldName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetNodeText(NodeName As String, cs As CPCSBaseClass, fieldName As String) As String
            GetNodeText = ""
            Try
                Dim fieldValue As String = cs.GetText(fieldName)
                If (String.IsNullOrEmpty(fieldValue)) Then
                    GetNodeText = vbCrLf & vbTab & "<" & NodeName & "></" & NodeName & ">"
                Else
                    GetNodeText = vbCrLf & vbTab & "<" & NodeName & ">" & EncodeCData(fieldValue) & "</" & NodeName & ">"
                End If
            Catch ex As Exception
                errorReport(cp, ex, "getNodeText")
            End Try
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' create a simple text node with a name and content
        ''' </summary>
        ''' <param name="NodeName"></param>
        ''' <param name="NodeContent"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetNodeText(NodeName As String, NodeContent As String) As String
            GetNodeText = ""
            Try
                GetNodeText = ""
                If NodeContent = "" Then
                    GetNodeText = GetNodeText & vbCrLf & vbTab & "<" & NodeName & "></" & NodeName & ">"
                Else
                    GetNodeText = GetNodeText & vbCrLf & vbTab & "<" & NodeName & ">" & EncodeCData(NodeContent) & "</" & NodeName & ">"
                End If
            Catch ex As Exception
                errorReport(cp, ex, "getNodeText")
            End Try
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' create a simple boolean node with a name and content
        ''' </summary>
        ''' <param name="NodeName"></param>
        ''' <param name="NodeContent"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetNodeBoolean(NodeName As String, NodeContent As Boolean) As String
            Try
                Return vbCrLf & vbTab & "<" & NodeName & ">" & kmaGetYesNo(NodeContent) & "</" & NodeName & ">"
            Catch ex As Exception
                errorReport(cp, ex, "GetNodeBoolean")
            End Try
            Return ""
        End Function
        '
        Private Function GetNodeBoolean(NodeName As String, cs As CPCSBaseClass, FieldName As String) As String
            Try
                Return GetNodeBoolean(NodeName, cs.GetBoolean(FieldName))
            Catch ex As Exception
                errorReport(cp, ex, "GetNodeBoolean")
            End Try
            Return ""
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' create a simple integer node with a name and content
        ''' </summary>
        ''' <param name="NodeName"></param>
        ''' <param name="NodeContent"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetNodeInteger(NodeName As String, NodeContent As Integer) As String
            GetNodeInteger = ""
            Try
                GetNodeInteger = vbCrLf & vbTab & "<" & NodeName & ">" & CStr(NodeContent) & "</" & NodeName & ">"
            Catch ex As Exception
                errorReport(cp, ex, "GetNodeInteger")
            End Try
        End Function
        '
        '====================================================================================================
        Function replaceMany(Source As String, ArrayOfSource() As String, ArrayOfReplacement() As String) As String
            replaceMany = ""
            Try
                Dim Count As Integer = UBound(ArrayOfSource) + 1
                replaceMany = Source
                For Pointer = 0 To Count - 1
                    replaceMany = Replace(replaceMany, ArrayOfSource(Pointer), ArrayOfReplacement(Pointer))
                Next
            Catch ex As Exception
                errorReport(cp, ex, "replaceMany")
            End Try
        End Function
        '
        '====================================================================================================
        Function encodeFilename(Filename As String) As String
            encodeFilename = ""
            Try
                Dim Source() As String
                Dim Replacement() As String
                '
                Source = {"""", "*", "/", ":", "<", ">", "?", "\", "|"}
                Replacement = {"_", "_", "_", "_", "_", "_", "_", "_", "_"}
                '
                encodeFilename = replaceMany(Filename, Source, Replacement)
                If Len(encodeFilename) > 254 Then
                    encodeFilename = Left(encodeFilename, 254)
                End If
            Catch ex As Exception
                errorReport(cp, ex, "encodeFilename")
            End Try
        End Function
        '
        '====================================================================================================
        Friend Sub GetLocalCollectionArgs(CollectionGuid As String, ByRef Return_CollectionPath As String, ByRef Return_LastChagnedate As Date)
            Try
                Const CollectionListRootNode = "collectionlist"
                '
                Dim LocalPath As String
                Dim LocalGuid As String = ""
                Dim Doc As New Xml.XmlDocument
                Dim CollectionNode As Xml.XmlNode
                Dim LocalListNode As Xml.XmlNode
                Dim CollectionFound As Boolean
                Dim CollectionPath As String = ""
                Dim LastChangeDate As Date
                Dim MatchFound As Boolean
                Dim LocalName As String
                '
                MatchFound = False
                Return_CollectionPath = ""
                Return_LastChagnedate = Date.MinValue
                Call Doc.LoadXml(GetConfig)
                If True Then
                    If LCase(Doc.DocumentElement.Name) <> LCase(CollectionListRootNode) Then
                        'Call AppendClassLogFile("Server", "", "GetLocalCollectionArgs, Hint=[" & Hint & "], The Collections.xml file has an invalid root node, [" & Doc.documentElement.name & "] was received and [" & CollectionListRootNode & "] was expected.")
                    Else
                        With Doc.DocumentElement
                            If LCase(.Name) <> "collectionlist" Then
                                'Call AppendClassLogFile("Server", "", "GetLocalCollectionArgs, basename was not collectionlist, [" & .name & "].")
                            Else
                                CollectionFound = False
                                'hint = hint & ",checking nodes [" & .childNodes.length & "]"
                                For Each LocalListNode In .ChildNodes
                                    LocalName = "no name found"
                                    LocalPath = ""
                                    Select Case LCase(LocalListNode.Name)
                                        Case "collection"
                                            LocalGuid = ""
                                            For Each CollectionNode In LocalListNode.ChildNodes
                                                Select Case LCase(CollectionNode.Name)
                                                    Case "name"
                                                        '
                                                        LocalName = LCase(CollectionNode.InnerText)
                                                    Case "guid"
                                                        '
                                                        LocalGuid = LCase(CollectionNode.InnerText)
                                                    Case "path"
                                                        '
                                                        CollectionPath = LCase(CollectionNode.InnerText)
                                                    Case "lastchangedate"
                                                        LastChangeDate = cp.Utils.EncodeDate(CollectionNode.InnerText)
                                                End Select
                                            Next
                                    End Select
                                    'hint = hint & ",checking node [" & LocalName & "]"
                                    If LCase(CollectionGuid) = LocalGuid Then
                                        Return_CollectionPath = CollectionPath
                                        Return_LastChagnedate = LastChangeDate
                                        'Call AppendClassLogFile("Server", "GetCollectionConfigArg", "GetLocalCollectionArgs, match found, CollectionName=" & LocalName & ", CollectionPath=" & CollectionPath & ", LastChangeDate=" & LastChangeDate)
                                        MatchFound = True
                                        Exit For
                                    End If
                                Next
                            End If
                        End With
                    End If
                End If
                If Not MatchFound Then
                    'Call AppendClassLogFile("Server", "GetCollectionConfigArg", "GetLocalCollectionArgs, no local collection match found, Hint=[" & Hint & "]")
                End If
            Catch ex As Exception
                errorReport(cp, ex, "GetLocalCollectionArgs")
            End Try
        End Sub
        '
        '====================================================================================================
        Public Function GetConfig() As String
            GetConfig = ""
            Try
                Dim AddonPath As String
                '
                AddonPath = cp.Site.PhysicalInstallPath & "\addons"
                AddonPath = AddonPath & "\Collections.xml"
                GetConfig = cp.File.Read(AddonPath)
            Catch ex As Exception
                errorReport(cp, ex, "GetConfig")
            End Try
        End Function
        '
        '====================================================================================================
        Private Function AddCompatibilityResources(CollectionPath As String, ArchiveFilename As String, SubPath As String) As String
            AddCompatibilityResources = ""
            Dim s As String = ""
            Try
                Dim AddFilename As String
                Dim FileExt As String
                Dim FileList As String
                Dim Files() As String
                Dim Filename As String
                Dim Ptr As Integer
                Dim FileArgs() As String
                Dim FolderList As String
                Dim Folders() As String
                Dim FolderArgs() As String
                Dim Folder As String
                Dim Pos As Integer
                '
                ' Process all SubPaths
                '
                FolderList = cp.File.folderList(CollectionPath & SubPath)
                If FolderList <> "" Then
                    Folders = Split(FolderList, vbCrLf)
                    For Ptr = 0 To UBound(Folders)
                        Folder = Folders(Ptr)
                        If Folder <> "" Then
                            FolderArgs = Split(Folders(Ptr), ",")
                            Folder = FolderArgs(0)
                            If Folder <> "" Then
                                s = s & AddCompatibilityResources(CollectionPath, ArchiveFilename, SubPath & Folder & "\")
                            End If
                        End If
                    Next
                End If
                '
                ' Process files in this path
                '
                'Set Remote = CreateObject("ccRemote.RemoteClass")
                FileList = cp.File.fileList(CollectionPath)
                If FileList <> "" Then
                    Files = Split(FileList, vbCrLf)
                    For Ptr = 0 To UBound(Files)
                        Filename = Files(Ptr)
                        If Filename <> "" Then
                            FileArgs = Split(Filename, ",")
                            If UBound(FileArgs) > 0 Then
                                Filename = FileArgs(0)
                                Pos = InStrRev(Filename, ".")
                                FileExt = ""
                                If Pos > 0 Then
                                    FileExt = Mid(Filename, Pos + 1)
                                End If
                                If LCase(Filename) = "collection.hlp" Then
                                    '
                                    ' legacy help system, ignore this file
                                    '
                                ElseIf LCase(FileExt) = "xml" Then
                                    '
                                    ' compatibility resources can not include an xml file in the wwwroot
                                    '
                                ElseIf InStr(1, CollectionPath, "\ContensiveFiles\", vbTextCompare) <> 0 Then
                                    '
                                    ' Content resources
                                    '
                                    s = s & vbCrLf & vbTab & "<Resource name=""" & cp.Utils.EncodeHTML(Filename) & """ type=""content"" path=""" & cp.Utils.EncodeHTML(SubPath) & """ />"
                                    AddFilename = CollectionPath & SubPath & "\" & Filename
                                    'Call zipFile(ArchiveFilename, AddFilename)
                                    'Call runAtServer("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable(AddFilename))
                                    'Call Remote.executeCmd("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable(AddFilename))
                                ElseIf LCase(FileExt) = "dll" Then
                                    '
                                    ' Executable resources
                                    '
                                    s = s & vbCrLf & vbTab & "<Resource name=""" & cp.Utils.EncodeHTML(Filename) & """ type=""executable"" path=""" & cp.Utils.EncodeHTML(SubPath) & """ />"
                                    AddFilename = CollectionPath & SubPath & "\" & Filename
                                    'Call zipFile(ArchiveFilename, AddFilename)
                                    'Call runAtServer("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable(AddFilename))
                                    'Call Remote.executeCmd("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable(AddFilename))
                                Else
                                    '
                                    ' www resources
                                    '
                                    s = s & vbCrLf & vbTab & "<Resource name=""" & cp.Utils.EncodeHTML(Filename) & """ type=""www"" path=""" & cp.Utils.EncodeHTML(SubPath) & """ />"
                                    AddFilename = CollectionPath & SubPath & "\" & Filename
                                    'Call zipFile(ArchiveFilename, AddFilename)
                                    'Call runAtServer("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable(AddFilename))
                                    'Call Remote.executeCmd("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable(AddFilename))
                                End If
                            End If
                        End If
                    Next
                End If
                '
                AddCompatibilityResources = s
            Catch ex As Exception
                errorReport(cp, ex, "GetNodeInteger")
            End Try
        End Function
        '
        '====================================================================================================
        Friend Function EncodeCData(Source As String) As String
            EncodeCData = ""
            Try
                EncodeCData = Source
                If EncodeCData <> "" Then
                    EncodeCData = "<![CDATA[" & Replace(EncodeCData, "]]>", "]]]]><![CDATA[>") & "]]>"
                End If
            Catch ex As Exception
                errorReport(cp, ex, "EncodeCData")
            End Try
        End Function
        '
        '====================================================================================================
        Public Function kmaGetYesNo(Key As Boolean) As String
            If Key Then
                kmaGetYesNo = "Yes"
            Else
                kmaGetYesNo = "No"
            End If
        End Function
        '
        '=======================================================================================
        ''' <summary>
        ''' zip
        ''' </summary>
        ''' <param name="PathFilename"></param>
        ''' <remarks></remarks>
        Public Sub UnzipFile(ByVal PathFilename As String)
            Try
                '
                Dim fastZip As ICSharpCode.SharpZipLib.Zip.FastZip = New ICSharpCode.SharpZipLib.Zip.FastZip()
                Dim fileFilter As String = Nothing

                fastZip.ExtractZip(PathFilename, getPath(PathFilename), fileFilter)                '
            Catch ex As Exception
                errorReport(cp, ex, "UnzipFile")
            End Try
        End Sub        '
        '
        '=======================================================================================
        ''' <summary>
        ''' unzip
        ''' </summary>
        ''' <param name="archivePathFilename"></param>
        ''' <param name="addPathFilename"></param>
        ''' <remarks></remarks>
        Public Sub zipFile(archivePathFilename As String, ByVal addPathFilename As List(Of String))
            Try
                '
                'Dim fastZip As FastZip = New ICSharpCode.SharpZipLib.Zip.FastZip()
                'Dim fileFilter As String = Nothing
                'Dim recurse As Boolean = True
                'Dim archivepath As String = getPath(archivePathFilename)
                'Dim archiveFilename As String = GetFilename(archivePathFilename)
                '
                Dim z As Zip.ZipFile
                If cp.File.fileExists(archivePathFilename) Then
                    '
                    ' update existing zip with list of files
                    '
                    z = New Zip.ZipFile(archivePathFilename)
                Else
                    '
                    ' create new zip
                    '
                    z = Zip.ZipFile.Create(archivePathFilename)
                End If
                z.BeginUpdate()
                For Each pathFilename In addPathFilename
                    If Not cp.File.fileExists(pathFilename) Then
                        ' -- should report this back to user
                        cp.UserError.Add("During export, this file was included in the collections file list, but was not found on the server. Either remove it from the resources tab of the collection record, or add it to the addons folder. [" & pathFilename & "]")
                    Else
                        z.Add(pathFilename, System.IO.Path.GetFileName(pathFilename))
                    End If
                Next
                z.CommitUpdate()
                z.Close()
                'fastZip.CreateZip(archivePathFilename, addPathFilename, recurse, fileFilter)
            Catch ex As Exception
                errorReport(cp, ex, "zipFile")
            End Try
        End Sub        '
        '
        '=======================================================================================
        Private Function getPath(ByVal pathFilename As String) As String
            getPath = ""
            Try
                Dim Position As Integer
                '
                Position = InStrRev(pathFilename, "\")
                If Position <> 0 Then
                    getPath = Mid(pathFilename, 1, Position)
                End If
            Catch ex As Exception
                errorReport(cp, ex, "getPath")
            End Try
        End Function
        '
        '=======================================================================================
        Public Function GetFilename(ByVal PathFilename As String) As String
            Dim Position As Integer
            '
            GetFilename = PathFilename
            Position = InStrRev(GetFilename, "/")
            If Position <> 0 Then
                GetFilename = Mid(GetFilename, Position + 1)
            End If
        End Function
        '
        '=======================================================================================
        '
        '   Indent every line by 1 tab
        '
        Public Function tabIndent(Source As String) As String
            Dim posStart As Integer
            Dim posEnd As Integer
            Dim pre As String
            Dim post As String
            Dim target As String
            '
            posStart = InStr(1, Source, "<![CDATA[", CompareMethod.Text)
            If posStart = 0 Then
                '
                ' no cdata
                '
                posStart = InStr(1, Source, "<textarea", CompareMethod.Text)
                If posStart = 0 Then
                    '
                    ' no textarea
                    '
                    tabIndent = Replace(Source, vbCrLf & vbTab, vbCrLf & vbTab & vbTab)
                Else
                    '
                    ' text area found, isolate it and indent before and after
                    '
                    posEnd = InStr(posStart, Source, "</textarea>", CompareMethod.Text)
                    pre = Mid(Source, 1, posStart - 1)
                    If posEnd = 0 Then
                        target = Mid(Source, posStart)
                        post = ""
                    Else
                        target = Mid(Source, posStart, posEnd - posStart + Len("</textarea>"))
                        post = Mid(Source, posEnd + Len("</textarea>"))
                    End If
                    tabIndent = pre & target & post
                    'tabIndent = tabIndent(pre) & target & tabIndent(post)
                End If
            Else
                '
                ' cdata found, isolate it and indent before and after
                '
                posEnd = InStr(posStart, Source, "]]>", CompareMethod.Text)
                pre = Mid(Source, 1, posStart - 1)
                If posEnd = 0 Then
                    target = Mid(Source, posStart)
                    post = ""
                Else
                    target = Mid(Source, posStart, posEnd - posStart + Len("]]>"))
                    post = Mid(Source, posEnd + 3)
                End If
                tabIndent = pre & target & post
                'tabIndent = tabIndent(pre) & target & tabIndent(post)
            End If
            '    kmaIndent = Source
            '    If InStr(1, kmaIndent, "<textarea", vbTextCompare) = 0 Then
            '        kmaIndent = Replace(Source, vbCrLf & vbTab, vbCrLf & vbTab & vbTab)
            '    End If
        End Function

    End Class
End Namespace
