VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExportClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
'========================================================================
'
' Sample file to create a Contensive Aggregate Object
'
'========================================================================
'
Private Main As Object
'Private Main As ccWeb3.MainClass
Private Csv As Object
'
Const RequestnameVer40Compatibility = "Version40Compatibility"
'
Const FormIDSelectCollection = 0
Const FormIDDisplayResults = 1
'
Const RequestNameButton = "button"
Const RequestNameFormID = "formid"
Const RequestnameExecutableFile = "executablefile"
Const RequestNameCollectionID = "collectionid"
'
'Private Main As Object
'
'=================================================================================
'   Execute Method, v3.4 Interface
'=================================================================================
'
Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Ver40Compatibility As Boolean
    Dim AddExecutableFilename As String
    Dim Argument1 As String
    Dim Argument2 As String
    Dim DoSomethingPtr As Long
    Dim Button As String
    Dim FormID As Long
    'Dim RequestNameCollectionID As Long
    Dim CollectionID As Long
    Dim CollectionName As String
    Dim CollectionFile As String
    Dim CollectionFilename As String
    Dim s As String
    '
    ' Load any Add-on Instance Options selected by the administrator
    '
    Set Main = MainObject
    Set Csv = CsvObject
    Argument1 = CsvObject.GetAddonOption("Argument1", OptionString)
    Argument2 = CsvObject.GetAddonOption("Argument2", OptionString)
    '
    ' Every form returns a button and a formid
    '
    Button = Main.GetStreamText("button")
    FormID = Main.GetStreamInteger("formid")
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
                '
                CollectionID = Main.GetStreamInteger(RequestNameCollectionID)
                CollectionName = Main.GetRecordName("Add-on Collections", CollectionID)
                If CollectionName = "" Then
                    Call Main.AddUserError("The collection file you selected could not be found. Please select another.")
                Else
                    Ver40Compatibility = Main.GetStreamBoolean(RequestnameVer40Compatibility)
                    AddExecutableFilename = Main.GetStreamText(RequestnameExecutableFile)
                    If AddExecutableFilename <> "" Then
                        AddExecutableFilename = "CollectionExport\" & AddExecutableFilename
                        Call Main.SaveStreamFile(RequestnameExecutableFile, "CollectionExport\")
                    End If
                    CollectionFilename = GetCollection(CollectionID, AddExecutableFilename, Ver40Compatibility)
                    'CollectionFile = GetCollection(CollectionID)
                    'If CollectionFile = "" Then
                    '    Call Main.AddUserError("There was a problem creating the collection file.")
                    'Else
                    '    CollectionFilename = "upload\" & CollectionName & ".xml"
                    '    Call Main.SaveVirtualFile(CollectionFilename, CollectionFile)
                    'End If
                End If
                If Main.IsUserError() Then
                    FormID = FormIDSelectCollection
                Else
                    
                    FormID = FormIDDisplayResults
                End If
        End Select
    End If
    '
    ' Reply with the next form
    '
    Select Case FormID
        Case FormIDDisplayResults
            '
            ' Diplay the results page
            '
            s = Main.GetUserError() _
                & vbCrLf & vbTab & "<div class=""responseForm"">" _
                & vbCrLf & vbTab & vbTab & "<p>Click <a href=""" & Main.serverFilePath & Replace(CollectionFilename, "\", "/") & """>here</a> to download the collection file</p>" _
                & vbCrLf & vbTab & "</div>"
        Case Else
            '
            ' ask them to select a collectioin to export
            '
            s = "" _
                & vbCrLf & vbTab & "<div class=""mainForm"">" _
                & vbCrLf & vbTab & vbTab & Main.GetUserError() _
                & vbCrLf & vbTab & vbTab & Main.GetUploadFormStart() _
                & vbCrLf & vbTab & vbTab & vbTab & "<p>Select a collection to be exported. If the project is being developed and you need to add an executable resource that is not installed as an add-on on this site, use the file upload.</p>" _
                & vbCrLf & vbTab & vbTab & vbTab & "<p>" & Main.GetFormInputSelect(RequestNameCollectionID, 0, "Add-on Collections") & "<br>The collection to export</p>" _
                & vbCrLf & vbTab & vbTab & vbTab & "<p>" & Main.GetFormInputCheckBox(RequestnameVer40Compatibility) & "&nbsp;Version 4.0 Compatible</p>" _
                & vbCrLf & vbTab & vbTab & vbTab & "<p>" & Main.GetFormInputFile(RequestnameExecutableFile) & "<br>An optional executable resource to be added to the collection</p>" _
                & vbCrLf & vbTab & vbTab & vbTab & "<p>" & Main.GetFormButton2("Export Collection") & "</p>" _
                & vbCrLf & vbTab & vbTab & Main.GetFormEnd() _
                & vbCrLf & vbTab & "</div>"
    End Select
    '
    Execute = "" _
        & vbCrLf & vbTab & "<div class=""collectionExport"">" _
        & kmaIndent(s) _
        & vbCrLf & vbTab & "</div>"
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("Execute")
End Function
'
'
'
Private Function GetCollection(CollectionID As Long, AddExecutableFilename As String, Version40Compatibility As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim IncludeSharedStyleGuidList As String
    Dim isUpdatable As Boolean
    Dim fieldLookupListValue As String
    Dim CSlookup As Long
    Dim FieldValueInteger As Long
    Dim FieldLookupContentName As String
    Dim fieldPtr As Long
    Dim fieldCnt As Long
    Dim fieldNames() As String
    Dim fieldTypes() As Long
    Dim fieldLookupContent() As String
    Dim fieldLookupList() As String
    Dim FieldLookupContentID As Long
    Dim Criteria As String
    Dim supportsGuid As Boolean
    Dim reload As Boolean
    Dim ContentID As Long
    Dim FieldValue As String
    Dim FieldSkip As Boolean
    Dim FieldTypeNumber As Long
    Dim CSData As Long
    Dim DataRecordList As String
    Dim DataRecords() As String
    Dim DataRecord As String
    Dim DataSplit() As String
    Dim DataContentName As String
    Dim DataContentId As Long
    Dim DataRecordGuid As String
    Dim DataRecordName As String
    Dim TestString As String
    Dim FieldName As String
    Dim FieldNodes As String
    Dim RecordNodes As String
    Dim Modules() As String
    Dim ModuleGuid As String
    Dim Code As String
    Dim ManualFilename As String
    Dim FileArgs() As String
    Dim ResourceCnt As Long
    Dim Remote As Object
    Dim ContentName As String
    Dim FileList As String
    Dim Files() As String
    Dim Ptr As Long
    Dim PathFilename As String
    Dim Filename As String
    Dim Path As String
    Dim Pos As Long
    Dim s As String
    Dim CS As Long
    Dim CS2 As Long
    Dim CS3 As Long
    Dim GuidGenerator As New GuidGenerator
    Dim Node As String
    Dim CollectionGuid As String
    Dim Guid As String
    Dim ArchiveFilename As String
    Dim ArchivePath As String
    Dim InstallFilename As String
    Dim CollectionName As String
    Dim AddFilename As String
    Dim PhysicalWWWPath As String
    Dim CollectionPath As String
    Dim LastChangeDate As Date
    Dim AddonPath As String
    Dim fs As New kmaFileSystem3.FileSystemClass
    Dim AddFileList As String
    Dim AddFileListFilename As String
    Dim IncludeModuleGuidList As String
    Dim Version40DLLList As String
    Dim ExecFileListNode As String
    Dim blockNavigatorNode As Boolean
    '
    IncludeSharedStyleGuidList = ""
    '
    CS = Main.OpenCSContentRecord("Add-on Collections", CollectionID)
    If Not Main.IsCSOK(CS) Then
        Call Main.AddUserError("The collection you selected could not be found")
    Else
        CollectionGuid = Main.GetCSText(CS, "ccGuid")
        If CollectionGuid = "" Then
            CollectionGuid = GuidGenerator.CreateGUID("")
            Call Main.SetCS(CS, "ccGuid", CollectionGuid)
        End If
        CollectionName = Main.GetCSText(CS, "name")
        If Not Main.IsCSFieldSupported(CS, "updatable") Then
            isUpdatable = True
        Else
            isUpdatable = Main.GetCSBoolean(CS, "updatable")
        End If
        If Not Main.IsCSFieldSupported(CS, "blockNavigatorNode") Then
            blockNavigatorNode = False
        Else
            blockNavigatorNode = Main.GetCSBoolean(CS, "blockNavigatorNode")
        End If
        s = "" _
            & "<?xml version=""1.0"" encoding=""windows-1252""?>" _
            & vbCrLf & "<Collection name=""" & kmaEncodeHTML(CollectionName) & """ guid=""" & CollectionGuid & """ system=""" & kmaGetYesNo(Main.GetCSBoolean(CS, "system")) & """ updatable=""" & kmaGetYesNo(isUpdatable) & """ blockNavigatorNode=""" & kmaGetYesNo(blockNavigatorNode) & """>"
        '
        ' Archive Filenames
        '
        ArchivePath = Main.PhysicalFilePath & "CollectionExport\"
        InstallFilename = encodeFilename(CollectionName & ".xml")
        InstallFilename = ArchivePath & InstallFilename
        ArchiveFilename = encodeFilename(CollectionName & ".zip")
        ArchiveFilename = ArchivePath & ArchiveFilename
        AddFileListFilename = ArchivePath & "AddFileList.txt"
        GetCollection = "CollectionExport\" & encodeFilename(CollectionName & ".zip")
        '
        ' Delete old archive file
        '
        Call fs.DeleteFile(ArchiveFilename)
        '
        ' Manual Upload Executable
        ' Do not add node yet, just get the file and make the node so executables can be added for V4 compatibility
        '
        Version40DLLList = ""
        If AddExecutableFilename <> "" Then
            '
            ' Add the uploaded executable
            '
            PathFilename = Replace(AddExecutableFilename, "\", "/")
            Path = ""
            ManualFilename = PathFilename
            Pos = InStrRev(ManualFilename, "/")
            If Pos > 0 Then
                ManualFilename = Mid(ManualFilename, Pos + 1)
            End If
            AddFilename = Main.PhysicalFilePath & AddExecutableFilename
            If InStr(1, AddFileList, "\" & ManualFilename, vbTextCompare) <> 0 Then
                Call Main.AddUserError("There was an error exporting this collection because there were multiple files with the same filename [" & AddExecutableFilename & "]")
            Else
                ExecFileListNode = ExecFileListNode & vbCrLf & vbTab & "<Resource name=""" & kmaEncodeHTML(ManualFilename) & """ type=""executable"" path=""" & kmaEncodeHTML(Path) & """ />"
                AddFileList = AddFileList & vbCrLf & AddFilename
            End If
            Version40DLLList = Version40DLLList & vbCrLf & ManualFilename
        End If
        
        
        '
        ' Build executable file list Resource Node so executables can be added to addons for Version40compatibility
        '   but save it for the end, executableFileList
        '
        AddonPath = Replace(Main.PhysicalccLibPath(), "\cclib", "\addons\", , , vbTextCompare)
        FileList = Main.GetCSText(CS, "execFileList")
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
                        If InStr(1, AddFileList, "\" & Filename, vbTextCompare) <> 0 Then
                            Call Main.AddUserError("There was an error exporting this collection because there were multiple files with the same filename [" & Filename & "]")
                        Else
                            ExecFileListNode = ExecFileListNode & vbCrLf & vbTab & "<Resource name=""" & kmaEncodeHTML(Filename) & """ type=""executable"" path=""" & kmaEncodeHTML(Path) & """ />"
                            AddFileList = AddFileList & vbCrLf & AddFilename
                        End If
                        Version40DLLList = Version40DLLList & vbCrLf & Filename
                    End If
                    ResourceCnt = ResourceCnt + 1
                End If
            Next
        End If
        If (ResourceCnt = 0) And (CollectionPath <> "") Then
            '
            ' If no resources were in the collection record, this might be an old installation
            ' Add all .dll files in the CollectionPath
            '
            ExecFileListNode = ExecFileListNode & AddCompatibilityResources(AddonPath & CollectionPath, ArchiveFilename, "", Version40DLLList)
        End If
        '
        ' helpLink
        '
        If Main.IsCSFieldSupported(CS, "HelpLink") Then
            s = s & vbCrLf & vbTab & "<HelpLink>" & kmaEncodeHTML(Main.GetCSText(CS, "HelpLink")) & "</HelpLink>"
        End If
        '
        ' Help
        '
        s = s & vbCrLf & vbTab & "<Help>" & kmaEncodeHTML(Main.GetCSText(CS, "Help")) & "</Help>"
        '
        ' Addons
        '
        CS2 = Main.OpenCSContent("Add-ons", "collectionid=" & CollectionID, , , "Process")
        Do While Main.IsCSOK(CS2)
            s = s & GetAddonNode(Main.GetCSInteger(CS2, "id"), IncludeModuleGuidList, Version40Compatibility, Version40DLLList, IncludeSharedStyleGuidList)
            Call Main.NextCSRecord(CS2)
        Loop
        '
        ' Data Records
        '
        DataRecordList = Main.GetCSText(CS, "DataRecordList")
        If DataRecordList <> "" Then
            DataRecords = Split(DataRecordList, vbCrLf)
            RecordNodes = ""
            For Ptr = 0 To UBound(DataRecords)
                FieldNodes = ""
                DataRecordName = ""
                DataRecordGuid = ""
                DataRecord = DataRecords(Ptr)
                If DataRecord <> "" Then
                    DataSplit = Split(DataRecord, ",")
                    If UBound(DataSplit) >= 0 Then
                        DataContentName = Trim(DataSplit(0))
                        DataContentId = Main.GetContentID(DataContentName)
                        If DataContentId <= 0 Then
                            RecordNodes = "" _
                                & RecordNodes _
                                & vbCrLf & vbTab & "<!-- data missing, content not found during export, content=""" & DataContentName & """ guid=""" & DataRecordGuid & """ name=""" & DataRecordName & """ -->"
                        Else
                            supportsGuid = Main.IsContentFieldSupported(DataContentName, "ccguid")
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
                                    Criteria = "name=" & KmaEncodeSQLText(DataRecordName)
                                ElseIf (Len(TestString) = 38) And (Left(TestString, 1) = "{") And (Right(TestString, 1) = "}") Then
                                    '
                                    ' guid {726ED098-5A9E-49A9-8840-767A74F41D01} format
                                    '
                                    DataRecordGuid = TestString
                                    DataRecordName = ""
                                    Criteria = "ccguid=" & KmaEncodeSQLText(DataRecordGuid)
                                ElseIf (Len(TestString) = 36) And (Mid(TestString, 9, 1) = "-") Then
                                    '
                                    ' guid 726ED098-5A9E-49A9-8840-767A74F41D01 format
                                    '
                                    DataRecordGuid = TestString
                                    DataRecordName = ""
                                    Criteria = "ccguid=" & KmaEncodeSQLText(DataRecordGuid)
                                ElseIf (Len(TestString) = 32) And (InStr(1, TestString, " ") = 0) Then
                                    '
                                    ' guid 726ED0985A9E49A98840767A74F41D01 format
                                    '
                                    DataRecordGuid = TestString
                                    DataRecordName = ""
                                    Criteria = "ccguid=" & KmaEncodeSQLText(DataRecordGuid)
                                Else
                                    '
                                    ' use name
                                    '
                                    DataRecordName = TestString
                                    DataRecordGuid = ""
                                    Criteria = "name=" & KmaEncodeSQLText(DataRecordName)
                                End If
                            End If
                            CSData = Main.OpenCSContent(DataContentName, Criteria, "id")
                            If Not Main.IsCSOK(CSData) Then
                                RecordNodes = "" _
                                    & RecordNodes _
                                    & vbCrLf & vbTab & "<!-- data missing, record not found during export, content=""" & DataContentName & """ guid=""" & DataRecordGuid & """ name=""" & DataRecordName & """ -->"
                            Else
                                '
                                ' determine all valid fields
                                '
                                fieldCnt = 0
                                FieldName = Main.GetCSFirstFieldName(CSData)
                                Do While (FieldName <> "")
                                    FieldName = Main.GetCSNextFieldName(CSData)
                                    FieldLookupContentID = 0
                                    FieldLookupContentName = ""
                                    If FieldName <> "" Then
                                        Select Case LCase(FieldName)
                                            Case "ccguid", "name", "id", "dateadded", "createdby", "modifiedby", "modifieddate", "createkey", "contentcontrolid", "editsourceid", "editarchive", "editblank", "contentcategoryid"
                                            Case Else
                                                FieldTypeNumber = kmaEncodeInteger(Main.GetContentFieldProperty(DataContentName, FieldName, "fieldtype"))
                                                If FieldTypeNumber = 7 Then
                                                    CSlookup = Main.OpenCSSQL("", "select top 1 LookupContentId,LookupList from ccFields where contentid=" & DataContentId & " and name=" & KmaEncodeSQLText(FieldName))
                                                    If Main.IsCSOK(CSlookup) Then
                                                        FieldLookupContentID = Main.GetCSInteger(CSlookup, "Lookupcontentid")
                                                        fieldLookupListValue = Main.GetCSText(CSlookup, "LookupList")
                                                    End If
                                                    Call Main.CloseCS(CSlookup)
                                                    If FieldLookupContentID <> 0 Then
                                                        FieldLookupContentName = Main.GetContentNameByID(FieldLookupContentID)
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
                                Loop
                                '
                                ' output records
                                '
                                DataRecordGuid = ""
                                Do While Main.IsCSOK(CSData)
                                    FieldNodes = ""
                                    DataRecordName = Main.GetCS(CSData, "name")
                                    If supportsGuid Then
                                        DataRecordGuid = Main.GetCS(CSData, "ccguid")
                                        If DataRecordGuid = "" Then
                                            DataRecordGuid = GuidGenerator.CreateGUID("")
                                            Call Main.SetCS(CSData, "ccGuid", DataRecordGuid)
                                        End If
                                    End If
                                    For fieldPtr = 0 To fieldCnt - 1
                                        FieldName = fieldNames(fieldPtr)
                                        FieldTypeNumber = kmaEncodeInteger(fieldTypes(fieldPtr))
                                        Select Case FieldTypeNumber
                                            Case FieldTypeBoolean
                                                '
                                                ' true/false
                                                '
                                                FieldValue = Main.GetCSBoolean(CSData, FieldName)
                                            Case FieldTypeCSSFile, FieldTypeJavascriptFile, FieldTypeTextFile, FieldTypeXMLFile
                                                '
                                                ' text files
                                                '
                                                FieldValue = Main.GetCS(CSData, FieldName)
                                                FieldValue = EncodeCData(FieldValue)
                                            Case FieldTypeCurrency, FieldTypeFloat, FieldTypeInteger
                                                '
                                                ' numbers
                                                '
                                                FieldValue = Main.GetCSNumber(CSData, FieldName)
                                            Case FieldTypeDate
                                                '
                                                ' date
                                                '
                                                FieldValue = Main.GetCSDate(CSData, FieldName)
                                            Case FieldTypeLookup
                                                '
                                                ' lookup
                                                '
                                                FieldValue = ""
                                                FieldValueInteger = Main.GetCSInteger(CSData, FieldName)
                                                If (FieldValueInteger <> 0) Then
                                                    FieldLookupContentName = fieldLookupContent(fieldPtr)
                                                    fieldLookupListValue = fieldLookupList(fieldPtr)
                                                    If (FieldLookupContentName <> "") Then
                                                        '
                                                        ' content lookup
                                                        '
                                                        If Main.IsContentFieldSupported(FieldLookupContentName, "ccguid") Then
                                                            CSlookup = Main.OpenCSContentRecord(FieldLookupContentName, FieldValueInteger)
                                                            If Main.IsCSOK(CSlookup) Then
                                                                FieldValue = Main.GetCSText(CSlookup, "ccguid")
                                                                If FieldValue = "" Then
                                                                    FieldValue = GuidGenerator.CreateGUID("")
                                                                    Call Main.SetCS(CSlookup, "ccGuid", FieldValue)
                                                                End If
                                                            End If
                                                            Call Main.CloseCS(CSlookup)
                                                        End If
                                                    ElseIf fieldLookupListValue <> "" Then
                                                        '
                                                        ' list lookup, ok to save integer
                                                        '
                                                        FieldValue = FieldValueInteger
                                                    End If
                                                End If
                                            Case Else
                                                '
                                                ' text types
                                                '
                                                FieldValue = Main.GetCS(CSData, FieldName)
                                                FieldValue = EncodeCData(FieldValue)
                                        End Select
                                        FieldNodes = FieldNodes & vbCrLf & vbTab & "<field name=""" & kmaEncodeHTML(FieldName) & """>" & FieldValue & "</field>"
                                    Next
                                    RecordNodes = "" _
                                        & RecordNodes _
                                        & vbCrLf & vbTab & "<record content=""" & kmaEncodeHTML(DataContentName) & """ guid=""" & DataRecordGuid & """ name=""" & kmaEncodeHTML(DataRecordName) & """>" _
                                        & kmaIndent(FieldNodes) _
                                        & vbCrLf & vbTab & "</record>"
                                    Call Main.NextCSRecord(CSData)
                                Loop
                            End If
                            Call Main.CloseCS(CSData)
                            ' include name only records
                            'If FieldNodes <> "" Then
                            'End If
                        End If
                    End If
                End If
            Next
            If RecordNodes <> "" Then
                s = "" _
                    & s _
                    & vbCrLf & vbTab & "<data>" _
                    & kmaIndent(RecordNodes) _
                    & vbCrLf & vbTab & "</data>"
            End If
        End If
        '
        ' CDef
        '
        CS2 = Main.OpenCSContent("Add-on Collection CDef Rules", "CollectionID=" & CollectionID)
        Do While Main.IsCSOK(CS2)
            ContentID = Main.GetCSInteger(CS2, "contentid")
            '
            ' get name and make sure there is a guid
            '
            reload = False
            CS3 = Main.OpenCSContentRecord("content", ContentID)
            If Main.IsCSOK(CS3) Then
                ContentName = Main.GetCS(CS3, "name")
                If Main.GetCSText(CS3, "ccguid") = "" Then
                    Call Main.SetCS(CS3, "ccGuid", GuidGenerator.CreateGUID(""))
                    reload = True
                End If
            End If
            Call Main.CloseCS(CS3)
            If reload Then
                Call Main.LoadContentDefinitions
            End If
            '
            Node = Csv.GetXMLContentDefinition(ContentName)
            '
            ' remove the <collection> top node
            '
            Pos = InStr(1, Node, "<cdef", vbTextCompare)
            If Pos > 0 Then
                Node = Mid(Node, Pos)
                Pos = InStr(1, Node, "</cdef>", vbTextCompare)
                If Pos > 0 Then
                    Node = Mid(Node, 1, Pos + 6)
                    s = s & vbCrLf & vbTab & Node
                End If
            End If
            Call Main.NextCSRecord(CS2)
        Loop
        Call Main.CloseCS(CS2)
        '
        ' Scripting Modules
        '
        If IncludeModuleGuidList <> "" Then
            Modules = Split(IncludeModuleGuidList, vbCrLf)
            For Ptr = 0 To UBound(Modules)
                ModuleGuid = Modules(Ptr)
                If ModuleGuid <> "" Then
                    CS2 = Main.OpenCSContent("Scripting Modules", "ccguid=" & KmaEncodeSQLText(ModuleGuid))
                    If Main.IsCSOK(CS2) Then
                        Code = Trim(Main.GetCS(CS2, "code"))
                        Code = EncodeCData(Code)
                        s = s & vbCrLf & vbTab & "<ScriptingModule Name=""" & kmaEncodeHTML(Main.GetCS(CS2, "name")) & """ guid=""" & ModuleGuid & """>" & Code & "</ScriptingModule>"
                    End If
                    Call Main.CloseCS(CS2)
                End If
            Next
        End If
        '
        ' shared styles
        '
Dim recordGuids() As String
Dim recordGuid As String
        If (IncludeSharedStyleGuidList <> "") Then
            recordGuids = Split(IncludeSharedStyleGuidList, vbCrLf)
            For Ptr = 0 To UBound(recordGuids)
                recordGuid = recordGuids(Ptr)
                If recordGuid <> "" Then
                    CS2 = Main.OpenCSContent("Shared Styles", "ccguid=" & KmaEncodeSQLText(recordGuid))
                    If Main.IsCSOK(CS2) Then
                        s = s & vbCrLf & vbTab & "<SharedStyle" _
                            & " Name=""" & kmaEncodeHTML(Main.GetCS(CS2, "name")) & """" _
                            & " guid=""" & recordGuid & """" _
                            & " alwaysInclude=""" & Main.GetCSBoolean(CS2, "alwaysInclude") & """" _
                            & " prefix=""" & kmaEncodeHTML(Main.GetCSText(CS2, "prefix")) & """" _
                            & " suffix=""" & kmaEncodeHTML(Main.GetCSText(CS2, "suffix")) & """" _
                            & " sortOrder=""" & kmaEncodeHTML(Main.GetCSText(CS2, "sortOrder")) & """" _
                            & ">" _
                            & EncodeCData(Trim(Main.GetCS(CS2, "styleFilename"))) _
                            & "</SharedStyle>"
                    End If
                    Call Main.CloseCS(CS2)
                End If
            Next
        End If
        '
        ' Import Collections
        '
        Node = ""
        CS3 = Main.OpenCSContent("Add-on Collection Parent Rules", "parentid=" & CollectionID)
        Do While Main.IsCSOK(CS3)
            CS2 = Main.OpenCSContentRecord("Add-on Collections", Main.GetCSInteger(CS3, "childid"))
            If Main.IsCSOK(CS2) Then
                Guid = Main.GetCSText(CS2, "ccGuid")
                If Guid = "" Then
                    Guid = GuidGenerator.CreateGUID("")
                    Call Main.SetCS(CS2, "ccGuid", Guid)
                End If
                Node = Node & vbCrLf & vbTab & "<ImportCollection name=""" & kmaEncodeHTML(Main.GetCSText(CS2, "name")) & """>" & Guid & "</ImportCollection>"
            End If
            Call Main.CloseCS(CS2)
            Call Main.NextCSRecord(CS3)
        Loop
        Call Main.CloseCS(CS3)
        s = s & Node
        '
        ' wwwFileList
        '
        ResourceCnt = 0
        FileList = Main.GetCSText(CS, "wwwFileList")
        If FileList <> "" Then
            PhysicalWWWPath = Main.PhysicalWWWPath
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
                        If InStr(1, AddFileList, "\" & Filename, vbTextCompare) <> 0 Then
                            Call Main.AddUserError("There was an error exporting this collection because there were multiple files with the same filename [" & Filename & "]")
                        Else
                            s = s & vbCrLf & vbTab & "<Resource name=""" & kmaEncodeHTML(Filename) & """ type=""www"" path=""" & kmaEncodeHTML(Path) & """ />"
                            AddFileList = AddFileList & vbCrLf & AddFilename
                        End If
                        ResourceCnt = ResourceCnt + 1
                    End If
                End If
            Next
        End If
        '
        ' ContentFileList
        '
        FileList = Main.GetCSText(CS, "ContentFileList")
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
                    AddFilename = Main.PhysicalFilePath & PathFilename
                    If InStr(1, AddFileList, "\" & Filename, vbTextCompare) <> 0 Then
                        Call Main.AddUserError("There was an error exporting this collection because there were multiple files with the same filename [" & Filename & "]")
                    Else
                        s = s & vbCrLf & vbTab & "<Resource name=""" & kmaEncodeHTML(Filename) & """ type=""content"" path=""" & kmaEncodeHTML(Path) & """ />"
                        AddFileList = AddFileList & vbCrLf & AddFilename
                    End If
                    ResourceCnt = ResourceCnt + 1
                End If
            Next
        End If
        '
        ' ExecFileListNode
        '
        s = s & ExecFileListNode
        '
        ' Other XML
        '
Dim OtherXML As String
        OtherXML = Main.GetCSText(CS, "otherxml")
        If Trim(OtherXML) <> "" Then
            s = s & vbCrLf & OtherXML
        End If
        's = s & GetNodeText("OtherXML", Main.GetCSText(CS, "otherxml"))
        '    & "<Data></Data>" _
        '    & "<NavigatorEntry></NavigatorEntry>" _
        '    & "<SQLIndex></SQLIndex>" _
        '    & "<Styles></Styles>" _
        '    & "<ScriptingModule Name="" guid="" />" _
        '    & "<Resource name="" type="www|content|executable" path="" />"
        s = s & vbCrLf & "</Collection>"
        '
        Call Main.CloseCS(CS)
        '
        ' Save the installation file and add it to the archive
        '
        Call fs.SaveFile(InstallFilename, s)
        If InStr(1, vbCrLf & AddFileList, vbCrLf & InstallFilename, vbTextCompare) = 0 Then
            AddFileList = AddFileList & vbCrLf & InstallFilename
        End If
        Call fs.SaveFile(AddFileListFilename, AddFileList)
        Call runAtServer("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable("@" & AddFileListFilename))
        'Set Remote = CreateObject("ccRemote.RemoteClass")
        'Call Remote.executeCmd("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable("@" & AddFileListFilename))
    End If
    '
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetCollection")
End Function
'
'
'
Private Function GetCDef(CollectionID As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim s As String
    '
    GetCDef = s
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetCDef")
End Function
'
'
'
Private Function getNavigatorNode(navigatorId As Long, Return_IncludeModuleGuidList As String, Ver40Compatibility As Boolean, Ver40DLLList As String) As String
    On Error GoTo ErrorTrap
    '
    Dim s As String
    Dim CS As Long
    Dim GuidGenerator As New GuidGenerator
    Dim nameSpace As String
    '
    CS = Main.OpenCSContentRecord("Navigator Entries", navigatorId)
    If Main.IsCSOK(CS) Then
        '
        ' ActiveX DLL node is being deprecated. This should be in the collection resource section
        '
        s = s & GetNodeText("Active", Main.GetCS(CS, "active"))
        '
        '
        '
        s = "" _
            & vbCrLf & vbTab & "<NavigatorEntry name=""" & kmaEncodeHTML(Main.GetCS(CS, "name")) & """ guid=""" & kmaEncodeHTML(Main.GetCS(CS, "guid")) & """ namespace=""" & nameSpace & """>" _
            & kmaIndent(s) _
            & vbCrLf & vbTab & "</NavigatorEntry>"
    End If
    Call Main.CloseCS(CS)
    '
    getNavigatorNode = s
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("getNavigatorNode")
End Function
'
'
'
Private Function GetAddonNode(addonid As Long, ByRef Return_IncludeModuleGuidList As String, Ver40Compatibility As Boolean, Ver40DLLList As String, ByRef Return_IncludeSharedStyleGuidList As String) As String
    On Error GoTo ErrorTrap
    '
    Dim styleId As Long
    Dim fieldType As String
    Dim fieldTypeID As Long
    Dim TriggerContentID As Long
    Dim StylesTest As String
    Dim BlockEditTools As Boolean
    Dim NavType As String
    Dim Styles As String
    Dim s As String
    Dim CS As Long
    Dim CS2 As Long
    Dim CS3 As Long
    Dim Node As String
    Dim NodeInnerText As String
    Dim GuidGenerator As New GuidGenerator
    Dim FilterAddonID As Long
    Dim IncludedAddonID As Long
    Dim ScriptingModuleID As Long
    Dim Guid As String
    Dim ListSplit() As String
    Dim Ptr As Long
    Dim Filename As String
    Dim addonName As String
    Dim processRunOnce As Boolean
    '
    CS = Main.OpenCSContentRecord("Add-ons", addonid)
    If Main.IsCSOK(CS) Then
        addonName = Main.GetCS(CS, "name")
        processRunOnce = Main.GetCSBoolean(CS, "ProcessRunOnce")
        If ((LCase(addonName) = "oninstall") Or (LCase(addonName) = "_oninstall")) Then
            processRunOnce = True
        End If
        '
        ' ActiveX DLL node is being deprecated. This should be in the collection resource section
        '
        s = s & GetNodeText("Copy", Main.GetCS(CS, "Copy"))
        s = s & GetNodeText("CopyText", Main.GetCS(CS, "CopyText"))
        '
        ' DLL
        '
        
        s = s & GetNodeText("ActiveXProgramID", Main.GetCS(CS, "objectprogramid"))
        s = s & GetNodeText("DotNetClass", Main.GetCS(CS, "DotNetClass"))
        If Ver40Compatibility And (Ver40DLLList <> "") Then
            ListSplit = Split(Ver40DLLList, vbCrLf)
            For Ptr = 0 To UBound(ListSplit)
                Filename = Trim(ListSplit(Ptr))
                If Filename <> "" Then
                    s = s & GetNodeText("ActiveXDLL", Filename) & "<!-- Version 40 Compatibility -->"
                End If
            Next
        End If
        '
        ' Features
        '
        s = s & GetNodeText("ArgumentList", Main.GetCS(CS, "ArgumentList"))
        s = s & GetNodeBoolean("AsAjax", Main.GetCSBoolean(CS, "AsAjax"))
        s = s & GetNodeBoolean("Filter", Main.GetCSBoolean(CS, "Filter"))
        s = s & GetNodeText("Help", Main.GetCS(CS, "Help"))
        s = s & GetNodeText("HelpLink", Main.GetCS(CS, "HelpLink"))
        s = s & vbCrLf & vbTab & "<Icon Link=""" & Main.GetCSText(CS, "iconfilename") & """ width=""" & Main.GetCSInteger(CS, "iconWidth") & """ height=""" & Main.GetCSInteger(CS, "iconHeight") & """ sprites=""" & Main.GetCSInteger(CS, "iconSprites") & """ />"
        s = s & GetNodeBoolean("InIframe", Main.GetCSBoolean(CS, "InFrame"))
        BlockEditTools = False
        If Main.IsCSFieldSupported(CS, "BlockEditTools") Then
            BlockEditTools = Main.GetCSBoolean(CS, "BlockEditTools")
        End If
        s = s & GetNodeBoolean("BlockEditTools", BlockEditTools)
        '
        ' Form XML
        '
        s = s & GetNodeText("FormXML", Main.GetCS(CS, "FormXML"))
        '
        NodeInnerText = ""
        CS2 = Main.OpenCSContent("Add-on Include Rules", "addonid=" & addonid)
        Do While Main.IsCSOK(CS2)
            IncludedAddonID = Main.GetCSInteger(CS2, "IncludedAddonID")
            CS3 = Main.OpenCSContent("Add-ons", "ID=" & IncludedAddonID)
            If Main.IsCSOK(CS3) Then
                Guid = Main.GetCS(CS3, "ccGuid")
                If Guid = "" Then
                    Guid = GuidGenerator.CreateGUID("")
                    Call Main.SetCS(CS3, "ccGuid", Guid)
                End If
                s = s & vbCrLf & vbTab & "<IncludeAddon name=""" & kmaEncodeHTML(Main.GetCS(CS3, "name")) & """ guid=""" & Guid & """/>"
            End If
            Call Main.CloseCS(CS3)
            Call Main.NextCSRecord(CS2)
        Loop
        Call Main.CloseCS(CS2)
        '
        s = s & GetNodeBoolean("IsInline", Main.GetCSBoolean(CS, "IsInline"))
        s = s & GetNodeText("JavascriptOnLoad", Main.GetCS(CS, "JavascriptOnLoad"))
        s = s & GetNodeText("JavascriptInHead", Main.GetCS(CS, "JSFilename"))
        s = s & GetNodeText("JavascriptBodyEnd", Main.GetCS(CS, "JavascriptBodyEnd"))
        s = s & GetNodeText("MetaDescription", Main.GetCS(CS, "MetaDescription"))
        s = s & GetNodeText("OtherHeadTags", Main.GetCS(CS, "OtherHeadTags"))
        '
        ' Placements
        '
        s = s & GetNodeBoolean("Content", Main.GetCSBoolean(CS, "Content"))
        s = s & GetNodeBoolean("Template", Main.GetCSBoolean(CS, "Template"))
        s = s & GetNodeBoolean("Email", Main.GetCSBoolean(CS, "Email"))
        s = s & GetNodeBoolean("Admin", Main.GetCSBoolean(CS, "Admin"))
        s = s & GetNodeBoolean("OnPageEndEvent", Main.GetCSBoolean(CS, "OnPageEndEvent"))
        s = s & GetNodeBoolean("OnPageStartEvent", Main.GetCSBoolean(CS, "OnPageStartEvent"))
        s = s & GetNodeBoolean("OnBodyStart", Main.GetCSBoolean(CS, "OnBodyStart"))
        s = s & GetNodeBoolean("OnBodyEnd", Main.GetCSBoolean(CS, "OnBodyEnd"))
        s = s & GetNodeBoolean("RemoteMethod", Main.GetCSBoolean(CS, "RemoteMethod"))
        's = s & GetNodeBoolean("OnNewVisitEvent", Main.GetCSBoolean(CS, "OnNewVisitEvent"))
        '
        ' Process
        '
        s = s & GetNodeBoolean("ProcessRunOnce", processRunOnce)
        s = s & GetNodeInteger("ProcessInterval", Main.GetCSInteger(CS, "ProcessInterval"))
        '
        ' Meta
        '
        s = s & GetNodeText("PageTitle", Main.GetCS(CS, "PageTitle"))
        s = s & GetNodeText("RemoteAssetLink", Main.GetCS(CS, "RemoteAssetLink"))
        '
        ' Styles
        '
        Styles = ""
        If Not Main.GetCSBoolean(CS, "BlockDefaultStyles") Then
            Styles = Trim(Main.GetCS(CS, "StylesFilename"))
        End If
        StylesTest = Trim(Main.GetCS(CS, "CustomStylesFilename"))
        If StylesTest <> "" Then
            If Styles <> "" Then
                Styles = Styles & vbCrLf & StylesTest
            Else
                Styles = StylesTest
            End If
        End If
        s = s & GetNodeText("Styles", Styles)
        '
        ' Scripting
        '
        NodeInnerText = Trim(Main.GetCS(CS, "ScriptingCode"))
        If NodeInnerText <> "" Then
            NodeInnerText = vbCrLf & vbTab & vbTab & "<Code>" & EncodeCData(NodeInnerText) & "</Code>"
        End If
        CS2 = Main.OpenCSContent("Add-on Scripting Module Rules", "addonid=" & addonid)
        Do While Main.IsCSOK(CS2)
            ScriptingModuleID = Main.GetCSInteger(CS2, "ScriptingModuleID")
            CS3 = Main.OpenCSContent("Scripting Modules", "ID=" & ScriptingModuleID)
            If Main.IsCSOK(CS3) Then
                Guid = Main.GetCS(CS3, "ccGuid")
                If Guid = "" Then
                    Guid = GuidGenerator.CreateGUID("")
                    Call Main.SetCS(CS3, "ccGuid", Guid)
                End If
                Return_IncludeModuleGuidList = Return_IncludeModuleGuidList & vbCrLf & Guid
                NodeInnerText = NodeInnerText & vbCrLf & vbTab & vbTab & "<IncludeModule name=""" & kmaEncodeHTML(Main.GetCS(CS3, "name")) & """ guid=""" & Guid & """/>"
            End If
            Call Main.CloseCS(CS3)
            Call Main.NextCSRecord(CS2)
        Loop
        Call Main.CloseCS(CS2)
        If NodeInnerText = "" Then
            s = s & vbCrLf & vbTab & "<Scripting Language=""" & Main.GetCSLookup(CS, "ScriptingLanguageID") & """ EntryPoint=""" & Main.GetCS(CS, "ScriptingEntryPoint") & """ Timeout=""" & Main.GetCS(CS, "ScriptingTimeout") & """/>"
        Else
            s = s & vbCrLf & vbTab & "<Scripting Language=""" & Main.GetCSLookup(CS, "ScriptingLanguageID") & """ EntryPoint=""" & Main.GetCS(CS, "ScriptingEntryPoint") & """ Timeout=""" & Main.GetCS(CS, "ScriptingTimeout") & """>" & NodeInnerText & vbCrLf & vbTab & "</Scripting>"
        End If
        '
        ' Shared Styles
        '
        CS2 = Main.OpenCSContent("Shared Styles Add-on Rules", "addonid=" & addonid)
        Do While Main.IsCSOK(CS2)
            styleId = Main.GetCSInteger(CS2, "styleId")
            CS3 = Main.OpenCSContent("shared styles", "ID=" & styleId)
            If Main.IsCSOK(CS3) Then
                Guid = Main.GetCS(CS3, "ccGuid")
                If Guid = "" Then
                    Guid = GuidGenerator.CreateGUID("")
                    Call Main.SetCS(CS3, "ccGuid", Guid)
                End If
                Return_IncludeSharedStyleGuidList = Return_IncludeSharedStyleGuidList & vbCrLf & Guid
                s = s & vbCrLf & vbTab & "<IncludeSharedStyle name=""" & kmaEncodeHTML(Main.GetCS(CS3, "name")) & """ guid=""" & Guid & """/>"
            End If
            Call Main.CloseCS(CS3)
            Call Main.NextCSRecord(CS2)
        Loop
        Call Main.CloseCS(CS2)
        '
        ' Process Triggers
        '
        NodeInnerText = ""
        CS2 = Main.OpenCSContent("Add-on Content Trigger Rules", "addonid=" & addonid)
        Do While Main.IsCSOK(CS2)
            TriggerContentID = Main.GetCSInteger(CS2, "ContentID")
            CS3 = Main.OpenCSContent("content", "ID=" & TriggerContentID)
            If Main.IsCSOK(CS3) Then
                Guid = Main.GetCS(CS3, "ccGuid")
                If Guid = "" Then
                    Guid = GuidGenerator.CreateGUID("")
                    Call Main.SetCS(CS3, "ccGuid", Guid)
                End If
                NodeInnerText = NodeInnerText & vbCrLf & vbTab & vbTab & "<ContentChange name=""" & kmaEncodeHTML(Main.GetCS(CS3, "name")) & """ guid=""" & Guid & """/>"
            End If
            Call Main.CloseCS(CS3)
            Call Main.NextCSRecord(CS2)
        Loop
        Call Main.CloseCS(CS2)
        If NodeInnerText <> "" Then
            s = s & vbCrLf & vbTab & "<ProcessTriggers>" & NodeInnerText & vbCrLf & vbTab & "</ProcessTriggers>"
        End If
        '
        ' Editors
        '
        If Main.IsContentFieldSupported("Add-on Content Field Type Rules", "id") Then
            NodeInnerText = ""
            CS2 = Main.OpenCSContent("Add-on Content Field Type Rules", "addonid=" & addonid)
            Do While Main.IsCSOK(CS2)
                fieldTypeID = Main.GetCSInteger(CS2, "contentFieldTypeID")
                fieldType = Main.GetRecordName("Content Field Types", fieldTypeID)
                If fieldType <> "" Then
                    NodeInnerText = NodeInnerText & vbCrLf & vbTab & vbTab & "<type>" & fieldType & "</type>"
                End If
                Call Main.NextCSRecord(CS2)
            Loop
            Call Main.CloseCS(CS2)
            If NodeInnerText <> "" Then
                s = s & vbCrLf & vbTab & "<Editors>" & NodeInnerText & vbCrLf & vbTab & "</Editors>"
            End If
        End If
        '
        '
        '
        Guid = Main.GetCS(CS, "ccGuid")
        If Guid = "" Then
            Guid = GuidGenerator.CreateGUID("")
            Call Main.SetCS(CS, "ccGuid", Guid)
        End If
        NavType = Main.GetCS(CS, "NavTypeID")
        If NavType = "" Then
            NavType = "Add-on"
        End If
    If Ver40Compatibility Then
        s = "" _
            & vbCrLf & vbTab & "<Page name=""" & kmaEncodeHTML(addonName) & """ guid=""" & Guid & """ type=""" & NavType & """>" _
            & kmaIndent(s) _
            & vbCrLf & vbTab & "</Page>"
        s = "" _
            & vbCrLf & vbTab & "<Interfaces><!-- Version 4.0 Compatibility -->" _
            & kmaIndent(s) _
            & vbCrLf & vbTab & "</Interfaces>"
    Else
        s = "" _
            & vbCrLf & vbTab & "<Addon name=""" & kmaEncodeHTML(addonName) & """ guid=""" & Guid & """ type=""" & NavType & """>" _
            & kmaIndent(s) _
            & vbCrLf & vbTab & "</Addon>"
    End If
    End If
    Call Main.CloseCS(CS)
    '
    GetAddonNode = s
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetAddonNode")
End Function
''
''
''
'Private Function GetProcessInterfaceNode(AddonID As Long, Return_IncludeModuleGuidList As String) As String
'    On Error GoTo ErrorTrap
'    '
'    Dim s As String
'    Dim CS As Long
'    Dim CS2 As Long
'    Dim CS3 As Long
'    Dim Node As String
'    Dim GuidGenerator As New GuidGenerator
'    Dim FilterAddonID As Long
'    Dim IncludedAddonID As Long
'    Dim ScriptingModuleID As Long
'    Dim Guid As String
'    '
'    CS = Main.OpenCSContentRecord("Add-ons", AddonID)
'    If Main.iscsok(CS) Then
'        '
'        ' ActiveX DLL node is being deprecated. This should be in the collection resource section
'        '
'        's = s & GetNodeText("Description", Main.GetCS(CS, "Description"))
'        s = s & GetNodeText("Help", Main.GetCS(CS, "Help"))
'        s = s & GetNodeText("HelpLink", Main.GetCS(CS, "HelpLink"))
'        s = s & GetNodeText("ArgumentList", Main.GetCS(CS, "ArgumentList"))
'        s = s & GetNodeText("ActiveXProgramID", Main.GetCS(CS, "objectprogramid"))
'        s = s & GetNodeBoolean("ProcessRunOnce", Main.GetCSBoolean(CS, "ProcessRunOnce"))
'        s = s & GetNodeInteger("ProcessInterval", Main.getcsinteger(CS, "ProcessInterval"))
'        '
'        ' Add scripting modules
'        '
'        Node = ""
'        CS2 = Main.OpenCSContent("Add-on Scripting Module Rules", "addonid=" & AddonID)
'        Do While Main.iscsok(CS2)
'            ScriptingModuleID = Main.getcsinteger(CS2, "ScriptingModuleID")
'            CS3 = Main.OpenCSContent("Scripting Modules", "ID=" & ScriptingModuleID)
'            If Main.iscsok(CS3) Then
'                Guid = Main.GetCS(CS3, "ccGuid")
'                If Guid = "" Then
'                    Guid = GuidGenerator.CreateGUID("")
'                    Call Main.SetCS(CS3, "ccGuid", Guid)
'                End If
'                Return_IncludeModuleGuidList = Return_IncludeModuleGuidList & vbCrLf & Guid
'                Node = Node & vbCrLf & vbTab & "<IncludeModule name=""" & Main.GetCS(CS3, "name") & """ guid=""" & Guid & """/>"
'            End If
'            Call Main.CloseCS(CS3)
'            Call Main.nextcsrecord(CS2)
'        Loop
'        Call Main.CloseCS(CS2)
'        If Node = "" Then
'            s = s & vbCrLf & vbTab & "<Scripting Language=""" & Main.GetCSLookup(CS, "ScriptingLanguageID") & """ EntryPoint=""" & Main.GetCS(CS, "ScriptingEntryPoint") & """/>"
'        Else
'            s = s & vbCrLf & vbTab & "<Scripting Language=""" & Main.GetCSLookup(CS, "ScriptingLanguageID") & """ EntryPoint=""" & Main.GetCS(CS, "ScriptingEntryPoint") & """>" & Node & vbCrLf & vbTab & "</Scripting>"
'        End If
'        '
'        Guid = Main.GetCS(CS, "ccGuid")
'        If Guid = "" Then
'            Guid = GuidGenerator.CreateGUID("")
'            Call Main.SetCS(CS, "ccGuid", Guid)
'        End If
'        s = "" _
'            & vbCrLf & vbTab & "<Process name=""" & addonName & """ guid=""" & Guid & """>" _
'            & KmaIndent(s) _
'            & vbCrLf & vbTab & "</Process>"
'    End If
'    Call Main.CloseCS(CS)
'    '
'    GetProcessInterfaceNode = s
'    '
'    Exit Function
'ErrorTrap:
'    Call HandleClassTrapError("GetProcessInterfaceNode")
'End Function
'
'
'
Private Function GetNodeText(NodeName As String, NodeContent As String) As String
    On Error GoTo ErrorTrap
    '
    If NodeContent = "" Then
        GetNodeText = GetNodeText & vbCrLf & vbTab & "<" & NodeName & "></" & NodeName & ">"
    Else
        GetNodeText = GetNodeText & vbCrLf & vbTab & "<" & NodeName & ">" & EncodeCData(NodeContent) & "</" & NodeName & ">"
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetNodeText")
End Function
'
'
'
Private Function GetNodeBoolean(NodeName As String, NodeContent As Boolean) As String
    On Error GoTo ErrorTrap
    '
    GetNodeBoolean = GetNodeBoolean & vbCrLf & vbTab & "<" & NodeName & ">" & kmaGetYesNo(NodeContent) & "</" & NodeName & ">"
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetNodeBoolean")
End Function
'
'
'
Private Function GetNodeInteger(NodeName As String, NodeContent As Integer) As String
    On Error GoTo ErrorTrap
    '
    GetNodeInteger = GetNodeInteger & vbCrLf & vbTab & "<" & NodeName & ">" & CStr(NodeContent) & "</" & NodeName & ">"
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetNodeInteger")
End Function
'
'
'
Private Sub HandleClassTrapError(MethodName As String)
    Call HandleError("CExportClass", MethodName, Err.Number, Err.Source, Err.Description, True, False, "")
End Sub
'
'========================================================================
' From ccCommon
'========================================================================
'
Function ReplaceMany(Source As String, ArrayOfSource() As Variant, ArrayOfReplacement() As Variant) As String
    ' ##### removed to catch err<>0 problem on error resume next
    '
    Dim Count As Long
    Dim Pointer As Long
    '
    Count = UBound(ArrayOfSource) + 1
    ReplaceMany = Source
    For Pointer = 0 To Count - 1
        ReplaceMany = Replace(ReplaceMany, ArrayOfSource(Pointer), ArrayOfReplacement(Pointer))
        Next
    '
End Function
'
'========================================================================
' From ccCommon
'========================================================================
'
Function encodeFilename(Filename As String) As String
    ' ##### removed to catch err<>0 problem on error resume next
    '
    Dim Source() As Variant
    Dim Replacement() As Variant
    '
    Source = Array("""", "*", "/", ":", "<", ">", "?", "\", "|")
    Replacement = Array("_", "_", "_", "_", "_", "_", "_", "_", "_")
    '
    encodeFilename = ReplaceMany(Filename, Source, Replacement)
    If Len(encodeFilename) > 254 Then
        encodeFilename = Left(encodeFilename, 254)
    End If
    '
End Function
'
'
'
Friend Sub GetLocalCollectionArgs(CollectionGuid As String, Return_CollectionPath As String, Return_LastChagnedate As Date)
    On Error GoTo ErrorTrap
    '
    Const CollectionListRootNode = "collectionlist"
    '
    Dim LocalPath As String
    Dim LocalFilename As String
    Dim LocalGuid As String
    Dim Doc As New DOMDocument
    Dim CollectionNode As IXMLDOMNode
    Dim LocalListNode As IXMLDOMNode
    Dim NewCollectionNode As IXMLDOMNode
    Dim NewAttrNode As IXMLDOMNode
    Dim CollectionFound As Boolean
    Dim Ptr As Long
    Dim CollectionPath As String
    Dim LastChangeDate As Date
    Dim Hint As String
    Dim MatchFound As Boolean
    Dim LocalName As String
    '
    MatchFound = False
    Return_CollectionPath = ""
    Return_LastChagnedate = CDate(0)
    Hint = "Match guid [" & CollectionGuid & "], Loading"
    Call Doc.loadXML(GetConfig)
    Do While Doc.readyState <> 4 And Ptr < 100
        Hint = Hint & ",waiting for load"
        Sleep (100)
        DoEvents
        Ptr = Ptr + 1
    Loop
    If Doc.parseError.errorCode <> 0 Then
        Hint = Hint & ",parse error"
        'Call AppendClassLogFile("Server", "", "GetLocalCollectionArgs, Hint=[" & Hint & "], Error loading Collections.xml file. The error was [" & Doc.parseError.reason & ", line " & Doc.parseError.Line & ", character " & Doc.parseError.linepos & "]")
    Else
        If LCase(Doc.documentElement.baseName) <> LCase(CollectionListRootNode) Then
            'Call AppendClassLogFile("Server", "", "GetLocalCollectionArgs, Hint=[" & Hint & "], The Collections.xml file has an invalid root node, [" & Doc.documentElement.BaseName & "] was received and [" & CollectionListRootNode & "] was expected.")
        Else
            With Doc.documentElement
                If LCase(.baseName) <> "collectionlist" Then
                    'Call AppendClassLogFile("Server", "", "GetLocalCollectionArgs, basename was not collectionlist, [" & .BaseName & "].")
                Else
                    CollectionFound = False
                    Hint = Hint & ",checking nodes [" & .childNodes.length & "]"
                    For Each LocalListNode In .childNodes
                        LocalName = "no name found"
                        LocalPath = ""
                        Select Case LCase(LocalListNode.baseName)
                            Case "collection"
                                LocalGuid = ""
                                For Each CollectionNode In LocalListNode.childNodes
                                    Select Case LCase(CollectionNode.baseName)
                                        Case "name"
                                            '
                                            LocalName = LCase(CollectionNode.Text)
                                        Case "guid"
                                            '
                                            LocalGuid = LCase(CollectionNode.Text)
                                        Case "path"
                                            '
                                            CollectionPath = LCase(CollectionNode.Text)
                                        Case "lastchangedate"
                                            LastChangeDate = KmaEncodeDate(CollectionNode.Text)
                                    End Select
                                Next
                        End Select
                        Hint = Hint & ",checking node [" & LocalName & "]"
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
    '
    Exit Sub
ErrorTrap:
    Call HandleClassTrapError("GetLocalCollectionArgs")
End Sub
'
'
'
Public Function GetConfig() As String
    On Error GoTo ErrorTrap
    '
    Dim LocalFilename As String
    Dim AddonPath As String
    Dim Pos As Long
    '
    AddonPath = Replace(Main.PhysicalccLibPath(), "\cclib", "\addons", , , vbTextCompare)
    AddonPath = AddonPath & "\Collections.xml"
    GetConfig = Main.ReadFile(AddonPath)
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetConfig")
End Function
'
'
'
Private Function AddCompatibilityResources(CollectionPath As String, ArchiveFilename As String, SubPath As String, Return_Version40DLLList As String) As String
    On Error GoTo ErrorTrap
    '
    Dim AddFilename As String
    Dim FileExt As String
    Dim FileList As String
    Dim Files() As String
    Dim Filename As String
    Dim Ptr As Long
    Dim FileArgs() As String
    Dim s As String
    'Dim Remote As Object
    Dim FolderList As String
    Dim Folders() As String
    Dim FolderArgs() As String
    Dim Folder As String
    Dim Pos As Long
    '
    ' Process all SubPaths
    '
    FolderList = Main.GetFolderList(CollectionPath & SubPath)
    If FolderList <> "" Then
        Folders = Split(FolderList, vbCrLf)
        For Ptr = 0 To UBound(Folders)
            Folder = Folders(Ptr)
            If Folder <> "" Then
            FolderArgs = Split(Folders(Ptr), ",")
            Folder = FolderArgs(0)
            If Folder <> "" Then
                s = s & AddCompatibilityResources(CollectionPath, ArchiveFilename, SubPath & Folder & "\", Return_Version40DLLList)
            End If
            End If
        Next
    End If
    '
    ' Process files in this path
    '
    'Set Remote = CreateObject("ccRemote.RemoteClass")
    FileList = Main.GetFileList(CollectionPath)
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
                        s = s & vbCrLf & vbTab & "<Resource name=""" & kmaEncodeHTML(Filename) & """ type=""content"" path=""" & kmaEncodeHTML(SubPath) & """ />"
                        AddFilename = CollectionPath & SubPath & "\" & Filename
                        Call runAtServer("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable(AddFilename))
                        'Call Remote.executeCmd("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable(AddFilename))
                    ElseIf LCase(FileExt) = "dll" Then
                        '
                        ' Executable resources
                        '
                        s = s & vbCrLf & vbTab & "<Resource name=""" & kmaEncodeHTML(Filename) & """ type=""executable"" path=""" & kmaEncodeHTML(SubPath) & """ />"
                        Return_Version40DLLList = Return_Version40DLLList & vbCrLf & Filename
                        AddFilename = CollectionPath & SubPath & "\" & Filename
                        Call runAtServer("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable(AddFilename))
                        'Call Remote.executeCmd("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable(AddFilename))
                    Else
                        '
                        ' www resources
                        '
                        s = s & vbCrLf & vbTab & "<Resource name=""" & kmaEncodeHTML(Filename) & """ type=""www"" path=""" & kmaEncodeHTML(SubPath) & """ />"
                        AddFilename = CollectionPath & SubPath & "\" & Filename
                        Call runAtServer("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable(AddFilename))
                        'Call Remote.executeCmd("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable(AddFilename))
                    End If
                End If
            End If
        Next
    End If
    '
    AddCompatibilityResources = s
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("AddCompatibilityResources")
End Function
'
' Encode a string into CData format
'   <![CDATA[" & NodeContent & "]]>
'
Private Function EncodeCData(Source As String) As String
    On Error GoTo ErrorTrap
    '
    EncodeCData = Source
    If EncodeCData <> "" Then
        EncodeCData = "<![CDATA[" & Replace(EncodeCData, "]]>", "]]]]><![CDATA[>") & "]]>"
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("EncodeCData")
End Function
'
'
'
Private Sub runAtServer(cmd As String, arg As String)
    On Error GoTo ErrorTrap
    '
    Dim runAtServer As Object
    '
    On Error Resume Next
    Set runAtServer = CreateObject("contensive.runAtServerClass")
    If Err.Number <> 0 Then
        '
        ' old object
        '
        Set runAtServer = CreateObject("ccRemote.RemoteClass")
    End If
    Call runAtServer.executeCmd(cmd, arg)
    'Call Remote.executeCmd("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable("@" & AddFileListFilename))
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassTrapError("runAtServer")
End Sub
