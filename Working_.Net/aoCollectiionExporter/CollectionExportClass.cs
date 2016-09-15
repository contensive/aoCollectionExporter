
using System;
using System.Collections.Generic;
using System.Text;
using Contensive.BaseClasses;

namespace Contensive.Addons
{
    //
    // 2) Change this class name to the addon name
    // 3) Create a Contensive Addon record with the namespace apCollectionName.ad
    // 3) add reference to CPBase.DLL, typically installed in c:\program files\kma\contensive\
    //
    public class CollectionExportClass : Contensive.BaseClasses.AddonBaseClass
    {
        const int FormIDSelectCollection = 0;
        const int FormIDDisplayResults = 1;
        const string RequestNameButton = "button";
        const string RequestNameFormID = "formid";
        const string RequestnameExecutableFile = "executablefile";
        const string RequestNameCollectionID = "collectionid";
        //
        //====================================================================================================
        /// <summary>
        /// addon interface
        /// </summary>
        /// <param name="cp"></param>
        /// <returns></returns>
        public override object Execute(Contensive.BaseClasses.CPBaseClass cp)
        {
            string returnHtml = "";
            try
            {
                returnHtml = exporter( cp );
            }
            catch (Exception ex)
            {
                cp.Site.ErrorReport("There was a unexpected exception, " + ex.ToString());
                returnHtml = "";
            }
            return returnHtml;
        }
        //====================================================================================================
        /// <summary>
        /// Create and process the export form
        /// </summary>
        /// <param name="cp"></param>
        /// <returns></returns>
        public string exporter(Contensive.BaseClasses.CPBaseClass cp)
        {
            string returnString = "";
            try
            {
                //
                string AddExecutableFilename = "";
                int DoSomethingPtr = 0;
                string Button = "";
                int FormID = 0;
                int CollectionID = 0;
                string CollectionName = "";
                string CollectionFile = "";
                string CollectionFilename = "";
                string s = "";
                //
                // Every form returns a button and a formid
                //
                Button = cp.Doc.GetText( "button" );
                FormID = cp.Doc.GetInteger( "formid" );
                //
                // Process the current form submission
                //
                if (Button != "")
                {
                    switch (FormID)
                    {
                        case FormIDDisplayResults:
                            // 
                            // nothing to process 
                            // 
                            break;
                        default:
                            // 
                            // process the Select Collection Form button 
                            // 
                            //hint = hint + ",200";
                            CollectionID = cp.Doc.GetInteger( RequestNameCollectionID );
                            CollectionName = cp.Content.GetRecordName( "Add-on Collections", CollectionID );
                            if (CollectionName == "")
                            {
                                cp.UserError.Add( "The collection file you selected could not be found. Please select another." );
                            }
                            else
                            {
                                AddExecutableFilename = cp.Doc.GetText ( RequestnameExecutableFile );
                                if (AddExecutableFilename != "")
                                {
                                    AddExecutableFilename = "CollectionExport\\" + AddExecutableFilename;
                                    cp.File.save
                                    ReflectionHelper.Invoke(Main, "SaveStreamFile", new object[] { RequestnameExecutableFile, "CollectionExport\\");
                                }
                                CollectionFilename = GetCollection(CollectionID, AddExecutableFilename);
                            }
                            if (ReflectionHelper.Invoke<bool>(Main, "IsUserError", new object[] { }))
                            {
                                FormID = FormIDSelectCollection;
                            }
                            else
                            {

                                FormID = FormIDDisplayResults;
                            }
                            break;
                    }
                }
                //
                // Reply with the next form
                //
                switch (FormID)
                {
                    case FormIDDisplayResults:
                        // 
                        // Diplay the results page 
                        // 
                        //hint = hint + ",500";
                        s = cp.UserError.GetList()  + "\r\n\t" + "<div class=\"responseForm\">\n\r\t" + "\t" + "<p>Click <a href=\"" + cp.Site.FilePath + CollectionFilename.Replace("\\", "/") + "\">here</a> to download the collection file</p>\n\r\t" + "</div>";
                        break;
                    default:
                        // 
                        // ask them to select a collectioin to export 
                        // 
                        s = cp.File.Read( cp.Site.PhysicalWWWPath + "aoCollectionExport\\layout.html" );
                        break;
                }
                //
                return "<div class=\"collectionExport\">" + s + "</div>";
            }
            catch( Exception ex )
            {
                cp.Site.ErrorReport(ex, "execute" );
            }
            return returnString;
        }
        //
        //
        //
        private string GetCollection( CPBaseClass cp,  int CollectionID, string AddExecutableFilename)
        {
            string result = "";
            try
            {
                CPCSBaseClass cs = cp.CSNew();
                CPCSBaseClass cs2 = cp.CSNew();
                CPCSBaseClass cs3 = cp.CSNew();
                CPCSBaseClass cs4 = cp.CSNew();
                //
                string IncludeSharedStyleGuidList = "";
                bool isUpdatable = false;
                string fieldLookupListValue = "";
                int CSlookup = 0;
                int FieldValueInteger = 0;
                string FieldLookupContentName = "";
                int fieldCnt = 0;
                string[] fieldNames = null;
                int[] fieldTypes = null;
                string[] fieldLookupContent = null;
                string[] fieldLookupList = null;
                int FieldLookupContentID = 0;
                string Criteria = "";
                bool supportsGuid = false;
                bool reload = false;
                int ContentID = 0;
                string FieldValue = "";
                bool FieldSkip = false;
                int FieldTypeNumber = 0;
                int CSData = 0;
                string DataRecordList = "";
                string[] DataRecords = null;
                string DataRecord = "";
                string[] DataSplit = null;
                string DataContentName = "";
                int DataContentId = 0;
                string DataRecordGuid = "";
                string DataRecordName = "";
                string TestString = "";
                string FieldName = "";
                string FieldNodes = "";
                string RecordNodes = "";
                string[] Modules = null;
                string ModuleGuid = "";
                string Code = "";
                string ManualFilename = "";
                string[] FileArgs = null;
                int ResourceCnt = 0;
                object Remote = null;
                string ContentName = "";
                string FileList = "";
                string[] Files = null;
                string PathFilename = "";
                string Filename = "";
                string Path = "";
                int Pos = 0;
                string s = "";
                StringBuilder Node = new StringBuilder();
                string CollectionGuid = "";
                string Guid = "";
                string ArchiveFilename = "";
                string ArchivePath = "";
                string InstallFilename = "";
                string CollectionName = "";
                string AddFilename = "";
                string PhysicalWWWPath = "";
                string CollectionPath = "";
                System.DateTime LastChangeDate = DateTime.FromOADate(0);
                string AddonPath = "";
                string AddFileList = "";
                string AddFileListFilename = "";
                string IncludeModuleGuidList = "";
                string Version40DLLList = "";
                StringBuilder ExecFileListNode = new StringBuilder();
                bool blockNavigatorNode = false;
                string[] recordGuids = null;
                string recordGuid = "";
                string OtherXML = "";
                //
                IncludeSharedStyleGuidList = "";
                //
                cs.OpenRecord( "Add-on Collections", CollectionID );
                if (cs.OK())
                {
                    cp.UserError.Add( "The collection you selected could not be found" );
                }
                else
                {
                    CollectionGuid = cs.GetText(  "ccGuid");
                    if (CollectionGuid == "")
                    {
                        string tempRefParam = "";
                        CollectionGuid = cp.Utils.CreateGuid();
                        cs.SetField(  "ccGuid", CollectionGuid );
                    }
                    CollectionName = cs.GetText( "name");
                    if (cs.FieldOK( "updatable" ))
                    {
                        isUpdatable = true;
                    }
                    else
                    {
                        isUpdatable = cs.GetBoolean( "updatable");
                    }
                    if (cs.FieldOK( "blockNavigatorNode" ))
                    {
                        blockNavigatorNode = false;
                    }
                    else
                    {
                        blockNavigatorNode = cs.GetBoolean( "blockNavigatorNode");
                    }
                    s = "" + "<?xml version=\"1.0\" encoding=\"windows-1252\"?>\n\r<Collection name=\"" + cp.Utils.EncodeHTML(CollectionName) + "\" guid=\"" + CollectionGuid + "\" system=\"" +  getYesNo(cs.GetBoolean( "system" )) + "\" updatable=\"" + getYesNo(isUpdatable) + "\" blockNavigatorNode=\"" + getYesNo(blockNavigatorNode) + "\">";
                    //
                    // Archive Filenames
                    //
                    ArchivePath = cp.Site.PhysicalFilePath + "CollectionExport\\";
                    InstallFilename = CollectionName + ".xml";
                    InstallFilename = ArchivePath + InstallFilename;
                    ArchiveFilename = encodeFilename(CollectionName + ".zip");
                    ArchiveFilename = ArchivePath + ArchiveFilename;
                    AddFileListFilename = ArchivePath + "AddFileList.txt";
                    result = "CollectionExport\\" + encodeFilename(CollectionName + ".zip");
                    //
                    // Delete old archive file
                    //
                    kmafs.DeleteFile(ref ArchiveFilename);
                    //
                    // Manual Upload Executable
                    // Do not add node yet, just get the file and make the node so executables can be added for V4 compatibility
                    //
                    Version40DLLList = "";
                    if (AddExecutableFilename != "")
                    {
                        //
                        // Add the uploaded executable
                        //
                        PathFilename = Strings.Replace(AddExecutableFilename, "\\", "/", 1, -1, CompareMethod.Binary);
                        Path = "";
                        ManualFilename = PathFilename;
                        Pos = ManualFilename.LastIndexOf("/") + 1;
                        if (Pos > 0)
                        {
                            ManualFilename = ManualFilename.Substring(Pos);
                        }
                        AddFilename = cp.Site.PhysicalFilePath + AddExecutableFilename;
                        if ((AddFileList.IndexOf("\\" + ManualFilename, StringComparison.CurrentCultureIgnoreCase) + 1) != 0)
                        {
                            cp.UserError.Add( "There was an error exporting this collection because there were multiple files with the same filename [" + AddExecutableFilename + "]");
                        }
                        else
                        {
                            ExecFileListNode.Append("\r\n\t" + "<Resource name=\"" + cp.Utils.EncodeHTML(ManualFilename) + "\" type=\"executable\" path=\"" + cp.Utils.EncodeHTML(Path) + "\" />");
                            AddFileList = AddFileList + Environment.NewLine + AddFilename;
                        }
                        Version40DLLList = Version40DLLList + Environment.NewLine + ManualFilename;
                    }


                    //
                    AddonPath = Strings.Replace(ReflectionHelper.Invoke<string>(Main, "PhysicalccLibPath", new object[] { }), "\\cclib", "\\addons\\", 1, -1, CompareMethod.Text);
                    FileList = cs.GetText( "execFileList");
                    if (FileList != "")
                    {
                        //
                        // There are executable files to include in the collection
                        //   If installed, source path is collectionpath, if not installed, collectionpath will be empty
                        //   and file will be sourced right from addon path
                        //
                        GetLocalCollectionArgs(CollectionGuid, ref CollectionPath, ref LastChangeDate);
                        if (CollectionPath != "")
                        {
                            CollectionPath = CollectionPath + "\\";
                        }
                        Files = (string[])FileList.Split(Environment.NewLine[0]);
                        for (int Ptr = 0; Ptr <= Files.GetUpperBound(0); Ptr++)
                        {
                            PathFilename = Files[Ptr];
                            if (PathFilename != "")
                            {
                                PathFilename = Strings.Replace(PathFilename, "\\", "/", 1, -1, CompareMethod.Binary);
                                Path = "";
                                Filename = PathFilename;
                                Pos = PathFilename.LastIndexOf("/") + 1;
                                if (Pos > 0)
                                {
                                    Filename = PathFilename.Substring(Pos);
                                    Path = PathFilename.Substring(0, Math.Min(Pos - 1, PathFilename.Length));
                                }
                                if (Filename.ToLower() != ManualFilename.ToLower())
                                {
                                    AddFilename = AddonPath + CollectionPath + Filename;
                                    if ((AddFileList.IndexOf("\\" + Filename, StringComparison.CurrentCultureIgnoreCase) + 1) != 0)
                                    {
                                        cp.UserError.Add( "There was an error exporting this collection because there were multiple files with the same filename [" + Filename + "]");
                                    }
                                    else
                                    {
                                        ExecFileListNode.Append("\r\n\t" + "<Resource name=\"" + cp.Utils.EncodeHTML(Filename) + "\" type=\"executable\" path=\"" + cp.Utils.EncodeHTML(Path) + "\" />");
                                        AddFileList = AddFileList + Environment.NewLine + AddFilename;
                                    }
                                    Version40DLLList = Version40DLLList + Environment.NewLine + Filename;
                                }
                                ResourceCnt++;
                            }
                        }
                    }
                    //ReflectionHelper.Invoke(Main, "testpoint", new object[] { "getCollection, 500");
                    if ((ResourceCnt == 0) && (CollectionPath != ""))
                    {
                        //
                        // If no resources were in the collection record, this might be an old installation
                        // Add all .dll files in the CollectionPath
                        //
                        string tempRefParam2 = AddonPath + CollectionPath;
                        ExecFileListNode.Append(AddCompatibilityResources(ref tempRefParam2, ArchiveFilename, "", ref Version40DLLList));
                    }
                    //
                    // helpLink
                    //
                    if ( cs.FieldOK("HelpLink" ))
                    {
                        s = s + "\r\n\t" + "<HelpLink>" + cp.Utils.EncodeHTML(cs.GetText( "HelpLink" )) + "</HelpLink>";
                    }
                    //
                    // Help
                    //
                    s = s + "\r\n\t" + "<Help>" + cp.Utils.EncodeHTML(cs.GetText( "Help" )) + "</Help>";
                    //
                    // Addons
                    //
                    cs2.open( "Add-ons", "collectionid=" + CollectionID.ToString(), null, null, "Process");

                    while (ReflectionHelper.Invoke<bool>(Main, "IsCSOK", new object[] { CS2 }))
                    {
                        s = s + GetAddonNode(ReflectionHelper.Invoke<int>(Main, "GetCSInteger", new object[] { CS2, "id" }), ref IncludeModuleGuidList, false , Version40DLLList, ref IncludeSharedStyleGuidList);
                        ReflectionHelper.Invoke(Main, "NextCSRecord", new object[] { CS2 });
                    };
                    //
                    // Data Records
                    //
                    //ReflectionHelper.Invoke(Main, "testpoint", new object[] { "getCollection, 600");
                    DataRecordList = cs.GetText( "DataRecordList");
                    if (DataRecordList != "")
                    {
                        DataRecords = (string[])DataRecordList.Split(Environment.NewLine[0]);
                        RecordNodes = "";
                        foreach (string DataRecords_item in DataRecords)
                        {
                            FieldNodes = "";
                            DataRecordName = "";
                            DataRecordGuid = "";
                            DataRecord = DataRecords_item;
                            if (DataRecord != "")
                            {
                                DataSplit = (string[])DataRecord.Split(',');
                                if (DataSplit.GetUpperBound(0) >= 0)
                                {
                                    DataContentName = DataSplit[0].Trim();
                                    DataContentId = ReflectionHelper.Invoke<int>(Main, "GetContentID", new object[] { DataContentName });
                                    if (DataContentId <= 0)
                                    {
                                        RecordNodes = "" + RecordNodes + "\r\n\t" + "<!-- data missing, content not found during export, content=\"" + DataContentName + "\" guid=\"" + DataRecordGuid + "\" name=\"" + DataRecordName + "\" -->";
                                    }
                                    else
                                    {
                                        supportsGuid = ReflectionHelper.Invoke<bool>(Main, "IsContentFieldSupported", new object[] { DataContentName, "ccguid");
                                        if (DataSplit.GetUpperBound(0) == 0)
                                        {
                                            Criteria = "";
                                        }
                                        else
                                        {
                                            TestString = DataSplit[1].Trim();
                                            if (TestString == "")
                                            {
                                                //
                                                // blank is a select all
                                                //
                                                Criteria = "";
                                                DataRecordName = "";
                                                DataRecordGuid = "";
                                            }
                                            else if (!supportsGuid)
                                            {
                                                //
                                                // if no guid, this is name
                                                //
                                                DataRecordName = TestString;
                                                DataRecordGuid = "";
                                                Criteria = "name=" + cp.Db.EncodeSQLText(DataRecordName);
                                            }
                                            else if (((TestString.Length) == 38) && (TestString.StartsWith("{")) && (TestString.Substring(Math.Max(TestString.Length - 1, 0)) == "}"))
                                            {
                                                //
                                                // guid {726ED098-5A9E-49A9-8840-767A74F41D01} format
                                                //
                                                DataRecordGuid = TestString;
                                                DataRecordName = "";
                                                Criteria = "ccguid=" +  cp.Db.EncodeSQLText(DataRecordGuid);
                                            }
                                            else if (((TestString.Length) == 36) && (TestString.Substring(8, Math.Min(1, TestString.Length - 8)) == "-"))
                                            {
                                                //
                                                // guid 726ED098-5A9E-49A9-8840-767A74F41D01 format
                                                //
                                                DataRecordGuid = TestString;
                                                DataRecordName = "";
                                                Criteria = "ccguid=" + cp.Db.EncodeSQLText(DataRecordGuid);
                                            }
                                            else if (((TestString.Length) == 32) && ((TestString.IndexOf(' ') + 1) == 0))
                                            {
                                                //
                                                // guid 726ED0985A9E49A98840767A74F41D01 format
                                                //
                                                DataRecordGuid = TestString;
                                                DataRecordName = "";
                                                Criteria = "ccguid=" + cp.Db.EncodeSQLText(DataRecordGuid);
                                            }
                                            else
                                            {
                                                //
                                                // use name
                                                //
                                                DataRecordName = TestString;
                                                DataRecordGuid = "";
                                                Criteria = "name=" + cp.Db.EncodeSQLText(DataRecordName);
                                            }
                                        }
                                        CSData = ReflectionHelper.Invoke<int>(Main, "OpenCSContent", new object[] { DataContentName, Criteria, "id");
                                        if (~ReflectionHelper.Invoke<int>(Main, "IsCSOK", new object[] { CSData }) != 0)
                                        {
                                            RecordNodes = "" + RecordNodes + "\r\n\t" + "<!-- data missing, record not found during export, content=\"" + DataContentName + "\" guid=\"" + DataRecordGuid + "\" name=\"" + DataRecordName + "\" -->";
                                        }
                                        else
                                        {
                                            //
                                            // determine all valid fields
                                            //
                                            fieldCnt = 0;
                                            FieldName = ReflectionHelper.Invoke<string>(Main, "GetCSFirstFieldName", new object[] { CSData });

                                            while ((FieldName != ""))
                                            {
                                                FieldName = ReflectionHelper.Invoke<string>(Main, "GetCSNextFieldName", new object[] { CSData });
                                                FieldLookupContentID = 0;
                                                FieldLookupContentName = "";
                                                if (FieldName != "")
                                                {
                                                    string switchVar = FieldName.ToLower();
                                                    if (switchVar == "ccguid" || switchVar == "name" || switchVar == "id" || switchVar == "dateadded" || switchVar == "createdby" || switchVar == "modifiedby" || switchVar == "modifieddate" || switchVar == "createkey" || switchVar == "contentcontrolid" || switchVar == "editsourceid" || switchVar == "editarchive" || switchVar == "editblank" || switchVar == "contentcategoryid")
                                                    {
                                                    }
                                                    else
                                                    {
                                                        FieldTypeNumber = kmaCommonModule.kmaEncodeInteger(ReflectionHelper.Invoke(Main, "GetContentFieldProperty", new object[] { DataContentName, FieldName, "fieldtype" }));
                                                        if (FieldTypeNumber == 7)
                                                        {
                                                            CSlookup = ReflectionHelper.Invoke<int>(Main, "OpenCSSQL", new object[] { "", "select top 1 LookupContentId,LookupList from ccFields where contentid=" + DataContentId.ToString() + " and name=" + cp.Db.EncodeSQLText(FieldName) });
                                                            if (ReflectionHelper.Invoke<bool>(Main, "IsCSOK", new object[] { CSlookup }))
                                                            {
                                                                FieldLookupContentID = ReflectionHelper.Invoke<int>(Main, "GetCSInteger", new object[] { CSlookup, "Lookupcontentid");
                                                                fieldLookupListValue = cs.GetText( lookup, "LookupList");
                                                            }
                                                            ReflectionHelper.Invoke(Main, "CloseCS", new object[] { CSlookup });
                                                            if (FieldLookupContentID != 0)
                                                            {
                                                                FieldLookupContentName = ReflectionHelper.Invoke<string>(Main, "GetContentNameByID", new object[] { FieldLookupContentID });
                                                            }
                                                        }
                                                        switch (FieldTypeNumber)
                                                        {
                                                            case ccCommonModule.FieldTypeLookup:
                                                            case ccCommonModule.FieldTypeBoolean:
                                                            case ccCommonModule.FieldTypeCSSFile:
                                                            case ccCommonModule.FieldTypeJavascriptFile:
                                                            case ccCommonModule.FieldTypeTextFile:
                                                            case ccCommonModule.FieldTypeXMLFile:
                                                            case ccCommonModule.FieldTypeCurrency:
                                                            case ccCommonModule.FieldTypeFloat:
                                                            case ccCommonModule.FieldTypeInteger:
                                                            case ccCommonModule.FieldTypeDate:
                                                            case ccCommonModule.FieldTypeLink:
                                                            case ccCommonModule.FieldTypeLongText:
                                                            case ccCommonModule.FieldTypeResourceLink:
                                                            case ccCommonModule.FieldTypeText:
                                                            case ccCommonModule.FieldTypeHTML:
                                                            case ccCommonModule.FieldTypeHTMLFile:
                                                                // 
                                                                // this is a keeper 
                                                                // 
                                                                fieldNames = ArraysHelper.RedimPreserve(fieldNames, new int[] { fieldCnt + 1 });
                                                                fieldTypes = ArraysHelper.RedimPreserve(fieldTypes, new int[] { fieldCnt + 1 });
                                                                fieldLookupContent = ArraysHelper.RedimPreserve(fieldLookupContent, new int[] { fieldCnt + 1 });
                                                                fieldLookupList = ArraysHelper.RedimPreserve(fieldLookupList, new int[] { fieldCnt + 1 });
                                                                //fieldLookupContent 
                                                                fieldNames[fieldCnt] = FieldName;
                                                                fieldTypes[fieldCnt] = FieldTypeNumber;
                                                                fieldLookupContent[fieldCnt] = FieldLookupContentName;
                                                                fieldLookupList[fieldCnt] = fieldLookupListValue;
                                                                fieldCnt++;
                                                                //end case 
                                                                break;
                                                        }
                                                        //end case
                                                    }
                                                }
                                            };
                                            //
                                            // output records
                                            //
                                            DataRecordGuid = "";

                                            while (ReflectionHelper.Invoke<bool>(Main, "IsCSOK", new object[] { CSData }))
                                            {
                                                FieldNodes = "";
                                                DataRecordName = ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CSData, "name");
                                                if (supportsGuid)
                                                {
                                                    DataRecordGuid = ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CSData, "ccguid");
                                                    if (DataRecordGuid == "")
                                                    {
                                                        string tempRefParam3 = "";
                                                        DataRecordGuid = GuidGenerator.CreateGUID(ref tempRefParam3);
                                                        ReflectionHelper.Invoke(Main, "SetCS", new object[] { CSData, "ccGuid", DataRecordGuid });
                                                    }
                                                }
                                                for (int fieldPtr = 0; fieldPtr <= fieldCnt - 1; fieldPtr++)
                                                {
                                                    FieldName = fieldNames[fieldPtr];
                                                    FieldTypeNumber = kmaCommonModule.kmaEncodeInteger(fieldTypes[fieldPtr]);
                                                    switch (FieldTypeNumber)
                                                    {
                                                        case ccCommonModule.FieldTypeBoolean:
                                                            // 
                                                            // true/false 
                                                            // 
                                                            FieldValue = ReflectionHelper.Invoke<string>(Main, "GetCSBoolean", new object[] { CSData, FieldName });
                                                            break;
                                                        case ccCommonModule.FieldTypeCSSFile:
                                                        case ccCommonModule.FieldTypeJavascriptFile:
                                                        case ccCommonModule.FieldTypeTextFile:
                                                        case ccCommonModule.FieldTypeXMLFile:
                                                            // 
                                                            // text files 
                                                            // 
                                                            FieldValue = ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CSData, FieldName });
                                                            FieldValue = EncodeCData(FieldValue);
                                                            break;
                                                        case ccCommonModule.FieldTypeCurrency:
                                                        case ccCommonModule.FieldTypeFloat:
                                                        case ccCommonModule.FieldTypeInteger:
                                                            // 
                                                            // numbers 
                                                            // 
                                                            FieldValue = ReflectionHelper.Invoke<string>(Main, "GetCSNumber", new object[] { CSData, FieldName });
                                                            break;
                                                        case ccCommonModule.FieldTypeDate:
                                                            // 
                                                            // date 
                                                            // 
                                                            FieldValue = ReflectionHelper.Invoke<string>(Main, "GetCSDate", new object[] { CSData, FieldName });
                                                            break;
                                                        case ccCommonModule.FieldTypeLookup:
                                                            // 
                                                            // lookup 
                                                            // 
                                                            FieldValue = "";
                                                            FieldValueInteger = ReflectionHelper.Invoke<int>(Main, "GetCSInteger", new object[] { CSData, FieldName });
                                                            if (FieldValueInteger != 0)
                                                            {
                                                                FieldLookupContentName = fieldLookupContent[fieldPtr];
                                                                fieldLookupListValue = fieldLookupList[fieldPtr];
                                                                if (FieldLookupContentName != "")
                                                                {
                                                                    //
                                                                    // content lookup
                                                                    //
                                                                    if (ReflectionHelper.Invoke<bool>(Main, "IsContentFieldSupported", new object[] { FieldLookupContentName, "ccguid" }))
                                                                    {
                                                                        CSlookup = ReflectionHelper.Invoke<int>(Main, "OpenCSContentRecord", new object[] { FieldLookupContentName, FieldValueInteger });
                                                                        if (ReflectionHelper.Invoke<bool>(Main, "IsCSOK", new object[] { CSlookup }))
                                                                        {
                                                                            FieldValue = cs.GetText( lookup, "ccguid");
                                                                            if (FieldValue == "")
                                                                            {
                                                                                string tempRefParam4 = "";
                                                                                FieldValue = GuidGenerator.CreateGUID(ref tempRefParam4);
                                                                                ReflectionHelper.Invoke(Main, "SetCS", new object[] { CSlookup, "ccGuid", FieldValue });
                                                                            }
                                                                        }
                                                                        ReflectionHelper.Invoke(Main, "CloseCS", new object[] { CSlookup });
                                                                    }
                                                                }
                                                                else if (fieldLookupListValue != "")
                                                                {
                                                                    //
                                                                    // list lookup, ok to save integer
                                                                    //
                                                                    FieldValue = FieldValueInteger.ToString();
                                                                }
                                                            }
                                                            break;
                                                        default:
                                                            // 
                                                            // text types 
                                                            // 
                                                            FieldValue = ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CSData, FieldName });
                                                            FieldValue = EncodeCData(FieldValue);
                                                            break;
                                                    }
                                                    FieldNodes = FieldNodes + "\r\n\t" + "<field name=\"" + cp.Utils.EncodeHTML(FieldName) + "\">" + FieldValue + "</field>";
                                                }
                                                RecordNodes = "" + RecordNodes + "\r\n\t" + "<record content=\"" + cp.Utils.EncodeHTML(DataContentName) + "\" guid=\"" + DataRecordGuid + "\" name=\"" + cp.Utils.EncodeHTML(DataRecordName) + "\">" + kmaCommonModule.kmaIndent(ref FieldNodes) + "\r\n\t" + "</record>";
                                                ReflectionHelper.Invoke(Main, "NextCSRecord", new object[] { CSData });
                                            };
                                        }
                                        ReflectionHelper.Invoke(Main, "CloseCS", new object[] { CSData });
                                        // include name only records
                                        //If FieldNodes <> "" Then
                                        //End If
                                    }
                                }
                            }
                        }
                        if (RecordNodes != "")
                        {
                            s = "" + s + "\r\n\t" + "<data>" + kmaCommonModule.kmaIndent(ref RecordNodes) + "\r\n\t" + "</data>";
                        }
                    }
                    //
                    // CDef
                    //
                    //ReflectionHelper.Invoke(Main, "testpoint", new object[] { "getCollection, 700");
                    cs2.open( "Add-on Collection CDef Rules", "CollectionID=" + CollectionID.ToString() });

                    while (ReflectionHelper.Invoke<bool>(Main, "IsCSOK", new object[] { CS2 }))
                    {
                        ContentID = ReflectionHelper.Invoke<int>(Main, "GetCSInteger", new object[] { CS2, "contentid");
                        //
                        // get name and make sure there is a guid
                        //
                        reload = false;
                        CS3 = ReflectionHelper.Invoke<int>(Main, "OpenCSContentRecord", new object[] { "content", ContentID });
                        if (ReflectionHelper.Invoke<bool>(Main, "IsCSOK", new object[] { CS3 }))
                        {
                            ContentName = ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS3, "name");
                            if (cs.GetText( 3, "ccguid" }) == "")
                            {
                                string tempRefParam5 = "";
                                ReflectionHelper.Invoke(Main, "SetCS", new object[] { CS3, "ccGuid", GuidGenerator.CreateGUID(ref tempRefParam5) });
                                reload = true;
                            }
                        }
                        ReflectionHelper.Invoke(Main, "CloseCS", new object[] { CS3 });
                        if (reload)
                        {
                            ReflectionHelper.Invoke(Main, "LoadContentDefinitions", new object[] { });
                        }
                        //
                        Node = new StringBuilder(ReflectionHelper.Invoke<string>(Csv, "GetXMLContentDefinition", new object[] { ContentName }));
                        //
                        // remove the <collection> top node
                        //
                        Pos = (Node.ToString().IndexOf("<cdef", StringComparison.CurrentCultureIgnoreCase) + 1);
                        if (Pos > 0)
                        {
                            Node = new StringBuilder(Node.ToString().Substring(Pos - 1));
                            Pos = (Node.ToString().IndexOf("</cdef>", StringComparison.CurrentCultureIgnoreCase) + 1);
                            if (Pos > 0)
                            {
                                Node = new StringBuilder(Node.ToString().Substring(0, Math.Min(Pos + 6, Node.ToString().Length)));
                                s = s + "\r\n\t" + Node.ToString();
                            }
                        }
                        ReflectionHelper.Invoke(Main, "NextCSRecord", new object[] { CS2 });
                    };
                    ReflectionHelper.Invoke(Main, "CloseCS", new object[] { CS2 });
                    //
                    // Scripting Modules
                    //
                    //ReflectionHelper.Invoke(Main, "testpoint", new object[] { "getCollection, 800");

                    if (IncludeModuleGuidList != "")
                    {
                        Modules = (string[])IncludeModuleGuidList.Split(Environment.NewLine[0]);
                        foreach (string Modules_item in Modules)
                        {
                            ModuleGuid = Modules_item;
                            if (ModuleGuid != "")
                            {
                                cs2.open( "Scripting Modules", "ccguid=" + cp.Db.EncodeSQLText(ModuleGuid) });
                                if (ReflectionHelper.Invoke<bool>(Main, "IsCSOK", new object[] { CS2 }))
                                {
                                    Code = ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS2, "code" }).Trim();
                                    Code = EncodeCData(Code);
                                    s = s + "\r\n\t" + "<ScriptingModule Name=\"" + cp.Utils.EncodeHTML(ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS2, "name" })) + "\" guid=\"" + ModuleGuid + "\">" + Code + "</ScriptingModule>";
                                }
                                ReflectionHelper.Invoke(Main, "CloseCS", new object[] { CS2 });
                            }
                        }
                    }
                    //
                    // shared styles
                    //
                    if (IncludeSharedStyleGuidList != "")
                    {
                        recordGuids = (string[])IncludeSharedStyleGuidList.Split(Environment.NewLine[0]);
                        foreach (string recordGuids_item in recordGuids)
                        {
                            recordGuid = recordGuids_item;
                            if (recordGuid != "")
                            {
                                cs2.open( "Shared Styles", "ccguid=" + cp.Db.EncodeSQLText(recordGuid) });
                                if (ReflectionHelper.Invoke<bool>(Main, "IsCSOK", new object[] { CS2 }))
                                {
                                    s = s + "\r\n\t" + "<SharedStyle" + " Name=\"" + cp.Utils.EncodeHTML(ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS2, "name" })) + "\"" + " guid=\"" + recordGuid + "\"" + " alwaysInclude=\"" + ReflectionHelper.Invoke<string>(Main, "GetCSBoolean", new object[] { CS2, "alwaysInclude" }) + "\"" + " prefix=\"" + cp.Utils.EncodeHTML(cs.GetText( 2, "prefix" })) + "\"" + " suffix=\"" + cp.Utils.EncodeHTML(cs.GetText( 2, "suffix" })) + "\"" + " sortOrder=\"" + cp.Utils.EncodeHTML(cs.GetText( 2, "sortOrder" })) + "\"" + ">" + EncodeCData(ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS2, "styleFilename" }).Trim()) + "</SharedStyle>";
                                }
                                ReflectionHelper.Invoke(Main, "CloseCS", new object[] { CS2 });
                            }
                        }
                    }
                    //
                    // Import Collections
                    //
                    Node = new StringBuilder("");
                    cs3.open( "Add-on Collection Parent Rules", "parentid=" + CollectionID.ToString() });

                    while (ReflectionHelper.Invoke<bool>(Main, "IsCSOK", new object[] { CS3 }))
                    {
                        CS2 = ReflectionHelper.Invoke<int>(Main, "OpenCSContentRecord", new object[] { "Add-on Collections", ReflectionHelper.Invoke(Main, "GetCSInteger", new object[] { CS3, "childid" }) });
                        if (ReflectionHelper.Invoke<bool>(Main, "IsCSOK", new object[] { CS2 }))
                        {
                            Guid = cs.GetText( 2, "ccGuid");
                            if (Guid == "")
                            {
                                string tempRefParam6 = "";
                                Guid = GuidGenerator.CreateGUID(ref tempRefParam6);
                                ReflectionHelper.Invoke(Main, "SetCS", new object[] { CS2, "ccGuid", Guid });
                            }
                            Node.Append("\r\n\t" + "<ImportCollection name=\"" + cp.Utils.EncodeHTML(cs.GetText( 2, "name" })) + "\">" + Guid + "</ImportCollection>");
                        }
                        ReflectionHelper.Invoke(Main, "CloseCS", new object[] { CS2 });
                        ReflectionHelper.Invoke(Main, "NextCSRecord", new object[] { CS3 });
                    };
                    ReflectionHelper.Invoke(Main, "CloseCS", new object[] { CS3 });
                    s = s + Node.ToString();
                    //
                    // wwwFileList
                    //
                    ResourceCnt = 0;
                    FileList = cs.GetText( "wwwFileList");
                    if (FileList != "")
                    {
                        PhysicalWWWPath = ReflectionHelper.GetMember<string>(Main, "PhysicalWWWPath");
                        if (PhysicalWWWPath.Substring(Math.Max(PhysicalWWWPath.Length - 1, 0)) != "\\")
                        {
                            PhysicalWWWPath = PhysicalWWWPath + "\\";
                        }
                        Files = (string[])FileList.Split(Environment.NewLine[0]);
                        foreach (string Files_item_2 in Files)
                        {
                            PathFilename = Files_item_2;
                            if (PathFilename != "")
                            {
                                PathFilename = Strings.Replace(PathFilename, "\\", "/", 1, -1, CompareMethod.Binary);
                                Path = "";
                                Filename = PathFilename;
                                Pos = PathFilename.LastIndexOf("/") + 1;
                                if (Pos > 0)
                                {
                                    Filename = PathFilename.Substring(Pos);
                                    Path = PathFilename.Substring(0, Math.Min(Pos - 1, PathFilename.Length));
                                }
                                if (Filename.ToLower() == "collection.hlp")
                                {
                                    //
                                    // legacy file, remove it
                                    //
                                }
                                else
                                {
                                    PathFilename = Strings.Replace(PathFilename, "/", "\\", 1, -1, CompareMethod.Binary);
                                    AddFilename = PhysicalWWWPath + PathFilename;
                                    if ((AddFileList.IndexOf("\\" + Filename, StringComparison.CurrentCultureIgnoreCase) + 1) != 0)
                                    {
                                        cp.UserError.Add( "There was an error exporting this collection because there were multiple files with the same filename [" + Filename + "]");
                                    }
                                    else
                                    {
                                        s = s + "\r\n\t" + "<Resource name=\"" + cp.Utils.EncodeHTML(Filename) + "\" type=\"www\" path=\"" + cp.Utils.EncodeHTML(Path) + "\" />";
                                        AddFileList = AddFileList + Environment.NewLine + AddFilename;
                                    }
                                    ResourceCnt++;
                                }
                            }
                        }
                    }
                    //
                    // ContentFileList
                    //
                    FileList = cs.GetText( "ContentFileList");
                    if (FileList != "")
                    {
                        Files = (string[])FileList.Split(Environment.NewLine[0]);
                        foreach (string Files_item_3 in Files)
                        {
                            PathFilename = Files_item_3;
                            if (PathFilename != "")
                            {
                                PathFilename = Strings.Replace(PathFilename, "\\", "/", 1, -1, CompareMethod.Binary);
                                Path = "";
                                Filename = PathFilename;
                                Pos = PathFilename.LastIndexOf("/") + 1;
                                if (Pos > 0)
                                {
                                    Filename = PathFilename.Substring(Pos);
                                    Path = PathFilename.Substring(0, Math.Min(Pos - 1, PathFilename.Length));
                                }
                                PathFilename = Strings.Replace(PathFilename, "/", "\\", 1, -1, CompareMethod.Binary);
                                if (PathFilename.StartsWith("\\"))
                                {
                                    PathFilename = PathFilename.Substring(1);
                                }
                                AddFilename = cp.Site.PhysicalFilePath + PathFilename;
                                if ((AddFileList.IndexOf("\\" + Filename, StringComparison.CurrentCultureIgnoreCase) + 1) != 0)
                                {
                                    cp.UserError.Add( "There was an error exporting this collection because there were multiple files with the same filename [" + Filename + "]");
                                }
                                else
                                {
                                    s = s + "\r\n\t" + "<Resource name=\"" + cp.Utils.EncodeHTML(Filename) + "\" type=\"content\" path=\"" + cp.Utils.EncodeHTML(Path) + "\" />";
                                    AddFileList = AddFileList + Environment.NewLine + AddFilename;
                                }
                                ResourceCnt++;
                            }
                        }
                    }
                    //
                    // ExecFileListNode
                    //
                    s = s + ExecFileListNode.ToString();
                    //
                    // Other XML
                    //
                    OtherXML = cs.GetText( "otherxml");
                    if (OtherXML.Trim() != "")
                    {
                        s = s + Environment.NewLine + OtherXML;
                    }
                    //s = s & GetNodeText("OtherXML", Main.GetCSText(CS, "otherxml"))
                    //    & "<Data></Data>" _
                    //'    & "<NavigatorEntry></NavigatorEntry>" _
                    //'    & "<SQLIndex></SQLIndex>" _
                    //'    & "<Styles></Styles>" _
                    //'    & "<ScriptingModule Name="" guid="" />" _
                    //'    & "<Resource name="" type="www|content|executable" path="" />"
                    s = s + "\r\n</Collection>";
                    //
                    ReflectionHelper.Invoke(Main, "CloseCS", new object[] { CS });
                    //
                    // Save the installation file and add it to the archive
                    //
                    kmafs.SaveFile(ref InstallFilename, ref s);
                    //Call fs.SaveFile(InstallFilename, s)
                    if (((Environment.NewLine + AddFileList).IndexOf(Environment.NewLine + InstallFilename, StringComparison.CurrentCultureIgnoreCase) + 1) == 0)
                    {
                        AddFileList = AddFileList + Environment.NewLine + InstallFilename;
                    }
                    kmafs.SaveFile(ref AddFileListFilename, ref AddFileList);
                    //Call fs.SaveFile(AddFileListFilename, AddFileList)
                    runAtServer("zipfile", "archive=" + kmaCommonModule.kmaEncodeRequestVariable(ArchiveFilename) + "&add=" + kmaCommonModule.kmaEncodeRequestVariable("@" + AddFileListFilename));
                    //Set Remote = CreateObject("ccRemote.RemoteClass")
                    //Call Remote.executeCmd("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable("@" & AddFileListFilename))
                }
                //ReflectionHelper.Invoke(Main, "testpoint", new object[] { "getCollection, exit");
                //
                //
            }
            catch
            {
                HandleClassTrapError("GetCollection");
            }
            return result;
        }
        //
        //
        //
        //UPGRADE_NOTE: (7001) The following declaration (GetCDef) seems to be dead code More Information: http://www.vbtonet.com/ewis/ewi7001.aspx
        //private string GetCDef(int CollectionID)
        //{
        //try
        //{
        ////
        //string s = "";
        ////
        ////
        //return s;
        //}
        //catch
        //{
        //HandleClassTrapError("GetCDef");
        //}
        //return "";
        //}
        //
        //
        //
        //UPGRADE_NOTE: (7001) The following declaration (getNavigatorNode) seems to be dead code More Information: http://www.vbtonet.com/ewis/ewi7001.aspx
        //private string getNavigatorNode(int navigatorId, string Return_IncludeModuleGuidList, bool Ver40Compatibility, string Ver40DLLList)
        //{
        //try
        //{
        ////
        //string s = "";
        //int CS = 0;
        //CSGUID.GUIDGenerator GuidGenerator = new CSGUID.GUIDGenerator();
        //string nameSpace = "";
        ////
        //CS = ReflectionHelper.Invoke<int>(Main, "OpenCSContentRecord", new object[]{"Navigator Entries", navigatorId});
        //if (ReflectionHelper.Invoke<bool>(Main, "IsCSOK", new object[]{CS}))
        //{
        ////
        // ActiveX DLL node is being deprecated. This should be in the collection resource section
        ////
        //s = s + GetNodeText("Active", ReflectionHelper.Invoke<string>(Main, "GetCS", new object[]{CS, "active"}));
        ////
        ////
        ////
        //s = "\n\r\t" + "<NavigatorEntry name=\"" + cp.Utils.EncodeHTML(ReflectionHelper.Invoke<string>(Main, "GetCS", new object[]{CS, "name"})) + "\" guid=\"" + cp.Utils.EncodeHTML(ReflectionHelper.Invoke<string>(Main, "GetCS", new object[]{CS, "guid"})) + "\" namespace=\"" + nameSpace + "\">" + kmaCommonModule.kmaIndent(ref s) + "\r\n\t" + "</NavigatorEntry>";
        //}
        //ReflectionHelper.Invoke(Main, "CloseCS", new object[]{CS});
        ////
        ////
        //return s;
        //}
        //catch
        //{
        //HandleClassTrapError("getNavigatorNode");
        //}
        //return "";
        //}
        //
        //
        //
        private string GetAddonNode(int addonid, ref string Return_IncludeModuleGuidList, bool Ver40Compatibility, string Ver40DLLList, ref string Return_IncludeSharedStyleGuidList)
        {
            try
            {
                //
                int styleId = 0;
                string fieldType = "";
                int fieldTypeID = 0;
                int TriggerContentID = 0;
                string StylesTest = "";
                bool BlockEditTools = false;
                string NavType = "";
                string Styles = "";
                string s = "";
                int CS = 0;
                int CS2 = 0;
                int CS3 = 0;
                string Node = "";
                string NodeInnerText = "";
                CSGUID.GUIDGenerator GuidGenerator = new CSGUID.GUIDGenerator();
                int FilterAddonID = 0;
                int IncludedAddonID = 0;
                int ScriptingModuleID = 0;
                string Guid = "";
                string[] ListSplit = null;
                string Filename = "";
                string addonName = "";
                bool processRunOnce = false;
                //
                CS = ReflectionHelper.Invoke<int>(Main, "OpenCSContentRecord", new object[] { "Add-ons", addonid });
                if (ReflectionHelper.Invoke<bool>(Main, "IsCSOK", new object[] { CS }))
                {
                    addonName = ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "name");
                    processRunOnce = cs.GetBoolean( "ProcessRunOnce");
                    if ((addonName.ToLower() == "oninstall") || (addonName.ToLower() == "_oninstall"))
                    {
                        processRunOnce = true;
                    }
                    //
                    // ActiveX DLL node is being deprecated. This should be in the collection resource section
                    //
                    s = s + GetNodeText("Copy", ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "Copy" }));
                    s = s + GetNodeText("CopyText", ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "CopyText" }));
                    //
                    // DLL
                    //

                    s = s + GetNodeText("ActiveXProgramID", ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "objectprogramid" }));
                    s = s + GetNodeText("DotNetClass", ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "DotNetClass" }));
                    if (Ver40Compatibility && (Ver40DLLList != ""))
                    {
                        ListSplit = (string[])Ver40DLLList.Split(Environment.NewLine[0]);
                        foreach (string ListSplit_item in ListSplit)
                        {
                            Filename = ListSplit_item.Trim();
                            if (Filename != "")
                            {
                                s = s + GetNodeText("ActiveXDLL", Filename) + "<!-- Version 40 Compatibility -->";
                            }
                        }
                    }
                    //
                    // Features
                    //
                    s = s + GetNodeText("ArgumentList", ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "ArgumentList" }));
                    s = s + GetNodeBoolean("AsAjax", cs.GetBoolean( "AsAjax" }));
                    s = s + GetNodeBoolean("Filter", cs.GetBoolean( "Filter"));
                    s = s + GetNodeText("Help", ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "Help"));
                    s = s + GetNodeText("HelpLink", ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "HelpLink"));
                    s = s + "\r\n\t" + "<Icon Link=\"" + cs.GetText( "iconfilename") + "\" width=\"" + ReflectionHelper.Invoke<string>(Main, "GetCSInteger", new object[] { CS, "iconWidth") + "\" height=\"" + ReflectionHelper.Invoke<string>(Main, "GetCSInteger", new object[] { CS, "iconHeight") + "\" sprites=\"" + ReflectionHelper.Invoke<string>(Main, "GetCSInteger", new object[] { CS, "iconSprites") + "\" />";
                    s = s + GetNodeBoolean("InIframe", cs.GetBoolean( "InFrame"));
                    BlockEditTools = false;
                    if (ReflectionHelper.Invoke<bool>(Main, "IsCSFieldSupported", new object[] { CS, "BlockEditTools"))
                    {
                        BlockEditTools = cs.GetBoolean( "BlockEditTools");
                    }
                    s = s + GetNodeBoolean("BlockEditTools", BlockEditTools);
                    //
                    // Form XML
                    //
                    s = s + GetNodeText("FormXML", ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "FormXML"));
                    //
                    NodeInnerText = "";
                    cs2.open( "Add-on Include Rules", "addonid=" + addonid.ToString() });

                    while (ReflectionHelper.Invoke<bool>(Main, "IsCSOK", new object[] { CS2 }))
                    {
                        IncludedAddonID = ReflectionHelper.Invoke<int>(Main, "GetCSInteger", new object[] { CS2, "IncludedAddonID");
                        cs3.open( "Add-ons", "ID=" + IncludedAddonID.ToString() });
                        if (ReflectionHelper.Invoke<bool>(Main, "IsCSOK", new object[] { CS3 }))
                        {
                            Guid = ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS3, "ccGuid");
                            if (Guid == "")
                            {
                                string tempRefParam = "";
                                Guid = GuidGenerator.CreateGUID(ref tempRefParam);
                                ReflectionHelper.Invoke(Main, "SetCS", new object[] { CS3, "ccGuid", Guid });
                            }
                            s = s + "\r\n\t" + "<IncludeAddon name=\"" + cp.Utils.EncodeHTML(ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS3, "name")) + "\" guid=\"" + Guid + "\"/>";
                        }
                        ReflectionHelper.Invoke(Main, "CloseCS", new object[] { CS3 });
                        ReflectionHelper.Invoke(Main, "NextCSRecord", new object[] { CS2 });
                    };
                    ReflectionHelper.Invoke(Main, "CloseCS", new object[] { CS2 });
                    //
                    s = s + GetNodeBoolean("IsInline", cs.GetBoolean( "IsInline" }));
                    s = s + GetNodeText("JavascriptOnLoad", ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "JavascriptOnLoad" }));
                    s = s + GetNodeText("JavascriptInHead", ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "JSFilename" }));
                    s = s + GetNodeText("JavascriptBodyEnd", ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "JavascriptBodyEnd" }));
                    s = s + GetNodeText("MetaDescription", ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "MetaDescription" }));
                    s = s + GetNodeText("OtherHeadTags", ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "OtherHeadTags" }));
                    //
                    // Placements
                    //
                    s = s + GetNodeBoolean("Content", cs.GetBoolean( "Content" }));
                    s = s + GetNodeBoolean("Template", cs.GetBoolean( "Template" }));
                    s = s + GetNodeBoolean("Email", cs.GetBoolean( "Email" }));
                    s = s + GetNodeBoolean("Admin", cs.GetBoolean( "Admin" }));
                    s = s + GetNodeBoolean("OnPageEndEvent", cs.GetBoolean( "OnPageEndEvent" }));
                    s = s + GetNodeBoolean("OnPageStartEvent", cs.GetBoolean( "OnPageStartEvent" }));
                    s = s + GetNodeBoolean("OnBodyStart", cs.GetBoolean( "OnBodyStart" }));
                    s = s + GetNodeBoolean("OnBodyEnd", cs.GetBoolean( "OnBodyEnd" }));
                    s = s + GetNodeBoolean("RemoteMethod", cs.GetBoolean( "RemoteMethod" }));
                    //s = s & GetNodeBoolean("OnNewVisitEvent", Main.GetCSBoolean(CS, "OnNewVisitEvent"))
                    //
                    // Process
                    //
                    s = s + GetNodeBoolean("ProcessRunOnce", processRunOnce);
                    s = s + GetNodeInteger("ProcessInterval", ReflectionHelper.Invoke<int>(Main, "GetCSInteger", new object[] { CS, "ProcessInterval" }));
                    //
                    // Meta
                    //
                    s = s + GetNodeText("PageTitle", ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "PageTitle" }));
                    s = s + GetNodeText("RemoteAssetLink", ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "RemoteAssetLink" }));
                    //
                    // Styles
                    //
                    Styles = "";
                    if (~ReflectionHelper.Invoke<int>(Main, "GetCSBoolean", new object[] { CS, "BlockDefaultStyles" }) != 0)
                    {
                        Styles = ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "StylesFilename" }).Trim();
                    }
                    StylesTest = ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "CustomStylesFilename" }).Trim();
                    if (StylesTest != "")
                    {
                        if (Styles != "")
                        {
                            Styles = Styles + Environment.NewLine + StylesTest;
                        }
                        else
                        {
                            Styles = StylesTest;
                        }
                    }
                    s = s + GetNodeText("Styles", Styles);
                    //
                    // Scripting
                    //
                    NodeInnerText = ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "ScriptingCode" }).Trim();
                    if (NodeInnerText != "")
                    {
                        NodeInnerText = "\r\n\t" + "\t" + "<Code>" + EncodeCData(NodeInnerText) + "</Code>";
                    }
                    cs2.open( "Add-on Scripting Module Rules", "addonid=" + addonid.ToString() });

                    while (ReflectionHelper.Invoke<bool>(Main, "IsCSOK", new object[] { CS2 }))
                    {
                        ScriptingModuleID = ReflectionHelper.Invoke<int>(Main, "GetCSInteger", new object[] { CS2, "ScriptingModuleID");
                        cs3.open( "Scripting Modules", "ID=" + ScriptingModuleID.ToString() });
                        if (ReflectionHelper.Invoke<bool>(Main, "IsCSOK", new object[] { CS3 }))
                        {
                            Guid = ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS3, "ccGuid");
                            if (Guid == "")
                            {
                                string tempRefParam2 = "";
                                Guid = GuidGenerator.CreateGUID(ref tempRefParam2);
                                ReflectionHelper.Invoke(Main, "SetCS", new object[] { CS3, "ccGuid", Guid });
                            }
                            Return_IncludeModuleGuidList = Return_IncludeModuleGuidList + Environment.NewLine + Guid;
                            NodeInnerText = NodeInnerText + "\r\n\t" + "\t" + "<IncludeModule name=\"" + cp.Utils.EncodeHTML(ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS3, "name" })) + "\" guid=\"" + Guid + "\"/>";
                        }
                        ReflectionHelper.Invoke(Main, "CloseCS", new object[] { CS3 });
                        ReflectionHelper.Invoke(Main, "NextCSRecord", new object[] { CS2 });
                    };
                    ReflectionHelper.Invoke(Main, "CloseCS", new object[] { CS2 });
                    if (NodeInnerText == "")
                    {
                        s = s + "\r\n\t" + "<Scripting Language=\"" + ReflectionHelper.Invoke<string>(Main, "GetCSLookup", new object[] { CS, "ScriptingLanguageID" }) + "\" EntryPoint=\"" + ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "ScriptingEntryPoint" }) + "\" Timeout=\"" + ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "ScriptingTimeout" }) + "\"/>";
                    }
                    else
                    {
                        s = s + "\r\n\t" + "<Scripting Language=\"" + ReflectionHelper.Invoke<string>(Main, "GetCSLookup", new object[] { CS, "ScriptingLanguageID" }) + "\" EntryPoint=\"" + ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "ScriptingEntryPoint" }) + "\" Timeout=\"" + ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "ScriptingTimeout" }) + "\">" + NodeInnerText + "\r\n\t" + "</Scripting>";
                    }
                    //
                    // Shared Styles
                    //
                    cs2.open( "Shared Styles Add-on Rules", "addonid=" + addonid.ToString() });

                    while (ReflectionHelper.Invoke<bool>(Main, "IsCSOK", new object[] { CS2 }))
                    {
                        styleId = ReflectionHelper.Invoke<int>(Main, "GetCSInteger", new object[] { CS2, "styleId");
                        cs3.open( "shared styles", "ID=" + styleId.ToString() });
                        if (ReflectionHelper.Invoke<bool>(Main, "IsCSOK", new object[] { CS3 }))
                        {
                            Guid = ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS3, "ccGuid");
                            if (Guid == "")
                            {
                                string tempRefParam3 = "";
                                Guid = GuidGenerator.CreateGUID(ref tempRefParam3);
                                ReflectionHelper.Invoke(Main, "SetCS", new object[] { CS3, "ccGuid", Guid });
                            }
                            Return_IncludeSharedStyleGuidList = Return_IncludeSharedStyleGuidList + Environment.NewLine + Guid;
                            s = s + "\r\n\t" + "<IncludeSharedStyle name=\"" + cp.Utils.EncodeHTML(ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS3, "name" })) + "\" guid=\"" + Guid + "\"/>";
                        }
                        ReflectionHelper.Invoke(Main, "CloseCS", new object[] { CS3 });
                        ReflectionHelper.Invoke(Main, "NextCSRecord", new object[] { CS2 });
                    };
                    ReflectionHelper.Invoke(Main, "CloseCS", new object[] { CS2 });
                    //
                    // Process Triggers
                    //
                    NodeInnerText = "";
                    cs2.open( "Add-on Content Trigger Rules", "addonid=" + addonid.ToString() });

                    while (ReflectionHelper.Invoke<bool>(Main, "IsCSOK", new object[] { CS2 }))
                    {
                        TriggerContentID = ReflectionHelper.Invoke<int>(Main, "GetCSInteger", new object[] { CS2, "ContentID");
                        cs3.open( "content", "ID=" + TriggerContentID.ToString() });
                        if (ReflectionHelper.Invoke<bool>(Main, "IsCSOK", new object[] { CS3 }))
                        {
                            Guid = ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS3, "ccGuid");
                            if (Guid == "")
                            {
                                string tempRefParam4 = "";
                                Guid = GuidGenerator.CreateGUID(ref tempRefParam4);
                                ReflectionHelper.Invoke(Main, "SetCS", new object[] { CS3, "ccGuid", Guid });
                            }
                            NodeInnerText = NodeInnerText + "\r\n\t" + "\t" + "<ContentChange name=\"" + cp.Utils.EncodeHTML(ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS3, "name" })) + "\" guid=\"" + Guid + "\"/>";
                        }
                        ReflectionHelper.Invoke(Main, "CloseCS", new object[] { CS3 });
                        ReflectionHelper.Invoke(Main, "NextCSRecord", new object[] { CS2 });
                    };
                    ReflectionHelper.Invoke(Main, "CloseCS", new object[] { CS2 });
                    if (NodeInnerText != "")
                    {
                        s = s + "\r\n\t" + "<ProcessTriggers>" + NodeInnerText + "\r\n\t" + "</ProcessTriggers>";
                    }
                    //
                    // Editors
                    //
                    if (ReflectionHelper.Invoke<bool>(Main, "IsContentFieldSupported", new object[] { "Add-on Content Field Type Rules", "id" }))
                    {
                        NodeInnerText = "";
                        cs2.open( "Add-on Content Field Type Rules", "addonid=" + addonid.ToString() });

                        while (ReflectionHelper.Invoke<bool>(Main, "IsCSOK", new object[] { CS2 }))
                        {
                            fieldTypeID = ReflectionHelper.Invoke<int>(Main, "GetCSInteger", new object[] { CS2, "contentFieldTypeID");
                            fieldType = ReflectionHelper.Invoke<string>(Main, "GetRecordName", new object[] { "Content Field Types", fieldTypeID });
                            if (fieldType != "")
                            {
                                NodeInnerText = NodeInnerText + "\r\n\t" + "\t" + "<type>" + fieldType + "</type>";
                            }
                            ReflectionHelper.Invoke(Main, "NextCSRecord", new object[] { CS2 });
                        };
                        ReflectionHelper.Invoke(Main, "CloseCS", new object[] { CS2 });
                        if (NodeInnerText != "")
                        {
                            s = s + "\r\n\t" + "<Editors>" + NodeInnerText + "\r\n\t" + "</Editors>";
                        }
                    }
                    //
                    //
                    //
                    Guid = ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "ccGuid");
                    if (Guid == "")
                    {
                        string tempRefParam5 = "";
                        Guid = GuidGenerator.CreateGUID(ref tempRefParam5);
                        ReflectionHelper.Invoke(Main, "SetCS", new object[] { CS, "ccGuid", Guid });
                    }
                    NavType = ReflectionHelper.Invoke<string>(Main, "GetCS", new object[] { CS, "NavTypeID");
                    if (NavType == "")
                    {
                        NavType = "Add-on";
                    }
                    if (Ver40Compatibility)
                    {
                        s = "\n\r\t" + "<Page name=\"" + cp.Utils.EncodeHTML(addonName) + "\" guid=\"" + Guid + "\" type=\"" + NavType + "\">" + kmaCommonModule.kmaIndent(ref s) + "\r\n\t" + "</Page>";
                        s = "\n\r\t" + "<Interfaces><!-- Version 4.0 Compatibility -->" + kmaCommonModule.kmaIndent(ref s) + "\r\n\t" + "</Interfaces>";
                    }
                    else
                    {
                        s = "\n\r\t" + "<Addon name=\"" + cp.Utils.EncodeHTML(addonName) + "\" guid=\"" + Guid + "\" type=\"" + NavType + "\">" + kmaCommonModule.kmaIndent(ref s) + "\r\n\t" + "</Addon>";
                    }
                }
                ReflectionHelper.Invoke(Main, "CloseCS", new object[] { CS });
                //
                //
                return s;
            }
            catch
            {
                HandleClassTrapError("GetAddonNode");
            }
            return "";
        }
        //'
        //'
        //'
        //Private Function GetProcessInterfaceNode(AddonID As Long, Return_IncludeModuleGuidList As String) As String
        //    On Error GoTo ErrorTrap
        //    '
        //    Dim s As String
        //    Dim CS As Long
        //    Dim CS2 As Long
        //    Dim CS3 As Long
        //    Dim Node As String
        //    Dim GuidGenerator As New GuidGenerator
        //    Dim FilterAddonID As Long
        //    Dim IncludedAddonID As Long
        //    Dim ScriptingModuleID As Long
        //    Dim Guid As String
        //    '
        //    CS = Main.OpenCSContentRecord("Add-ons", AddonID)
        //    If Main.iscsok(CS) Then
        //        '
        //        ' ActiveX DLL node is being deprecated. This should be in the collection resource section
        //        '
        //        's = s & GetNodeText("Description", Main.GetCS(CS, "Description"))
        //        s = s & GetNodeText("Help", Main.GetCS(CS, "Help"))
        //        s = s & GetNodeText("HelpLink", Main.GetCS(CS, "HelpLink"))
        //        s = s & GetNodeText("ArgumentList", Main.GetCS(CS, "ArgumentList"))
        //        s = s & GetNodeText("ActiveXProgramID", Main.GetCS(CS, "objectprogramid"))
        //        s = s & GetNodeBoolean("ProcessRunOnce", Main.GetCSBoolean(CS, "ProcessRunOnce"))
        //        s = s & GetNodeInteger("ProcessInterval", Main.getcsinteger(CS, "ProcessInterval"))
        //        '
        //        ' Add scripting modules
        //        '
        //        Node = ""
        //        CS2 = Main.OpenCSContent("Add-on Scripting Module Rules", "addonid=" & AddonID)
        //        Do While Main.iscsok(CS2)
        //            ScriptingModuleID = Main.getcsinteger(CS2, "ScriptingModuleID")
        //            CS3 = Main.OpenCSContent("Scripting Modules", "ID=" & ScriptingModuleID)
        //            If Main.iscsok(CS3) Then
        //                Guid = Main.GetCS(CS3, "ccGuid")
        //                If Guid = "" Then
        //                    Guid = GuidGenerator.CreateGUID("")
        //                    Call Main.SetCS(CS3, "ccGuid", Guid)
        //                End If
        //                Return_IncludeModuleGuidList = Return_IncludeModuleGuidList & vbCrLf & Guid
        //                Node = Node & vbCrLf & vbTab & "<IncludeModule name=""" & Main.GetCS(CS3, "name") & """ guid=""" & Guid & """/>"
        //            End If
        //            Call Main.CloseCS(CS3)
        //            Call Main.nextcsrecord(CS2)
        //        Loop
        //        Call Main.CloseCS(CS2)
        //        If Node = "" Then
        //            s = s & vbCrLf & vbTab & "<Scripting Language=""" & Main.GetCSLookup(CS, "ScriptingLanguageID") & """ EntryPoint=""" & Main.GetCS(CS, "ScriptingEntryPoint") & """/>"
        //        Else
        //            s = s & vbCrLf & vbTab & "<Scripting Language=""" & Main.GetCSLookup(CS, "ScriptingLanguageID") & """ EntryPoint=""" & Main.GetCS(CS, "ScriptingEntryPoint") & """>" & Node & vbCrLf & vbTab & "</Scripting>"
        //        End If
        //        '
        //        Guid = Main.GetCS(CS, "ccGuid")
        //        If Guid = "" Then
        //            Guid = GuidGenerator.CreateGUID("")
        //            Call Main.SetCS(CS, "ccGuid", Guid)
        //        End If
        //        s = "" _
        //'            & vbCrLf & vbTab & "<Process name=""" & addonName & """ guid=""" & Guid & """>" _
        //'            & KmaIndent(s) _
        //'            & vbCrLf & vbTab & "</Process>"
        //    End If
        //    Call Main.CloseCS(CS)
        //    '
        //    GetProcessInterfaceNode = s
        //    '
        //    Exit Function
        //ErrorTrap:
        //    Call HandleClassTrapError("GetProcessInterfaceNode")
        //End Function
        //
        //
        //
        private string GetNodeText(string NodeName, string NodeContent)
        {
            string result = "";
            try
            {
                //
                //
                if (NodeContent == "")
                {
                    return result + "\r\n\t" + "<" + NodeName + "></" + NodeName + ">";
                }
                else
                {
                    return result + "\r\n\t" + "<" + NodeName + ">" + EncodeCData(NodeContent) + "</" + NodeName + ">";
                }
            }
            catch
            {
                HandleClassTrapError("GetNodeText");
            }
            return result;
        }
        //
        //
        //
        private string GetNodeBoolean(string NodeName, bool NodeContent)
        {
            string result = "";
            try
            {
                //
                //
                return result + "\r\n\t" + "<" + NodeName + ">" + getYesNo(NodeContent) + "</" + NodeName + ">";
            }
            catch
            {
                HandleClassTrapError("GetNodeBoolean");
            }
            return result;
        }
        //
        //
        //
        private string GetNodeInteger(string NodeName, int NodeContent)
        {
            string result = "";
            try
            {
                //
                //
                return result + "\r\n\t" + "<" + NodeName + ">" + NodeContent.ToString() + "</" + NodeName + ">";
            }
            catch
            {
                HandleClassTrapError("GetNodeInteger");
            }
            return result;
        }
        //
        //
        //
        private void HandleClassTrapError(string MethodName)
        {
            //UPGRADE_WARNING: (2081) Err.Number has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2081.aspx
            kmaCommonModule.HandleError("CExportClass", MethodName, Information.Err().Number, Information.Err().Source, Information.Err().Description, true, false, "");
        }
        //
        //========================================================================
        // From ccCommon
        //========================================================================
        //
        public string ReplaceMany(string Source, string[] ArrayOfSource, string[] ArrayOfReplacement)
        {
            // ##### removed to catch err<>0 problem on error resume next
            //
            //
            string result = "";
            int Count = ArrayOfSource.GetUpperBound(0) + 1;
            result = Source;
            for (int Pointer = 0; Pointer <= Count - 1; Pointer++)
            {
                result = Strings.Replace(result, ArrayOfSource[Pointer], ArrayOfReplacement[Pointer], 1, -1, CompareMethod.Binary);
            }
            //
            return result;
        }
        //
        //========================================================================
        // From ccCommon
        //========================================================================
        //
        public string encodeFilename(string Filename)
        {
            // ##### removed to catch err<>0 problem on error resume next
            //
            //
            //UPGRADE_WARNING: (2081) Array has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2081.aspx
            string result = "";
            string[] Source = new string[] { "\"", "*", "/", ":", "<", ">", "?", "\\", "|" };
            //UPGRADE_WARNING: (2081) Array has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2081.aspx
            string[] Replacement = new string[] { "_", "_", "_", "_", "_", "_", "_", "_", "_" };
            //
            result = ReplaceMany(Filename, Source, Replacement);
            if ((result.Length) > 254)
            {
                result = result.Substring(0, Math.Min(254, result.Length));
            }
            //
            return result;
        }
        //
        //
        //
        internal void GetLocalCollectionArgs(string CollectionGuid, ref string Return_CollectionPath, ref System.DateTime Return_LastChagnedate)
        {
            try
            {
                //
                const string CollectionListRootNode = "collectionlist";
                //
                string LocalPath = "";
                string LocalFilename = "";
                string LocalGuid = "";
                XmlDocument Doc = new XmlDocument();
                XmlNode NewCollectionNode = null;
                XmlNode NewAttrNode = null;
                bool CollectionFound = false;
                int Ptr = 0;
                string CollectionPath = "";
                System.DateTime LastChangeDate = DateTime.FromOADate(0);
                StringBuilder //hint = new StringBuilder();
                bool MatchFound = false;
                string LocalName = "";
                //
                MatchFound = false;
                Return_CollectionPath = "";
                Return_LastChagnedate = DateTime.FromOADate(0);
                //hint = new StringBuilder("Match guid [" + CollectionGuid + "], Loading");
                Doc.LoadXml(GetConfig());
                //UPGRADE_ISSUE: (2064) MSXML2.DOMDocument property Doc.readyState was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx

                while (Doc.getreadyState() != 4 && Ptr < 100)
                {
                    hint.Append(",waiting for load");
                    UpgradeSolution1Support.PInvoke.SafeNative.kernel32.Sleep(100);
                    Application.DoEvents();
                    Ptr++;
                };
                XmlElement withVar = null;
                //UPGRADE_ISSUE: (2064) MSXML2.DOMDocument property Doc.parseError was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                //UPGRADE_ISSUE: (2064) MSXML2.IXMLDOMParseError property parseError.errorCode was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                if (Doc.getparseError().geterrorCode() != 0)
                {
                    hint.Append(",parse error");
                    //Call AppendClassLogFile("Server", "", "GetLocalCollectionArgs, Hint=[" & Hint & "], Error loading Collections.xml file. The error was [" & Doc.parseError.reason & ", line " & Doc.parseError.Line & ", character " & Doc.parseError.linepos & "]")
                }
                else
                {
                    if (Doc.DocumentElement.LocalName.ToLower() != CollectionListRootNode.ToLower())
                    {
                        //Call AppendClassLogFile("Server", "", "GetLocalCollectionArgs, Hint=[" & Hint & "], The Collections.xml file has an invalid root node, [" & Doc.documentElement.BaseName & "] was received and [" & CollectionListRootNode & "] was expected.")
                    }
                    else
                    {
                        withVar = Doc.DocumentElement;
                        if (withVar.LocalName.ToLower() != "collectionlist")
                        {
                            //Call AppendClassLogFile("Server", "", "GetLocalCollectionArgs, basename was not collectionlist, [" & .BaseName & "].")
                        }
                        else
                        {
                            CollectionFound = false;
                            hint.Append(",checking nodes [" + withVar.ChildNodes.Count.ToString() + "]");
                            foreach (XmlNode LocalListNode in withVar.ChildNodes)
                            {
                                LocalName = "no name found";
                                LocalPath = "";
                                switch (LocalListNode.LocalName.ToLower())
                                {
                                    case "collection":
                                        LocalGuid = "";
                                        foreach (XmlNode CollectionNode in LocalListNode.ChildNodes)
                                        {
                                            switch (CollectionNode.LocalName.ToLower())
                                            {
                                                case "name":
                                                    // 
                                                    LocalName = CollectionNode.InnerText.ToLower();
                                                    break;
                                                case "guid":
                                                    // 
                                                    LocalGuid = CollectionNode.InnerText.ToLower();
                                                    break;
                                                case "path":
                                                    // 
                                                    CollectionPath = CollectionNode.InnerText.ToLower();
                                                    break;
                                                case "lastchangedate":
                                                    LastChangeDate = kmaCommonModule.KmaEncodeDate(CollectionNode.InnerText);
                                                    break;
                                            }
                                        }
                                        break;
                                }
                                hint.Append(",checking node [" + LocalName + "]");
                                if (CollectionGuid.ToLower() == LocalGuid)
                                {
                                    Return_CollectionPath = CollectionPath;
                                    Return_LastChagnedate = LastChangeDate;
                                    //Call AppendClassLogFile("Server", "GetCollectionConfigArg", "GetLocalCollectionArgs, match found, CollectionName=" & LocalName & ", CollectionPath=" & CollectionPath & ", LastChangeDate=" & LastChangeDate)
                                    MatchFound = true;
                                    break;
                                }
                            }
                        }
                    }
                }
                if (!MatchFound)
                {
                    //Call AppendClassLogFile("Server", "GetCollectionConfigArg", "GetLocalCollectionArgs, no local collection match found, Hint=[" & Hint & "]")
                }
                //
            }
            catch
            {
                HandleClassTrapError("GetLocalCollectionArgs");
            }
        }
        //
        //
        //
        public string GetConfig()
        {
            try
            {
                //
                string LocalFilename = "";
                string AddonPath = "";
                int Pos = 0;
                //
                AddonPath = Strings.Replace(ReflectionHelper.Invoke<string>(Main, "PhysicalccLibPath", new object[] { }), "\\cclib", "\\addons", 1, -1, CompareMethod.Text);
                AddonPath = AddonPath + "\\Collections.xml";
                //
                return ReflectionHelper.Invoke<string>(Main, "ReadFile", new object[] { AddonPath });
            }
            catch
            {
                HandleClassTrapError("GetConfig");
            }
            return "";
        }
        //
        //
        //
        private string AddCompatibilityResources(ref string CollectionPath, string ArchiveFilename, string SubPath, ref string Return_Version40DLLList)
        {
            try
            {
                //
                string AddFilename = "";
                string FileExt = "";
                string FileList = "";
                string[] Files = null;
                string Filename = "";
                string[] FileArgs = null;
                StringBuilder s = new StringBuilder();
                //Dim Remote As Object
                string FolderList = "";
                string[] Folders = null;
                string[] FolderArgs = null;
                string Folder = "";
                int Pos = 0;
                //
                // Process all SubPaths
                //
                FolderList = ReflectionHelper.Invoke<string>(Main, "GetFolderList", new object[] { CollectionPath + SubPath });
                if (FolderList != "")
                {
                    Folders = (string[])FolderList.Split(Environment.NewLine[0]);
                    foreach (string Folders_item in Folders)
                    {
                        Folder = Folders_item;
                        if (Folder != "")
                        {
                            FolderArgs = (string[])Folders_item.Split(',');
                            Folder = FolderArgs[0];
                            if (Folder != "")
                            {
                                s.Append(AddCompatibilityResources(ref CollectionPath, ArchiveFilename, SubPath + Folder + "\\", ref Return_Version40DLLList));
                            }
                        }
                    }
                }
                //
                // Process files in this path
                //
                //Set Remote = CreateObject("ccRemote.RemoteClass")
                FileList = ReflectionHelper.Invoke<string>(Main, "GetFileList", new object[] { CollectionPath });
                if (FileList != "")
                {
                    Files = (string[])FileList.Split(Environment.NewLine[0]);
                    for (int Ptr = 0; Ptr <= Files.GetUpperBound(0); Ptr++)
                    {
                        Filename = Files[Ptr];
                        if (Filename != "")
                        {
                            FileArgs = (string[])Filename.Split(',');
                            if (FileArgs.GetUpperBound(0) > 0)
                            {
                                Filename = FileArgs[0];
                                Pos = Filename.LastIndexOf(".") + 1;
                                FileExt = "";
                                if (Pos > 0)
                                {
                                    FileExt = Filename.Substring(Pos);
                                }
                                if (Filename.ToLower() == "collection.hlp")
                                {
                                    //
                                    // legacy help system, ignore this file
                                    //
                                }
                                else if (FileExt.ToLower() == "xml")
                                {
                                    //
                                    // compatibility resources can not include an xml file in the wwwroot
                                    //
                                }
                                else if ((CollectionPath.IndexOf("\\ContensiveFiles\\", StringComparison.CurrentCultureIgnoreCase) + 1) != 0)
                                {
                                    //
                                    // Content resources
                                    //
                                    s.Append("\r\n\t" + "<Resource name=\"" + cp.Utils.EncodeHTML(Filename) + "\" type=\"content\" path=\"" + cp.Utils.EncodeHTML(SubPath) + "\" />");
                                    AddFilename = CollectionPath + SubPath + "\\" + Filename;
                                    runAtServer("zipfile", "archive=" + kmaCommonModule.kmaEncodeRequestVariable(ArchiveFilename) + "&add=" + kmaCommonModule.kmaEncodeRequestVariable(AddFilename));
                                    //Call Remote.executeCmd("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable(AddFilename))
                                }
                                else if (FileExt.ToLower() == "dll")
                                {
                                    //
                                    // Executable resources
                                    //
                                    s.Append("\r\n\t" + "<Resource name=\"" + cp.Utils.EncodeHTML(Filename) + "\" type=\"executable\" path=\"" + cp.Utils.EncodeHTML(SubPath) + "\" />");
                                    Return_Version40DLLList = Return_Version40DLLList + Environment.NewLine + Filename;
                                    AddFilename = CollectionPath + SubPath + "\\" + Filename;
                                    runAtServer("zipfile", "archive=" + kmaCommonModule.kmaEncodeRequestVariable(ArchiveFilename) + "&add=" + kmaCommonModule.kmaEncodeRequestVariable(AddFilename));
                                    //Call Remote.executeCmd("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable(AddFilename))
                                }
                                else
                                {
                                    //
                                    // www resources
                                    //
                                    s.Append("\r\n\t" + "<Resource name=\"" + cp.Utils.EncodeHTML(Filename) + "\" type=\"www\" path=\"" + cp.Utils.EncodeHTML(SubPath) + "\" />");
                                    AddFilename = CollectionPath + SubPath + "\\" + Filename;
                                    runAtServer("zipfile", "archive=" + kmaCommonModule.kmaEncodeRequestVariable(ArchiveFilename) + "&add=" + kmaCommonModule.kmaEncodeRequestVariable(AddFilename));
                                    //Call Remote.executeCmd("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable(AddFilename))
                                }
                            }
                        }
                    }
                }
                //
                //
                return s.ToString();
            }
            catch
            {
                HandleClassTrapError("AddCompatibilityResources");
            }
            return "";
        }
        //
        // Encode a string into CData format
        //   <![CDATA[" & NodeContent & "]]>
        //
        private string EncodeCData(string Source)
        {
            string result = "";
            try
            {
                //
                result = Source;
                if (result != "")
                {
                    result = "<![CDATA[" + Strings.Replace(result, "]]>", "]]]]><![CDATA[>", 1, -1, CompareMethod.Binary) + "]]>";
                }
                //
            }
            catch
            {
                HandleClassTrapError("EncodeCData");
            }
            return result;
        }
        //
        //
        //
        private void runAtServer(string Cmd, string Arg)
        {
            //UPGRADE_TODO: (1065) Error handling statement (On Error Goto) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Goto Label (ErrorTrap)");
            //
            object runAtServer = null;
            //
            //UPGRADE_TODO: (1069) Error handling statement (On Error Resume Next) was converted to a pattern that might have a different behavior. More Information: http://www.vbtonet.com/ewis/ewi1069.aspx
            try
            {
                //UPGRADE_WARNING: (7008) The ProgId could not be found on computer where this application was migrated More Information: http://www.vbtonet.com/ewis/ewi7008.aspx
                runAtServer = Activator.CreateInstance(Type.GetTypeFromProgID("contensive.runAtServerClass"));
                //UPGRADE_WARNING: (2081) Err.Number has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2081.aspx
                if (Information.Err().Number != 0)
                {
                    //
                    // old object
                    //
                    runAtServer = new ccRemote.RemoteClass();
                }
                ReflectionHelper.Invoke(runAtServer, "executeCmd", new object[] { Cmd, Arg });
                //Call Remote.executeCmd("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable("@" & AddFileListFilename))
                //
                return;
            //
            // ----- Error Trap
            //
            ErrorTrap:
                HandleClassTrapError("runAtServer");
            }
            catch (Exception exc)
            {
                NotUpgradedHelper.NotifyNotUpgradedElement("Resume in On-Error-Resume-Next Block");
            }
        }
        //====================================================================================================
        private bool getYesNo( bool q)
        {
if ( q ) {
return "Yes";
} else {
return "No";
}

        }
    }
}
