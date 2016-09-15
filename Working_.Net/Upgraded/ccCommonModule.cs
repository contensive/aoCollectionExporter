#undef DebugBuild
using Microsoft.VisualBasic;
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using UpgradeHelpers.Helpers;

namespace aoCExport
{
	internal static class ccCommonModule
	{


		//
		//
		//=======================================================================
		//   sitepropertyNames
		//=======================================================================
		//
		public const string siteproperty_serverPageDefault_name = "serverPageDefault";
		public const string siteproperty_serverPageDefault_defaultValue = "index.php";
		//
		//=======================================================================
		//   content replacements
		//=======================================================================
		//
		public const string contentReplaceEscapeStart = "{%";
		public const string contentReplaceEscapeEnd = "%}";
		//
		public struct fieldEditorType
		{
			public int fieldId;
			public int addonid;
		}
		//
		public const string protectedContentSetControlFieldList = "ID,CREATEDBY,DATEADDED,MODIFIEDBY,MODIFIEDDATE,EDITSOURCEID,EDITARCHIVE,EDITBLANK,CONTENTCONTROLID";
		//Public Const protectedContentSetControlFieldList = "ID,CREATEDBY,DATEADDED,MODIFIEDBY,MODIFIEDDATE,EDITSOURCEID,EDITARCHIVE,EDITBLANK"
		//
		public const string HTMLEditorDefaultCopyStartMark = "<!-- cc -->";
		public const string HTMLEditorDefaultCopyEndMark = "<!-- /cc -->";
		public const string HTMLEditorDefaultCopyNoCr = HTMLEditorDefaultCopyStartMark + "<p><br></p>" + HTMLEditorDefaultCopyEndMark;
		public const string HTMLEditorDefaultCopyNoCr2 = "<p><br></p>";
		//
		public const string IconWidthHeight = " width=21 height=22 ";
		//Public Const IconWidthHeight = " width=18 height=22 "
		//
		public const string CoreCollectionGuid = "{8DAABAE6-8E45-4CEE-A42C-B02D180E799B}"; // contains core Contensive objects, loaded from Library
		public const string ApplicationCollectionGuid = "{C58A76E2-248B-4DE8-BF9C-849A960F79C6}"; // exported from application during upgrade
		//
		public const string adminCommonAddonGuid = "{76E7F79E-489F-4B0F-8EE5-0BAC3E4CD782}";
		public const string DashboardAddonGuid = "{4BA7B4A2-ED6C-46C5-9C7B-8CE251FC8FF5}";
		public const string PersonalizationGuid = "{C82CB8A6-D7B9-4288-97FF-934080F5FC9C}";
		public const string TextBoxGuid = "{7010002E-5371-41F7-9C77-0BBFF1F8B728}";
		public const string ContentBoxGuid = "{E341695F-C444-4E10-9295-9BEEC41874D8}";
		public const string DynamicMenuGuid = "{DB1821B3-F6E4-4766-A46E-48CA6C9E4C6E}";
		public const string ChildListGuid = "{D291F133-AB50-4640-9A9A-18DB68FF363B}";
		public const string DynamicFormGuid = "{8284FA0C-6C9D-43E1-9E57-8E9DD35D2DCC}";
		public const string AddonManagerGuid = "{1DC06F61-1837-419B-AF36-D5CC41E1C9FD}";
		public const string FormWizardGuid = "{2B1384C4-FD0E-4893-B3EA-11C48429382F}";
		public const string ImportWizardGuid = "{37F66F90-C0E0-4EAF-84B1-53E90A5B3B3F}";
		public const string JQueryGuid = "{9C882078-0DAC-48E3-AD4B-CF2AA230DF80}";
		public const string JQueryUIGuid = "{840B9AEF-9470-4599-BD47-7EC0C9298614}";
		public const string ImportProcessAddonGuid = "{5254FAC6-A7A6-4199-8599-0777CC014A13}";
		public const string StructuredDataProcessorGuid = "{65D58FE9-8B76-4490-A2BE-C863B372A6A4}";
		public const string jQueryFancyBoxGuid = "{24C2DBCF-3D84-44B6-A5F7-C2DE7EFCCE3D}";
		//
		public const string DefaultLandingPageGuid = "{925F4A57-32F7-44D9-9027-A91EF966FB0D}";
		public const string DefaultLandingSectionGuid = "{D882ED77-DB8F-4183-B12C-F83BD616E2E1}";
		public const string DefaultTemplateGuid = "{47BE95E4-5D21-42CC-9193-A343241E2513}";
		public const string DefaultDynamicMenuGuid = "{E8D575B9-54AE-4BF9-93B7-C7E7FE6F2DB3}";
		//
		public const string fpoContentBox = "{1571E62A-972A-4BFF-A161-5F6075720791}";
		//
		public const string sfImageExtList = "jpg,jpeg,gif,png";
		//
		public const string PageChildListInstanceID = "{ChildPageList}";
		//
		static readonly public string cr = Environment.NewLine + "\t";
		static readonly public string cr2 = cr + "\t";
		static readonly public string cr3 = cr2 + "\t";
		static readonly public string cr4 = cr3 + "\t";
		static readonly public string cr5 = cr4 + "\t";
		static readonly public string cr6 = cr5 + "\t";
		//
		static readonly public string AddonOptionConstructor_BlockNoAjax = "Wrapper=[Default:0|None:-1|ListID(Wrappers)]" + Environment.NewLine + "css Container id" + Environment.NewLine + "css Container class";
		static readonly public string AddonOptionConstructor_Block = "Wrapper=[Default:0|None:-1|ListID(Wrappers)]" + Environment.NewLine + "As Ajax=[If Add-on is Ajax:0|Yes:1]" + Environment.NewLine + "css Container id" + Environment.NewLine + "css Container class";
		static readonly public string AddonOptionConstructor_Inline = "As Ajax=[If Add-on is Ajax:0|Yes:1]" + Environment.NewLine + "css Container id" + Environment.NewLine + "css Container class";
		//
		// Constants used as arguments to SiteBuilderClass.CreateNewSite
		//
		public const int SiteTypeBaseAsp = 1;
		public const int sitetypebaseaspx = 2;
		public const int SiteTypeDemoAsp = 3;
		public const int SiteTypeBasePhp = 4;
		//
		//Public Const AddonNewParse = True
		//
		public const string AddonOptionConstructor_ForBlockText = "AllowGroups=[listid(groups)]checkbox";
		public const string AddonOptionConstructor_ForBlockTextEnd = "";
		public const string BlockTextStartMarker = "<!-- BLOCKTEXTSTART -->";
		public const string BlockTextEndMarker = "<!-- BLOCKTEXTEND -->";
		//
		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static void Sleep(int dwMilliseconds);
		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static int GetExitCodeProcess(int hProcess, ref int lpExitCode);
		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static int timeGetTime();
		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static int OpenProcess(int dwDesiredAccess, int bInheritHandle, int dwProcessId);
		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static int CloseHandle(int hObject);
		//
		public const string InstallFolderName = "Install";
		public const string DownloadFileRootNode = "collectiondownload";
		public const string CollectionFileRootNode = "collection";
		public const string CollectionFileRootNodeOld = "addoncollection";
		public const string CollectionListRootNode = "collectionlist";
		//
		public const string LegacyLandingPageName = "Landing Page Content";
		public const string DefaultNewLandingPageName = "Home";
		public const string DefaultLandingSectionName = "Home";
		//
		// ----- Errors Specific to the Contensive Objects
		//
		static readonly public int KmaccErrorUpgrading = kmaCommonModule.KmaObjectError + 1;
		static readonly public int KmaccErrorServiceStopped = kmaCommonModule.KmaObjectError + 2;
		//
		public const string UserErrorHeadline = "<p class=\"ccError\">There was a problem with this page.</p>";
		//
		// ----- Errors connecting to server
		//
		public const int ccError_InvalidAppName = 100;
		public const int ccError_ErrorAddingApp = 101;
		public const int ccError_ErrorDeletingApp = 102;
		public const int ccError_InvalidFieldName = 103; // Invalid parameter name
		public const int ccError_InvalidCommand = 104;
		public const int ccError_InvalidAuthentication = 105;
		public const int ccError_NotConnected = 106; // Attempt to execute a command without a connection
		//
		//
		//
		static readonly public int ccStatusCode_Base = kmaCommonModule.KmaErrorBase;
		static readonly public int ccStatusCode_ControllerCreateFailed = ccStatusCode_Base + 1;
		static readonly public int ccStatusCode_ControllerInProcess = ccStatusCode_Base + 2;
		static readonly public int ccStatusCode_ControllerStartedWithoutService = ccStatusCode_Base + 3;
		//
		// ----- Previous errors, can be replaced
		//
		//Public Const KmaError_UnderlyingObject_Msg = "An error occurred in an underlying routine."
		public const string KmaccErrorServiceStopped_Msg = "The Contensive CSv Service is not running.";
		public const string KmaError_BadObject_Msg = "Server Object is not valid.";
		public const string KmaError_UpgradeInProgress_Msg = "Server is busy with internal upgrade.";
		//
		//Public Const KmaError_InvalidArgument_Msg = "Invalid Argument"
		//Public Const KmaError_UnderlyingObject_Msg = "An error occurred in an underlying routine."
		//Public Const KmaccErrorServiceStopped_Msg = "The Contensive CSv Service is not running."
		//Public Const KmaError_BadObject_Msg = "Server Object is not valid."
		//Public Const KmaError_UpgradeInProgress_Msg = "Server is busy with internal upgrade."
		//Public Const KmaError_InvalidArgument_Msg = "Invalid Argument"
		//
		//-----------------------------------------------------------------------
		//   GetApplicationList indexes
		//-----------------------------------------------------------------------
		//
		public const int AppList_Name = 0;
		public const int AppList_Status = 1;
		public const int AppList_ConnectionsActive = 2;
		public const int AppList_ConnectionString = 3;
		public const int AppList_DataBuildVersion = 4;
		public const int AppList_LicenseKey = 5;
		public const int AppList_RootPath = 6;
		public const int AppList_PhysicalFilePath = 7;
		public const int AppList_DomainName = 8;
		public const int AppList_DefaultPage = 9;
		public const int AppList_AllowSiteMonitor = 10;
		public const int AppList_HitCounter = 11;
		public const int AppList_ErrorCount = 12;
		public const int AppList_DateStarted = 13;
		public const int AppList_AutoStart = 14;
		public const int AppList_Progress = 15;
		public const int AppList_PhysicalWWWPath = 16;
		public const int AppListCount = 17;
		//
		//-----------------------------------------------------------------------
		//   System MemberID - when the system does an update, it uses this member
		//-----------------------------------------------------------------------
		//
		public const int SystemMemberID = 0;
		//
		//-----------------------------------------------------------------------
		// ----- old (OptionKeys for available Options)
		//-----------------------------------------------------------------------
		//
		public const int OptionKeyProductionLicense = 0;
		public const int OptionKeyDeveloperLicense = 1;
		//
		//-----------------------------------------------------------------------
		// ----- LicenseTypes, replaced OptionKeys
		//-----------------------------------------------------------------------
		//
		public const int LicenseTypeInvalid = -1;
		public const int LicenseTypeProduction = 0;
		public const int LicenseTypeTrial = 1;
		//
		//-----------------------------------------------------------------------
		// ----- Active Content Definitions
		//-----------------------------------------------------------------------
		//
		public const string ACTypeDate = "DATE";
		public const string ACTypeVisit = "VISIT";
		public const string ACTypeVisitor = "VISITOR";
		public const string ACTypeMember = "MEMBER";
		public const string ACTypeOrganization = "ORGANIZATION";
		public const string ACTypeChildList = "CHILDLIST";
		public const string ACTypeContact = "CONTACT";
		public const string ACTypeFeedback = "FEEDBACK";
		public const string ACTypeLanguage = "LANGUAGE";
		public const string ACTypeAggregateFunction = "AGGREGATEFUNCTION";
		public const string ACTypeAddon = "ADDON";
		public const string ACTypeImage = "IMAGE";
		public const string ACTypeDownload = "DOWNLOAD";
		public const string ACTypeEnd = "END";
		public const string ACTypeTemplateContent = "CONTENT";
		public const string ACTypeTemplateText = "TEXT";
		public const string ACTypeDynamicMenu = "DYNAMICMENU";
		public const string ACTypeWatchList = "WATCHLIST";
		public const string ACTypeRSSLink = "RSSLINK";
		public const string ACTypePersonalization = "PERSONALIZATION";
		public const string ACTypeDynamicForm = "DYNAMICFORM";
		//
		public const string ACTagEnd = "<ac type=\"" + ACTypeEnd + "\">";
		//
		// ----- PropertyType Definitions
		//
		public const int PropertyTypeMember = 0;
		public const int PropertyTypeVisit = 1;
		public const int PropertyTypeVisitor = 2;
		//
		//-----------------------------------------------------------------------
		// ----- Port Assignments
		//-----------------------------------------------------------------------
		//
		public const int WinsockPortWebOut = 4000;
		public const int WinsockPortServerFromWeb = 4001;
		public const int WinsockPortServerToClient = 4002;
		//
		public const int Port_ContentServerControlDefault = 4531;
		public const int Port_SiteMonitorDefault = 4532;
		//
		public const int RMBMethodHandShake = 1;
		public const int RMBMethodMessage = 3;
		public const int RMBMethodTestPoint = 4;
		public const int RMBMethodInit = 5;
		public const int RMBMethodClosePage = 6;
		public const int RMBMethodOpenCSContent = 7;
		//
		// ----- Position equates for the Remote Method Block
		//
		const int RMBPositionLength = 0; // Length of the RMB
		const int RMBPositionSourceHandle = 4; // Handle generated by the source of the command
		const int RMBPositionMethod = 8; // Method in the method block
		const int RMBPositionArgumentCount = 12; // The number of arguments in the Block
		const int RMBPositionFirstArgument = 16; // The offset to the first argu
		//
		//-----------------------------------------------------------------------
		//   Remote Connections
		//   List of current remove connections for Remote Monitoring/administration
		//-----------------------------------------------------------------------
		//
		public struct RemoteAdministratorType
		{
			public string RemoteIP;
			public int RemotePort;
			public static RemoteAdministratorType CreateInstance()
			{
				RemoteAdministratorType result = new RemoteAdministratorType();
				result.RemoteIP = String.Empty;
				return result;
			}
		}
		//
		// Default username/password
		//
		public const string DefaultServerUsername = "root";
		public const string DefaultServerPassword = "contensive";
		//
		//-----------------------------------------------------------------------
		//   Form Contension Strategy
		//
		//       all Contensive Forms contain a hidden "ccFormSN"
		//       The value in the hidden is the FormID string. All input
		//       elements of the form are named FormID & "ElementName"
		//
		//       This prevents form elements from different forms from interfearing
		//       with each other, and with developer generated forms.
		//
		//       GetFormSN gets a new valid random formid to be used.
		//       All forms requires:
		//           a FormId (text), containing the formid string
		//           a [formid]Type (text), as defined in FormTypexxx in CommonModule
		//
		//       Forms have two primary sections: GetForm and ProcessForm
		//
		//       Any form that has a GetForm method, should have the process form
		//       in the main.init, selected with this [formid]type hidden (not the
		//       GetForm method). This is so the process can alter the stream
		//       output for areas before the GetForm call.
		//
		//       System forms, like tools panel, that may appear on any page, have
		//       their process call in the main.init.
		//
		//       Popup forms, like ImageSelector have their processform call in the
		//       main.init because no .asp page exists that might contain a call
		//       the process section.
		//
		//-----------------------------------------------------------------------
		//
		public const string FormTypeToolsPanel = "do30a8vl29";
		public const string FormTypeActiveEditor = "l1gk70al9n";
		public const string FormTypeImageSelector = "ila9c5s01m";
		public const string FormTypePageAuthoring = "2s09lmpalb";
		public const string FormTypeMyProfile = "89aLi180j5";
		public const string FormTypeLogin = "login";
		//Public Const FormTypeLogin = "l09H58a195"
		public const string FormTypeSendPassword = "lk0q56am09";
		public const string FormTypeJoin = "6df38abv00";
		public const string FormTypeHelpBubbleEditor = "9df019d77sA";
		public const string FormTypeAddonSettingsEditor = "4ed923aFGw9d";
		public const string FormTypeAddonStyleEditor = "ar5028jklkfd0s";
		public const string FormTypeSiteStyleEditor = "fjkq4w8794kdvse";
		//Public Const FormTypeAggregateFunctionProperties = "9wI751270"
		//
		//-----------------------------------------------------------------------
		//   Hardcoded profile form const
		//-----------------------------------------------------------------------
		//
		public const string rnMyProfileTopics = "profileTopics";
		//
		//-----------------------------------------------------------------------
		// Legacy - replaced with HardCodedPages
		//   Intercept Page Strategy
		//
		//       RequestnameInterceptpage = InterceptPage number from the input stream
		//       InterceptPage = Global variant with RequestnameInterceptpage value read during early Init
		//
		//       Intercept pages are complete pages that appear instead of what
		//       the physical page calls.
		//-----------------------------------------------------------------------
		//
		public const string RequestNameInterceptpage = "ccIPage";
		//
		public const string LegacyInterceptPageSNResourceLibrary = "s033l8dm15";
		public const string LegacyInterceptPageSNSiteExplorer = "kdif3318sd";
		public const string LegacyInterceptPageSNImageUpload = "ka983lm039";
		public const string LegacyInterceptPageSNMyProfile = "k09ddk9105";
		public const string LegacyInterceptPageSNLogin = "6ge42an09a";
		public const string LegacyInterceptPageSNPrinterVersion = "l6d09a10sP";
		public const string LegacyInterceptPageSNUploadEditor = "k0hxp2aiOZ";
		//
		//-----------------------------------------------------------------------
		// Ajax functions intercepted during init, answered and response closed
		//   These are hard-coded internal Contensive functions
		//   These should eventually be replaced with (HardcodedAddons) remote methods
		//   They should all be prefixed "cc"
		//   They are called with cj.ajax.qs(), setting RequestNameAjaxFunction=name in the qs
		//   These name=value pairs go in the QueryString argument of the javascript cj.ajax.qs() function
		//-----------------------------------------------------------------------
		//
		//Public Const RequestNameOpenSettingPage = "settingpageid"
		public const string RequestNameAjaxFunction = "ajaxfn";
		public const string RequestNameAjaxFastFunction = "ajaxfastfn";
		//
		public const string AjaxOpenAdminNav = "aps89102kd";
		public const string AjaxOpenAdminNavGetContent = "d8475jkdmfj2";
		public const string AjaxCloseAdminNav = "3857fdjdskf91";
		public const string AjaxAdminNavOpenNode = "8395j2hf6jdjf";
		public const string AjaxAdminNavOpenNodeGetContent = "eieofdwl34efvclaeoi234598";
		public const string AjaxAdminNavCloseNode = "w325gfd73fhdf4rgcvjk2";
		//
		public const string AjaxCloseIndexFilter = "k48smckdhorle0";
		public const string AjaxOpenIndexFilter = "Ls8jCDt87kpU45YH";
		public const string AjaxOpenIndexFilterGetContent = "llL98bbJQ38JC0KJm";
		public const string AjaxStyleEditorAddStyle = "ajaxstyleeditoradd";
		public const string AjaxPing = "ajaxalive";
		public const string AjaxGetFormEditTabContent = "ajaxgetformedittabcontent";
		public const string AjaxData = "data";
		public const string AjaxGetVisitProperty = "getvisitproperty";
		public const string AjaxSetVisitProperty = "setvisitproperty";
		public const string AjaxGetDefaultAddonOptionString = "ccGetDefaultAddonOptionString";
		public const string ajaxGetFieldEditorPreferenceForm = "ajaxgetfieldeditorpreference";
		//
		//-----------------------------------------------------------------------
		//
		// no - for now just use ajaxfn in the cj.ajax.qs call
		//   this is more work, and I do not see why it buys anything new or better
		//
		//   Hard-coded addons
		//       these are internal Contensive functions
		//       can be called with just /addonname?querystring
		//       call them with cj.ajax.addon() or cj.ajax.addonCallback()
		//       are first in the list of checks when a URL rewrite is detected in Init()
		//       should all be prefixed with 'cc'
		//-----------------------------------------------------------------------
		//
		//Public Const HardcodedAddonGetDefaultAddonOptionString = "ccGetDefaultAddonOptionString"
		//
		//-----------------------------------------------------------------------
		//   Remote Methods
		//       ?RemoteMethodAddon=string
		//       calls an addon (if marked to run as a remote method)
		//       blocks all other Contensive output (tools panel, javascript, etc)
		//-----------------------------------------------------------------------
		//
		public const string RequestNameRemoteMethodAddon = "remotemethodaddon";
		//
		//-----------------------------------------------------------------------
		//   Hard Coded Pages
		//       ?Method=string
		//       Querystring based so they can be added to URLs, preserving the current page for a return
		//       replaces output stream with html output
		//-----------------------------------------------------------------------
		//
		public const string RequestNameHardCodedPage = "method";
		//
		public const string HardCodedPageLogin = "login";
		public const string HardCodedPageLoginDefault = "logindefault";
		public const string HardCodedPageMyProfile = "myprofile";
		public const string HardCodedPagePrinterVersion = "printerversion";
		public const string HardCodedPageResourceLibrary = "resourcelibrary";
		public const string HardCodedPageLogoutLogin = "logoutlogin";
		public const string HardCodedPageLogout = "logout";
		public const string HardCodedPageSiteExplorer = "siteexplorer";
		//Public Const HardCodedPageForceMobile = "forcemobile"
		//Public Const HardCodedPageForceNonMobile = "forcenonmobile"
		public const string HardCodedPageNewOrder = "neworderpage";
		public const string HardCodedPageStatus = "status";
		public const string HardCodedPageGetJSPage = "getjspage";
		public const string HardCodedPageGetJSLogin = "getjslogin";
		public const string HardCodedPageRedirect = "redirect";
		public const string HardCodedPageExportAscii = "exportascii";
		public const string HardCodedPagePayPalConfirm = "paypalconfirm";
		public const string HardCodedPageSendPassword = "sendpassword";
		//
		//-----------------------------------------------------------------------
		//   Option values
		//       does not effect output directly
		//-----------------------------------------------------------------------
		//
		public const string RequestNamePageOptions = "ccoptions";
		//
		public const string PageOptionForceMobile = "forcemobile";
		public const string PageOptionForceNonMobile = "forcenonmobile";
		public const string PageOptionLogout = "logout";
		public const string PageOptionPrinterVersion = "printerversion";
		//
		// convert to options later
		//
		public const string RequestNameDashboardReset = "ResetDashboard";
		//
		//-----------------------------------------------------------------------
		//   DataSource constants
		//-----------------------------------------------------------------------
		//
		public const int DefaultDataSourceID = -1;
		//
		//-----------------------------------------------------------------------
		// ----- Type compatibility between databases
		//       Boolean
		//           Access      YesNo       true=1, false=0
		//           SQL Server  bit         true=1, false=0
		//           MySQL       bit         true=1, false=0
		//           Oracle      integer(1)  true=1, false=0
		//           Note: false does not equal NOT true
		//       Integer (Number)
		//           Access      Long        8 bytes, about E308
		//           SQL Server  int
		//           MySQL       integer
		//           Oracle      integer(8)
		//       Float
		//           Access      Double      8 bytes, about E308
		//           SQL Server  Float
		//           MySQL
		//           Oracle
		//       Text
		//           Access
		//           SQL Server
		//           MySQL
		//           Oracle
		//-----------------------------------------------------------------------
		//
		//Public Const SQLFalse = "0"
		//Public Const SQLTrue = "1"
		//
		//-----------------------------------------------------------------------
		// ----- Style sheet definitions
		//-----------------------------------------------------------------------
		//
		public const string defaultStyleFilename = "ccDefault.r5.css";
		public const string StyleSheetStart = "<STYLE TYPE=\"text/css\">";
		public const string StyleSheetEnd = "</STYLE>";
		//
		public const string SpanClassAdminNormal = "<span class=\"ccAdminNormal\">";
		public const string SpanClassAdminSmall = "<span class=\"ccAdminSmall\">";
		//
		// remove these from ccWebx
		//
		public const string SpanClassNormal = "<span class=\"ccNormal\">";
		public const string SpanClassSmall = "<span class=\"ccSmall\">";
		public const string SpanClassLarge = "<span class=\"ccLarge\">";
		public const string SpanClassHeadline = "<span class=\"ccHeadline\">";
		public const string SpanClassList = "<span class=\"ccList\">";
		public const string SpanClassListCopy = "<span class=\"ccListCopy\">";
		public const string SpanClassError = "<span class=\"ccError\">";
		public const string SpanClassSeeAlso = "<span class=\"ccSeeAlso\">";
		public const string SpanClassEnd = "</span>";
		//
		//-----------------------------------------------------------------------
		// ----- XHTML definitions
		//-----------------------------------------------------------------------
		//
		public const string DTDTransitional = "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">";
		//
		public const string BR = "<br>";
		//
		//-----------------------------------------------------------------------
		// AuthoringControl Types
		//-----------------------------------------------------------------------
		//
		public const int AuthoringControlsEditing = 1;
		public const int AuthoringControlsSubmitted = 2;
		public const int AuthoringControlsApproved = 3;
		public const int AuthoringControlsModified = 4;
		//
		//-----------------------------------------------------------------------
		// ----- Panel and header colors
		//-----------------------------------------------------------------------
		//
		//Public Const "ccPanel" = "#E0E0E0"    ' The background color of a panel (black copy visible on it)
		//Public Const "ccPanelHilite" = "#F8F8F8"  '
		//Public Const "ccPanelShadow" = "#808080"  '
		//
		//Public Const HeaderColorBase = "#0320B0"   ' The background color of a panel header (reverse copy visible)
		//Public Const "ccPanelHeaderHilite" = "#8080FF" '
		//Public Const "ccPanelHeaderShadow" = "#000000" '
		//
		//-----------------------------------------------------------------------
		// ----- Field type Definitions
		//       Field Types are numeric values that describe how to treat values
		//       stored as ContentFieldDefinitionType (FieldType property of FieldType Type.. ;)
		//-----------------------------------------------------------------------
		//
		public const int FieldTypeInteger = 1; // An long number
		public const int FieldTypeText = 2; // A text field (up to 255 characters)
		public const int FieldTypeLongText = 3; // A memo field (up to 8000 characters)
		public const int FieldTypeBoolean = 4; // A yes/no field
		public const int FieldTypeDate = 5; // A date field
		public const int FieldTypeFile = 6; // A filename of a file in the files directory.
		public const int FieldTypeLookup = 7; // A lookup is a FieldTypeInteger that indexes into another table
		public const int FieldTypeRedirect = 8; // creates a link to another section
		public const int FieldTypeCurrency = 9; // A Float that prints in dollars
		public const int FieldTypeTextFile = 10; // Text saved in a file in the files area.
		public const int FieldTypeImage = 11; // A filename of a file in the files directory.
		public const int FieldTypeFloat = 12; // A float number
		public const int FieldTypeAutoIncrement = 13; //long that automatically increments with the new record
		public const int FieldTypeManyToMany = 14; // no database field - sets up a relationship through a Rule table to another table
		public const int FieldTypeMemberSelect = 15; // This ID is a ccMembers record in a group defined by the MemberSelectGroupID field
		public const int FieldTypeCSSFile = 16; // A filename of a CSS compatible file
		public const int FieldTypeXMLFile = 17; // the filename of an XML compatible file
		public const int FieldTypeJavascriptFile = 18; // the filename of a javascript compatible file
		public const int FieldTypeLink = 19; // Links used in href tags -- can go to pages or resources
		public const int FieldTypeResourceLink = 20; // Links used in resources, link <img or <object. Should not be pages
		public const int FieldTypeHTML = 21; // LongText field that expects HTML content
		public const int FieldTypeHTMLFile = 22; // TextFile field that expects HTML content
		public const int FieldTypeMax = 22;
		//
		// ----- Field Descriptors for these type
		//       These are what are publicly displayed for each type
		//       See GetFieldDescriptorByType and vise-versa to translater
		//
		public const string FieldDescriptorInteger = "Integer";
		public const string FieldDescriptorText = "Text";
		public const string FieldDescriptorLongText = "LongText";
		public const string FieldDescriptorBoolean = "Boolean";
		public const string FieldDescriptorDate = "Date";
		public const string FieldDescriptorFile = "File";
		public const string FieldDescriptorLookup = "Lookup";
		public const string FieldDescriptorRedirect = "Redirect";
		public const string FieldDescriptorCurrency = "Currency";
		public const string FieldDescriptorImage = "Image";
		public const string FieldDescriptorFloat = "Float";
		public const string FieldDescriptorManyToMany = "ManyToMany";
		public const string FieldDescriptorTextFile = "TextFile";
		public const string FieldDescriptorCSSFile = "CSSFile";
		public const string FieldDescriptorXMLFile = "XMLFile";
		public const string FieldDescriptorJavascriptFile = "JavascriptFile";
		public const string FieldDescriptorLink = "Link";
		public const string FieldDescriptorResourceLink = "ResourceLink";
		public const string FieldDescriptorMemberSelect = "MemberSelect";
		public const string FieldDescriptorHTML = "HTML";
		public const string FieldDescriptorHTMLFile = "HTMLFile";
		//
		public const string FieldDescriptorLcaseInteger = "integer";
		public const string FieldDescriptorLcaseText = "text";
		public const string FieldDescriptorLcaseLongText = "longtext";
		public const string FieldDescriptorLcaseBoolean = "boolean";
		public const string FieldDescriptorLcaseDate = "date";
		public const string FieldDescriptorLcaseFile = "file";
		public const string FieldDescriptorLcaseLookup = "lookup";
		public const string FieldDescriptorLcaseRedirect = "redirect";
		public const string FieldDescriptorLcaseCurrency = "currency";
		public const string FieldDescriptorLcaseImage = "image";
		public const string FieldDescriptorLcaseFloat = "float";
		public const string FieldDescriptorLcaseManyToMany = "manytomany";
		public const string FieldDescriptorLcaseTextFile = "textfile";
		public const string FieldDescriptorLcaseCSSFile = "cssfile";
		public const string FieldDescriptorLcaseXMLFile = "xmlfile";
		public const string FieldDescriptorLcaseJavascriptFile = "javascriptfile";
		public const string FieldDescriptorLcaseLink = "link";
		public const string FieldDescriptorLcaseResourceLink = "resourcelink";
		public const string FieldDescriptorLcaseMemberSelect = "memberselect";
		public const string FieldDescriptorLcaseHTML = "html";
		public const string FieldDescriptorLcaseHTMLFile = "htmlfile";
		//
		//------------------------------------------------------------------------
		// ----- Payment Options
		//------------------------------------------------------------------------
		//
		public const int PayTypeCreditCardOnline = 0; // Pay by credit card online
		public const int PayTypeCreditCardByPhone = 1; // Phone in a credit card
		public const int PayTypeCreditCardByFax = 9; // Phone in a credit card
		public const int PayTypeCHECK = 2; // pay by check to be mailed
		public const int PayTypeCREDIT = 3; // pay on account
		public const int PayTypeNONE = 4; // order total is $0.00. Nothing due
		public const int PayTypeCHECKCOMPANY = 5; // pay by company check
		public const int PayTypeNetTerms = 6;
		public const int PayTypeCODCompanyCheck = 7;
		public const int PayTypeCODCertifiedFunds = 8;
		public const int PayTypePAYPAL = 10;
		public const int PAYDEFAULT = 0;
		//
		//------------------------------------------------------------------------
		// ----- Credit card options
		//------------------------------------------------------------------------
		//
		public const int CCTYPEVISA = 0; // Visa
		public const int CCTYPEMC = 1; // MasterCard
		public const int CCTYPEAMEX = 2; // American Express
		public const int CCTYPEDISCOVER = 3; // Discover
		public const int CCTYPENOVUS = 4; // Novus Card
		public const int CCTYPEDEFAULT = 0;
		//
		//------------------------------------------------------------------------
		// ----- Shipping Options
		//------------------------------------------------------------------------
		//
		public const int SHIPGROUND = 0; // ground
		public const int SHIPOVERNIGHT = 1; // overnight
		public const int SHIPSTANDARD = 2; // standard, whatever that is
		public const int SHIPOVERSEAS = 3; // to overseas
		public const int SHIPCANADA = 4; // to Canada
		public const int SHIPDEFAULT = 0;
		//
		//------------------------------------------------------------------------
		// Debugging info
		//------------------------------------------------------------------------
		//
		public const int TestPointTab = 2;
		public const string TestPointTabChr = "-";
		public static double CPTickCountBase = 0;
		//
		//------------------------------------------------------------------------
		//   project width button defintions
		//------------------------------------------------------------------------
		//
		public const string ButtonApply = "  Apply ";
		public const string ButtonLogin = "  Login  ";
		public const string ButtonLogout = "  Logout  ";
		public const string ButtonSendPassword = "  Send Password  ";
		public const string ButtonJoin = "   Join   ";
		public const string ButtonSave = "  Save  ";
		public const string ButtonOK = "     OK     ";
		public const string ButtonReset = "  Reset  ";
		public const string ButtonSaveAddNew = " Save + Add ";
		//Public Const ButtonSaveAddNew = " Save > Add "
		public const string ButtonCancel = " Cancel ";
		public const string ButtonRestartContensiveApplication = " Restart Contensive Application ";
		public const string ButtonCancelAll = "  Cancel  ";
		public const string ButtonFind = "   Find   ";
		public const string ButtonDelete = "  Delete  ";
		public const string ButtonDeletePerson = " Delete Person ";
		public const string ButtonDeleteRecord = " Delete Record ";
		public const string ButtonDeleteEmail = " Delete Email ";
		public const string ButtonDeletePage = " Delete Page ";
		public const string ButtonFileChange = "   Upload   ";
		public const string ButtonFileDelete = "    Delete    ";
		public const string ButtonClose = "  Close   ";
		public const string ButtonAdd = "   Add    ";
		public const string ButtonAddChildPage = " Add Child ";
		public const string ButtonAddSiblingPage = " Add Sibling ";
		public const string ButtonContinue = " Continue >> ";
		public const string ButtonBack = "  << Back  ";
		public const string ButtonNext = "   Next   ";
		public const string ButtonPrevious = " Previous ";
		public const string ButtonFirst = "  First   ";
		public const string ButtonSend = "  Send   ";
		public const string ButtonSendTest = "Send Test";
		public const string ButtonCreateDuplicate = " Create Duplicate ";
		public const string ButtonActivate = "  Activate   ";
		public const string ButtonDeactivate = "  Deactivate   ";
		public const string ButtonOpenActiveEditor = "Active Edit";
		public const string ButtonPublish = " Publish Changes ";
		public const string ButtonAbortEdit = " Abort Edits ";
		public const string ButtonPublishSubmit = " Submit for Publishing ";
		public const string ButtonPublishApprove = " Approve for Publishing ";
		public const string ButtonPublishDeny = " Deny for Publishing ";
		public const string ButtonWorkflowPublishApproved = " Publish Approved Records ";
		public const string ButtonWorkflowPublishSelected = " Publish Selected Records ";
		public const string ButtonSetHTMLEdit = " Edit WYSIWYG ";
		public const string ButtonSetTextEdit = " Edit HTML ";
		public const string ButtonRefresh = " Refresh ";
		public const string ButtonOrder = " Order ";
		public const string ButtonSearch = " Search ";
		public const string ButtonSpellCheck = " Spell Check ";
		public const string ButtonLibraryUpload = " Upload ";
		public const string ButtonCreateReport = " Create Report ";
		public const string ButtonClearTrapLog = " Clear Trap Log ";
		public const string ButtonNewSearch = " New Search ";
		public const string ButtonReloadCDef = " Reload Content Definitions ";
		public const string ButtonImportTemplates = " Import Templates ";
		public const string ButtonRSSRefresh = " Update RSS Feeds Now ";
		public const string ButtonRequestDownload = " Request Download ";
		public const string ButtonFinish = " Finish ";
		public const string ButtonRegister = " Register ";
		public const string ButtonBegin = "Begin";
		public const string ButtonAbort = "Abort";
		public const string ButtonCreateGUID = " Create GUID ";
		public const string ButtonEnable = " Enable ";
		public const string ButtonDisable = " Disable ";
		public const string ButtonMarkReviewed = " Mark Reviewed ";
		//
		//------------------------------------------------------------------------
		//   member actions
		//------------------------------------------------------------------------
		//
		public const int MemberActionNOP = 0;
		public const int MemberActionLogin = 1;
		public const int MemberActionLogout = 2;
		public const int MemberActionForceLogin = 3;
		public const int MemberActionSendPassword = 4;
		public const int MemberActionForceLogout = 5;
		public const int MemberActionToolsApply = 6;
		public const int MemberActionJoin = 7;
		public const int MemberActionSaveProfile = 8;
		public const int MemberActionEditProfile = 9;
		//
		//-----------------------------------------------------------------------
		// ----- note pad info
		//-----------------------------------------------------------------------
		//
		public const int NoteFormList = 1;
		public const int NoteFormRead = 2;
		//
		public const string NoteButtonPrevious = " Previous ";
		public const string NoteButtonNext = "   Next   ";
		public const string NoteButtonDelete = "  Delete  ";
		public const string NoteButtonClose = "  Close   ";
		//                       ' Submit button is created in CommonDim, so it is simple
		public const string NoteButtonSubmit = "Submit";
		//
		//-----------------------------------------------------------------------
		// ----- Admin site storage
		//-----------------------------------------------------------------------
		//
		public const int AdminMenuModeHidden = 0; //   menu is hidden
		public const int AdminMenuModeLeft = 1; //   menu on the left
		public const int AdminMenuModeTop = 2; //   menu as dropdowns from the top
		//
		// ----- AdminActions - describes the form processing to do
		//
		public const int AdminActionNop = 0; // do nothing
		public const int AdminActionDelete = 4; // delete record
		public const int AdminActionFind = 5; //
		public const int AdminActionDeleteFilex = 6; //
		public const int AdminActionUpload = 7; //
		public const int AdminActionSaveNormal = 3; // save fields to database
		public const int AdminActionSaveEmail = 8; // save email record (and update EmailGroups) to database
		public const int AdminActionSaveMember = 11; //
		public const int AdminActionSaveSystem = 12;
		public const int AdminActionSavePaths = 13; // Save a record that is in the BathBlocking Format
		public const int AdminActionSendEmail = 9; //
		public const int AdminActionSendEmailTest = 10; //
		public const int AdminActionNext = 14; //
		public const int AdminActionPrevious = 15; //
		public const int AdminActionFirst = 16; //
		public const int AdminActionSaveContent = 17; //
		public const int AdminActionSaveField = 18; // Save a single field, fieldname = fn input
		public const int AdminActionPublish = 19; // Publish record live
		public const int AdminActionAbortEdit = 20; // Publish record live
		public const int AdminActionPublishSubmit = 21; // Submit for Workflow Publishing
		public const int AdminActionPublishApprove = 22; // Approve for Workflow Publishing
		public const int AdminActionWorkflowPublishApproved = 23; // Publish what was approved
		public const int AdminActionSetHTMLEdit = 24; // Set Member Property for this field to HTML Edit
		public const int AdminActionSetTextEdit = 25; // Set Member Property for this field to Text Edit
		public const int AdminActionSave = 26; // Save Record
		public const int AdminActionActivateEmail = 27; // Activate a Conditional Email
		public const int AdminActionDeactivateEmail = 28; // Deactivate a conditional email
		public const int AdminActionDuplicate = 29; // Duplicate the (sent email) record
		public const int AdminActionDeleteRows = 30; // Delete from rows of records, row0 is boolean, rowid0 is ID, rowcnt is count
		public const int AdminActionSaveAddNew = 31; // Save Record and add a new record
		public const int AdminActionReloadCDef = 32; // Load Content Definitions
		public const int AdminActionWorkflowPublishSelected = 33; // Publish what was selected
		public const int AdminActionMarkReviewed = 34; // Mark the record reviewed without making any changes
		public const int AdminActionEditRefresh = 35; // reload the page just like a save, but do not save
		//
		// ----- Adminforms (0-99)
		//
		public const int AdminFormRoot = 0; // intro page
		public const int AdminFormIndex = 1; // record list page
		public const int AdminFormHelp = 2; // popup help window
		public const int AdminFormUpload = 3; // encoded file upload form
		public const int AdminFormEdit = 4; // Edit form for system format records
		public const int AdminFormEditSystem = 5; // Edit form for system format records
		public const int AdminFormEditNormal = 6; // record edit page
		public const int AdminFormEditEmail = 7; // Edit form for Email format records
		public const int AdminFormEditMember = 8; // Edit form for Member format records
		public const int AdminFormEditPaths = 9; // Edit form for Paths format records
		public const int AdminFormClose = 10; // Special Case - do a window close instead of displaying a form
		public const int AdminFormReports = 12; // Call Reports form (admin only)
		//Public Const AdminFormSpider = 13          ' Call Spider
		public const int AdminFormEditContent = 14; // Edit form for Content records
		public const int AdminFormDHTMLEdit = 15; // ActiveX DHTMLEdit form
		public const int AdminFormEditPageContent = 16; //
		public const int AdminFormPublishing = 17; // Workflow Authoring Publish Control form
		public const int AdminFormQuickStats = 18; // Quick Stats (from Admin root)
		public const int AdminFormResourceLibrary = 19; // Resource Library without Selects
		public const int AdminFormEDGControl = 20; // Control Form for the EDG publishing controls
		public const int AdminFormSpiderControl = 21; // Control Form for the Content Spider
		public const int AdminFormContentChildTool = 22; // Admin Create Content Child tool
		public const int AdminformPageContentMap = 23; // Map all content to a single map
		public const int AdminformHousekeepingControl = 24; // Housekeeping control
		public const int AdminFormCommerceControl = 25;
		public const int AdminFormContactManager = 26;
		public const int AdminFormStyleEditor = 27;
		public const int AdminFormEmailControl = 28;
		public const int AdminFormCommerceInterface = 29;
		public const int AdminFormDownloads = 30;
		public const int AdminformRSSControl = 31;
		public const int AdminFormMeetingSmart = 32;
		public const int AdminFormMemberSmart = 33;
		public const int AdminFormEmailWizard = 34;
		public const int AdminFormImportWizard = 35;
		public const int AdminFormCustomReports = 36;
		public const int AdminFormFormWizard = 37;
		public const int AdminFormLegacyAddonManager = 38;
		public const int AdminFormIndex_SubFormAdvancedSearch = 39;
		public const int AdminFormIndex_SubFormSetColumns = 40;
		public const int AdminFormPageControl = 41;
		public const int AdminFormSecurityControl = 42;
		public const int AdminFormEditorConfig = 43;
		public const int AdminFormBuilderCollection = 44;
		public const int AdminFormClearCache = 45;
		public const int AdminFormMobileBrowserControl = 46;
		public const int AdminFormMetaKeywordTool = 47;
		public const int AdminFormIndex_SubFormExport = 48;
		//
		// ----- AdminFormTools (11,100-199)
		//
		public const int AdminFormTools = 11; // Call Tools form (developer only)
		public const int AdminFormToolRoot = 11; // These should match for compatibility
		public const int AdminFormToolCreateContentDefinition = 101;
		public const int AdminFormToolContentTest = 102;
		public const int AdminFormToolConfigureMenu = 103;
		public const int AdminFormToolConfigureListing = 104;
		public const int AdminFormToolConfigureEdit = 105;
		public const int AdminFormToolManualQuery = 106;
		public const int AdminFormToolWriteUpdateMacro = 107;
		public const int AdminFormToolDuplicateContent = 108;
		public const int AdminFormToolDuplicateDataSource = 109;
		public const int AdminFormToolDefineContentFieldsFromTable = 110;
		public const int AdminFormToolContentDiagnostic = 111;
		public const int AdminFormToolCreateChildContent = 112;
		public const int AdminFormToolClearContentWatchLink = 113;
		public const int AdminFormToolSyncTables = 114;
		public const int AdminFormToolBenchmark = 115;
		public const int AdminFormToolSchema = 116;
		public const int AdminFormToolContentFileView = 117;
		public const int AdminFormToolDbIndex = 118;
		public const int AdminFormToolContentDbSchema = 119;
		public const int AdminFormToolLogFileView = 120;
		public const int AdminFormToolLoadCDef = 121;
		public const int AdminFormToolLoadTemplates = 122;
		public const int AdminformToolFindAndReplace = 123;
		public const int AdminformToolCreateGUID = 124;
		public const int AdminformToolIISReset = 125;
		public const int AdminFormToolRestart = 126;
		public const int AdminFormToolWebsiteFileView = 127;
		//
		// ----- Define the index column structure
		//       IndexColumnVariant( 0, n ) is the first column on the left
		//       IndexColumnVariant( 0, IndexColumnField ) = the index into the fields array
		//
		public const int IndexColumnField = 0; // The field displayed in the column
		public const int IndexColumnWIDTH = 1; // The width of the column
		public const int IndexColumnSORTPRIORITY = 2; // lowest columns sorts first
		public const int IndexColumnSORTDIRECTION = 3; // direction of the sort on this column
		public const int IndexColumnSATTRIBUTEMAX = 3; // the number of attributes here
		public const int IndexColumnsMax = 50;
		//
		// ----- ReportID Constants, moved to ccCommonModule
		//
		public const int ReportFormRoot = 1;
		public const int ReportFormDailyVisits = 2;
		public const int ReportFormWeeklyVisits = 12;
		public const int ReportFormSitePath = 4;
		public const int ReportFormSearchKeywords = 5;
		public const int ReportFormReferers = 6;
		public const int ReportFormBrowserList = 8;
		public const int ReportFormAddressList = 9;
		public const int ReportFormContentProperties = 14;
		public const int ReportFormSurveyList = 15;
		public const int ReportFormOrdersList = 13;
		public const int ReportFormOrderDetails = 21;
		public const int ReportFormVisitorList = 11;
		public const int ReportFormMemberDetails = 16;
		public const int ReportFormPageList = 10;
		public const int ReportFormVisitList = 3;
		public const int ReportFormVisitDetails = 17;
		public const int ReportFormVisitorDetails = 20;
		public const int ReportFormSpiderDocList = 22;
		public const int ReportFormSpiderErrorList = 23;
		public const int ReportFormEDGDocErrors = 24;
		public const int ReportFormDownloadLog = 25;
		public const int ReportFormSpiderDocDetails = 26;
		public const int ReportFormSurveyDetails = 27;
		public const int ReportFormEmailDropList = 28;
		public const int ReportFormPageTraffic = 29;
		public const int ReportFormPagePerformance = 30;
		public const int ReportFormEmailDropDetails = 31;
		public const int ReportFormEmailOpenDetails = 32;
		public const int ReportFormEmailClickDetails = 33;
		public const int ReportFormGroupList = 34;
		public const int ReportFormGroupMemberList = 35;
		public const int ReportFormTrapList = 36;
		public const int ReportFormCount = 36;
		//
		//=============================================================================
		// Page Scope Meetings Related Storage
		//=============================================================================
		//
		public const int MeetingFormIndex = 0;
		public const int MeetingFormAttendees = 1;
		public const int MeetingFormLinks = 2;
		public const int MeetingFormFacility = 3;
		public const int MeetingFormHotel = 4;
		public const int MeetingFormDetails = 5;
		//
		//------------------------------------------------------------------------------
		// Form actions
		//------------------------------------------------------------------------------
		//
		// ----- DataSource Types
		//
		public const int DataSourceTypeODBCSQL99 = 0;
		public const int DataSourceTypeODBCAccess = 1;
		public const int DataSourceTypeODBCSQLServer = 2;
		public const int DataSourceTypeODBCMySQL = 3;
		public const int DataSourceTypeXMLFile = 4; // Use MSXML Interface to open a file
		//
		//------------------------------------------------------------------------------
		//   Application Status
		//------------------------------------------------------------------------------
		//
		public const int ApplicationStatusNotFound = 0;
		public const int ApplicationStatusLoadedNotRunning = 1;
		public const int ApplicationStatusRunning = 2;
		public const int ApplicationStatusStarting = 3;
		public const int ApplicationStatusUpgrading = 4;
		// Public Const ApplicationStatusConnectionBusy = 5    ' can not open connection because already open
		public const int ApplicationStatusKernelFailure = 6; // can not create Kernel
		public const int ApplicationStatusNoHostService = 7; // host service process ID not set
		public const int ApplicationStatusLicenseFailure = 8; // failed to start because of License failure
		public const int ApplicationStatusDbFailure = 9; // failed to start because ccSetup table not found
		public const int ApplicationStatusUnknownFailure = 10; // failed to start because of unknown error, see trace log
		public const int ApplicationStatusDbBad = 11; // ccContent,ccFields no records found
		public const int ApplicationStatusConnectionObjectFailure = 12; // Connection Object FAiled
		public const int ApplicationStatusConnectionStringFailure = 13; // Connection String FAiled to open the ODBC connection
		public const int ApplicationStatusDataSourceFailure = 14; // DataSource failed to open
		public const int ApplicationStatusDuplicateDomains = 15; // Can not locate application because there are 1+ apps that match the domain
		public const int ApplicationStatusPaused = 16; // Running, but all activity is blocked (for backup)
		//
		// Document (HTML, graphic, etc) retrieved from site
		//
		public const int ResponseHeaderCountMax = 20;
		public const int ResponseCookieCountMax = 20;
		//
		// ----- text delimiter that divides the text and html parts of an email message stored in the queue folder
		//
		static readonly public string EmailTextHTMLDelimiter = Environment.NewLine + " ----- End Text Begin HTML -----" + Environment.NewLine;
		//
		//------------------------------------------------------------------------
		//   Common RequestName Variables
		//------------------------------------------------------------------------
		//
		public const string RequestNameDynamicFormID = "dformid";
		//
		public const string RequestNameRunAddon = "addonid";
		public const string RequestNameEditReferer = "EditReferer";
		public const string RequestNameRefreshBlock = "ccFormRefreshBlockSN";
		public const string RequestNameCatalogOrder = "CatalogOrderID";
		public const string RequestNameCatalogCategoryID = "CatalogCatID";
		public const string RequestNameCatalogForm = "CatalogFormID";
		public const string RequestNameCatalogItemID = "CatalogItemID";
		public const string RequestNameCatalogItemAge = "CatalogItemAge";
		public const string RequestNameCatalogRecordTop = "CatalogTop";
		public const string RequestNameCatalogFeatured = "CatalogFeatured";
		public const string RequestNameCatalogSpan = "CatalogSpan";
		public const string RequestNameCatalogKeywords = "CatalogKeywords";
		public const string RequestNameCatalogSource = "CatalogSource";
		//
		public const string RequestNameLibraryFileID = "fileEID";
		public const string RequestNameDownloadID = "downloadid";
		public const string RequestNameLibraryUpload = "LibraryUpload";
		public const string RequestNameLibraryName = "LibraryName";
		public const string RequestNameLibraryDescription = "LibraryDescription";

		public const string RequestNameTestHook = "CC"; // input request that sets debugging hooks

		public const string RequestNameRootPage = "RootPageName";
		public const string RequestNameRootPageID = "RootPageID";
		public const string RequestNameContent = "ContentName";
		public const string RequestNameOrderByClause = "OrderByClause";
		public const string RequestNameAllowChildPageList = "AllowChildPageList";
		//
		public const string RequestNameCRKey = "crkey";
		public const string RequestNameAdminForm = "af";
		public const string RequestNameAdminSubForm = "subform";
		public const string RequestNameButton = "button";
		public const string RequestNameAdminSourceForm = "asf";
		public const string RequestNameAdminFormSpelling = "SpellingRequest";
		public const string RequestNameInlineStyles = "InlineStyles";
		public const string RequestNameAllowCSSReset = "AllowCSSReset";
		//
		public const string RequestNameReportForm = "rid";
		//
		public const string RequestNameToolContentID = "ContentID";
		//
		public const string RequestNameCut = "a904o2pa0cut";
		public const string RequestNamePaste = "dp29a7dsa6paste";
		public const string RequestNamePasteParentContentID = "dp29a7dsa6cid";
		public const string RequestNamePasteParentRecordID = "dp29a7dsa6rid";
		public const string RequestNamePasteFieldList = "dp29a7dsa6key";
		public const string RequestNameCutClear = "dp29a7dsa6clear";
		//
		public const string RequestNameRequestBinary = "RequestBinary";
		// removed -- this was an old method of blocking form input for file uploads
		//Public Const RequestNameFormBlock = "RB"
		public const string RequestNameJSForm = "RequestJSForm";
		public const string RequestNameJSProcess = "ProcessJSForm";
		//
		public const string RequestNameFolderID = "FolderID";
		//
		public const string RequestNameEmailMemberID = "emi8s9Kj";
		public const string RequestNameEmailOpenFlag = "eof9as88";
		public const string RequestNameEmailOpenCssFlag = "8aa41pM3";
		public const string RequestNameEmailClickFlag = "ecf34Msi";
		public const string RequestNameEmailSpamFlag = "9dq8Nh61";
		public const string RequestNameEmailBlockRequestDropID = "BlockEmailRequest";
		public const string RequestNameVisitTracking = "s9lD1088";
		public const string RequestNameBlockContentTracking = "BlockContentTracking";
		public const string RequestNameCookieDetectVisitID = "f92vo2a8d";

		public const string RequestNamePageNumber = "PageNumber";
		public const string RequestNamePageSize = "PageSize";
		//
		public const string RequestValueNull = "[NULL]";
		//
		public const string SpellCheckUserDictionaryFilename = "SpellCheck\\UserDictionary.txt";
		//
		public const string RequestNameStateString = "vstate";
		//
		//------------------------------------------------------------------------------
		// name value pairs
		//------------------------------------------------------------------------------
		//
		public struct NameValuePairType
		{
			public string Name;
			public string Value;
			public static NameValuePairType CreateInstance()
			{
				NameValuePairType result = new NameValuePairType();
				result.Name = String.Empty;
				result.Value = String.Empty;
				return result;
			}
		}
		//'
		//' ----- ContentSetMirror Type
		//'       Used on the WebClient, not the CSv
		//'       Stores info about the ContentSet, and caches the current row
		//'
		//Public Type ContentSetMirrorType
		//    Open As Boolean                     ' If true, it is in use
		//    Updateable As Boolean               ' Can not update an OpenCSSQL because Fields are not accessable
		//    ContentName As String               ' If updateable, this is the contentname
		//    CSPointer As Long                ' CSPointer for this ContentSet
		//    '
		//    ' ----- a cache of the current row, passed in during open and nextrecord, back during save and nextrecord
		//    '
		//    EOF As Boolean                      ' if true, Row is empty and at end of records
		//    RowCache() As ContentSetRowCacheType ' array of fields buffered for this set
		//    RowCacheSize As Long             ' the total number of fields in the row
		//    RowCacheCount As Long            ' the number of field() values to write
		//    End Type
		//
		// ----- Dataset for graphing
		//
		public struct ColumnDataType
		{
			public string Name;
			public int[] row;
			public static ColumnDataType CreateInstance()
			{
				ColumnDataType result = new ColumnDataType();
				result.Name = String.Empty;
				return result;
			}
		}
		//
		public struct ChartDataType
		{
			public string Title;
			public string XLabel;
			public string YLabel;
			public int RowCount;
			public string[] RowLabel;
			public int ColumnCount;
			public ccCommonModule.ColumnDataType[] Column;
			public static ChartDataType CreateInstance()
			{
				ChartDataType result = new ChartDataType();
				result.Title = String.Empty;
				result.XLabel = String.Empty;
				result.YLabel = String.Empty;
				return result;
			}
		}
		//'
		// PrivateStorage to hold the DebugTimer
		//
		public struct TimerStackType
		{
			public string Label;
			public int StartTicks;
			public static TimerStackType CreateInstance()
			{
				TimerStackType result = new TimerStackType();
				result.Label = String.Empty;
				return result;
			}
		}
		private const int TimerStackMax = 20;
		private static ccCommonModule.TimerStackType[] TimerStack = ArraysHelper.InitializeArray<ccCommonModule.TimerStackType>(TimerStackMax + 1);
		private static int TimerStackCount = 0;
		//
		public const string TextSearchStartTagDefault = "<!--TextSearchStart-->";
		public const string TextSearchEndTagDefault = "<!--TextSearchEnd-->";
		//
		//-------------------------------------------------------------------------------------
		//   IPDaemon communication objects
		//-------------------------------------------------------------------------------------
		//
		public struct IPDaemonConnectionType
		{
			public short ConnectionID;
			public int BytesToSend;
			public string HTTPVersion;
			public string HTTPMethod;
			public string Path;
			public string Query;
			public string Headers;
			public string PostData;
			public bool SendData;
			public short State;
			public short ContentLength;
			public static IPDaemonConnectionType CreateInstance()
			{
				IPDaemonConnectionType result = new IPDaemonConnectionType();
				result.HTTPVersion = String.Empty;
				result.HTTPMethod = String.Empty;
				result.Path = String.Empty;
				result.Query = String.Empty;
				result.Headers = String.Empty;
				result.PostData = String.Empty;
				return result;
			}
		}

		public static ccCommonModule.IPDaemonConnectionType[] IPDaemonConnection = null;

		public const int IPDaemon_DISCONNECTED = 0;
		public const int IPDaemon_CONNECTED = 1;
		public const int IPDaemon_HEADERS = 2;
		public const int IPDaemon_POSTDATA = 3;
		public const int IPDaemon_SERVE = 4;
		public const int IPDaemon_SERVEDIR = 5;
		public const int IPDaemon_SERVEFILE = 6;
		//
		//-------------------------------------------------------------------------------------
		//   Email
		//-------------------------------------------------------------------------------------
		//
		public const int EmailLogTypeDrop = 1; // Email was dropped
		public const int EmailLogTypeOpen = 2; // System detected the email was opened
		public const int EmailLogTypeClick = 3; // System detected a click from a link on the email
		public const int EmailLogTypeBounce = 4; // Email was processed by bounce processing
		public const int EmailLogTypeBlockRequest = 5; // recipient asked us to stop sending email
		public const int EmailLogTypeImmediateSend = 6; // Email was dropped
		//
		public const string DefaultSpamFooter = "<p>To block future emails from this site, <link>click here</link></p>";
		//
		static readonly public string FeedbackFormNotSupportedComment = "<!--" + Environment.NewLine + "Feedback form is not supported in this context" + Environment.NewLine + "-->";
		//
		//-------------------------------------------------------------------------------------
		//   Page Content constants
		//-------------------------------------------------------------------------------------
		//
		public const string ContentBlockCopyName = "Content Block Copy";
		//
		public const string BubbleCopy_AdminAddPage = "Use the Add page to create new content records. The save button puts you in edit mode. The OK button creates the record and exits.";
		public const string BubbleCopy_AdminIndexPage = "Use the Admin Listing page to locate content records through the Admin Site.";
		public const string BubbleCopy_SpellCheckPage = "Use the Spell Check page to verify and correct spelling throught the content.";
		public const string BubbleCopy_AdminEditPage = "Use the Edit page to add and modify content.";
		//
		//
		public const string TemplateDefaultName = "Default";
		//Public Const TemplateDefaultBody = "<!--" & vbCrLf & "Default Template - edit this Page Template, or select a different template for your page or section" & vbCrLf & "-->{{DYNAMICMENU?MENU=}}<br>{{CONTENT}}"
		static readonly public string TemplateDefaultBody = "" + Environment.NewLine + "\t" + "<!--" + Environment.NewLine + "\t" + "Default Template - edit this Page Template, or select a different template for your page or section" + Environment.NewLine + "\t" + "-->" + Environment.NewLine + "\t" + "<ac type=\"AGGREGATEFUNCTION\" name=\"Dynamic Menu\" querystring=\"Menu Name=Default\" acinstanceid=\"{6CBADABB-5B0D-43E1-B3CA-46A3D60DA3E1}\" >" + Environment.NewLine + "\t" + "<ac type=\"AGGREGATEFUNCTION\" name=\"Content Box\" acinstanceid=\"{49E0D0C0-D323-49B6-B211-B9599673A265}\" >";
		public const string TemplateDefaultBodyTag = "<body class=\"ccBodyWeb\">";
		//
		//=======================================================================
		//   Internal Tab interface storage
		//=======================================================================
		//
		private struct TabType
		{
			public string Caption;
			public string Link;
			public string StylePrefix;
			public bool IsHit;
			public string LiveBody;
			public static TabType CreateInstance()
			{
				TabType result = new TabType();
				result.Caption = String.Empty;
				result.Link = String.Empty;
				result.StylePrefix = String.Empty;
				result.LiveBody = String.Empty;
				return result;
			}
		}
		private static ccCommonModule.TabType[] Tabs = null;
		private static int TabsCnt = 0;
		private static int TabsSize = 0;
		//
		// Admin Navigator Nodes
		//
		public const int NavigatorNodeCollectionList = -1;
		public const int NavigatorNodeAddonList = -1;
		//
		// Pointers into index of PCC (Page Content Cache) array
		//
		public const int PCC_ID = 0;
		public const int PCC_Active = 1;
		public const int PCC_ParentID = 2;
		public const int PCC_Name = 3;
		public const int PCC_Headline = 4;
		public const int PCC_MenuHeadline = 5;
		public const int PCC_DateArchive = 6;
		public const int PCC_DateExpires = 7;
		public const int PCC_PubDate = 8;
		public const int PCC_ChildListSortMethodID = 9;
		public const int PCC_ContentControlID = 10;
		public const int PCC_TemplateID = 11;
		public const int PCC_BlockContent = 12;
		public const int PCC_BlockPage = 13;
		public const int PCC_Link = 14;
		public const int PCC_RegistrationGroupID = 15;
		public const int PCC_BlockSourceID = 16;
		public const int PCC_CustomBlockMessageFilename = 17;
		public const int PCC_JSOnLoad = 18;
		public const int PCC_JSHead = 19;
		public const int PCC_JSEndBody = 20;
		public const int PCC_Viewings = 21;
		public const int PCC_ContactMemberID = 22;
		public const int PCC_AllowHitNotification = 23;
		public const int PCC_TriggerSendSystemEmailID = 24;
		public const int PCC_TriggerConditionID = 25;
		public const int PCC_TriggerConditionGroupID = 26;
		public const int PCC_TriggerAddGroupID = 27;
		public const int PCC_TriggerRemoveGroupID = 28;
		public const int PCC_AllowMetaContentNoFollow = 29;
		public const int PCC_ParentListName = 30;
		public const int PCC_CopyFilename = 31;
		public const int PCC_BriefFilename = 32;
		public const int PCC_AllowChildListDisplay = 33;
		public const int PCC_SortOrder = 34;
		public const int PCC_DateAdded = 35;
		public const int PCC_ModifiedDate = 36;
		public const int PCC_ChildPagesFound = 37;
		public const int PCC_AllowInMenus = 38;
		public const int PCC_AllowInChildLists = 39;
		public const int PCC_JSFilename = 40;
		public const int PCC_ChildListInstanceOptions = 41;
		public const int PCC_IsSecure = 42;
		public const int PCC_AllowBrief = 43;
		public const int PCC_ColCnt = 44;
		//
		// Indexes into the SiteSectionCache
		// Created from "ID, Name,TemplateID,ContentID,MenuImageFilename,Caption,MenuImageOverFilename,HideMenu,BlockSection,RootPageID,JSOnLoad,JSHead,JSEndBody"
		//
		public const int SSC_ID = 0;
		public const int SSC_Name = 1;
		public const int SSC_TemplateID = 2;
		public const int SSC_ContentID = 3;
		public const int SSC_MenuImageFilename = 4;
		public const int SSC_Caption = 5;
		public const int SSC_MenuImageOverFilename = 6;
		public const int SSC_HideMenu = 7;
		public const int SSC_BlockSection = 8;
		public const int SSC_RootPageID = 9;
		public const int SSC_JSOnLoad = 10;
		public const int SSC_JSHead = 11;
		public const int SSC_JSEndBody = 12;
		public const int SSC_JSFilename = 13;
		public const int SSC_cnt = 14;
		//
		// Indexes into the TemplateCache
		// Created from "t.ID,t.Name,t.Link,t.BodyHTML,t.JSOnLoad,t.JSHead,t.JSEndBody,t.StylesFilename,r.StyleID"
		//
		public const int TC_ID = 0;
		public const int TC_Name = 1;
		public const int TC_Link = 2;
		public const int TC_BodyHTML = 3;
		public const int TC_JSOnLoad = 4;
		public const int TC_JSInHeadLegacy = 5;
		//Public Const TC_JSHead = 5
		public const int TC_JSEndBody = 6;
		public const int TC_StylesFilename = 7;
		public const int TC_SharedStylesIDList = 8;
		public const int TC_MobileBodyHTML = 9;
		public const int TC_MobileStylesFilename = 10;
		public const int TC_OtherHeadTags = 11;
		public const int TC_BodyTag = 12;
		public const int TC_JSInHeadFilename = 13;
		//Public Const TC_JSFilename = 13
		public const int TC_IsSecure = 14;
		public const int TC_DomainIdList = 15;
		// for now, Mobile templates do not have shared styles
		//Public Const TC_MobileSharedStylesIDList = 11
		public const int TC_cnt = 16;
		//
		// DTD
		//
		public const string DTDDefault = "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\" \"http://www.w3.org/TR/html4/loose.dtd\">";
		public const string DTDDefaultAdmin = "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\" \"http://www.w3.org/TR/html4/loose.dtd\">";
		//
		// innova Editor feature list
		//
		public const string InnovaEditorFeaturefilename = "Config\\EditorCongif.txt";
		public const string InnovaEditorFeatureList = "FullScreen,Preview,Print,Search,Cut,Copy,Paste,PasteWord,PasteText,SpellCheck,Undo,Redo,Image,Flash,Media,CustomObject,CustomTag,Bookmark,Hyperlink,HTMLSource,XHTMLSource,Numbering,Bullets,Indent,Outdent,JustifyLeft,JustifyCenter,JustifyRight,JustifyFull,Table,Guidelines,Absolute,Characters,Line,Form,RemoveFormat,ClearAll,StyleAndFormatting,TextFormatting,ListFormatting,BoxFormatting,ParagraphFormatting,CssText,Styles,Paragraph,FontName,FontSize,Bold,Italic,Underline,Strikethrough,Superscript,Subscript,ForeColor,BackColor";
		public const string InnovaEditorPublicFeatureList = "FullScreen,Preview,Print,Search,Cut,Copy,Paste,PasteWord,PasteText,SpellCheck,Undo,Redo,Bookmark,Hyperlink,HTMLSource,XHTMLSource,Numbering,Bullets,Indent,Outdent,JustifyLeft,JustifyCenter,JustifyRight,JustifyFull,Table,Guidelines,Absolute,Characters,Line,Form,RemoveFormat,ClearAll,StyleAndFormatting,TextFormatting,ListFormatting,BoxFormatting,ParagraphFormatting,CssText,Styles,Paragraph,FontName,FontSize,Bold,Italic,Underline,Strikethrough,Superscript,Subscript,ForeColor,BackColor";
		//'
		//' Content Type
		//'
		//Enum contentTypeEnum
		//    contentTypeWeb = 1
		//    ContentTypeEmail = 2
		//    contentTypeWebTemplate = 3
		//    contentTypeEmailTemplate = 4
		//End Enum
		//Public EditorContext As contentTypeEnum
		//Enum EditorContextEnum
		//    contentTypeWeb = 1
		//    contentTypeEmail = 2
		//End Enum
		//Public EditorContext As EditorContextEnum
		//'
		//Public Const EditorAddonMenuEmailTemplateFilename = "templates/EditorAddonMenuTemplateEmail.js"
		//Public Const EditorAddonMenuEmailContentFilename = "templates/EditorAddonMenuContentEmail.js"
		//Public Const EditorAddonMenuWebTemplateFilename = "templates/EditorAddonMenuTemplateWeb.js"
		//Public Const EditorAddonMenuWebContentFilename = "templates/EditorAddonMenuContentWeb.js"
		//
		public const string DynamicStylesFilename = "templates/styles.css";
		public const string AdminSiteStylesFilename = "templates/AdminSiteStyles.css";
		public const string EditorStyleRulesFilenamePattern = "templates/EditorStyleRules$TemplateID$.js";
		// deprecated 11/24/3009 - StyleRules destinction between web/email not needed b/c body background blocked
		//Public Const EditorStyleWebRulesFilename = "templates/EditorStyleWebRules.js"
		//Public Const EditorStyleEmailRulesFilename = "templates/EditorStyleEmailRules.js"
		//
		// ----- ccGroupRules storage for list of Content that a group can author
		//
		public struct ContentGroupRuleType
		{
			public int ContentID;
			public int GroupID;
			public bool AllowAdd;
			public bool AllowDelete;
		}
		//
		// ----- This should match the Lookup List in the NavIconType field in the Navigator Entry content definition
		//
		public const string navTypeIDList = "Add-on,Report,Setting,Tool";
		public const int NavTypeIDAddon = 1;
		public const int NavTypeIDReport = 2;
		public const int NavTypeIDSetting = 3;
		public const int NavTypeIDTool = 4;
		//
		public const string NavIconTypeList = "Custom,Advanced,Content,Folder,Email,User,Report,Setting,Tool,Record,Addon,help";
		public const int NavIconTypeCustom = 1;
		public const int NavIconTypeAdvanced = 2;
		public const int NavIconTypeContent = 3;
		public const int NavIconTypeFolder = 4;
		public const int NavIconTypeEmail = 5;
		public const int NavIconTypeUser = 6;
		public const int NavIconTypeReport = 7;
		public const int NavIconTypeSetting = 8;
		public const int NavIconTypeTool = 9;
		public const int NavIconTypeRecord = 10;
		public const int NavIconTypeAddon = 11;
		public const int NavIconTypeHelp = 12;
		//
		public const int QueryTypeSQL = 1;
		public const int QueryTypeOpenContent = 2;
		public const int QueryTypeUpdateContent = 3;
		public const int QueryTypeInsertContent = 4;
		//
		// Google Data Object construction in GetRemoteQuery
		//
		public struct ColsType
		{
			public string Type;
			public string Id;
			public string Label;
			public string Pattern;
			public static ColsType CreateInstance()
			{
				ColsType result = new ColsType();
				result.Type = String.Empty;
				result.Id = String.Empty;
				result.Label = String.Empty;
				result.Pattern = String.Empty;
				return result;
			}
		}
		//
		public struct CellType
		{
			public string v;
			public string f;
			public string p;
			public static CellType CreateInstance()
			{
				CellType result = new CellType();
				result.v = String.Empty;
				result.f = String.Empty;
				result.p = String.Empty;
				return result;
			}
		}
		//
		public struct RowsType
		{
			public ccCommonModule.CellType[] Cell;
		}
		//
		public struct GoogleDataType
		{
			public bool IsEmpty;
			public ccCommonModule.ColsType[] col;
			public ccCommonModule.RowsType[] row;
		}
		//
		internal enum GoogleVisualizationStatusEnum
		{
			OK = 1,
			warning = 2,
			Error = 3
		}
		//
		public struct GoogleVisualizationType
		{
			public string version;
			public string reqid;
			public ccCommonModule.GoogleVisualizationStatusEnum status;
			public string[] warnings;
			public string[] errors;
			public string sig;
			public ccCommonModule.GoogleDataType table;
			public static GoogleVisualizationType CreateInstance()
			{
				GoogleVisualizationType result = new GoogleVisualizationType();
				result.version = String.Empty;
				result.reqid = String.Empty;
				result.sig = String.Empty;
				return result;
			}
		}

		//Public Const ReturnFormatTypeGoogleTable = 1
		//Public Const ReturnFormatTypeNameValue = 2

		internal enum RemoteFormatEnum
		{
			RemoteFormatJsonTable = 1,
			RemoteFormatJsonNameArray = 2,
			RemoteFormatJsonNameValue = 3
		}
		//
		//
		//
		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("advapi32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static int RegCloseKey(int hKey);
		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("advapi32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static int RegOpenKeyExA(int hKey, [MarshalAs(UnmanagedType.VBByRefStr)] ref string lpszSubKey, ref int dwOptions, int samDesired, ref int lpHKey);
		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("advapi32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static int RegQueryValueExA(int hKey, [MarshalAs(UnmanagedType.VBByRefStr)] ref string lpszValueName, int lpdwRes, ref int lpdwType, [MarshalAs(UnmanagedType.VBByRefStr)] ref string lpDataBuff, ref int nSize);
		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("advapi32.dll", EntryPoint = "RegQueryValueExA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static int RegQueryValueEx(int hKey, [MarshalAs(UnmanagedType.VBByRefStr)] ref string lpszValueName, int lpdwRes, ref int lpdwType, ref int lpDataBuff, ref int nSize);

		public const int HKEY_CLASSES_ROOT = unchecked((int) 0x80000000);
		public const int HKEY_CURRENT_USER = unchecked((int) 0x80000001);
		public const int HKEY_LOCAL_MACHINE = unchecked((int) 0x80000002);
		public const int HKEY_USERS = unchecked((int) 0x80000003);

		public const int ERROR_SUCCESS = 0;
		public const int REG_SZ = 1; // Unicode nul terminated string
		public const int REG_DWORD = 4; // 32-bit number

		public const int KEY_QUERY_VALUE = 0x1;
		public const int KEY_SET_VALUE = 0x2;
		public const int KEY_CREATE_SUB_KEY = 0x4;
		public const int KEY_ENUMERATE_SUB_KEYS = 0x8;
		public const int KEY_NOTIFY = 0x10;
		public const int KEY_CREATE_LINK = 0x20;
		public const int READ_CONTROL = 0x20000;
		public const int WRITE_DAC = 0x40000;
		public const int WRITE_OWNER = 0x80000;
		public const int SYNCHRONIZE = 0x100000;
		public const int STANDARD_RIGHTS_REQUIRED = 0xF0000;
		public const int STANDARD_RIGHTS_READ = READ_CONTROL;
		public const int STANDARD_RIGHTS_WRITE = READ_CONTROL;
		public const int STANDARD_RIGHTS_EXECUTE = READ_CONTROL;
		static readonly public int KEY_READ = STANDARD_RIGHTS_READ | KEY_QUERY_VALUE | KEY_ENUMERATE_SUB_KEYS | KEY_NOTIFY;
		static readonly public int KEY_WRITE = STANDARD_RIGHTS_WRITE | KEY_SET_VALUE | KEY_CREATE_SUB_KEY;
		static readonly public int KEY_EXECUTE = KEY_READ;

		//======================================================================================
		//
		//======================================================================================
		//
		internal static void StartDebugTimer(bool Enabled, string Label)
		{
			// ##### removed to catch err<>0 problem on error resume next
			if (Enabled)
			{
				if (TimerStackCount < TimerStackMax)
				{
					TimerStack[TimerStackCount].Label = Label;
					TimerStack[TimerStackCount].StartTicks = Environment.TickCount;
				}
				else
				{
					kmaCommonModule.AppendLogFile(Path.GetFileNameWithoutExtension(Application.ExecutablePath) + ".?.StartDebugTimer, " + "Timer Stack overflow, attempting push # [" + TimerStackCount.ToString() + "], but max = [" + TimerStackMax.ToString() + "]");
				}
				TimerStackCount++;
			}
		}
		//
		internal static void StopDebugTimer(bool Enabled, string Label)
		{
			// ##### removed to catch err<>0 problem on error resume next
			if (Enabled)
			{
				if (TimerStackCount <= 0)
				{
					kmaCommonModule.AppendLogFile(Path.GetFileNameWithoutExtension(Application.ExecutablePath) + ".?.StopDebugTimer, " + "Timer Error, attempting to Pop, but the stack is empty");
				}
				else
				{
					if (TimerStackCount <= TimerStackMax)
					{
						if (TimerStack[TimerStackCount - 1].Label == Label)
						{
							kmaCommonModule.AppendLogFile(Path.GetFileNameWithoutExtension(Application.ExecutablePath) + ".?.StopDebugTimer, " + "Timer [" + new string('.', 2 * TimerStackCount) + Label + "] took " + ((Environment.TickCount - TimerStack[TimerStackCount - 1].StartTicks).ToString()) + " msec");
						}
						else
						{
							kmaCommonModule.AppendLogFile(Path.GetFileNameWithoutExtension(Application.ExecutablePath) + ".?.StopDebugTimer, " + "Timer Error, [" + Label + "] was popped, but [" + TimerStack[TimerStackCount].Label + "] was on the top of the stack");
						}
					}
					TimerStackCount--;
				}
			}
		}
		//
		//
		//
		internal static string PayString(int Index)
		{
			// ##### removed to catch err<>0 problem on error resume next
			switch(Index)
			{
				case PayTypeCreditCardOnline : 
					return "Credit Card";
				case PayTypeCreditCardByPhone : 
					return "Credit Card by phone";
				case PayTypeCreditCardByFax : 
					return "Credit Card by fax";
				case PayTypeCHECK : 
					return "Personal Check";
				case PayTypeCHECKCOMPANY : 
					return "Company Check";
				case PayTypeCREDIT : 
					return "You will be billed";
				case PayTypeNetTerms : 
					return "Net Terms (Approved customers only)";
				case PayTypeCODCompanyCheck : 
					return "COD- Pre-Approved Only";
				case PayTypeCODCertifiedFunds : 
					return "COD- Certified Funds";
				case PayTypePAYPAL : 
					return "PayPal";
				default:
					// Case PayTypeNONE 
					return "No payment required";
			}
		}
		//
		//
		//
		internal static string CCString(int Index)
		{
			// ##### removed to catch err<>0 problem on error resume next
			switch(Index)
			{
				case CCTYPEVISA : 
					return "Visa";
				case CCTYPEMC : 
					return "MasterCard";
				case CCTYPEAMEX : 
					return "American Express";
				case CCTYPEDISCOVER : 
					return "Discover";
				default:
					// Case CCTYPENOVUS 
					return "Novus Card";
			}
		}
		//
		//========================================================================
		// Get a Long from a CommandPacket
		//   position+0, 4 byte value
		//========================================================================
		//
		internal static int GetLongFromByteArray(byte[] ByteArray, ref int Position)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			int result = 0;
			result = ByteArray[Position + 3];
			result = ByteArray[Position + 2] + (256 * result);
			result = ByteArray[Position + 1] + (256 * result);
			result = ByteArray[Position + 0] + (256 * result);
			Position += 4;
			//
			return result;
		}
		//
		//========================================================================
		// Get a Long from a byte array
		//   position+0, 4 byte size of the number
		//   position+3, start of the number
		//========================================================================
		//
		internal static int GetNumberFromByteArray(byte[] ByteArray, ref int Position)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			//
			int result = 0;
			int ArgumentLength = GetLongFromByteArray(ByteArray, ref Position);
			//
			if (ArgumentLength > 0)
			{
				result = 0;
				for (int ArgumentCount = ArgumentLength - 1; ArgumentCount >= 0; ArgumentCount--)
				{
					result = ByteArray[Position + ArgumentCount] + (256 * result);
				}
			}
			Position += ArgumentLength;
			//
			return result;
		}
		//
		//========================================================================
		// Get a String a byte array
		//   position+0, 4 byte length of the string
		//   position+3, start of the string
		//========================================================================
		//
		internal static string GetStringFromByteArray(byte[] ByteArray, ref int Position)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			//
			string result = "";
			int ArgumentLength = GetLongFromByteArray(ByteArray, ref Position);
			//
			if (ArgumentLength > 0)
			{
				for (int Pointer = 0; Pointer <= ArgumentLength - 1; Pointer++)
				{
					result = result + Strings.Chr(ByteArray[Position + Pointer]).ToString();
				}
			}
			Position += ArgumentLength;
			//
			return result;
		}
		//
		//========================================================================
		// Get a Long from a byte array
		//========================================================================
		//
		internal static void SetLongByteArray(byte[] ByteArray, ref int Position, int LongValue)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			ByteArray[Position + 0] = (byte) (LongValue & (0xFF));
			ByteArray[Position + 1] = (byte) (Convert.ToInt32(Math.Floor((double) (LongValue / 256d))) & (0xFF));
			ByteArray[Position + 2] = (byte) (Convert.ToInt32(Math.Floor((double) (LongValue / ((double) ((int) Math.Pow(256, 2)))))) & (0xFF));
			ByteArray[Position + 3] = (byte) (Convert.ToInt32(Math.Floor((double) (LongValue / ((double) ((int) Math.Pow(256, 3)))))) & (0xFF));
			Position += 4;
			//
		}
		//
		//========================================================================
		// Set a string in a byte array
		//========================================================================
		//
		internal static void SetStringByteArray(byte[] ByteArray, ref int Position, string StringValue)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			//
			int LenStringValue = Strings.Len(StringValue);
			if (LenStringValue > 0)
			{
				for (int Pointer = 0; Pointer <= LenStringValue - 1; Pointer++)
				{
					ByteArray[Position + Pointer] = (byte) (Strings.Asc(StringValue.Substring(Pointer, Math.Min(1, StringValue.Length - Pointer))[0]) & (0xFF));
				}
				Position += LenStringValue;
			}
			//
		}

		//
		//========================================================================
		//   Set a Long long on the end of a RMB (Remote Method Block)
		//       You determine the position, or it will add it to the end
		//========================================================================
		//
		internal static void SetRMBLong(ref byte[] ByteArray, int LongValue, int Position = 0)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			int MyPosition = 0;
			int ByteArraySize = 0;
			//
			// ----- determine the position
			//
			//UPGRADE_WARNING: (2065) Boolean method Information.IsMissing has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2065.aspx
			if (Position != null)
			{
				MyPosition = Position;
			}
			else
			{
				//
				// ----- Add it to the end, determine length
				//
				MyPosition = ByteArray[RMBPositionLength + 3];
				MyPosition = ByteArray[RMBPositionLength + 2] + (256 * MyPosition);
				MyPosition = ByteArray[RMBPositionLength + 1] + (256 * MyPosition);
				MyPosition = ByteArray[RMBPositionLength + 0] + (256 * MyPosition);
				//
				// ----- adjust size of array if necessary
				//
				ByteArraySize = ByteArray.GetUpperBound(0);
				if (ByteArraySize < (MyPosition + 8))
				{
					ByteArray = ArraysHelper.RedimPreserve(ByteArray, new int[]{ByteArraySize + 9});
				}
			}
			//
			// ----- set the length
			//
			//ByteArray(MyPosition + 0) = 4
			//ByteArray(MyPosition + 1) = 0
			//ByteArray(MyPosition + 2) = 0
			//ByteArray(MyPosition + 3) = 0
			//MyPosition = MyPosition + 4
			//
			// ----- set the value
			//
			ByteArray[MyPosition + 0] = (byte) (LongValue & (0xFF));
			ByteArray[MyPosition + 1] = (byte) (Convert.ToInt32(Math.Floor((double) (LongValue / 256d))) & (0xFF));
			ByteArray[MyPosition + 2] = (byte) (Convert.ToInt32(Math.Floor((double) (LongValue / ((double) ((int) Math.Pow(256, 2)))))) & (0xFF));
			ByteArray[MyPosition + 3] = (byte) (Convert.ToInt32(Math.Floor((double) (LongValue / ((double) ((int) Math.Pow(256, 3)))))) & (0xFF));
			MyPosition += 4;
			//
			//UPGRADE_WARNING: (2065) Boolean method Information.IsMissing has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2065.aspx
			if (Position == null)
			{
				//
				// ----- Adjust the RMB length if length not given
				//
				ByteArray[RMBPositionLength + 0] = (byte) (MyPosition & (0xFF));
				ByteArray[RMBPositionLength + 1] = (byte) (Convert.ToInt32(Math.Floor((double) (MyPosition / 256d))) & (0xFF));
				ByteArray[RMBPositionLength + 2] = (byte) (Convert.ToInt32(Math.Floor((double) (MyPosition / ((double) ((int) Math.Pow(256, 2)))))) & (0xFF));
				ByteArray[RMBPositionLength + 3] = (byte) (Convert.ToInt32(Math.Floor((double) (MyPosition / ((double) ((int) Math.Pow(256, 3)))))) & (0xFF));
			}
			//
		}
		//
		//========================================================================
		//   Set a Long long on the end of a RMB (Remote Method Block)
		//       You determine the position, or it will add it to the end
		//========================================================================
		//
		internal static void SetRMBString(ref byte[] ByteArray, string StringValue, int Position = 0)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			int MyPosition = 0;
			int ByteArraySize = 0;
			//
			// ----- determine the position
			//
			//UPGRADE_WARNING: (2065) Boolean method Information.IsMissing has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2065.aspx
			if (Position != null)
			{
				MyPosition = Position;
			}
			else
			{
				//
				// ----- Add it to the end, determine length
				//
				MyPosition = ByteArray[RMBPositionLength + 3];
				MyPosition = ByteArray[RMBPositionLength + 2] + (256 * MyPosition);
				MyPosition = ByteArray[RMBPositionLength + 1] + (256 * MyPosition);
				MyPosition = ByteArray[RMBPositionLength + 0] + (256 * MyPosition);
				//
				// ----- adjust size of array if necessary
				//
				ByteArraySize = ByteArray.GetUpperBound(0);
				if (ByteArraySize < (MyPosition + 8))
				{
					ByteArray = ArraysHelper.RedimPreserve(ByteArray, new int[]{ByteArraySize + 4 + Strings.Len(StringValue) + 1});
				}
			}
			//
			// ----- set the value
			//

			//
			//
			int LenStringValue = Strings.Len(StringValue);
			if (LenStringValue > 0)
			{
				for (int Pointer = 0; Pointer <= LenStringValue - 1; Pointer++)
				{
					ByteArray[MyPosition + Pointer] = (byte) (Strings.Asc(StringValue.Substring(Pointer, Math.Min(1, StringValue.Length - Pointer))[0]) & (0xFF));
				}
				MyPosition += LenStringValue;
			}
			//
			//UPGRADE_WARNING: (2065) Boolean method Information.IsMissing has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2065.aspx
			if (Position == null)
			{
				//
				// ----- Adjust the RMB length if length not given
				//
				ByteArray[RMBPositionLength + 0] = (byte) (MyPosition & (0xFF));
				ByteArray[RMBPositionLength + 1] = (byte) (Convert.ToInt32(Math.Floor((double) (MyPosition / 256d))) & (0xFF));
				ByteArray[RMBPositionLength + 2] = (byte) (Convert.ToInt32(Math.Floor((double) (MyPosition / ((double) ((int) Math.Pow(256, 2)))))) & (0xFF));
				ByteArray[RMBPositionLength + 3] = (byte) (Convert.ToInt32(Math.Floor((double) (MyPosition / ((double) ((int) Math.Pow(256, 3)))))) & (0xFF));
			}
			//
		}
		//
		//========================================================================
		//   IsTrue
		//       returns true or false depending on the state of the variant input
		//========================================================================
		//
		internal static bool IsTrue(object ValueVariant)
		{
			return kmaCommonModule.kmaEncodeBoolean(ValueVariant);
		}
		//
		//========================================================================
		// EncodeXML
		//
		//========================================================================
		//
		internal static string EncodeXML(object ValueVariant, int fieldType)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			string result = "";
			//
			switch(fieldType)
			{
				case FieldTypeInteger : case FieldTypeLookup : case FieldTypeRedirect : case FieldTypeManyToMany : case FieldTypeMemberSelect : 
					double dbNumericTemp = 0; 
					//UPGRADE_WARNING: (1049) Use of Null/IsNull() detected. More Information: http://www.vbtonet.com/ewis/ewi1049.aspx 
					if (Convert.IsDBNull(ValueVariant))
					{
						result = "null";
					}
					else if (ReflectionHelper.GetPrimitiveValue<string>(ValueVariant) == "")
					{ 
						result = "null";
					}
					else if (Double.TryParse(ReflectionHelper.GetPrimitiveValue<string>(ValueVariant), NumberStyles.Float , CultureInfo.CurrentCulture.NumberFormat, out dbNumericTemp))
					{ 
						result = Math.Floor((double) ReflectionHelper.GetPrimitiveValue<double>(ValueVariant)).ToString();
					}
					else
					{
						result = "null";
					} 
					break;
				case FieldTypeBoolean : 
					//UPGRADE_WARNING: (1049) Use of Null/IsNull() detected. More Information: http://www.vbtonet.com/ewis/ewi1049.aspx 
					if (Convert.IsDBNull(ValueVariant))
					{
						result = "0";
					}
					else if (ReflectionHelper.GetPrimitiveValue<bool>(ValueVariant))
					{ 
						result = "1";
					}
					else
					{
						result = "0";
					} 
					break;
				case FieldTypeCurrency : 
					double dbNumericTemp2 = 0; 
					//UPGRADE_WARNING: (1049) Use of Null/IsNull() detected. More Information: http://www.vbtonet.com/ewis/ewi1049.aspx 
					if (Convert.IsDBNull(ValueVariant))
					{
						result = "null";
					}
					else if (ReflectionHelper.GetPrimitiveValue<string>(ValueVariant) == "")
					{ 
						result = "null";
					}
					else if (Double.TryParse(ReflectionHelper.GetPrimitiveValue<string>(ValueVariant), NumberStyles.Float , CultureInfo.CurrentCulture.NumberFormat, out dbNumericTemp2))
					{ 
						result = ReflectionHelper.GetPrimitiveValue<string>(ValueVariant);
					}
					else
					{
						result = "null";
					} 
					break;
				case FieldTypeFloat : 
					double dbNumericTemp3 = 0; 
					//UPGRADE_WARNING: (1049) Use of Null/IsNull() detected. More Information: http://www.vbtonet.com/ewis/ewi1049.aspx 
					if (Convert.IsDBNull(ValueVariant))
					{
						result = "null";
					}
					else if (ReflectionHelper.GetPrimitiveValue<string>(ValueVariant) == "")
					{ 
						result = "null";
					}
					else if (Double.TryParse(ReflectionHelper.GetPrimitiveValue<string>(ValueVariant), NumberStyles.Float , CultureInfo.CurrentCulture.NumberFormat, out dbNumericTemp3))
					{ 
						result = ReflectionHelper.GetPrimitiveValue<string>(ValueVariant);
					}
					else
					{
						result = "null";
					} 
					break;
				case FieldTypeDate : 
					//UPGRADE_WARNING: (1049) Use of Null/IsNull() detected. More Information: http://www.vbtonet.com/ewis/ewi1049.aspx 
					if (Convert.IsDBNull(ValueVariant))
					{
						result = "null";
					}
					else if (ReflectionHelper.GetPrimitiveValue<string>(ValueVariant) == "")
					{ 
						result = "null";
					}
					else if (Information.IsDate(ValueVariant))
					{ 
						//TimeVar = CDate(ValueVariant)
						//TimeValuething = 86400! * (TimeVar - Int(TimeVar))
						//TimeHours = Int(TimeValuething / 3600!)
						//TimeMinutes = Int(TimeValuething / 60!) - (TimeHours * 60)
						//TimeSeconds = TimeValuething - (TimeHours * 3600!) - (TimeMinutes * 60!)
						//EncodeXML = Year(ValueVariant) & "-" & Right("0" & Month(ValueVariant), 2) & "-" & Right("0" & Day(ValueVariant), 2) & " " & Right("0" & TimeHours, 2) & ":" & Right("0" & TimeMinutes, 2) & ":" & Right("0" & TimeSeconds, 2)
						result = kmaCommonModule.kmaEncodeText(ValueVariant);
					} 
					break;
				default:
					// 
					// ----- FieldTypeText 
					// ----- FieldTypeLongText 
					// ----- FieldTypeFile 
					// ----- FieldTypeImage 
					// ----- FieldTypeTextFile 
					// ----- FieldTypeCSSFile 
					// ----- FieldTypeXMLFile 
					// ----- FieldTypeJavascriptFile 
					// ----- FieldTypeLink 
					// ----- FieldTypeResourceLink 
					// ----- FieldTypeHTML 
					// ----- FieldTypeHTMLFile 
					// 
					//UPGRADE_WARNING: (1049) Use of Null/IsNull() detected. More Information: http://www.vbtonet.com/ewis/ewi1049.aspx 
					if (Convert.IsDBNull(ValueVariant))
					{
						result = "null";
					}
					else if (ReflectionHelper.GetPrimitiveValue<string>(ValueVariant) == "")
					{ 
						result = "";
					}
					else
					{
						//EncodeXML = ASPServer.HTMLEncode(ValueVariant)
						//EncodeXML = Replace(ValueVariant, "&", "&lt;")
						//EncodeXML = Replace(ValueVariant, "<", "&lt;")
						//EncodeXML = Replace(EncodeXML, ">", "&gt;")
					} 
					break;
			}
			//
			return result;
		}
		//
		//========================================================================
		// EncodeFilename
		//
		//========================================================================
		//
		internal static string encodeFilename(string Source)
		{
			string chr = "";
			//
			StringBuilder returnString = new StringBuilder();
			int cnt = Strings.Len(Source);
			if (cnt > 254)
			{
				cnt = 254;
			}
			string allowed = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ^&'@{}[],$-#()%.+~_";
			for (int Ptr = 1; Ptr <= cnt; Ptr++)
			{
				chr = Source.Substring(Ptr - 1, Math.Min(1, Source.Length - (Ptr - 1)));
				if (allowed.IndexOf(chr) >= 0)
				{
					returnString.Append(chr);
				}
			}
			return returnString.ToString();
		}
		//
		//Function encodeFilename(Filename As String) As String
		//    ' ##### removed to catch err<>0 problem on error resume next
		//    '
		//    Dim Source() As Variant
		//    Dim Replacement() As Variant
		//    '
		//    Source = Array("""", "*", "/", ":", "<", ">", "?", "\", "|", "=")
		//    Replacement = Array("_", "_", "_", "_", "_", "_", "_", "_", "_", "_")
		//    '
		//    encodeFilename = ReplaceMany(Filename, Source, Replacement)
		//    If Len(encodeFilename) > 254 Then
		//        encodeFilename = Left(encodeFilename, 254)
		//    End If
		//    encodeFilename = Replace(encodeFilename, vbCr, "_")
		//    encodeFilename = Replace(encodeFilename, vbLf, "_")
		//    '
		//    End Function
		//
		//
		//

		//
		//========================================================================
		// DecodeHTML
		//
		//========================================================================
		//
		internal static string DecodeHTML(string Source)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			return kmaCommonModule.kmaDecodeHTML(Source);
			//Dim SourceChr() As Variant
			//Dim ReplacementChr() As Variant
			//'
			//SourceChr = Array("&gt;", "&lt;", "&nbsp;", "&amp;")
			//ReplacementChr = Array(">", "<", " ", "&")
			//'
			//DecodeHTML = ReplaceMany(Source, SourceChr, ReplacementChr)
			//
		}
		//
		//========================================================================
		// EncodeFilename
		//
		//========================================================================
		//
		internal static string ReplaceMany(string Source, string[] ArrayOfSource, string[] ArrayOfReplacement)
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
		//
		//
		internal static string GetURIHost(string URI)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			//   Divide the URI into URIHost, URIPath, and URIPage
			//
			int LastSlash = 0;
			string URIPath = "";
			string URIPage = "";
			string URIWorking = URI;
			if (URIWorking.ToUpper().StartsWith("HTTP"))
			{
				URIWorking = URIWorking.Substring(URIWorking.IndexOf("//") + 2);
			}
			string URIHost = URIWorking;
			int Slash = (URIHost.IndexOf('/') + 1);
			if (Slash == 0)
			{
				URIPath = "/";
				URIPage = "";
			}
			else
			{
				URIPath = URIHost.Substring(Slash - 1);
				URIHost = URIHost.Substring(0, Math.Min(Slash - 1, URIHost.Length));
				Slash = (URIPath.IndexOf('/') + 1);

				while(Slash != 0)
				{
					LastSlash = Slash;
					Slash = Strings.InStr(LastSlash + 1, URIPath, "/", CompareMethod.Binary);
					Application.DoEvents();
				};
				URIPage = URIPath.Substring(LastSlash);
				URIPath = URIPath.Substring(0, Math.Min(LastSlash, URIPath.Length));
			}
			return URIHost;
			//
		}
		//
		//
		//
		internal static string GetURIPage(string URI)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			//   Divide the URI into URIHost, URIPath, and URIPage
			//
			int LastSlash = 0;
			string URIPath = "";
			string URIPage = "";
			string URIWorking = URI;
			if (URIWorking.ToUpper().StartsWith("HTTP"))
			{
				URIWorking = URIWorking.Substring(URIWorking.IndexOf("//") + 2);
			}
			string URIHost = URIWorking;
			int Slash = (URIHost.IndexOf('/') + 1);
			if (Slash == 0)
			{
				URIPath = "/";
				URIPage = "";
			}
			else
			{
				URIPath = URIHost.Substring(Slash - 1);
				URIHost = URIHost.Substring(0, Math.Min(Slash - 1, URIHost.Length));
				Slash = (URIPath.IndexOf('/') + 1);

				while(Slash != 0)
				{
					LastSlash = Slash;
					Slash = Strings.InStr(LastSlash + 1, URIPath, "/", CompareMethod.Binary);
					Application.DoEvents();
				};
				URIPage = URIPath.Substring(LastSlash);
				URIPath = URIPath.Substring(0, Math.Min(LastSlash, URIPath.Length));
			}
			return URIPage;
			//
		}
		//
		//
		//
		internal static System.DateTime GetDateFromGMT(string GMTDate)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			System.DateTime result = DateTime.FromOADate(0);
			string WorkString = "";
			if (GMTDate != "")
			{
				WorkString = GMTDate.Substring(5, Math.Min(11, GMTDate.Length - 5));
				if (Information.IsDate(WorkString))
				{
					result = DateTime.Parse(WorkString);
					WorkString = GMTDate.Substring(17, Math.Min(8, GMTDate.Length - 17));
					if (Information.IsDate(WorkString))
					{
						result = result.AddDays(DateTime.Parse(WorkString).ToOADate()).AddDays(4 / 24d);
					}
				}
			}
			//
			return result;
		}
        ////
        //// Wdy, DD-Mon-YYYY HH:MM:SS GMT
        ////
        //internal static string GetGMTFromDate(System.DateTime DateValue)
        //{
        //    //
        //    string result = "";
        //    int WorkLong = 0;
        //    //
        //    if (Information.IsDate(DateValue))
        //    {
        //        FirstDayOfWeek switchVar = InvalidCastException DateAndTime.Weekday( DateValue, FirstDayOfWeek.Sunday);
        //        if (switchVar == FirstDayOfWeek.Sunday)
        //        {
        //            result = "Sun, ";
        //        }
        //        else if (switchVar == FirstDayOfWeek.Monday)
        //        { 
        //            result = "Mon, ";
        //        }
        //        else if (switchVar == FirstDayOfWeek.Tuesday)
        //        { 
        //            result = "Tue, ";
        //        }
        //        else if (switchVar == FirstDayOfWeek.Wednesday)
        //        { 
        //            result = "Wed, ";
        //        }
        //        else if (switchVar == FirstDayOfWeek.Thursday)
        //        { 
        //            result = "Thu, ";
        //        }
        //        else if (switchVar == FirstDayOfWeek.Friday)
        //        { 
        //            result = "Fri, ";
        //        }
        //        else if (switchVar == FirstDayOfWeek.Saturday)
        //        { 
        //            result = "Sat, ";
        //        }
        //        //
        //        WorkLong = DateAndTime.Day(DateValue);
        //        if (WorkLong < 10)
        //        {
        //            result = result + "0" + WorkLong.ToString() + " ";
        //        }
        //        else
        //        {
        //            result = result + WorkLong.ToString() + " ";
        //        }
        //        //
        //        switch(DateValue.Month)
        //        {
        //            case 1 : 
        //                result = result + "Jan "; 
        //                break;
        //            case 2 : 
        //                result = result + "Feb "; 
        //                break;
        //            case 3 : 
        //                result = result + "Mar "; 
        //                break;
        //            case 4 : 
        //                result = result + "Apr "; 
        //                break;
        //            case 5 : 
        //                result = result + "May "; 
        //                break;
        //            case 6 : 
        //                result = result + "Jun "; 
        //                break;
        //            case 7 : 
        //                result = result + "Jul "; 
        //                break;
        //            case 8 : 
        //                result = result + "Aug "; 
        //                break;
        //            case 9 : 
        //                result = result + "Sep "; 
        //                break;
        //            case 10 : 
        //                result = result + "Oct "; 
        //                break;
        //            case 11 : 
        //                result = result + "Nov "; 
        //                break;
        //            case 12 : 
        //                result = result + "Dec "; 
        //                break;
        //        }
        //        //
        //        result = result + DateValue.Year.ToString() + " ";
        //        //
        //        WorkLong = DateAndTime.Hour(DateValue);
        //        if (WorkLong < 10)
        //        {
        //            result = result + "0" + WorkLong.ToString() + ":";
        //        }
        //        else
        //        {
        //            result = result + WorkLong.ToString() + ":";
        //        }
        //        //
        //        WorkLong = DateValue.Minute;
        //        if (WorkLong < 10)
        //        {
        //            result = result + "0" + WorkLong.ToString() + ":";
        //        }
        //        else
        //        {
        //            result = result + WorkLong.ToString() + ":";
        //        }
        //        //
        //        WorkLong = DateValue.Second;
        //        if (WorkLong < 10)
        //        {
        //            result = result + "0" + WorkLong.ToString();
        //        }
        //        else
        //        {
        //            result = result + WorkLong.ToString();
        //        }
        //        //
        //        result = result + " GMT";
        //    }
        //    //
        //    return result;
        //}
		//
		//========================================================================
		//   EncodeSQL
		//       encode a variable to go in an sql expression
		//       NOT supported
		//========================================================================
		//
		internal static string EncodeSQL(object ExpressionVariant, object fieldType = null)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			//
			string result = "";
			//
			int iFieldType = kmaCommonModule.KmaEncodeMissingInteger(fieldType, FieldTypeText);
			switch(iFieldType)
			{
				case FieldTypeBoolean : 
					return kmaCommonModule.KmaEncodeSQLBoolean(ExpressionVariant);
				case FieldTypeCurrency : case FieldTypeAutoIncrement : case FieldTypeFloat : case FieldTypeInteger : case FieldTypeLookup : case FieldTypeMemberSelect : 
					return kmaCommonModule.KmaEncodeSQLNumber(ExpressionVariant);
				case FieldTypeDate : 
					return kmaCommonModule.KmaEncodeSQLDate(ExpressionVariant);
				case FieldTypeLongText : case FieldTypeHTML : 
					return kmaCommonModule.KmaEncodeSQLLongText(ExpressionVariant);
				case FieldTypeFile : case FieldTypeImage : case FieldTypeLink : case FieldTypeResourceLink : case FieldTypeRedirect : case FieldTypeManyToMany : case FieldTypeText : case FieldTypeTextFile : case FieldTypeJavascriptFile : case FieldTypeXMLFile : case FieldTypeCSSFile : case FieldTypeHTMLFile : 
					return cp.Db.EncodeSQLText(ExpressionVariant);
				default:
					result = cp.Db.EncodeSQLText(ExpressionVariant); 
					UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Goto Label (0)"); 
					throw new System.Exception(kmaCommonModule.KmaErrorBase.ToString() + ", " + Path.GetFileNameWithoutExtension(Application.ExecutablePath) + ", " + "Unknown Field Type [" + ReflectionHelper.GetPrimitiveValue<string>(fieldType) + "] used FieldTypeText.");
			}
			//
			return result;
		}
		//'
		//'
		//'
		//Public Sub AppendLogFile(Text)
		//    On Error GoTo 0
		//    Dim kmafs As New kmaFileSystem3.FileSystemClass
		//    Dim Filename As String
		//    Filename = GetProgramPath() & "\logs\Trace" & kmaEncodeText(CLng(Int(Now()))) & ".log"
		//    Call kmafs.AppendLogFile2(Filename, """" & FormatDateTime(Now(), vbGeneralDate) & """,""" & Text & """" & vbCrLf)
		//    End Sub
		//'
		//'========================================================================
		//'   HandleError
		//'       Logs the error and either resumes next, or raises it to the next level
		//'========================================================================
		//'
		//Public Function HandleError(ClassName As String, MethodName As String, ErrNumber As Long, ErrSource As String, ErrDescription As String, ErrorTrap As Boolean, ResumeNext As Boolean, Optional URL As String) As String
		//    ' ##### removed to catch err<>0 problem on error resume next
		//    '
		//    Dim ErrorMessage As String
		//    '
		//    If ErrorTrap Then
		//        ErrorMessage = ErrorMessage & " Unexpected ErrorTrap"
		//    Else
		//        ErrorMessage = ErrorMessage & " Error"
		//        End If
		//    '
		//    If URL <> "" Then
		//        ErrorMessage = ErrorMessage & " on page [" & URL & "]"
		//        End If
		//    '
		//    If ErrorTrap Then
		//        If ResumeNext Then
		//            Call AppendLogFile(App.EXEName & "." & ClassName & "." & MethodName & ErrorMessage & ", will resume after logging [" & ErrSource & " #" & ErrNumber & ", " & ErrDescription & "]")
		//        Else
		//            Call AppendLogFile(App.EXEName & "." & ClassName & "." & MethodName & ErrorMessage & ", will abort after logging [" & ErrSource & " #" & ErrNumber & ", " & ErrDescription & "]")
		//            On Error GoTo 0
		//            Call Err.Raise(ErrNumber, ErrSource, ErrDescription)
		//            End If
		//    Else
		//        If ResumeNext Then
		//            Call AppendLogFile(App.EXEName & "." & ClassName & "." & MethodName & ErrorMessage & ", will resume after logging  [" & ErrSource & " #" & ErrNumber & ", " & ErrDescription & "]")
		//        Else
		//            Call AppendLogFile(App.EXEName & "." & ClassName & "." & MethodName & ErrorMessage & ", will abort after logging [" & ErrSource & " #" & ErrNumber & ", " & ErrDescription & "]")
		//            On Error GoTo 0
		//            Call Err.Raise(ErrNumber, ErrSource, ErrDescription, , -1)
		//            End If
		//        End If
		//    '
		//    End Function
		//
		//
		//
		internal static void cpTick(string Text)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			int Duration = 0;
			if (CPTickCountBase != 0)
			{
				Duration = (Convert.ToInt32(Environment.TickCount - CPTickCountBase));
			}
			string iText = "cpTick " + StringsHelper.Format(Duration / 1000d, "0000.000") + " " + Path.GetFileNameWithoutExtension(Application.ExecutablePath) + " " + Text;
			kmaCommonModule.AppendLogFile(Path.GetFileNameWithoutExtension(Application.ExecutablePath) + ".?.cpTick, " + iText);
			CPTickCountBase = Environment.TickCount;
			//
		}
		//
		//=====================================================================================================
		//   Set a value in a name/value pair
		//=====================================================================================================
		//
		internal static void SetNameValueArrays(string InputName, string InputValue, string[] SQLName, string[] SQLValue, ref int Index)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			SQLName[Index] = InputName;
			SQLValue[Index] = InputValue;
			Index++;
			//
		}
		//
		//
		//
		internal static string GetApplicationStatusMessage(int ApplicationStatus)
		{
			switch(ApplicationStatus)
			{
				case ApplicationStatusNoHostService : 
					return "Contensive server not running";
				case ApplicationStatusNotFound : 
					return "Contensive application not found";
				case ApplicationStatusLoadedNotRunning : 
					return "Contensive application not running";
				case ApplicationStatusRunning : 
					return "Contensive application running";
				case ApplicationStatusStarting : 
					return "Contensive application starting";
				case ApplicationStatusUpgrading : 
					return "Contensive database upgrading";
				case ApplicationStatusDbBad : 
					return "Error verifying core database records";
				case ApplicationStatusDbFailure : 
					return "Error opening application database";
				case ApplicationStatusKernelFailure : 
					return "Error contacting Contensive kernel services";
				case ApplicationStatusLicenseFailure : 
					return "Error verifying Contensive site license, see Http://www.Contensive.com/License";
				case ApplicationStatusConnectionObjectFailure : 
					return "Error creating ODBC Connection object";
				case ApplicationStatusConnectionStringFailure : 
					return "ODBC Data Source connection failed";
				case ApplicationStatusDataSourceFailure : 
					return "Error opening default data source";
				case ApplicationStatusDuplicateDomains : 
					return "Can not determine application because there are multiple applications with domain names that match this site's domain (See Application Manager)";
				case ApplicationStatusUnknownFailure : 
					return "Unknown error, see trace log for details (/Contensive/Logs/trace____.log)";
				case ApplicationStatusPaused : 
					return "Contensive application paused";
				default:
					return "Unknown status code [" + ApplicationStatus.ToString() + "], see trace log for details";
			}
		}
		//
		//
		//
		internal static string GetFormInputSelectNameValue(string SelectName, ccCommonModule.NameValuePairType[] NameValueArray)
		{
			//
			string result = "";
			ccCommonModule.NameValuePairType[] Source = null;
			result = "<SELECT name=\"" + SelectName + "\" Size=\"1\">";
			for (int Pointer = 0; Pointer <= NameValueArray.GetUpperBound(0); Pointer++)
			{
				result = result + "<OPTION value=\"" + Source[Pointer].Value + "\">" + Source[Pointer].Name + "</OPTION>";
			}
			return result + "</SELECT>";
		}
		//
		//
		//
		internal static string kmaGetSpacer(int Width, int Height)
		{
			return "<img alt=\"space\" src=\"/ccLib/images/spacer.gif\" width=\"" + Width.ToString() + "\" height=\"" + Height.ToString() + "\" border=\"0\">";
		}
		//
		//
		//
		internal static string kmaProcessReplacement(object NameValueLines, object Source)
		{
			//
			string result = "";
			string[] Lines = null;
			//
			string[] Names = null;
			string[] Values = null;
			int PairCnt = 0;
			string[] Splits = null;
			//
			// ----- read pairs in from NameValueLines
			//
			string iNameValueLines = kmaCommonModule.kmaEncodeText(NameValueLines);
			if (iNameValueLines.IndexOf('=') >= 0)
			{
				PairCnt = 0;
				Lines = (string[]) SplitCRLF(ref iNameValueLines);
				foreach (string Lines_item in Lines)
				{
					if (Lines_item.IndexOf('=') >= 0)
					{
						Splits = (string[]) Lines_item.Split('=');
						Names = ArraysHelper.RedimPreserve(Names, new int[]{PairCnt + 1});
						Names = ArraysHelper.RedimPreserve(Names, new int[]{PairCnt + 1});
						Values = ArraysHelper.RedimPreserve(Values, new int[]{PairCnt + 1});
						Names[PairCnt] = Splits[0].Trim();
						Names[PairCnt] = Strings.Replace(Names[PairCnt], "\t", "", 1, -1, CompareMethod.Binary);
						Splits[0] = "";
						Values[PairCnt] = Splits[1].Trim();
						PairCnt++;
					}
				}
			}
			//
			// ----- Process replacements on Source
			//
			result = kmaCommonModule.kmaEncodeText(Source);
			if (PairCnt > 0)
			{
				for (int PairPtr = 0; PairPtr <= PairCnt - 1; PairPtr++)
				{
					result = Strings.Replace(result, Names[PairPtr], Values[PairPtr], 1, 999, CompareMethod.Text);
				}
			}
			//
			return result;
		}
		//
		//==========================================================================================================================
		//   To convert from site license to server licenses, we still need the URLEncoder in the site license
		//   This routine generates a site license that is just the URL encoder.
		//==========================================================================================================================
		//
		internal static string GetURLEncoder()
		{
			VBMath.Randomize();
			return ((float) Math.Floor((double) (1 + (VBMath.Rnd() * 8)))).ToString() + ((float) Math.Floor((double) (1 + (VBMath.Rnd() * 8)))).ToString() + ((float) Math.Floor((double) (1000000000 + (VBMath.Rnd() * 899999999)))).ToString();
		}
		//
		//==========================================================================================================================
		//   To convert from site license to server licenses, we still need the URLEncoder in the site license
		//   This routine generates a site license that is just the URL encoder.
		//==========================================================================================================================
		//
		internal static string GetSiteLicenseKey()
		{
			return "00000-00000-00000-" + GetURLEncoder();
		}
		//
		//
		//
		internal static void ccAddTabEntry(string Caption, string Link, bool IsHit, string StylePrefix = "", string LiveBody = "")
		{
			try
			{
				//
				if (TabsCnt <= TabsSize)
				{
					TabsSize += 10;
					Tabs = ArraysHelper.RedimPreserve(Tabs, new int[]{TabsSize + 1});
				}
				Tabs[TabsCnt].Caption = Caption;
				Tabs[TabsCnt].Link = Link;
				Tabs[TabsCnt].IsHit = IsHit;
				Tabs[TabsCnt].StylePrefix = kmaCommonModule.KmaEncodeMissingText(StylePrefix, "cc");
				Tabs[TabsCnt].LiveBody = kmaCommonModule.KmaEncodeMissingText(LiveBody, "");
				TabsCnt++;
				//
			}
			catch (System.Exception excep)
			{
				//
				//UPGRADE_WARNING: (2081) Err.Number has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2081.aspx
				throw new System.Exception(Information.Err().Number.ToString() + ", " + excep.Source + ", " + "Error in ccAddTabEntry-" + excep.Message);
			}
		}
		//
		//
		//
		internal static string OldccGetTabs()
		{
			string result = "";
			try
			{
				//
				int HitPtr = 0;
				bool IsLiveTab = false;
				StringBuilder TabBody = new StringBuilder();
				string TabLink = "";
				string TabID = "";
				bool FirstLiveBodyShown = false;
				//
				if (TabsCnt > 0)
				{
					HitPtr = 0;
					//
					// Create TabBar
					//
					result = result + "<table border=0 cellspacing=0 cellpadding=0 align=center ><tr>";
					for (int TabPtr = 0; TabPtr <= TabsCnt - 1; TabPtr++)
					{
						TabID = kmaCommonModule.GetRandomInteger().ToString();
						if (Tabs[TabPtr].LiveBody == "")
						{
							//
							// This tab is linked to a page
							//
							TabLink = kmaCommonModule.kmaEncodeHTML(Tabs[TabPtr].Link);
						}
						else
						{
							//
							// This tab has a live body
							//
							TabLink = kmaCommonModule.kmaEncodeHTML(Tabs[TabPtr].Link);
							if (!FirstLiveBodyShown)
							{
								FirstLiveBodyShown = true;
								TabBody.Append("<div style=\"visibility: visible; position: absolute; left: 0px;\" class=\"" + Tabs[TabPtr].StylePrefix + "Body\" id=\"" + TabID + "\"></div>");
							}
							else
							{
								TabBody.Append("<div style=\"visibility: hidden; position: absolute; left: 0px;\" class=\"" + Tabs[TabPtr].StylePrefix + "Body\" id=\"" + TabID + "\"></div>");
							}
						}
						result = result + "<td valign=bottom>";
						if (Tabs[TabPtr].IsHit && (HitPtr == 0))
						{
							HitPtr = TabPtr;
							//
							// This tab is hit
							//
							result = result + "<table cellspacing=0 cellPadding=0 border=0>";
							result = result + "<tr>" + "<td colspan=2 height=1 width=2></td>" + "<td colspan=1 height=1 bgcolor=black></td>" + "<td colspan=3 height=1 width=3></td>" + "</tr>";
							result = result + "<tr>" + "<td colspan=1 height=1 width=1></td>" + "<td colspan=1 height=1 width=1 bgcolor=black></td>" + "<td colspan=1 height=1></td>" + "<td colspan=1 height=1 width=1 bgcolor=black></td>" + "<td colspan=2 height=1 width=2></td>" + "</tr>";
							result = result + "<tr>" + "<td colspan=1 height=2 bgcolor=black></td>" + "<td colspan=1 height=2></td>" + "<td colspan=1 height=2></td>" + "<td colspan=1 height=2></td>" + "<td colspan=1 height=2 width=1 bgcolor=black></td>" + "<td colspan=1 height=2 width=1></td>" + "</tr>";
							result = result + "<tr>" + "<td bgcolor=black></td>" + "<td></td>" + "<td>" + "<table cellspacing=0 cellPadding=2 border=0><tr>" + "<td Class=\"ccTabHit\">&nbsp;<a href=\"" + TabLink + "\" Class=\"ccTabHit\">" + Tabs[TabPtr].Caption + "</a>&nbsp;</td>" + "</tr></table >" + "</td>" + "<td></td>" + "<td bgcolor=black></td>" + "<td></td>" + "</tr>";
							result = result + "<tr>" + "<td bgcolor=black></td>" + "<td></td>" + "<td></td>" + "<td></td>" + "<td bgcolor=black></td>" + "<td bgcolor=black></td>" + "</tr>" + "</table >";
						}
						else
						{
							//
							// This tab is not hit
							//
							result = result + "<table cellspacing=0 cellPadding=0 border=0>";
							result = result + "<tr>" + "<td colspan=6 height=1></td>" + "</tr>";
							result = result + "<tr>" + "<td colspan=2 height=1></td>" + "<td colspan=1 height=1 bgcolor=black></td>" + "<td colspan=3 height=1></td>" + "</tr>";
							result = result + "<tr>" + "<td width=1></td>" + "<td width=1 bgcolor=black></td>" + "<td></td>" + "<td width=1 bgcolor=black></td>" + "<td width=2 colspan=2></td>" + "</tr>";
							result = result + "<tr>" + "<td width=1 bgcolor=black></td>" + "<td width=1></td>" + "<td nowrap>" + "<table cellspacing=0 cellPadding=2 border=0><tr>" + "<td Class=\"ccTab\">&nbsp;<a href=\"" + TabLink + "\" Class=\"ccTab\">" + Tabs[TabPtr].Caption + "</a>&nbsp;</td>" + "</tr></table >" + "</td>" + "<td width=1></td>" + "<td width=1 bgcolor=black></td>" + "<td width=1></td>" + "</tr>";
							result = result + "<tr>" + "<td colspan=6 height=1 bgcolor=black></td>" + "</tr>" + "</table >";
						}
						result = result + "</td>";
					}
					result = result + "<td class=\"ccTabEnd\">&nbsp;</td></tr>";
					if (TabBody.ToString() != "")
					{
						result = result + "<tr><td colspan=6>" + TabBody.ToString() + "</td></tr>";
					}
					result = result + "</tr></table >";
					TabsCnt = 0;
				}
				//
			}
			catch (System.Exception excep)
			{
				//
				//UPGRADE_WARNING: (2081) Err.Number has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2081.aspx
				throw new System.Exception(Information.Err().Number.ToString() + ", " + excep.Source + ", " + "Error in OldccGetTabs-" + excep.Message);
			}
			return result;
		}


		//
		//
		//
		internal static string ccGetTabs()
		{
			string result = "";
			try
			{
				//
				int HitPtr = 0;
				bool IsLiveTab = false;
				StringBuilder TabBody = new StringBuilder();
				string TabLink = "";
				string TabID = "";
				bool FirstLiveBodyShown = false;
				//
				if (TabsCnt > 0)
				{
					HitPtr = 0;
					//
					// Create TabBar
					//
					result = result + "<table border=0 cellspacing=0 cellpadding=0 align=center ><tr>";
					for (int TabPtr = 0; TabPtr <= TabsCnt - 1; TabPtr++)
					{
						TabID = kmaCommonModule.GetRandomInteger().ToString();
						if (Tabs[TabPtr].LiveBody == "")
						{
							//
							// This tab is linked to a page
							//
							TabLink = kmaCommonModule.kmaEncodeHTML(Tabs[TabPtr].Link);
						}
						else
						{
							//
							// This tab has a live body
							//
							TabLink = kmaCommonModule.kmaEncodeHTML(Tabs[TabPtr].Link);
							if (!FirstLiveBodyShown)
							{
								FirstLiveBodyShown = true;
								TabBody.Append("<div style=\"visibility: visible; position: absolute; left: 0px;\" class=\"" + Tabs[TabPtr].StylePrefix + "Body\" id=\"" + TabID + "\">" + Tabs[TabPtr].LiveBody + "</div>");
							}
							else
							{
								TabBody.Append("<div style=\"visibility: hidden; position: absolute; left: 0px;\" class=\"" + Tabs[TabPtr].StylePrefix + "Body\" id=\"" + TabID + "\">" + Tabs[TabPtr].LiveBody + "</div>");
							}
						}
						result = result + "<td valign=bottom>";
						if (Tabs[TabPtr].IsHit && (HitPtr == 0))
						{
							HitPtr = TabPtr;
							//
							// This tab is hit
							//
							result = result + "<table cellspacing=0 cellPadding=0 border=0>";
							result = result + "<tr>" + "<td colspan=2 height=1 width=2></td>" + "<td colspan=1 height=1 bgcolor=black></td>" + "<td colspan=3 height=1 width=3></td>" + "</tr>";
							result = result + "<tr>" + "<td colspan=1 height=1 width=1></td>" + "<td colspan=1 height=1 width=1 bgcolor=black></td>" + "<td Class=\"ccTabHit\" colspan=1 height=1></td>" + "<td colspan=1 height=1 width=1 bgcolor=black></td>" + "<td colspan=2 height=1 width=2></td>" + "</tr>";
							result = result + "<tr>" + "<td colspan=1 height=2 bgcolor=black></td>" + "<td Class=\"ccTabHit\" colspan=1 height=2></td>" + "<td Class=\"ccTabHit\" colspan=1 height=2></td>" + "<td Class=\"ccTabHit\" colspan=1 height=2></td>" + "<td colspan=1 height=2 bgcolor=black></td>" + "<td colspan=1 height=2></td>" + "</tr>";
							result = result + "<tr>" + "<td bgcolor=black></td>" + "<td Class=\"ccTabHit\"></td>" + "<td Class=\"ccTabHit\">" + "<table cellspacing=0 cellPadding=2 border=0><tr>" + "<td Class=\"ccTabHit\">&nbsp;<a href=\"" + TabLink + "\" Class=\"ccTabHit\">" + Tabs[TabPtr].Caption + "</a>&nbsp;</td>" + "</tr></table >" + "</td>" + "<td Class=\"ccTabHit\"></td>" + "<td bgcolor=black></td>" + "<td></td>" + "</tr>";
							result = result + "<tr>" + "<td bgcolor=black></td>" + "<td Class=\"ccTabHit\"></td>" + "<td Class=\"ccTabHit\"></td>" + "<td Class=\"ccTabHit\"></td>" + "<td bgcolor=black></td>" + "<td bgcolor=black></td>" + "</tr>" + "</table >";
						}
						else
						{
							//
							// This tab is not hit
							//
							result = result + "<table cellspacing=0 cellPadding=0 border=0>";
							result = result + "<tr>" + "<td colspan=6 height=1></td>" + "</tr>";
							result = result + "<tr>" + "<td colspan=2 height=1></td>" + "<td colspan=1 height=1 bgcolor=black></td>" + "<td colspan=3 height=1></td>" + "</tr>";
							result = result + "<tr>" + "<td width=1></td>" + "<td width=1 bgcolor=black></td>" + "<td Class=\"ccTab\"></td>" + "<td width=1 bgcolor=black></td>" + "<td width=2 colspan=2></td>" + "</tr>";
							result = result + "<tr>" + "<td width=1 bgcolor=black></td>" + "<td width=1 Class=\"ccTab\"></td>" + "<td nowrap Class=\"ccTab\">" + "<table cellspacing=0 cellPadding=2 border=0><tr>" + "<td Class=\"ccTab\">&nbsp;<a href=\"" + TabLink + "\" Class=\"ccTab\">" + Tabs[TabPtr].Caption + "</a>&nbsp;</td>" + "</tr></table >" + "</td>" + "<td width=1 Class=\"ccTab\"></td>" + "<td width=1 bgcolor=black></td>" + "<td width=1></td>" + "</tr>";
							result = result + "<tr>" + "<td colspan=6 height=1 bgcolor=black></td>" + "</tr>" + "</table >";
						}
						result = result + "</td>";
					}
					result = result + "<td class=\"ccTabEnd\">&nbsp;</td></tr>";
					if (TabBody.ToString() != "")
					{
						result = result + "<tr><td colspan=6>" + TabBody.ToString() + "</td></tr>";
					}
					result = result + "</tr></table >";
					TabsCnt = 0;
				}
				//
			}
			catch (System.Exception excep)
			{
				//
				//UPGRADE_WARNING: (2081) Err.Number has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2081.aspx
				throw new System.Exception(Information.Err().Number.ToString() + ", " + excep.Source + ", " + "Error in ccGetTabs-" + excep.Message);
			}
			return result;
		}
		//
		//
		//
		internal static string ConvertLinksToAbsolute(string Source, string RootLink)
		{
			try
			{
				//
				string s = "";
				//
				s = Source;
				//
				s = Strings.Replace(s, " href=\"", " href=\"/", 1, -1, CompareMethod.Text);
				s = Strings.Replace(s, " href=\"/http", " href=\"http", 1, -1, CompareMethod.Text);
				s = Strings.Replace(s, " href=\"/mailto", " href=\"mailto", 1, -1, CompareMethod.Text);
				s = Strings.Replace(s, " href=\"//", " href=\"" + RootLink, 1, -1, CompareMethod.Text);
				s = Strings.Replace(s, " href=\"/?", " href=\"" + RootLink + "?", 1, -1, CompareMethod.Text);
				s = Strings.Replace(s, " href=\"/", " href=\"" + RootLink, 1, -1, CompareMethod.Text);
				//
				s = Strings.Replace(s, " href=", " href=/", 1, -1, CompareMethod.Text);
				s = Strings.Replace(s, " href=/\"", " href=\"", 1, -1, CompareMethod.Text);
				s = Strings.Replace(s, " href=/http", " href=http", 1, -1, CompareMethod.Text);
				s = Strings.Replace(s, " href=//", " href=" + RootLink, 1, -1, CompareMethod.Text);
				s = Strings.Replace(s, " href=/?", " href=" + RootLink + "?", 1, -1, CompareMethod.Text);
				s = Strings.Replace(s, " href=/", " href=" + RootLink, 1, -1, CompareMethod.Text);
				//
				s = Strings.Replace(s, " src=\"", " src=\"/", 1, -1, CompareMethod.Text);
				s = Strings.Replace(s, " src=\"/http", " src=\"http", 1, -1, CompareMethod.Text);
				s = Strings.Replace(s, " src=\"//", " src=\"" + RootLink, 1, -1, CompareMethod.Text);
				s = Strings.Replace(s, " src=\"/?", " src=\"" + RootLink + "?", 1, -1, CompareMethod.Text);
				s = Strings.Replace(s, " src=\"/", " src=\"" + RootLink, 1, -1, CompareMethod.Text);
				//
				s = Strings.Replace(s, " src=", " src=/", 1, -1, CompareMethod.Text);
				s = Strings.Replace(s, " src=/\"", " src=\"", 1, -1, CompareMethod.Text);
				s = Strings.Replace(s, " src=/http", " src=http", 1, -1, CompareMethod.Text);
				s = Strings.Replace(s, " src=//", " src=" + RootLink, 1, -1, CompareMethod.Text);
				s = Strings.Replace(s, " src=/?", " src=" + RootLink + "?", 1, -1, CompareMethod.Text);
				s = Strings.Replace(s, " src=/", " src=" + RootLink, 1, -1, CompareMethod.Text);
				//
				//
				return s;
			}
			catch (System.Exception excep)
			{
				//
				//UPGRADE_WARNING: (2081) Err.Number has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2081.aspx
				throw new System.Exception(Information.Err().Number.ToString() + ", " + excep.Source + ", " + "Error in ConvertLinksToAbsolute-" + excep.Message);
			}
		}
		//
		//
		//
		internal static string GetProgramPath()
		{
			string result = "";
			result = Path.GetDirectoryName(Application.ExecutablePath);
			if ((result.IndexOf("c:\\h\\contensive\\", StringComparison.CurrentCultureIgnoreCase) + 1) != 0)
			{
				result = "c:\\h\\Contensive";
			}
			else if ((result.IndexOf("c:\\release\\", StringComparison.CurrentCultureIgnoreCase) + 1) != 0)
			{ 
				result = "c:\\h\\Contensive";
			}
			return result;
		}
		//
		//
		//
		internal static string GetAddonRootPath()
		{
			string result = "";
			result = GetProgramPath();
			if ((result.IndexOf("c:\\h\\contensive", StringComparison.CurrentCultureIgnoreCase) + 1) != 0)
			{
				//
				// debugging - change program path to dummy path so addon builds all copy to
				//
				result = "c:\\program files\\kma\\contensive";
			}
			return result + "\\addons";
		}
		//
		//
		//
		internal static string GetHTMLComment(string Comment)
		{
			return "<!-- " + Comment + " -->";
		}
		//
		//
		//
		internal static string[] SplitCRLF(ref string Expression)
		{
			string[] result = null;
			//
			if (Expression.IndexOf(Environment.NewLine) >= 0)
			{
				return (string[]) Strings.Split(Expression, Environment.NewLine, -1, CompareMethod.Text);
			}
			else if (Expression.IndexOf("\r") >= 0)
			{ 
				return (string[]) Strings.Split(Expression, "\r", -1, CompareMethod.Text);
			}
			else if (Expression.IndexOf(Constants.vbLf) >= 0)
			{ 
				return (string[]) Strings.Split(Expression, Constants.vbLf, -1, CompareMethod.Text);
			}
			else
			{
				result = new string[]{""};
				return (string[]) Expression.Split(Environment.NewLine[0]);
			}
			return result;
		}
		//
		//
		//
		internal static void kmaShell(string Cmd, ProcessWindowStyle eWindowStyle, ref bool WaitForReturn)
		{
			try
			{
				//
				IWshRuntimeLibrary.WshShell ShellObj = null;
				//
				ShellObj = new IWshRuntimeLibrary.WshShell();
				if (ShellObj != null)
				{
					object tempRefParam = 0;
					object tempRefParam2 = WaitForReturn;
					ShellObj.Run(Cmd, ref tempRefParam, ref tempRefParam2);
					WaitForReturn = ReflectionHelper.GetPrimitiveValue<bool>(tempRefParam2);
				}
				ShellObj = null;
			}
			catch
			{
				//
				kmaCommonModule.AppendLogFile("ErrorTrap, kmaShell running command [" + Cmd + "], WaitForReturn=" + WaitForReturn.ToString() + ", err=" + kmaCommonModule.GetErrString(Information.Err()));
			}
		}

		internal static void kmaShell(string Cmd, ProcessWindowStyle eWindowStyle)
		{
			bool tempRefParam3 = false;
			kmaShell(Cmd, eWindowStyle, ref tempRefParam3);
		}

		internal static void kmaShell(string Cmd)
		{
			bool tempRefParam4 = false;
			kmaShell(Cmd, ProcessWindowStyle.Hidden, ref tempRefParam4);
		}
		//
		//------------------------------------------------------------------------------------------------------------
		//   Encodes an argument in an Addon OptionString (QueryString) for all non-allowed characters
		//       call this before parsing them together
		//       call decodeAddonConstructorArgument after parsing them apart
		//
		//       Arg0,Arg1,Arg2,Arg3,Name=Value&Name=VAlue[Option1|Option2]
		//
		//       This routine is needed for all Arg, Name, Value, Option values
		//
		//------------------------------------------------------------------------------------------------------------
		//
		internal static string EncodeAddonConstructorArgument(string Arg)
		{
			string result = "";
			string a = "";
			if (Arg != "")
			{
				a = Arg;
				if (true)
				{
					//If AddonNewParse Then
					a = Strings.Replace(a, "\\", "\\\\", 1, -1, CompareMethod.Binary);
					a = Strings.Replace(a, Environment.NewLine, "\\n", 1, -1, CompareMethod.Binary);
					a = Strings.Replace(a, "\t", "\\t", 1, -1, CompareMethod.Binary);
					a = Strings.Replace(a, "&", "\\&", 1, -1, CompareMethod.Binary);
					a = Strings.Replace(a, "=", "\\=", 1, -1, CompareMethod.Binary);
					a = Strings.Replace(a, ",", "\\,", 1, -1, CompareMethod.Binary);
					a = Strings.Replace(a, "\"", "\\\"", 1, -1, CompareMethod.Binary);
					a = Strings.Replace(a, "'", "\\'", 1, -1, CompareMethod.Binary);
					a = Strings.Replace(a, "|", "\\|", 1, -1, CompareMethod.Binary);
					a = Strings.Replace(a, "[", "\\[", 1, -1, CompareMethod.Binary);
					a = Strings.Replace(a, "]", "\\]", 1, -1, CompareMethod.Binary);
					a = Strings.Replace(a, ":", "\\:", 1, -1, CompareMethod.Binary);
				}
				result = a;
			}
			return result;
		}
		//
		//------------------------------------------------------------------------------------------------------------
		//   Decodes an argument parsed from an AddonConstructorString for all non-allowed characters
		//       AddonConstructorString is a & delimited string of name=value[selector]descriptor
		//
		//       to get a value from an AddonConstructorString, first use getargument() to get the correct value[selector]descriptor
		//       then remove everything to the right of any '['
		//
		//       call encodeAddonConstructorargument before parsing them together
		//       call decodeAddonConstructorArgument after parsing them apart
		//
		//       Arg0,Arg1,Arg2,Arg3,Name=Value&Name=VAlue[Option1|Option2]
		//
		//       This routine is needed for all Arg, Name, Value, Option values
		//
		//------------------------------------------------------------------------------------------------------------
		//
		internal static string DecodeAddonConstructorArgument(string EncodedArg)
		{
			//
			string a = EncodedArg;
			if (true)
			{
				//If AddonNewParse Then
				a = Strings.Replace(a, "\\:", ":", 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, "\\]", "]", 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, "\\[", "[", 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, "\\|", "|", 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, "\\'", "'", 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, "\\\"", "\"", 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, "\\,", ",", 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, "\\=", "=", 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, "\\&", "&", 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, "\\t", "\t", 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, "\\n", Environment.NewLine, 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, "\\\\", "\\", 1, -1, CompareMethod.Binary);
			}
			return a;
		}
		//
		//------------------------------------------------------------------------------------------------------------
		//   use only internally
		//
		//   encode an argument to be used in a name=value& (N-V-A) string
		//
		//   an argument can be any one of these is this format:
		//       Arg0,Arg1,Arg2,Arg3,Name=Value&Name=Value[Option1|Option2]descriptor
		//
		//   to create an nva string
		//       string = encodeNvaArgument( name ) & "=" & encodeNvaArgument( value ) & "&"
		//
		//   to decode an nva string
		//       split on ampersand then on equal, and decodeNvaArgument() each part
		//
		//------------------------------------------------------------------------------------------------------------
		//
		internal static string encodeNvaArgument(string Arg)
		{
			string a = Arg;
			if (a != "")
			{
				a = Strings.Replace(a, Environment.NewLine, "#0013#", 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, Constants.vbLf, "#0013#", 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, "\r", "#0013#", 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, "&", "#0038#", 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, "=", "#0061#", 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, ",", "#0044#", 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, "\"", "#0034#", 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, "'", "#0039#", 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, "|", "#0124#", 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, "[", "#0091#", 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, "]", "#0093#", 1, -1, CompareMethod.Binary);
				a = Strings.Replace(a, ":", "#0058#", 1, -1, CompareMethod.Binary);
			}
			return a;
		}
		//
		//------------------------------------------------------------------------------------------------------------
		//   use only internally
		//       decode an argument removed from a name=value& string
		//       see encodeNvaArgument for details on how to use this
		//------------------------------------------------------------------------------------------------------------
		//
		internal static string decodeNvaArgument(string EncodedArg)
		{
			//
			string a = EncodedArg;
			a = Strings.Replace(a, "#0058#", ":", 1, -1, CompareMethod.Binary);
			a = Strings.Replace(a, "#0093#", "]", 1, -1, CompareMethod.Binary);
			a = Strings.Replace(a, "#0091#", "[", 1, -1, CompareMethod.Binary);
			a = Strings.Replace(a, "#0124#", "|", 1, -1, CompareMethod.Binary);
			a = Strings.Replace(a, "#0039#", "'", 1, -1, CompareMethod.Binary);
			a = Strings.Replace(a, "#0034#", "\"", 1, -1, CompareMethod.Binary);
			a = Strings.Replace(a, "#0044#", ",", 1, -1, CompareMethod.Binary);
			a = Strings.Replace(a, "#0061#", "=", 1, -1, CompareMethod.Binary);
			a = Strings.Replace(a, "#0038#", "&", 1, -1, CompareMethod.Binary);
			a = Strings.Replace(a, "#0013#", Environment.NewLine, 1, -1, CompareMethod.Binary);
			return a;
		}
		//
		// returns true of the link is a valid link on the source host
		//
		internal static bool IsLinkToThisHost(string Host, ref string Link)
		{
			//
			bool result = false;
			string LinkHost = "";
			int Pos = 0;
			//
			if (Link.Trim() == "")
			{
				//
				// Blank is not a link
				//
				result = false;
			}
			else if (Link.IndexOf("://") >= 0)
			{ 
				//
				// includes protocol, may be link to another site
				//
				LinkHost = Link.ToLower();
				Pos = 1;
				Pos = Strings.InStr(Pos, LinkHost, "://", CompareMethod.Binary);
				if (Pos > 0)
				{
					Pos = Strings.InStr(Pos + 3, LinkHost, "/", CompareMethod.Binary);
					if (Pos > 0)
					{
						LinkHost = LinkHost.Substring(0, Math.Min(Pos - 1, LinkHost.Length));
					}
					result = (Host.ToLower() == LinkHost);
					if (!result)
					{
						//
						// try combinations including/excluding www.
						//
						if ((LinkHost.IndexOf("www.", StringComparison.CurrentCultureIgnoreCase) + 1) != 0)
						{
							//
							// remove it
							//
							LinkHost = Strings.Replace(LinkHost, "www.", "", 1, -1, CompareMethod.Text);
							result = (Host.ToLower() == LinkHost);
						}
						else
						{
							//
							// add it
							//
							LinkHost = Strings.Replace(LinkHost, "://", "://www.", 1, -1, CompareMethod.Text);
							result = (Host.ToLower() == LinkHost);
						}
					}
				}
			}
			else if ((Link.IndexOf('#') + 1) == 1)
			{ 
				//
				// Is a bookmark, not a link
				//
				result = false;
			}
			else
			{
				//
				// all others are links on the source
				//
				result = true;
			}
			if (!result)
			{
				Link = Link;
			}
			return result;
		}
		//
		//========================================================================================================
		//   ConvertLinkToRootRelative
		//
		//   /images/logo-main.jpg with any Basepath to /images/logo-main.jpg
		//   http://gcm.brandeveolve.com/images/logo-main.jpg with any BasePath  to /images/logo-main.jpg
		//   images/logo-main.jpg with Basepath '/' to /images/logo-main.jpg
		//   logo-main.jpg with Basepath '/images/' to /images/logo-main.jpg
		//
		//========================================================================================================
		//
		internal static string ConvertLinkToRootRelative(ref string Link, string BasePath)
		{
			//
			string result = "";
			int Pos = 0;
			//
			result = Link;
			if ((Link.IndexOf('/') + 1) == 1)
			{
				//
				//   case /images/logo-main.jpg with any Basepath to /images/logo-main.jpg
				//
			}
			else if (Link.IndexOf("://") >= 0)
			{ 
				//
				//   case http://gcm.brandeveolve.com/images/logo-main.jpg with any BasePath  to /images/logo-main.jpg
				//
				Pos = (Link.IndexOf("://") + 1);
				if (Pos > 0)
				{
					Pos = Strings.InStr(Pos + 3, Link, "/", CompareMethod.Binary);
					if (Pos > 0)
					{
						result = Link.Substring(Pos - 1);
					}
					else
					{
						//
						// This is just the domain name, RootRelative is the root
						//
						result = "/";
					}
				}
			}
			else
			{
				//
				//   case images/logo-main.jpg with Basepath '/' to /images/logo-main.jpg
				//   case logo-main.jpg with Basepath '/images/' to /images/logo-main.jpg
				//
				result = BasePath + Link;
			}
			//
			return result;
		}
		//
		//
		//
		internal static string GetAddonIconImg(string AdminURL, ref int IconWidth, ref int IconHeight, ref int IconSprites, bool IconIsInline, string IconImgID, ref string IconFilename, string serverFilePath, ref string IconAlt, ref string IconTitle, string ACInstanceID, int IconSpriteColumn)
		{
			//
			string result = "";
			//
			if (IconAlt == "")
			{
				IconAlt = "Add-on";
			}
			if (IconTitle == "")
			{
				IconTitle = "Rendered as Add-on";
			}
			if (IconFilename == "")
			{
				//
				// No icon given, use the default
				//
				if (IconIsInline)
				{
					IconFilename = "/ccLib/images/IconAddonInlineDefault.png";
					IconWidth = 62;
					IconHeight = 17;
					IconSprites = 0;
				}
				else
				{
					IconFilename = "/ccLib/images/IconAddonBlockDefault.png";
					IconWidth = 57;
					IconHeight = 59;
					IconSprites = 4;
				}
			}
			else if (IconFilename.IndexOf("://") >= 0)
			{ 
				//
				// icon is an Absolute URL - leave it
				//
			}
			else if (IconFilename.StartsWith("/"))
			{ 
				//
				// icon is Root Relative, leave it
				//
			}
			else
			{
				//
				// icon is a virtual file, add the serverfilepath
				//
				IconFilename = serverFilePath + IconFilename;
			}
			//IconFilename = kmaEncodeJavascript(IconFilename)
			if ((IconWidth == 0) || (IconHeight == 0))
			{
				IconSprites = 0;
			}

			if (IconSprites == 0)
			{
				//
				// just the icon
				//
				result = "<img" + " border=0" + " id=\"" + IconImgID + "\"" + " onDblClick=\"window.parent.OpenAddonPropertyWindow(this,'" + AdminURL + "');\"" + " alt=\"" + IconAlt + "\"" + " title=\"" + IconTitle + "\"" + " src=\"" + IconFilename + "\"";
				//GetAddonIconImg = "<img" _
				//'    & " id=""AC,AGGREGATEFUNCTION,0," & FieldName & "," & ArgumentList & """" _
				//'    & " onDblClick=""window.parent.OpenAddonPropertyWindow(this);""" _
				//'    & " alt=""" & IconAlt & """" _
				//'    & " title=""" & IconTitle & """" _
				//'    & " src=""" & IconFilename & """"
				if (IconWidth != 0)
				{
					result = result + " width=\"" + IconWidth.ToString() + "px\"";
				}
				if (IconHeight != 0)
				{
					result = result + " height=\"" + IconHeight.ToString() + "px\"";
				}
				if (IconIsInline)
				{
					result = result + " style=\"vertical-align:middle;display:inline;\" ";
				}
				else
				{
					result = result + " style=\"display:block\" ";
				}
				if (ACInstanceID != "")
				{
					result = result + " ACInstanceID=\"" + ACInstanceID + "\"";
				}
				return result + ">";
			}
			else
			{
				//
				// Sprite Icon
				//
				return GetIconSprite(IconImgID, IconSpriteColumn, IconFilename, IconWidth, IconHeight, IconAlt, IconTitle, "window.parent.OpenAddonPropertyWindow(this,'" + AdminURL + "');", IconIsInline, ACInstanceID);
				//        GetAddonIconImg = "<img" _
				//'            & " border=0" _
				//'            & " id=""" & IconImgID & """" _
				//'            & " onMouseOver=""this.style.backgroundPosition='" & (-1 * IconSpriteColumn * IconWidth) & "px -" & (2 * IconHeight) & "px'""" _
				//'            & " onMouseOut=""this.style.backgroundPosition='" & (-1 * IconSpriteColumn * IconWidth) & "px 0px'""" _
				//'            & " onDblClick=""window.parent.OpenAddonPropertyWindow(this,'" & AdminURL & "');""" _
				//'            & " alt=""" & IconAlt & """" _
				//'            & " title=""" & IconTitle & """" _
				//'            & " src=""/ccLib/images/spacer.gif"""
				//        ImgStyle = "background:url(" & IconFilename & ") " & (-1 * IconSpriteColumn * IconWidth) & "px 0px no-repeat;"
				//        ImgStyle = ImgStyle & "width:" & IconWidth & "px;"
				//        ImgStyle = ImgStyle & "height:" & IconHeight & "px;"
				//        If IconIsInline Then
				//            'GetAddonIconImg = GetAddonIconImg & " align=""middle"""
				//            ImgStyle = ImgStyle & "vertical-align:middle;display:inline;"
				//        Else
				//            ImgStyle = ImgStyle & "display:block;"
				//        End If
				//
				//
				//        'Return_IconStyleMenuEntries = Return_IconStyleMenuEntries & vbCrLf & ",["".icon" & AddonID & """,false,"".icon" & AddonID & """,""background:url(" & IconFilename & ") 0px 0px no-repeat;"
				//        'GetAddonIconImg = "<img" _
				//'        '    & " id=""AC,AGGREGATEFUNCTION,0," & FieldName & "," & ArgumentList & """" _
				//'        '    & " onMouseOver=""this.style.backgroundPosition=\'0px -" & (2 * IconHeight) & "px\'""" _
				//'        '    & " onMouseOut=""this.style.backgroundPosition=\'0px 0px\'""" _
				//'        '    & " onDblClick=""window.parent.OpenAddonPropertyWindow(this);""" _
				//'        '    & " alt=""" & IconAlt & """" _
				//'        '    & " title=""" & IconTitle & """" _
				//'        '    & " src=""/ccLib/images/spacer.gif"""
				//        If ACInstanceID <> "" Then
				//            GetAddonIconImg = GetAddonIconImg & " ACInstanceID=""" & ACInstanceID & """"
				//        End If
				//        GetAddonIconImg = GetAddonIconImg & " style=""" & ImgStyle & """>"
				//        'Return_IconStyleMenuEntries = Return_IconStyleMenuEntries & """]"
			}
			return result;
		}
		//
		//
		//
		internal static string ConvertRSTypeToGoogleType(int RecordFieldType)
		{
			switch(RecordFieldType)
			{
				case 2 : case 3 : case 4 : case 5 : case 6 : case 14 : case 16 : case 17 : case 18 : case 19 : case 20 : case 21 : case 131 : 
					return "number";
				default:
					return "string";
			}
		}

		//
		//========================================================================
		//   HandleError
		//       Logs the error and either resumes next, or raises it to the next level
		//========================================================================
		//
		internal static void AppendLogFile2(string ContensiveAppName, string Context, string ProgramName, string ClassName, string MethodName, int ErrNumber, string ErrSource, string ErrDescription, bool ErrorTrap, bool ResumeNextAfterLogging, string URL, string LogFolder, string LogNamePrefix)
		{
			//UPGRADE_TODO: (1065) Error handling statement (On Error Goto) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
			UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Goto Label (ErrorTrap)");
			//
			int MonthNumber = 0;
			int DayNumber = 0;
			string FilenameNoExt = "";
			kmaFileSystem3.FileSystemClass kmafs = null;
			string ErrorMessage = "";
			string LogLine = "";
			string ResumeMessage = "";
			string FolderFileList = "";
			string[] FolderFiles = null;
			string PathFilenameNoExt = "";
			string[] FileDetails = null;
			int fileSize = 0;
			int RetryCnt = 0;
			bool SaveOK = false;
			string FileSuffix = "";
			string iLogFolder = "";
			//
			iLogFolder = LogFolder;
			//
			if (ErrorTrap)
			{
				ErrorMessage = "Error Trap";
			}
			else
			{
				ErrorMessage = "Log Entry";
			}
			//
			if (ResumeNextAfterLogging)
			{
				ResumeMessage = "Resume after logging";
			}
			else
			{
				ResumeMessage = "Abort after logging";
			}
			//
			LogLine = "" + LogFileCopyPrep(DateTime.Now.ToString()) + "," + LogFileCopyPrep(ContensiveAppName) + "," + LogFileCopyPrep(ProgramName) + "," + LogFileCopyPrep(ClassName) + "," + LogFileCopyPrep(MethodName) + "," + LogFileCopyPrep(Context) + "," + LogFileCopyPrep(ErrorMessage) + "," + LogFileCopyPrep(ResumeMessage) + "," + LogFileCopyPrep(ErrSource) + "," + LogFileCopyPrep(ErrNumber.ToString()) + "," + LogFileCopyPrep(ErrDescription) + "," + LogFileCopyPrep(URL) + Environment.NewLine;
			//
			DayNumber = DateAndTime.Day(DateTime.Now);
			MonthNumber = DateTime.Now.Month;
			FilenameNoExt = DateTime.Now.Year.ToString();
			if (MonthNumber < 10)
			{
				FilenameNoExt = FilenameNoExt + "0";
			}
			FilenameNoExt = FilenameNoExt + MonthNumber.ToString();
			if (DayNumber < 10)
			{
				FilenameNoExt = FilenameNoExt + "0";
			}
			FilenameNoExt = LogNamePrefix + FilenameNoExt + DayNumber.ToString();
			if (iLogFolder != "")
			{
				iLogFolder = iLogFolder + "\\";
			}
			iLogFolder = GetProgramPath() + "\\logs\\" + iLogFolder;
			PathFilenameNoExt = iLogFolder + FilenameNoExt;
			//
			kmafs = new kmaFileSystem3.FileSystemClass();
			FolderFileList = kmafs.GetFolderFiles2(ref iLogFolder);
			FolderFiles = (string[]) FolderFileList.Split(Environment.NewLine[0]);
			foreach (string FolderFiles_item in FolderFiles)
			{
				if ((FolderFiles_item.IndexOf(FilenameNoExt + ".log" + ",", StringComparison.CurrentCultureIgnoreCase) + 1) != 0)
				{
					FileDetails = (string[]) FolderFiles_item.Split('\t');
					fileSize = kmaCommonModule.kmaEncodeInteger(FileDetails[5]);
					break;
				}
			}
			if (fileSize < 10000000)
			{
				RetryCnt = 0;
				SaveOK = false;
				FileSuffix = "";
				//UPGRADE_TODO: (1065) Error handling statement (On Error Resume Next) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
				UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Resume Next");

				while((!SaveOK) && (RetryCnt < 10))
				{
					SaveOK = true;
					string tempRefParam = (PathFilenameNoExt + FileSuffix + ".log").ToLower();
					kmafs.AppendFile(ref tempRefParam, ref LogLine);
					//UPGRADE_WARNING: (2081) Err.Number has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2081.aspx
					if (Information.Err().Number != 0)
					{
						//UPGRADE_WARNING: (2081) Err.Number has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2081.aspx
						if (Information.Err().Number == 70)
						{
							//
							// permission denied - happens when more then one process are writing at once, go to the next suffix
							//
							FileSuffix = "-" + (RetryCnt + 1).ToString();
							SaveOK = false;
						}
						else
						{
							//
							// ignore all other errors - this routine logs errors, so there is nothing to do if it fails
							//
						}
						RetryCnt++;
						Information.Err().Clear();
					}
				};
			}
			return;
			//
ErrorTrap:
			Information.Err().Clear();
		}
		//
		//========================================================================
		//   HandleError
		//       Logs the error and either resumes next, or raises it to the next level
		//========================================================================
		//
		internal static void HandleError2(string ContensiveAppName, string Context, string ProgramName, string ClassName, string MethodName, int ErrNumber, string ErrSource, string ErrDescription, bool ErrorTrap, bool ResumeNext, string URL)
		{
			//
			AppendLogFile2(ContensiveAppName, Context, ProgramName, ClassName, MethodName, ErrNumber, ErrSource, ErrDescription, ErrorTrap, ResumeNext, URL, "", "Trace");
			//
			if (!ResumeNext)
			{
				UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Goto Label (0)");
				if (ErrNumber == 0)
				{
					throw new System.Exception(kmaCommonModule.KmaErrorInternal.ToString() + ", " + ErrSource + ", " + Context);
				}
				else
				{
					throw new System.Exception(ErrNumber.ToString() + ", " + ErrSource + ", " + ErrDescription);
				}
			}
			//
		}
		//
		//
		//
		private static string LogFileCopyPrep(string Source)
		{
			string Copy = Source;
			Copy = Strings.Replace(Copy, Environment.NewLine, " ", 1, -1, CompareMethod.Binary);
			Copy = Strings.Replace(Copy, Constants.vbLf, " ", 1, -1, CompareMethod.Binary);
			Copy = Strings.Replace(Copy, "\r", " ", 1, -1, CompareMethod.Binary);
			Copy = Strings.Replace(Copy, "\"", "\"\"", 1, -1, CompareMethod.Binary);
			Copy = "\"" + Copy + "\"";
			return Copy;
		}
		// moved to csv
		//'
		//'=================================================================================================================
		//'   GetAddonOptionStringValue
		//'
		//'   gets the value from a list matching the name
		//'
		//'   InstanceOptionstring is an "AddonEncoded" name=AddonEncodedValue[selector]descriptor&name=value string
		//'=================================================================================================================
		//'
		//Public Function GetAddonOptionStringValue(OptionName As String, AddonOptionString As String) As String
		//    On Error GoTo ErrorTrap
		//    '
		//    Dim Pos As Long
		//    Dim s As String
		//    '
		//    s = GetArgument(OptionName, AddonOptionString, "", "&")
		//    Pos = InStr(1, s, "[")
		//    If Pos > 0 Then
		//        s = Left(s, Pos - 1)
		//    End If
		//    s = decodeNvaArgument(s)
		//    '
		//    GetAddonOptionStringValue = Trim(s)
		//    '
		//    Exit Function
		//ErrorTrap:
		//    Call HandleError2("", "", App.EXEName, "ccCommonModule", "GetAddonOptionStringValue", Err.Number, Err.Source, Err.Description, True, False, "")
		//End Function
		//
		//
		//
		internal static string GetIconSprite(string TagID, int SpriteColumn, string IconSrc, int IconWidth, int IconHeight, string IconAlt, string IconTitle, string onDblClick, bool IconIsInline, string ACInstanceID)
		{
			//
			//
			string result = "";
			result = "<img" + " border=0" + " id=\"" + TagID + "\"" + " onMouseOver=\"this.style.backgroundPosition='" + ((-1 * SpriteColumn * IconWidth).ToString()) + "px -" + ((2 * IconHeight).ToString()) + "px';\"" + " onMouseOut=\"this.style.backgroundPosition='" + ((-1 * SpriteColumn * IconWidth).ToString()) + "px 0px'\"" + " onDblClick=\"" + onDblClick + "\"" + " alt=\"" + IconAlt + "\"" + " title=\"" + IconTitle + "\"" + " src=\"/ccLib/images/spacer.gif\"";
			string ImgStyle = "background:url(" + IconSrc + ") " + ((-1 * SpriteColumn * IconWidth).ToString()) + "px 0px no-repeat;";
			ImgStyle = ImgStyle + "width:" + IconWidth.ToString() + "px;";
			ImgStyle = ImgStyle + "height:" + IconHeight.ToString() + "px;";
			if (IconIsInline)
			{
				ImgStyle = ImgStyle + "vertical-align:middle;display:inline;";
			}
			else
			{
				ImgStyle = ImgStyle + "display:block;";
			}
			if (ACInstanceID != "")
			{
				result = result + " ACInstanceID=\"" + ACInstanceID + "\"";
			}
			return result + " style=\"" + ImgStyle + "\">";
		}
		//
		//
		//
		internal static string RegGetValue(int MainKey, ref string SubKey, ref string Value)
		{
			// MainKey must be one of the Publicly declared HKEY constants.
			string result = "";
			int sKeyType = 0; //to return the key type.  This function expects REG_SZ or REG_DWORD
			int ret = 0; //returned by registry functions, should be 0&
			int lpHKey = 0; //return handle to opened key
			int lpcbData = 0; //length of data in returned string
			string ReturnedString = ""; //returned string value
			int ReturnedLong = 0; //returned long value
			if (MainKey >= 0x80000000 && MainKey <= 0x80000006)
			{
				// Open key
				int tempRefParam = 0;
				ret = UpgradeSolution1Support.PInvoke.SafeNative.advapi32.RegOpenKeyExA(MainKey, ref SubKey, ref tempRefParam, KEY_READ, ref lpHKey);
				if (ret != ERROR_SUCCESS)
				{
					return ""; //No key open, so leave
				}

				// Set up buffer for data to be returned in.
				// Adjust next value for larger buffers.
				lpcbData = 255;
				ReturnedString = new String(' ', lpcbData);

				// Read key
				ret = UpgradeSolution1Support.PInvoke.SafeNative.advapi32.RegQueryValueExA(lpHKey, ref Value, 0, ref sKeyType, ref ReturnedString, ref lpcbData);
				if (ret != ERROR_SUCCESS)
				{
					result = ""; //Value probably doesn't exist
				}
				else
				{
					if (sKeyType == REG_DWORD)
					{
						int tempRefParam2 = 4;
						ret = UpgradeSolution1Support.PInvoke.SafeNative.advapi32.RegQueryValueEx(lpHKey, ref Value, 0, ref sKeyType, ref ReturnedLong, ref tempRefParam2);
						if (ret == ERROR_SUCCESS)
						{
							result = ReturnedLong.ToString();
						}
					}
					else
					{
						result = ReturnedString.Substring(0, Math.Min(lpcbData - 1, ReturnedString.Length));
					}
				}
				// Always close opened keys.
				ret = UpgradeSolution1Support.PInvoke.SafeNative.advapi32.RegCloseKey(lpHKey);
			}
			return result;
		}
	}
}