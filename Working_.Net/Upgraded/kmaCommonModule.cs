using Microsoft.VisualBasic;
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using UpgradeHelpers.Helpers;

namespace aoCExport
{
	internal static class kmaCommonModule
	{


		//
		//========================================================================
		//   kma defined errors
		//       1000-1999 Contensive
		//       2000-2999 Datatree
		//
		//   see kmaErrorDescription() for transations
		//========================================================================
		//
		const int Error_DataTree_RootNodeNext = 2000;
		const int Error_DataTree_NoGoNext = 2001;
		//
		//========================================================================
		//
		//========================================================================
		//
		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static int GetTickCount();
		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static int GetCurrentProcessId();
		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static void Sleep(int dwMilliseconds);
		//
		//========================================================================
		//   Declarations for SetPiorityClass
		//========================================================================
		//
		const int THREAD_BASE_PRIORITY_IDLE = -15;
		const int THREAD_BASE_PRIORITY_LOWRT = 15;
		const int THREAD_BASE_PRIORITY_MIN = -2;
		const int THREAD_BASE_PRIORITY_MAX = 2;
		const int THREAD_PRIORITY_LOWEST = THREAD_BASE_PRIORITY_MIN;
		const int THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX;
		const int THREAD_PRIORITY_BELOW_NORMAL = (THREAD_PRIORITY_LOWEST + 1);
		const int THREAD_PRIORITY_ABOVE_NORMAL = (THREAD_PRIORITY_HIGHEST - 1);
		const int THREAD_PRIORITY_IDLE = THREAD_BASE_PRIORITY_IDLE;
		const int THREAD_PRIORITY_NORMAL = 0;
		const int THREAD_PRIORITY_TIME_CRITICAL = THREAD_BASE_PRIORITY_LOWRT;
		const int HIGH_PRIORITY_CLASS = 0x80;
		const int IDLE_PRIORITY_CLASS = 0x40;
		const int NORMAL_PRIORITY_CLASS = 0x20;
		const int REALTIME_PRIORITY_CLASS = 0x100;
		//
		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static int SetThreadPriority(int hThread, int nPriority);
		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static int SetPriorityClass(int hProcess, int dwPriorityClass);
		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static int GetThreadPriority(int hThread);
		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static int GetPriorityClass(int hProcess);
		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static int GetCurrentThread();
		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static int GetCurrentProcess();
		//

		//
		//========================================================================
		//Converts unsafe characters,
		//such as spaces, into their
		//corresponding escape sequences.
		//========================================================================
		//
		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("shlwapi.dll", EntryPoint = "UrlEscapeA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static int UrlEscape([MarshalAs(UnmanagedType.VBByRefStr)] ref string pszURL, [MarshalAs(UnmanagedType.VBByRefStr)] ref string pszEscaped, ref int pcchEscaped, int dwFlags);
		//
		//Converts escape sequences back into
		//ordinary characters.
		//
		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("shlwapi.dll", EntryPoint = "UrlUnescapeA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static int UrlUnescape([MarshalAs(UnmanagedType.VBByRefStr)] ref string pszURL, [MarshalAs(UnmanagedType.VBByRefStr)] ref string pszUnescaped, ref int pcchUnescaped, int dwFlags);

		//
		//   Error reporting strategy
		//       Popups are pop-up boxes that tell the user what to do
		//       Logs are files with error details for developers to use debugging
		//
		//       Attended Programs
		//           - errors that do not effect the operation, resume next
		//           - all errors trickle up to the user interface level
		//           - User Errors, like file not found, return "UserError" code and a description
		//           - Internal Errors, like div-by-0, User should see no detail, log gets everything
		//           - Dependant Object Errors, codes that return from objects:
		//               - If UserError, translate ErrSource for raise, but log all original info
		//               - If InternalError, log info and raise InternalError
		//               - If you can not tell, call it InternalError
		//
		//       UnattendedMode
		//           The same, except each routine decides when
		//
		//       When an error happens in-line (bad condition without a raise)
		//           Log the error
		//           Raise the appropriate Code/Description in the current Source
		//
		//       When an ErrorTrap occurs
		//           If ErrSource is not AppTitle, it is a dependantObjectError, log and translate code
		//           If ErrNumber is not an ObjectError, call it internal error, log and translate code
		//           Error must be either "InternalError" or "UserError", just raise it again
		//
		// old - If an error is raised that is not a KmaCode, it is logged and translated
		// old - If an error is raised and the soure is not he current App.EXEname, it is logged and translated
		//
		static readonly public int KmaErrorBase = Constants.vbObjectError; // Base on which Internal errors should start
		//
		//Public Const KmaError_UnderlyingObject = vbObjectError + 1     ' An error occurec in an underlying object
		//Public Const KmaccErrorServiceStopped = vbObjectError + 2       ' The service is not running
		//Public Const KmaError_BadObject = vbObjectError + 3            ' The Server Pointer is not valid
		//Public Const KmaError_UpgradeInProgress = vbObjectError + 4    ' page is blocked because an upgrade is in progress
		//Public Const KmaError_InvalidArgument = vbObjectError + 5      ' and input argument is not valid. Put details at end of description
		//
		static readonly public int KmaErrorUser = KmaErrorBase + 16; // Generic Error code that passes the description back to the user
		static readonly public int KmaErrorInternal = KmaErrorBase + 17; // Internal error which the user should not see
		static readonly public int KmaErrorPage = KmaErrorBase + 18; // Error from the page which called Contensive
		//
		static readonly public int KmaObjectError = KmaErrorBase + 256; // Internal error which the user should not see
		//
		//==========================================================================
		//       NTSvc.ocx, LogEvent Constants
		//==========================================================================
		//
		public const int NTServiceEventError = 1; // Error event.
		public const int NTServiceEventWarning = 2; // Warning event.
		public const int NTServiceEventInformation = 4; // Information event.
		public const int NTServiceEventAuditSuccess = 8; // Audit success event.
		public const int NTServiceEventAuditFailure = 10; // Audit failure event.

		public const int NTServiceIDDebug = 108; // Debugging message
		public const int NTServiceIDError = 109; // Error message
		public const int NTServiceIDInfo = 110; // Information message
		//
		//==========================================================================
		//       NTSvc.ocx, LogEvent Constants
		//==========================================================================
		//
		public const string SQLTrue = "1";
		public const string SQLFalse = "0";
		//
		//
		//
		public const string kmaEndTable = "</table >";
		public const string kmaEndTableCell = "</td>";
		public const string kmaEndTableRow = "</tr>";
		//
		//==========================================================================
		// kmaByteArrayToString / kmaStringToByteArray
		//==========================================================================
		//
		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("kernel32.dll", EntryPoint = "RtlMoveMemory", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static void CopyMemory(System.IntPtr hpvDest, System.IntPtr hpvSource, int cbCopy);
		//The WideCharToMultiByte function maps a wide-character string to a new character string.
		//The function is faster when both lpDefaultChar and lpUsedDefaultChar are NULL.
		//CodePage
		private const int CP_ACP = 0; //ANSI
		private const int CP_MACCP = 2; //Mac
		private const int CP_OEMCP = 1; //OEM
		private const int CP_UTF7 = 65000;
		private const int CP_UTF8 = 65001;
		//dwFlags
		private const int WC_NO_BEST_FIT_CHARS = 0x400;
		private const int WC_COMPOSITECHECK = 0x200;
		private const int WC_DISCARDNS = 0x10;
		private const int WC_SEPCHARS = 0x20; //Default
		private const int WC_DEFAULTCHAR = 0x40;

		//UPGRADE_NOTE: (2041) The following line was commented. More Information: http://www.vbtonet.com/ewis/ewi2041.aspx
		//[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		//extern public static int WideCharToMultiByte(int CodePage, int dwFlags, int lpWideCharStr, int cchWideChar, int lpMultiByteStr, int cbMultiByte, int lpDefaultChar, int lpUsedDefaultChar);
		//
		//==========================================================================
		//   Convert a variant to an long (long)
		//   returns 0 if the input is not an integer
		//   if float, rounds to integer
		//==========================================================================
		//
		internal static int kmaEncodeInteger(object ExpressionVariant)
		{
			// 7/14/2009 - cover the overflow case, return 0
			//UPGRADE_TODO: (1069) Error handling statement (On Error Resume Next) was converted to a pattern that might have a different behavior. More Information: http://www.vbtonet.com/ewis/ewi1069.aspx
			int result = 0;
			try
			{
				//
				if (!Information.IsArray(ExpressionVariant))
				{
					//UPGRADE_WARNING: (2065) Boolean method Information.IsMissing has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2065.aspx
					if (ExpressionVariant != null)
					{
						//UPGRADE_WARNING: (1049) Use of Null/IsNull() detected. More Information: http://www.vbtonet.com/ewis/ewi1049.aspx
						if (!Convert.IsDBNull(ExpressionVariant))
						{
							if (ReflectionHelper.GetPrimitiveValue<string>(ExpressionVariant) != "")
							{
								double dbNumericTemp = 0;
								if (Double.TryParse(ReflectionHelper.GetPrimitiveValue<string>(ExpressionVariant), NumberStyles.Float , CultureInfo.CurrentCulture.NumberFormat, out dbNumericTemp))
								{
									result = ReflectionHelper.GetPrimitiveValue<int>(ExpressionVariant);
								}
							}
						}
					}
				}
				//
				return result;
				//
				// ----- ErrorTrap
				//

				Information.Err().Clear();
				result = 0;
			}
			catch (Exception exc)
			{
				NotUpgradedHelper.NotifyNotUpgradedElement("Resume in On-Error-Resume-Next Block");
			}
			return result;
		}
		//
		//==========================================================================
		//   Convert a variant to a number (double)
		//   returns 0 if the input is not a number
		//==========================================================================
		//
		internal static double KmaEncodeNumber(object ExpressionVariant)
		{
			double result = 0;
			try
			{
				//
				//KmaEncodeNumber = 0
				//UPGRADE_WARNING: (2065) Boolean method Information.IsMissing has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2065.aspx
				if (ExpressionVariant != null)
				{
					//UPGRADE_WARNING: (1049) Use of Null/IsNull() detected. More Information: http://www.vbtonet.com/ewis/ewi1049.aspx
					if (!Convert.IsDBNull(ExpressionVariant))
					{
						if (ReflectionHelper.GetPrimitiveValue<string>(ExpressionVariant) != "")
						{
							double dbNumericTemp = 0;
							if (Double.TryParse(ReflectionHelper.GetPrimitiveValue<string>(ExpressionVariant), NumberStyles.Float , CultureInfo.CurrentCulture.NumberFormat, out dbNumericTemp))
							{
								result = ReflectionHelper.GetPrimitiveValue<double>(ExpressionVariant);
							}
						}
					}
				}
				//
			}
			catch
			{
				//
				// ----- ErrorTrap
				//
				result = 0;
			}
			return result;
		}
		//
		//==========================================================================
		//   Convert a variant to a date
		//   returns 0 if the input is not a number
		//==========================================================================
		//
		internal static System.DateTime KmaEncodeDate(object ExpressionVariant)
		{
			System.DateTime result = DateTime.FromOADate(0);
			try
			{
				//
				//    KmaEncodeDate = CDate(ExpressionVariant)
				//    KmaEncodeDate = CDate("1/1/1980")
				//KmaEncodeDate = CDate(0)
				//UPGRADE_WARNING: (2065) Boolean method Information.IsMissing has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2065.aspx
				if (ExpressionVariant != null)
				{
					//UPGRADE_WARNING: (1049) Use of Null/IsNull() detected. More Information: http://www.vbtonet.com/ewis/ewi1049.aspx
					if (!Convert.IsDBNull(ExpressionVariant))
					{
						if (ReflectionHelper.GetPrimitiveValue<string>(ExpressionVariant) != "")
						{
							if (Information.IsDate(ReflectionHelper.GetPrimitiveValue(ExpressionVariant)))
							{
								result = ReflectionHelper.GetPrimitiveValue<System.DateTime>(ExpressionVariant);
							}
						}
					}
				}
				//
			}
			catch
			{
				//
				// ----- ErrorTrap
				//
				result = DateTime.FromOADate(0);
			}
			return result;
		}
		//
		//==========================================================================
		//   Convert a variant to a boolean
		//   Returns true if input is not false, else false
		//==========================================================================
		//
		internal static bool kmaEncodeBoolean(object ExpressionVariant)
		{
			bool result = false;
			try
			{
				//
				//KmaEncodeBoolean = False
				//UPGRADE_WARNING: (2065) Boolean method Information.IsMissing has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2065.aspx
				if (ExpressionVariant != null)
				{
					//UPGRADE_WARNING: (1049) Use of Null/IsNull() detected. More Information: http://www.vbtonet.com/ewis/ewi1049.aspx
					if (!Convert.IsDBNull(ExpressionVariant))
					{
						if (ReflectionHelper.GetPrimitiveValue<string>(ExpressionVariant) != "")
						{
							double dbNumericTemp = 0;
							if (Double.TryParse(ReflectionHelper.GetPrimitiveValue<string>(ExpressionVariant), NumberStyles.Float , CultureInfo.CurrentCulture.NumberFormat, out dbNumericTemp))
							{
								if (ReflectionHelper.GetPrimitiveValue<string>(ExpressionVariant) != "0")
								{
									if (ReflectionHelper.GetPrimitiveValue<double>(ExpressionVariant) != 0)
									{
										result = true;
									}
								}
							}
							else if (ReflectionHelper.GetPrimitiveValue<string>(ExpressionVariant).ToUpper() == "ON")
							{ 
								result = true;
							}
							else if (ReflectionHelper.GetPrimitiveValue<string>(ExpressionVariant).ToUpper() == "YES")
							{ 
								result = true;
							}
							else if (ReflectionHelper.GetPrimitiveValue<string>(ExpressionVariant).ToUpper() == "TRUE")
							{ 
								result = true;
							}
							else
							{
								result = false;
							}
						}
					}
				}
			}
			catch
			{
				//
				// ----- ErrorTrap
				//
				result = false;
			}
			return result;
		}
		//
		//==========================================================================
		//   Convert a variant into 0 or 1
		//   Returns 1 if input is not false, else 0
		//==========================================================================
		//
		internal static int KmaEncodeBit(object ExpressionVariant)
		{
			int result = 0;
			try
			{
				//
				//KmaEncodeBit = 0
				if (kmaEncodeBoolean(ExpressionVariant))
				{
					result = 1;
				}
				//
			}
			catch
			{
				//
				// ----- ErrorTrap
				//
				result = 0;
			}
			return result;
		}
		//
		//==========================================================================
		//   Convert a variant to a string
		//   returns emptystring if the input is not a string
		//==========================================================================
		//
		internal static string kmaEncodeText(object ExpressionVariant)
		{
			string result = "";
			try
			{
				//
				//KmaEncodeText = ""
				//UPGRADE_WARNING: (2065) Boolean method Information.IsMissing has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2065.aspx
				if (ExpressionVariant != null)
				{
					//UPGRADE_WARNING: (1049) Use of Null/IsNull() detected. More Information: http://www.vbtonet.com/ewis/ewi1049.aspx
					if (!Convert.IsDBNull(ExpressionVariant))
					{
						result = ReflectionHelper.GetPrimitiveValue<string>(ExpressionVariant);
					}
				}
				//
			}
			catch
			{
				//
				// ----- ErrorTrap
				//
				result = "";
			}
			return result;
		}
		//
		//==========================================================================
		//   Converts a possibly missing value to variant
		//==========================================================================
		//
		internal static object KmaEncodeMissing(object ExpressionVariant, object DefaultVariant)
		{
			//On Error GoTo ErrorTrap
			//
			//UPGRADE_WARNING: (2065) Boolean method Information.IsMissing has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2065.aspx
			//
			if (ExpressionVariant == null)
			{
				return DefaultVariant;
			}
			else
			{
				return ExpressionVariant;
			}
			//
			// ----- ErrorTrap
			//

			Information.Err().Clear();
			return null;
		}
		//
		//
		//
		internal static string KmaEncodeMissingText(object ExpressionVariant, string DefaultText)
		{
			return kmaEncodeText(KmaEncodeMissing(ExpressionVariant, DefaultText));
		}
		//
		//
		//
		internal static int KmaEncodeMissingInteger(object ExpressionVariant, int DefaultInteger)
		{
			return kmaEncodeInteger(KmaEncodeMissing(ExpressionVariant, DefaultInteger));
		}
		//
		//
		//
		internal static System.DateTime KmaEncodeMissingDate(object ExpressionVariant, System.DateTime DefaultDate)
		{
			return KmaEncodeDate(KmaEncodeMissing(ExpressionVariant, DefaultDate));
		}
		//
		//
		//
		internal static double KmaEncodeMissingNumber(object ExpressionVariant, double DefaultNumber)
		{
			return KmaEncodeNumber(KmaEncodeMissing(ExpressionVariant, DefaultNumber));
		}
		//
		//
		//
		internal static bool KmaEncodeMissingBoolean(object ExpressionVariant, bool DefaultState)
		{
			return kmaEncodeBoolean(KmaEncodeMissing(ExpressionVariant, DefaultState));
		}
		//
		//================================================================================================================
		//   Separate a URL into its host, path, page parts
		//================================================================================================================
		//
		internal static void SeparateURL(string SourceURL, ref string Protocol, ref string Host, ref string Path, ref string Page, ref string QueryString)
		{
			//On Error GoTo ErrorTrap
			//
			//   Divide the URL into URLHost, URLPath, and URLPage
			//
			//
			// Get Protocol (before the first :)
			//
			string WorkingURL = SourceURL;
			int Position = (WorkingURL.IndexOf(':') + 1);
			//Position = InStr(1, WorkingURL, "://")
			if (Position != 0)
			{
				Protocol = WorkingURL.Substring(0, Math.Min(Position + 2, WorkingURL.Length));
				WorkingURL = WorkingURL.Substring(Position + 2);
			}
			//
			// compatibility fix
			//
			if ((WorkingURL.IndexOf("//") + 1) == 1)
			{
				if (Protocol == "")
				{
					Protocol = "http:";
				}
				Protocol = Protocol + "//";
				WorkingURL = WorkingURL.Substring(2);
			}
			//
			// Get QueryString
			//
			Position = (WorkingURL.IndexOf('?') + 1);
			if (Position > 0)
			{
				QueryString = WorkingURL.Substring(Position - 1);
				WorkingURL = WorkingURL.Substring(0, Math.Min(Position - 1, WorkingURL.Length));
			}
			//
			// separate host from pathpage
			//
			//iURLHost = WorkingURL
			Position = (WorkingURL.IndexOf('/') + 1);
			if ((Position == 0) && (Protocol == ""))
			{
				//
				// Page without path or host
				//
				Page = WorkingURL;
				Path = "";
				Host = "";
			}
			else if ((Position == 0))
			{ 
				//
				// host, without path or page
				//
				Page = "";
				Path = "/";
				Host = WorkingURL;
			}
			else
			{
				//
				// host with a path (at least)
				//
				Path = WorkingURL.Substring(Position - 1);
				Host = WorkingURL.Substring(0, Math.Min(Position - 1, WorkingURL.Length));
				//
				// separate page from path
				//
				Position = Path.LastIndexOf("/") + 1;
				if (Position == 0)
				{
					//
					// no path, just a page
					//
					Page = Path;
					Path = "/";
				}
				else
				{
					Page = Path.Substring(Position);
					Path = Path.Substring(0, Math.Min(Position, Path.Length));
				}
			}
			return;
			//
			// ----- ErrorTrap
			//

			Information.Err().Clear();
		}
		//
		//================================================================================================================
		//   Separate a URL into its host, path, page parts
		//================================================================================================================
		//
		internal static void ParseURL(string SourceURL, ref string Protocol, ref string Host, ref string Port, ref string Path, ref string Page, ref string QueryString)
		{
			//On Error GoTo ErrorTrap
			//
			//   Divide the URL into URLHost, URLPath, and URLPage
			//
			string iURLProtocol = "";
			string iURLPort = "";
			string iURLPath = "";
			string iURLPage = "";
			string iURLQueryString = "";
			//
			string iURLWorking = SourceURL; // internal storage for GetURL functions
			int Position = (iURLWorking.IndexOf("://") + 1);
			if (Position != 0)
			{
				iURLProtocol = iURLWorking.Substring(0, Math.Min(Position + 2, iURLWorking.Length));
				iURLWorking = iURLWorking.Substring(Position + 2);
			}
			//
			// separate Host:Port from pathpage
			//
			string iURLHost = iURLWorking;
			Position = (iURLHost.IndexOf('/') + 1);
			if (Position == 0)
			{
				//
				// just host, no path or page
				//
				iURLPath = "/";
				iURLPage = "";
			}
			else
			{
				iURLPath = iURLHost.Substring(Position - 1);
				iURLHost = iURLHost.Substring(0, Math.Min(Position - 1, iURLHost.Length));
				//
				// separate page from path
				//
				Position = iURLPath.LastIndexOf("/") + 1;
				if (Position == 0)
				{
					//
					// no path, just a page
					//
					iURLPage = iURLPath;
					iURLPath = "/";
				}
				else
				{
					iURLPage = iURLPath.Substring(Position);
					iURLPath = iURLPath.Substring(0, Math.Min(Position, iURLPath.Length));
				}
			}
			//
			// Divide Host from Port
			//
			Position = (iURLHost.IndexOf(':') + 1);
			if (Position == 0)
			{
				//
				// host not given, take a guess
				//
				switch(iURLProtocol.ToUpper())
				{
					case "FTP://" : 
						iURLPort = "21"; 
						break;
					case "HTTP://" : case "HTTPS://" : 
						iURLPort = "80"; 
						break;
					default:
						iURLPort = "80"; 
						break;
				}
			}
			else
			{
				iURLPort = iURLHost.Substring(Position);
				iURLHost = iURLHost.Substring(0, Math.Min(Position - 1, iURLHost.Length));
			}
			Position = (iURLPage.IndexOf('?') + 1);
			if (Position > 0)
			{
				iURLQueryString = iURLPage.Substring(Position - 1);
				iURLPage = iURLPage.Substring(0, Math.Min(Position - 1, iURLPage.Length));
			}
			Protocol = iURLProtocol;
			Host = iURLHost;
			Port = iURLPort;
			Path = iURLPath;
			Page = iURLPage;
			QueryString = iURLQueryString;
			return;
			//
			// ----- ErrorTrap
			//

			Information.Err().Clear();
		}
		//
		//
		//
		internal static System.DateTime DecodeGMTDate(string GMTDate)
		{
			//On Error GoTo ErrorTrap
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
			return result;
			//
			// ----- ErrorTrap
			//

		}
		//
		//
		//
		internal static string EncodeGMTDate(System.DateTime MSDate)
		{
			//On Error GoTo ErrorTrap
			//
			return "";
			//
			// ----- ErrorTrap
			//

		}
		//
		//=================================================================================
		//   Renamed to catch all the cases that used it in addons
		//
		//   Do not use this routine in Addons to get the addon option string value
		//   to get the value in an option string, use csv.getAddonOption("name")
		//
		// Get the value of a name in a string of name value pairs parsed with vrlf and =
		//   the legacy line delimiter was a '&' -> name1=value1&name2=value2"
		//   new format is "name1=value1 crlf name2=value2 crlf ..."
		//   There can be no extra spaces between the delimiter, the name and the "="
		//=================================================================================
		//
		internal static string getSimpleNameValue(string Name, ref string ArgumentString, string DefaultValue, ref string Delimiter)
		{
			//Function getArgument(Name As String, ArgumentString As String, Optional DefaultValue As Variant, Optional Delimiter As String) As String
			//
			string result = "";
			int NameLength = 0;
			int ValueStart = 0;
			int ValueEnd = 0;
			bool IsQuoted = false;
			//
			// determine delimiter
			//
			if (Delimiter == "")
			{
				//
				// If not explicit
				//
				if (ArgumentString.IndexOf(Environment.NewLine) >= 0)
				{
					//
					// crlf can only be here if it is the delimiter
					//
					Delimiter = Environment.NewLine;
				}
				else
				{
					//
					// either only one option, or it is the legacy '&' delimit
					//
					Delimiter = "&";
				}
			}
			string iDefaultValue = ReflectionHelper.GetPrimitiveValue<string>(KmaEncodeMissing(DefaultValue, ""));
			string WorkingString = ArgumentString;
			result = iDefaultValue;
			if (WorkingString != "")
			{
				WorkingString = Delimiter + WorkingString + Delimiter;
				ValueStart = (WorkingString.IndexOf(Delimiter + Name + "=", StringComparison.CurrentCultureIgnoreCase) + 1);
				if (ValueStart != 0)
				{
					NameLength = Strings.Len(Name);
					ValueStart = ValueStart + Strings.Len(Delimiter) + NameLength + 1;
					if (WorkingString.Substring(ValueStart - 1, Math.Min(1, WorkingString.Length - (ValueStart - 1))) == "\"")
					{
						IsQuoted = true;
						ValueStart++;
					}
					if (IsQuoted)
					{
						ValueEnd = Strings.InStr(ValueStart, WorkingString, "\"" + Delimiter, CompareMethod.Binary);
					}
					else
					{
						ValueEnd = Strings.InStr(ValueStart, WorkingString, Delimiter, CompareMethod.Binary);
					}
					if (ValueEnd == 0)
					{
						result = WorkingString.Substring(ValueStart - 1);
					}
					else
					{
						result = WorkingString.Substring(ValueStart - 1, Math.Min(ValueEnd - ValueStart, WorkingString.Length - (ValueStart - 1)));
					}
				}
			}
			//

			return result;
			//
			// ----- ErrorTrap
			//

		}
		//
		//=================================================================================
		//   Do not use this code
		//
		//   To retrieve a value from an option string, use csv.getAddonOption("name")
		//
		//   This was left here to work through any code issues that might arrise during
		//   the conversion.
		//
		//   Return the value from a name value pair, parsed with =,&[|].
		//   For example:
		//       name=Jay[Jay|Josh|Dwayne]
		//       the answer is Jay. If a select box is displayed, it is a dropdown of all three
		//=================================================================================
		//
		internal static string GetAggrOption_old(string Name, ref string SegmentCMDArgs)
		{
			//
			//
			string result = "";
			string tempRefParam = Environment.NewLine;
			result = getSimpleNameValue(Name, ref SegmentCMDArgs, "", ref tempRefParam);
			//
			// remove the manual select list syntax "answer[choice1|choice2]"
			//
			int Pos = (result.IndexOf('[') + 1);
			if (Pos != 0)
			{
				result = result.Substring(0, Math.Min(Pos - 1, result.Length));
			}
			//
			// remove any function syntax "answer{selectcontentname RSS Feeds}"
			//
			Pos = (result.IndexOf('{') + 1);
			if (Pos != 0)
			{
				result = result.Substring(0, Math.Min(Pos - 1, result.Length));
			}
			//
			return result;
		}
		//
		//=================================================================================
		//   Do not use this code
		//
		//   To retrieve a value from an option string, use csv.getAddonOption("name")
		//
		//   This was left here to work through any code issues that might arrise during
		//   Compatibility for GetArgument
		//=================================================================================
		//
		internal static string kmaGetNameValue_old(string Name, ref string ArgumentString, string DefaultValue = "")
		{
			string tempRefParam = Environment.NewLine;
			return getSimpleNameValue(Name, ref ArgumentString, DefaultValue, ref tempRefParam);
		}
		//
		//========================================================================
		//   KmaEncodeSQLText
		//========================================================================
		//
		internal static string KmaEncodeSQLText(object ExpressionVariant)
		{
			//On Error GoTo ErrorTrap
			//
			//Dim MethodName As String
			//
			//MethodName = "KmaEncodeSQLText"
			//
			//UPGRADE_WARNING: (2065) Boolean method Information.IsMissing has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2065.aspx
			//UPGRADE_WARNING: (1049) Use of Null/IsNull() detected. More Information: http://www.vbtonet.com/ewis/ewi1049.aspx
			string result = "";
			if (Convert.IsDBNull(ExpressionVariant))
			{
				return "null";
			}
			else if (ExpressionVariant == null)
			{ 
				return "null";
			}
			else if (ReflectionHelper.GetPrimitiveValue<string>(ExpressionVariant) == "")
			{ 
				return "null";
			}
			else
			{
				result = ReflectionHelper.GetPrimitiveValue<string>(ExpressionVariant);
				// ??? this should not be here -- to correct a field used in a CDef, truncate in SaveCS by fieldtype
				//KmaEncodeSQLText = Left(ExpressionVariant, 255)
				//remove-can not find a case where | is not allowed to be saved.
				//KmaEncodeSQLText = Replace(KmaEncodeSQLText, "|", "_")
				return "'" + Strings.Replace(result, "'", "''", 1, -1, CompareMethod.Binary) + "'";
			}
			//
			// ----- Error Trap
			//

			return result;
		}
		//
		//========================================================================
		//   KmaEncodeSQLLongText
		//========================================================================
		//
		internal static string KmaEncodeSQLLongText(object ExpressionVariant)
		{
			//On Error GoTo ErrorTrap
			//
			//Dim MethodName As String
			//
			//MethodName = "KmaEncodeSQLLongText"
			//
			//UPGRADE_WARNING: (2065) Boolean method Information.IsMissing has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2065.aspx
			//UPGRADE_WARNING: (1049) Use of Null/IsNull() detected. More Information: http://www.vbtonet.com/ewis/ewi1049.aspx
			string result = "";
			if (Convert.IsDBNull(ExpressionVariant))
			{
				return "null";
			}
			else if (ExpressionVariant == null)
			{ 
				return "null";
			}
			else if (ReflectionHelper.GetPrimitiveValue<string>(ExpressionVariant) == "")
			{ 
				return "null";
			}
			else
			{
				result = ReflectionHelper.GetPrimitiveValue<string>(ExpressionVariant);
				//KmaEncodeSQLLongText = Replace(ExpressionVariant, "|", "_")
				return "'" + Strings.Replace(result, "'", "''", 1, -1, CompareMethod.Binary) + "'";
			}
			//
			// ----- Error Trap
			//

			return result;
		}
		//
		//========================================================================
		//   KmaEncodeSQLDate
		//       encode a date variable to go in an sql expression
		//========================================================================
		//
		internal static string KmaEncodeSQLDate(object ExpressionVariant)
		{
			//On Error GoTo ErrorTrap
			//
			System.DateTime TimeVar = DateTime.FromOADate(0);
			float TimeValuething = 0;
			int TimeHours = 0;
			int TimeMinutes = 0;
			int TimeSeconds = 0;
			//Dim MethodName As String
			//'
			//MethodName = "KmaEncodeSQLDate"
			//
			//UPGRADE_WARNING: (2065) Boolean method Information.IsMissing has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2065.aspx
			//UPGRADE_WARNING: (1049) Use of Null/IsNull() detected. More Information: http://www.vbtonet.com/ewis/ewi1049.aspx
			if (Convert.IsDBNull(ExpressionVariant))
			{
				return "null";
			}
			else if (ExpressionVariant == null)
			{ 
				return "null";
			}
			else if (ReflectionHelper.GetPrimitiveValue<string>(ExpressionVariant) == "")
			{ 
				return "null";
			}
			else if (Information.IsDate(ExpressionVariant))
			{ 
				TimeVar = ReflectionHelper.GetPrimitiveValue<System.DateTime>(ExpressionVariant);
				if (TimeVar == DateTime.FromOADate(0))
				{
					return "null";
				}
				else
				{
					TimeValuething = (float) (86400f * (TimeVar.ToOADate() - Math.Floor((double) TimeVar.AddDays(0.000011f).ToOADate())));
					TimeHours = Convert.ToInt32((float) Math.Floor((double) (TimeValuething / ((double) 3600f))));
					if (TimeHours >= 24)
					{
						TimeHours = 23;
					}
					TimeMinutes = Convert.ToInt32(((float) Math.Floor((double) (TimeValuething / ((double) 60f)))) - (TimeHours * 60));
					if (TimeMinutes >= 60)
					{
						TimeMinutes = 59;
					}
					TimeSeconds = Convert.ToInt32(TimeValuething - (TimeHours * 3600f) - (TimeMinutes * 60f));
					if (TimeSeconds >= 60)
					{
						TimeSeconds = 59;
					}
					return "{ts '" + ReflectionHelper.GetPrimitiveValue<System.DateTime>(ExpressionVariant).Year.ToString() + "-" + ("0" + ReflectionHelper.GetPrimitiveValue<System.DateTime>(ExpressionVariant).Month.ToString()).Substring(Math.Max(("0" + ReflectionHelper.GetPrimitiveValue<System.DateTime>(ExpressionVariant).Month.ToString()).Length - 2, 0)) + "-" + ("0" + DateAndTime.Day(ReflectionHelper.GetPrimitiveValue<System.DateTime>(ExpressionVariant)).ToString()).Substring(Math.Max(("0" + DateAndTime.Day(ReflectionHelper.GetPrimitiveValue<System.DateTime>(ExpressionVariant)).ToString()).Length - 2, 0)) + " " + ("0" + TimeHours.ToString()).Substring(Math.Max(("0" + TimeHours.ToString()).Length - 2, 0)) + ":" + ("0" + TimeMinutes.ToString()).Substring(Math.Max(("0" + TimeMinutes.ToString()).Length - 2, 0)) + ":" + ("0" + TimeSeconds.ToString()).Substring(Math.Max(("0" + TimeSeconds.ToString()).Length - 2, 0)) + "'}";
				}
			}
			else
			{
				return "null";
			}
			//
			// ----- Error Trap
			//

		}
		//
		//========================================================================
		//   KmaEncodeSQLNumber
		//       encode a number variable to go in an sql expression
		//========================================================================
		//
		internal static string KmaEncodeSQLNumber(object ExpressionVariant)
		{
			//On Error GoTo ErrorTrap
			//
			//Dim MethodName As String
			//'
			//MethodName = "KmaEncodeSQLNumber"
			//
			double dbNumericTemp = 0;
			//UPGRADE_WARNING: (2065) Boolean method Information.IsMissing has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2065.aspx
			//UPGRADE_WARNING: (1049) Use of Null/IsNull() detected. More Information: http://www.vbtonet.com/ewis/ewi1049.aspx
			if (Convert.IsDBNull(ExpressionVariant))
			{
				return "null";
			}
			else if (ExpressionVariant == null)
			{ 
				return "null";
			}
			else if (ReflectionHelper.GetPrimitiveValue<string>(ExpressionVariant) == "")
			{ 
				return "null";
			}
			else if (Double.TryParse(ReflectionHelper.GetPrimitiveValue<string>(ExpressionVariant), NumberStyles.Float , CultureInfo.CurrentCulture.NumberFormat, out dbNumericTemp))
			{ 
				VariantType switchVar = Information.VarType(ReflectionHelper.GetPrimitiveValue(ExpressionVariant));
				if (switchVar == VariantType.Boolean)
				{
					if (ReflectionHelper.GetPrimitiveValue<bool>(ExpressionVariant))
					{
						return SQLTrue;
					}
					else
					{
						return SQLFalse;
					}
				}
				else
				{
					return ReflectionHelper.GetPrimitiveValue<string>(ExpressionVariant);
				}
			}
			else
			{
				return "null";
			}
			//
			// ----- Error Trap
			//

		}
		//
		//========================================================================
		//   KmaEncodeSQLBoolean
		//       encode a boolean variable to go in an sql expression
		//========================================================================
		//
		internal static string KmaEncodeSQLBoolean(object ExpressionVariant)
		{
			//
			string result = "";
			//
			result = SQLFalse;
			if (kmaEncodeBoolean(ExpressionVariant))
			{
				result = SQLTrue;
			}
			//    If Not IsNull(ExpressionVariant) Then
			//        If Not IsMissing(ExpressionVariant) Then
			//            If ExpressionVariant <> False Then
			//                    KmaEncodeSQLBoolean = SQLTrue
			//                End If
			//            End If
			//        End If
			//    End If
			//
			return result;
		}
		//
		//========================================================================
		//   Gets the next line from a string, and removes the line
		//========================================================================
		//
		internal static string KmaGetLine(ref string Body)
		{
			string result = "";
			string EOL = "";
			int BOL = 0;
			//
			int NextCR = (Body.IndexOf("\r") + 1);
			int NextLF = (Body.IndexOf(Constants.vbLf) + 1);

			if (NextCR != 0 || NextLF != 0)
			{
				if (NextCR != 0)
				{
					if (NextLF != 0)
					{
						if (NextCR < NextLF)
						{
							EOL = (NextCR - 1).ToString();
							if (NextLF == NextCR + 1)
							{
								BOL = NextLF + 1;
							}
							else
							{
								BOL = NextCR + 1;
							}

						}
						else
						{
							EOL = (NextLF - 1).ToString();
							BOL = NextLF + 1;
						}
					}
					else
					{
						EOL = (NextCR - 1).ToString();
						BOL = NextCR + 1;
					}
				}
				else
				{
					EOL = (NextLF - 1).ToString();
					BOL = NextLF + 1;
				}
				result = Body.Substring(0, Math.Min(Convert.ToInt32(Double.Parse(EOL)), Body.Length));
				Body = Body.Substring(BOL - 1);
			}
			else
			{
				result = Body;
				Body = "";
			}

			//EOL = InStr(1, Body, vbCrLf)

			//If EOL <> 0 Then
			//    KmaGetLine = Mid(Body, 1, EOL - 1)
			//    Body = Mid(Body, EOL + 2)
			//    End If
			//
			return result;
		}
		//
		//=================================================================================
		//   Get a Random Long Value
		//=================================================================================
		//
		internal static int GetRandomInteger()
		{
			//
			//
			int RandomBase = Thread.CurrentThread.ManagedThreadId;
			RandomBase = RandomBase & (((int) Math.Pow(2, 30)) - 1);
			int RandomLimit = ((int) Math.Pow(2, 31)) - RandomBase - 1;
			VBMath.Randomize();
			return Convert.ToInt32(RandomBase + (VBMath.Rnd() * RandomLimit));
			//
		}
		//
		//=================================================================================
		//
		//=================================================================================
		//
		internal static bool IsRSOK(object RS)
		{
			bool result = false;
			if (RS != null)
			{
				if (ReflectionHelper.GetMember<double>(RS, "State") != 0)
				{
					if (~ReflectionHelper.GetMember<int>(RS, "EOF") != 0)
					{
						result = true;
					}
				}
			}
			return result;
		}
		//
		//=================================================================================
		//
		//=================================================================================
		//
		internal static void CloseRS(object RS)
		{
			if (RS != null)
			{
				if (ReflectionHelper.GetMember<double>(RS, "State") != 0)
				{
					ReflectionHelper.Invoke(RS, "Close", new object[]{});
				}
			}
		}
		//
		//=============================================================================
		// Create the part of the sql where clause that is modified by the user
		//   WorkingQuery is the original querystring to change
		//   QueryName is the name part of the name pair to change
		//   If the QueryName is not found in the string, ignore call
		//=============================================================================
		//
		internal static string ModifyQueryString(ref string WorkingQuery, string QueryName, string QueryValue, object AddIfMissing = null)
		{
			//
			if (WorkingQuery.IndexOf('?') >= 0)
			{
				return kmaModifyLinkQuery(ref WorkingQuery, QueryName, QueryValue, AddIfMissing);
			}
			else
			{
				string tempRefParam = "?" + WorkingQuery;
				return kmaModifyLinkQuery(ref tempRefParam, QueryName, QueryValue, AddIfMissing).Substring(1);
			}
		}
		//
		//=============================================================================
		//   Modify a querystring name/value pair in a Link
		//=============================================================================
		//
		internal static string kmaModifyLinkQuery(ref string Link, string QueryName, string QueryValue, object AddIfMissing = null)
		{
			//
			string result = "";
			string[] Element = null;
			int ElementCount = 0;
			string[] NameValue = null;
			bool ElementFound = false;
			string QueryString = "";
			//
			bool iAddIfMissing = KmaEncodeMissingBoolean(AddIfMissing, true);
			if (Link.IndexOf('?') >= 0)
			{
				result = Link.Substring(0, Math.Min(Link.IndexOf('?'), Link.Length));
				QueryString = Link.Substring(Strings.Len(result) + 1);
			}
			else
			{
				result = Link;
				QueryString = "";
			}
			string UcaseQueryName = kmaEncodeRequestVariable(QueryName).ToUpper();
			if (QueryString != "")
			{
				Element = (string[]) QueryString.Split('&');
				ElementCount = Element.GetUpperBound(0) + 1;
				for (int ElementPointer = 0; ElementPointer <= ElementCount - 1; ElementPointer++)
				{
					NameValue = (string[]) Element[ElementPointer].Split('=');
					if (NameValue.GetUpperBound(0) == 1)
					{
						if (NameValue[0].ToUpper() == UcaseQueryName)
						{
							if (QueryValue == "")
							{
								Element[ElementPointer] = "";
							}
							else
							{
								Element[ElementPointer] = QueryName + "=" + QueryValue;
							}
							ElementFound = true;
							break;
						}
					}
				}
			}
			if (!ElementFound && (QueryValue != ""))
			{
				//
				// element not found, it needs to be added
				//
				if (iAddIfMissing)
				{
					if (QueryString == "")
					{
						QueryString = kmaEncodeRequestVariable(QueryName) + "=" + kmaEncodeRequestVariable(QueryValue);
					}
					else
					{
						QueryString = QueryString + "&" + kmaEncodeRequestVariable(QueryName) + "=" + kmaEncodeRequestVariable(QueryValue);
					}
				}
			}
			else
			{
				//
				// element found
				//
				QueryString = String.Join("&", Element);
				if ((QueryString != "") && (QueryValue == ""))
				{
					//
					// element found and needs to be removed
					//
					QueryString = Strings.Replace(QueryString, "&&", "&", 1, -1, CompareMethod.Binary);
					if (QueryString.StartsWith("&"))
					{
						QueryString = QueryString.Substring(1);
					}
					if (QueryString.Substring(Math.Max(QueryString.Length - 1, 0)) == "&")
					{
						QueryString = QueryString.Substring(0, Math.Min(Strings.Len(QueryString) - 1, QueryString.Length));
					}
				}
			}
			if (QueryString != "")
			{
				result = result + "?" + QueryString;
			}
			return result;
		}
		//
		//=================================================================================
		//
		//=================================================================================
		//
		internal static string GetIntegerString(int Value, int DigitCount)
		{
			//UPGRADE_WARNING: (2081) Len has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2081.aspx
			if (Marshal.SizeOf(Value) <= DigitCount)
			{
				return new string('0', DigitCount - Strings.Len(Value.ToString())) + Value.ToString();
			}
			else
			{
				return Value.ToString();
			}
		}
		//
		//==========================================================================================
		//   Set the current process to a high priority
		//       Should be called once from the objects parent when it is first created.
		//
		//   taken from an example labeled
		//       KPD-Team 2000
		//       URL: http://www.allapi.net/
		//       Email: KPDTeam@Allapi.net
		//==========================================================================================
		//
		internal static void SetProcessHighPriority()
		{
			//
			//set the new priority class
			//
			int hProcess = UpgradeSolution1Support.PInvoke.SafeNative.kernel32.GetCurrentProcess();
			UpgradeSolution1Support.PInvoke.SafeNative.kernel32.SetPriorityClass(hProcess, HIGH_PRIORITY_CLASS);
			//
		}
		//
		//==========================================================================================
		//   Format the current error object into a standard string
		//==========================================================================================
		//
		internal static string GetErrString(ErrObject ErrorObject = null)
		{
			string Copy = "";
			if (ErrorObject == null)
			{
				//UPGRADE_WARNING: (2081) Err.Number has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2081.aspx
				if (Information.Err().Number == 0)
				{
					return "[no error]";
				}
				else
				{
					Copy = Information.Err().Description;
					Copy = Strings.Replace(Copy, Environment.NewLine, "-", 1, -1, CompareMethod.Binary);
					Copy = Strings.Replace(Copy, Constants.vbLf, "-", 1, -1, CompareMethod.Binary);
					Copy = Strings.Replace(Copy, Environment.NewLine, "", 1, -1, CompareMethod.Binary);
					//UPGRADE_WARNING: (2081) Err.Number has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2081.aspx
					return "[" + Information.Err().Source + " #" + Information.Err().Number.ToString() + ", " + Copy + "]";
				}
			}
			else
			{
				//UPGRADE_WARNING: (2081) Err.Number has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2081.aspx
				if (ErrorObject.Number == 0)
				{
					return "[no error]";
				}
				else
				{
					Copy = ErrorObject.Description;
					Copy = Strings.Replace(Copy, Environment.NewLine, "-", 1, -1, CompareMethod.Binary);
					Copy = Strings.Replace(Copy, Constants.vbLf, "-", 1, -1, CompareMethod.Binary);
					Copy = Strings.Replace(Copy, Environment.NewLine, "", 1, -1, CompareMethod.Binary);
					//UPGRADE_WARNING: (2081) Err.Number has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2081.aspx
					return "[" + ErrorObject.Source + " #" + ErrorObject.Number.ToString() + ", " + Copy + "]";
				}
			}
			//
		}
		//
		//==========================================================================================
		//   Format the current error object into a standard string
		//==========================================================================================
		//
		internal static int GetProcessID()
		{
			return UpgradeSolution1Support.PInvoke.SafeNative.kernel32.GetCurrentProcessId();
		}
		//
		//==========================================================================================
		//   Test if a test string is in a delimited string
		//==========================================================================================
		//
		internal static bool IsInDelimitedString(string DelimitedString, string TestString, string Delimiter)
		{
			return 0 != ((Delimiter + DelimitedString + Delimiter).IndexOf(Delimiter + TestString + Delimiter, StringComparison.CurrentCultureIgnoreCase) + 1);
		}
		//
		//========================================================================
		// kmaEncodeURL
		//
		//   Encodes only what is to the left of the first ?
		//   All URL path characters are assumed to be correct (/:#)
		//========================================================================
		//
		internal static string kmaEncodeURL(string Source)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			string result = "";
			string[] URLSplit = null;
			//
			result = Source;
			if (Source != "")
			{
				URLSplit = (string[]) Source.Split('?');
				result = URLSplit[0];
				result = Strings.Replace(result, "%", "%25", 1, -1, CompareMethod.Binary);
				//
				result = Strings.Replace(result, "\"", "%22", 1, -1, CompareMethod.Binary);
				result = Strings.Replace(result, " ", "%20", 1, -1, CompareMethod.Binary);
				result = Strings.Replace(result, "$", "%24", 1, -1, CompareMethod.Binary);
				result = Strings.Replace(result, "+", "%2B", 1, -1, CompareMethod.Binary);
				result = Strings.Replace(result, ",", "%2C", 1, -1, CompareMethod.Binary);
				result = Strings.Replace(result, ";", "%3B", 1, -1, CompareMethod.Binary);
				result = Strings.Replace(result, "<", "%3C", 1, -1, CompareMethod.Binary);
				result = Strings.Replace(result, "=", "%3D", 1, -1, CompareMethod.Binary);
				result = Strings.Replace(result, ">", "%3E", 1, -1, CompareMethod.Binary);
				result = Strings.Replace(result, "@", "%40", 1, -1, CompareMethod.Binary);
				if (URLSplit.GetUpperBound(0) > 0)
				{
					result = result + "?" + kmaEncodeQueryString(URLSplit[1]);
				}
			}
			//
			return result;
		}
		//
		//========================================================================
		// kmaEncodeQueryString
		//
		//   This routine encodes the URL QueryString to conform to rules
		//========================================================================
		//
		internal static string kmaEncodeQueryString(string Source)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			string result = "";
			string[] QSSplit = null;
			string[] NVSplit = null;
			string NV = "";
			//
			if (Source != "")
			{
				QSSplit = (string[]) Source.Split('&');
				foreach (string QSSplit_item in QSSplit)
				{
					NV = QSSplit_item;
					if (NV != "")
					{
						NVSplit = (string[]) NV.Split('=');
						if (NVSplit.GetUpperBound(0) == 0)
						{
							NVSplit[0] = kmaEncodeRequestVariable(NVSplit[0]);
							result = result + "&" + NVSplit[0];
						}
						else
						{
							NVSplit[0] = kmaEncodeRequestVariable(NVSplit[0]);
							NVSplit[1] = kmaEncodeRequestVariable(NVSplit[1]);
							result = result + "&" + NVSplit[0] + "=" + NVSplit[1];
						}
					}
				}
				if (result != "")
				{
					result = result.Substring(1);
				}
			}
			//
			return result;
		}
		//
		//========================================================================
		// kmaEncodeRequestVariable
		//
		//   This routine encodes a request variable for a URL Query String
		//       ...can be the requestname or the requestvalue
		//========================================================================
		//
		internal static string kmaEncodeRequestVariable(string Source)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			string result = "";
			string Character = "";
			string LocalSource = "";
			//
			if (Source != "")
			{
				LocalSource = Source;
				// "+" is an allowed character for filenames. If you add it, the wrong file will be looked up
				//LocalSource = Replace(LocalSource, " ", "+")
				for (int SourcePointer = 1; SourcePointer <= Strings.Len(LocalSource); SourcePointer++)
				{
					Character = LocalSource.Substring(SourcePointer - 1, Math.Min(1, LocalSource.Length - (SourcePointer - 1)));
					// "%" added so if this is called twice, it will not destroy "%20" values
					//If Character = " " Then
					//    kmaEncodeRequestVariable = kmaEncodeRequestVariable & "+"
					if ((("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.-_!*()").IndexOf(Character, StringComparison.CurrentCultureIgnoreCase) + 1) != 0)
					{
						//ElseIf InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789./:-_!*()", Character, vbTextCompare) <> 0 Then
						//ElseIf InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789./:?#-_!~*'()%", Character, vbTextCompare) <> 0 Then
						result = result + Character;
					}
					else
					{
						result = result + "%" + Strings.Asc(Character[0]).ToString("X");
					}
				}
			}
			//
			return result;
		}
		//
		//========================================================================
		// kmaEncodeHTML
		//
		//   Convert all characters that are not allowed in HTML to their Text equivalent
		//   in preperation for use on an HTML page
		//========================================================================
		//
		internal static string kmaEncodeHTML(string Source)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			string result = "";
			result = Source;
			result = Strings.Replace(result, "&", "&amp;", 1, -1, CompareMethod.Binary);
			result = Strings.Replace(result, "<", "&lt;", 1, -1, CompareMethod.Binary);
			result = Strings.Replace(result, ">", "&gt;", 1, -1, CompareMethod.Binary);
			result = Strings.Replace(result, "\"", "&quot;", 1, -1, CompareMethod.Binary);
			return Strings.Replace(result, "'", "&#39;", 1, -1, CompareMethod.Binary);
			//kmaEncodeHTML = Replace(kmaEncodeHTML, "'", "&apos;")
			//
		}
		//
		//========================================================================
		// kmaDecodeHTML
		//
		//   Convert HTML equivalent characters to their equivalents
		//========================================================================
		//
		internal static string kmaDecodeHTML(string Source)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			string result = "";
			int Pos = 0;
			string s = "";
			string CharCodeString = "";
			int CharCode = 0;
			int posEnd = 0;
			//
			// 11/26/2009 - basically re-wrote it, I commented the old one out below
			//
			if (Source != "")
			{
				s = Source;
				//
				Pos = Strings.Len(s);
				Pos = s.LastIndexOf("&#", Pos - 1) + 1;

				while(Pos != 0)
				{
					CharCodeString = "";
					if (s.Substring(Pos + 2, Math.Min(1, s.Length - (Pos + 2))) == ";")
					{
						CharCodeString = s.Substring(Pos + 1, Math.Min(1, s.Length - (Pos + 1)));
						posEnd = Pos + 4;
					}
					else if (s.Substring(Pos + 3, Math.Min(1, s.Length - (Pos + 3))) == ";")
					{ 
						CharCodeString = s.Substring(Pos + 1, Math.Min(2, s.Length - (Pos + 1)));
						posEnd = Pos + 5;
					}
					else if (s.Substring(Pos + 4, Math.Min(1, s.Length - (Pos + 4))) == ";")
					{ 
						CharCodeString = s.Substring(Pos + 1, Math.Min(3, s.Length - (Pos + 1)));
						posEnd = Pos + 6;
					}
					if (CharCodeString != "")
					{
						double dbNumericTemp = 0;
						if (Double.TryParse(CharCodeString, NumberStyles.Float , CultureInfo.CurrentCulture.NumberFormat, out dbNumericTemp))
						{
							CharCode = Convert.ToInt32(Double.Parse(CharCodeString));
							s = s.Substring(0, Math.Min(Pos - 1, s.Length)) + Strings.Chr(CharCode).ToString() + s.Substring(posEnd - 1);
						}
					}
					//
					Pos = s.LastIndexOf("&#", Pos - 1) + 1;
				};
				//
				// replace out all common names (at least the most common for now)
				//
				s = Strings.Replace(s, "&lt;", "<", 1, -1, CompareMethod.Binary);
				s = Strings.Replace(s, "&gt;", ">", 1, -1, CompareMethod.Binary);
				s = Strings.Replace(s, "&quot;", "\"", 1, -1, CompareMethod.Binary);
				s = Strings.Replace(s, "&apos;", "'", 1, -1, CompareMethod.Binary);
				//
				// Always replace the amp last
				//
				s = Strings.Replace(s, "&amp;", "&", 1, -1, CompareMethod.Binary);
				//
				result = s;
			}
			// pre-11/26/2009
			//kmaDecodeHTML = Source
			//kmaDecodeHTML = Replace(kmaDecodeHTML, "&amp;", "&")
			//kmaDecodeHTML = Replace(kmaDecodeHTML, "&lt;", "<")
			//kmaDecodeHTML = Replace(kmaDecodeHTML, "&gt;", ">")
			//kmaDecodeHTML = Replace(kmaDecodeHTML, "&quot;", """")
			//kmaDecodeHTML = Replace(kmaDecodeHTML, "&nbsp;", " ")
			//
			return result;
		}
		//
		//========================================================================
		// kmaAddSpanClass
		//
		//   Adds a span around the copy with the class name provided
		//========================================================================
		//
		internal static string kmaAddSpan(string Copy, string ClassName)
		{
			//
			return "<SPAN Class=\"" + ClassName + "\">" + Copy + "</SPAN>";
			//
		}
		//
		//========================================================================
		// kmaDecodeResponseVariable
		//
		//   Converts a querystring name or value back into the characters it represents
		//   This is the same code as the decodeurl
		//========================================================================
		//
		internal static string kmaDecodeResponseVariable(string Source)
		{
			//
			string result = "";
			string ESCString = "";
			int ESCValue = 0;
			string Digit0 = "";
			string Digit1 = "";
			//
			string iURL = kmaEncodeText(Source);
			result = Strings.Replace(iURL, "+", " ", 1, -1, CompareMethod.Binary);
			int Position = (result.IndexOf('%') + 1);

			while(Position != 0)
			{
				ESCString = result.Substring(Position - 1, Math.Min(3, result.Length - (Position - 1)));
				Digit0 = ESCString.Substring(1, Math.Min(1, ESCString.Length - 1)).ToUpper();
				Digit1 = ESCString.Substring(2, Math.Min(1, ESCString.Length - 2)).ToUpper();
				if (((String.CompareOrdinal(Digit0, "0") >= 0) && (String.CompareOrdinal(Digit0, "9") <= 0)) || ((String.CompareOrdinal(Digit0, "A") >= 0) && (String.CompareOrdinal(Digit0, "F") <= 0)))
				{
					if (((String.CompareOrdinal(Digit1, "0") >= 0) && (String.CompareOrdinal(Digit1, "9") <= 0)) || ((String.CompareOrdinal(Digit1, "A") >= 0) && (String.CompareOrdinal(Digit1, "F") <= 0)))
					{
						ESCValue = Convert.ToInt32(Double.Parse("&H" + ESCString.Substring(1)));
						result = result.Substring(0, Math.Min(Position - 1, result.Length)) + Strings.Chr(ESCValue).ToString() + result.Substring(Position + 2);
						//  & Replace(kmaDecodeResponseVariable, ESCString, Chr(ESCValue), Position, 1)
					}
				}
				Position = Strings.InStr(Position + 1, result, "%", CompareMethod.Binary);
			};
			//
			return result;
		}
		//
		//========================================================================
		// kmadecodeURL
		//   Converts a querystring from an Encoded URL (with %20 and +), to non incoded (with spaced)
		//========================================================================
		//
		internal static string kmaDecodeURL(string Source)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			string result = "";
			string ESCString = "";
			int ESCValue = 0;
			string Digit0 = "";
			string Digit1 = "";
			//
			string iURL = kmaEncodeText(Source);
			result = Strings.Replace(iURL, "+", " ", 1, -1, CompareMethod.Binary);
			int Position = (result.IndexOf('%') + 1);

			while(Position != 0)
			{
				ESCString = result.Substring(Position - 1, Math.Min(3, result.Length - (Position - 1)));
				Digit0 = ESCString.Substring(1, Math.Min(1, ESCString.Length - 1)).ToUpper();
				Digit1 = ESCString.Substring(2, Math.Min(1, ESCString.Length - 2)).ToUpper();
				if (((String.CompareOrdinal(Digit0, "0") >= 0) && (String.CompareOrdinal(Digit0, "9") <= 0)) || ((String.CompareOrdinal(Digit0, "A") >= 0) && (String.CompareOrdinal(Digit0, "F") <= 0)))
				{
					if (((String.CompareOrdinal(Digit1, "0") >= 0) && (String.CompareOrdinal(Digit1, "9") <= 0)) || ((String.CompareOrdinal(Digit1, "A") >= 0) && (String.CompareOrdinal(Digit1, "F") <= 0)))
					{
						ESCValue = Convert.ToInt32(Double.Parse("&H" + ESCString.Substring(1)));
						result = Strings.Replace(result, ESCString, Strings.Chr(ESCValue).ToString(), 1, -1, CompareMethod.Binary);
					}
				}
				Position = Strings.InStr(Position + 1, result, "%", CompareMethod.Binary);
			};
			//
			return result;
		}
		//
		//========================================================================
		// kmaGetFirstNonZeroDate
		//
		//   Converts a querystring name or value back into the characters it represents
		//========================================================================
		//
		internal static System.DateTime kmaGetFirstNonZeroDate(System.DateTime Date0, System.DateTime Date1)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			//
			System.DateTime NullDate = DateTime.FromOADate(0);
			if (Date0 == NullDate)
			{
				if (Date1 == NullDate)
				{
					//
					// Both 0, return 0
					//
					return NullDate;
				}
				else
				{
					//
					// Date0 is NullDate, return Date1
					//
					return Date1;
				}
			}
			else
			{
				if (Date1 == NullDate)
				{
					//
					// Date1 is nulldate, return Date0
					//
					return Date0;
				}
				else if (Date0 < Date1)
				{ 
					//
					// Date0 is first
					//
					return Date0;
				}
				else
				{
					//
					// Date1 is first
					//
					return Date1;
				}
			}
			//
		}
		//
		//========================================================================
		// kmaGetFirstposition
		//
		//   returns 0 if both are zero
		//   returns 1 if the first integer is non-zero and less then the second
		//   returns 2 if the second integer is non-zero and less then the first
		//========================================================================
		//
		internal static int kmaGetFirstNonZeroLong(int Integer1, int Integer2)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			if (Integer1 == 0)
			{
				if (Integer2 == 0)
				{
					//
					// Both 0, return 0
					//
					return 0;
				}
				else
				{
					//
					// Integer1 is 0, return Integer2
					//
					return 2;
				}
			}
			else
			{
				if (Integer2 == 0)
				{
					//
					// Integer2 is 0, return Integer1
					//
					return 1;
				}
				else if (Integer1 < Integer2)
				{ 
					//
					// Integer1 is first
					//
					return 1;
				}
				else
				{
					//
					// Integer2 is first
					//
					return 2;
				}
			}
			//
		}
		//
		//========================================================================
		// kmaSplit
		//   returns the result of a Split, except it honors quoted text
		//   if a quote is found, it is assumed to also be a delimiter ( 'this"that"theother' = 'this "that" theother' )
		//========================================================================
		//
		internal static string[] kmaSplit(string WordList, string Delimiter)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			string[] QuoteSplit = null;
			int QuoteSplitCount = 0;
			bool InQuote = false;
			string[] Out = null;
			string[] SpaceSplit = null;
			int SpaceSplitCount = 0;
			string Fragment = "";
			//
			int OutPointer = 0;
			Out = new string[]{""};
			int OutSize = 1;
			if (WordList != "")
			{
				QuoteSplit = (string[]) WordList.Split(new string[]{"\""}, StringSplitOptions.None);
				QuoteSplitCount = QuoteSplit.GetUpperBound(0) + 1;
				InQuote = (WordList.StartsWith(""));
				for (int QuoteSplitPointer = 0; QuoteSplitPointer <= QuoteSplitCount - 1; QuoteSplitPointer++)
				{
					Fragment = QuoteSplit[QuoteSplitPointer];
					if (Fragment == "")
					{
						//
						// empty fragment
						// this is a quote at the end, or two quotes together
						// do not skip to the next out pointer
						//
						if (OutPointer >= OutSize)
						{
							OutSize += 10;
							Out = ArraysHelper.RedimPreserve(Out, new int[]{OutSize + 1});
						}
						//OutPointer = OutPointer + 1
					}
					else
					{
						if (!InQuote)
						{
							SpaceSplit = (string[]) Fragment.Split(new string[]{Delimiter}, StringSplitOptions.None);
							SpaceSplitCount = SpaceSplit.GetUpperBound(0) + 1;
							for (int SpaceSplitPointer = 0; SpaceSplitPointer <= SpaceSplitCount - 1; SpaceSplitPointer++)
							{
								if (OutPointer >= OutSize)
								{
									OutSize += 10;
									Out = ArraysHelper.RedimPreserve(Out, new int[]{OutSize + 1});
								}
								Out[OutPointer] = Out[OutPointer] + SpaceSplit[SpaceSplitPointer];
								if (SpaceSplitPointer != (SpaceSplitCount - 1))
								{
									//
									// divide output between splits
									//
									OutPointer++;
									if (OutPointer >= OutSize)
									{
										OutSize += 10;
										Out = ArraysHelper.RedimPreserve(Out, new int[]{OutSize + 1});
									}
								}
							}
						}
						else
						{
							Out[OutPointer] = Out[OutPointer] + "\"" + Fragment + "\"";
						}
					}
					InQuote = !InQuote;
				}
			}
			Out = ArraysHelper.RedimPreserve(Out, new int[]{OutPointer + 1});
			//
			//
			return Out;
			//
		}
		//
		//========================================================================
		// kmaSplit_Old
		//   returns the result of a Split, except it honors quoted text
		//   if a quote is found, it is assumed to also be a delimiter ( 'this"that"theother' = 'this "that" theother' )
		//========================================================================
		//
		internal static string[] kmaSplit_Old(string WordList, ref string Delimiter)
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			int CurrentPosition = 0;
			int NextDelimiterPosition = 0;
			int NextQuotePosition = 0;
			int ResultCount = 0;
			string[] Result = null;
			int LenWorkingWordList = 0;
			//
			string WorkingWordList = WordList.Trim();
			if (WorkingWordList != "")
			{
				LenWorkingWordList = Strings.Len(WorkingWordList);
				CurrentPosition = 1;
				ResultCount = 0;

				while(CurrentPosition != 0)
				{
					NextDelimiterPosition = Strings.InStr(CurrentPosition, WorkingWordList, Delimiter, CompareMethod.Text);
					NextQuotePosition = Strings.InStr(CurrentPosition, WorkingWordList, "\"", CompareMethod.Text);
					switch(kmaGetFirstNonZeroLong(NextDelimiterPosition, NextQuotePosition))
					{
						case 0 : 
							// 
							// no more left 
							// 
							if (CurrentPosition <= LenWorkingWordList)
							{
								Result = ArraysHelper.RedimPreserve(Result, new int[]{ResultCount + 1});
								Result[ResultCount] = WorkingWordList.Substring(CurrentPosition - 1);
								ResultCount++;
							} 
							CurrentPosition = 0; 
							break;
						case 1 : 
							// 
							// Delimiter found before quote 
							// 
							Result = ArraysHelper.RedimPreserve(Result, new int[]{ResultCount + 1}); 
							Result[ResultCount] = WorkingWordList.Substring(CurrentPosition - 1, Math.Min(NextDelimiterPosition - CurrentPosition, WorkingWordList.Length - (CurrentPosition - 1))); 
							ResultCount++; 
							if (NextDelimiterPosition >= LenWorkingWordList)
							{
								CurrentPosition = 0;
							}
							else
							{
								CurrentPosition = NextDelimiterPosition + 1;
							} 
							break;
						case 2 : 
							// 
							// Quote Found before delimiter 
							// 
							CurrentPosition = NextQuotePosition + 1; 
							NextQuotePosition = Strings.InStr(CurrentPosition, WorkingWordList, "\"", CompareMethod.Text); 
							if (NextQuotePosition == 0)
							{
								//
								// Problem, as single quote. Just end the phrase here
								//
								NextQuotePosition = LenWorkingWordList + 1;
							} 
							Result = ArraysHelper.RedimPreserve(Result, new int[]{ResultCount + 1}); 
							Result[ResultCount] = WorkingWordList.Substring(CurrentPosition - 1, Math.Min(NextQuotePosition - CurrentPosition, WorkingWordList.Length - (CurrentPosition - 1))); 
							ResultCount++; 
							if (NextQuotePosition >= LenWorkingWordList)
							{
								CurrentPosition = 0;
							}
							else
							{
								CurrentPosition = NextQuotePosition + 1;
							} 
							break;
					}
					//
					// pass any delimiters
					//
					if (CurrentPosition != 0)
					{

						while(WorkingWordList.Substring(CurrentPosition - 1, Math.Min(1, WorkingWordList.Length - (CurrentPosition - 1))) == Delimiter)
						{
							CurrentPosition++;
							if (CurrentPosition >= LenWorkingWordList)
							{
								CurrentPosition = 0;
								break;
							}
						};
					}
				};
			}
			return Result;
			//

			//
			//kmaSplit_Old = Split(WorkingWordList, Delimiter, , vbTextCompare)
			//
		}
		//
		//
		//
		internal static string kmaGetYesNo(bool Key)
		{
			if (Key)
			{
				return "Yes";
			}
			else
			{
				return "No";
			}
		}
		//
		//
		//
		internal static string kmaGetFilename(string PathFilename)
		{
			//
			string result = "";
			result = PathFilename;
			int Position = result.LastIndexOf("/") + 1;
			if (Position != 0)
			{
				result = result.Substring(Position);
			}
			return result;
		}
		//
		//
		//
		internal static string kmaStartTable(int Padding, int Spacing, int Border, string ClassStyle = "")
		{
			return "<table border=\"" + Border.ToString() + "\" cellpadding=\"" + Padding.ToString() + "\" cellspacing=\"" + Spacing.ToString() + "\" class=\"" + ClassStyle + "\" width=\"100%\">";
		}
		//
		//
		//
		internal static string kmaStartTableRow()
		{
			return "<tr>";
		}
		//
		//
		//
		internal static string kmaStartTableCell(string Width = "", int ColSpan = 0, bool EvenRow = false, string Align = "", string BGColor = "")
		{
			string result = "";
			if (Width != "")
			{
				result = " width=\"" + Width + "\"";
			}
			if (BGColor != "")
			{
				result = result + " bgcolor=\"" + BGColor + "\"";
			}
			else if (EvenRow)
			{ 
				result = result + " class=\"ccPanelRowEven\"";
			}
			else
			{
				result = result + " class=\"ccPanelRowOdd\"";
			}
			if (ColSpan != 0)
			{
				result = result + " colspan=\"" + ColSpan.ToString() + "\"";
			}
			if (Align != "")
			{
				result = result + " align=\"" + Align + "\"";
			}
			return "<TD" + result + ">";
		}
		//
		//
		//
		internal static string kmaGetTableCell(string Copy, string Width = "", int ColSpan = 0, bool EvenRow = false, string Align = "", string BGColor = "")
		{
			return kmaStartTableCell(Width, ColSpan, EvenRow, Align, BGColor) + Copy + kmaEndTableCell;
		}
		//
		//
		//
		internal static string kmaGetTableRow(string Cell, int ColSpan = 0, bool EvenRow = false)
		{
			return kmaStartTableRow() + kmaGetTableCell(Cell, "100%", ColSpan, EvenRow) + kmaEndTableRow;
		}
		//
		// remove the host and approotpath, leaving the "active" path and all else
		//
		internal static string kmaConvertShortLinkToLink(string URL, ref string PathPagePrefix)
		{
			string result = "";
			result = URL;
			if (URL != "" && PathPagePrefix != "")
			{
				if ((result.IndexOf(PathPagePrefix, StringComparison.CurrentCultureIgnoreCase) + 1) == 1)
				{
					result = result.Substring(Strings.Len(PathPagePrefix));
				}
			}
			return result;
		}
		//
		// ------------------------------------------------------------------------------------------------------
		//   Preserve URLs that do not start HTTP or HTTPS
		//   Preserve URLs from other sites (offsite)
		//   Preserve HTTP://ServerHost/ServerVirtualPath/Files/ in all cases
		//   Convert HTTP://ServerHost/ServerVirtualPath/folder/page -> /folder/page
		//   Convert HTTP://ServerHost/folder/page -> /folder/page
		// ------------------------------------------------------------------------------------------------------
		//
		internal static string kmaConvertLinkToShortLink(string URL, string ServerHost, string ServerVirtualPath)
		{
			//
			string BadString = "";
			string GoodString = "";
			string Protocol = "";
			//
			string WorkingLink = URL;
			//
			// ----- Determine Protocol
			//
			if ((WorkingLink.IndexOf("HTTP://", StringComparison.CurrentCultureIgnoreCase) + 1) == 1)
			{
				//
				// HTTP
				//
				Protocol = WorkingLink.Substring(0, Math.Min(7, WorkingLink.Length));
			}
			else if ((WorkingLink.IndexOf("HTTPS://", StringComparison.CurrentCultureIgnoreCase) + 1) == 1)
			{ 
				//
				// HTTPS
				//
				// try this -- a ssl link can not be shortened
				return WorkingLink;
				Protocol = WorkingLink.Substring(0, Math.Min(8, WorkingLink.Length));
			}
			if (Protocol != "")
			{
				//
				// ----- Protcol found, determine if is local
				//
				GoodString = Protocol + ServerHost;
				if ((WorkingLink.IndexOf(GoodString, StringComparison.CurrentCultureIgnoreCase) + 1) != 0)
				{
					//
					// URL starts with Protocol ServerHost
					//
					GoodString = Protocol + ServerHost + ServerVirtualPath + "/files/";
					if ((WorkingLink.IndexOf(GoodString, StringComparison.CurrentCultureIgnoreCase) + 1) != 0)
					{
						//
						// URL is in the virtual files directory
						//
						BadString = GoodString;
						GoodString = ServerVirtualPath + "/files/";
						WorkingLink = Strings.Replace(WorkingLink, BadString, GoodString, 1, -1, CompareMethod.Text);
					}
					else
					{
						//
						// URL is not in files virtual directory
						//
						BadString = Protocol + ServerHost + ServerVirtualPath + "/";
						GoodString = "/";
						WorkingLink = Strings.Replace(WorkingLink, BadString, GoodString, 1, -1, CompareMethod.Text);
						//
						BadString = Protocol + ServerHost + "/";
						GoodString = "/";
						WorkingLink = Strings.Replace(WorkingLink, BadString, GoodString, 1, -1, CompareMethod.Text);
					}
				}
			}
			return WorkingLink;
		}
		//
		// Correct the link for the virtual path, either add it or remove it
		//
		internal static string kmaEncodeAppRootPath(ref string Link, ref string VirtualPath, ref string AppRootPath, ref string ServerHost)
		{
			//
			string result = "";
			string Protocol = "";
			string Host = "";
			string Path = "";
			string Page = "";
			string QueryString = "";
			bool VirtualHosted = false;
			//
			result = Link;
			if (((result.IndexOf(ServerHost, StringComparison.CurrentCultureIgnoreCase) + 1) != 0) || ((Link.IndexOf('/') + 1) == 1))
			{
				//If (InStr(1, kmaEncodeAppRootPath, ServerHost, vbTextCompare) <> 0) And (InStr(1, Link, "/") <> 0) Then
				//
				// This link is onsite and has a path
				//
				VirtualHosted = ((AppRootPath.IndexOf(VirtualPath, StringComparison.CurrentCultureIgnoreCase) + 1) != 0);
				if (VirtualHosted && ((Link.IndexOf(AppRootPath, StringComparison.CurrentCultureIgnoreCase) + 1) == 1))
				{
					//
					// quick - virtual hosted and link starts at AppRootPath
					//
				}
				else if ((!VirtualHosted) && (Link.StartsWith("/")) && ((Link.IndexOf(AppRootPath, StringComparison.CurrentCultureIgnoreCase) + 1) == 1))
				{ 
					//
					// quick - not virtual hosted and link starts at Root
					//
				}
				else
				{
					SeparateURL(Link, ref Protocol, ref Host, ref Path, ref Page, ref QueryString);
					if (VirtualHosted)
					{
						//
						// Virtual hosted site, add VirualPath if it is not there
						//
						if ((Path.IndexOf(AppRootPath, StringComparison.CurrentCultureIgnoreCase) + 1) == 0)
						{
							if (Path == "/")
							{
								Path = AppRootPath;
							}
							else
							{
								Path = AppRootPath + Path.Substring(1);
							}
						}
					}
					else
					{
						//
						// Root hosted site, remove virtual path if it is there
						//
						if ((Path.IndexOf(AppRootPath, StringComparison.CurrentCultureIgnoreCase) + 1) != 0)
						{
							Path = Strings.Replace(Path, AppRootPath, "/", 1, -1, CompareMethod.Binary);
						}
					}
					result = Protocol + Host + Path + Page + QueryString;
				}
			}
			return result;
		}
		//
		// Return just the tablename from a tablename reference (database.object.tablename->tablename)
		//
		internal static string GetDbObjectTableName(string DbObject)
		{
			//
			string result = "";
			result = DbObject;
			int Position = result.LastIndexOf(".") + 1;
			if (Position > 0)
			{
				result = result.Substring(Position);
			}
			return result;
		}
		//
		//
		//
		internal static string kmaGetLinkedText(object AnchorTag, object AnchorText)
		{
			//
			string result = "";
			int LinkPosition = 0;
			//
			//
			string iAnchorTag = kmaEncodeText(AnchorTag);
			string iAnchorText = kmaEncodeText(AnchorText);
			string UcaseAnchorText = iAnchorText.ToUpper();
			if ((iAnchorTag != "") && (iAnchorText != ""))
			{
				LinkPosition = UcaseAnchorText.LastIndexOf("<LINK>") + 1;
				if (LinkPosition == 0)
				{
					result = iAnchorTag + iAnchorText + "</A>";
				}
				else
				{
					result = iAnchorText;
					LinkPosition = UcaseAnchorText.LastIndexOf("</LINK>") + 1;

					while(LinkPosition > 1)
					{
						result = result.Substring(0, Math.Min(LinkPosition - 1, result.Length)) + "</A>" + result.Substring(LinkPosition + 6);
						LinkPosition = UcaseAnchorText.LastIndexOf("<LINK>", LinkPosition - 2) + 1;
						if (LinkPosition != 0)
						{
							result = result.Substring(0, Math.Min(LinkPosition - 1, result.Length)) + iAnchorTag + result.Substring(LinkPosition + 5);
						}
						LinkPosition = UcaseAnchorText.LastIndexOf("</LINK>", LinkPosition - 1) + 1;
					};
				}
			}
			//
			return result;
		}
		//
		//========================================================================
		//   HandleError
		//       Logs the error and either resumes next, or raises it to the next level
		//========================================================================
		//
		internal static string HandleError(string ClassName, string MethodName, int ErrNumber, string ErrSource, string ErrDescription, bool ErrorTrap, bool ResumeNext, string URL = "")
		{
			// ##### removed to catch err<>0 problem on error resume next
			//
			string ErrorMessage = "";
			//
			if (ErrorTrap)
			{
				ErrorMessage = ErrorMessage + " Unexpected ErrorTrap";
			}
			else
			{
				ErrorMessage = ErrorMessage + " Error";
			}
			//
			if (URL != "")
			{
				ErrorMessage = ErrorMessage + " on page [" + URL + "]";
			}
			//
			if (ErrorTrap)
			{
				if (ResumeNext)
				{
					AppendLogFile(Path.GetFileNameWithoutExtension(Application.ExecutablePath) + "." + ClassName + "." + MethodName + ErrorMessage + ", will resume after logging [" + ErrSource + " #" + ErrNumber.ToString() + ", " + ErrDescription + "]");
				}
				else
				{
					AppendLogFile(Path.GetFileNameWithoutExtension(Application.ExecutablePath) + "." + ClassName + "." + MethodName + ErrorMessage + ", will abort after logging [" + ErrSource + " #" + ErrNumber.ToString() + ", " + ErrDescription + "]");
					UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Goto Label (0)");
					throw new System.Exception(ErrNumber.ToString() + ", " + ErrSource + ", " + ErrDescription);
				}
			}
			else
			{
				if (ResumeNext)
				{
					AppendLogFile(Path.GetFileNameWithoutExtension(Application.ExecutablePath) + "." + ClassName + "." + MethodName + ErrorMessage + ", will resume after logging  [" + ErrSource + " #" + ErrNumber.ToString() + ", " + ErrDescription + "]");
				}
				else
				{
					AppendLogFile(Path.GetFileNameWithoutExtension(Application.ExecutablePath) + "." + ClassName + "." + MethodName + ErrorMessage + ", will abort after logging [" + ErrSource + " #" + ErrNumber.ToString() + ", " + ErrDescription + "]");
					UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Goto Label (0)");
					throw new System.Exception(ErrNumber.ToString() + ", " + ErrSource + ", " + ErrDescription + ", " + -1);
				}
			}
			//
			return "";
		}
		//
		//==========================================================================================
		// handle error and resume next
		//==========================================================================================
		//
		internal static void HandleErrorAndResumeNext(string ClassName, string MethodName, string Description = "", int ErrorNumber = 0)
		{
			//
			//UPGRADE_WARNING: (2081) Err.Number has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2081.aspx
			int ErrNumber = Information.Err().Number;
			string ErrSource = Information.Err().Source;
			string ErrDescription = Information.Err().Description;
			//
			if (ErrNumber == 0)
			{
				//
				// internal error, no VB error
				//
				if (Description == "")
				{
					ErrDescription = "Unknown error";
				}
				else
				{
					ErrDescription = Description;
				}
				if (ErrorNumber == 0)
				{
					HandleError(ClassName, MethodName, KmaErrorInternal, Path.GetFileNameWithoutExtension(Application.ExecutablePath), Description, false, true);
				}
				else
				{
					HandleError(ClassName, MethodName, ErrNumber, Path.GetFileNameWithoutExtension(Application.ExecutablePath), Description, false, true);
				}
			}
			else
			{
				//
				// VB Error
				//
				if (Description != "")
				{
					ErrDescription = Description + ",VB Error [" + Information.Err().Description + "]";
				}
				else
				{
					ErrDescription = "Unexpected VB Error [" + Information.Err().Description + "]";
				}
				HandleError(ClassName, MethodName, ErrNumber, ErrSource, ErrDescription, true, true);
			}
		}
		//
		//
		//
		internal static void AppendLogFile(object Text)
		{
			UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Goto Label (0)");
			//
			//
			int DayNumber = DateAndTime.Day(DateTime.Now);
			int MonthNumber = DateTime.Now.Month;
			string Filename = DateTime.Now.Year.ToString();
			if (MonthNumber < 10)
			{
				Filename = Filename + "0";
			}
			Filename = Filename + MonthNumber.ToString();
			if (DayNumber < 10)
			{
				Filename = Filename + "0";
			}
			Filename = Filename + DayNumber.ToString();
			//
			AppendLog("Trace" + Filename + ".log", kmaEncodeText(Text));
			//
		}
		//
		//
		//
		internal static void AppendLog(string Filename, string Text)
		{
			UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Goto Label (0)");
			kmaFileSystem3.FileSystemClass kmafs = null;
			//
			if (Filename != "")
			{
				kmafs = new kmaFileSystem3.FileSystemClass();
				string tempRefParam = Path.GetDirectoryName(Application.ExecutablePath) + "\\logs\\" + Filename;
				string tempRefParam2 = "\"" + DateTime.Now.ToString() + "\",\"" + Text + "\"" + Environment.NewLine;
				kmafs.AppendFile(ref tempRefParam, ref tempRefParam2);
				kmafs = null;
			}
			//
		}
		//
		//
		//
		internal static string kmaErrorMsg(int ErrorNumber)
		{
			return "";
		}
		//
		//
		//
		internal static string kmaEncodeInitialCaps(string Source)
		{
			string result = "";
			string[] SegSplit = null;
			int SegMax = 0;
			//
			if (Source != "")
			{
				SegSplit = (string[]) Source.Split(' ');
				SegMax = SegSplit.GetUpperBound(0);
				if (SegMax >= 0)
				{
					for (int SegPtr = 0; SegPtr <= SegMax; SegPtr++)
					{
						SegSplit[SegPtr] = SegSplit[SegPtr].Substring(0, Math.Min(1, SegSplit[SegPtr].Length)).ToUpper() + SegSplit[SegPtr].Substring(1).ToLower();
					}
				}
				result = String.Join(" ", SegSplit);
			}
			return result;
		}
		//
		//
		//
		internal static string kmaRemoveTag(ref string Source, string TagName)
		{
			string result = "";
			int posEnd = 0;
			result = Source;
			int Pos = (Source.IndexOf("<" + TagName, StringComparison.CurrentCultureIgnoreCase) + 1);
			if (Pos != 0)
			{
				posEnd = Strings.InStr(Pos, Source, ">", CompareMethod.Binary);
				if (posEnd > 0)
				{
					result = Source.Substring(0, Math.Min(Pos - 1, Source.Length)) + Source.Substring(posEnd);
				}
			}
			return result;
		}
		//
		//
		//
		internal static string kmaRemoveStyleTags(string Source)
		{
			string result = "";
			result = Source;

			while((result.IndexOf("<style", StringComparison.CurrentCultureIgnoreCase) + 1) != 0)
			{
				string tempRefParam = result;
				result = kmaRemoveTag(ref tempRefParam, "style");
			};

			while((result.IndexOf("</style", StringComparison.CurrentCultureIgnoreCase) + 1) != 0)
			{
				string tempRefParam2 = result;
				result = kmaRemoveTag(ref tempRefParam2, "/style");
			};
			return result;
		}
		//
		//
		//
		internal static string kmaGetSingular(string PluralSource)
		{
			//
			string result = "";
			bool UpperCase = false;
			string LastCharacter = "";
			//
			result = PluralSource;
			if (Strings.Len(result) > 1)
			{
				LastCharacter = result.Substring(Math.Max(result.Length - 1, 0));
				if (LastCharacter != LastCharacter.ToUpper())
				{
					UpperCase = true;
				}
				if (result.Substring(Math.Max(result.Length - 3, 0)).ToUpper() == "IES")
				{
					if (UpperCase)
					{
						result = result.Substring(0, Math.Min(Strings.Len(result) - 3, result.Length)) + "Y";
					}
					else
					{
						result = result.Substring(0, Math.Min(Strings.Len(result) - 3, result.Length)) + "Y";
					}
				}
				else if (result.Substring(Math.Max(result.Length - 2, 0)).ToUpper() == "SS")
				{ 
					// nothing
				}
				else if (result.Substring(Math.Max(result.Length - 1, 0)).ToUpper() == "S")
				{ 
					result = result.Substring(0, Math.Min(Strings.Len(result) - 1, result.Length));
				}
				else
				{
					// nothing
				}
			}
			return result;
		}
		//
		//
		//
		internal static string kmaEncodeJavascript(string Source)
		{
			//
			string result = "";
			result = Source;
			result = Strings.Replace(result, "'", "\\'", 1, -1, CompareMethod.Binary);
			//kmaEncodeJavascript = Replace(kmaEncodeJavascript, "'", "'+""'""+'")
			result = Strings.Replace(result, Environment.NewLine, "\\n", 1, -1, CompareMethod.Binary);
			result = Strings.Replace(result, "\r", "\\n", 1, -1, CompareMethod.Binary);
			result = Strings.Replace(result, Constants.vbLf, "\\n", 1, -1, CompareMethod.Binary);
			return Strings.Replace(result, "</script", "</scr'+'ipt", 1, 99, CompareMethod.Text);
			//
		}
		//
		//   Indent every line by 1 tab
		//
		internal static string kmaIndent(ref string Source)
		{
			int posEnd = 0;
			string pre = "";
			string post = "";
			string target = "";
			//
			int posStart = (Source.IndexOf("<![CDATA[", StringComparison.CurrentCultureIgnoreCase) + 1);
			if (posStart == 0)
			{
				//
				// no cdata
				//
				posStart = (Source.IndexOf("<textarea", StringComparison.CurrentCultureIgnoreCase) + 1);
				if (posStart == 0)
				{
					//
					// no textarea
					//
					return Strings.Replace(Source, Environment.NewLine + "\t", Environment.NewLine + "\t" + "\t", 1, -1, CompareMethod.Binary);
				}
				else
				{
					//
					// text area found, isolate it and indent before and after
					//
					posEnd = Strings.InStr(posStart, Source, "</textarea>", CompareMethod.Text);
					pre = Source.Substring(0, Math.Min(posStart - 1, Source.Length));
					if (posEnd == 0)
					{
						target = Source.Substring(posStart - 1);
						post = "";
					}
					else
					{
						target = Source.Substring(posStart - 1, Math.Min(posEnd - posStart + Strings.Len("</textarea>"), Source.Length - (posStart - 1)));
						post = Source.Substring(posEnd + Strings.Len("</textarea>") - 1);
					}
					return kmaIndent(ref pre) + target + kmaIndent(ref post);
				}
			}
			else
			{
				//
				// cdata found, isolate it and indent before and after
				//
				posEnd = Strings.InStr(posStart, Source, "]]>", CompareMethod.Text);
				pre = Source.Substring(0, Math.Min(posStart - 1, Source.Length));
				if (posEnd == 0)
				{
					target = Source.Substring(posStart - 1);
					post = "";
				}
				else
				{
					target = Source.Substring(posStart - 1, Math.Min(posEnd - posStart + Strings.Len("]]>"), Source.Length - (posStart - 1)));
					post = Source.Substring(posEnd + 2);
				}
				return kmaIndent(ref pre) + target + kmaIndent(ref post);
			}
			//    kmaIndent = Source
			//    If InStr(1, kmaIndent, "<textarea", vbTextCompare) = 0 Then
			//        kmaIndent = Replace(Source, vbCrLf & vbTab, vbCrLf & vbTab & vbTab)
			//    End If
		}
		//
		//
		//
		internal static int kmaGetListIndex(string Item, string ListOfItems)
		{
			//
			int result = 0;
			string[] Items = null;
			string LcaseItem = "";
			string LcaseList = "";
			//
			if (ListOfItems != "")
			{
				LcaseItem = Item.ToLower();
				LcaseList = ListOfItems.ToLower();
				Items = (string[]) LcaseList.Split(',');
				for (int Ptr = 0; Ptr <= Items.GetUpperBound(0); Ptr++)
				{
					if (Items[Ptr] == LcaseItem)
					{
						result = Ptr + 1;
						break;
					}
				}
			}
			//
			return result;
		}
		//
		//========================================================================================================
		//
		// Finds all tags matching the input, and concatinates them into the output
		// does NOT account for nested tags, use for body, script, style
		//
		// ReturnAll - if true, it returns all the occurances, back-to-back
		//
		//========================================================================================================
		//
		internal static string GetTagInnerHTML(ref string PageSource, string Tag, bool ReturnAll)
		{
			string result = "";
			try
			{
				//
				int TagStart = 0;
				int TagEnd = 0;
				int LoopCnt = 0;
				string WB = "";
				int Pos = 0;
				int posEnd = 0;
				int CommentPos = 0;
				int ScriptPos = 0;
				//
				Pos = 1;

				while((Pos > 0) && (LoopCnt < 100))
				{
					TagStart = Strings.InStr(Pos, PageSource, "<" + Tag, CompareMethod.Text);
					if (TagStart == 0)
					{
						Pos = 0;
					}
					else
					{
						//
						// tag found, skip any comments that start between current position and the tag
						//
						CommentPos = Strings.InStr(Pos, PageSource, "<!--", CompareMethod.Binary);
						if ((CommentPos != 0) && (CommentPos < TagStart))
						{
							//
							// skip comment and start again
							//
							Pos = Strings.InStr(CommentPos, PageSource, "-->", CompareMethod.Binary);
						}
						else
						{
							ScriptPos = Strings.InStr(Pos, PageSource, "<script", CompareMethod.Binary);
							if ((ScriptPos != 0) && (ScriptPos < TagStart))
							{
								//
								// skip comment and start again
								//
								Pos = Strings.InStr(ScriptPos, PageSource, "</script", CompareMethod.Binary);
							}
							else
							{
								//
								// Get the tags innerHTML
								//
								TagStart = Strings.InStr(TagStart, PageSource, ">", CompareMethod.Text);
								Pos = TagStart;
								if (TagStart != 0)
								{
									TagStart++;
									TagEnd = Strings.InStr(TagStart, PageSource, "</" + Tag, CompareMethod.Text);
									if (TagEnd != 0)
									{
										result = result + PageSource.Substring(TagStart - 1, Math.Min(TagEnd - TagStart, PageSource.Length - (TagStart - 1)));
									}
								}
							}
						}
						LoopCnt++;
						if (ReturnAll)
						{
							TagStart = Strings.InStr(TagEnd, PageSource, "<" + Tag, CompareMethod.Text);
						}
						else
						{
							TagStart = 0;
						}
					}
				};
				//
			}
			catch
			{
				//
				//Call HandleError("EncodePage_SplitBody")
			}
			return result;
		}
        ////
        ////========================================================================================================
        ////Place code in a form module
        ////Add a Command button.
        ////========================================================================================================
        ////
        //internal static string kmaByteArrayToString(byte[] Bytes)
        //{
        //    string result = "";
        //    int iUnicode = 0;

        //    //UPGRADE_TODO: (1069) Error handling statement (On Error Resume Next) was converted to a pattern that might have a different behavior. More Information: http://www.vbtonet.com/ewis/ewi1069.aspx
        //    try
        //    {
        //        int i = Bytes.GetUpperBound(0);

        //        if (i < 1)
        //        {
        //            //ANSI, just convert to unicode and return
        //            //UPGRADE_WARNING: (1059) Code was upgraded to use UpgradeHelpers.Helpers.StringsHelper.ByteArrayToString() which may not have the same behavior. More Information: http://www.vbtonet.com/ewis/ewi1059.aspx
        //            return StringsHelper.StrConv(StringsHelper.ByteArrayToString(Bytes), StringsHelper.VbStrConvEnum.VbUnicode, 0);
        //        }
        //        i++;

        //        //Examine the first two bytes
        //        UpgradeSolution1Support.PInvoke.SafeNative.kernel32.CopyMemory(ref iUnicode, ref Bytes[0], 2);

        //        if (iUnicode == Bytes[0])
        //        { //Unicode
        //            //Account for terminating null
        //            if ((i % 2 != 0))
        //            {
        //                i--;
        //            }
        //            //Set up a buffer to recieve the string
        //            result = new string((char) 0, i / 2);
        //            //Copy to string
        //            GCHandle gh = GCHandle.Alloc(result, GCHandleType.Pinned);
        //            IntPtr tmpPtr = gh.AddrOfPinnedObject();
        //            UpgradeSolution1Support.PInvoke.SafeNative.kernel32.CopyMemory(tmpPtr.ToInt32(), ref Bytes[0], i);
        //            gh.Free();
        //        }
        //        else
        //        {
        //            //ANSI
        //            //UPGRADE_WARNING: (1059) Code was upgraded to use UpgradeHelpers.Helpers.StringsHelper.ByteArrayToString() which may not have the same behavior. More Information: http://www.vbtonet.com/ewis/ewi1059.aspx
        //            result = StringsHelper.StrConv(StringsHelper.ByteArrayToString(Bytes), StringsHelper.VbStrConvEnum.VbUnicode, 0);
        //        }
        //    }
        //    catch (Exception exc)
        //    {
        //        NotUpgradedHelper.NotifyNotUpgradedElement("Resume in On-Error-Resume-Next Block");
        //    }

        //    return result;
        //}
		//
		//========================================================================================================
		//
		//========================================================================================================
		//
		internal static byte[] kmaStringToByteArray(string strInput, bool bReturnAsUnicode = true, bool bAddNullTerminator = false)
		{

			int lRet = 0;
			byte[] bytBuffer = null;
			int lLenB = 0;

			if (bReturnAsUnicode)
			{
				//Number of bytes
				//UPGRADE_WARNING: (2081) LenB has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2081.aspx
				lLenB = Encoding.Unicode.GetByteCount(strInput);
				//Resize buffer, do we want terminating null?
				if (bAddNullTerminator)
				{
					bytBuffer = new byte[lLenB + 1];
				}
				else
				{
					bytBuffer = new byte[lLenB];
				}
				//Copy characters from string to byte array
				GCHandle gh = GCHandle.Alloc(strInput, GCHandleType.Pinned);
				IntPtr tmpPtr = gh.AddrOfPinnedObject();
				UpgradeSolution1Support.PInvoke.SafeNative.kernel32.CopyMemory(ref bytBuffer[0], tmpPtr.ToInt32(), lLenB);
				gh.Free();
			}
			else
			{
				//METHOD ONE
				//        'Get rid of embedded nulls
				//        strRet = StrConv(strInput, vbFromUnicode)
				//        lLenB = LenB(strRet)
				//        If bAddNullTerminator Then
				//            ReDim bytBuffer(lLenB)
				//        Else
				//            ReDim bytBuffer(lLenB - 1)
				//        End If
				//        CopyMemory bytBuffer(0), ByVal StrPtr(strInput), lLenB

				//METHOD TWO
				//Num of characters
				lLenB = Strings.Len(strInput);
				if (bAddNullTerminator)
				{
					bytBuffer = new byte[lLenB + 1];
				}
				else
				{
					bytBuffer = new byte[lLenB];
				}
				GCHandle gh3 = GCHandle.Alloc(bytBuffer[0], GCHandleType.Pinned);
				IntPtr tmpPtr3 = gh3.AddrOfPinnedObject();
				GCHandle gh2 = GCHandle.Alloc(strInput, GCHandleType.Pinned);
				IntPtr tmpPtr2 = gh2.AddrOfPinnedObject();
				lRet = UpgradeSolution1Support.PInvoke.SafeNative.kernel32.WideCharToMultiByte(CP_ACP, 0, tmpPtr2.ToInt32(), -1, tmpPtr3.ToInt32(), lLenB, 0, 0);
				gh3.Free();
				gh2.Free();
			}

			return bytBuffer;

		}
		//
		//========================================================================================================
		//   Sample kmaStringToByteArray
		//========================================================================================================
		//
		//UPGRADE_NOTE: (7001) The following declaration (SampleStringToByteArray) seems to be dead code More Information: http://www.vbtonet.com/ewis/ewi7001.aspx
		//private void SampleStringToByteArray()
		//{
			////
			//string str = "Convert";
			//byte[] bAnsi = (byte[]) kmaStringToByteArray(str, false);
			//byte[] bUni = (byte[]) kmaStringToByteArray(str);
			////
			//for (int i = 0; i <= bAnsi.GetUpperBound(0); i++)
			//{
				//Debug.WriteLine("=" + bAnsi[i].ToString());
			//}
			////
			//Debug.WriteLine("========");
			////
			//for (int i = 0; i <= bUni.GetUpperBound(0); i++)
			//{
				//Debug.WriteLine("=" + bUni[i].ToString());
			//}
			////
			//Debug.WriteLine("ANSI= " + kmaByteArrayToString(bAnsi));
			//Debug.WriteLine("UNICODE= " + kmaByteArrayToString(bUni));
			//Using StrConv to convert a Unicode character array directly
			//will cause the resultant string to have extra embedded nulls
			//reason, StrConv does not know the difference between Unicode and ANSI
			////UPGRADE_WARNING: (1059) Code was upgraded to use UpgradeHelpers.Helpers.StringsHelper.ByteArrayToString() which may not have the same behavior. More Information: http://www.vbtonet.com/ewis/ewi1059.aspx
			//Debug.WriteLine("Resull= " + StringsHelper.StrConv(StringsHelper.ByteArrayToString(bUni), StringsHelper.VbStrConvEnum.VbUnicode, 0));
		//}
	}
}