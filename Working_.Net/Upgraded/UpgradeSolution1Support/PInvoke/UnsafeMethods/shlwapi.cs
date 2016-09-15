using System.Runtime.InteropServices;
using System;

namespace UpgradeSolution1Support.PInvoke.UnsafeNative
{
	[System.Security.SuppressUnmanagedCodeSecurity]
	public static class shlwapi
	{

		[DllImport("shlwapi.dll", EntryPoint = "UrlEscapeA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		extern public static int UrlEscape([MarshalAs(UnmanagedType.VBByRefStr)] ref string pszURL, [MarshalAs(UnmanagedType.VBByRefStr)] ref string pszEscaped, ref int pcchEscaped, int dwFlags);
		[DllImport("shlwapi.dll", EntryPoint = "UrlUnescapeA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		extern public static int UrlUnescape([MarshalAs(UnmanagedType.VBByRefStr)] ref string pszURL, [MarshalAs(UnmanagedType.VBByRefStr)] ref string pszUnescaped, ref int pcchUnescaped, int dwFlags);
	}
}