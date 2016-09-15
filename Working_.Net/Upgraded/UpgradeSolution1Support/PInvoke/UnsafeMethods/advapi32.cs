using System.Runtime.InteropServices;
using System;

namespace UpgradeSolution1Support.PInvoke.UnsafeNative
{
	[System.Security.SuppressUnmanagedCodeSecurity]
	public static class advapi32
	{

		[DllImport("advapi32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		extern public static int RegCloseKey(int hKey);
		[DllImport("advapi32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		extern public static int RegOpenKeyExA(int hKey, [MarshalAs(UnmanagedType.VBByRefStr)] ref string lpszSubKey, ref int dwOptions, int samDesired, ref int lpHKey);
		[DllImport("advapi32.dll", EntryPoint = "RegQueryValueExA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		extern public static int RegQueryValueEx(int hKey, [MarshalAs(UnmanagedType.VBByRefStr)] ref string lpszValueName, int lpdwRes, ref int lpdwType, ref int lpDataBuff, ref int nSize);
		[DllImport("advapi32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		extern public static int RegQueryValueExA(int hKey, [MarshalAs(UnmanagedType.VBByRefStr)] ref string lpszValueName, int lpdwRes, ref int lpdwType, [MarshalAs(UnmanagedType.VBByRefStr)] ref string lpDataBuff, ref int nSize);
	}
}