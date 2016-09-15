using System.Runtime.InteropServices;
using System;

namespace UpgradeSolution1Support.PInvoke.SafeNative
{
	public static class advapi32
	{

		public static int RegCloseKey(int hKey)
		{
			return UpgradeSolution1Support.PInvoke.UnsafeNative.advapi32.RegCloseKey(hKey);
		}
		public static int RegOpenKeyExA(int hKey, ref string lpszSubKey, ref int dwOptions, int samDesired, ref int lpHKey)
		{
			return UpgradeSolution1Support.PInvoke.UnsafeNative.advapi32.RegOpenKeyExA(hKey, ref lpszSubKey, ref dwOptions, samDesired, ref lpHKey);
		}
		public static int RegQueryValueEx(int hKey, ref string lpszValueName, int lpdwRes, ref int lpdwType, ref int lpDataBuff, ref int nSize)
		{
			return UpgradeSolution1Support.PInvoke.UnsafeNative.advapi32.RegQueryValueEx(hKey, ref lpszValueName, lpdwRes, ref lpdwType, ref lpDataBuff, ref nSize);
		}
		public static int RegQueryValueExA(int hKey, ref string lpszValueName, int lpdwRes, ref int lpdwType, ref string lpDataBuff, ref int nSize)
		{
			return UpgradeSolution1Support.PInvoke.UnsafeNative.advapi32.RegQueryValueExA(hKey, ref lpszValueName, lpdwRes, ref lpdwType, ref lpDataBuff, ref nSize);
		}
	}
}