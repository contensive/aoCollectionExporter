using System.Runtime.InteropServices;
using System;

namespace UpgradeSolution1Support.PInvoke.UnsafeNative
{
	[System.Security.SuppressUnmanagedCodeSecurity]
	public static class winmm
	{

		[DllImport("winmm.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		extern public static int timeGetTime();
	}
}