using System.Runtime.InteropServices;
using System;

namespace UpgradeSolution1Support.PInvoke.UnsafeNative
{
	[System.Security.SuppressUnmanagedCodeSecurity]
	public static class kernel32
	{

		[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		extern public static int CloseHandle(int hObject);
		[DllImport("kernel32.dll", EntryPoint = "RtlMoveMemory", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		extern public static void CopyMemory(System.IntPtr hpvDest, System.IntPtr hpvSource, int cbCopy);
		[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		extern public static int GetCurrentProcess();
		[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		extern public static int GetCurrentProcessId();
		[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		extern public static int GetCurrentThread();
		[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		extern public static int GetExitCodeProcess(int hProcess, ref int lpExitCode);
		[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		extern public static int GetPriorityClass(int hProcess);
		[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		extern public static int GetThreadPriority(int hThread);
		[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		extern public static int GetTickCount();
		[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		extern public static int OpenProcess(int dwDesiredAccess, int bInheritHandle, int dwProcessId);
		[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		extern public static int SetPriorityClass(int hProcess, int dwPriorityClass);
		[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		extern public static int SetThreadPriority(int hThread, int nPriority);
		[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		extern public static void Sleep(int dwMilliseconds);
		[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		extern public static int WideCharToMultiByte(int CodePage, int dwFlags, int lpWideCharStr, int cchWideChar, int lpMultiByteStr, int cbMultiByte, int lpDefaultChar, int lpUsedDefaultChar);
	}
}