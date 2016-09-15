using System.Runtime.InteropServices;
using System;

namespace UpgradeSolution1Support.PInvoke.SafeNative
{
	public static class kernel32
	{

		public static void CopyMemory(ref int hpvDest, ref byte hpvSource, int cbCopy)
		{
			GCHandle handle = GCHandle.Alloc(hpvDest, GCHandleType.Pinned);
			GCHandle handle2 = GCHandle.Alloc(hpvSource, GCHandleType.Pinned);
			try
			{
				IntPtr tmpPtr2 = handle2.AddrOfPinnedObject();
				IntPtr tmpPtr = handle.AddrOfPinnedObject();
				UpgradeSolution1Support.PInvoke.UnsafeNative.kernel32.CopyMemory(tmpPtr, tmpPtr2, cbCopy);
				hpvSource = Marshal.ReadByte(tmpPtr2);
				hpvDest = Marshal.ReadInt32(tmpPtr);
			}
			finally
			{
				handle.Free();
				handle2.Free();
			}
		}
		public static void CopyMemory(ref byte hpvDest, int hpvSource, int cbCopy)
		{
			GCHandle handle = GCHandle.Alloc(hpvDest, GCHandleType.Pinned);
			GCHandle handle2 = GCHandle.Alloc(hpvSource, GCHandleType.Pinned);
			try
			{
				IntPtr tmpPtr2 = handle2.AddrOfPinnedObject();
				IntPtr tmpPtr = handle.AddrOfPinnedObject();
				UpgradeSolution1Support.PInvoke.UnsafeNative.kernel32.CopyMemory(tmpPtr, tmpPtr2, cbCopy);
				hpvSource = Marshal.ReadInt32(tmpPtr2);
				hpvDest = Marshal.ReadByte(tmpPtr);
			}
			finally
			{
				handle.Free();
				handle2.Free();
			}
		}
		public static int GetCurrentProcess()
		{
			return UpgradeSolution1Support.PInvoke.UnsafeNative.kernel32.GetCurrentProcess();
		}
		public static int GetCurrentProcessId()
		{
			return UpgradeSolution1Support.PInvoke.UnsafeNative.kernel32.GetCurrentProcessId();
		}
		public static int SetPriorityClass(int hProcess, int dwPriorityClass)
		{
			return UpgradeSolution1Support.PInvoke.UnsafeNative.kernel32.SetPriorityClass(hProcess, dwPriorityClass);
		}
		public static void Sleep(int dwMilliseconds)
		{
			UpgradeSolution1Support.PInvoke.UnsafeNative.kernel32.Sleep(dwMilliseconds);
		}
		public static int WideCharToMultiByte(int CodePage, int dwFlags, int lpWideCharStr, int cchWideChar, int lpMultiByteStr, int cbMultiByte, int lpDefaultChar, int lpUsedDefaultChar)
		{
			return UpgradeSolution1Support.PInvoke.UnsafeNative.kernel32.WideCharToMultiByte(CodePage, dwFlags, lpWideCharStr, cchWideChar, lpMultiByteStr, cbMultiByte, lpDefaultChar, lpUsedDefaultChar);
		}
	}
}