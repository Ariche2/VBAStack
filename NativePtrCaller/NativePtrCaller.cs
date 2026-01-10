using System;

namespace NativePtrCaller
{
    public static unsafe class NativePtrCaller
    {
        public static int EbMode(IntPtr pEbMode)
        {
            delegate* unmanaged[Stdcall]<int> del = (delegate* unmanaged[Stdcall]<int>)pEbMode;
            return del();
        }
        public static int EbSetMode(IntPtr pEbSetMode, int Mode)
        {
            delegate* unmanaged[Stdcall]<int, int> del = (delegate* unmanaged[Stdcall]<int, int>)pEbSetMode;
            return del(Mode);
        }
        public static int EbGetCallstackCount(IntPtr pEbCallstackCount, ref int Count)
        {
            delegate* unmanaged[Stdcall]<ref int, int> del = (delegate* unmanaged[Stdcall]<ref int, int>)pEbCallstackCount;
            return del(ref Count);
        }
        public static int ErrGetCallstackString(IntPtr pErrGetCallstackString, int StackIndex, ref string CallstackString, ref int mysteryNumber)
        {
            delegate* unmanaged[Stdcall]<int, IntPtr, ref int, int> del;
            del = (delegate* unmanaged[Stdcall]<int, IntPtr, ref int, int>)pErrGetCallstackString;
            using LPSTR pCallstackString = new(255);
            int returnVal = del(StackIndex, pCallstackString, ref mysteryNumber);
            CallstackString = pCallstackString.ToString();
            return returnVal;
        }
    }
}
