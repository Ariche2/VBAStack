using System;
using System.Runtime.InteropServices;

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
        public static IntPtr BtrootOfExframe(IntPtr pBtrootOfExframe, IntPtr pExFrame)
        {
            delegate* unmanaged[Stdcall]<IntPtr, IntPtr> del = (delegate* unmanaged[Stdcall]<IntPtr, IntPtr>)pBtrootOfExframe;
            return del(pExFrame);
        }
        public static IntPtr ExecGetExframeTOS(IntPtr pExecGetExframeTOS)
        {
            delegate* unmanaged[Stdcall]<IntPtr> del = (delegate* unmanaged[Stdcall]<IntPtr>)pExecGetExframeTOS;
            return del();
        }
        public static int BASIC_TYPEROOT_GetBTSrc(IntPtr pGetBTSrc, IntPtr pThis, ref IntPtr ppBTSrc)
        {
            fixed (IntPtr* ptr = &ppBTSrc)
            {
                if (IntPtr.Size == 8)
                {
                    // this goes in RCX, first param in RDX
                    delegate* unmanaged[Fastcall]<IntPtr, IntPtr*, int> del = (delegate* unmanaged[Fastcall]<IntPtr, IntPtr*, int>)pGetBTSrc;
                    return del(pThis, ptr);
                }
                else
                {
                    // Microsoft's C# compiler should handle ECX for us
                    var del = (delegate* unmanaged[Thiscall]<IntPtr, IntPtr*, int>)pGetBTSrc;
                    return del(pThis, ptr);
                }
            }
        }
        public static int BASIC_TYPESRC_GetFrameNames(IntPtr pGetFrameNames, IntPtr pThis, IntPtr pExFrame, out string moduleName, out string functionName)
        {
            IntPtr pModuleBstr = IntPtr.Zero;
            IntPtr pFunctionBstr = IntPtr.Zero;
            moduleName = null;
            functionName = null;
            
            try
            {
                int result;
                if (IntPtr.Size == 8)
                {
                    // x64: Use Windows x64 calling convention (fastcall variant)
                    delegate* unmanaged[Fastcall]<IntPtr, IntPtr, IntPtr*, IntPtr*, int> del = (delegate* unmanaged[Fastcall]<IntPtr, IntPtr, IntPtr*, IntPtr*, int>)pGetFrameNames;
                    result = del(pThis, pExFrame, &pModuleBstr, &pFunctionBstr);
                }
                else
                {
                    // x86: Use thiscall
                    delegate* unmanaged[Thiscall]<IntPtr, IntPtr, IntPtr*, IntPtr*, int> del = (delegate* unmanaged[Thiscall]<IntPtr, IntPtr, IntPtr*, IntPtr*, int>)pGetFrameNames;
                    result = del(pThis, pExFrame, &pModuleBstr, &pFunctionBstr);
                }
                
                if (result >= 0)
                {
                    // GetFrameNames returns ANSI BSTRs, not Unicode BSTRs
                    // Must use PtrToStringAnsi, not PtrToStringBSTR
                    if (pModuleBstr != IntPtr.Zero)
                    {
                        moduleName = Marshal.PtrToStringAnsi(pModuleBstr);
                    }
                    
                    if (pFunctionBstr != IntPtr.Zero)
                    {
                        functionName = Marshal.PtrToStringAnsi(pFunctionBstr);
                    }
                }
                
                return result;
            }
            finally
            {   //Free the ANSI BSTRs allocated by the native function
                // Note: These are ANSI BSTRs, but FreeBSTR works for both ANSI and Unicode
                if (pModuleBstr != IntPtr.Zero)
                {
                    Marshal.FreeBSTR(pModuleBstr);
                }
                
                if (pFunctionBstr != IntPtr.Zero)
                {
                    Marshal.FreeBSTR(pFunctionBstr);
                }                

            }
        }

        /// <summary>
        /// Calls epiModule::MemidOfPrtmi to get a Member ID from an RTMI pointer.
        /// Signature: long __thiscall epiModule::MemidOfPrtmi(epiModule *this, RTMI *pRtmi, tagINVOKEKIND *pInvokeKind)
        /// Returns: Memid as int, or -1 on failure
        /// </summary>
        public static int EpiModule_MemidOfPrtmi(IntPtr pMemidOfPrtmi, IntPtr pThis, IntPtr pRtmi, out int invokeKind)
        {
            int invKind = 0;
            int result;
            
            if (IntPtr.Size == 8)
            {
                // x64: this in RCX, pRtmi in RDX, pInvokeKind in R8
                delegate* unmanaged[Fastcall]<IntPtr, IntPtr, int*, int> del = (delegate* unmanaged[Fastcall]<IntPtr, IntPtr, int*, int>)pMemidOfPrtmi;
                result = del(pThis, pRtmi, &invKind);
            }
            else
            {
                // x86: thiscall, this in ECX, pRtmi and pInvokeKind on stack
                delegate* unmanaged[Thiscall]<IntPtr, IntPtr, int*, int> del = (delegate* unmanaged[Thiscall]<IntPtr, IntPtr, int*, int>)pMemidOfPrtmi;
                result = del(pThis, pRtmi, &invKind);
            }
            
            invokeKind = invKind;
            return result;
        }

        /// <summary>
        /// Calls TYPE_DATA::HfdefnOfHmember to get a function definition handle from a Memid.
        /// Signature: ulong __thiscall TYPE_DATA::HfdefnOfHmember(TYPE_DATA *this, ulong memid, uint invokeKindFlags)
        /// Returns: Hfdefn handle as IntPtr, or 0xFFFFFFFF on failure
        /// </summary>
        public static IntPtr TYPE_DATA_HfdefnOfHmember(IntPtr pHfdefnOfHmember, IntPtr pThis, int memid, uint invokeKindFlags)
        {
            if (IntPtr.Size == 8)
            {
                // x64: this in RCX, memid in RDX, invokeKindFlags in R8
                delegate* unmanaged[Fastcall]<IntPtr, int, uint, IntPtr> del = (delegate* unmanaged[Fastcall]<IntPtr, int, uint, IntPtr>)pHfdefnOfHmember;
                return del(pThis, memid, invokeKindFlags);
            }
            else
            {
                // x86: thiscall, this in ECX, parameters on stack
                delegate* unmanaged[Thiscall]<IntPtr, int, uint, IntPtr> del = (delegate* unmanaged[Thiscall]<IntPtr, int, uint, IntPtr>)pHfdefnOfHmember;
                return del(pThis, memid, invokeKindFlags);
            }
        }

        /// <summary>
        /// Calls BASIC_TYPEINFO::GetFunctionNameOfHfdefn to get the function name from an Hfdefn.
        /// Signature: long __thiscall BASIC_TYPEINFO::GetFunctionNameOfHfdefn(BASIC_TYPEINFO *this, ulong hfdefn, char **pName, tagINVOKEKIND *pInvokeKind)
        /// Returns: HRESULT (0 = success, negative = error)
        /// </summary>
        public static int BASIC_TYPEINFO_GetFunctionNameOfHfdefn(IntPtr pGetFunctionNameOfHfdefn, IntPtr pThis, IntPtr hfdefn, out string functionName, out int invokeKind)
        {
            IntPtr pNameBstr = IntPtr.Zero;
            int invKind = 0;
            functionName = null;
            
            try
            {
                int result;
                if (IntPtr.Size == 8)
                {
                    // x64: this in RCX, hfdefn in RDX, ppName in R8, pInvokeKind in R9
                    delegate* unmanaged[Fastcall]<IntPtr, IntPtr, IntPtr*, int*, int> del = (delegate* unmanaged[Fastcall]<IntPtr, IntPtr, IntPtr*, int*, int>)pGetFunctionNameOfHfdefn;
                    result = del(pThis, hfdefn, &pNameBstr, &invKind);
                }
                else
                {
                    // x86: thiscall, this in ECX, parameters on stack
                    delegate* unmanaged[Thiscall]<IntPtr, IntPtr, IntPtr*, int*, int> del = (delegate* unmanaged[Thiscall]<IntPtr, IntPtr, IntPtr*, int*, int>)pGetFunctionNameOfHfdefn;
                    result = del(pThis, hfdefn, &pNameBstr, &invKind);
                }
                
                if (result >= 0 && pNameBstr != IntPtr.Zero)
                {
                    // VBE returns ANSI strings (char*), not Unicode BSTRs
                    functionName = Marshal.PtrToStringAnsi(pNameBstr);
                    Marshal.FreeCoTaskMem(pNameBstr);
                }
                
                invokeKind = invKind;
                return result;
            }
            catch
            {
                if (pNameBstr != IntPtr.Zero)
                {
                    Marshal.FreeCoTaskMem(pNameBstr);
                }
                throw;
            }
        }

        /// <summary>
        /// Calls GetBtinfoOfExframe to get BASIC_TYPEINFO from an EXFRAME.
        /// Signature: BASIC_TYPEINFO* __stdcall GetBtinfoOfExframe(EXFRAME *pExFrame)
        /// Returns: Pointer to BASIC_TYPEINFO structure
        /// </summary>
        public static IntPtr GetBtinfoOfExframe(IntPtr pGetBtinfoOfExframe, IntPtr pExFrame)
        {
            delegate* unmanaged[Stdcall]<IntPtr, IntPtr> del = (delegate* unmanaged[Stdcall]<IntPtr, IntPtr>)pGetBtinfoOfExframe;
            return del(pExFrame);
        }

        /// <summary>
        /// Calls TipGetModuleName to get the module name from an ITypeInfo pointer.
        /// This is a safer wrapper than calling BASIC_TYPEINFO::GetName directly (which has multiple overloads).
        /// Signature: long __stdcall TipGetModuleName(ITypeInfo *pTypeInfo, BSTR *ppModuleName)
        /// Returns: VBE error code (0 = success, non-zero = error)
        /// </summary>
        public static int TipGetModuleName(IntPtr pTipGetModuleName, IntPtr pTypeInfo, out string moduleName)
        {
            IntPtr pNameBstr = IntPtr.Zero;
            moduleName = null;
            
            try
            {
                // TipGetModuleName uses stdcall convention
                delegate* unmanaged[Stdcall]<IntPtr, IntPtr*, int> del = (delegate* unmanaged[Stdcall]<IntPtr, IntPtr*, int>)pTipGetModuleName;
                int result = del(pTypeInfo, &pNameBstr);
                
                if (result == 0 && pNameBstr != IntPtr.Zero)
                {
                    // VBE returns ANSI strings (char*), not Unicode BSTRs
                    moduleName = Marshal.PtrToStringAnsi(pNameBstr);
                    Marshal.FreeBSTR(pNameBstr);
                }
                
                return result;
            }
            catch
            {
                if (pNameBstr != IntPtr.Zero)
                {
                    Marshal.FreeBSTR(pNameBstr);
                }
                throw;
            }
        }
    }
}
