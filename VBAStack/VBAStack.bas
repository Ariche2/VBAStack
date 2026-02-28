Attribute VB_Name = "VBAStack"
Option Explicit

'If these show up in red for you, ignore them. I promise.
#If VBA7 = False Then
    Private Enum LongPtr
        [_]
    End Enum
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByVal lpSource As LongPtr, ByVal cbCopy As Long)
    Private Declare Function SysReAllocString Lib "OleAut32" (ByVal pBSTR As LongPtr, ByVal psz As LongPtr) As Long
    Private Declare Function VariantCopy Lib "OleAut32" (ByRef pVarDest As Variant, ByVal pVarSource As LongPtr) As Long
    Private Declare Function VirtualQuery Lib "kernel32" (ByVal lpAddress As LongPtr, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As LongPtr) As LongPtr
    Private Declare Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As LongPtr, ByRef retVAR As Variant) As Long
    Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As LongPtr, ByRef lpiid As GUIDt) As Long
#Else
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByVal lpSource As LongPtr, ByVal cbCopy As Long)
    Private Declare PtrSafe Function SysReAllocString Lib "OleAut32" (ByVal pDestBSTR As LongPtr, ByVal pSourceBSTR As LongPtr) As Long
    Private Declare PtrSafe Function VariantCopy Lib "OleAut32" (ByRef pVarDest As Variant, ByVal pVarSource As LongPtr) As Long
    Private Declare PtrSafe Function VirtualQuery Lib "kernel32" (ByVal lpAddress As LongPtr, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As LongPtr) As LongPtr
    Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As LongPtr, ByRef retVAR As Variant) As Long
    Private Declare PtrSafe Function IIDFromString Lib "ole32.dll" (ByVal lpsz As LongPtr, ByRef lpiid As GUIDt) As Long
#End If


Private Enum AllocationProtectEnum
    PAGE_EXECUTE = &H10
    PAGE_EXECUTE_READ = &H20
    PAGE_EXECUTE_READWRITE = &H40
    PAGE_EXECUTE_WRITECOPY = &H80
    PAGE_NOACCESS = &H1
    PAGE_READONLY = &H2
    PAGE_READWRITE = &H4
    PAGE_WRITECOPY = &H8
    PAGE_GUARD = &H100
    PAGE_NOCACHE = &H200
    PAGE_WRITECOMBINE = &H400
End Enum

Private Type MEMORY_BASIC_INFORMATION
    BaseAddress As LongPtr
    AllocationBase As LongPtr
    AllocationProtect As AllocationProtectEnum
    RegionSize As LongPtr
    State As StateEnum
    Protect As AllocationProtectEnum
    lType As TypeEnum
End Type

Private Enum StateEnum
    MEM_COMMIT = &H1000
    MEM_FREE = &H10000
    MEM_RESERVE = &H2000
End Enum

Private Enum TypeEnum
    MEM_IMAGE = &H1000000
    MEM_MAPPED = &H40000
    MEM_PRIVATE = &H20000
End Enum


#If Win64 Then
    Const PtrSize As Integer = 8
#Else
    Const PtrSize As Integer = 4
#End If


#Const DEBUGConst = False

Private SafeAddressCache() As AddressRangeSafety

Private Type AddressRangeSafety
    pRangeStart As LongPtr
    pRangeEnd As LongPtr
    Safe As Boolean
    MBI As MEMORY_BASIC_INFORMATION 'This only gets populated if the address safety check fails
End Type

Public Type StackFrame
    pExFrame As Currency 'Stored like this so can still use StackFrame outside this module, even in VBA6 - people might already have their own version of the "LongPtr" trick declared (LibMemory uses it, for example) so want to keep mine Private
    ProjectName As String
    ObjectName As String
    ProcedureName As String
    FrameNumber As Integer
    MethodIndex As Integer
    Errored As Boolean
End Type

Public Type paramInfo
    ParamName As String
    TypeName As String
    TypeEnumVal As VbInternal_Type
    Value As String
    IsByRef As Boolean
    IsArray As Boolean
    IsOptional As Boolean
    ParamSize As Byte
    DataSize As Byte
    hasExtraPointer As Boolean
    pExtraData As Currency 'Stored like this for same reason as StackFrame.pExFrame
    Errored As Boolean
End Type

'Directly and shamelessly copied from David Zimmer, again.

'had to manually map these out watching variations
'internal to vb not variant types
Public Enum VbInternal_Type
#If Win64 = False Then
    VbIT_Boolean = 3
    VbIT_Byte = 5
    VbIT_Integer = 6
    VbIT_Long = 8
    VbIT_Single = &HA
    VbIT_Double = &HB
    VbIT_Date = &HC
    VbIT_Currency = &HD
    VbIT_Variant = &HF
    VbIT_String = &H10
    VbIT_UserDefinedType = &H14
    VbIT_EnumMaybe = &H18
    VbIT_Object = &H1B

    VbIT_Internal = &H13 ' \
    VbIT_ComIface = &H1C '  \__ these add 32bit pointer <--- oh god he means literally. so if we encounter one of these, the next 4 (or 8) bytes are a pointer to a structure with COM info. So this'll need a special case to skip those bytes when reading arg types. yay
    VbIT_ComObj = &H1D   '  /

    VbIT_HRESULT = &H1E
    VbIT_LongLong = &HFF 'not possible in x86
#Else
    VbIT_Boolean = 3
    VbIT_Byte = 5
    VbIT_Integer = 6
    VbIT_Long = 8
    VbIT_Single = &HA
    VbIT_Double = &HB
    VbIT_Date = &HC
    VbIT_Currency = &HD
    VbIT_Variant = &HF
    VbIT_String = &H10
    VbIT_LongLong = &H11
    VbIT_UserDefinedType = &H15
    VbIT_EnumMaybe = &H19
    VbIT_Object = &H1C
    'Only tested above this comment - the rest are incremented by one from the x86 ones.
    VbIT_Internal = &H14
    VbIT_ComIface = &H1D
    VbIT_ComObj = &H1E
    VbIT_HRESULT = &H1F
#End If
End Enum

'From Greedquest - used with QueryInterface (link in comment near that function)
Private Type GUIDt
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Const iidITypeLib As String = "{00020402-0000-0000-C000-000000000046}"
Private Const CC_STDCALL As Long = 4
Private Const offset_ITypeLib_GetTypeInfo = PtrSize * 4
Private Const offset_ITypeInfo_GetTypeAttr = PtrSize * 3
Private Const offset_ITypeInfo_ReleaseTypeAttr = PtrSize * 19

'Public methods

Public Function GetCurrentProcedure() As StackFrame

    Dim pTopFrame As LongPtr 'This is a pointer to the top-of-stack ExFrame
    pTopFrame = ReadPtr(GetExFrameTOS())
    
    Dim topFrame As LongPtr 'This is the actual top frame - i.e., *this* currently executing procedure
    topFrame = ReadPtr(pTopFrame)
    
    Dim frameBefore As LongPtr 'This is the thing that *called* GetCurrentProcedure, the one we want.
    frameBefore = ReadPtr(topFrame)
    
    
    Dim retVal As StackFrame
    retVal = VBAStack.FrameFromPointer(frameBefore)
    retVal.FrameNumber = 1
    GetCurrentProcedure = retVal

    Exit Function
ErrorOccurred:
retVal.Errored = True
retVal.FrameNumber = -1
GetCurrentProcedure = retVal
End Function

Public Function GetProcedureBeforeCaller() As StackFrame

    Dim pTopFrame As LongPtr 'This is a pointer to the top-of-stack ExFrame
    pTopFrame = ReadPtr(GetExFrameTOS())
    
    Dim topFrame As LongPtr 'This is the actual top frame - i.e., *this* currently executing procedure
    topFrame = ReadPtr(pTopFrame)
    
    Dim callerFrame As LongPtr 'This is the procedure that called GetProcedureBeforeCaller
    callerFrame = ReadPtr(topFrame)
    If callerFrame = 0 Then GoTo ErrorOccurred
    
    Dim callerOfOurCallerFrame As LongPtr 'This is the procedure that led to our caller - the one we want
    callerOfOurCallerFrame = ReadPtr(callerFrame)
    
    
    Dim retVal As StackFrame
    retVal = VBAStack.FrameFromPointer(callerOfOurCallerFrame)
    retVal.FrameNumber = 2
    GetProcedureBeforeCaller = retVal
    
    Exit Function
ErrorOccurred:
retVal.Errored = True
retVal.FrameNumber = -1
GetProcedureBeforeCaller = retVal
End Function

'Returns an array of stackframes
Public Function GetCallstack() As StackFrame()

    Dim FrameArray() As StackFrame

    Dim pTopFrame As LongPtr 'This is a pointer to the top-of-stack ExFrame
    pTopFrame = ReadPtr(GetExFrameTOS())
    
    Dim topFrame As LongPtr 'This is the actual top frame - i.e., *this* currently executing procedure
    topFrame = ReadPtr(pTopFrame)
    
    Dim frameBefore As LongPtr 'This is the thing that *called* GetCallstack, the one we want.
    frameBefore = ReadPtr(topFrame)
    
    
    'Now walk the linked list of frames, building up the array.
    Dim workingFrame As LongPtr
    Dim count As Integer
    workingFrame = frameBefore
    
    
    Do Until workingFrame = 0
        
        ReDim Preserve FrameArray(count)
        FrameArray(count) = VBAStack.FrameFromPointer(workingFrame)
        FrameArray(count).FrameNumber = count + 1
        workingFrame = ReadPtr(workingFrame)
        count = count + 1
        
    Loop
    
    GetCallstack = FrameArray
End Function

Public Function GetParamInfoForFrame(frame As StackFrame) As paramInfo()

    'guess what confused the hell out of me and required me to change this variable name, lol
    Dim ArrayOfParams() As paramInfo

    If frame.Errored Or frame.FrameNumber < 0 Then GoTo ErrorOccurred

    'Get our frame's pointer out of the struct
    Dim pExFrame As LongPtr
    pExFrame = CurToLongPtr(frame.pExFrame)
    If pExFrame = 0 Then GoTo ErrorOccurred


    'Check it hasn't gone out of scope
    If IsFrameOnStack(pExFrame) = False Then GoTo ErrorOccurred


    'Get RTMI
    Dim pRTMI As LongPtr
    If Not CheckAddressSafe(pExFrame, PtrSize * 3) Then GoTo ErrorOccurred
    pRTMI = ReadPtr(pExFrame, PtrSize * 3)
    If pRTMI = 0 Then GoTo ErrorOccurred


    'Get ObjectInfo
    Dim pObjectInfo As LongPtr
    If Not CheckAddressSafe(pRTMI) Then GoTo ErrorOccurred
    pObjectInfo = ReadPtr(pRTMI)
    If pObjectInfo = 0 Then GoTo ErrorOccurred


    'Get ObjectTable
    Dim pObjectTable As LongPtr
    If Not CheckAddressSafe(pObjectInfo, PtrSize * 1) Then GoTo ErrorOccurred
    pObjectTable = ReadPtr(pObjectInfo, PtrSize * 1)
    If pObjectTable = 0 Then GoTo ErrorOccurred

    #If DEBUGConst Then
        MsgBox ("pObjectInfo:" & vbCrLf & DumpMemoryStr(pObjectInfo, 20))
    #End If


    'Attempt 1 to get PrivateObject
    Dim pPrivateObject As LongPtr
    If Not CheckAddressSafe(pObjectInfo, PtrSize * 3) Then GoTo ErrorOccurred
    pPrivateObject = ReadPtr(pObjectInfo, PtrSize * 3)


    'When working with a compiled Access MDE, the private object isn't supposed to be loaded. The LoadAndRelease_TypeAttrs sub causes it to load in.
    If pPrivateObject = 0 Or Not CheckAddressSafe(pPrivateObject) Then

        'Get ExecProj
        Dim pExecProj As LongPtr
        If Not CheckAddressSafe(pObjectTable, PtrSize * 1) Then GoTo ErrorOccurred
        pExecProj = ReadPtr(pObjectTable, PtrSize * 1)
        If pExecProj = 0 Then GoTo ErrorOccurred


        If LoadAndRelease_TypeAttrs(pExecProj) = False Then GoTo ErrorOccurred

        #If DEBUGConst Then
            MsgBox ("pObjectInfo:" & vbCrLf & DumpMemoryStr(pObjectInfo, 20))
        #End If

        'Attempt 2 to get PrivateObject
        pPrivateObject = ReadPtr(pObjectInfo, PtrSize * 3)

        'if that didn't work then we're fucked, so
        If pPrivateObject = 0 Or Not CheckAddressSafe(pPrivateObject) Then GoTo ErrorOccurred

    End If


    'Get array of function prototypes from PrivateObject
    Dim pFunctionPrototypeArr As LongPtr
    #If Win64 Then
        If Not CheckAddressSafe(pPrivateObject, &H28) Then GoTo ErrorOccurred
        pFunctionPrototypeArr = ReadPtr(pPrivateObject, &H28)
    #Else
        If Not CheckAddressSafe(pPrivateObject, &H18) Then GoTo ErrorOccurred
        pFunctionPrototypeArr = ReadPtr(pPrivateObject, &H18)
    #End If
    If pFunctionPrototypeArr = 0 Then GoTo ErrorOccurred


    'Get our prototype
    Dim pFunctionPrototype As LongPtr
    If Not CheckAddressSafe(pFunctionPrototypeArr, PtrSize * frame.MethodIndex) Then GoTo ErrorOccurred
    pFunctionPrototype = ReadPtr(pFunctionPrototypeArr, PtrSize * frame.MethodIndex)
    If pFunctionPrototype = 0 Then GoTo ErrorOccurred


    'Get pointer to array of parameters names from prototype
    Dim pArgNamesArr As LongPtr
    #If Win64 Then
        If Not CheckAddressSafe(pFunctionPrototype, &H18) Then GoTo ErrorOccurred
        pArgNamesArr = ReadPtr(pFunctionPrototype, &H18)
    #Else
        If Not CheckAddressSafe(pFunctionPrototype, &H10) Then GoTo ErrorOccurred
        pArgNamesArr = ReadPtr(pFunctionPrototype, &H10)
    #End If
    If pArgNamesArr = 0 Then GoTo ErrorOccurred
    
    
    'I discovered while porting this to x64 that Microsoft had to make the "type bytes" not bloody bytes any more - they're 9 bits long in x64. Sometimes I'll refer to them as type bytes, sometimes as type numbers.
    'The type bytes themselves are both an enum value, and 3 flag bits (isByRef, isOptional, isArray). All Microsoft did was left shift the 3 flag bits by 1, to give themselves another bit for the enum value.
    'So in x86, an Optional ByRef Integer shows up as this: 00000000-10100110
    'But in x64 that same parameter shows up as this:       00000001-01000110

    'At the end of the function prototype structure, there's a dynamically sized array of bytes - each one corresponds to a parameter of our function prototype, and defines it's type, ByRef/ByVal, if it's an array, etc.
    'That array starts at the address below, and ends wherever the array of parameter names (pArgNamesArr) starts. In x86 each type is 1 byte, x64 it's 2.
    'The first one is the return type of the function - we ignore this, but we do need to read it anyway since it could be one of the types that has an extra pointer attached, which we will need to skip over.
    Dim pTypeNumArr As LongPtr
    #If Win64 Then
        pTypeNumArr = pFunctionPrototype + &H34
    #Else
        pTypeNumArr = pFunctionPrototype + &H20
    #End If
    

    'Read that dynamic array (this is gonna get messy. apologies)
    ArrayOfParams = ParseTypes(pTypeNumArr, pArgNamesArr)

    
    #If DEBUGConst Then
        Debug.Print frame.ProcedureName
    #End If
    'god help us
    ReadParamValues pExFrame, ArrayOfParams

    GetParamInfoForFrame = ArrayOfParams
    Exit Function

ErrorOccurred:
Dim i As Integer
For i = 0 To UBound(ArrayOfParams)
    ArrayOfParams(i).Errored = True
Next i
GetParamInfoForFrame = ArrayOfParams
End Function

Private Function LoadAndRelease_TypeAttrs(pExecProj As LongPtr) As Boolean

    'Get ExecProj ITypeLib object
    Dim pITypeLib As LongPtr
    pITypeLib = QueryInterface(pExecProj, iidITypeLib)
    If pITypeLib = 0 Then GoTo ErrorOccurred
    
    
    'Call ITypeLib.GetTypeInfo to get ITypeInfo (hits GEN_PROJECT::GetTypeInfo(int,ITypeInfo * *) in VBE7.dll)
    Dim pITypeInfo As LongPtr, ppITypeInfo As Variant
    pITypeInfo = 0: ppITypeInfo = VarPtr(pITypeInfo)
    Dim index As Variant
    index = &H0
    Dim varResult As Variant: varResult = &H0
    
    Dim varTypes(1) As Integer, varValues(1) As LongPtr
    varTypes(0) = VarType(index)
    varTypes(1) = VarType(ppITypeInfo)
    varValues(0) = VarPtr(index)
    varValues(1) = VarPtr(ppITypeInfo)
    
    Dim dispCallFuncReturn As Long
    dispCallFuncReturn = DispCallFunc(pITypeLib, offset_ITypeLib_GetTypeInfo, CC_STDCALL, vbLong, 2, varTypes(0), varValues(0), varResult)
    If dispCallFuncReturn <> 0 Or varResult <> 0 Then GoTo ErrorOccurred
    
    
    'Call ITypeInfo.GetTypeAttr (hits BASIC_TYPEINFO::GetTypeAttr(tagTYPEATTR * *) in VBE7.dll)
    Dim pTypeAttr As LongPtr, ppTypeAttr As Variant
    pTypeAttr = 0: ppTypeAttr = VarPtr(pTypeAttr)
    
    varTypes(0) = VarType(ppTypeAttr)
    varValues(0) = VarPtr(ppTypeAttr)
    
    dispCallFuncReturn = DispCallFunc(pITypeInfo, offset_ITypeInfo_GetTypeAttr, CC_STDCALL, vbLong, 1, varTypes(0), varValues(0), varResult)
    If dispCallFuncReturn <> 0 Or varResult <> 0 Then GoTo ErrorOccurred
    
    
    'Call ITypeInfo.ReleaseTypeAttr (hits BASIC_TYPEINFO::ReleaseTypeAttr(tagTYPEATTR *) in VBE7.dll)
    Dim pTypeAttrVariant As Variant: pTypeAttrVariant = pTypeAttr
    varTypes(0) = VarType(pTypeAttrVariant)
    varValues(0) = VarPtr(pTypeAttrVariant)
    dispCallFuncReturn = DispCallFunc(pITypeInfo, offset_ITypeInfo_ReleaseTypeAttr, CC_STDCALL, vbEmpty, 1, varTypes(0), varValues(0), varResult)
    If dispCallFuncReturn <> 0 Then GoTo ErrorOccurred
    
    LoadAndRelease_TypeAttrs = True
    Exit Function
    
ErrorOccurred:
LoadAndRelease_TypeAttrs = False
End Function

'Private methods

Private Sub ReadParamValues(ByVal pExFrame As LongPtr, ByRef ArrayOfParams() As paramInfo)

    'Array isn't initialised, so leave
    If (Not ArrayOfParams) = -1 Then Exit Sub

    Dim pParamBase As LongPtr
    #If Win64 Then
        pParamBase = ReadPtr(pExFrame - &H38) 'Magic number time!
    #Else
        pParamBase = ReadPtr(pExFrame - &H28) 'Magic number time!
    #End If
    Dim i As Integer
    
    Dim totalArgSize As Integer
    For i = 0 To UBound(ArrayOfParams)
        totalArgSize = totalArgSize + ArrayOfParams(i).ParamSize
    Next i
    
    
    Dim curPtr As LongPtr
    #If Win64 Then
        curPtr = pParamBase
    #Else
        curPtr = pParamBase - totalArgSize
    #End If
    
    #If DEBUGConst Then
        DumpMemory curPtr, totalArgSize
    #End If
    
    For i = 0 To UBound(ArrayOfParams)
        
        Dim thisParam As paramInfo
        thisParam = ArrayOfParams(i)
        
        With thisParam
            Dim pParamVal As LongPtr
            If .IsByRef Then
                pParamVal = ReadPtr(curPtr)
            Else
                pParamVal = curPtr
            End If
                            
            'and here's the real nightmare
            .Value = ParamValAsString(pParamVal, thisParam)
            
            #If DEBUGConst Then
                Debug.Print .Value
            #End If
            
            curPtr = curPtr + .ParamSize
        End With
        
        ArrayOfParams(i) = thisParam
    Next i
    
End Sub

Private Function ParamValAsString(paramPtr As LongPtr, thisParam As paramInfo) As String
    Dim retVal As String
    
    If paramPtr = 0 Then
        If thisParam.TypeEnumVal <> VbInternal_Type.VbIT_Variant Then
            ParamValAsString = "[Nothing]"
        Else
            ParamValAsString = "[Empty]"
        End If
        Exit Function
    End If
    
    If Not CheckAddressSafe(paramPtr) Then GoTo ErrorOccurred

    Select Case thisParam.TypeEnumVal
    
        Case VbInternal_Type.VbIT_String
            Dim valString As String
            Dim bstrPtr As LongPtr 'I don't know why but for some reason this isn't a BSTR* like i thought, but a BSTR** - so read pointer first
            bstrPtr = ReadPtr(paramPtr)
            If bstrPtr <> 0 Then
                If Not CheckAddressSafe(bstrPtr) Then GoTo ErrorOccurred
                valString = ReadUniStr(bstrPtr)
                retVal = (valString)
            Else
                retVal = "[Null string]"
            End If

        Case VbInternal_Type.VbIT_Byte
            Dim valByte As Byte
            CopyMemory valByte, paramPtr, thisParam.DataSize
            retVal = valByte
            
        Case VbInternal_Type.VbIT_Boolean
            Dim valBoolean As Boolean
            CopyMemory valBoolean, paramPtr, thisParam.DataSize
            retVal = valBoolean

        Case VbInternal_Type.VbIT_Integer
            Dim valInteger As Integer
            CopyMemory valInteger, paramPtr, thisParam.DataSize
            retVal = valInteger

        Case VbInternal_Type.VbIT_Single
            Dim valSingle As Single
            CopyMemory valSingle, paramPtr, thisParam.DataSize
            retVal = valSingle

        Case VbInternal_Type.VbIT_Long
            Dim valLong As Long
            CopyMemory valLong, paramPtr, thisParam.DataSize
            retVal = valLong

        Case VbInternal_Type.VbIT_Double
            Dim valDouble As Double
            CopyMemory valDouble, paramPtr, thisParam.DataSize
            retVal = valDouble

         Case VbInternal_Type.VbIT_Date
            Dim valDate As Date
            CopyMemory valDate, paramPtr, thisParam.DataSize
            retVal = valDate
            
        Case VbInternal_Type.VbIT_Currency
            Dim valCurrency As Currency
            CopyMemory valCurrency, paramPtr, thisParam.DataSize
            retVal = valCurrency
            
        #If Win64 Then
            Case VbInternal_Type.VbIT_LongLong
                Dim valLongPtr As LongPtr
                CopyMemory valLongPtr, paramPtr, thisParam.DataSize
                retVal = IIf(valLongPtr <> 0, Hex(valLongPtr), "[Nothing]")
        #End If
        
        Case VbInternal_Type.VbIT_ComIface, VbInternal_Type.VbIT_Internal
            Dim valComObjPtr As LongPtr
            CopyMemory valComObjPtr, paramPtr, thisParam.DataSize
            retVal = IIf(valComObjPtr <> 0, "[ComObject]", "[Nothing]")
            
        Case VbInternal_Type.VbIT_Object, VbInternal_Type.VbIT_ComObj
            Dim valObj As Object
            'This is *VERY* dirty.
            CopyMemory valObj, paramPtr, thisParam.DataSize 'just briefly set valObj to the parameter value
            retVal = "[" & TypeName(valObj) & "]"
            If retVal <> "[Nothing]" Then thisParam.TypeName = thisParam.TypeName & "/" & TypeName(valObj)
            CopyMemory valObj, VarPtr(Nothing), thisParam.DataSize 'now set it back so VBA can't "clean it up" and screw with the COM reference count
            
        Case VbInternal_Type.VbIT_Variant
            Dim valVariant As Variant
            Dim copyresult As Long
            On Error Resume Next
            copyresult = VariantCopy(valVariant, paramPtr)
            thisParam.TypeName = thisParam.TypeName & "/" & TypeName(valVariant)
                
            retVal = valVariant
            If err.Number <> 0 Then
                retVal = "[Variant to string failed]"
            End If
            On Error GoTo 0

        #If DEBUGConst Then
            Case VbInternal_Type.VbIT_UserDefinedType
                MsgBox ("ParamPtr is: " & Hex(paramPtr))
        #End If
            
        Case Else
            'Not handled
            retVal = "[Unknown]"
            
    End Select
    
    ParamValAsString = retVal
    Exit Function
    
ErrorOccurred:
    ParamValAsString = "[Invalid parameter pointer]"
End Function

Private Function GetParamInfoForTypeNum(typeNum As Integer) As paramInfo

    ' 100 - Left flag bit: IsOptional
    ' 010 - Middle flag bit: IsArray
    ' 001 - Right flag bit: IsByRef

    Dim retVal As paramInfo
    
    Dim flagBits As Byte
    Dim enumValue As Byte
    #If Win64 Then
        'Get 3 most sig bits of a 9 bit value
        flagBits = LShift(typeNum, 6)
        'Bitwise AND the typenum with 00111111 to get the 6 least sig bits
        enumValue = typeNum And 63
    #Else
        'Get 3 most sig bits of a byte
        flagBits = LShift(typeNum, 5)
        'Bitwise AND the typenum with 00011111 to get the 5 least sig bits
        enumValue = typeNum And 31
    #End If
    
    
    'Get name of that type
    retVal.TypeName = VbInternal_Type_ToString(enumValue)
    If retVal.TypeName = "Unknown" Then Stop 'Very not good! If you get stopped here, tell Leo
    retVal.TypeEnumVal = enumValue
    
    
    'Left flag bit
    retVal.IsOptional = ((flagBits And 4) <> 0)
    'Middle flag bit
    retVal.IsArray = ((flagBits And 2) <> 0)
    'Right flag bit
    retVal.IsByRef = ((flagBits And 1) <> 0)
    
    If retVal.IsArray And Not retVal.IsByRef Then Debug.Print ("Type is somehow an array, but not ByRef - this should never happen!")
    
    
    If enumValue = VbInternal_Type.VbIT_ComIface Or enumValue = VbInternal_Type.VbIT_ComObj Or enumValue = VbInternal_Type.VbIT_Internal Then
        retVal.hasExtraPointer = True
    End If
    
    Select Case enumValue
        Case VbInternal_Type.VbIT_Byte
            retVal.DataSize = 1
        Case VbInternal_Type.VbIT_Integer, VbInternal_Type.VbIT_Boolean
            'Here's something fun - booleans are 2 bytes in size, because they are actually integers, they just have a value restriction on them.
            'Why not use a byte, and put a value restriction on that, you ask? In my (uninformed) opinion, it's because bytes are unsigned. And for some
            'ungodly reason they decided that true should be -1.
            retVal.DataSize = 2
        Case VbInternal_Type.VbIT_Single, VbInternal_Type.VbIT_Long
            retVal.DataSize = 4
        Case VbInternal_Type.VbIT_Double, VbInternal_Type.VbIT_Currency, VbInternal_Type.VbIT_LongLong, VbInternal_Type.VbIT_Date
            retVal.DataSize = 8
        Case VbInternal_Type.VbIT_Variant
            If PtrSize = 4 Then
                retVal.DataSize = 16
            Else
                retVal.DataSize = 24
            End If
        'All of the pointer types HAVE to be the size of a pointer, even if they're passed byVal. Right??
        Case VbInternal_Type.VbIT_Object, VbInternal_Type.VbIT_ComObj, VbInternal_Type.VbIT_ComIface, VbInternal_Type.VbIT_Internal
            retVal.DataSize = PtrSize
        
        Case VbInternal_Type.VbIT_String
            'purposefully do nothing here
    End Select

    
    If retVal.IsByRef Then
        retVal.ParamSize = PtrSize
    Else
        If PtrSize > retVal.DataSize Then
            retVal.ParamSize = PtrSize
        Else
            retVal.ParamSize = retVal.DataSize
        End If
    End If
    
    GetParamInfoForTypeNum = retVal
    
    
    
    #If DEBUGConst Then
        Dim hx As String
        hx = Hex(typeNum)
        If Len(hx) = 1 Then hx = hx & "  "
        If Len(hx) = 2 Then hx = hx & " "
        Debug.Print hx & " - " & DisplayByteAsBinary(typeNum)
    #End If
    Exit Function
    
ErrorOccurred:
    retVal.Errored = True
End Function

Private Function ParseTypes(pTypeNumsArr As LongPtr, pArgNamesArr As LongPtr) As paramInfo()

    Dim params() As paramInfo
    
    Dim typeNum As Integer
    Dim curParamInfo As paramInfo
    Dim curPtr As LongPtr
    curPtr = pTypeNumsArr
    
    
    'Read first type, if it exists, as it is the return value. If it's one of the types that has extra data, we need to skip that.
    If Not CheckAddressSafe(curPtr) Then GoTo ErrorOccurred
    CopyMemory typeNum, curPtr, (PtrSize / 4)
    curPtr = curPtr + (PtrSize / 4)
    If typeNum <> 0 Then
        curParamInfo = GetParamInfoForTypeNum(typeNum)
        If curParamInfo.hasExtraPointer Then
            curPtr = curPtr + IIf((curPtr Mod PtrSize) <> 0, (PtrSize - (curPtr Mod PtrSize)), 0) 'Skip alignment padding
            'Don't read it because we don't care
            'curParamInfo.pExtraData = LongPtrToCur(ReadPtr(curPtr))
            curPtr = curPtr + PtrSize
        End If
    End If
    
    
    'This needs to read 2 bytes per type number on x64, and 1 byte on x86.
    Do
        If curPtr = pArgNamesArr Then Exit Do
        If curPtr > pArgNamesArr Then Debug.Print ("Holy christ we've overshot the bounds of the dynamic array - must've started at the wrong address while reading 2-bytes per element in an x64 environment")
        If Not CheckAddressSafe(curPtr) Then GoTo ErrorOccurred
        CopyMemory typeNum, curPtr, (PtrSize / 4)
        curPtr = curPtr + (PtrSize / 4)
        
        If typeNum <> 0 Then
        
            If (Not params) = -1 Then 'If it has no upper bound yet, then redim it to 0
                ReDim Preserve params(0)
            Else 'Otherwise redim it to current bounds + 1
                ReDim Preserve params(UBound(params) + 1)
            End If
            
            'Parse the type num - we need to do this now as the type could be one of a few specific ones that just dumps a pointer into the middle of the damn array we're trying to read here
            curParamInfo = GetParamInfoForTypeNum(typeNum)
            
            'If it does, store it in case I decide to parse those later - NOTE: stdole.StdPicture triggers this behaviour
            If curParamInfo.hasExtraPointer Then
                curPtr = curPtr + IIf((curPtr Mod PtrSize) <> 0, (PtrSize - (curPtr Mod PtrSize)), 0) 'Skip alignment padding
                curParamInfo.pExtraData = LongPtrToCur(ReadPtr(curPtr))
                curPtr = curPtr + PtrSize
            End If
            
            'Get the param's name
            Dim pArgName As LongPtr
            'This array grows with every loop, so the current element is always the top bound (note mostly for myself because I keep looking at the use of UBound for the index into the names array and getting confused)
            If Not CheckAddressSafe(pArgNamesArr, PtrSize * UBound(params)) Then GoTo ErrorOccurred
            pArgName = ReadPtr(pArgNamesArr, PtrSize * UBound(params))
            
            If Not CheckAddressSafe(pArgName) Then GoTo ErrorOccurred
            curParamInfo.ParamName = ReadANSIStr(pArgName)
            
            params(UBound(params)) = curParamInfo
        End If
    Loop
    
    ParseTypes = params
    Exit Function
ErrorOccurred:

End Function

Private Function GetExFrameTOS() As LongPtr

    'Get ptr to VBA.Err
    Dim errObj As LongPtr
    errObj = ObjPtr(VBA.err)
    
    'Get g_ebThread
    Dim g_ebThread As LongPtr
    g_ebThread = ReadPtr(errObj, PtrSize * 6)
    If g_ebThread = 0 Then GoTo ErrorOccurred
    
    
    'Get g_ExFrameTOS
    #If Win64 Then
        GetExFrameTOS = g_ebThread + (&H10)
    #Else
        GetExFrameTOS = g_ebThread + (&HC)
    #End If
    Exit Function
    
ErrorOccurred:
GetExFrameTOS = 0
End Function

Private Function IsFrameOnStack(pExFrame As LongPtr) As Boolean

    Dim pTmpFrame As LongPtr

    IsFrameOnStack = False
    pTmpFrame = GetExFrameTOS()

    
    Do While pTmpFrame <> 0
        
        If pTmpFrame = pExFrame Then 'Found it - return true
            IsFrameOnStack = True
            Exit Function
        Else 'Try next one down the stack
            pTmpFrame = ReadPtr(pTmpFrame)
        End If
        
    Loop

End Function

'Takes a pointer to a VBA ExFrame and makes one of my StackFrame structs
Private Function FrameFromPointer(pExFrame As LongPtr) As StackFrame

On Error GoTo ErrorOccurred
    
    Dim retVal As StackFrame
    If IsFrameOnStack(pExFrame) = False Then GoTo ErrorOccurred
    
    retVal.pExFrame = LongPtrToCur(pExFrame)
    
    'Get RTMI
    Dim pRTMI As LongPtr
    pRTMI = ReadPtr(pExFrame, PtrSize * 3)
    If pRTMI = 0 Then GoTo ErrorOccurred
    
    
    'Get ObjectInfo
    Dim pObjectInfo As LongPtr
    pObjectInfo = ReadPtr(pRTMI)
    If pObjectInfo = 0 Then GoTo ErrorOccurred
    
    
    'Get Public Object Descriptor
    Dim pPublicObject As LongPtr
    pPublicObject = ReadPtr(pObjectInfo, PtrSize * 6)
    If pPublicObject = 0 Then GoTo ErrorOccurred
    
    
    'Get pointer to module name string from Public Object Descriptor
    Dim pObjectName As LongPtr
    pObjectName = ReadPtr(pPublicObject, PtrSize * 6)
    If pObjectName = 0 Then GoTo ErrorOccurred
    
    
    'Read the object name string
    retVal.ObjectName = ReadANSIStr(pObjectName)
    
    
    'Get pointer to methods array from ObjectInfo
    Dim pMethodsArr As LongPtr
    pMethodsArr = ReadPtr(pObjectInfo, PtrSize * 9)
    If pMethodsArr = 0 Then GoTo ErrorOccurred
    
    
    'Get count of methods from Public Object Descriptor
    Dim methodCount As Long
    methodCount = Read4Byte(pPublicObject, PtrSize * 7)
    If methodCount = 0 Then GoTo ErrorOccurred '...I don't think anything can have 0 methods and still actually HAVE a stackframe? Considering this seems to count all methods, public, private, whatever.
    
    
    'Search the method array to find our RTMI
    Dim iMethodIndex As Integer: iMethodIndex = -1
    Dim i As Integer
    Dim pMethodRTMI As LongPtr
    For i = methodCount - 1 To 0 Step -1
        pMethodRTMI = ReadPtr(pMethodsArr, PtrSize * i)
        If pMethodRTMI = 0 Then GoTo ErrorOccurred
        If pMethodRTMI = pRTMI Then
            iMethodIndex = i
            Exit For
        End If
    Next
    
    If iMethodIndex = -1 Then GoTo ErrorOccurred
    retVal.MethodIndex = iMethodIndex
    
    'Get array of method names from Public Object Descriptor
    Dim pMethodNamesArr As LongPtr
    pMethodNamesArr = ReadPtr(pPublicObject, PtrSize * 8)
    If pMethodNamesArr = 0 Then GoTo ErrorOccurred
    
    
    'Get pointer to our method name
    Dim pMethodName As LongPtr
    pMethodName = ReadPtr(pMethodNamesArr + PtrSize * iMethodIndex)
    If pMethodName = 0 Then GoTo ErrorOccurred
    
    
    'Read the method name string
    retVal.ProcedureName = ReadANSIStr(pMethodName)
    
    
    'Get ObjectTable
    Dim pObjectTable As LongPtr
    pObjectTable = ReadPtr(pObjectInfo, PtrSize * 1)
    If pObjectTable = 0 Then GoTo ErrorOccurred
    
    
    'Get project name from ObjectTable
    Dim pProjName As LongPtr
    #If Win64 Then
        pProjName = ReadPtr(pObjectTable, &H68)
    #Else
        pProjName = ReadPtr(pObjectTable, &H40)
    #End If
    If pProjName = 0 Then GoTo ErrorOccurred
    
    
    'Read the project name string
    retVal.ProjectName = ReadANSIStr(pProjName)
    FrameFromPointer = retVal
    
Exit Function

ErrorOccurred:
    retVal.Errored = True
    retVal.FrameNumber = -1
    FrameFromPointer = retVal
End Function


'Helpers of every flavour

Public Function ParamInfoToString(paramInfo As paramInfo) As String
    ParamInfoToString = IIf(paramInfo.IsByRef, "ByRef ", "ByVal ") & paramInfo.ParamName & " As " & paramInfo.TypeName & " : " & paramInfo.Value
End Function

Public Function ParamInfoArrayToString(paramInfo() As paramInfo) As String
    Dim i As Integer
    If (Not paramInfo) = -1 Then Exit Function
    For i = 0 To UBound(paramInfo)
        If ParamInfoArrayToString <> vbNullString Then
            ParamInfoArrayToString = ParamInfoArrayToString & vbCrLf
        End If
        ParamInfoArrayToString = ParamInfoArrayToString & ParamInfoToString(paramInfo(i))
    Next
End Function

Public Function StackFrameToString(frame As StackFrame, Optional IncludeProject As Boolean = False) As String
    StackFrameToString = IIf(IncludeProject, frame.ProjectName & "::", "") & frame.ObjectName & "::" & frame.ProcedureName
End Function

Public Function StackFramesToString(frame() As StackFrame, Optional IncludeProject As Boolean = False) As String
    Dim i As Integer
    If (Not frame) = -1 Then Exit Function
    For i = 0 To UBound(frame)
        If StackFramesToString <> vbNullString Then
            StackFramesToString = StackFramesToString & vbCrLf
        End If
        StackFramesToString = StackFramesToString & StackFrameToString(frame(i), IncludeProject)
    Next
End Function

Private Function ReadPtr(ByVal lpSource As LongPtr, Optional ByVal offset As LongPtr = 0) As LongPtr
    CopyMemory ReadPtr, lpSource + offset, PtrSize
End Function

Private Function Read4Byte(ByVal lpSource As LongPtr, Optional ByVal offset As LongPtr = 0) As Long
    CopyMemory Read4Byte, lpSource + offset, 4
End Function

Private Function Read2Byte(ByVal lpSource As LongPtr, Optional ByVal offset As LongPtr = 0) As Integer
    CopyMemory Read2Byte, lpSource + offset, 2
End Function

Private Function LongPtrToCur(ByVal ptr As LongPtr) As Currency
    CopyMemory LongPtrToCur, VarPtr(ptr), PtrSize
End Function

Private Function CurToLongPtr(ByVal cur As Currency) As LongPtr
    CopyMemory CurToLongPtr, VarPtr(cur), PtrSize
End Function

Private Function CheckAddressSafe(ByVal pAddr As LongPtr, Optional ByVal offset As LongPtr = 0, Optional ByVal Name As String = "") As Boolean
    
    Dim arrayIndex As Integer
    
    pAddr = pAddr + offset
    
    'Does the cache array exist?
    If (Not SafeAddressCache) = -1 Then
        'No - create it
        ReDim SafeAddressCache(0)
        arrayIndex = 0
    Else
        'Cache exists
        For arrayIndex = 0 To UBound(SafeAddressCache)
            'Check if it's within the memory range
            If pAddr > SafeAddressCache(arrayIndex).pRangeStart And pAddr < SafeAddressCache(arrayIndex).pRangeEnd Then
                CheckAddressSafe = SafeAddressCache(arrayIndex).Safe
                Exit Function
            End If
        Next arrayIndex
        
        'The address wasn't in any of our previously checked ranges, so enlarge the array by one.
        arrayIndex = UBound(SafeAddressCache) + 1
        ReDim Preserve SafeAddressCache(arrayIndex)
    End If

    'If we got down here, then the address wasn't in any of the cached memory ranges, and arrayIndex will point us to a new, empty cache entry. So fill it!
    
    If Name <> vbNullString Then Debug.Print Name
    SafeAddressCache(arrayIndex) = GetAddrRangeSafety(pAddr)
    If Name <> vbNullString Then Debug.Print vbCrLf
    #If DEBUGConst Then
        If Not SafeAddressCache(arrayIndex).Safe Then Stop 'Christ almighty if you get stopped here, tell Leo
    #End If
    
    CheckAddressSafe = SafeAddressCache(arrayIndex).Safe
End Function

Private Function GetAddrRangeSafety(pAddr As LongPtr) As AddressRangeSafety
    Dim MBI As MEMORY_BASIC_INFORMATION
    Dim vqret As LongPtr
    
    'From MS documentation for VirtualQuery,
        'The return value is the actual number of bytes returned in the information buffer.
        'If the function fails, the return value is zero.
    vqret = VirtualQuery(pAddr, MBI, LenB(MBI))
    
    GetAddrRangeSafety.pRangeStart = MBI.BaseAddress
    GetAddrRangeSafety.pRangeEnd = MBI.BaseAddress + MBI.RegionSize
    
    If err.LastDLLError <> 0 Or vqret = 0 Then
        GetAddrRangeSafety.Safe = False
        Exit Function
    End If
    
    If MBI.Protect = PAGE_EXECUTE_READWRITE Or MBI.Protect = PAGE_READWRITE Then
        GetAddrRangeSafety.Safe = True
        Exit Function
    End If
    
    GetAddrRangeSafety.MBI = MBI

End Function

Private Function RShift(ByVal lNum As Long, ByVal lBits As Long) As Long
    If lBits <= 0 Then RShift = lNum
    If (lBits <= 0) Or (lBits > 31) Then Exit Function
    
    RShift = (lNum And (2 ^ (31 - lBits) - 1)) * _
        IIf(lBits = 31, &H80000000, 2 ^ lBits) Or _
        IIf((lNum And 2 ^ (31 - lBits)) = 2 ^ (31 - lBits), _
        &H80000000, 0)
End Function

Private Function LShift(ByVal lNum As Long, ByVal lBits As Long) As Long
    If lBits <= 0 Then LShift = lNum
    If (lBits <= 0) Or (lBits > 31) Then Exit Function
    
    If lNum < 0 Then
        LShift = (lNum And &H7FFFFFFF) \ (2 ^ lBits) Or 2 ^ (31 - lBits)
    Else
        LShift = lNum \ (2 ^ lBits)
    End If
End Function

Private Function ReadUniStr(ByVal lpStr As LongPtr) As String
    SysReAllocString VarPtr(ReadUniStr), lpStr
End Function

Private Function ReadANSIStr(ByVal lpStr As LongPtr) As String
    Dim tmpByte As Byte
    
    Do
        CopyMemory tmpByte, lpStr, 1
        lpStr = lpStr + 1
        If tmpByte = 0 Then Exit Do
        ReadANSIStr = ReadANSIStr & Chr(tmpByte)
    Loop
End Function

Private Function ReadByte(ptr As LongPtr) As Byte
    CopyMemory ReadByte, ptr, 1
End Function

'well this sucked
Private Function VbInternal_Type_ToString(ByVal val As VbInternal_Type) As String
    Select Case val
        Case VbInternal_Type.VbIT_Boolean
            VbInternal_Type_ToString = "Boolean"
        Case VbInternal_Type.VbIT_Byte
            VbInternal_Type_ToString = "Byte"
        Case VbInternal_Type.VbIT_ComIface
            VbInternal_Type_ToString = "ComIface"
        Case VbInternal_Type.VbIT_ComObj
            VbInternal_Type_ToString = "ComObj"
        Case VbInternal_Type.VbIT_Currency
            VbInternal_Type_ToString = "Currency"
        Case VbInternal_Type.VbIT_Date
            VbInternal_Type_ToString = "Date"
        Case VbInternal_Type.VbIT_Double
            VbInternal_Type_ToString = "Double"
        Case VbInternal_Type.VbIT_HRESULT
            VbInternal_Type_ToString = "HRESULT"
        Case VbInternal_Type.VbIT_Integer
            VbInternal_Type_ToString = "Integer"
        Case VbInternal_Type.VbIT_Internal
            VbInternal_Type_ToString = "Internal"
        Case VbInternal_Type.VbIT_Long
            VbInternal_Type_ToString = "Long"
        Case VbInternal_Type.VbIT_LongLong
            VbInternal_Type_ToString = "LongLong"
        Case VbInternal_Type.VbIT_Object
            VbInternal_Type_ToString = "Object"
        Case VbInternal_Type.VbIT_Single
            VbInternal_Type_ToString = "Single"
        Case VbInternal_Type.VbIT_String
            VbInternal_Type_ToString = "String"
        Case VbInternal_Type.VbIT_UserDefinedType
            VbInternal_Type_ToString = "UserDefinedType"
        Case VbInternal_Type.VbIT_Variant
            VbInternal_Type_ToString = "Variant"
        Case Else
            VbInternal_Type_ToString = "Unknown"
    End Select
End Function


'Debugging helpers

Private Sub DumpMemory(ptr As LongPtr, Optional Length As Integer = 128)
    Debug.Print DumpMemoryStr(ptr, Length)
End Sub

Private Function DumpMemoryStr(ptr As LongPtr, Optional Length As Integer = 128) As String
    
    Dim dmp As Byte
    Dim dmpStr As String
    Dim z As LongPtr
    For z = 0 To Length - 1
        dmp = ReadByte(ptr + z)
        If z Mod PtrSize = 0 Then
            dmpStr = dmpStr & vbCrLf & Hex(ptr + z) & ": " & IIf(Len(Hex(dmp)) = 1, "0" & Hex(dmp), Hex(dmp))
        Else
            dmpStr = dmpStr & " " & IIf(Len(Hex(dmp)) = 1, "0" & Hex(dmp), Hex(dmp))
        End If
    Next z
    DumpMemoryStr = dmpStr
    
End Function

'I know this takes an integer not a byte, but Microsoft forced my hand by making the type enum for parameters 9 bits in x64. So here we are.
Private Function DisplayByteAsBinary(b As Integer) As String
Dim i As Integer

For i = 15 To 0 Step -1
    
    If i = 7 Then DisplayByteAsBinary = DisplayByteAsBinary & "-"
    
    If (b And (2 ^ i)) <> 0 Then
        DisplayByteAsBinary = DisplayByteAsBinary & "1"
    Else
        DisplayByteAsBinary = DisplayByteAsBinary & "0"
    End If

Next

End Function

'These were used earlier in development but don't seem necessary any more - keeping just in case

''For VBA6 compat
'#If VBA7 = False Then
'    Public Function CLngPtr(var) As LongPtr
'        CLngPtr = CLng(var)
'    End Function
'#End If
'
''This will return the number of frames BELOW the one you give it.
'Public Function FrameCountBelowFramePtr(pTargetFrame As LongPtr) As Integer
'
'On Error GoTo ErrorOccurred
'
'    FrameCountBelowFramePtr = 0
'    If IsFrameOnStack(pTargetFrame) = False Then GoTo ErrorOccurred
'
'
'    'Loop over frames to count
'    Dim pExFrame As LongPtr: pExFrame = pTargetFrame
'
'    Do While pExFrame <> 0
'        pExFrame = ReadPtr(pExFrame)
'        FrameCountBelowFramePtr = FrameCountBelowFramePtr + 1
'    Loop
'
'Exit Function
'
'ErrorOccurred:
'
'End Function
'
'Public Function FrameCountBelowFrame(frame As StackFrame) As Integer
'    FrameCountBelowFrame = FrameCountBelowFramePtr(CurToLongPtr(frame.pExFrame))
'End Function
'
'
'Public Function FrameCount() As Integer
'
'On Error GoTo ErrorOccurred
'
'    FrameCount = -1
'
'
'    'Get top ExFrame
'    Dim pTopExFrame As LongPtr
'    pTopExFrame = ReadPtr(GetExFrameTOS)
'    If pTopExFrame = 0 Then GoTo ErrorOccurred
'
'
'    'Loop over frames to count
'    Dim pExFrame As LongPtr: pExFrame = pTopExFrame
'    Do
'        CopyMemory pExFrame, pExFrame, PtrSize
'        FrameCount = FrameCount + 1
'        If pExFrame = 0 Then Exit Do
'    Loop
'
'ErrorOccurred:
'
'End Function
'
'Private Function GetCurrentFramePointer() As LongPtr
'
'    Dim pTopFrame As LongPtr 'This is a pointer to the top-of-stack ExFrame
'    pTopFrame = ReadPtr(GetExFrameTOS())
'    If pTopFrame = 0 Then GoTo ErrorOccurred
'
'    Dim topFrame As LongPtr 'This is the actual top frame - i.e., *this* currently executing procedure
'    topFrame = ReadPtr(pTopFrame)
'    If topFrame = 0 Then GoTo ErrorOccurred
'
'    Dim frameBefore As LongPtr 'This is the thing that *called* GetCurrentFramePointer, the one we want.
'    frameBefore = ReadPtr(topFrame)
'
'    GetCurrentFramePointer = frameBefore
'
'ErrorOccurred:
'GetCurrentFramePointer = 0
'End Function


'InterfaceQuerying.bas - Pilfered from Greedquest (https://web.archive.org/web/20260209204217/https://raw.githubusercontent.com/Greedquest/CodeReviewFiles/refs/heads/master/VBAHack/COMToolsSrc/InterfaceQuerying.bas)

Private Function QueryInterface(ByVal pClassInstance As LongPtr, ByVal InterfaceIID As String) As LongPtr

    Dim InterfaceGUID As GUIDt
    IIDFromString StrPtr(InterfaceIID), InterfaceGUID

    Dim valueWrapper0 As Variant
    Dim valueWrapper1 As Variant

    valueWrapper0 = VarPtr(InterfaceGUID)

    Dim retVal As LongPtr
    valueWrapper1 = VarPtr(retVal)

    Dim ptrVarValues(1) As LongPtr
    ptrVarValues(0) = VarPtr(valueWrapper0)
    ptrVarValues(1) = VarPtr(valueWrapper1)
    

    Dim varTypes(1) As Integer
    varTypes(0) = VbVarType.vbLong
    varTypes(1) = VarType(retVal)
    
    Const paramCount As Long = 2
    
    Dim apiRetVal As Variant
    Dim hResult As Long

    hResult = DispCallFunc(pClassInstance, 0, CC_STDCALL, VbVarType.vbLong, paramCount, varTypes(0), ptrVarValues(0), apiRetVal)

    If hResult = 0 Then
        hResult = apiRetVal
        
        If hResult = 0 Then
        
            QueryInterface = retVal
        Else
            err.Raise hResult, "QueryInterface", "Failed to cast to interface pointer. IUnknown::QueryInterface HRESULT: 0x" & Hex$(hResult)
        End If
    Else
        err.Raise hResult, "DispCallFunc", "Failed to cast to interface pointer. DispCallFunc HRESULT: 0x" & Hex$(hResult)
    End If
        
End Function


