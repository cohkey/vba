VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' derived from  https://stackoverflow.com/questions/62172551/error-with-createpipe-in-vba-office-64bit



Private hStdOutWr As LongPtr
Private hStdOutRd As LongPtr
Private hStdInWr As LongPtr
Private hStdInRd As LongPtr
Private hCDPOutWr As LongPtr
Private hCDPOutRd As LongPtr
Private hCDPInWr As LongPtr
Private hCDPInRd As LongPtr
Private boolSerializable As Boolean

Private hProcess As LongPtr

Public Function serialize() As String

    If Not boolSerializable Then
        Call err.Raise(-904, , "this instance is not serializable")
    End If

    Dim objSerialize As Dictionary
    Set objSerialize = New Dictionary
    Call objSerialize.Add("hStdOutRd", hStdOutRd)
    Call objSerialize.Add("hStdInWr", hStdInWr)
    Call objSerialize.Add("hCDPOutRd", hCDPOutRd)
    Call objSerialize.Add("hCDPInWr", hCDPInWr)
    serialize = ConvertToJson(objSerialize)


End Function

Public Sub deserialize(strSerialized As String)
    Dim objSerialize As Dictionary
    Set objSerialize = ParseJson(strSerialized)
    
    hStdOutRd = objSerialize.Item("hStdOutRd")
    hStdInWr = objSerialize.Item("hStdInWr")
    hCDPOutRd = objSerialize.Item("hCDPOutRd")
    hCDPInWr = objSerialize.Item("hCDPInWr")

End Sub

Public Function init(strExec As String, Optional aboolSerializable As Boolean = False) As Integer
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim sa As SECURITY_ATTRIBUTES
    Dim hReadPipe As LongPtr, hWritePipe As LongPtr
    Dim L As Long, result As Long, bSuccess As Long
    Dim Buffer As String
    Dim k As Long
    
    Dim pipes As STDIO_BUFFER
    Dim pipes2 As STDIO_BUFFER2
    
    boolSerializable = aboolSerializable


    ' First we create all 4 pipes
    
    ' We start with stdout of the edge process
    ' This pipe is used for stderr, too
    
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&
    result = CreatePipe(hStdOutRd, hStdOutWr, sa, 0)
    
    If result = 0 Then
        init = -2
        Exit Function
    End If
    
    ' then stdin
    
    result = CreatePipe(hStdInRd, hStdInWr, sa, 0)


    If result = 0 Then
        init = -2
        Exit Function
    End If
    
    ' then the out pipe for the CDP Protocol
    
    result = CreatePipe(hCDPOutRd, hCDPOutWr, sa, 2 ^ 20)
    
    If result = 0 Then
        init = -2
        Exit Function
    End If
    
    ' and finally the in pipe

    
    result = CreatePipe(hCDPInRd, hCDPInWr, sa, 0)


    If result = 0 Then
        init = -2
        Exit Function
    End If
    
    ' then we fill the special structure for passing arbitrary pipes (i.e. fds)
    ' to a process
    
    pipes.number_of_fds = 5
    
    pipes.os_handle(0) = hStdInRd
    pipes.os_handle(1) = hStdOutWr
    pipes.os_handle(2) = hStdOutWr
    pipes.os_handle(3) = hCDPInRd
    pipes.os_handle(4) = hCDPOutWr
    
    pipes.crt_flags(0) = 9
    pipes.crt_flags(1) = 9
    pipes.crt_flags(2) = 9
    pipes.crt_flags(3) = 9
    pipes.crt_flags(4) = 9
    
    ' pipes2 is filled by copying memory from pipes
    
    pipes2.number_of_fds = pipes.number_of_fds
    
    Call MoveMemory(pipes2.raw_bytes(0), pipes.crt_flags(0), 5)
    Call MoveMemory(pipes2.raw_bytes(5), pipes.os_handle(0), UBound(pipes2.raw_bytes) - 4)


    With start
        .cb = Len(start)
        .dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
        .hStdOutput = hStdOutWr
        .hStdInput = hStdInRd
        .hStdError = hStdOutWr
        .wShowWindow = vbHide ' hide console window, seems not to work
        .cbReserved2 = Len(pipes2)
        .lpReserved2 = VarPtr(pipes2)
    End With
    

    result = CreateProcessA(0&, strExec, sa, sa, 1&, _
        NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)

    If result = 0 Then
        init = -1
    End If
    
    ' We close the sides of the handles that we dont need anymore
    
    Call CloseHandle(hStdOutWr)
    Call CloseHandle(hStdInRd)
    Call CloseHandle(hCDPOutWr)
    Call CloseHandle(hCDPInRd)
    
'    Call Sleep(10)
'
'    EnumThreadWindows proc.dwThreadId, AddressOf EnumThreadWndProc, 0

    hProcess = proc.hProcess
    
    init = 0
    
End Function

' This function tries to read from the CDP out pipe
' Reading is non-blocking, if there are no bytes ro read the function returns 0
' otherwise the number of bytes read
Public Function readProcCDP(ByRef strData As String) As Long
    Dim lPeekData As Long
    
    Dim lngRes As Long
    
    lngRes = PeekNamedPipe(hCDPOutRd, ByVal 0&, 0&, ByVal 0&, _
        lPeekData, ByVal 0&)
        
    If lngRes = 0 Then
        Call err.Raise(901, , "Error PeekNamedPipe in readProcCDP")
    End If
    
    
    If lPeekData > 0 Then
        Dim Buffer As String
        Dim L As Long
        Dim bSuccess As Long
        Buffer = Space$(lPeekData)
        bSuccess = ReadFile(hCDPOutRd, Buffer, Len(Buffer), L, 0&)
        If bSuccess = 1 Then
            strData = Buffer
            
            readProcCDP = Len(strData)
        Else
            readProcCDP = -2
        End If
    Else
        readProcCDP = -1
    End If

End Function

' Same as ReadProcCDP

Public Function readProcSTD(ByRef strData As String) As Integer
    Dim lPeekData As Long
    
    Call PeekNamedPipe(hStdOutRd, ByVal 0&, 0&, ByVal 0&, _
    lPeekData, ByVal 0&)
    
    
    If lPeekData > 0 Then
        Dim Buffer As String
        Dim L As Long
        Dim bSuccess As Long
        Buffer = Space$(lPeekData)
        bSuccess = ReadFile(hStdOutRd, Buffer, Len(Buffer), L, 0&)
        If bSuccess = 1 Then
            strData = Buffer
            readProcSTD = Len(strData)
        Else
            readProcSTD = -2
        End If
    Else
        readProcSTD = -1
    End If

End Function

' This functions sends a CDP message to edge

Public Function writeProc(ByVal strData As String) As Integer
    Dim lngWritten As Long
    writeProc = WriteFile(hCDPInWr, ByVal strData, Len(strData), lngWritten, ByVal 0&)

End Function

Private Sub Class_Terminate()

    If boolSerializable Then Exit Sub

    Call CloseHandle(hStdOutRd)
    Call CloseHandle(hStdOutWr)
    Call CloseHandle(hStdInRd)
    Call CloseHandle(hStdInWr)

    Call CloseHandle(hCDPOutRd)
    Call CloseHandle(hCDPOutWr)
    Call CloseHandle(hCDPInRd)
    Call CloseHandle(hCDPInWr)
    
End Sub
