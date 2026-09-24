Option Explicit On
Option Strict Off

Imports System.Runtime.InteropServices
Imports System.Threading

' IOleMessageFilter: kad je CATIA zauzeta, vanjski COM poziv vraca RPC_E_CALL_REJECTED
' ("Call was rejected by callee"). Ovaj filter automatski ponavlja odbijeni poziv
' umjesto da aplikacija pukne. Registrira se na STA niti (Application.Run / Form).
Friend Class OleMessageFilter
    Implements IOleMessageFilter

    Private Shared _registered As IOleMessageFilter

    ' Pozovi jednom na UI (STA) niti.
    Public Shared Sub Register()
        If _registered IsNot Nothing Then Return
        Dim newFilter As IOleMessageFilter = New OleMessageFilter()
        Dim oldFilter As IOleMessageFilter = Nothing
        CoRegisterMessageFilter(newFilter, oldFilter)
        _registered = newFilter
    End Sub

    Public Shared Sub Revoke()
        Dim oldFilter As IOleMessageFilter = Nothing
        CoRegisterMessageFilter(Nothing, oldFilter)
        _registered = Nothing
    End Sub

    ' Dolazni pozivi: uvijek prihvati (0 = SERVERCALL_ISHANDLED).
    Private Function HandleInComingCall(dwCallType As Integer, hTaskCaller As IntPtr,
                                        dwTickCount As Integer, lpInterfaceInfo As IntPtr) As Integer _
                                        Implements IOleMessageFilter.HandleInComingCall
        Return 0
    End Function

    ' Odbijeni poziv: ako je server zauzet (SERVERCALL_RETRYLATER = 2), ponovi.
    Private Function RetryRejectedCall(hTaskCallee As IntPtr, dwTickCount As Integer,
                                       dwRejectType As Integer) As Integer _
                                       Implements IOleMessageFilter.RetryRejectedCall
        If dwRejectType = 2 Then
            ' kratko pricekaj pa ponovi (vrijednost < 100 = retry nakon N ms)
            Thread.Sleep(50)
            Return 99
        End If
        ' -1 = odustani (poziv ce vratiti gresku pozivatelju)
        Return -1
    End Function

    Private Function MessagePending(hTaskCallee As IntPtr, dwTickCount As Integer,
                                    dwPendingType As Integer) As Integer _
                                    Implements IOleMessageFilter.MessagePending
        ' 2 = PENDINGMSG_WAITDEFPROCESS
        Return 2
    End Function

    <DllImport("ole32.dll")>
    Private Shared Sub CoRegisterMessageFilter(newFilter As IOleMessageFilter,
                                               ByRef oldFilter As IOleMessageFilter)
    End Sub

End Class

<ComImport(), Guid("00000016-0000-0000-C000-000000000046"),
 InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Friend Interface IOleMessageFilter
    <PreserveSig()>
    Function HandleInComingCall(dwCallType As Integer, hTaskCaller As IntPtr,
                                dwTickCount As Integer, lpInterfaceInfo As IntPtr) As Integer
    <PreserveSig()>
    Function RetryRejectedCall(hTaskCallee As IntPtr, dwTickCount As Integer,
                              dwRejectType As Integer) As Integer
    <PreserveSig()>
    Function MessagePending(hTaskCallee As IntPtr, dwTickCount As Integer,
                            dwPendingType As Integer) As Integer
End Interface
