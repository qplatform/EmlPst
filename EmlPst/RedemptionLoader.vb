
Imports System.Collections.Generic
Imports System.Text
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports Redemption


Namespace Redemption

    Public NotInheritable Class RedemptionLoader

#Region "public methods"

        '64 bit dll location - defaults to <assemblydir>\Redemption64.dll
        Public Shared DllLocation64Bit As String
        '32 bit dll location - defaults to <assemblydir>\Redemption.dll
        Public Shared DllLocation32Bit As String
        'The only creatable RDO object - RDOSession
        Public Shared Function new_RDOSession() As RDOSession
            Return DirectCast(NewRedemptionObject(New Guid("29AB7A12-B531-450E-8F7A-EA94C2F3C05F")), Object)
        End Function
        'Safe*Item objects
        Public Shared Function new_SafeMailItem() As SafeMailItem
            Return DirectCast(NewRedemptionObject(New Guid("{741BEEFD-AEC0-4AFF-84AF-4F61D15F5526}")), Object)
        End Function
        Public Shared Function new_SafeContactItem() As SafeContactItem
            Return DirectCast(NewRedemptionObject(New Guid("4FD5C4D3-6C15-4EA0-9EB9-EEE8FC74A91B")), Object)
        End Function
        Public Shared Function new_SafeAppointmentItem() As SafeAppointmentItem
            Return DirectCast(NewRedemptionObject(New Guid("620D55B0-F2FB-464E-A278-B4308DB1DB2B")), Object)
        End Function
        Public Shared Function new_SafeTaskItem() As SafeTaskItem
            Return DirectCast(NewRedemptionObject(New Guid("7A41359E-0407-470F-B3F7-7C6A0F7C449A")), Object)
        End Function
        Public Shared Function new_SafeJournalItem() As SafeJournalItem
            Return DirectCast(NewRedemptionObject(New Guid("C5AA36A1-8BD1-47E0-90F8-47E7239C6EA1")), Object)
        End Function
        Public Shared Function new_SafeMeetingItem() As SafeMeetingItem
            Return DirectCast(NewRedemptionObject(New Guid("FA2CBAFB-F7B1-4F41-9B7A-73329A6C1CB7")), Object)
        End Function
        Public Shared Function new_SafePostItem() As SafePostItem
            Return DirectCast(NewRedemptionObject(New Guid("11E2BC0C-5D4F-4E0C-B438-501FFE05A382")), Object)
        End Function
        Public Shared Function new_SafeReportItem() As SafeReportItem
            Return DirectCast(NewRedemptionObject(New Guid("D46BA7B2-899F-4F60-85C7-4DF5713F6F18")), Object)
        End Function
        Public Shared Function new_MAPIFolder() As MAPIFolder
            Return DirectCast(NewRedemptionObject(New Guid("03C4C5F4-1893-444C-B8D8-002F0034DA92")), Object)
        End Function
        Public Shared Function new_SafeCurrentUser() As SafeCurrentUser
            Return DirectCast(NewRedemptionObject(New Guid("7ED1E9B1-CB57-4FA0-84E8-FAE653FE8E6B")), Object)
        End Function
        Public Shared Function new_SafeDistList() As SafeDistList
            Return DirectCast(NewRedemptionObject(New Guid("7C4A630A-DE98-4E3E-8093-E8F5E159BB72")), Object)
        End Function
        Public Shared Function new_AddressLists() As AddressLists
            Return DirectCast(NewRedemptionObject(New Guid("37587889-FC28-4507-B6D3-8557305F7511")), Object)
        End Function
        Public Shared Function new_MAPITable() As MAPITable
            Return DirectCast(NewRedemptionObject(New Guid("A6931B16-90FA-4D69-A49F-3ABFA2C04060")), Object)
        End Function
        Public Shared Function new_MAPIUtils() As MAPIUtils
            Return DirectCast(NewRedemptionObject(New Guid("4A5E947E-C407-4DCC-A0B5-5658E457153B")), Object)
        End Function
        Public Shared Function new_SafeInspector() As SafeInspector
            Return DirectCast(NewRedemptionObject(New Guid("ED323630-B4FD-4628-BC6A-D4CC44AE3F00")), Object)
        End Function
#End Region

#Region "private methods"

        Shared Sub New()
            Dim thisPath As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
            DllLocation64Bit = IO.Path.Combine(thisPath, "Redemption64.dll")
            DllLocation32Bit = IO.Path.Combine(thisPath, "Redemption.dll")
        End Sub

        <ComVisible(False)> _
        <ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00000001-0000-0000-C000-000000000046")> _
        Private Interface IClassFactory
            Sub CreateInstance(<MarshalAs(UnmanagedType.[Interface])> ByVal pUnkOuter As Object, ByRef refiid As Guid, <MarshalAs(UnmanagedType.[Interface])> ByRef ppunk As Object)
            Sub LockServer(ByVal fLock As Boolean)
        End Interface

        <ComVisible(False)> _
        <ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("00000000-0000-0000-C000-000000000046")> _
        Private Interface IUnknown
        End Interface

        Private Delegate Function DllGetClassObject(ByRef ClassId As Guid, ByRef InterfaceId As Guid, <Out(), MarshalAs(UnmanagedType.[Interface])> ByRef ppunk As Object) As Integer
        Private Delegate Function DllCanUnloadNow() As Integer

        'COM GUIDs
        Private Shared IID_IClassFactory As New Guid("00000001-0000-0000-C000-000000000046")
        Private Shared IID_IUnknown As New Guid("00000000-0000-0000-C000-000000000046")

        'win32 functions to load\unload dlls and get a function pointer
        Private Class Win32NativeMethods
            <DllImport("kernel32.dll", CharSet:=CharSet.Ansi)> _
            Public Shared Function GetProcAddress(ByVal hModule As IntPtr, ByVal lpProcName As String) As IntPtr
            End Function
            <DllImport("kernel32.dll")> _
            Public Shared Function FreeLibrary(ByVal hModule As IntPtr) As Boolean
            End Function
            <DllImport("kernel32.dll", CharSet:=CharSet.Unicode)> _
            Public Shared Function LoadLibraryW(ByVal lpFileName As String) As IntPtr
            End Function
        End Class

        'private variables
        Private Shared _redemptionDllHandle As IntPtr = IntPtr.Zero
        Private Shared _DllGetClassObject As IntPtr = IntPtr.Zero

        Private Shared Function NewRedemptionObject(ByVal guid As Guid) As IUnknown
            Dim res As Object = Nothing
            Dim ClassFactory As IClassFactory
            If _redemptionDllHandle.Equals(IntPtr.Zero) Then
                Dim dllPath As String
                If IntPtr.Size = 8 Then
                    dllPath = DllLocation64Bit
                Else
                    dllPath = DllLocation32Bit
                End If
                _redemptionDllHandle = Win32NativeMethods.LoadLibraryW(dllPath)
                If _redemptionDllHandle.Equals(IntPtr.Zero) Then
                    Throw New Exception(String.Format("Could not load '{0}'" & vbLf & "Make sure the dll exists.", dllPath))
                End If
                _DllGetClassObject = Win32NativeMethods.GetProcAddress(_redemptionDllHandle, "DllGetClassObject")
                If _DllGetClassObject.Equals(IntPtr.Zero) Then
                    Throw New Exception("Could not retrieve a pointer to the 'DllGetClassObject' function exported by the dll")
                End If
            End If
            Dim dllGetClassObject As DllGetClassObject = DirectCast(Marshal.GetDelegateForFunctionPointer(_DllGetClassObject, GetType(DllGetClassObject)), DllGetClassObject)
            Dim unk As Object
            Dim hr As Integer = dllGetClassObject(guid, IID_IClassFactory, unk)
            If hr <> 0 Then
                Throw New Exception("DllGetClassObject failed")
            End If
            ClassFactory = TryCast(unk, IClassFactory)
            ClassFactory.CreateInstance(Nothing, IID_IUnknown, res)
            Return TryCast(res, IUnknown)
        End Function

#End Region

    End Class



End Namespace