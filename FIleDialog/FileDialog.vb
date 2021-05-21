Option Strict On
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

' Reference : https://github.com/microsoft/Windows-classic-samples/tree/master/Samples/Win7Samples/winui/shell/appplatform/CommonFileDialogModes

''' <summary>
''' Explorateur de fichiers
''' </summary>
Public Class FileDialog

    Public Sub New(Optional nOptions As FOS = 0, Optional sTitle As String = Nothing, Optional sFolder As String = Nothing, Optional sButtonText As String = Nothing, Optional ByVal listener As PushButtonClickedDelegate = Nothing)
        Options = nOptions
        Title = sTitle
        Folder = sFolder
        ButtonText = sButtonText
        OnPushButtonClickedListener = listener
    End Sub

    Public Delegate Function PushButtonClickedDelegate(ByVal sender As FileDialog) As DialogResult

    ''' <summary>
    ''' Le délégué à appeler lorsque l'utilisateur clique sur le bouton personnalisé.
    ''' </summary>
    ''' <returns></returns>
    Public Property OnPushButtonClickedListener As PushButtonClickedDelegate

    ''' <summary>
    ''' Si défini, le texte à afficher pour le bouton personnalisé. Si non défini, le bouton personnalisé n'est pas affiché.
    ''' </summary>
    ''' <returns></returns>
    Public Property ButtonText As String

    Public Property Options As FOS
    Public Property Title As String

    Private _folder As String

    ''' <summary>
    ''' Si défini avant l'affichage, le répertoire initialement affiché.
    ''' </summary>
    ''' <returns></returns>
    Public Property Folder As String
        <Obsolete>
        Get
            Return _folder
        End Get
        Set(value As String)
            _folder = value
        End Set
    End Property

    ''' <summary>
    ''' Le fichier sélectionné par l'utilisateur.
    ''' </summary>
    ''' <returns></returns>
    Public Property FileName As String


    ''' <summary>
    ''' Un objet personnalisé associé à cette boîte de dialogue. 
    ''' </summary>
    ''' <returns></returns>
    Public Property Tag As Object

    Private eventSink As DialogEventSink

    Public Function ShowDialog(ByVal hwndOwner As IntPtr) As DialogResult

        Dim hr As HRESULT
        Dim fod As IFileOpenDialog = CType(New FileOpenDialog(), IFileOpenDialog)
        Try
            Dim nOptions As FOS = 0
            hr = fod.GetOptions(nOptions)
            nOptions = nOptions Or Options
            hr = fod.SetOptions(nOptions)

            If Not String.IsNullOrEmpty(Title) Then hr = fod.SetTitle(Title)

            If Not String.IsNullOrEmpty(_folder) Then
                Dim GUID_IShellItem As Guid = GetType(IShellItem).GUID
                Dim psi As IntPtr = IntPtr.Zero
                hr = SHCreateItemFromParsingName(_folder, IntPtr.Zero, GUID_IShellItem, psi)
                If (hr = HRESULT.S_OK) Then
                    Dim si As IShellItem = CType(Marshal.GetObjectForIUnknown(psi), IShellItem)
                    hr = fod.SetFolder(si)
                    Marshal.ReleaseComObject(si)
                End If
            End If

            Dim nCookie As UInteger
            eventSink = New DialogEventSink(Me)
            hr = fod.Advise(eventSink, nCookie)
            eventSink.Cookie = nCookie

            If ButtonText <> Nothing Then
                Dim pfdc As IFileDialogCustomize
                pfdc = CType(fod, IFileDialogCustomize)
                hr = pfdc.AddPushButton(601, ButtonText)
            End If

            hr = fod.Show(hwndOwner)
            If (hr = HRESULT.S_OK) Then
                Dim pShellItemResult As IShellItem = Nothing
                Dim sbResult As System.Text.StringBuilder = New System.Text.StringBuilder(MAX_PATH)
                hr = fod.GetResult(pShellItemResult)
                If (hr = HRESULT.S_OK) Then
                    hr = pShellItemResult.GetDisplayName(SIGDN.SIGDN_FILESYSPATH, sbResult)
                    If (hr = HRESULT.S_OK) Then FileName = sbResult.ToString()
                    Marshal.ReleaseComObject(pShellItemResult)
                End If
            ElseIf hr = HRESULT.E_CANCELLED Then
                Return DialogResult.Cancel
            End If

            Return DialogResult.OK
        Finally
            Marshal.ReleaseComObject(fod)
        End Try
    End Function

    Public Const MAX_PATH = 260
    Public Enum HRESULT As Integer
        S_OK = 0
        S_FALSE = 1
        E_NOINTERFACE = &H80004002
        E_NOTIMPL = &H80004001
        E_FAIL = &H80004005
        E_INVALIDARG = &H80070057
        E_CANCELLED = &H80070000 + ERROR_CANCELLED
        E_UNEXPECTED = &H8000FFFF
    End Enum

    Public Const ERROR_CANCELLED = 1223L

    <DllImport("Shell32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function SHCreateItemFromParsingName(pszPath As String, pbc As IntPtr, ByRef riid As Guid, ByRef ppv As IntPtr) As HRESULT
    End Function


    <DllImport("Shlwapi.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function IUnknown_QueryService(ByVal punk As IntPtr, ByRef guidService As Guid, ByRef riid As Guid, <Out> ByRef ppvOut As IntPtr) As HRESULT
    End Function

    Private Class DialogEventSink : Implements IFileDialogEvents, IFileDialogControlEvents

        Private ReadOnly parent As FileDialog

        Public Sub New(ByVal parent As FileDialog)
            Me.parent = parent
        End Sub

        Public Property Cookie As UInteger

        ' If defined as Function => "Attempted to read or write protected memory. This is often an indication that other memory is corrupt "
        Public Sub OnButtonClicked(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal pfdc As IFileDialogCustomize, <[In]> ByVal dwIDCtl As Integer) Implements IFileDialogControlEvents.OnButtonClicked
            Select Case dwIDCtl
                Case 601
                    If parent.OnPushButtonClickedListener IsNot Nothing Then
                        Dim result As DialogResult = parent.OnPushButtonClickedListener.Invoke(parent)
                        If result = DialogResult.OK Or result = DialogResult.Yes Then
                            CType(pfdc, IFileDialog).Close(HRESULT.S_OK)
                        ElseIf result = DialogResult.No Or result = DialogResult.Abort Then
                            CType(pfdc, IFileDialog).Close(HRESULT.E_CANCELLED)
                        End If
                    End If
                Case Else
            End Select
        End Sub

        Public Function OnCheckButtonToggled(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal pfdc As IFileDialogCustomize, <[In]> ByVal dwIDCtl As Integer, <[In]> ByVal bChecked As Boolean) As HRESULT Implements IFileDialogControlEvents.OnCheckButtonToggled
            Return HRESULT.E_NOTIMPL
        End Function

        Public Function OnControlActivating(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal pfdc As IFileDialogCustomize, <[In]> ByVal dwIDCtl As Integer) As HRESULT Implements IFileDialogControlEvents.OnControlActivating
            Return HRESULT.E_NOTIMPL
        End Function
        Public Function OnItemSelected(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal pfdc As IFileDialogCustomize, <[In]> ByVal dwIDCtl As Integer, <[In]> ByVal dwIDItem As Integer) As HRESULT Implements IFileDialogControlEvents.OnItemSelected
            Return HRESULT.E_NOTIMPL
        End Function

        Private Function OnFileOk(<[In]> <MarshalAs(UnmanagedType.Interface)> pfd As IFileDialog) As HRESULT Implements IFileDialogEvents.OnFileOk
            Return HRESULT.S_OK
        End Function

        Private Function OnFolderChanging(<[In]> <MarshalAs(UnmanagedType.Interface)> pfd As IFileDialog, <[In]> <MarshalAs(UnmanagedType.Interface)> psiFolder As IShellItem) As HRESULT Implements IFileDialogEvents.OnFolderChanging
            Return HRESULT.E_NOTIMPL
        End Function

        ' If defined as Function => "Attempted to read or write protected memory. This is often an indication that other memory is corrupt "
        Private Sub OnFolderChange(<[In]> <MarshalAs(UnmanagedType.Interface)> pfd As IFileDialog) Implements IFileDialogEvents.OnFolderChange

        End Sub

        Private Sub OnSelectionChange(<[In]> <MarshalAs(UnmanagedType.Interface)> pfd As IFileDialog) Implements IFileDialogEvents.OnSelectionChange

        End Sub

        Private Sub OnShareViolation(<[In]> <MarshalAs(UnmanagedType.Interface)> pfd As IFileDialog, <[In]> <MarshalAs(UnmanagedType.Interface)> psi As IShellItem, <Out> ByRef pResponse As FDE_SHAREVIOLATION_RESPONSE) Implements IFileDialogEvents.OnShareViolation

        End Sub

        Private Sub OnTypeChange(<[In]> <MarshalAs(UnmanagedType.Interface)> pfd As IFileDialog) Implements IFileDialogEvents.OnTypeChange

        End Sub

        Private Sub OnOverwrite(<[In]> <MarshalAs(UnmanagedType.Interface)> pfd As IFileDialog, <[In]> <MarshalAs(UnmanagedType.Interface)> psi As IShellItem, <Out> ByRef pResponse As FDE_OVERWRITE_RESPONSE) Implements IFileDialogEvents.OnOverwrite

        End Sub
    End Class


    <ComImport>
    <Guid("cde725b0-ccc9-4519-917e-325d72fab4ce")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IFolderView
        Function GetCurrentViewMode(ByRef pViewMode As UInteger) As HRESULT
        Function SetCurrentViewMode(ViewMode As UInteger) As HRESULT
        Function GetFolder(ByRef riid As Guid, ByRef ppv As IntPtr) As HRESULT
        Function Item(iItemIndex As Integer, ByRef ppidl As IntPtr) As HRESULT
        Function ItemCount(uFlags As UInteger, ByRef pcItems As Integer) As HRESULT
        Function Items(uFlags As UInteger, ByRef riid As Guid, ByRef ppv As IntPtr) As HRESULT
        Function GetSelectionMarkedItem(ByRef piItem As Integer) As HRESULT
        Function GetFocusedItem(ByRef piItem As Integer) As HRESULT
        Function GetItemPosition(pidl As IntPtr, ByRef ppt As Point) As HRESULT
        Function GetSpacing(ByRef ppt As Point) As HRESULT
        Function GetDefaultSpacing(ByRef ppt As Point) As HRESULT
        Function GetAutoArrange() As HRESULT
        Function SelectItem(iItem As Integer, dwFlags As Integer) As HRESULT
        Function SelectAndPositionItems(cidl As UInteger, apidl As IntPtr, apt As Point, dwFlags As Integer) As HRESULT
    End Interface

    <ComImport>
    <Guid("1af3a467-214f-4298-908e-06b03e0b39f9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IFolderView2
        Inherits IFolderView
#Region "IFolderView"
        Overloads Function GetCurrentViewMode(ByRef pViewMode As UInteger) As HRESULT
        Overloads Function SetCurrentViewMode(ViewMode As UInteger) As HRESULT
        Overloads Function GetFolder(ByRef riid As Guid, ByRef ppv As IntPtr) As HRESULT
        Overloads Function Item(iItemIndex As Integer, ByRef ppidl As IntPtr) As HRESULT
        Overloads Function ItemCount(uFlags As UInteger, ByRef pcItems As Integer) As HRESULT
        Overloads Function Items(uFlags As UInteger, ByRef riid As Guid, ByRef ppv As IntPtr) As HRESULT
        Overloads Function GetSelectionMarkedItem(ByRef piItem As Integer) As HRESULT
        Overloads Function GetFocusedItem(ByRef piItem As Integer) As HRESULT
        Overloads Function GetItemPosition(pidl As IntPtr, ByRef ppt As Point) As HRESULT
        Overloads Function GetSpacing(ByRef ppt As Point) As HRESULT
        Overloads Function GetDefaultSpacing(ByRef ppt As Point) As HRESULT
        Overloads Function GetAutoArrange() As HRESULT
        Overloads Function SelectItem(iItem As Integer, dwFlags As Integer) As HRESULT
        Overloads Function SelectAndPositionItems(cidl As UInteger, apidl As IntPtr, apt As Point, dwFlags As Integer) As HRESULT
#End Region
        Function SetGroupBy(key As PROPERTYKEY, fAscending As Boolean) As HRESULT

        Function GetGroupBy(ByRef pkey As PROPERTYKEY, ByRef pfAscending As Boolean) As HRESULT

        'DEPRECATED
        Function SetViewProperty(pidl As IntPtr, propkey As PROPERTYKEY, propvar As PROPVARIANT) As HRESULT
        'DEPRECATED
        Function GetViewProperty(pidl As IntPtr, propkey As PROPERTYKEY, ByRef ppropvar As PROPVARIANT) As HRESULT
        'DEPRECATED
        Function SetTileViewProperties(pidl As IntPtr, pszPropList As String) As HRESULT
        'DEPRECATED
        Function SetExtendedTileViewProperties(pidl As IntPtr, pszPropList As String) As HRESULT

        Function SetText(iType As FVTEXTTYPE, pwszText As String) As HRESULT
        Function SetCurrentFolderFlags(dwMask As Integer, dwFlags As Integer) As HRESULT
        Function GetCurrentFolderFlags(ByRef pdwFlags As Integer) As HRESULT
        Function GetSortColumnCount(ByRef pcColumns As Integer) As HRESULT
        Function SetSortColumns(rgSortColumns As SORTCOLUMN, cColumns As Integer) As HRESULT
        Function GetSortColumns(ByRef rgSortColumns As SORTCOLUMN, cColumns As Integer) As HRESULT
        Function GetItem(iItem As Integer, ByRef riid As Guid, ByRef ppv As IntPtr) As HRESULT
        Function GetVisibleItem(iStart As Integer, fPrevious As Boolean, ByRef piItem As Integer) As HRESULT
        Function GetSelectedItem(iStart As Integer, ByRef piItem As Integer) As HRESULT
        Function GetSelection(fNoneImpliesFolder As Boolean, ByRef ppsia As IShellItemArray) As HRESULT
        Function GetSelectionState(pidl As IntPtr, ByRef pdwFlags As Integer) As HRESULT
        Function InvokeVerbOnSelection(pszVerb As String) As HRESULT
        Function SetViewModeAndIconSize(uViewMode As FOLDERVIEWMODE, iImageSize As Integer) As HRESULT
        Function GetViewModeAndIconSize(ByRef puViewMode As FOLDERVIEWMODE, ByRef piImageSize As Integer) As HRESULT
        Function SetGroupSubsetCount(cVisibleRows As UInteger) As HRESULT
        Function GetGroupSubsetCount(ByRef pcVisibleRows As UInteger) As HRESULT
        Function SetRedraw(fRedrawOn As Boolean) As HRESULT
        Function IsMoveInSameFolder() As HRESULT
        Function DoRename() As HRESULT
    End Interface

    Public Enum FVTEXTTYPE
        FVST_EMPTYTEXT = 0
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure PROPARRAY
        Public cElems As UInteger
        Public pElems As IntPtr
    End Structure

    <StructLayout(LayoutKind.Explicit, Pack:=1)>
    Public Structure PROPVARIANT
        <FieldOffset(0)>
        Public varType As UShort
        <FieldOffset(2)>
        Public wReserved1 As UShort
        <FieldOffset(4)>
        Public wReserved2 As UShort
        <FieldOffset(6)>
        Public wReserved3 As UShort
        <FieldOffset(8)>
        Public bVal As Byte
        <FieldOffset(8)>
        Public cVal As SByte
        <FieldOffset(8)>
        Public uiVal As UShort
        <FieldOffset(8)>
        Public iVal As Short
        <FieldOffset(8)>
        Public uintVal As UInt32
        <FieldOffset(8)>
        Public intVal As Int32
        <FieldOffset(8)>
        Public ulVal As UInt64
        <FieldOffset(8)>
        Public lVal As Int64
        <FieldOffset(8)>
        Public fltVal As Single
        <FieldOffset(8)>
        Public dblVal As Double
        <FieldOffset(8)>
        Public boolVal As Short
        <FieldOffset(8)>
        Public pclsidVal As IntPtr
        <FieldOffset(8)>
        Public pszVal As IntPtr
        <FieldOffset(8)>
        Public pwszVal As IntPtr
        <FieldOffset(8)>
        Public punkVal As IntPtr
        <FieldOffset(8)>
        Public ca As PROPARRAY
        <FieldOffset(8)>
        Public filetime As System.Runtime.InteropServices.ComTypes.FILETIME
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure SORTCOLUMN
        Public propkey As PROPERTYKEY
        Public direction As SORTDIRECTION
    End Structure

    Public Enum SORTDIRECTION
        SORT_DESCENDING = -1
        SORT_ASCENDING = 1
    End Enum

    Public Enum FOLDERVIEWMODE As Integer
        FVM_AUTO = -1
        FVM_FIRST = 1
        FVM_ICON = 1
        FVM_SMALLICON = 2
        FVM_LIST = 3
        FVM_DETAILS = 4
        FVM_THUMBNAIL = 5
        FVM_TILE = 6
        FVM_THUMBSTRIP = 7
        FVM_CONTENT = 8
        FVM_LAST = 8
    End Enum

    <ComImport, Guid("b4db1657-70d7-485e-8e3e-6fcb5a5c1802"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Interface IModalWindow
        <Runtime.CompilerServices.MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), PreserveSig>
        Function Show(<[In]> ByVal hwndOwner As IntPtr) As UInteger
    End Interface


    <ComImport, Guid("42f85136-db7e-439c-85f1-e4075d135fc8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Interface IFileDialog : Inherits IModalWindow
#Region "IModalWindow"
        <Runtime.CompilerServices.MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), PreserveSig>
        Overloads Function Show(<[In]> ByVal parent As IntPtr) As HRESULT
#End Region
        Function SetFileTypes(<[In]> ByVal cFileTypes As UInteger, <[In]> <MarshalAs(UnmanagedType.LPArray)> ByVal rgFilterSpec As COMDLG_FILTERSPEC()) As HRESULT
        Function SetFileTypeIndex(<[In]> ByVal iFileType As UInteger) As HRESULT
        Function GetFileTypeIndex(<Out> ByRef piFileType As UInteger) As HRESULT
        Function Advise(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal pfde As IFileDialogEvents, <Out> ByRef pdwCookie As UInteger) As HRESULT
        Function Unadvise(<[In]> ByVal dwCookie As UInteger) As HRESULT
        Function SetOptions(<[In]> ByVal fos As FOS) As HRESULT
        Function GetOptions(<Out> ByRef pfos As FOS) As HRESULT
        Function SetDefaultFolder(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal psi As IShellItem) As HRESULT
        Function SetFolder(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal psi As IShellItem) As HRESULT
        Function GetFolder(<Out> <MarshalAs(UnmanagedType.[Interface])> ByRef ppsi As IShellItem) As HRESULT
        Function GetCurrentSelection(<Out> <MarshalAs(UnmanagedType.[Interface])> ByRef ppsi As IShellItem) As HRESULT
        Function SetFileName(<[In], MarshalAs(UnmanagedType.LPWStr)> ByVal pszName As String) As HRESULT
        Function GetFileName(<Out> <MarshalAs(UnmanagedType.LPWStr)> ByRef pszName As String) As HRESULT
        Function SetTitle(<[In], MarshalAs(UnmanagedType.LPWStr)> ByVal pszTitle As String) As HRESULT
        Function SetOkButtonLabel(<[In], MarshalAs(UnmanagedType.LPWStr)> ByVal pszText As String) As HRESULT
        Function SetFileNameLabel(<[In], MarshalAs(UnmanagedType.LPWStr)> ByVal pszLabel As String) As HRESULT
        <PreserveSig>
        Function GetResult(<Out> <MarshalAs(UnmanagedType.[Interface])> ByRef ppsi As IShellItem) As HRESULT
        Function AddPlace(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal psi As IShellItem, ByVal fdap As FDAP) As HRESULT
        Function SetDefaultExtension(<[In], MarshalAs(UnmanagedType.LPWStr)> ByVal pszDefaultExtension As String) As HRESULT
        Function Close(<MarshalAs(UnmanagedType.[Error])> ByVal hr As Integer) As HRESULT
        Function SetClientGuid(<[In]> ByRef guid As Guid) As HRESULT
        Function ClearClientData() As HRESULT
        Function SetFilter(<MarshalAs(UnmanagedType.[Interface])> ByVal pFilter As IntPtr) As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure COMDLG_FILTERSPEC
        <MarshalAs(UnmanagedType.LPWStr)>
        Public pszName As String
        <MarshalAs(UnmanagedType.LPWStr)>
        Public pszSpec As String
    End Structure

    Public Enum FOS As Integer
        FOS_OVERWRITEPROMPT = &H2
        FOS_STRICTFILETYPES = &H4
        FOS_NOCHANGEDIR = &H8
        FOS_PICKFOLDERS = &H20
        FOS_FORCEFILESYSTEM = &H40
        FOS_ALLNONSTORAGEITEMS = &H80
        FOS_NOVALIDATE = &H100
        FOS_ALLOWMULTISELECT = &H200
        FOS_PATHMUSTEXIST = &H800
        FOS_FILEMUSTEXIST = &H1000
        FOS_CREATEPROMPT = &H2000
        FOS_SHAREAWARE = &H4000
        FOS_NOREADONLYRETURN = &H8000
        FOS_NOTESTFILECREATE = &H10000
        FOS_HIDEMRUPLACES = &H20000
        FOS_HIDEPINNEDPLACES = &H40000
        FOS_NODEREFERENCELINKS = &H100000
        FOS_DONTADDTORECENT = &H2000000
        FOS_FORCESHOWHIDDEN = &H10000000
        FOS_DEFAULTNOMINIMODE = &H20000000
    End Enum

    <ComImport()> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> <Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")>
    Public Interface IShellItem
        Function BindToHandler(ByVal pbc As IntPtr, ByRef bhid As Guid, ByRef riid As Guid, ByRef ppv As IntPtr) As HRESULT
        Function GetParent(ByRef ppsi As IShellItem) As HRESULT
        Function GetDisplayName(ByVal sigdnName As SIGDN, ByRef ppszName As System.Text.StringBuilder) As HRESULT
        Function GetAttributes(ByVal sfgaoMask As UInteger, ByRef psfgaoAttribs As UInteger) As HRESULT
        Function Compare(ByVal psi As IShellItem, ByVal hint As UInteger, ByRef piOrder As Integer) As HRESULT
    End Interface

    Public Const SFGAO_DROPTARGET = &H100L 'Objects are drop target
    Public Const SFGAO_FOLDER = &H20000000L 'Support BindToObject(IID_IShellFolder)
    Public Const SFGAO_FILESYSTEM = &H40000000L 'Is a win32 file system Object (file/folder/root)
    Public Const SFGAO_STREAM = &H400000L 'Supports BindToObject(IID_IStream)
    Public Enum SIGDN As Integer
        SIGDN_NORMALDISPLAY = &H0
        SIGDN_PARENTRELATIVEPARSING = &H80018001
        SIGDN_DESKTOPABSOLUTEPARSING = &H80028000
        SIGDN_PARENTRELATIVEEDITING = &H80031001
        SIGDN_DESKTOPABSOLUTEEDITING = &H8004C000
        SIGDN_FILESYSPATH = &H80058000
        SIGDN_URL = &H80068000
        SIGDN_PARENTRELATIVEFORADDRESSBAR = &H8007C001
        SIGDN_PARENTRELATIVE = &H80080001
    End Enum


    <ComImport, Guid("973510DB-7D7F-452B-8975-74A85828D354"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Interface IFileDialogEvents
        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), PreserveSig>
        Function OnFileOk(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal pfd As IFileDialog) As HRESULT

        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), PreserveSig>
        Function OnFolderChanging(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal pfd As IFileDialog, <[In], MarshalAs(UnmanagedType.[Interface])> ByVal psiFolder As IShellItem) As HRESULT

        ' If defined as Function => "Attempted to read or write protected memory. This is often an indication that other memory is corrupt "
        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Sub OnFolderChange(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal pfd As IFileDialog)

        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Sub OnSelectionChange(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal pfd As IFileDialog)

        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Sub OnShareViolation(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal pfd As IFileDialog, <[In], MarshalAs(UnmanagedType.[Interface])> ByVal psi As IShellItem, <Out> ByRef pResponse As FDE_SHAREVIOLATION_RESPONSE)

        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Sub OnTypeChange(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal pfd As IFileDialog)

        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Sub OnOverwrite(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal pfd As IFileDialog, <[In], MarshalAs(UnmanagedType.[Interface])> ByVal psi As IShellItem, <Out> ByRef pResponse As FDE_OVERWRITE_RESPONSE)
    End Interface

    Public Enum FDE_SHAREVIOLATION_RESPONSE
        FDESVR_DEFAULT = &H0
        FDESVR_ACCEPT = &H1
        FDESVR_REFUSE = &H2
    End Enum

    Public Enum FDE_OVERWRITE_RESPONSE
        FDEOR_DEFAULT = &H0
        FDEOR_ACCEPT = &H1
        FDEOR_REFUSE = &H2
    End Enum

    <ComImport, Guid("36116642-D713-4b97-9B83-7484A9D00433"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IFileDialogControlEvents
        Function OnItemSelected(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal pfdc As IFileDialogCustomize, <[In]> ByVal dwIDCtl As Integer, <[In]> ByVal dwIDItem As Integer) As HRESULT

        ' If defined as Function => "Attempted to read or write protected memory. This is often an indication that other memory is corrupt "
        Sub OnButtonClicked(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal pfdc As IFileDialogCustomize, <[In]> ByVal dwIDCtl As Integer)
        Function OnCheckButtonToggled(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal pfdc As IFileDialogCustomize, <[In]> ByVal dwIDCtl As Integer, <[In]> ByVal bChecked As Boolean) As HRESULT
        Function OnControlActivating(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal pfdc As IFileDialogCustomize, <[In]> ByVal dwIDCtl As Integer) As HRESULT
    End Interface


    <ComImport, Guid("e6fdd21a-163f-4975-9c8c-a69f1ba37034"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IFileDialogCustomize
        Function EnableOpenDropDown(<[In]> ByVal dwIDCtl As Integer) As HRESULT
        Function AddMenu(<[In]> ByVal dwIDCtl As Integer, <[In], MarshalAs(UnmanagedType.LPWStr)> ByVal pszLabel As String) As HRESULT
        Function AddPushButton(<[In]> ByVal dwIDCtl As Integer, <[In], MarshalAs(UnmanagedType.LPWStr)> ByVal pszLabel As String) As HRESULT
        Function AddComboBox(<[In]> ByVal dwIDCtl As Integer) As HRESULT
        Function AddRadioButtonList(<[In]> ByVal dwIDCtl As Integer) As HRESULT
        Function AddCheckButton(<[In]> ByVal dwIDCtl As Integer, <[In], MarshalAs(UnmanagedType.LPWStr)> ByVal pszLabel As String, <[In]> ByVal bChecked As Boolean) As HRESULT
        Function AddEditBox(<[In]> ByVal dwIDCtl As Integer, <[In], MarshalAs(UnmanagedType.LPWStr)> ByVal pszText As String) As HRESULT
        Function AddSeparator(<[In]> ByVal dwIDCtl As Integer) As HRESULT
        Function AddText(<[In]> ByVal dwIDCtl As Integer, <[In], MarshalAs(UnmanagedType.LPWStr)> ByVal pszText As String) As HRESULT
        Function SetControlLabel(<[In]> ByVal dwIDCtl As Integer, <[In], MarshalAs(UnmanagedType.LPWStr)> ByVal pszLabel As String) As HRESULT
        Function GetControlState(<[In]> ByVal dwIDCtl As Integer, <Out> ByRef pdwState As CDCONTROLSTATEF) As HRESULT
        Function SetControlState(<[In]> ByVal dwIDCtl As Integer, <[In]> ByVal dwState As CDCONTROLSTATEF) As HRESULT
        Function GetEditBoxText(<[In]> ByVal dwIDCtl As Integer, <Out> <MarshalAs(UnmanagedType.LPWStr)> ByRef ppszText As String) As HRESULT
        Function SetEditBoxText(<[In]> ByVal dwIDCtl As Integer, <[In], MarshalAs(UnmanagedType.LPWStr)> ByVal pszText As String) As HRESULT
        Function GetCheckButtonState(<[In]> ByVal dwIDCtl As Integer, <Out> ByRef pbChecked As Boolean) As HRESULT
        Function SetCheckButtonState(<[In]> ByVal dwIDCtl As Integer, <[In]> ByVal bChecked As Boolean) As HRESULT
        Function AddControlItem(<[In]> ByVal dwIDCtl As Integer, <[In]> ByVal dwIDItem As Integer, <[In], MarshalAs(UnmanagedType.LPWStr)> ByVal pszLabel As String) As HRESULT
        Function RemoveControlItem(<[In]> ByVal dwIDCtl As Integer, <[In]> ByVal dwIDItem As Integer) As HRESULT
        Function RemoveAllControlItems(<[In]> ByVal dwIDCtl As Integer) As HRESULT
        Function GetControlItemState(<[In]> ByVal dwIDCtl As Integer, <[In]> ByVal dwIDItem As Integer, <Out> ByRef pdwState As CDCONTROLSTATEF) As HRESULT
        Function SetControlItemState(<[In]> ByVal dwIDCtl As Integer, <[In]> ByVal dwIDItem As Integer, <[In]> ByVal dwState As CDCONTROLSTATEF) As HRESULT
        Function GetSelectedControlItem(<[In]> ByVal dwIDCtl As Integer, <Out> ByRef pdwIDItem As Integer) As HRESULT
        Function SetSelectedControlItem(<[In]> ByVal dwIDCtl As Integer, <[In]> ByVal dwIDItem As Integer) As HRESULT
        Function StartVisualGroup(<[In]> ByVal dwIDCtl As Integer, <[In], MarshalAs(UnmanagedType.LPWStr)> ByVal pszLabel As String) As HRESULT
        Function EndVisualGroup() As HRESULT
        Function MakeProminent(<[In]> ByVal dwIDCtl As Integer) As HRESULT
        Function SetControlItemText(<[In]> ByVal dwIDCtl As Integer, <[In]> ByVal dwIDItem As Integer, <[In]> ByVal pszLabel As String) As HRESULT
    End Interface

    Public Enum CDCONTROLSTATEF
        CDCS_INACTIVE = 0
        CDCS_ENABLED = &H1
        CDCS_VISIBLE = &H2
        CDCS_ENABLEDVISIBLE = &H3
    End Enum

    <ComImport, Guid("d57c7288-d4ad-4768-be02-9d969532d960"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Interface IFileOpenDialog : Inherits IFileDialog

#Region "IFileDialog"
#Region "IModalWindow"
        <Runtime.CompilerServices.MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), PreserveSig>
        Overloads Function Show(<[In]> ByVal hwndOwner As IntPtr) As HRESULT
#End Region
        Overloads Function SetFileTypes(<[In]> ByVal cFileTypes As UInteger, <[In]> <MarshalAs(UnmanagedType.LPArray)> ByVal rgFilterSpec As COMDLG_FILTERSPEC()) As HRESULT
        Overloads Function SetFileTypeIndex(<[In]> ByVal iFileType As UInteger) As HRESULT
        Overloads Function GetFileTypeIndex(<Out> ByRef piFileType As UInteger) As HRESULT
        Overloads Function Advise(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal pfde As IFileDialogEvents, <Out> ByRef pdwCookie As UInteger) As HRESULT
        Overloads Function Unadvise(<[In]> ByVal dwCookie As UInteger) As HRESULT
        Overloads Function SetOptions(<[In]> ByVal fos As FOS) As HRESULT
        Overloads Function GetOptions(<Out> ByRef pfos As FOS) As HRESULT
        Overloads Function SetDefaultFolder(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal psi As IShellItem) As HRESULT
        Overloads Function SetFolder(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal psi As IShellItem) As HRESULT
        Overloads Function GetFolder(<Out> <MarshalAs(UnmanagedType.[Interface])> ByRef ppsi As IShellItem) As HRESULT
        Overloads Function GetCurrentSelection(<Out> <MarshalAs(UnmanagedType.[Interface])> ByRef ppsi As IShellItem) As HRESULT
        Overloads Function SetFileName(<[In], MarshalAs(UnmanagedType.LPWStr)> ByVal pszName As String) As HRESULT
        Overloads Function GetFileName(<Out> <MarshalAs(UnmanagedType.LPWStr)> ByRef pszName As String) As HRESULT
        Overloads Function SetTitle(<[In], MarshalAs(UnmanagedType.LPWStr)> ByVal pszTitle As String) As HRESULT
        Overloads Function SetOkButtonLabel(<[In], MarshalAs(UnmanagedType.LPWStr)> ByVal pszText As String) As HRESULT
        Overloads Function SetFileNameLabel(<[In], MarshalAs(UnmanagedType.LPWStr)> ByVal pszLabel As String) As HRESULT
        <PreserveSig>
        Overloads Function GetResult(<Out> <MarshalAs(UnmanagedType.[Interface])> ByRef ppsi As IShellItem) As HRESULT
        Overloads Function AddPlace(<[In], MarshalAs(UnmanagedType.[Interface])> ByVal psi As IShellItem, ByVal fdap As FDAP) As HRESULT
        Overloads Function SetDefaultExtension(<[In], MarshalAs(UnmanagedType.LPWStr)> ByVal pszDefaultExtension As String) As HRESULT
        Overloads Function Close(<MarshalAs(UnmanagedType.[Error])> ByVal hr As Integer) As HRESULT
        Overloads Function SetClientGuid(<[In]> ByRef guid As Guid) As HRESULT
        Overloads Function ClearClientData() As HRESULT
        Overloads Function SetFilter(<MarshalAs(UnmanagedType.[Interface])> ByVal pFilter As IntPtr) As HRESULT
#End Region
        <PreserveSig>
        Function GetResults(<Out> <MarshalAs(UnmanagedType.[Interface])> ByRef ppenum As IShellItemArray) As HRESULT
        Function GetSelectedItems(<Out> <MarshalAs(UnmanagedType.[Interface])> ByRef ppsai As IShellItemArray) As HRESULT
    End Interface

    Public Enum FDAP
        FDAP_BOTTOM
        FDAP_TOP
    End Enum

    <ComImport()> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> <Guid("b63ea76d-1f85-456f-a19c-48159efa858b")>
    Public Interface IShellItemArray
        Function BindToHandler(pbc As IntPtr, ByRef bhid As Guid, ByRef riid As Guid, ByRef ppvOut As IntPtr) As HRESULT
        Function GetPropertyStore(flags As GETPROPERTYSTOREFLAGS, ByRef riid As Guid, ByRef ppv As IntPtr) As HRESULT
        Function GetPropertyDescriptionList(keyType As PROPERTYKEY, ByRef riid As Guid, ByRef ppv As IntPtr) As HRESULT
        'Function GetAttributes(AttribFlags As SIATTRIBFLAGS, sfgaoMask As SFGAOF, ByRef psfgaoAttribs As SFGAOF) As HRESULT
        Function GetAttributes(AttribFlags As SIATTRIBFLAGS, sfgaoMask As Integer, ByRef psfgaoAttribs As Integer) As HRESULT
        Function GetCount(ByRef pdwNumItems As Integer) As HRESULT
        Function GetItemAt(dwIndex As Integer, ByRef ppsi As IShellItem) As HRESULT

        'Function EnumItems(ByRef ppenumShellItems As IEnumShellItems) As HRESULT
        Function EnumItems(ByRef ppenumShellItems As IntPtr) As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential, Pack:=4)>
    Public Structure PROPERTYKEY
        Private fmtid As Guid
        Private ReadOnly pid As Integer
        Public ReadOnly Property FormatId() As Guid
            Get
                Return Me.fmtid
            End Get
        End Property
        Public ReadOnly Property PropertyId() As Integer
            Get
                Return Me.pid
            End Get
        End Property
        Public Sub New(ByVal formatId As Guid, ByVal propertyId As Integer)
            Me.fmtid = formatId
            Me.pid = propertyId
        End Sub
        Public Shared ReadOnly PKEY_DateCreated As PROPERTYKEY = New PROPERTYKEY(New Guid("B725F130-47EF-101A-A5F1-02608C9EEBAC"), 15)
    End Structure

    Public Enum GETPROPERTYSTOREFLAGS
        GPS_DEFAULT = 0
        GPS_HANDLERPROPERTIESONLY = &H1
        GPS_READWRITE = &H2
        GPS_TEMPORARY = &H4
        GPS_FASTPROPERTIESONLY = &H8
        GPS_OPENSLOWITEM = &H10
        GPS_DELAYCREATION = &H20
        GPS_BESTEFFORT = &H40
        GPS_NO_OPLOCK = &H80
        GPS_PREFERQUERYPROPERTIES = &H100
        GPS_EXTRINSICPROPERTIES = &H200
        GPS_EXTRINSICPROPERTIESONLY = &H400
        GPS_MASK_VALID = &H7FF
    End Enum

    Public Enum SIATTRIBFLAGS
        SIATTRIBFLAGS_AND = &H1
        SIATTRIBFLAGS_OR = &H2
        SIATTRIBFLAGS_APPCOMPAT = &H3
        SIATTRIBFLAGS_MASK = &H3
        SIATTRIBFLAGS_ALLITEMS = &H4000
    End Enum


    <ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")>
    Private Class FileOpenDialog
    End Class

End Class

''' <summary>
''' Explorateur de dossiers
''' </summary>
Public Class FolderBrowserDialog
    Inherits FileDialog

    Public Sub New(Optional ByVal options As FOS = 0, Optional ByVal title As String = Nothing, Optional ByVal folder As String = Nothing, Optional ByVal textButton As String = Nothing, Optional ByVal listener As PushButtonClickedDelegate = Nothing)
        MyBase.New(options Or FOS.FOS_PICKFOLDERS, title, folder, textButton, listener)
    End Sub
End Class
