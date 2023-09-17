Attribute VB_Name = "CoreModule"
'This module contains this program's core procedures.
Option Explicit

'Defines the Microsoft Office constants used by this program.
Private Const xlWorkbookNormal As Long = -4143

'This structure defines a user's information.
Private Type UserStr
   LAN As String               'Defines a user's LAN id.
   Groups() As String          'Defines the list of groups a user belongs to.
   Descriptions() As String    'Defines the descriptions for the groups a user belongs to.
   UserDoesNotHave() As Long   'Defines the indexes of groups a user does not belong to, but an example user does belong to.
   Domain As String            'Defines the domain a user has logged onto.
   PDC As String               'Defines the primary domain controller.
End Type

Private Const UNKNOWN_NUMBER As Long = -1   'Defines unknown number of items in a given dimension in an array.

'This procedure compares the specified users's groups.
Private Sub CompareGroups(User As UserStr, ExampleUser As UserStr)
On Error GoTo ErrorTrap
Dim ExampleIndex As Long
Dim Index As Long
Dim UserHasGroupToo As Boolean

   With ExampleUser
      ReDim .UserDoesNotHave(0 To 0) As Long
      
      For ExampleIndex = LBound(.Groups()) To UBound(.Groups()) - 1
         UserHasGroupToo = False
         For Index = 0 To UBound(User.Groups()) - 1
            If User.Groups(Index) = .Groups(ExampleIndex) Then
               UserHasGroupToo = True
               Exit For
            End If
         Next Index
         If Not UserHasGroupToo Then
            .UserDoesNotHave(UBound(.UserDoesNotHave())) = ExampleIndex
            ReDim Preserve .UserDoesNotHave(LBound(.UserDoesNotHave()) To UBound(.UserDoesNotHave()) + 1) As Long
         End If
      Next ExampleIndex
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure returns the specified path's absolute path.
Private Function GetAbsolutePath(Path As String) As String
On Error GoTo ErrorTrap
Dim AbsolutePath As String

   AbsolutePath = CreateObject("Scripting.Filesystemobject").GetAbsolutePathName(Path)
   
EndRoutine:
   GetAbsolutePath = AbsolutePath
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure handes any errors that occur.
Private Sub HandleError(Optional ExtraInformation As String = vbNullString)
Dim ErrorCode As Long
Dim Message As String

   Message = Err.Description
   ErrorCode = Err.Number
   
   On Error Resume Next
   If Not ExtraInformation = vbNullString Then Message = Message & vbCr & ExtraInformation
   Message = Message & vbCr & "Error code: " & CStr(ErrorCode)
   
   Screen.MousePointer = vbDefault
   MsgBox Message, vbExclamation, App.Title
End Sub

'This procedure returns the zero based number if items in an array for the specified dimension.
Private Function ItemCount(ArrayV As Variant, Optional Dimension As Long = 1) As Long
On Error GoTo ErrorTrap
Dim Count As Long

   Count = UBound(ArrayV, Dimension) - LBound(ArrayV, Dimension)
EndRoutine:
   ItemCount = Count
   Exit Function

ErrorTrap:
   Count = UNKNOWN_NUMBER
   HandleError
   Resume EndRoutine
End Function



'This procedure is executed when this program starts.
Private Sub Main()
On Error GoTo ErrorTrap
Dim ExampleUser As UserStr
Dim ListPath As String
Dim Message As String
Dim User As UserStr

   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   ListPath = GetAbsolutePath(Command$())
   
   If ListPath = vbNullString Then
      User.LAN = InputBox$("Specify the user's LAN id:", ProgramInformation())
      If Not User.LAN = vbNullString Then
         ExampleUser.LAN = InputBox$("Specify the example user's LAN id:", ProgramInformation())
         ProcessUser User, ExampleUser
      End If
   Else
      ProcessUserList ListPath
   End If

   
   If Not User.LAN = vbNullString Then
      With User
         Message = "Done:" & vbCr
         Message = Message & "User: """ & .LAN & """" & vbCr
         Message = Message & "Example: """ & ExampleUser.LAN & """" & vbCr
         Message = Message & "Domain: """ & .Domain & """" & vbCr
         Message = Message & "PDC: """ & .PDC & """"
   
         MsgBox Message, vbInformation
      End With
   End If
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure manages the reference to a Microsoft Excel session.
Private Function MSExcel(Optional NewSession As Object = Nothing) As Object
On Error GoTo ErrorTrap
Static Excel As Object

   If Not NewSession Is Nothing Then Set Excel = NewSession
   
EndRoutine:
   Set MSExcel = Excel
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure processes the specified user(s)'s information.
Private Sub ProcessUser(User As UserStr, ExampleUser As UserStr)
On Error GoTo ErrorTrap
   MSExcel NewSession:=CreateObject("Excel.Application")
   MSExcel().DisplayAlerts = False
   MSExcel().Interactive = False
   MSExcel().ScreenUpdating = False
   
   Screen.MousePointer = vbHourglass: DoEvents
   If ExampleUser.LAN = vbNullString Or ExampleUser.LAN = User.LAN Then
      RetrieveGroups User
      SaveGroups User, ProgramPath() & User.LAN & ".xls"
   Else
      RetrieveGroups User
      RetrieveGroups ExampleUser
      CompareGroups User, ExampleUser
      CompareGroups ExampleUser, User
      
      SaveGroups User, ProgramPath() & User.LAN & ".xls"
      SaveGroups ExampleUser, ProgramPath() & ExampleUser.LAN & ".xls"
      SaveComparisonResult ExampleUser, ProgramPath() & ExampleUser.LAN & " does have and " & User.LAN & " does not.xls"
      SaveComparisonResult User, ProgramPath() & User.LAN & " does have and " & ExampleUser.LAN & " does not.xls"
   End If
   Screen.MousePointer = vbDefault
   
EndRoutine:
   If Not MSExcel() Is Nothing Then
      MSExcel().DisplayAlerts = True
      MSExcel().Interactive = True
      MSExcel().ScreenUpdating = True
      MSExcel().Quit
   End If
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure process the specified user list.
Private Sub ProcessUserList(Path As String)
On Error GoTo ErrorTrap
Dim ExampleUser As UserStr
Dim Row As Long
Dim User As UserStr
Dim Workbook As Object

   Set Workbook = CreateObject(Path)
   
   With Workbook.Worksheets(1)
      For Row = 1 To .UsedRange.Rows.Count
         User.LAN = .Range("A" & CStr(Row)).Value
         If .UsedRange.Columns.Count = 2 Then ExampleUser.LAN = .Range("B" & CStr(Row)).Value
         ProcessUser User, ExampleUser
      Next Row
   End With
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub



'This procedure returns this program's information.
Private Function ProgramInformation() As String
On Error GoTo ErrorTrap
   With App
      ProgramInformation = App.Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & App.CompanyName
   End With
EndRoutine:
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure returns this program's path ensuring it ends with a backslash.
Private Function ProgramPath() As String
On Error GoTo ErrorTrap
   ProgramPath = App.Path
   If Not Right$(ProgramPath, 1) = "\" Then ProgramPath = ProgramPath & "\"
EndRoutine:
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure gives the command to save the comparison's result.
Private Sub SaveComparisonResult(User As UserStr, Path As String)
On Error GoTo ErrorTrap
Dim Index As Long
Dim Table() As String

   If ItemCount(User.Groups()) > 0 And ItemCount(User.UserDoesNotHave()) > 0 Then
      With User
         ReDim Table(LBound(.UserDoesNotHave()) To UBound(.UserDoesNotHave()) - 1, 0 To 1) As String
         
         For Index = LBound(Table(), 1) To UBound(Table(), 1)
            Table(Index, 0) = .Groups(.UserDoesNotHave(Index))
            Table(Index, 1) = .Descriptions(.UserDoesNotHave(Index))
         Next Index
      End With
   
      SaveToExcel Table(), Path
   End If

EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure gives the command to save the specified user's groups.
Private Sub SaveGroups(User As UserStr, Path As String)
On Error GoTo ErrorTrap
Dim Index As Long
Dim Table() As String

   If ItemCount(User.Groups()) > 0 Then
      With User
         ReDim Table(LBound(.Groups()) To UBound(.Groups()) - 1, 0 To 1) As String
         
         For Index = LBound(Table(), 1) To UBound(Table(), 1)
            Table(Index, 0) = .Groups(Index)
            Table(Index, 1) = .Descriptions(Index)
         Next Index
      End With
      
      SaveToExcel Table(), Path
   End If
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub
'This procedure saves the specifed information at the specified path.
Private Sub SaveToExcel(Table() As String, Path As String)
On Error GoTo ErrorTrap
Dim Workbook As Object
Dim Worksheet As Object

   MSExcel().Workbooks.Add
   
   Set Workbook = MSExcel().Workbooks.Item(1)
   Workbook.Activate
   Do While Workbook.Worksheets.Count > 1
      Workbook.Worksheets(2).Delete
   Loop
   
   Set Worksheet = Workbook.Worksheets.Item(1)
   Worksheet.Activate
   Worksheet.Range("A1:B" & CStr(UBound(Table()) + 1)).Value = Table()
   Worksheet.Range("A:B").Columns.AutoFit
   
   Workbook.SaveAs Path, xlWorkbookNormal
   Workbook.Close
   
EndRoutine:
   Set Worksheet = Nothing
   Set Workbook = Nothing
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub



'This procedure retrieves the specified user's groups.
Private Sub RetrieveGroups(User As UserStr)
On Error GoTo ErrorTrap
Dim PropertyO As Variant
Dim SystemInformation As Object
Dim UserO As Object

   With User
      ReDim .Groups(0 To 0) As String
      ReDim .Descriptions(0 To 0) As String
   
      Set SystemInformation = CreateObject("WinNTSystemInfo")
      .PDC = SystemInformation.PDC
      
      If .PDC = vbNullString Then
         .Domain = Environ$("userdomain")
      Else
         .Domain = CreateObject("AdSystemInfo").DomainDNSName
      End If

      Set UserO = GetObject("WinNT://" & .Domain & "/" & .LAN)
       
      For Each PropertyO In UserO.Groups
         .Groups(UBound(.Groups())) = PropertyO.Name
         .Descriptions(UBound(.Descriptions())) = PropertyO.Description
         
         ReDim Preserve .Groups(LBound(.Groups()) To UBound(.Groups()) + 1) As String
         ReDim Preserve .Descriptions(LBound(.Descriptions()) To UBound(.Descriptions()) + 1) As String
      Next PropertyO
      
      Set UserO = Nothing
      Set SystemInformation = Nothing
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError ExtraInformation:="LAN: """ & User.LAN & """" & vbCr & "Domain: """ & User.Domain & """" & vbCr & "PDC: """ & User.PDC & """"
   Resume EndRoutine
End Sub

