'==============================================================================
'  MOKONIC - repair shop management workbook (one standard module)
'
'  1. Alt+F11 > Insert > Module, paste this whole module.
'  2. Run the macro Setup_Mokonic (Alt+F8).
'  3. Save the workbook as "Excel Macro-Enabled Workbook (*.xlsm)".
'
'  Persian text: every Persian string in this module is stored as plain
'  ASCII and decoded at run time by T():  "~XX" = Unicode U+06XX,
'  "^XXXX" = any other Unicode code point. This keeps the Persian text
'  intact on any Windows PC, whatever its "language for non-Unicode
'  programs" setting is. Messages use MessageBoxW for the same reason.
'==============================================================================
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

'------------------------------------------------------------------------------
' Settings you may change
'------------------------------------------------------------------------------
Private Const PROTECT_PWD As String = ""      ' sheet protection password ("" = none)
Private Const FONT_NAME As String = "Tahoma"  ' renders Persian on every Windows PC

'------------------------------------------------------------------------------
' Column positions inside each table (1-based)
'------------------------------------------------------------------------------
' Devices (tblDevices)
Private Const DC_CODE As Long = 1
Private Const DC_JDATE As Long = 2
Private Const DC_TIME As Long = 3
Private Const DC_GDATE As Long = 4
Private Const DC_CUSTCODE As Long = 5
Private Const DC_CUSTNAME As Long = 6
Private Const DC_PHONE As Long = 7
Private Const DC_TYPE As Long = 8
Private Const DC_BRAND As Long = 9
Private Const DC_SERIAL As Long = 10
Private Const DC_PROBLEM As Long = 11
Private Const DC_ACCESS As Long = 12
Private Const DC_COND As Long = 13
Private Const DC_SHELF As Long = 14
Private Const DC_STATUS As Long = 15
Private Const DC_TECH As Long = 16
Private Const DC_DIAG As Long = 17
Private Const DC_PARTS As Long = 18
Private Const DC_ESTIMATE As Long = 19
Private Const DC_PARTSCOST As Long = 20
Private Const DC_LABOR As Long = 21
Private Const DC_TOTAL As Long = 22
Private Const DC_PAID As Long = 23
Private Const DC_BALANCE As Long = 24
Private Const DC_LASTCHG As Long = 25
Private Const DC_DELIVERY As Long = 26
Private Const DC_WARRANTY As Long = 27
Private Const DC_STAFF As Long = 28
Private Const DC_NOTES As Long = 29
Private Const DC_GDELIV As Long = 30
' Customers (tblCustomers)
Private Const CC_CODE As Long = 1
Private Const CC_NAME As Long = 2
Private Const CC_PHONE As Long = 3
Private Const CC_PHONE2 As Long = 4
Private Const CC_ADDR As Long = 5
Private Const CC_FIRST As Long = 6
Private Const CC_COUNT As Long = 7
Private Const CC_PAID As Long = 8
Private Const CC_NOTES As Long = 9
' Accounting (tblAcc)
Private Const AC_JDATE As Long = 1
Private Const AC_GDATE As Long = 2
Private Const AC_TYPE As Long = 3
Private Const AC_CAT As Long = 4
Private Const AC_REF As Long = 5
Private Const AC_DESC As Long = 6
Private Const AC_AMOUNT As Long = 7
Private Const AC_METHOD As Long = 8
Private Const AC_MONTH As Long = 9
' Parts (tblParts)
Private Const PC_CODE As Long = 1
Private Const PC_NAME As Long = 2
Private Const PC_CAT As Long = 3
Private Const PC_SPEC As Long = 4
Private Const PC_SHELF As Long = 5
Private Const PC_QTY As Long = 6
Private Const PC_UNIT As Long = 7
Private Const PC_MIN As Long = 8
Private Const PC_BUY As Long = 9
Private Const PC_SELL As Long = 10
Private Const PC_VALUE As Long = 11
Private Const PC_SUPP As Long = 12
' Purchases (tblPurchases)
Private Const BC_CODE As Long = 1
Private Const BC_JDATE As Long = 2
Private Const BC_SUPP As Long = 3
Private Const BC_PART As Long = 4
Private Const BC_PNAME As Long = 5
Private Const BC_QTY As Long = 6
Private Const BC_PRICE As Long = 7
Private Const BC_TOTAL As Long = 8
Private Const BC_METHOD As Long = 9
Private Const BC_NOTES As Long = 10
Private Const BC_GDATE As Long = 11
' Tools (tblTools)
Private Const TC_CODE As Long = 1
Private Const TC_NAME As Long = 2
Private Const TC_CAT As Long = 3
Private Const TC_BRAND As Long = 4
Private Const TC_LOC As Long = 5
Private Const TC_COND As Long = 6
Private Const TC_PDATE As Long = 7
Private Const TC_PRICE As Long = 8
Private Const TC_HOLDER As Long = 9
Private Const TC_NOTES As Long = 10
' Shelves (tblShelves)
Private Const SC_CODE As Long = 1
Private Const SC_DESC As Long = 2
Private Const SC_TYPE As Long = 3
Private Const SC_CAP As Long = 4
Private Const SC_DEV As Long = 5
Private Const SC_PARTS As Long = 6
Private Const SC_TOOLS As Long = 7
Private Const SC_FREE As Long = 8

' Status positions in the status list (Settings sheet)
Private Const ST_RECEIVED As Long = 1
Private Const ST_DIAGNOSING As Long = 2
Private Const ST_APPROVAL As Long = 3
Private Const ST_PARTS As Long = 4
Private Const ST_REPAIRING As Long = 5
Private Const ST_READY As Long = 6
Private Const ST_DELIVERED As Long = 7
Private Const ST_RETURNED As Long = 8
Private Const STATUS_COUNT As Long = 8

' Accounting category positions (Settings sheet)
Private Const CAT_DEPOSIT As Long = 1
Private Const CAT_REPAIR As Long = 2
Private Const CAT_PARTSALE As Long = 3
Private Const CAT_PARTBUY As Long = 4

' Transaction type positions
Private Const TT_INCOME As Long = 1
Private Const TT_EXPENSE As Long = 2

' Layout anchors
Private Const SEARCH_FIRST_ROW As Long = 9
Private Const SEARCH_MAX_ROWS As Long = 500
Private Const DASH_LOW_FIRST As Long = 22
Private Const DASH_LOW_LAST As Long = 33

Private mStep As String   ' current build step, shown in error messages

'==============================================================================
'  SETUP
'==============================================================================

' Builds the whole workbook. Safe to re-run: asks before deleting anything.
Public Sub Setup_Mokonic()
    Dim injected As Boolean, s As String
    On Error GoTo EH

    If MokonicSheetsExist() Then
        If Msg(T("~47~34~2F~27~31: ~28~31~AF~47^200C~47~27~CC ~45~48~A9~48~46~CC~A9 ~27~32 ~42~28~44 ~2F~31 ~27~CC~46 ~41~27~CC~44 ~48~2C~48~2F ~2F~27~31~46~2F.") & vbLf & vbLf & _
               T("~33~27~2E~2A ~2F~48~28~27~31~47~0C ~47~45~47 ~28~31~AF~47^200C~47~27~CC ~45~48~A9~48~46~CC~A9 ~48 ~47~45~47 ~27~37~44~27~39~27~2A ~2B~28~2A^200C~34~2F~47 ~2F~31 ~22~46^200C~47~27") & vbLf & _
               T("(~2F~33~2A~AF~27~47^200C~47~27~0C ~45~34~2A~31~CC~27~46~0C ~2D~33~27~28^200C~47~27~0C ~27~46~28~27~31~0C ~2E~31~CC~2F~47~27 ~48 ~27~28~32~27~31) ~31~27 ~28~31~27~CC ~47~45~CC~34~47 ~7E~27~A9 ~45~CC^200C~A9~46~2F.") & vbLf & vbLf & _
               T("~22~CC~27 ~45~37~45~26~46 ~47~33~2A~CC~2F ~A9~47 ~45~CC^200C~2E~48~27~47~CC~2F ~47~45~47 ~86~CC~32 ~7E~27~A9 ~34~48~2F ~48 ~27~32 ~46~48 ~33~27~2E~2A~47 ~34~48~2F~1F"), _
               vbYesNo + vbCritical + vbDefaultButton2, T("~33~27~2E~2A ~2F~48~28~27~31~47 ~45~48~A9~48~46~CC~A9")) <> vbYes Then
            Msg T("~39~45~44~CC~27~2A ~44~3A~48 ~34~2F. ~47~CC~86 ~27~37~44~27~39~27~2A~CC ~7E~27~A9 ~46~34~2F."), vbInformation
            Exit Sub
        End If
    End If

    BeginBusy True
    Application.StatusBar = "Mokonic: building workbook..."

    mStep = "DeleteMokonicSheets": DeleteMokonicSheets
    mStep = "DeleteMokonicNames": DeleteMokonicNames
    mStep = "CreateSheets": CreateSheets

    ' 1) tables and static layout
    mStep = "Build_Settings": Build_Settings
    mStep = "Build_Shelves": Build_Shelves
    mStep = "Build_Customers": Build_Customers
    mStep = "Build_Stock": Build_Stock
    mStep = "Build_Devices": Build_Devices
    mStep = "Build_Accounting": Build_Accounting
    mStep = "Build_Purchases": Build_Purchases
    mStep = "Build_Tools": Build_Tools

    ' 2) names, formulas, dropdowns, colors (need every table to exist)
    mStep = "Build_Names": Build_Names
    mStep = "Build_Formulas": Build_Formulas
    mStep = "Build_Validation": Build_Validation
    mStep = "Build_ConditionalFormats": Build_ConditionalFormats

    ' 3) screens that use the names
    mStep = "Build_Intake": Build_Intake
    mStep = "Build_Search": Build_Search
    mStep = "Build_AccSummary": Build_AccSummary
    mStep = "Build_Dashboard": Build_Dashboard

    mStep = "RemoveBlankExtraSheets": RemoveBlankExtraSheets
    injected = TryInjectDevicesHandler()

    Application.Calculation = xlCalculationAutomatic
    mStep = "Dash_RefreshCore": Dash_RefreshCore
    mStep = "FinishSheets": FinishSheets
    mStep = ""
    EndBusy
    Application.StatusBar = False

    Sh("DASH").Activate
    Sh("DASH").Range("A1").Select

    s = T("~33~27~2E~2A ~41~27~CC~44 ~45~48~A9~48~46~CC~A9 ~28~27 ~45~48~41~42~CC~2A ~27~46~2C~27~45 ~34~2F.") & vbLf & vbLf
    If injected Then
        s = s & T("~2B~28~2A ~2E~48~2F~A9~27~31 ~2A~27~31~CC~2E ~2A~3A~CC~CC~31 ~48~36~39~CC~2A (~47~46~AF~27~45 ~27~46~2A~2E~27~28 ~48~36~39~CC~2A ~27~32 ~41~47~31~33~2A) ~41~39~27~44 ~34~2F.") & vbLf & vbLf
    Else
        s = s & T("~28~31~27~CC ~2A~3A~CC~CC~31 ~48~36~39~CC~2A ~2F~33~2A~AF~27~47 ~27~32 ~2F~A9~45~47 ^00AB~2A~3A~CC~CC~31 ~48~36~39~CC~2A^00BB ~27~33~2A~41~27~2F~47 ~A9~46~CC~2F.") & vbLf & vbLf
    End If
    s = s & T("~45~47~45: ~27~A9~46~48~46 ~41~27~CC~44 ~31~27 ~28~27 ~41~31~45~2A xlsm ~30~2E~CC~31~47 ~A9~46~CC~2F:") & vbLf & _
            "File > Save As > Excel Macro-Enabled Workbook (*.xlsm)"
    Msg s, vbInformation, T("~45~48~A9~48~46~CC~A9 ~22~45~27~2F~47 ~27~33~2A")
    Exit Sub
EH:
    Application.StatusBar = False
    FailWith "Setup_Mokonic"
End Sub

Private Sub CreateSheets()
    Dim keys As Variant, i As Long, ws As Worksheet, prev As Worksheet, k As String
    keys = SheetKeys()
    For i = 0 To UBound(keys)
        k = CStr(keys(i))
        If prev Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
        Else
            Set ws = ThisWorkbook.Worksheets.Add(After:=prev)
        End If
        ws.Name = SN(k)
        ws.Tab.Color = SheetTabColor(k)
        ws.DisplayRightToLeft = True
        ws.Cells.Font.Name = FONT_NAME
        ws.Cells.Font.Size = 10
        ws.Cells.VerticalAlignment = xlCenter
        ws.Columns(1).ColumnWidth = 2
        ws.Activate
        ActiveWindow.DisplayGridlines = Not (k = "DASH" Or k = "INTAKE" Or k = "SEARCH")
        ActiveWindow.Zoom = 100
        Set prev = ws
    Next i
End Sub

Private Sub DeleteMokonicSheets()
    Dim keys As Variant, i As Long, ws As Worksheet, tmp As Worksheet
    Set tmp = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))   ' a workbook needs one sheet
    keys = SheetKeys()
    For i = 0 To UBound(keys)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(SN(CStr(keys(i))))
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Unprotect PROTECT_PWD
            ws.Delete
        End If
    Next i
End Sub

' Deletes empty sheets that are not Mokonic sheets (e.g. "Sheet1" of a new workbook).
Private Sub RemoveBlankExtraSheets()
    Dim i As Long, ws As Worksheet
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        If Not IsMokonicSheet(ws.Name) Then
            If Application.WorksheetFunction.CountA(ws.Cells) = 0 And ws.Shapes.Count = 0 Then
                If ThisWorkbook.Worksheets.Count > 1 Then ws.Delete
            End If
        End If
    Next i
End Sub

Private Sub DeleteMokonicNames()
    Dim i As Long, nm As String
    For i = ThisWorkbook.Names.Count To 1 Step -1
        nm = ThisWorkbook.Names(i).Name
        If InStr(nm, "!") > 0 Then nm = Mid$(nm, InStr(nm, "!") + 1)
        If LCase$(Left$(nm, 3)) = "mk_" Then ThisWorkbook.Names(i).Delete
    Next i
End Sub

Private Function MokonicSheetsExist() As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsMokonicSheet(ws.Name) Then MokonicSheetsExist = True: Exit Function
    Next ws
End Function

Private Function IsMokonicSheet(ByVal sheetName As String) As Boolean
    Dim keys As Variant, i As Long
    keys = SheetKeys()
    For i = 0 To UBound(keys)
        If SN(CStr(keys(i))) = sheetName Then IsMokonicSheet = True: Exit Function
    Next i
End Function

' Final per-sheet touches: freeze panes, protection, first cell.
Private Sub FinishSheets()
    FreezeAt Sh("DEV"), "C6"
    FreezeAt Sh("CUST"), "A6"
    FreezeAt Sh("ACC"), "A6"
    FreezeAt Sh("STOCK"), "A6"
    FreezeAt Sh("TOOLS"), "A6"
    FreezeAt Sh("SHELF"), "A6"
    ProtectAll
    Sh("INTAKE").Activate
    NC("mk_in_Phone").Select
End Sub

Private Sub FreezeAt(ByVal ws As Worksheet, ByVal addr As String)
    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ws.Range(addr).Select
    ActiveWindow.FreezePanes = True
End Sub

' Optional: injects a Worksheet_Change handler into the Devices sheet so that
' choosing a status from the dropdown stamps the date automatically.
' Works only if "Trust access to the VBA project object model" is enabled.
Private Function TryInjectDevicesHandler() As Boolean
    Dim vbp As Object, comp As Object, found As Object, cm As Object, code As String
    On Error GoTo Fail
    Set vbp = ThisWorkbook.VBProject
    For Each comp In vbp.VBComponents
        If comp.Type = 100 Then
            If comp.Properties("Name").Value = SN("DEV") Then Set found = comp: Exit For
        End If
    Next comp
    If found Is Nothing Then GoTo Fail
    Set cm = found.CodeModule
    code = "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
           "    ' Added by Setup_Mokonic: stamps status-change dates automatically." & vbCrLf & _
           "    On Error Resume Next" & vbCrLf & _
           "    Mokonic_OnDevicesChange Target" & vbCrLf & _
           "End Sub"
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString code
    TryInjectDevicesHandler = True
    Exit Function
Fail:
    TryInjectDevicesHandler = False
End Function

'==============================================================================
'  BUILDERS - one per sheet
'==============================================================================

Private Sub Build_Settings()
    Dim ws As Worksheet, lo As ListObject, c As Long
    Set ws = Sh("SET")
    ws.Columns(2).ColumnWidth = 44
    ws.Columns(3).ColumnWidth = 34
    ws.Columns(4).ColumnWidth = 3

    SectionTitle ws.Range("B4:C4"), T("~27~37~44~27~39~27~2A ~41~31~48~34~AF~27~47 (~31~48~CC ~31~33~CC~2F ~48 ~33~31~28~31~AF^200C~47~27 ~86~27~7E ~45~CC^200C~34~48~2F)")
    KV ws, 5, T("~46~27~45 ~41~31~48~34~AF~27~47"), T("~45~48~A9~48~46~CC~A9"), "mk_ShopName", True
    KV ws, 6, T("~34~39~27~31 / ~2A~48~36~CC~2D ~32~CC~31 ~46~27~45"), T("~2A~39~45~CC~31 ~2A~2E~35~35~CC ~28~31~2F~47~27~CC ~27~44~A9~2A~31~48~46~CC~A9~CC ~48 ~44~48~27~32~45 ~28~31~42~CC"), "mk_ShopSlogan", True
    KV ws, 7, T("~2A~44~41~46 ~41~31~48~34~AF~27~47"), "021-00000000", "mk_ShopPhone", True
    KV ws, 8, T("~22~2F~31~33 ~41~31~48~34~AF~27~47"), T("~22~2F~31~33 ~41~31~48~34~AF~27~47 ~31~27 ~27~CC~46~2C~27 ~28~46~48~CC~33~CC~2F"), "mk_ShopAddr", True
    KV ws, 9, T("~45~2F~2A ~AF~27~31~27~46~2A~CC (~31~48~32) - ~35~41~31 = ~28~2F~48~46 ~AF~27~31~27~46~2A~CC"), 30, "mk_WarrantyDays", False
    KV ws, 10, T("~27~46~2F~27~32~47 ~A9~27~3A~30 ~31~33~CC~2F (80mm ~CC~27 A5)"), "80mm", "mk_ReceiptSize", True

    SectionTitle ws.Range("B12:C12"), T("~34~45~27~31~46~2F~47 ~A9~2F~47~27 (~2F~33~2A~CC ~2A~3A~CC~CC~31 ~46~2F~47~CC~2F)")
    KV ws, 13, T("~33~27~44 ~34~45~33~CC ~34~45~27~31~46~2F~47 ~2F~33~2A~AF~27~47"), JalaliYear(Date), "mk_cntDevYear", False
    KV ws, 14, T("~22~2E~31~CC~46 ~34~45~27~31~47 ~2F~33~2A~AF~27~47 ~2F~31 ~27~CC~46 ~33~27~44"), 0, "mk_cntDev", False
    KV ws, 15, T("~22~2E~31~CC~46 ~34~45~27~31~47 ~45~34~2A~31~CC"), 0, "mk_cntCust", False
    KV ws, 16, T("~22~2E~31~CC~46 ~34~45~27~31~47 ~42~37~39~47"), 0, "mk_cntPart", False
    KV ws, 17, T("~22~2E~31~CC~46 ~34~45~27~31~47 ~2E~31~CC~2F"), 0, "mk_cntBuy", False
    KV ws, 18, T("~22~2E~31~CC~46 ~34~45~27~31~47 ~27~28~32~27~31"), 0, "mk_cntTool", False
    ws.Range("C13:C18").Interior.Color = RGB(242, 242, 242)

    With ws.Range("B20:C23")
        .Merge
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Color = RGB(89, 89, 89)
        .Value = T("~31~27~47~46~45~27: ~41~47~31~33~2A^200C~47~27~CC ~33~45~2A ~86~7E~0C ~45~46~28~39 ~47~45~47 ~45~46~48~47~27~CC ~A9~34~48~CC~CC ~41~27~CC~44 ~47~33~2A~46~2F.") & vbLf & _
                 T("~28~31~27~CC ~27~41~32~48~2F~46 ~45~48~31~2F~0C ~2F~31 ~27~48~44~CC~46 ~2E~27~46~47 ~2E~27~44~CC ~32~CC~31 ~47~31 ~2C~2F~48~44 ~2A~27~CC~7E ~A9~46~CC~2F ~2A~27 ~2C~2F~48~44 ~28~32~31~AF ~34~48~2F.") & vbLf & _
                 T("~2F~31 ~41~47~31~33~2A ~48~36~39~CC~2A^200C~47~27 ~48 ~2F~33~2A~47^200C~47~27~CC ~2D~33~27~28~2F~27~31~CC ~41~42~37 ~46~27~45^200C~47~27 ~31~27 ~48~CC~31~27~CC~34 ~A9~46~CC~2F~1B ~2A~31~2A~CC~28 ~31~2F~CC~41^200C~47~27 ~31~27 ~2A~3A~CC~CC~31 ~46~2F~47~CC~2F ~48 ~31~2F~CC~41~CC ~2D~30~41 ~46~A9~46~CC~2F.")
    End With

    c = 5
    SetListCol ws, c, "tblDevTypes", T("~46~48~39 ~2F~33~2A~AF~27~47"), Array(T("~28~31~2F ~27~44~A9~2A~31~48~46~CC~A9~CC"), T("~2A~44~48~CC~32~CC~48~46"), T("~45~27~46~CC~2A~48~31"), T("~44~7E^200C~2A~27~7E"), T("~45~48~28~27~CC~44 ~48 ~2A~28~44~2A"), _
        T("~45~46~28~39 ~2A~3A~30~CC~47"), T("~27~CC~46~48~31~2A~31 ~48 UPS"), T("~34~27~31~98~31 ~48 ~22~2F~27~7E~2A~48~31"), T("~33~CC~33~2A~45 ~35~48~2A~CC"), T("~44~48~27~32~45 ~2E~27~46~AF~CC"), T("~27~28~32~27~31 ~28~31~42~CC"), T("~33~27~CC~31"))
    SetListCol ws, c, "tblStatus", T("~48~36~39~CC~2A ~2A~39~45~CC~31"), Array(T("~2A~2D~48~CC~44 ~AF~31~41~2A~47 ~34~2F~47"), T("~2F~31 ~2D~27~44 ~28~31~31~33~CC"), T("~45~46~2A~38~31 ~2A~23~CC~CC~2F ~45~34~2A~31~CC"), _
        T("~45~46~2A~38~31 ~42~37~39~47"), T("~2F~31 ~2D~27~44 ~2A~39~45~CC~31"), T("~2A~39~45~CC~31 ~34~2F~47 ^2013 ~22~45~27~2F~47 ~2A~2D~48~CC~44"), T("~2A~2D~48~CC~44 ~34~2F~47"), T("~3A~CC~31~42~27~28~44 ~2A~39~45~CC~31 / ~45~31~2C~48~39~CC"))
    SetListCol ws, c, "tblTech", T("~2A~A9~46~33~CC~46 / ~A9~27~31~45~46~2F"), Array(T("~2A~A9~46~33~CC~46 ~F1"), T("~2A~A9~46~33~CC~46 ~F2"), T("~2A~A9~46~33~CC~46 ~F3"))
    SetListCol ws, c, "tblAccCat", T("~2F~33~2A~47 ~2D~33~27~28~2F~27~31~CC"), Array(T("~28~CC~39~27~46~47"), T("~7E~31~2F~27~2E~2A ~2A~39~45~CC~31"), T("~41~31~48~34 ~42~37~39~47"), T("~2E~31~CC~2F ~42~37~39~47"), _
        T("~27~2C~27~31~47"), T("~2D~42~48~42"), T("~42~28~48~36 ~48 ~2E~2F~45~27~2A"), T("~27~28~32~27~31"), T("~33~27~CC~31"))
    SetListCol ws, c, "tblPay", T("~31~48~34 ~7E~31~2F~27~2E~2A"), Array(T("~46~42~2F~CC"), T("~A9~27~31~2A~2E~48~27~46"), T("~A9~27~31~2A ~28~47 ~A9~27~31~2A"), T("~27~46~2A~42~27~44 ~28~27~46~A9~CC"))
    SetListCol ws, c, "tblPartCat", T("~2F~33~2A~47 ~42~37~39~47"), Array(T("~22~CC^200C~33~CC"), T("~2E~27~32~46"), T("~45~42~27~48~45~2A"), T("~28~31~2F"), T("~A9~27~46~A9~2A~48~31"), T("~46~45~27~CC~34~AF~31"), T("~28~27~2A~31~CC"), T("~33~27~CC~31"))
    SetListCol ws, c, "tblUnits", T("~48~27~2D~2F"), Array(T("~39~2F~2F"), T("~45~2A~31"), T("~28~33~2A~47"), T("~2D~44~42~47"), T("~AF~31~45"))
    SetListCol ws, c, "tblToolCat", T("~2F~33~2A~47 ~27~28~32~27~31"), Array(T("~47~48~CC~47 ~48 ~44~2D~CC~45^200C~A9~27~31~CC"), T("~27~46~2F~27~32~47^200C~AF~CC~31~CC"), T("~45~46~28~39 ~2A~3A~30~CC~47"), T("~27~28~32~27~31 ~2F~33~2A~CC"), T("~33~27~CC~31"))
    SetListCol ws, c, "tblToolCond", T("~48~36~39~CC~2A ~27~28~32~27~31"), Array(T("~33~27~44~45"), T("~46~CC~27~32 ~28~47 ~2A~39~45~CC~31"), T("~2E~31~27~28"))
    SetListCol ws, c, "tblShelfType", T("~46~48~39 ~42~41~33~47"), Array(T("~2F~33~2A~AF~27~47 ~45~34~2A~31~CC"), T("~42~37~39~27~2A"), T("~27~28~32~27~31"))
    SetListCol ws, c, "tblAccType", T("~46~48~39 ~2A~31~27~A9~46~34"), Array(T("~2F~31~22~45~2F"), T("~47~32~CC~46~47"))
    SetListCol ws, c, "tblSearchBy", T("~2C~33~2A~2C~48 ~28~31 ~27~33~27~33"), Array(T("~2A~44~41~46"), T("~46~27~45 ~45~34~2A~31~CC"), T("~A9~2F ~2F~33~2A~AF~27~47"))

    ' status colors next to the status list, for reference
    Set lo = GetLO("tblStatus")
    For c = 1 To STATUS_COUNT
        lo.DataBodyRange.Cells(c, 1).Interior.Color = StatusFill(c)
        lo.DataBodyRange.Cells(c, 1).Font.Color = StatusFont(c)
    Next c

    HeaderBand ws, SheetTitle("SET"), 28
    BackButton ws
End Sub

' Adds one single-column list table at column c and moves c two columns on.
Private Sub SetListCol(ByVal ws As Worksheet, ByRef c As Long, ByVal tblName As String, ByVal header As String, ByVal items As Variant)
    Dim n As Long, i As Long, lo As ListObject
    n = UBound(items) - LBound(items) + 1
    ws.Cells(5, c).Value = header
    For i = 0 To n - 1
        ws.Cells(6 + i, c).Value = items(LBound(items) + i)
    Next i
    ws.Columns(c).ColumnWidth = 22
    ws.Columns(c + 1).ColumnWidth = 2
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(5, c), ws.Cells(5 + n, c)), , xlYes)
    lo.Name = tblName
    lo.TableStyle = "TableStyleMedium7"
    lo.HeaderRowRange.HorizontalAlignment = xlCenter
    c = c + 2
End Sub

Private Sub Build_Shelves()
    Dim ws As Worksheet, lo As ListObject, codes As Variant, i As Long, r As Long
    Set ws = Sh("SHELF")
    codes = Array("A1", "A2", "A3", "A4", "B1", "B2", "B3", "B4", "C1", "C2")
    For i = 0 To UBound(codes)
        r = 6 + i
        ws.Cells(r, 2).Value = codes(i)
        Select Case Left$(codes(i), 1)
            Case "A"
                ws.Cells(r, 3).Value = T("~2F~33~2A~AF~27~47^200C~47~27~CC ~45~34~2A~31~CC")
                ws.Cells(r, 4).Value = T("~2F~33~2A~AF~27~47 ~45~34~2A~31~CC")
                ws.Cells(r, 5).Value = 10
            Case "B"
                ws.Cells(r, 3).Value = T("~42~37~39~27~2A ~27~44~A9~2A~31~48~46~CC~A9~CC")
                ws.Cells(r, 4).Value = T("~42~37~39~27~2A")
                ws.Cells(r, 5).Value = 30
            Case Else
                ws.Cells(r, 3).Value = T("~27~28~32~27~31 ~A9~27~31~AF~27~47")
                ws.Cells(r, 4).Value = T("~27~28~32~27~31")
                ws.Cells(r, 5).Value = 20
        End Select
    Next i
    Set lo = MakeTable(ws, ws.Range("B5"), "tblShelves", ShelfHeaders(), _
        Array(10, 26, 16, 9, 13, 13, 11, 12), "TableStyleMedium4", UBound(codes) + 1)
    ColFmt lo, Array(SC_CODE, SC_CAP, SC_DEV, SC_PARTS, SC_TOOLS, SC_FREE), "", True
    ColFmt lo, Array(SC_CODE), "@"

    HeaderBand ws, SheetTitle("SHELF"), 9
    BackButton ws
    Note ws.Range("D2:I3"), T("~42~41~33~47 ~2C~2F~CC~2F: ~2F~31 ~27~48~44~CC~46 ~31~2F~CC~41 ~2E~27~44~CC ~32~CC~31 ~2C~2F~48~44 ~2A~27~CC~7E ~A9~46~CC~2F. ~33~2A~48~46^200C~47~27~CC ^00AB~45~48~2C~48~2F^00BB ~48 ^00AB~38~31~41~CC~2A ~22~32~27~2F^00BB ~2E~48~2F~A9~27~31 ~45~2D~27~33~28~47 ~45~CC^200C~34~48~46~2F. ~A9~2F~47~27~CC ~27~CC~46 ~2C~2F~48~44 ~2F~31 ~47~45~47 ~45~46~48~47~27~CC ^00AB~42~41~33~47^00BB ~46~45~27~CC~34 ~2F~27~2F~47 ~45~CC^200C~34~48~46~2F.")
End Sub

Private Sub Build_Customers()
    Dim ws As Worksheet, lo As ListObject
    Set ws = Sh("CUST")
    Set lo = MakeTable(ws, ws.Range("B5"), "tblCustomers", CustHeaders(), _
        Array(10, 24, 14, 14, 36, 13, 11, 16, 28), "TableStyleMedium2", 0)
    ColFmt lo, Array(CC_CODE, CC_PHONE, CC_PHONE2, CC_FIRST), "@", True
    ColFmt lo, Array(CC_COUNT), "0", True
    ColFmt lo, Array(CC_PAID), "#,##0"
    ColFmt lo, Array(CC_ADDR, CC_NOTES), "", False, True
    ApplyLockPattern lo

    HeaderBand ws, SheetTitle("CUST"), 10
    BackButton ws
    AddBtn ws, ws.Range("B2").Left + 120, ws.Range("B2").Top + 3, 120, 26, T("+ ~45~34~2A~31~CC ~2C~2F~CC~2F"), "Cust_Add", CLR("green")
    Note ws.Range("F2:J3"), T("~45~34~2A~31~CC~27~46 ~47~46~AF~27~45 ~7E~30~CC~31~34 ~2F~33~2A~AF~27~47 ~2E~48~2F~A9~27~31 ~33~27~2E~2A~47 ~45~CC^200C~34~48~46~2F. ~34~45~27~31~47 ~2A~A9~31~27~31~CC ~28~27 ~31~46~AF ~42~31~45~32 ~45~34~2E~35 ~45~CC^200C~34~48~2F. ~2A~39~2F~27~2F ~2F~33~2A~AF~27~47 ~48 ~2C~45~39 ~7E~31~2F~27~2E~2A~CC ~2E~48~2F~A9~27~31 ~45~2D~27~33~28~47 ~45~CC^200C~34~48~2F.")
End Sub

Private Sub Build_Stock()
    Dim ws As Worksheet, lo As ListObject, x As Double, y As Double
    Set ws = Sh("STOCK")
    Set lo = MakeTable(ws, ws.Range("B5"), "tblParts", PartHeaders(), _
        Array(10, 26, 12, 20, 9, 10, 9, 10, 14, 14, 15, 20), "TableStyleMedium3", 0)
    ColFmt lo, Array(PC_CODE, PC_SHELF), "@", True
    ColFmt lo, Array(PC_QTY, PC_MIN), "General", True
    ColFmt lo, Array(PC_UNIT, PC_CAT), "", True
    ColFmt lo, Array(PC_BUY, PC_SELL, PC_VALUE), "#,##0"
    ColFmt lo, Array(PC_SPEC), "", False, True
    ApplyLockPattern lo

    HeaderBand ws, SheetTitle("STOCK"), 13
    BackButton ws
    x = ws.Range("B2").Left + 120: y = ws.Range("B2").Top + 3
    AddBtn ws, x, y, 110, 26, T("+ ~42~37~39~47 ~2C~2F~CC~2F"), "Stock_AddPart", CLR("green")
    AddBtn ws, x + 116, y, 110, 26, T("~41~31~48~34 ~42~37~39~47"), "Stock_SellPart", CLR("blue")
    AddBtn ws, x + 232, y, 110, 26, T("~2B~28~2A ~2E~31~CC~2F"), "Go_Purchases", CLR("brown")
    Note ws.Range("I2:M3"), T("~31~2F~CC~41 ~42~31~45~32 ~CC~39~46~CC ~45~48~2C~48~2F~CC ~28~47 ~2D~2F~27~42~44 ~31~33~CC~2F~47 ~27~33~2A. ~27~41~32~27~CC~34 ~45~48~2C~48~2F~CC ~27~32 ~28~31~AF~47 ^00AB~2E~31~CC~2F^00BB ~27~46~2C~27~45 ~34~48~2F. ~47~45~47 ~45~28~27~44~3A ~28~47 ~2A~48~45~27~46 ~27~33~2A.")
End Sub

Private Sub Build_Devices()
    Dim ws As Worksheet, lo As ListObject, x As Double, y As Double, i As Long, w As Double
    Set ws = Sh("DEV")
    Set lo = MakeTable(ws, ws.Range("B5"), "tblDevices", DevHeaders(), _
        Array(15, 11, 7, 11, 9, 18, 13, 14, 18, 13, 30, 20, 18, 7, 24, 13, 30, 20, 13, 13, 13, 13, 13, 13, 17, 17, 11, 12, 26, 11), _
        "TableStyleMedium2", 0)
    ColFmt lo, Array(DC_CODE, DC_JDATE, DC_TIME, DC_CUSTCODE, DC_PHONE, DC_SERIAL, DC_SHELF, DC_LASTCHG, DC_DELIVERY, DC_WARRANTY), "@", True
    ColFmt lo, Array(DC_GDATE, DC_GDELIV), "yyyy/mm/dd", True, False, True
    ColFmt lo, Array(DC_ESTIMATE, DC_PARTSCOST, DC_LABOR, DC_TOTAL, DC_PAID), "#,##0"
    ColFmt lo, Array(DC_BALANCE), "#,##0;[Red]-#,##0"
    ColFmt lo, Array(DC_PROBLEM, DC_ACCESS, DC_COND, DC_DIAG, DC_PARTS, DC_NOTES), "", False, True
    ColFmt lo, Array(DC_STATUS, DC_TYPE, DC_TECH, DC_STAFF), "", True
    lo.ListColumns(DC_STATUS).DataBodyRange.Font.Bold = True
    ApplyLockPattern lo

    HeaderBand ws, SheetTitle("DEV"), 31
    ws.Rows(2).RowHeight = 30
    ws.Rows(3).RowHeight = 28
    BackButton ws
    x = ws.Range("B2").Left + 120: y = ws.Range("B2").Top + 3
    AddBtn ws, x, y, 120, 26, T("~2A~3A~CC~CC~31 ~48~36~39~CC~2A"), "Dev_ChangeStatus", CLR("green")
    AddBtn ws, x + 126, y, 120, 26, T("~2F~31~CC~27~41~2A ~48~2C~47"), "Dev_ReceivePayment", CLR("blue")
    AddBtn ws, x + 252, y, 130, 26, T("~45~35~31~41 ~42~37~39~47 ~27~32 ~27~46~28~27~31"), "Dev_UsePart", CLR("orange")
    AddBtn ws, x + 388, y, 110, 26, T("~86~27~7E ~31~33~CC~2F"), "Dev_PrintReceipt", CLR("gray")
    AddBtn ws, x + 504, y, 110, 26, T("~46~45~27~CC~34 ~47~45~47"), "Dev_ClearFilter", CLR("navy2")
    ' quick status filters
    x = ws.Range("B3").Left: y = ws.Range("B3").Top + 2: w = 128
    For i = 1 To STATUS_COUNT
        AddBtn ws, x + (i - 1) * (w + 4), y, w, 24, StatusName(i), "Dev_FilterByStatus", StatusFill(i), StatusFont(i), 8.5, "mk_flt_" & i
    Next i
    Note ws.Range("B4:P4"), T("~28~31~27~CC ^00AB~2A~3A~CC~CC~31 ~48~36~39~CC~2A^00BB~0C ^00AB~2F~31~CC~27~41~2A ~48~2C~47^00BB~0C ^00AB~45~35~31~41 ~42~37~39~47^00BB ~CC~27 ^00AB~86~27~7E ~31~33~CC~2F^00BB: ~27~48~44 ~31~48~CC ~31~2F~CC~41 ~2F~33~2A~AF~27~47 ~A9~44~CC~A9 ~A9~46~CC~2F~0C ~28~39~2F ~2F~A9~45~47 ~31~27 ~28~32~46~CC~2F. ~2F~A9~45~47^200C~47~27~CC ~31~46~AF~CC ~31~2F~CC~41 ~2F~48~45~0C ~41~47~31~33~2A ~31~27 ~28~31 ~27~33~27~33 ~48~36~39~CC~2A ~41~CC~44~2A~31 ~45~CC^200C~A9~46~46~2F.")
End Sub

Private Sub Build_Accounting()
    Dim ws As Worksheet, lo As ListObject
    Set ws = Sh("ACC")
    Set lo = MakeTable(ws, ws.Range("B5"), "tblAcc", AccHeaders(), _
        Array(12, 12, 9, 15, 15, 36, 15, 14, 10), "TableStyleMedium6", 0)
    ColFmt lo, Array(AC_JDATE, AC_REF), "@", True
    ColFmt lo, Array(AC_MONTH), "", True
    ColFmt lo, Array(AC_GDATE), "yyyy/mm/dd", True, False, True
    ColFmt lo, Array(AC_TYPE, AC_CAT, AC_METHOD), "", True
    ColFmt lo, Array(AC_AMOUNT), "#,##0"
    ColFmt lo, Array(AC_DESC), "", False, True
    ApplyLockPattern lo

    HeaderBand ws, SheetTitle("ACC"), 25
    BackButton ws
    AddBtn ws, ws.Range("B2").Left + 120, ws.Range("B2").Top + 3, 130, 26, T("+ ~2B~28~2A ~2F~33~2A~CC ~2F~31~22~45~2F/~47~32~CC~46~47"), "Acc_AddManual", CLR("green")
    AddBtn ws, ws.Range("B2").Left + 256, ws.Range("B2").Top + 3, 110, 26, T("~28~31~48~32~31~33~27~46~CC"), "Acc_Refresh", CLR("navy2")
    Note ws.Range("G2:L3"), T("~2F~31~22~45~2F~47~27 ~27~32 ~7E~30~CC~31~34~0C ~2A~2D~48~CC~44 ~48 ~41~31~48~34 ~42~37~39~47~0C ~48 ~47~32~CC~46~47^200C~47~27 ~27~32 ~2E~31~CC~2F ~42~37~39~47 ~2E~48~2F~A9~27~31 ~2B~28~2A ~45~CC^200C~34~48~46~2F. ~27~2C~27~31~47~0C ~2D~42~48~42 ~48 ~42~28~48~36 ~31~27 ~28~27 ^00AB~2B~28~2A ~2F~33~2A~CC^00BB ~48~27~31~2F ~A9~46~CC~2F.")
End Sub

Private Sub Build_Purchases()
    Dim ws As Worksheet, lo As ListObject, labels As Variant, nms As Variant, i As Long, r As Long, x As Double
    Set ws = Sh("BUY")
    Set lo = MakeTable(ws, ws.Range("B13"), "tblPurchases", BuyHeaders(), _
        Array(11, 11, 22, 11, 24, 9, 14, 15, 14, 26, 11), "TableStyleMedium10", 0)
    ColFmt lo, Array(BC_CODE, BC_JDATE, BC_PART), "@", True
    ColFmt lo, Array(BC_GDATE), "yyyy/mm/dd", True, False, True
    ColFmt lo, Array(BC_QTY), "General", True
    ColFmt lo, Array(BC_PRICE, BC_TOTAL), "#,##0"
    ColFmt lo, Array(BC_METHOD), "", True
    ColFmt lo, Array(BC_NOTES, BC_PNAME), "", False, True
    ApplyLockPattern lo

    HeaderBand ws, SheetTitle("BUY"), 12
    BackButton ws
    SectionTitle ws.Range("B4:K4"), T("~2B~28~2A ~2E~31~CC~2F ~2C~2F~CC~2F - ~41~31~45 ~31~27 ~7E~31 ~A9~46~CC~2F ~48 ^00AB~2B~28~2A ~2E~31~CC~2F^00BB ~31~27 ~28~32~46~CC~2F")
    labels = Array(T("~2A~23~45~CC~46^200C~A9~46~46~2F~47 (~41~31~48~34~46~2F~47)"), T("~A9~2F ~42~37~39~47 *"), T("~2A~39~2F~27~2F *"), T("~42~CC~45~2A ~48~27~2D~2F (~2A~48~45~27~46) *"), T("~31~48~34 ~7E~31~2F~27~2E~2A"), T("~2A~48~36~CC~2D~27~2A"))
    nms = Array("mk_buy_Supp", "mk_buy_Part", "mk_buy_Qty", "mk_buy_Price", "mk_buy_Method", "mk_buy_Notes")
    For i = 0 To UBound(labels)
        r = 5 + i
        ws.Rows(r).RowHeight = 22
        FormLabel ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)), CStr(labels(i))
        StyleInput ws.Range(ws.Cells(r, 4), ws.Cells(r, 6)), True
        NameCell CStr(nms(i)), ws.Cells(r, 4)
    Next i
    ws.Range("D6").NumberFormat = "@"
    ws.Range("D7").NumberFormat = "General"
    ws.Range("D8").NumberFormat = "#,##0"
    With ws.Range("G6:I6")
        .Merge
        .Font.Color = CLR("teal")
        .Font.Bold = True
    End With
    NameCell "mk_buy_PartName", ws.Range("G6")
    ws.Range("G6").Formula = "=IFERROR(INDEX(" & SRc(GetLO("tblParts"), PC_NAME) & ",MATCH(mk_buy_Part," & SRc(GetLO("tblParts"), PC_CODE) & ",0)),"""")"
    Note ws.Range("B11:K11"), T("~7E~33 ~27~32 ~2B~28~2A: ~45~48~2C~48~2F~CC ~42~37~39~47 ~2F~31 ~27~46~28~27~31 ~32~CC~27~2F ~45~CC^200C~34~48~2F~0C ~42~CC~45~2A ~2E~31~CC~2F ~42~37~39~47 ~28~47^200C~31~48~32 ~45~CC^200C~34~48~2F ~48 ~CC~A9 ~31~2F~CC~41 ^00AB~47~32~CC~46~47^00BB ~2F~31 ~2D~33~27~28~2F~27~31~CC ~2B~28~2A ~45~CC^200C~34~48~2F. ~42~37~39~47 ~2C~2F~CC~2F ~31~27 ~27~48~44 ~2F~31 ~28~31~AF~47 ^00AB~27~46~28~27~31^00BB ~2A~39~31~CC~41 ~A9~46~CC~2F.")
    x = ws.Range("J5").Left
    AddBtn ws, x, ws.Range("J5").Top + 2, 130, 40, T("~2B~28~2A ~2E~31~CC~2F"), "Buy_Save", CLR("green"), , 12
    AddBtn ws, x, ws.Range("J8").Top + 2, 130, 24, T("~7E~27~A9 ~A9~31~2F~46 ~41~31~45"), "Buy_Clear", CLR("gray")
    AddBtn ws, ws.Range("B2").Left + 120, ws.Range("B2").Top + 3, 110, 26, T("+ ~42~37~39~47 ~2C~2F~CC~2F"), "Stock_AddPart", CLR("orange")
End Sub

Private Sub Build_Tools()
    Dim ws As Worksheet, lo As ListObject
    Set ws = Sh("TOOLS")
    Set lo = MakeTable(ws, ws.Range("B5"), "tblTools", ToolHeaders(), _
        Array(9, 26, 17, 14, 15, 15, 12, 14, 15, 28), "TableStyleMedium1", 0)
    ColFmt lo, Array(TC_CODE, TC_PDATE, TC_LOC), "@", True
    ColFmt lo, Array(TC_CAT, TC_COND), "", True
    ColFmt lo, Array(TC_PRICE), "#,##0"
    ColFmt lo, Array(TC_NOTES), "", False, True
    ApplyLockPattern lo

    HeaderBand ws, SheetTitle("TOOLS"), 11
    BackButton ws
    AddBtn ws, ws.Range("B2").Left + 120, ws.Range("B2").Top + 3, 120, 26, T("+ ~27~28~32~27~31 ~2C~2F~CC~2F"), "Tools_Add", CLR("green")
    Note ws.Range("E2:K3"), T("~48~36~39~CC~2A ~27~28~32~27~31: ~33~28~32 = ~33~27~44~45~0C ~32~31~2F = ~46~CC~27~32 ~28~47 ~2A~39~45~CC~31~0C ~42~31~45~32 = ~2E~31~27~28. ~33~2A~48~46 ^00AB~2F~31 ~27~2E~2A~CC~27~31^00BB ~31~27 ~28~31~27~CC ~27~28~32~27~31~47~27~CC ~27~45~27~46~2A ~2F~27~2F~47^200C~34~2F~47 ~7E~31 ~A9~46~CC~2F.")
End Sub

Private Sub Build_Names()
    Dim i As Long
    AddListName "mk_lstDevTypes", "tblDevTypes"
    AddListName "mk_lstStatus", "tblStatus"
    AddListName "mk_lstTech", "tblTech"
    AddListName "mk_lstAccCat", "tblAccCat"
    AddListName "mk_lstPay", "tblPay"
    AddListName "mk_lstPartCat", "tblPartCat"
    AddListName "mk_lstUnits", "tblUnits"
    AddListName "mk_lstToolCat", "tblToolCat"
    AddListName "mk_lstToolCond", "tblToolCond"
    AddListName "mk_lstShelfType", "tblShelfType"
    AddListName "mk_lstAccType", "tblAccType"
    AddListName "mk_lstSearchBy", "tblSearchBy"
    AddListName "mk_lstShelves", "tblShelves"
    AddListName "mk_lstPartCodes", "tblParts"
    For i = 1 To STATUS_COUNT
        ThisWorkbook.Names.Add Name:="mk_st" & i, RefersTo:="=INDEX(mk_lstStatus," & i & ")"
    Next i
    For i = 1 To 3
        ThisWorkbook.Names.Add Name:="mk_cond" & i, RefersTo:="=INDEX(mk_lstToolCond," & i & ")"
    Next i
    ThisWorkbook.Names.Add Name:="mk_income", RefersTo:="=INDEX(mk_lstAccType,1)"
    ThisWorkbook.Names.Add Name:="mk_expense", RefersTo:="=INDEX(mk_lstAccType,2)"
End Sub

Private Sub AddListName(ByVal nm As String, ByVal tblName As String)
    ThisWorkbook.Names.Add Name:=nm, RefersTo:="=" & tblName & "[" & GetLO(tblName).ListColumns(1).Name & "]"
End Sub

' Calculated table columns.
Private Sub Build_Formulas()
    Dim dev As ListObject, cus As ListObject, prt As ListObject, buy As ListObject, acc As ListObject, shf As ListObject, tls As ListObject
    Set dev = GetLO("tblDevices"): Set cus = GetLO("tblCustomers"): Set prt = GetLO("tblParts")
    Set buy = GetLO("tblPurchases"): Set acc = GetLO("tblAcc"): Set shf = GetLO("tblShelves"): Set tls = GetLO("tblTools")

    SetColFormula dev, DC_TOTAL, IfCode(dev, "N(" & TRc(dev, DC_PARTSCOST) & ")+N(" & TRc(dev, DC_LABOR) & ")")
    SetColFormula dev, DC_BALANCE, IfCode(dev, "N(" & TRc(dev, DC_TOTAL) & ")-N(" & TRc(dev, DC_PAID) & ")")

    SetColFormula cus, CC_COUNT, IfCode(cus, "COUNTIF(" & SRc(dev, DC_CUSTCODE) & "," & TRc(cus, CC_CODE) & ")")
    SetColFormula cus, CC_PAID, IfCode(cus, "SUMIF(" & SRc(dev, DC_CUSTCODE) & "," & TRc(cus, CC_CODE) & "," & SRc(dev, DC_PAID) & ")")

    SetColFormula prt, PC_VALUE, IfCode(prt, "N(" & TRc(prt, PC_QTY) & ")*N(" & TRc(prt, PC_BUY) & ")")
    SetColFormula buy, BC_TOTAL, IfCode(buy, "N(" & TRc(buy, BC_QTY) & ")*N(" & TRc(buy, BC_PRICE) & ")")
    SetColFormula acc, AC_MONTH, IfCode(acc, "LEFT(" & TRc(acc, AC_JDATE) & ",7)")

    SetColFormula shf, SC_DEV, IfCode(shf, "COUNTIFS(" & SRc(dev, DC_SHELF) & "," & TRc(shf, SC_CODE) & "," & _
        SRc(dev, DC_STATUS) & ",""<>""&mk_st7," & SRc(dev, DC_STATUS) & ",""<>""&mk_st8)")
    SetColFormula shf, SC_PARTS, IfCode(shf, "COUNTIF(" & SRc(prt, PC_SHELF) & "," & TRc(shf, SC_CODE) & ")")
    SetColFormula shf, SC_TOOLS, IfCode(shf, "COUNTIF(" & SRc(tls, TC_LOC) & "," & TRc(shf, SC_CODE) & ")")
    SetColFormula shf, SC_FREE, IfCode(shf, "N(" & TRc(shf, SC_CAP) & ")-N(" & TRc(shf, SC_DEV) & ")-N(" & _
        TRc(shf, SC_PARTS) & ")-N(" & TRc(shf, SC_TOOLS) & ")")
End Sub

Private Sub Build_Validation()
    Dim lo As ListObject
    Set lo = GetLO("tblDevices")
    AddListDV lo.ListColumns(DC_TYPE).DataBodyRange, "mk_lstDevTypes", False
    AddListDV lo.ListColumns(DC_SHELF).DataBodyRange, "mk_lstShelves"
    AddListDV lo.ListColumns(DC_STATUS).DataBodyRange, "mk_lstStatus"
    AddListDV lo.ListColumns(DC_TECH).DataBodyRange, "mk_lstTech"
    AddListDV lo.ListColumns(DC_STAFF).DataBodyRange, "mk_lstTech"
    Set lo = GetLO("tblAcc")
    AddListDV lo.ListColumns(AC_TYPE).DataBodyRange, "mk_lstAccType"
    AddListDV lo.ListColumns(AC_CAT).DataBodyRange, "mk_lstAccCat"
    AddListDV lo.ListColumns(AC_METHOD).DataBodyRange, "mk_lstPay"
    Set lo = GetLO("tblParts")
    AddListDV lo.ListColumns(PC_CAT).DataBodyRange, "mk_lstPartCat"
    AddListDV lo.ListColumns(PC_SHELF).DataBodyRange, "mk_lstShelves"
    AddListDV lo.ListColumns(PC_UNIT).DataBodyRange, "mk_lstUnits", False
    Set lo = GetLO("tblPurchases")
    AddListDV lo.ListColumns(BC_METHOD).DataBodyRange, "mk_lstPay"
    Set lo = GetLO("tblTools")
    AddListDV lo.ListColumns(TC_CAT).DataBodyRange, "mk_lstToolCat", False
    AddListDV lo.ListColumns(TC_LOC).DataBodyRange, "mk_lstShelves", False
    AddListDV lo.ListColumns(TC_COND).DataBodyRange, "mk_lstToolCond"
    Set lo = GetLO("tblShelves")
    AddListDV lo.ListColumns(SC_TYPE).DataBodyRange, "mk_lstShelfType"
    ' purchase form
    AddListDV NC("mk_buy_Part").MergeArea, "mk_lstPartCodes"
    AddListDV NC("mk_buy_Method").MergeArea, "mk_lstPay"
    NC("mk_buy_Method").Value = ListItem("tblPay", 1)
End Sub

Private Sub Build_ConditionalFormats()
    Dim lo As ListObject, rng As Range, i As Long, f As String
    ' devices: whole row colored by status
    Set lo = GetLO("tblDevices")
    Set rng = lo.DataBodyRange
    rng.FormatConditions.Delete
    For i = 1 To STATUS_COUNT
        AddCF rng, "=" & lo.ListColumns(DC_STATUS).DataBodyRange.Cells(1, 1).Address(False, True) & "=mk_st" & i, StatusFill(i), StatusFont(i)
    Next i
    ' parts: red when quantity <= minimum
    Set lo = GetLO("tblParts")
    Set rng = lo.DataBodyRange
    rng.FormatConditions.Delete
    f = "=(" & lo.ListColumns(PC_CODE).DataBodyRange.Cells(1, 1).Address(False, True) & "<>"""")*(" & _
        lo.ListColumns(PC_QTY).DataBodyRange.Cells(1, 1).Address(False, True) & "<=" & _
        lo.ListColumns(PC_MIN).DataBodyRange.Cells(1, 1).Address(False, True) & ")"
    AddCF rng, f, StatusFill(ST_RETURNED), StatusFont(ST_RETURNED)
    ' tools: by condition
    Set lo = GetLO("tblTools")
    Set rng = lo.DataBodyRange
    rng.FormatConditions.Delete
    f = "=" & lo.ListColumns(TC_COND).DataBodyRange.Cells(1, 1).Address(False, True) & "=mk_cond"
    AddCF rng, f & "1", StatusFill(ST_READY), StatusFont(ST_READY)
    AddCF rng, f & "2", StatusFill(ST_APPROVAL), StatusFont(ST_APPROVAL)
    AddCF rng, f & "3", StatusFill(ST_RETURNED), StatusFont(ST_RETURNED)
    ' customers: duplicate phone numbers in red
    Set lo = GetLO("tblCustomers")
    Set rng = lo.ListColumns(CC_PHONE).DataBodyRange
    rng.FormatConditions.Delete
    With rng.FormatConditions.AddUniqueValues
        .DupeUnique = xlDuplicate
        .Interior.Color = StatusFill(ST_RETURNED)
        .Font.Color = StatusFont(ST_RETURNED)
        .Font.Bold = True
    End With
    ' shelves: red when full
    Set lo = GetLO("tblShelves")
    Set rng = lo.ListColumns(SC_FREE).DataBodyRange
    rng.FormatConditions.Delete
    AddCF rng, "=" & rng.Cells(1, 1).Address(False, True) & "<0", StatusFill(ST_RETURNED), StatusFont(ST_RETURNED)
End Sub

Private Sub Build_Intake()
    Dim ws As Worksheet, labels As Variant, nms As Variant, i As Long, r As Long
    Set ws = Sh("INTAKE")
    ws.Columns(2).ColumnWidth = 30
    ws.Columns(3).ColumnWidth = 42
    ws.Columns(4).ColumnWidth = 26
    ws.Columns(5).ColumnWidth = 4
    ws.Columns(6).ColumnWidth = 15
    ws.Columns(7).ColumnWidth = 29

    HeaderBand ws, SheetTitle("INTAKE"), 7
    BackButton ws
    Note ws.Range("C2:D2"), T("~28~27 ~A9~44~CC~2F Tab ~27~32 ~47~31 ~2E~27~46~47 ~28~47 ~2E~27~46~47 ~28~39~2F~CC ~28~31~48~CC~2F. ~2E~27~46~47^200C~47~27~CC ~33~2A~27~31~47^200C~2F~27~31 (*) ~27~44~32~27~45~CC ~47~33~2A~46~2F.")
    With ws.Range("B3:D3")
        .Merge
        .Value = T("~F1) ~34~45~27~31~47 ~2A~44~41~46 ~47~45~31~27~47 ~31~27 ~28~46~48~CC~33~CC~2F ~48 ^00AB~CC~27~41~2A~46 ~45~34~2A~31~CC^00BB ~31~27 ~28~32~46~CC~2F   ~F2) ~28~42~CC~47 ~41~31~45 ~31~27 ~7E~31 ~A9~46~CC~2F   ~F3) ^00AB~2B~28~2A ~2F~33~2A~AF~27~47^00BB")
        .Font.Bold = True
        .Font.Color = CLR("teal")
        .Font.Size = 10
        .WrapText = True
    End With
    ws.Rows(3).RowHeight = 34

    labels = Array(T("~2A~44~41~46 ~47~45~31~27~47 ~45~34~2A~31~CC *"), T("~46~27~45 ~48 ~46~27~45 ~2E~27~46~48~27~2F~AF~CC ~45~34~2A~31~CC *"), T("~46~48~39 ~2F~33~2A~AF~27~47 *"), T("~28~31~46~2F"), T("~45~2F~44"), T("~34~45~27~31~47 ~33~31~CC~27~44"), _
        T("~27~CC~31~27~2F ~27~39~44~27~45^200C~34~2F~47 ~2A~48~33~37 ~45~34~2A~31~CC *"), T("~44~48~27~32~45 ~47~45~31~27~47 (~34~27~31~98~31~0C ~A9~27~28~44~0C ~2C~39~28~47~0C ~31~CC~45~48~2A...)"), T("~48~36~39~CC~2A ~38~27~47~31~CC ~2F~33~2A~AF~27~47"), T("~28~31~22~48~31~2F ~47~32~CC~46~47 (~2A~48~45~27~46)"), _
        T("~28~CC~39~27~46~47 ~2F~31~CC~27~41~2A~CC (~2A~48~45~27~46)"), T("~31~48~34 ~7E~31~2F~27~2E~2A ~28~CC~39~27~46~47"), T("~42~41~33~47 ~46~AF~47~2F~27~31~CC ~2F~33~2A~AF~27~47"), T("~7E~30~CC~31~46~2F~47 (~A9~27~31~45~46~2F)"), T("~2A~48~36~CC~2D~27~2A"))
    nms = Array("mk_in_Phone", "mk_in_Name", "mk_in_Type", "mk_in_Brand", "mk_in_Model", "mk_in_Serial", "mk_in_Problem", _
        "mk_in_Access", "mk_in_Cond", "mk_in_Estimate", "mk_in_Deposit", "mk_in_PayMethod", "mk_in_Shelf", "mk_in_Staff", "mk_in_Notes")
    For i = 0 To UBound(labels)
        r = 5 + i
        ws.Rows(r).RowHeight = 24
        FormLabel ws.Cells(r, 2), CStr(labels(i))
        StyleInput ws.Cells(r, 3), True
        NameCell CStr(nms(i)), ws.Cells(r, 3)
    Next i
    ws.Rows(11).RowHeight = 48
    ws.Range("C11").WrapText = True
    ws.Range("C5").NumberFormat = "@"
    ws.Range("C5").Font.Size = 14
    ws.Range("C5").Font.Bold = True
    ws.Range("C10").NumberFormat = "@"
    ws.Range("C14:C15").NumberFormat = "#,##0"
    AddListDV ws.Range("C7"), "mk_lstDevTypes", False
    AddListDV ws.Range("C16"), "mk_lstPay"
    AddListDV ws.Range("C17"), "mk_lstShelves"
    AddListDV ws.Range("C18"), "mk_lstTech", False
    ws.Range("C16").Value = ListItem("tblPay", 1)

    With ws.Range("D6:D8")
        .Merge
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Color = CLR("teal")
        .Font.Bold = True
    End With
    NameCell "mk_in_CustInfo", ws.Range("D6")

    ' action buttons
    AddBtn ws, ws.Range("D5").Left + 4, ws.Range("D5").Top + 1, ws.Range("D5").Width - 8, 22, T("~CC~27~41~2A~46 ~45~34~2A~31~CC"), "Intake_FindCustomer", CLR("blue"), , 10
    ws.Rows(21).RowHeight = 24: ws.Rows(22).RowHeight = 24
    AddBtn ws, ws.Range("B21").Left, ws.Range("B21").Top, ws.Range("B21").Width + ws.Range("C21").Width, 46, T("~2B~28~2A ~2F~33~2A~AF~27~47"), "Intake_Save", CLR("green"), , 16
    AddBtn ws, ws.Range("D21").Left + 4, ws.Range("D21").Top, ws.Range("D21").Width - 8, 22, T("~7E~27~A9 ~A9~31~2F~46 ~41~31~45"), "Intake_Clear", CLR("gray"), , 10
    AddBtn ws, ws.Range("D21").Left + 4, ws.Range("D22").Top + 2, ws.Range("D21").Width - 8, 22, T("~86~27~7E ~31~33~CC~2F"), "Intake_PrintReceipt", CLR("navy2"), , 10

    ws.Rows(24).RowHeight = 40
    FormLabel ws.Range("B24"), T("~22~2E~31~CC~46 ~A9~2F ~2B~28~2A^200C~34~2F~47:")
    With ws.Range("C24")
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = CLR("teal")
        .HorizontalAlignment = xlCenter
        .NumberFormat = "@"
        .Interior.Color = CLR("tealLight")
    End With
    NameCell "mk_in_LastCode", ws.Range("C24")

    Build_Receipt ws
End Sub

' Small printable receipt on the intake sheet (columns F:G).
Private Sub Build_Receipt(ByVal ws As Worksheet)
    Dim labels As Variant, nms As Variant, i As Long, r As Long
    With ws.Range("F3:G3")
        .Merge
        .Formula = "=mk_ShopName"
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("F4:G4")
        .Merge
        .Formula = "=mk_ShopSlogan"
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("F5:G5")
        .Merge
        .Formula = T("=""~2A~44~41~46: ""&mk_ShopPhone&""  -  ""&mk_ShopAddr")
        .Font.Size = 8
        .WrapText = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(5).RowHeight = 28
    With ws.Range("F7:G7")
        .Merge
        .Value = T("~31~33~CC~2F ~7E~30~CC~31~34 ~2F~33~2A~AF~27~47")
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(230, 230, 230)
    End With
    NameCell "mk_rc_Title", ws.Range("F7")

    labels = Array(T("~A9~2F ~2F~33~2A~AF~27~47"), T("~2A~27~31~CC~2E ~48 ~33~27~39~2A"), T("~45~34~2A~31~CC"), T("~2A~44~41~46"), T("~2F~33~2A~AF~27~47"), T("~27~CC~31~27~2F ~27~39~44~27~45^200C~34~2F~47"), T("~44~48~27~32~45 ~47~45~31~27~47"), _
        T("~48~36~39~CC~2A"), T("~28~31~22~48~31~2F ~47~32~CC~46~47"), T("~2C~45~39 ~47~32~CC~46~47"), T("~7E~31~2F~27~2E~2A~CC"), T("~45~27~46~2F~47"), T("~AF~27~31~27~46~2A~CC ~2A~27"))
    nms = Array("mk_rc_Code", "mk_rc_Date", "mk_rc_Cust", "mk_rc_Phone", "mk_rc_Device", "mk_rc_Problem", "mk_rc_Access", _
        "mk_rc_Status", "mk_rc_Estimate", "mk_rc_Total", "mk_rc_Paid", "mk_rc_Balance", "mk_rc_Warranty")
    For i = 0 To UBound(labels)
        r = 8 + i
        ws.Cells(r, 6).Value = labels(i)
        ws.Cells(r, 6).Font.Bold = True
        ws.Cells(r, 6).Font.Size = 9
        ws.Cells(r, 7).Font.Size = 9
        ws.Cells(r, 7).WrapText = True
        ws.Cells(r, 7).HorizontalAlignment = xlRight
        NameCell CStr(nms(i)), ws.Cells(r, 7)
    Next i
    ws.Range("G8").Font.Size = 14
    ws.Range("G8").Font.Bold = True
    ws.Range("G8,G11").NumberFormat = "@"
    ws.Range("G16:G19").NumberFormat = "#,##0"
    ws.Rows(13).RowHeight = 34
    With ws.Range("F22:G23")
        .Merge
        .WrapText = True
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
    End With
    NameCell "mk_rc_Footer", ws.Range("F22")
    ws.Range("F3:G23").BorderAround xlContinuous, xlThin, , RGB(120, 120, 120)
    With ws.Range("F8:G20").Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Color = RGB(180, 180, 180)
    End With
    Note ws.Range("F24:G24"), T("~7E~CC~34^200C~46~45~27~CC~34 ~31~33~CC~2F - ~28~27 ~2F~A9~45~47 ^00AB~86~27~7E ~31~33~CC~2F^00BB ~86~27~7E ~45~CC^200C~34~48~2F.")
    SetupReceiptPage ws
End Sub

Private Sub SetupReceiptPage(ByVal ws As Worksheet)
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = "$F$3:$G$23"
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.CentimetersToPoints(0.3)
        .RightMargin = Application.CentimetersToPoints(0.3)
        .TopMargin = Application.CentimetersToPoints(0.3)
        .BottomMargin = Application.CentimetersToPoints(0.3)
        .HeaderMargin = 0
        .FooterMargin = 0
        .CenterHorizontally = True
        If InStr(1, UCase$(CStr(NC("mk_ReceiptSize").Value)), "A5") > 0 Then .PaperSize = xlPaperA5
    End With
    Application.PrintCommunication = True
    Err.Clear
End Sub

Private Sub Build_Search()
    Dim ws As Worksheet, heads As Variant, widths As Variant, i As Long
    Set ws = Sh("SEARCH")
    widths = Array(16, 12, 22, 14, 24, 24, 14, 14)
    For i = 0 To UBound(widths)
        ws.Columns(2 + i).ColumnWidth = widths(i)
    Next i
    HeaderBand ws, SheetTitle("SEARCH"), 9
    BackButton ws

    ws.Rows(4).RowHeight = 26: ws.Rows(5).RowHeight = 26
    FormLabel ws.Range("B4:C4"), T("~39~28~27~31~2A ~2C~33~2A~2C~48")
    StyleInput ws.Range("D4:F4"), True
    ws.Range("D4").Font.Size = 13
    ws.Range("D4").NumberFormat = "@"
    NameCell "mk_srch_Text", ws.Range("D4")
    FormLabel ws.Range("B5:C5"), T("~2C~33~2A~2C~48 ~28~31 ~27~33~27~33")
    StyleInput ws.Range("D5:F5"), True
    AddListDV ws.Range("D5:F5"), "mk_lstSearchBy"
    ws.Range("D5").Value = ListItem("tblSearchBy", 1)
    NameCell "mk_srch_By", ws.Range("D5")
    AddBtn ws, ws.Range("G4").Left + 4, ws.Range("G4").Top + 1, ws.Range("G4").Width - 8, 50, T("~2C~33~2A~2C~48"), "Search_Run", CLR("purple"), , 13
    AddBtn ws, ws.Range("H4").Left + 4, ws.Range("H4").Top + 1, ws.Range("H4").Width + ws.Range("I4").Width - 8, 24, T("~7E~27~A9 ~A9~31~2F~46"), "Search_Clear", CLR("gray")
    With ws.Range("B6:I6")
        .Merge
        .Font.Bold = True
        .Font.Color = CLR("purple")
    End With
    NameCell "mk_srch_Info", ws.Range("B6")
    Note ws.Range("B7:I7"), T("~46~27~45 ~45~34~2A~31~CC ~48 ~2A~44~41~46 ~28~27 ~28~2E~34~CC ~27~32 ~39~28~27~31~2A ~47~45 ~7E~CC~2F~27 ~45~CC^200C~34~48~46~2F (~45~2B~44~27~4B ~41~42~37 ~86~46~2F ~31~42~45 ~22~2E~31 ~2A~44~41~46). ~31~48~CC ~A9~2F ~2F~33~2A~AF~27~47 ~A9~44~CC~A9 ~A9~46~CC~2F ~2A~27 ~28~47 ~31~2F~CC~41 ~22~46 ~2F~31 ~28~31~AF~47 ^00AB~2F~33~2A~AF~27~47^200C~47~27^00BB ~28~31~48~CC~2F.")

    heads = Array(T("~A9~2F ~2F~33~2A~AF~27~47"), T("~2A~27~31~CC~2E ~7E~30~CC~31~34"), T("~45~34~2A~31~CC"), T("~2A~44~41~46"), T("~2F~33~2A~AF~27~47"), T("~48~36~39~CC~2A"), T("~2C~45~39 ~47~32~CC~46~47"), T("~45~27~46~2F~47"))
    For i = 0 To UBound(heads)
        ws.Cells(SEARCH_FIRST_ROW - 1, 2 + i).Value = heads(i)
    Next i
    With ws.Range(ws.Cells(SEARCH_FIRST_ROW - 1, 2), ws.Cells(SEARCH_FIRST_ROW - 1, 9))
        .Interior.Color = CLR("purple")
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .RowHeight = 26
    End With
End Sub

' Summary block, monthly table, chart and daily table on the accounting sheet.
Private Sub Build_AccSummary()
    Dim ws As Worksheet, acc As ListObject, dev As ListObject, i As Long, r As Long
    Dim amt As String, typ As String, mon As String, cat As String, jd As String, gd As String, y As String, m As String
    Dim co As ChartObject
    Set ws = Sh("ACC")
    Set acc = GetLO("tblAcc"): Set dev = GetLO("tblDevices")
    amt = SRc(acc, AC_AMOUNT): typ = SRc(acc, AC_TYPE): mon = SRc(acc, AC_MONTH)
    cat = SRc(acc, AC_CAT): jd = SRc(acc, AC_JDATE): gd = SRc(acc, AC_GDATE)

    ws.Columns(11).ColumnWidth = 3
    ws.Columns(12).ColumnWidth = 30
    ws.Columns(13).ColumnWidth = 15
    ws.Columns(14).ColumnWidth = 15
    ws.Columns(15).ColumnWidth = 15
    ws.Columns(16).ColumnWidth = 3

    SectionTitle ws.Range("L5:O5"), T("~2E~44~27~35~47 ~2D~33~27~28^200C~47~27 (~2A~48~45~27~46)")
    FormLabel ws.Range("L6"), T("~45~27~47 ~45~48~31~2F ~46~38~31 (~45~2B~44~27~4B 1405/07) - ~2E~27~44~CC = ~45~27~47 ~2C~27~31~CC")
    StyleInput ws.Range("M6"), True
    ws.Range("M6").NumberFormat = "@"
    ws.Range("M6").HorizontalAlignment = xlCenter
    NameCell "mk_acc_Month", ws.Range("M6")
    FormLabel ws.Range("L7"), T("~45~27~47 ~46~45~27~CC~34 ~2F~27~2F~47 ~34~2F~47")
    ws.Range("M7").Formula = "=IF(TRIM(mk_acc_Month)="""",LEFT(JalaliDate(TODAY()),7),TRIM(mk_acc_Month))"
    ws.Range("M7").Font.Bold = True
    ws.Range("M7").HorizontalAlignment = xlCenter
    NameCell "mk_acc_MonthEff", ws.Range("M7")

    TableHead ws.Range("L9:O9"), Array("", T("~2F~31~22~45~2F"), T("~47~32~CC~46~47"), T("~33~48~2F ~2E~27~44~35"))
    ws.Range("L10").Value = T("~27~45~31~48~32")
    ws.Range("M10").Formula = "=SUMIFS(" & amt & "," & typ & ",mk_income," & gd & ",TODAY())"
    ws.Range("N10").Formula = "=SUMIFS(" & amt & "," & typ & ",mk_expense," & gd & ",TODAY())"
    ws.Range("L11").Value = T("~45~27~47 ~27~46~2A~2E~27~28^200C~34~2F~47")
    ws.Range("M11").Formula = "=SUMPRODUCT(--(" & mon & "=mk_acc_MonthEff),--(" & typ & "=mk_income)," & amt & ")"
    ws.Range("N11").Formula = "=SUMPRODUCT(--(" & mon & "=mk_acc_MonthEff),--(" & typ & "=mk_expense)," & amt & ")"
    ws.Range("L12").Value = T("~27~32 ~27~28~2A~2F~27 ~2A~27 ~27~45~31~48~32")
    ws.Range("M12").Formula = "=SUMIF(" & typ & ",mk_income," & amt & ")"
    ws.Range("N12").Formula = "=SUMIF(" & typ & ",mk_expense," & amt & ")"
    For r = 10 To 12
        ws.Cells(r, 15).Formula = "=M" & r & "-N" & r
    Next r
    BodyBox ws.Range("L10:O12")
    ws.Range("O10:O12").Font.Bold = True

    FormLabel ws.Range("L14"), T("~37~44~28 ~27~32 ~45~34~2A~31~CC~27~46 (~45~27~46~2F~47 ~2F~33~2A~AF~27~47^200C~47~27)")
    ws.Range("M14").Formula = "=SUMIF(" & SRc(dev, DC_BALANCE) & ","">0"")"
    ws.Range("M14").Font.Bold = True
    ws.Range("M14").Font.Color = CLR("red")
    BodyBox ws.Range("M14")

    ' by category
    TableHead ws.Range("L16:N16"), Array(T("~2F~33~2A~47"), T("~27~32 ~27~28~2A~2F~27"), T("~45~27~47 ~27~46~2A~2E~27~28^200C~34~2F~47"))
    For i = 1 To 12
        r = 16 + i
        ws.Cells(r, 12).Formula = "=IFERROR(INDEX(mk_lstAccCat," & i & "),"""")"
        ws.Cells(r, 13).Formula = "=IF($L" & r & "="""","""",SUMIF(" & cat & ",$L" & r & "," & amt & "))"
        ws.Cells(r, 14).Formula = "=IF($L" & r & "="""","""",SUMPRODUCT(--(" & cat & "=$L" & r & "),--(" & mon & "=mk_acc_MonthEff)," & amt & "))"
    Next i
    BodyBox ws.Range("L17:N28")

    ' last 12 months (oldest first)
    y = "VALUE(LEFT(mk_acc_MonthEff,4))": m = "VALUE(MID(mk_acc_MonthEff,6,2))"
    TableHead ws.Range("L30:O30"), Array(T("~45~27~47"), T("~2F~31~22~45~2F"), T("~47~32~CC~46~47"), T("~33~48~2F"))
    For i = 0 To 11
        r = 31 + i
        ws.Cells(r, 12).Formula = "=IFERROR(INT((" & y & "*12+" & m & "-1-" & (11 - i) & ")/12)&""/""&TEXT(MOD(" & y & "*12+" & m & "-1-" & (11 - i) & ",12)+1,""00""),"""")"
        ws.Cells(r, 13).Formula = "=IF($L" & r & "="""",0,SUMPRODUCT(--(" & mon & "=$L" & r & "),--(" & typ & "=mk_income)," & amt & "))"
        ws.Cells(r, 14).Formula = "=IF($L" & r & "="""",0,SUMPRODUCT(--(" & mon & "=$L" & r & "),--(" & typ & "=mk_expense)," & amt & "))"
        ws.Cells(r, 15).Formula = "=M" & r & "-N" & r
    Next i
    BodyBox ws.Range("L31:O42")
    ws.Range("L31:L42").HorizontalAlignment = xlCenter

    ' days of the selected month
    TableHead ws.Range("L45:O45"), Array(T("~31~48~32 (~45~27~47 ~27~46~2A~2E~27~28^200C~34~2F~47)"), T("~2F~31~22~45~2F"), T("~47~32~CC~46~47"), T("~33~48~2F"))
    For i = 1 To 31
        r = 45 + i
        ws.Cells(r, 12).Formula = "=mk_acc_MonthEff&""/" & Format$(i, "00") & """"
        ws.Cells(r, 13).Formula = "=SUMPRODUCT(--(" & jd & "=$L" & r & "),--(" & typ & "=mk_income)," & amt & ")"
        ws.Cells(r, 14).Formula = "=SUMPRODUCT(--(" & jd & "=$L" & r & "),--(" & typ & "=mk_expense)," & amt & ")"
        ws.Cells(r, 15).Formula = "=M" & r & "-N" & r
    Next i
    BodyBox ws.Range("L46:O76")
    ws.Range("L46:L76").HorizontalAlignment = xlCenter
    ws.Range("M10:O76").NumberFormat = "#,##0;[Red]-#,##0"

    ' chart: monthly income vs expenses
    Set co = ws.ChartObjects.Add(ws.Range("Q5").Left, ws.Range("Q5").Top, 640, 320)
    co.Name = "mk_chtMonthly"
    With co.Chart
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        With .SeriesCollection.NewSeries
            .Name = "='" & ws.Name & "'!$M$30"
            .Values = ws.Range("M31:M42")
            .XValues = ws.Range("L31:L42")
        End With
        With .SeriesCollection.NewSeries
            .Name = "='" & ws.Name & "'!$N$30"
            .Values = ws.Range("N31:N42")
            .XValues = ws.Range("L31:L42")
        End With
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = T("~2F~31~22~45~2F ~48 ~47~32~CC~46~47 ~F1~F2 ~45~27~47 ~27~2E~CC~31")
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(76, 175, 80)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(229, 115, 115)
        On Error Resume Next
        .ChartArea.Format.TextFrame2.TextRange.Font.Name = FONT_NAME
        .ChartArea.Font.Name = FONT_NAME
        .Axes(xlCategory).TickLabels.Font.Size = 8
        On Error GoTo 0
    End With
End Sub

Private Sub Build_Dashboard()
    Dim ws As Worksheet, dev As ListObject, acc As ListObject, prt As ListObject, i As Long, c As Long
    Dim caps As Variant, macs As Variant, clrs As Variant, x As Double, y As Double, w As Double, stc As String
    Set ws = Sh("DASH")
    Set dev = GetLO("tblDevices"): Set acc = GetLO("tblAcc"): Set prt = GetLO("tblParts")

    For c = 2 To 16
        If (c - 2) Mod 3 = 2 Then ws.Columns(c).ColumnWidth = 2 Else ws.Columns(c).ColumnWidth = 12.5
    Next c

    ws.Rows(1).RowHeight = 52
    With ws.Range("B1:O1")
        .Merge
        .Formula = "=mk_ShopName&""      MOKONIC"""
        .Interior.Color = CLR("navy")
        .Font.Color = vbWhite
        .Font.Size = 26
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(2).RowHeight = 24
    With ws.Range("B2:O2")
        .Merge
        .Formula = "=mk_ShopSlogan"
        .Interior.Color = CLR("navy2")
        .Font.Color = vbWhite
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
    End With

    ws.Rows(3).RowHeight = 30
    FormLabel ws.Range("B3:C3"), T("~27~45~31~48~32:")
    With ws.Range("D3:F3")
        .Merge
        .Formula = "=JalaliDate(TODAY())"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = CLR("navy")
        .HorizontalAlignment = xlCenter
    End With
    NameCell "mk_dash_Today", ws.Range("D3")
    FormLabel ws.Range("H3:I3"), T("~45~27~47 ~2C~27~31~CC:")
    With ws.Range("J3:K3")
        .Merge
        .Formula = "=LEFT(mk_dash_Today,7)"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = CLR("navy")
        .HorizontalAlignment = xlCenter
    End With
    NameCell "mk_dash_Month", ws.Range("J3")
    AddBtn ws, ws.Range("N3").Left, ws.Range("N3").Top + 3, ws.Range("N3:O3").Width, 24, T("~28~31~48~32~31~33~27~46~CC ~2F~27~34~28~48~31~2F"), "Dash_Refresh", CLR("navy2"), , 9

    ' navigation buttons
    caps = Array(T("~7E~30~CC~31~34 ~33~31~CC~39 ~2F~33~2A~AF~27~47"), T("~2F~33~2A~AF~27~47^200C~47~27"), T("~2C~33~2A~2C~48"), T("~45~34~2A~31~CC~27~46"), T("~2D~33~27~28~2F~27~31~CC"), _
                 T("~27~46~28~27~31 ~42~37~39~27~2A"), T("~2E~31~CC~2F ~42~37~39~27~2A"), T("~27~28~32~27~31"), T("~42~41~33~47^200C~47~27"), T("~2A~46~38~CC~45~27~2A"))
    macs = Array("Go_Intake", "Go_Devices", "Go_Search", "Go_Customers", "Go_Accounting", _
                 "Go_Stock", "Go_Purchases", "Go_Tools", "Go_Shelves", "Go_Settings")
    clrs = Array(CLR("green"), CLR("teal"), CLR("purple"), CLR("blue"), CLR("gold"), _
                 CLR("orange"), CLR("brown"), CLR("gray"), CLR("olive"), CLR("dark"))
    For i = 5 To 8
        ws.Rows(i).RowHeight = 21
    Next i
    For i = 0 To 9
        c = 2 + 3 * (i Mod 5)
        x = ws.Cells(5, c).Left
        w = ws.Range(ws.Cells(5, c), ws.Cells(5, c + 1)).Width
        If i < 5 Then y = ws.Range("B5").Top Else y = ws.Range("B7").Top + 2
        AddBtn ws, x, y, w, 38, CStr(caps(i)), CStr(macs(i)), CLng(clrs(i)), , 12
    Next i

    ' KPI cards
    stc = SRc(dev, DC_STATUS)
    SectionTitle ws.Range("B10:O10"), T("~27~45~31~48~32 ~48 ~48~36~39~CC~2A ~41~39~44~CC")
    Card ws, 11, 2, T("~2F~33~2A~AF~27~47^200C~47~27~CC ~2F~27~2E~44 ~2A~39~45~CC~31~AF~27~47"), "=COUNTA(" & SRc(dev, DC_CODE) & ")-COUNTIF(" & stc & ",mk_st7)-COUNTIF(" & stc & ",mk_st8)", "#,##0", CLR("teal")
    Card ws, 11, 5, T("~7E~30~CC~31~34 ~27~45~31~48~32"), "=COUNTIF(" & SRc(dev, DC_GDATE) & ",TODAY())", "#,##0", CLR("blue")
    Card ws, 11, 8, T("~22~45~27~2F~47 ~2A~2D~48~CC~44"), "=COUNTIF(" & stc & ",mk_st6)", "#,##0", CLR("green")
    Card ws, 11, 11, T("~2F~31~22~45~2F ~27~45~31~48~32 (~2A~48~45~27~46)"), "=SUMIFS(" & SRc(acc, AC_AMOUNT) & "," & SRc(acc, AC_TYPE) & ",mk_income," & SRc(acc, AC_GDATE) & ",TODAY())", "#,##0", CLR("gold")
    Card ws, 11, 14, T("~37~44~28 ~27~32 ~45~34~2A~31~CC~27~46 (~2A~48~45~27~46)"), "=SUMIF(" & SRc(dev, DC_BALANCE) & ","">0"")", "#,##0", CLR("red")
    SectionTitle ws.Range("B15:O15"), T("~27~CC~46 ~45~27~47 (~34~45~33~CC) ~48 ~27~46~28~27~31")
    Card ws, 16, 2, T("~2F~31~22~45~2F ~27~CC~46 ~45~27~47"), "=SUMPRODUCT(--(" & SRc(acc, AC_MONTH) & "=mk_dash_Month),--(" & SRc(acc, AC_TYPE) & "=mk_income)," & SRc(acc, AC_AMOUNT) & ")", "#,##0", CLR("green")
    Card ws, 16, 5, T("~47~32~CC~46~47 ~27~CC~46 ~45~27~47"), "=SUMPRODUCT(--(" & SRc(acc, AC_MONTH) & "=mk_dash_Month),--(" & SRc(acc, AC_TYPE) & "=mk_expense)," & SRc(acc, AC_AMOUNT) & ")", "#,##0", CLR("red")
    Card ws, 16, 8, T("~33~48~2F ~27~CC~46 ~45~27~47"), "=B17-E17", "#,##0;[Red]-#,##0", CLR("navy")
    Card ws, 16, 11, T("~42~37~39~27~2A ~A9~45^200C~45~48~2C~48~2F~CC"), "=SUMPRODUCT((" & SRc(prt, PC_CODE) & "<>"""")*(" & SRc(prt, PC_QTY) & "<=" & SRc(prt, PC_MIN) & "))", "#,##0", CLR("orange")
    Card ws, 16, 14, T("~27~31~32~34 ~27~46~28~27~31 (~2A~48~45~27~46)"), "=SUM(" & SRc(prt, PC_VALUE) & ")", "#,##0", CLR("brown")

    ' devices by status
    SectionTitle ws.Range("B20:F20"), T("~2A~39~2F~27~2F ~2F~33~2A~AF~27~47^200C~47~27 ~28~31 ~27~33~27~33 ~48~36~39~CC~2A")
    For i = 1 To STATUS_COUNT
        With ws.Range(ws.Cells(20 + i, 2), ws.Cells(20 + i, 5))
            .Merge
            .Formula = "=mk_st" & i
            .Interior.Color = StatusFill(i)
            .Font.Color = StatusFont(i)
            .Font.Bold = True
        End With
        With ws.Cells(20 + i, 6)
            .Formula = "=COUNTIF(" & stc & ",mk_st" & i & ")"
            .Interior.Color = StatusFill(i)
            .Font.Color = StatusFont(i)
            .Font.Bold = True
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
        End With
        ws.Rows(20 + i).RowHeight = 21
    Next i
    With ws.Range("B29:E29")
        .Merge
        .Value = T("~2C~45~39 ~A9~44 ~2F~33~2A~AF~27~47^200C~47~27~CC ~2B~28~2A^200C~34~2F~47")
        .Font.Bold = True
    End With
    ws.Range("F29").Formula = "=COUNTA(" & SRc(dev, DC_CODE) & ")"
    ws.Range("F29").Font.Bold = True
    ws.Range("F29").HorizontalAlignment = xlCenter
    ws.Range("B21:F29").Borders.Color = RGB(255, 255, 255)
    ws.Range("B29:F29").Interior.Color = RGB(242, 242, 242)

    ' low stock list (filled by Dash_Refresh)
    SectionTitle ws.Range("H20:O20"), T("~42~37~39~27~2A ~A9~45^200C~45~48~2C~48~2F~CC (~45~48~2C~48~2F~CC ~A9~45~2A~31 ~CC~27 ~45~33~27~48~CC ~2D~2F~27~42~44)")
    TableHead ws.Range("H21:O21"), Array(T("~A9~2F"), T("~46~27~45 ~42~37~39~47"), "", "", T("~45~48~2C~48~2F~CC"), "", T("~2D~2F~27~42~44"), T("~42~41~33~47"))
    ws.Range("I21:K21").Merge
    ws.Range("L21:M21").Merge
    For i = DASH_LOW_FIRST To DASH_LOW_LAST
        ws.Range(ws.Cells(i, 9), ws.Cells(i, 11)).Merge
        ws.Range(ws.Cells(i, 12), ws.Cells(i, 13)).Merge
    Next i
    With ws.Range(ws.Cells(DASH_LOW_FIRST, 8), ws.Cells(DASH_LOW_LAST, 15))
        .Borders(xlInsideHorizontal).Color = RGB(230, 230, 230)
        .Borders(xlEdgeBottom).Color = RGB(200, 200, 200)
    End With
    ws.Range(ws.Cells(DASH_LOW_FIRST, 8), ws.Cells(DASH_LOW_LAST, 8)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(DASH_LOW_FIRST, 12), ws.Cells(DASH_LOW_LAST, 15)).HorizontalAlignment = xlCenter
    Note ws.Range("B35:O35"), T("~27~39~2F~27~2F ~27~CC~46 ~35~41~2D~47 ~28~27 ~47~31 ~2A~3A~CC~CC~31 ~2E~48~2F~A9~27~31 ~28~47^200C~31~48~32 ~45~CC^200C~34~48~46~2F. ~41~47~31~33~2A ~42~37~39~27~2A ~A9~45^200C~45~48~2C~48~2F~CC ~28~27 ~2F~A9~45~47 ^00AB~28~31~48~32~31~33~27~46~CC ~2F~27~34~28~48~31~2F^00BB ~CC~27 ~47~46~AF~27~45 ~28~27~32 ~34~2F~46 ~41~27~CC~44 ~28~47^200C~31~48~32 ~45~CC^200C~34~48~2F.")
End Sub

' One KPI card: title on row r, big value on rows r+1..r+2, two columns wide.
Private Sub Card(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal title As String, ByVal f As String, ByVal numFmt As String, ByVal color As Long)
    ws.Rows(r).RowHeight = 24
    ws.Rows(r + 1).RowHeight = 22
    ws.Rows(r + 2).RowHeight = 22
    With ws.Range(ws.Cells(r, c), ws.Cells(r, c + 1))
        .Merge
        .Value = title
        .Interior.Color = color
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range(ws.Cells(r + 1, c), ws.Cells(r + 2, c + 1))
        .Merge
        .Interior.Color = RGB(250, 250, 250)
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = color
        .HorizontalAlignment = xlCenter
        .NumberFormat = numFmt
        .BorderAround xlContinuous, xlThin, , color
    End With
    ws.Cells(r + 1, c).Formula = f
End Sub

'==============================================================================
'  NAVIGATION
'==============================================================================

Public Sub Go_Dashboard()
    On Error GoTo EH
    BeginBusy
    Dash_RefreshCore
    EndBusy
    Sh("DASH").Activate
    ActiveWindow.ScrollRow = 1
    Sh("DASH").Range("A1").Select
    Exit Sub
EH:
    FailWith "Go_Dashboard"
End Sub

Public Sub Go_Intake()
    GoSheet "INTAKE", "mk_in_Phone"
End Sub

Public Sub Go_Devices()
    GoSheet "DEV"
End Sub

Public Sub Go_Search()
    GoSheet "SEARCH", "mk_srch_Text"
End Sub

Public Sub Go_Customers()
    GoSheet "CUST"
End Sub

Public Sub Go_Accounting()
    GoSheet "ACC"
End Sub

Public Sub Go_Stock()
    GoSheet "STOCK"
End Sub

Public Sub Go_Purchases()
    GoSheet "BUY", "mk_buy_Supp"
End Sub

Public Sub Go_Tools()
    GoSheet "TOOLS"
End Sub

Public Sub Go_Shelves()
    GoSheet "SHELF"
End Sub

Public Sub Go_Settings()
    GoSheet "SET"
End Sub

Private Sub GoSheet(ByVal key As String, Optional ByVal cellName As String = "")
    On Error GoTo EH
    Sh(key).Activate
    If Len(cellName) > 0 Then
        NC(cellName).Select
    Else
        ActiveWindow.ScrollRow = 1
    End If
    Exit Sub
EH:
    FailWith "GoSheet"
End Sub

' Runs automatically when the workbook is opened (macros enabled).
Public Sub Auto_Open()
    On Error Resume Next
    If Not MokonicSheetsExist() Then Exit Sub
    BeginBusy
    Dash_RefreshCore
    EndBusy
    Sh("DASH").Activate
    Sh("DASH").Range("A1").Select
End Sub

'==============================================================================
'  DASHBOARD
'==============================================================================

Public Sub Dash_Refresh()
    On Error GoTo EH
    BeginBusy
    Dash_RefreshCore
    EndBusy
    Exit Sub
EH:
    FailWith "Dash_Refresh"
End Sub

' Rewrites the low-stock list on the dashboard.
Private Sub Dash_RefreshCore()
    Dim ws As Worksheet, lo As ListObject, rw As Range, r As Long, n As Long, maxShow As Long
    Set ws = Sh("DASH")
    Set lo = GetLO("tblParts")
    ws.Range(ws.Cells(DASH_LOW_FIRST, 8), ws.Cells(DASH_LOW_LAST, 15)).ClearContents
    ws.Range(ws.Cells(DASH_LOW_FIRST, 8), ws.Cells(DASH_LOW_LAST, 15)).Interior.Pattern = xlNone
    maxShow = DASH_LOW_LAST - DASH_LOW_FIRST     ' last row is kept for "... and N more"
    r = DASH_LOW_FIRST
    If Not lo.DataBodyRange Is Nothing Then
        For Each rw In lo.DataBodyRange.Rows
            If Len(CStr(rw.Cells(1, PC_CODE).Value)) > 0 Then
                If Num(rw.Cells(1, PC_QTY).Value) <= Num(rw.Cells(1, PC_MIN).Value) Then
                    n = n + 1
                    If n <= maxShow Then
                        ws.Cells(r, 8).Value = rw.Cells(1, PC_CODE).Value
                        ws.Cells(r, 9).Value = rw.Cells(1, PC_NAME).Value
                        ws.Cells(r, 12).Value = Num(rw.Cells(1, PC_QTY).Value)
                        ws.Cells(r, 14).Value = Num(rw.Cells(1, PC_MIN).Value)
                        ws.Cells(r, 15).Value = rw.Cells(1, PC_SHELF).Value
                        ws.Range(ws.Cells(r, 8), ws.Cells(r, 15)).Interior.Color = StatusFill(ST_RETURNED)
                        r = r + 1
                    End If
                End If
            End If
        Next rw
    End If
    If n = 0 Then
        ws.Cells(DASH_LOW_FIRST, 9).Value = T("~45~48~2C~48~2F~CC ~47~45~47 ~42~37~39~27~2A ~A9~27~41~CC ~27~33~2A.")
    ElseIf n > maxShow Then
        ws.Cells(DASH_LOW_LAST, 9).Value = T("... ~48 ") & (n - maxShow) & T(" ~42~37~39~47 ~2F~CC~AF~31 (~28~31~AF~47 ~27~46~28~27~31 ~31~27 ~28~28~CC~46~CC~2F)")
    End If
    Application.Calculate
End Sub

'==============================================================================
'  INTAKE (fast device registration)
'==============================================================================

Public Sub Intake_FindCustomer()
    Dim phone As String, r As Range, cnt As Long
    On Error GoTo EH
    Sh("INTAKE").Activate
    phone = NormalizePhone(CStr(NC("mk_in_Phone").Value))
    If Len(phone) = 0 Then
        Msg T("~27~28~2A~2F~27 ~34~45~27~31~47 ~2A~44~41~46 ~47~45~31~27~47 ~45~34~2A~31~CC ~31~27 ~48~27~31~2F ~A9~46~CC~2F."), vbExclamation
        NC("mk_in_Phone").Select
        Exit Sub
    End If
    If Not IsValidMobile(phone) Then
        Msg T("~34~45~27~31~47 ~2A~44~41~46 ~47~45~31~27~47 ~45~39~2A~28~31 ~46~CC~33~2A.") & vbLf & T("~28~27~CC~2F ~F1~F1 ~31~42~45 ~28~27~34~2F ~48 ~28~27 09 ~34~31~48~39 ~34~48~2F. ~46~45~48~46~47: 09121234567"), vbExclamation
        NC("mk_in_Phone").Select
        Exit Sub
    End If
    BeginBusy
    PutText NC("mk_in_Phone"), phone
    Set r = FindCustomerRow(phone)
    If r Is Nothing Then
        NC("mk_in_CustInfo").Value = T("~45~34~2A~31~CC ~2C~2F~CC~2F ~27~33~2A. ~46~27~45 ~31~27 ~48~27~31~2F ~A9~46~CC~2F~1B ~A9~2F ~45~34~2A~31~CC ~47~46~AF~27~45 ~2B~28~2A ~2E~48~2F~A9~27~31 ~33~27~2E~2A~47 ~45~CC^200C~34~48~2F.")
        EndBusy
        NC("mk_in_Name").Select
    Else
        cnt = Application.WorksheetFunction.CountIf(GetLO("tblDevices").ListColumns(DC_CUSTCODE).Range, CStr(r.Cells(1, CC_CODE).Value))
        NC("mk_in_Name").Value = r.Cells(1, CC_NAME).Value
        NC("mk_in_CustInfo").Value = T("~45~34~2A~31~CC ~42~2F~CC~45~CC: ") & r.Cells(1, CC_CODE).Value & vbLf & cnt & T(" ~2F~33~2A~AF~27~47 ~42~28~44~27~4B ~2B~28~2A ~34~2F~47 ~27~33~2A.")
        EndBusy
        NC("mk_in_Type").Select
    End If
    Exit Sub
EH:
    FailWith "Intake_FindCustomer"
End Sub

Public Sub Intake_Save()
    Dim phone As String, nm As String, dType As String, brand As String, model As String, serial As String
    Dim problem As String, accs As String, cond As String, payM As String, shelf As String, staff As String, notes As String
    Dim est As Double, dep As Double, custRow As Range, custCode As String, code As String, r As Range, s As String
    On Error GoTo EH
    Sh("INTAKE").Activate

    phone = NormalizePhone(CStr(NC("mk_in_Phone").Value))
    nm = CleanText(NC("mk_in_Name").Value)
    dType = CleanText(NC("mk_in_Type").Value)
    brand = CleanText(NC("mk_in_Brand").Value)
    model = CleanText(NC("mk_in_Model").Value)
    serial = CleanText(NC("mk_in_Serial").Value)
    problem = CleanText(NC("mk_in_Problem").Value)
    accs = CleanText(NC("mk_in_Access").Value)
    cond = CleanText(NC("mk_in_Cond").Value)
    payM = CleanText(NC("mk_in_PayMethod").Value)
    shelf = CleanText(NC("mk_in_Shelf").Value)
    staff = CleanText(NC("mk_in_Staff").Value)
    notes = CleanText(NC("mk_in_Notes").Value)

    ' ---- validation ----
    If Len(phone) = 0 Then FormError T("~34~45~27~31~47 ~2A~44~41~46 ~47~45~31~27~47 ~45~34~2A~31~CC ~31~27 ~48~27~31~2F ~A9~46~CC~2F."), "mk_in_Phone": Exit Sub
    If Not IsValidMobile(phone) Then FormError T("~34~45~27~31~47 ~2A~44~41~46 ~47~45~31~27~47 ~45~39~2A~28~31 ~46~CC~33~2A.") & vbLf & T("~28~27~CC~2F ~F1~F1 ~31~42~45 ~28~27~34~2F ~48 ~28~27 09 ~34~31~48~39 ~34~48~2F. ~46~45~48~46~47: 09121234567"), "mk_in_Phone": Exit Sub
    Set custRow = FindCustomerRow(phone)
    If custRow Is Nothing And Len(nm) = 0 Then FormError T("~27~CC~46 ~34~45~27~31~47 ~45~34~2A~31~CC ~2C~2F~CC~2F ~27~33~2A. ~46~27~45 ~45~34~2A~31~CC ~31~27 ~48~27~31~2F ~A9~46~CC~2F."), "mk_in_Name": Exit Sub
    If Len(dType) = 0 Then FormError T("~46~48~39 ~2F~33~2A~AF~27~47 ~31~27 ~27~32 ~41~47~31~33~2A ~27~46~2A~2E~27~28 ~A9~46~CC~2F."), "mk_in_Type": Exit Sub
    If Len(problem) = 0 Then FormError T("~27~CC~31~27~2F ~27~39~44~27~45^200C~34~2F~47 ~2A~48~33~37 ~45~34~2A~31~CC ~31~27 ~28~46~48~CC~33~CC~2F."), "mk_in_Problem": Exit Sub
    If Not CellAmount(NC("mk_in_Estimate").Value, est) Then FormError T("~28~31~22~48~31~2F ~47~32~CC~46~47 ~28~27~CC~2F ~39~2F~2F ~28~27~34~2F (~45~2B~44~27~4B 1500000)."), "mk_in_Estimate": Exit Sub
    If Not CellAmount(NC("mk_in_Deposit").Value, dep) Then FormError T("~28~CC~39~27~46~47 ~28~27~CC~2F ~39~2F~2F ~28~27~34~2F (~45~2B~44~27~4B 200000)."), "mk_in_Deposit": Exit Sub
    If est < 0 Or dep < 0 Then FormError T("~45~28~44~3A ~46~45~CC^200C~2A~48~27~46~2F ~45~46~41~CC ~28~27~34~2F."), "mk_in_Deposit": Exit Sub
    If Len(payM) = 0 Then payM = ListItem("tblPay", 1)

    ' ---- save ----
    BeginBusy
    If custRow Is Nothing Then
        custCode = CreateCustomer(phone, nm)
    Else
        custCode = CStr(custRow.Cells(1, CC_CODE).Value)
        If Len(CStr(custRow.Cells(1, CC_NAME).Value)) > 0 Then
            nm = CStr(custRow.Cells(1, CC_NAME).Value)
        Else
            custRow.Cells(1, CC_NAME).Value = nm
        End If
    End If

    code = NextDeviceCode()
    Set r = NewRow(GetLO("tblDevices"))
    PutText r.Cells(1, DC_CODE), code
    PutText r.Cells(1, DC_JDATE), GregorianToJalali(Date)
    PutText r.Cells(1, DC_TIME), HHMM()
    r.Cells(1, DC_GDATE).Value = Date
    PutText r.Cells(1, DC_CUSTCODE), custCode
    r.Cells(1, DC_CUSTNAME).Value = nm
    PutText r.Cells(1, DC_PHONE), phone
    r.Cells(1, DC_TYPE).Value = dType
    If Len(brand) > 0 And Len(model) > 0 Then
        r.Cells(1, DC_BRAND).Value = brand & " / " & model
    Else
        r.Cells(1, DC_BRAND).Value = brand & model
    End If
    PutText r.Cells(1, DC_SERIAL), serial
    r.Cells(1, DC_PROBLEM).Value = problem
    r.Cells(1, DC_ACCESS).Value = accs
    r.Cells(1, DC_COND).Value = cond
    PutText r.Cells(1, DC_SHELF), shelf
    r.Cells(1, DC_STATUS).Value = StatusName(ST_RECEIVED)
    r.Cells(1, DC_STAFF).Value = staff
    If est > 0 Then r.Cells(1, DC_ESTIMATE).Value = est
    r.Cells(1, DC_PARTSCOST).Value = 0
    r.Cells(1, DC_LABOR).Value = 0
    r.Cells(1, DC_PAID).Value = dep
    PutText r.Cells(1, DC_LASTCHG), JalaliNow()
    r.Cells(1, DC_NOTES).Value = notes

    If dep > 0 Then AddAccRow TT_INCOME, CAT_DEPOSIT, code, T("~28~CC~39~27~46~47 ~7E~30~CC~31~34 - ") & nm & " - " & dType, dep, payM

    FillReceipt r
    ClearIntakeForm
    PutText NC("mk_in_LastCode"), code
    EndBusy

    Sh("INTAKE").Activate
    NC("mk_in_Phone").Select
    s = T("~2F~33~2A~AF~27~47 ~28~27 ~45~48~41~42~CC~2A ~2B~28~2A ~34~2F.") & vbLf & vbLf & _
        T("~A9~2F ~2F~33~2A~AF~27~47:     ") & code & vbLf & vbLf & _
        T("~45~34~2A~31~CC: ") & nm & "  (" & custCode & ")" & vbLf
    If dep > 0 Then s = s & T("~28~CC~39~27~46~47 ~2B~28~2A^200C~34~2F~47: ") & Fmt(dep) & T(" ~2A~48~45~27~46") & vbLf
    s = s & vbLf & T("~27~CC~46 ~A9~2F ~31~27 ~31~48~CC ~28~31~86~33~28 ~2F~33~2A~AF~27~47 ~28~46~48~CC~33~CC~2F.") & vbLf & T("~28~31~27~CC ~86~27~7E ~31~33~CC~2F~0C ~2F~A9~45~47 ^00AB~86~27~7E ~31~33~CC~2F^00BB ~31~27 ~28~32~46~CC~2F.")
    Msg s, vbInformation, T("~2B~28~2A ~34~2F - ") & code
    Exit Sub
EH:
    FailWith "Intake_Save"
End Sub

Public Sub Intake_Clear()
    On Error GoTo EH
    BeginBusy
    ClearIntakeForm
    EndBusy
    Sh("INTAKE").Activate
    NC("mk_in_Phone").Select
    Exit Sub
EH:
    FailWith "Intake_Clear"
End Sub

Private Sub ClearIntakeForm()
    Dim nms As Variant, i As Long
    nms = Array("mk_in_Phone", "mk_in_Name", "mk_in_Type", "mk_in_Brand", "mk_in_Model", "mk_in_Serial", "mk_in_Problem", _
        "mk_in_Access", "mk_in_Cond", "mk_in_Estimate", "mk_in_Deposit", "mk_in_PayMethod", "mk_in_Shelf", "mk_in_Staff", "mk_in_Notes")
    For i = 0 To UBound(nms)
        NC(CStr(nms(i))).ClearContents
    Next i
    NC("mk_in_CustInfo").MergeArea.ClearContents
    NC("mk_in_PayMethod").Value = ListItem("tblPay", 1)
End Sub

Public Sub Intake_PrintReceipt()
    Dim v As Variant, r As Range
    On Error GoTo EH
    v = AskText(T("~A9~2F ~2F~33~2A~AF~27~47 ~28~31~27~CC ~86~27~7E ~31~33~CC~2F:"), T("~86~27~7E ~31~33~CC~2F"), CStr(NC("mk_in_LastCode").Value))
    If VarType(v) = vbBoolean Then Exit Sub
    If Len(v) = 0 Then Exit Sub
    Set r = FindDeviceRow(CStr(v))
    If r Is Nothing Then Msg T("~2F~33~2A~AF~27~47~CC ~28~27 ~27~CC~46 ~A9~2F ~7E~CC~2F~27 ~46~34~2F: ") & v, vbExclamation: Exit Sub
    ShowAndPrintReceipt r
    Exit Sub
EH:
    FailWith "Intake_PrintReceipt"
End Sub

Private Sub ShowAndPrintReceipt(ByVal r As Range)
    BeginBusy
    FillReceipt r
    SetupReceiptPage Sh("INTAKE")
    EndBusy
    Sh("INTAKE").Activate
    ActiveWindow.ScrollRow = 1: ActiveWindow.ScrollColumn = 1
    If Msg(T("~31~33~CC~2F ~2F~33~2A~AF~27~47 ") & r.Cells(1, DC_CODE).Value & T(" ~22~45~27~2F~47 ~27~33~2A.") & vbLf & vbLf & T("~86~27~7E ~34~48~2F~1F"), vbYesNo + vbQuestion, T("~86~27~7E ~31~33~CC~2F")) = vbYes Then
        On Error GoTo PrintFail
        Sh("INTAKE").PrintOut
    End If
    NC("mk_in_Phone").Select
    Exit Sub
PrintFail:
    Msg T("~86~27~7E ~27~46~2C~27~45 ~46~34~2F. ~27~2A~35~27~44 ~86~27~7E~AF~31 ~31~27 ~28~31~31~33~CC ~A9~46~CC~2F.") & vbLf & "(" & Err.Description & ")", vbExclamation
End Sub

Private Sub FillReceipt(ByVal r As Range)
    Dim idx As Long, wd As Long, total As Double, paid As Double, s As String
    idx = StatusIndex(CStr(r.Cells(1, DC_STATUS).Value))
    total = Num(r.Cells(1, DC_PARTSCOST).Value) + Num(r.Cells(1, DC_LABOR).Value)
    paid = Num(r.Cells(1, DC_PAID).Value)
    If idx = ST_DELIVERED Or idx = ST_RETURNED Then
        NC("mk_rc_Title").Value = T("~31~33~CC~2F ~2A~2D~48~CC~44 ~2F~33~2A~AF~27~47")
    Else
        NC("mk_rc_Title").Value = T("~31~33~CC~2F ~7E~30~CC~31~34 ~2F~33~2A~AF~27~47")
    End If
    PutText NC("mk_rc_Code"), CStr(r.Cells(1, DC_CODE).Value)
    PutText NC("mk_rc_Date"), r.Cells(1, DC_JDATE).Value & "  " & r.Cells(1, DC_TIME).Value
    NC("mk_rc_Cust").Value = r.Cells(1, DC_CUSTNAME).Value & "  (" & r.Cells(1, DC_CUSTCODE).Value & ")"
    PutText NC("mk_rc_Phone"), CStr(r.Cells(1, DC_PHONE).Value)
    NC("mk_rc_Device").Value = Trim$(r.Cells(1, DC_TYPE).Value & "  " & r.Cells(1, DC_BRAND).Value)
    NC("mk_rc_Problem").Value = r.Cells(1, DC_PROBLEM).Value
    NC("mk_rc_Access").Value = IIf(Len(CStr(r.Cells(1, DC_ACCESS).Value)) > 0, r.Cells(1, DC_ACCESS).Value, "-")
    NC("mk_rc_Status").Value = r.Cells(1, DC_STATUS).Value
    NC("mk_rc_Estimate").Value = Num(r.Cells(1, DC_ESTIMATE).Value)
    NC("mk_rc_Total").Value = total
    NC("mk_rc_Paid").Value = paid
    NC("mk_rc_Balance").Value = IIf(total > 0, total - paid, 0)
    PutText NC("mk_rc_Warranty"), IIf(Len(CStr(r.Cells(1, DC_WARRANTY).Value)) > 0, CStr(r.Cells(1, DC_WARRANTY).Value), "-")
    wd = CLng(Num(NC("mk_WarrantyDays").Value))
    s = T("~44~37~41~27~4B ~27~CC~46 ~31~33~CC~2F ~31~27 ~47~46~AF~27~45 ~2A~2D~48~CC~44 ~AF~31~41~2A~46 ~2F~33~2A~AF~27~47 ~47~45~31~27~47 ~2F~27~34~2A~47 ~28~27~34~CC~2F.")
    If wd > 0 Then s = s & vbLf & T("~AF~27~31~27~46~2A~CC ~2A~39~45~CC~31: ") & wd & T(" ~31~48~32 ~27~32 ~2A~27~31~CC~2E ~2A~2D~48~CC~44")
    NC("mk_rc_Footer").Value = s
End Sub

'==============================================================================
'  DEVICES
'==============================================================================

Public Sub Dev_ChangeStatus()
    Dim r As Range, idx As Long, cur As String, dflt As Long
    On Error GoTo EH
    Set r = SelectedDeviceRow()
    If r Is Nothing Then Exit Sub
    cur = CStr(r.Cells(1, DC_STATUS).Value)
    dflt = StatusIndex(cur) + 1
    If dflt < 1 Or dflt > STATUS_COUNT Then dflt = ST_DIAGNOSING
    idx = PickIndex("tblStatus", T("~2A~3A~CC~CC~31 ~48~36~39~CC~2A ") & r.Cells(1, DC_CODE).Value, T("~48~36~39~CC~2A ~41~39~44~CC: ") & cur & vbLf & T("~34~45~27~31~47 ~48~36~39~CC~2A ~2C~2F~CC~2F:"), dflt)
    If idx = 0 Then Exit Sub
    BeginBusy
    StatusChangeFlow r, idx, False
    EndBusy
    Exit Sub
EH:
    FailWith "Dev_ChangeStatus"
End Sub

' Applies a new status: stamps the date and, on delivery/return, settles the account.
' fromEvent = True when the user already picked the status in the dropdown.
Private Sub StatusChangeFlow(ByVal r As Range, ByVal idx As Long, ByVal fromEvent As Boolean)
    Dim code As String, total As Double, paid As Double, bal As Double, pay As Variant, method As String, wd As Long, s As String
    code = CStr(r.Cells(1, DC_CODE).Value)
    If idx = ST_DELIVERED Or idx = ST_RETURNED Then
        total = Num(r.Cells(1, DC_PARTSCOST).Value) + Num(r.Cells(1, DC_LABOR).Value)
        paid = Num(r.Cells(1, DC_PAID).Value)
        bal = total - paid
        If bal < 0 Then bal = 0
        If total = 0 And idx = ST_DELIVERED Then
            If Msg(T("~2C~45~39 ~47~32~CC~46~47 ~27~CC~46 ~2F~33~2A~AF~27~47 ~35~41~31 ~27~33~2A (~2F~33~2A~45~32~2F ~48 ~42~37~39~47 ~2B~28~2A ~46~34~2F~47).") & vbLf & T("~27~28~2A~2F~27 ~33~2A~48~46 ^00AB~2F~33~2A~45~32~2F^00BB ~31~27 ~7E~31 ~A9~46~CC~2F.") & vbLf & vbLf & T("~28~27 ~27~CC~46 ~2D~27~44 ~2A~2D~48~CC~44 ~2B~28~2A ~34~48~2F~1F"), _
                   vbYesNo + vbExclamation, T("~2A~2D~48~CC~44 ") & code) <> vbYes Then
                If Not fromEvent Then Exit Sub
            End If
        End If
        pay = AskNumber(T("~2A~33~48~CC~47 ~2D~33~27~28 ") & code & vbLf & T("~2C~45~39: ") & Fmt(total) & T("   ~7E~31~2F~27~2E~2A~CC: ") & Fmt(paid) & vbLf & T("~45~27~46~2F~47: ") & Fmt(bal) & T(" ~2A~48~45~27~46") & vbLf & vbLf & T("~45~28~44~3A ~2F~31~CC~27~41~2A~CC ~27~44~27~46:"), _
                        T("~2A~2D~48~CC~44 ~2F~33~2A~AF~27~47"), bal)
        If VarType(pay) = vbBoolean Then
            If Not fromEvent Then Msg T("~2A~3A~CC~CC~31 ~48~36~39~CC~2A ~44~3A~48 ~34~2F."), vbInformation: Exit Sub
            pay = 0
        End If
        If CDbl(pay) > 0 Then
            method = PickMethod()
            If Len(method) = 0 Then
                If Not fromEvent Then Msg T("~2A~3A~CC~CC~31 ~48~36~39~CC~2A ~44~3A~48 ~34~2F."), vbInformation: Exit Sub
                method = ListItem("tblPay", 1)
            End If
        End If
        r.Cells(1, DC_STATUS).Value = StatusName(idx)
        PutText r.Cells(1, DC_LASTCHG), JalaliNow()
        If CDbl(pay) > 0 Then RecordDevicePayment r, CDbl(pay), method
        PutText r.Cells(1, DC_DELIVERY), JalaliNow()
        r.Cells(1, DC_GDELIV).Value = Date
        wd = CLng(Num(NC("mk_WarrantyDays").Value))
        If idx = ST_DELIVERED And wd > 0 Then
            PutText r.Cells(1, DC_WARRANTY), GregorianToJalali(Date + wd)
        Else
            r.Cells(1, DC_WARRANTY).Value = ""
        End If
        bal = total - Num(r.Cells(1, DC_PAID).Value)
        s = T("~48~36~39~CC~2A ~2F~33~2A~AF~27~47 ") & code & T(" ~28~47 ^00AB") & StatusName(idx) & T("^00BB ~2A~3A~CC~CC~31 ~A9~31~2F.") & vbLf
        If CDbl(pay) > 0 Then s = s & T("~45~28~44~3A ~2F~31~CC~27~41~2A~CC: ") & Fmt(CDbl(pay)) & T(" ~2A~48~45~27~46 (~2F~31 ~2D~33~27~28~2F~27~31~CC ~2B~28~2A ~34~2F)") & vbLf
        If bal > 0 Then s = s & vbLf & T("~2A~48~2C~47: ~47~46~48~32 ") & Fmt(bal) & T(" ~2A~48~45~27~46 ~45~27~46~2F~47 ~27~33~2A.") & vbLf
        If Len(CStr(r.Cells(1, DC_WARRANTY).Value)) > 0 Then s = s & T("~AF~27~31~27~46~2A~CC ~2A~27: ") & r.Cells(1, DC_WARRANTY).Value & vbLf
        s = s & vbLf & T("~31~33~CC~2F ~2A~2D~48~CC~44 ~86~27~7E ~34~48~2F~1F")
        If Msg(s, vbYesNo + vbInformation, T("~2A~2D~48~CC~44 ") & code) = vbYes Then
            EndBusy
            ShowAndPrintReceipt r
            BeginBusy
        End If
    Else
        r.Cells(1, DC_STATUS).Value = StatusName(idx)
        PutText r.Cells(1, DC_LASTCHG), JalaliNow()
        If Len(CStr(r.Cells(1, DC_DELIVERY).Value)) > 0 Then
            ' device came back from "delivered": clear delivery stamps
            r.Cells(1, DC_DELIVERY).Value = ""
            r.Cells(1, DC_GDELIV).Value = ""
            r.Cells(1, DC_WARRANTY).Value = ""
        End If
    End If
End Sub

Private Sub RecordDevicePayment(ByVal r As Range, ByVal amount As Double, ByVal method As String)
    r.Cells(1, DC_PAID).Value = Num(r.Cells(1, DC_PAID).Value) + amount
    AddAccRow TT_INCOME, CAT_REPAIR, CStr(r.Cells(1, DC_CODE).Value), _
        T("~2F~31~CC~27~41~2A ~28~27~28~2A ~2A~39~45~CC~31 - ") & r.Cells(1, DC_CUSTNAME).Value & " - " & r.Cells(1, DC_TYPE).Value, amount, method
End Sub

Public Sub Dev_ReceivePayment()
    Dim r As Range, total As Double, paid As Double, bal As Double, pay As Variant, method As String
    On Error GoTo EH
    Set r = SelectedDeviceRow()
    If r Is Nothing Then Exit Sub
    total = Num(r.Cells(1, DC_PARTSCOST).Value) + Num(r.Cells(1, DC_LABOR).Value)
    paid = Num(r.Cells(1, DC_PAID).Value)
    bal = total - paid
    If bal < 0 Then bal = 0
    pay = AskNumber(T("~2F~33~2A~AF~27~47 ") & r.Cells(1, DC_CODE).Value & " - " & r.Cells(1, DC_CUSTNAME).Value & vbLf & _
                    T("~2C~45~39: ") & Fmt(total) & T("   ~7E~31~2F~27~2E~2A~CC: ") & Fmt(paid) & vbLf & T("~45~27~46~2F~47: ") & Fmt(bal) & T(" ~2A~48~45~27~46") & vbLf & vbLf & T("~45~28~44~3A ~2F~31~CC~27~41~2A~CC:"), _
                    T("~2F~31~CC~27~41~2A ~48~2C~47"), bal)
    If VarType(pay) = vbBoolean Then Exit Sub
    If CDbl(pay) <= 0 Then Msg T("~45~28~44~3A ~28~27~CC~2F ~28~CC~34~2A~31 ~27~32 ~35~41~31 ~28~27~34~2F."), vbExclamation: Exit Sub
    method = PickMethod()
    If Len(method) = 0 Then Exit Sub
    BeginBusy
    RecordDevicePayment r, CDbl(pay), method
    EndBusy
    Msg T("~45~28~44~3A ") & Fmt(CDbl(pay)) & T(" ~2A~48~45~27~46 ~28~31~27~CC ~2F~33~2A~AF~27~47 ") & r.Cells(1, DC_CODE).Value & T(" ~2B~28~2A ~34~2F.") & vbLf & _
        BalanceText(total - Num(r.Cells(1, DC_PAID).Value)), vbInformation
    Exit Sub
EH:
    FailWith "Dev_ReceivePayment"
End Sub

Public Sub Dev_UsePart()
    Dim r As Range, v As Variant, pr As Range, qty As Variant, price As Variant, stock As Double, defPrice As Double
    Dim pcode As String, pname As String, used As String
    On Error GoTo EH
    Set r = SelectedDeviceRow()
    If r Is Nothing Then Exit Sub
    v = AskText(PartPrompt(T("~A9~2F ~42~37~39~47 ~45~35~31~41~CC ~28~31~27~CC ") & r.Cells(1, DC_CODE).Value & ":"), T("~45~35~31~41 ~42~37~39~47 ~27~32 ~27~46~28~27~31"))
    If VarType(v) = vbBoolean Then Exit Sub
    If Len(v) = 0 Then Exit Sub
    Set pr = FindPartRow(CStr(v))
    If pr Is Nothing Then Msg T("~42~37~39~47^200C~27~CC ~28~27 ~27~CC~46 ~A9~2F ~2F~31 ~27~46~28~27~31 ~7E~CC~2F~27 ~46~34~2F: ") & v, vbExclamation: Exit Sub
    pcode = CStr(pr.Cells(1, PC_CODE).Value)
    pname = CStr(pr.Cells(1, PC_NAME).Value)
    stock = Num(pr.Cells(1, PC_QTY).Value)
    qty = AskNumber(pname & " (" & pcode & ")" & vbLf & T("~45~48~2C~48~2F~CC ~41~39~44~CC: ") & stock & vbLf & vbLf & T("~2A~39~2F~27~2F ~45~35~31~41~CC:"), T("~45~35~31~41 ~42~37~39~47"), 1)
    If VarType(qty) = vbBoolean Then Exit Sub
    If CDbl(qty) <= 0 Then Msg T("~2A~39~2F~27~2F ~28~27~CC~2F ~28~CC~34~2A~31 ~27~32 ~35~41~31 ~28~27~34~2F."), vbExclamation: Exit Sub
    If CDbl(qty) > stock Then Msg T("~45~48~2C~48~2F~CC ~A9~27~41~CC ~46~CC~33~2A.") & vbLf & T("~45~48~2C~48~2F~CC ~41~39~44~CC ") & pname & ": " & stock, vbExclamation: Exit Sub
    defPrice = Num(pr.Cells(1, PC_SELL).Value)
    If defPrice = 0 Then defPrice = Num(pr.Cells(1, PC_BUY).Value)
    price = AskNumber(T("~42~CC~45~2A ~47~31 ~48~27~2D~2F ~28~31~27~CC ~45~34~2A~31~CC (~2A~48~45~27~46):"), T("~45~35~31~41 ~42~37~39~47"), defPrice)
    If VarType(price) = vbBoolean Then Exit Sub
    BeginBusy
    pr.Cells(1, PC_QTY).Value = stock - CDbl(qty)
    used = CStr(r.Cells(1, DC_PARTS).Value)
    If Len(used) > 0 Then used = used & T("~0C ")
    r.Cells(1, DC_PARTS).Value = used & pcode & T("^00D7") & qty
    r.Cells(1, DC_PARTSCOST).Value = Num(r.Cells(1, DC_PARTSCOST).Value) + CDbl(qty) * CDbl(price)
    Dash_RefreshCore
    EndBusy
    If stock - CDbl(qty) <= Num(pr.Cells(1, PC_MIN).Value) Then
        Msg T("~42~37~39~47 ~2B~28~2A ~34~2F. ~45~48~2C~48~2F~CC ~2C~2F~CC~2F ") & pname & ": " & (stock - CDbl(qty)) & vbLf & vbLf & T("~47~34~2F~27~31: ~45~48~2C~48~2F~CC ~27~CC~46 ~42~37~39~47 ~28~47 ~2D~2F~27~42~44 ~31~33~CC~2F~47 ~27~33~2A."), vbExclamation
    Else
        Msg T("~42~37~39~47 ~2B~28~2A ~34~2F. ~45~48~2C~48~2F~CC ~2C~2F~CC~2F ") & pname & ": " & (stock - CDbl(qty)), vbInformation
    End If
    Exit Sub
EH:
    FailWith "Dev_UsePart"
End Sub

Public Sub Dev_PrintReceipt()
    Dim r As Range
    On Error GoTo EH
    Set r = SelectedDeviceRow()
    If r Is Nothing Then Exit Sub
    ShowAndPrintReceipt r
    Exit Sub
EH:
    FailWith "Dev_PrintReceipt"
End Sub

' Quick filter: each colored button on the devices sheet calls this; the button name says which status.
Public Sub Dev_FilterByStatus()
    Dim lo As ListObject, idx As Long, caller As String
    On Error GoTo EH
    If TypeName(Application.Caller) = "String" Then caller = Application.Caller
    If Left$(caller, 7) = "mk_flt_" Then
        idx = CLng(Val(Mid$(caller, 8)))
    Else
        idx = PickIndex("tblStatus", T("~41~CC~44~2A~31"), T("~34~45~27~31~47 ~48~36~39~CC~2A:"), ST_READY)
    End If
    If idx < 1 Or idx > STATUS_COUNT Then Exit Sub
    Set lo = GetLO("tblDevices")
    BeginBusy
    lo.Range.AutoFilter Field:=DC_STATUS, Criteria1:=StatusName(idx)
    EndBusy
    Sh("DEV").Activate
    ActiveWindow.ScrollRow = 1
    Exit Sub
EH:
    FailWith "Dev_FilterByStatus"
End Sub

Public Sub Dev_ClearFilter()
    Dim lo As ListObject
    On Error GoTo EH
    Set lo = GetLO("tblDevices")
    BeginBusy
    If lo.AutoFilter Is Nothing Then
        lo.ShowAutoFilter = True
    ElseIf lo.AutoFilter.FilterMode Then
        lo.AutoFilter.ShowAllData
    End If
    EndBusy
    Exit Sub
EH:
    FailWith "Dev_ClearFilter"
End Sub

' Called by the optional Worksheet_Change handler of the devices sheet.
Public Sub Mokonic_OnDevicesChange(ByVal Target As Range)
    Dim lo As ListObject, isect As Range, c As Range, r As Range, idx As Long
    On Error GoTo EH
    If Not Application.EnableEvents Then Exit Sub
    Set lo = GetLO("tblDevices")
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Not Target.Worksheet Is lo.Parent Then Exit Sub
    Set isect = Application.Intersect(Target, lo.ListColumns(DC_STATUS).DataBodyRange)
    If isect Is Nothing Then Exit Sub
    If isect.Cells.Count > 50 Then Exit Sub
    BeginBusy
    For Each c In isect.Cells
        Set r = lo.ListRows(c.Row - lo.HeaderRowRange.Row).Range
        If Len(CStr(r.Cells(1, DC_CODE).Value)) > 0 Then
            idx = StatusIndex(CStr(c.Value))
            If idx > 0 Then StatusChangeFlow r, idx, True
        End If
    Next c
    EndBusy
    Exit Sub
EH:
    FailWith "Mokonic_OnDevicesChange"
End Sub

Private Function SelectedDeviceRow() As Range
    Dim lo As ListObject, idx As Long, r As Range
    Set lo = GetLO("tblDevices")
    If Not ActiveSheet Is lo.Parent Then
        lo.Parent.Activate
        Msg T("~2F~31 ~28~31~AF~47 ^00AB~2F~33~2A~AF~27~47^200C~47~27^00BB ~31~48~CC ~31~2F~CC~41 ~2F~33~2A~AF~27~47 ~45~48~31~2F ~46~38~31 ~A9~44~CC~A9 ~A9~46~CC~2F ~48 ~2F~48~28~27~31~47 ~2F~A9~45~47 ~31~27 ~28~32~46~CC~2F."), vbExclamation
        Exit Function
    End If
    If Not lo.DataBodyRange Is Nothing Then
        If Not Application.Intersect(ActiveCell, lo.DataBodyRange) Is Nothing Then
            idx = ActiveCell.Row - lo.HeaderRowRange.Row
            Set r = lo.ListRows(idx).Range
            If Len(CStr(r.Cells(1, DC_CODE).Value)) > 0 Then
                Set SelectedDeviceRow = r
                Exit Function
            End If
        End If
    End If
    Msg T("~27~28~2A~2F~27 ~31~48~CC ~CC~A9~CC ~27~32 ~31~2F~CC~41^200C~47~27~CC ~2C~2F~48~44 ~2F~33~2A~AF~27~47^200C~47~27 ~A9~44~CC~A9 ~A9~46~CC~2F~0C ~33~7E~33 ~2F~A9~45~47 ~31~27 ~28~32~46~CC~2F."), vbExclamation
End Function

'==============================================================================
'  SEARCH
'==============================================================================

Public Sub Search_Run()
    Dim ws As Worksheet, lo As ListObject, q As String, qd As String, mode As Long, i As Long, n As Long
    Dim r As Range, hit As Boolean, outRow As Long, idx As Long, devName As String
    On Error GoTo EH
    Set ws = Sh("SEARCH")
    q = NormalizeFa(CStr(NC("mk_srch_Text").Value))
    If Len(q) = 0 Then
        Msg T("~39~28~27~31~2A ~2C~33~2A~2C~48 ~31~27 ~48~27~31~2F ~A9~46~CC~2F."), vbExclamation
        ws.Activate
        NC("mk_srch_Text").Select
        Exit Sub
    End If
    mode = ListIndexOf("tblSearchBy", CStr(NC("mk_srch_By").Value))
    If mode = 0 Then mode = 1
    qd = DigitsOnly(q)
    If mode = 1 And Len(qd) = 0 Then Msg T("~28~31~27~CC ~2C~33~2A~2C~48 ~28~27 ~2A~44~41~46~0C ~31~42~45 ~48~27~31~2F ~A9~46~CC~2F."), vbExclamation: Exit Sub

    BeginBusy
    ClearSearchResults ws
    Set lo = GetLO("tblDevices")
    devName = lo.Parent.Name
    outRow = SEARCH_FIRST_ROW
    If Not lo.DataBodyRange Is Nothing Then
        For i = lo.ListRows.Count To 1 Step -1          ' newest first
            Set r = lo.ListRows(i).Range
            If Len(CStr(r.Cells(1, DC_CODE).Value)) > 0 Then
                Select Case mode
                    Case 1
                        hit = InStr(1, DigitsOnly(CStr(r.Cells(1, DC_PHONE).Value)), qd) > 0
                    Case 2
                        hit = InStr(1, NormalizeFa(CStr(r.Cells(1, DC_CUSTNAME).Value)), q, vbTextCompare) > 0
                    Case Else
                        hit = InStr(1, CStr(r.Cells(1, DC_CODE).Value), q, vbTextCompare) > 0
                End Select
                If hit Then
                    n = n + 1
                    If n <= SEARCH_MAX_ROWS Then
                        ws.Hyperlinks.Add Anchor:=ws.Cells(outRow, 2), Address:="", _
                            SubAddress:="'" & devName & "'!" & r.Cells(1, 1).Address, TextToDisplay:=CStr(r.Cells(1, DC_CODE).Value)
                        PutText ws.Cells(outRow, 3), CStr(r.Cells(1, DC_JDATE).Value)
                        ws.Cells(outRow, 4).Value = r.Cells(1, DC_CUSTNAME).Value
                        PutText ws.Cells(outRow, 5), CStr(r.Cells(1, DC_PHONE).Value)
                        ws.Cells(outRow, 6).Value = Trim$(r.Cells(1, DC_TYPE).Value & "  " & r.Cells(1, DC_BRAND).Value)
                        ws.Cells(outRow, 7).Value = r.Cells(1, DC_STATUS).Value
                        ws.Cells(outRow, 8).Value = Num(r.Cells(1, DC_PARTSCOST).Value) + Num(r.Cells(1, DC_LABOR).Value)
                        ws.Cells(outRow, 9).Value = Num(ws.Cells(outRow, 8).Value) - Num(r.Cells(1, DC_PAID).Value)
                        idx = StatusIndex(CStr(r.Cells(1, DC_STATUS).Value))
                        If idx > 0 Then
                            ws.Cells(outRow, 7).Interior.Color = StatusFill(idx)
                            ws.Cells(outRow, 7).Font.Color = StatusFont(idx)
                            ws.Cells(outRow, 7).Font.Bold = True
                        End If
                        outRow = outRow + 1
                    End If
                End If
            End If
        Next i
    End If
    If n > 0 Then
        With ws.Range(ws.Cells(SEARCH_FIRST_ROW, 2), ws.Cells(outRow - 1, 9))
            .Font.Name = FONT_NAME
            .Font.Size = 10
            .Borders(xlInsideHorizontal).Color = RGB(220, 220, 220)
            .Borders(xlEdgeBottom).Color = RGB(200, 200, 200)
            .RowHeight = 22
        End With
        ws.Range(ws.Cells(SEARCH_FIRST_ROW, 2), ws.Cells(outRow - 1, 3)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(SEARCH_FIRST_ROW, 5), ws.Cells(outRow - 1, 5)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(SEARCH_FIRST_ROW, 7), ws.Cells(outRow - 1, 7)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(SEARCH_FIRST_ROW, 8), ws.Cells(outRow - 1, 9)).NumberFormat = "#,##0;[Red]-#,##0"
        NC("mk_srch_Info").Value = n & T(" ~2F~33~2A~AF~27~47 ~7E~CC~2F~27 ~34~2F.") & IIf(n > SEARCH_MAX_ROWS, T(" (~41~42~37 ") & SEARCH_MAX_ROWS & T(" ~45~48~31~2F ~27~48~44 ~46~45~27~CC~34 ~2F~27~2F~47 ~34~2F)"), "")
    Else
        NC("mk_srch_Info").Value = T("~45~48~31~2F~CC ~7E~CC~2F~27 ~46~34~2F.")
    End If
    EndBusy
    ws.Activate
    NC("mk_srch_Text").Select
    Exit Sub
EH:
    FailWith "Search_Run"
End Sub

Public Sub Search_Clear()
    Dim ws As Worksheet
    On Error GoTo EH
    Set ws = Sh("SEARCH")
    BeginBusy
    ClearSearchResults ws
    NC("mk_srch_Text").ClearContents
    NC("mk_srch_Info").ClearContents
    EndBusy
    ws.Activate
    NC("mk_srch_Text").Select
    Exit Sub
EH:
    FailWith "Search_Clear"
End Sub

Private Sub ClearSearchResults(ByVal ws As Worksheet)
    With ws.Range(ws.Cells(SEARCH_FIRST_ROW, 2), ws.Cells(SEARCH_FIRST_ROW + SEARCH_MAX_ROWS + 5, 9))
        .Hyperlinks.Delete
        .Clear
        .Font.Name = FONT_NAME
        .Font.Size = 10
    End With
End Sub

'==============================================================================
'  CUSTOMERS
'==============================================================================

Public Sub Cust_Add()
    Dim v As Variant, phone As String, nm As Variant, r As Range, code As String
    On Error GoTo EH
    v = AskText(T("~34~45~27~31~47 ~2A~44~41~46 ~47~45~31~27~47 ~45~34~2A~31~CC ~2C~2F~CC~2F (~45~2B~44~27~4B 09121234567):"), T("~45~34~2A~31~CC ~2C~2F~CC~2F"))
    If VarType(v) = vbBoolean Then Exit Sub
    phone = NormalizePhone(CStr(v))
    If Not IsValidMobile(phone) Then Msg T("~34~45~27~31~47 ~2A~44~41~46 ~47~45~31~27~47 ~45~39~2A~28~31 ~46~CC~33~2A. ~46~45~48~46~47 ~2F~31~33~2A: 09121234567"), vbExclamation: Exit Sub
    Set r = FindCustomerRow(phone)
    If Not r Is Nothing Then
        Msg T("~27~CC~46 ~34~45~27~31~47 ~42~28~44~27~4B ~2B~28~2A ~34~2F~47 ~27~33~2A:") & vbLf & r.Cells(1, CC_CODE).Value & " - " & r.Cells(1, CC_NAME).Value, vbExclamation
        Exit Sub
    End If
    nm = AskText(T("~46~27~45 ~48 ~46~27~45 ~2E~27~46~48~27~2F~AF~CC ~45~34~2A~31~CC:"), T("~45~34~2A~31~CC ~2C~2F~CC~2F"))
    If VarType(nm) = vbBoolean Then Exit Sub
    If Len(nm) = 0 Then Msg T("~46~27~45 ~45~34~2A~31~CC ~27~44~32~27~45~CC ~27~33~2A."), vbExclamation: Exit Sub
    BeginBusy
    code = CreateCustomer(phone, CleanText(nm))
    EndBusy
    Set r = FindCustomerRow(phone)
    Sh("CUST").Activate
    If Not r Is Nothing Then r.Cells(1, CC_PHONE2).Select
    Msg T("~45~34~2A~31~CC ~28~27 ~A9~2F ") & code & T(" ~2B~28~2A ~34~2F. ~45~CC^200C~2A~48~27~46~CC~2F ~22~2F~31~33 ~48 ~2A~44~41~46 ~2F~48~45 ~31~27 ~2F~31 ~47~45~27~46 ~31~2F~CC~41 ~28~46~48~CC~33~CC~2F."), vbInformation
    Exit Sub
EH:
    FailWith "Cust_Add"
End Sub

Private Function CreateCustomer(ByVal phone As String, ByVal nm As String) As String
    Dim lo As ListObject, r As Range, code As String
    Set lo = GetLO("tblCustomers")
    code = NextSeqCode("C-", 4, "mk_cntCust", lo, CC_CODE)
    Set r = NewRow(lo)
    PutText r.Cells(1, CC_CODE), code
    r.Cells(1, CC_NAME).Value = nm
    PutText r.Cells(1, CC_PHONE), phone
    PutText r.Cells(1, CC_FIRST), GregorianToJalali(Date)
    CreateCustomer = code
End Function

Private Function FindCustomerRow(ByVal phone As String) As Range
    Dim lo As ListObject, c As Range
    Set lo = GetLO("tblCustomers")
    If lo.DataBodyRange Is Nothing Then Exit Function
    For Each c In lo.ListColumns(CC_PHONE).DataBodyRange.Cells
        If Len(CStr(c.Value)) > 0 Then
            If NormalizePhone(CStr(c.Value)) = phone Then
                Set FindCustomerRow = lo.ListRows(c.Row - lo.HeaderRowRange.Row).Range
                Exit Function
            End If
        End If
    Next c
End Function

'==============================================================================
'  ACCOUNTING
'==============================================================================

Public Sub Acc_AddManual()
    Dim tIdx As Long, cIdx As Long, amt As Variant, desc As Variant, method As String, ref As Variant
    On Error GoTo EH
    tIdx = PickIndex("tblAccType", T("~2B~28~2A ~2F~33~2A~CC"), T("~46~48~39 ~2A~31~27~A9~46~34:"), TT_EXPENSE)
    If tIdx = 0 Then Exit Sub
    cIdx = PickIndex("tblAccCat", T("~2B~28~2A ~2F~33~2A~CC"), T("~2F~33~2A~47:"), IIf(tIdx = TT_INCOME, CAT_REPAIR, 5))
    If cIdx = 0 Then Exit Sub
    amt = AskNumber(T("~45~28~44~3A (~2A~48~45~27~46):"), T("~2B~28~2A ~2F~33~2A~CC"), 0)
    If VarType(amt) = vbBoolean Then Exit Sub
    If CDbl(amt) <= 0 Then Msg T("~45~28~44~3A ~28~27~CC~2F ~28~CC~34~2A~31 ~27~32 ~35~41~31 ~28~27~34~2F."), vbExclamation: Exit Sub
    desc = AskText(T("~34~31~2D (~45~2B~44~27~4B ~27~2C~27~31~47 ~22~28~27~46~0C ~42~28~36 ~28~31~42):"), T("~2B~28~2A ~2F~33~2A~CC"))
    If VarType(desc) = vbBoolean Then Exit Sub
    ref = AskText(T("~A9~2F ~45~31~2A~28~37 (~27~2E~2A~CC~27~31~CC - ~45~2B~44~27~4B ~A9~2F ~2F~33~2A~AF~27~47):"), T("~2B~28~2A ~2F~33~2A~CC"))
    If VarType(ref) = vbBoolean Then ref = ""
    method = PickMethod()
    If Len(method) = 0 Then Exit Sub
    BeginBusy
    AddAccRow tIdx, cIdx, CStr(ref), CStr(desc), CDbl(amt), method
    EndBusy
    Msg ListItem("tblAccType", tIdx) & T(" ~28~47 ~45~28~44~3A ") & Fmt(CDbl(amt)) & T(" ~2A~48~45~27~46 ~2B~28~2A ~34~2F."), vbInformation
    Exit Sub
EH:
    FailWith "Acc_AddManual"
End Sub

Public Sub Acc_Refresh()
    On Error GoTo EH
    BeginBusy
    NC("mk_acc_Month").ClearContents
    Application.CalculateFull
    EndBusy
    Sh("ACC").Activate
    Exit Sub
EH:
    FailWith "Acc_Refresh"
End Sub

' Appends one transaction. typeIdx: 1 income / 2 expense. catIdx: position in the category list.
Private Sub AddAccRow(ByVal typeIdx As Long, ByVal catIdx As Long, ByVal ref As String, ByVal desc As String, ByVal amount As Double, ByVal method As String)
    Dim r As Range
    Set r = NewRow(GetLO("tblAcc"))
    PutText r.Cells(1, AC_JDATE), GregorianToJalali(Date)
    r.Cells(1, AC_GDATE).Value = Date
    r.Cells(1, AC_TYPE).Value = ListItem("tblAccType", typeIdx)
    r.Cells(1, AC_CAT).Value = ListItem("tblAccCat", catIdx)
    PutText r.Cells(1, AC_REF), ref
    r.Cells(1, AC_DESC).Value = desc
    r.Cells(1, AC_AMOUNT).Value = amount
    r.Cells(1, AC_METHOD).Value = method
End Sub

'==============================================================================
'  WAREHOUSE / PURCHASES / TOOLS
'==============================================================================

Public Sub Stock_AddPart()
    Dim nm As Variant, cIdx As Long, qty As Variant, minQ As Variant, lo As ListObject, r As Range, code As String
    On Error GoTo EH
    nm = AskText(T("~46~27~45 ~42~37~39~47 (~45~2B~44~27~4B ~2E~27~32~46 470uF 25V):"), T("~42~37~39~47 ~2C~2F~CC~2F"))
    If VarType(nm) = vbBoolean Then Exit Sub
    If Len(nm) = 0 Then Msg T("~46~27~45 ~42~37~39~47 ~27~44~32~27~45~CC ~27~33~2A."), vbExclamation: Exit Sub
    cIdx = PickIndex("tblPartCat", T("~42~37~39~47 ~2C~2F~CC~2F"), T("~2F~33~2A~47 ~42~37~39~47:"), 1)
    If cIdx = 0 Then Exit Sub
    qty = AskNumber(T("~45~48~2C~48~2F~CC ~27~48~44~CC~47:"), T("~42~37~39~47 ~2C~2F~CC~2F"), 0)
    If VarType(qty) = vbBoolean Then Exit Sub
    minQ = AskNumber(T("~2D~2F~27~42~44 ~45~48~2C~48~2F~CC (~28~31~27~CC ~47~34~2F~27~31 ~A9~45~28~48~2F):"), T("~42~37~39~47 ~2C~2F~CC~2F"), 1)
    If VarType(minQ) = vbBoolean Then Exit Sub
    BeginBusy
    Set lo = GetLO("tblParts")
    code = NextSeqCode("P-", 4, "mk_cntPart", lo, PC_CODE)
    Set r = NewRow(lo)
    PutText r.Cells(1, PC_CODE), code
    r.Cells(1, PC_NAME).Value = CleanText(nm)
    r.Cells(1, PC_CAT).Value = ListItem("tblPartCat", cIdx)
    r.Cells(1, PC_QTY).Value = CDbl(qty)
    r.Cells(1, PC_UNIT).Value = ListItem("tblUnits", 1)
    r.Cells(1, PC_MIN).Value = CDbl(minQ)
    r.Cells(1, PC_BUY).Value = 0
    r.Cells(1, PC_SELL).Value = 0
    Dash_RefreshCore
    EndBusy
    Sh("STOCK").Activate
    r.Cells(1, PC_SPEC).Select
    Msg T("~42~37~39~47 ~28~27 ~A9~2F ") & code & T(" ~33~27~2E~2A~47 ~34~2F.") & vbLf & T("~45~34~2E~35~27~2A~0C ~42~41~33~47~0C ~42~CC~45~2A ~41~31~48~34 ~48 ~2A~23~45~CC~46^200C~A9~46~46~2F~47 ~31~27 ~2F~31 ~47~45~27~46 ~31~2F~CC~41 ~28~46~48~CC~33~CC~2F."), vbInformation
    Exit Sub
EH:
    FailWith "Stock_AddPart"
End Sub

Public Sub Stock_SellPart()
    Dim v As Variant, pr As Range, qty As Variant, price As Variant, method As String, stock As Double, pname As String, pcode As String
    On Error GoTo EH
    v = AskText(PartPrompt(T("~A9~2F ~42~37~39~47 ~28~31~27~CC ~41~31~48~34:")), T("~41~31~48~34 ~42~37~39~47"))
    If VarType(v) = vbBoolean Then Exit Sub
    If Len(v) = 0 Then Exit Sub
    Set pr = FindPartRow(CStr(v))
    If pr Is Nothing Then Msg T("~42~37~39~47^200C~27~CC ~28~27 ~27~CC~46 ~A9~2F ~7E~CC~2F~27 ~46~34~2F: ") & v, vbExclamation: Exit Sub
    pcode = CStr(pr.Cells(1, PC_CODE).Value)
    pname = CStr(pr.Cells(1, PC_NAME).Value)
    stock = Num(pr.Cells(1, PC_QTY).Value)
    qty = AskNumber(pname & " (" & pcode & ")" & vbLf & T("~45~48~2C~48~2F~CC: ") & stock & vbLf & vbLf & T("~2A~39~2F~27~2F ~41~31~48~34:"), T("~41~31~48~34 ~42~37~39~47"), 1)
    If VarType(qty) = vbBoolean Then Exit Sub
    If CDbl(qty) <= 0 Then Msg T("~2A~39~2F~27~2F ~28~27~CC~2F ~28~CC~34~2A~31 ~27~32 ~35~41~31 ~28~27~34~2F."), vbExclamation: Exit Sub
    If CDbl(qty) > stock Then Msg T("~45~48~2C~48~2F~CC ~A9~27~41~CC ~46~CC~33~2A. ~45~48~2C~48~2F~CC ~41~39~44~CC: ") & stock, vbExclamation: Exit Sub
    price = AskNumber(T("~42~CC~45~2A ~41~31~48~34 ~47~31 ~48~27~2D~2F (~2A~48~45~27~46):"), T("~41~31~48~34 ~42~37~39~47"), Num(pr.Cells(1, PC_SELL).Value))
    If VarType(price) = vbBoolean Then Exit Sub
    method = PickMethod()
    If Len(method) = 0 Then Exit Sub
    BeginBusy
    pr.Cells(1, PC_QTY).Value = stock - CDbl(qty)
    AddAccRow TT_INCOME, CAT_PARTSALE, pcode, T("~41~31~48~34 ~42~37~39~47 ") & pname & T(" ^00D7 ") & qty, CDbl(qty) * CDbl(price), method
    Dash_RefreshCore
    EndBusy
    Msg T("~41~31~48~34 ~2B~28~2A ~34~2F: ") & Fmt(CDbl(qty) * CDbl(price)) & T(" ~2A~48~45~27~46") & vbLf & T("~45~48~2C~48~2F~CC ~2C~2F~CC~2F: ") & (stock - CDbl(qty)), vbInformation
    Exit Sub
EH:
    FailWith "Stock_SellPart"
End Sub

Public Sub Buy_Save()
    Dim supp As String, pin As String, method As String, notes As String, qty As Double, price As Double
    Dim pr As Range, lo As ListObject, r As Range, code As String, pname As String, pcode As String, desc As String
    On Error GoTo EH
    Sh("BUY").Activate
    supp = CleanText(NC("mk_buy_Supp").Value)
    pin = CleanText(NC("mk_buy_Part").Value)
    method = CleanText(NC("mk_buy_Method").Value)
    notes = CleanText(NC("mk_buy_Notes").Value)
    If Len(pin) = 0 Then FormError T("~A9~2F ~42~37~39~47 ~31~27 ~27~32 ~41~47~31~33~2A ~27~46~2A~2E~27~28 ~A9~46~CC~2F."), "mk_buy_Part": Exit Sub
    Set pr = FindPartRow(pin)
    If pr Is Nothing Then FormError T("~27~CC~46 ~A9~2F ~2F~31 ~27~46~28~27~31 ~46~CC~33~2A.") & vbLf & T("~27~28~2A~2F~27 ~28~27 ~2F~A9~45~47 ^00AB~42~37~39~47 ~2C~2F~CC~2F^00BB ~42~37~39~47 ~31~27 ~2A~39~31~CC~41 ~A9~46~CC~2F."), "mk_buy_Part": Exit Sub
    If Not CellAmount(NC("mk_buy_Qty").Value, qty) Then FormError T("~2A~39~2F~27~2F ~28~27~CC~2F ~39~2F~2F ~28~27~34~2F."), "mk_buy_Qty": Exit Sub
    If qty <= 0 Then FormError T("~2A~39~2F~27~2F ~2E~31~CC~2F ~28~27~CC~2F ~28~CC~34~2A~31 ~27~32 ~35~41~31 ~28~27~34~2F."), "mk_buy_Qty": Exit Sub
    If Not CellAmount(NC("mk_buy_Price").Value, price) Then FormError T("~42~CC~45~2A ~48~27~2D~2F ~28~27~CC~2F ~39~2F~2F ~28~27~34~2F."), "mk_buy_Price": Exit Sub
    If price < 0 Then FormError T("~42~CC~45~2A ~46~45~CC^200C~2A~48~27~46~2F ~45~46~41~CC ~28~27~34~2F."), "mk_buy_Price": Exit Sub
    If Len(method) = 0 Then method = ListItem("tblPay", 1)
    pcode = CStr(pr.Cells(1, PC_CODE).Value)
    pname = CStr(pr.Cells(1, PC_NAME).Value)

    BeginBusy
    Set lo = GetLO("tblPurchases")
    code = NextSeqCode("B-", 4, "mk_cntBuy", lo, BC_CODE)
    Set r = NewRow(lo)
    PutText r.Cells(1, BC_CODE), code
    PutText r.Cells(1, BC_JDATE), GregorianToJalali(Date)
    r.Cells(1, BC_GDATE).Value = Date
    r.Cells(1, BC_SUPP).Value = supp
    PutText r.Cells(1, BC_PART), pcode
    r.Cells(1, BC_PNAME).Value = pname
    r.Cells(1, BC_QTY).Value = qty
    r.Cells(1, BC_PRICE).Value = price
    r.Cells(1, BC_METHOD).Value = method
    r.Cells(1, BC_NOTES).Value = notes
    ' warehouse
    pr.Cells(1, PC_QTY).Value = Num(pr.Cells(1, PC_QTY).Value) + qty
    pr.Cells(1, PC_BUY).Value = price
    If Len(supp) > 0 Then pr.Cells(1, PC_SUPP).Value = supp
    ' accounting
    desc = T("~2E~31~CC~2F ") & pname & " (" & pcode & T(") ^00D7 ") & qty
    If Len(supp) > 0 Then desc = desc & T(" ~27~32 ") & supp
    If qty * price > 0 Then AddAccRow TT_EXPENSE, CAT_PARTBUY, code, desc, qty * price, method
    ClearBuyForm
    Dash_RefreshCore
    EndBusy
    Sh("BUY").Activate
    NC("mk_buy_Supp").Select
    Msg T("~2E~31~CC~2F ") & code & T(" ~2B~28~2A ~34~2F.") & vbLf & pname & T(": ~45~48~2C~48~2F~CC ~2C~2F~CC~2F ") & pr.Cells(1, PC_QTY).Value & vbLf & _
        T("~47~32~CC~46~47 ~2B~28~2A^200C~34~2F~47: ") & Fmt(qty * price) & T(" ~2A~48~45~27~46"), vbInformation
    Exit Sub
EH:
    FailWith "Buy_Save"
End Sub

Public Sub Buy_Clear()
    On Error GoTo EH
    BeginBusy
    ClearBuyForm
    EndBusy
    Sh("BUY").Activate
    NC("mk_buy_Supp").Select
    Exit Sub
EH:
    FailWith "Buy_Clear"
End Sub

Private Sub ClearBuyForm()
    Dim nms As Variant, i As Long
    nms = Array("mk_buy_Supp", "mk_buy_Part", "mk_buy_Qty", "mk_buy_Price", "mk_buy_Method", "mk_buy_Notes")
    For i = 0 To UBound(nms)
        NC(CStr(nms(i))).MergeArea.ClearContents
    Next i
    NC("mk_buy_Method").Value = ListItem("tblPay", 1)
End Sub

Public Sub Tools_Add()
    Dim nm As Variant, cIdx As Long, lo As ListObject, r As Range, code As String
    On Error GoTo EH
    nm = AskText(T("~46~27~45 ~27~28~32~27~31 (~45~2B~44~27~4B ~47~48~CC~47 ~47~48~27~CC ~AF~31~45):"), T("~27~28~32~27~31 ~2C~2F~CC~2F"))
    If VarType(nm) = vbBoolean Then Exit Sub
    If Len(nm) = 0 Then Msg T("~46~27~45 ~27~28~32~27~31 ~27~44~32~27~45~CC ~27~33~2A."), vbExclamation: Exit Sub
    cIdx = PickIndex("tblToolCat", T("~27~28~32~27~31 ~2C~2F~CC~2F"), T("~2F~33~2A~47 ~27~28~32~27~31:"), 1)
    If cIdx = 0 Then Exit Sub
    BeginBusy
    Set lo = GetLO("tblTools")
    code = NextSeqCode("T-", 3, "mk_cntTool", lo, TC_CODE)
    Set r = NewRow(lo)
    PutText r.Cells(1, TC_CODE), code
    r.Cells(1, TC_NAME).Value = CleanText(nm)
    r.Cells(1, TC_CAT).Value = ListItem("tblToolCat", cIdx)
    r.Cells(1, TC_COND).Value = ListItem("tblToolCond", 1)
    PutText r.Cells(1, TC_PDATE), GregorianToJalali(Date)
    EndBusy
    Sh("TOOLS").Activate
    r.Cells(1, TC_BRAND).Select
    Msg T("~27~28~32~27~31 ~28~27 ~A9~2F ") & code & T(" ~2B~28~2A ~34~2F. ~28~31~46~2F~0C ~45~2D~44~0C ~42~CC~45~2A ~48 ~2A~48~36~CC~2D~27~2A ~31~27 ~2F~31 ~47~45~27~46 ~31~2F~CC~41 ~28~46~48~CC~33~CC~2F."), vbInformation
    Exit Sub
EH:
    FailWith "Tools_Add"
End Sub

Private Function FindPartRow(ByVal codeIn As String) As Range
    Dim lo As ListObject, c As Range, s As String
    s = UCase$(Trim$(ToLatinDigits(codeIn)))
    If Len(s) = 0 Then Exit Function
    If s Like "#*" And DigitsOnly(s) = s Then s = "P-" & Format$(CLng(s), "0000")
    Set lo = GetLO("tblParts")
    If lo.DataBodyRange Is Nothing Then Exit Function
    For Each c In lo.ListColumns(PC_CODE).DataBodyRange.Cells
        If UCase$(CStr(c.Value)) = s And Len(s) > 0 Then
            Set FindPartRow = lo.ListRows(c.Row - lo.HeaderRowRange.Row).Range
            Exit Function
        End If
    Next c
End Function

Private Function FindDeviceRow(ByVal codeIn As String) As Range
    Dim lo As ListObject, c As Range, s As String
    s = UCase$(Trim$(ToLatinDigits(codeIn)))
    If Len(s) = 0 Then Exit Function
    Set lo = GetLO("tblDevices")
    If lo.DataBodyRange Is Nothing Then Exit Function
    For Each c In lo.ListColumns(DC_CODE).DataBodyRange.Cells
        If UCase$(CStr(c.Value)) = s Then
            Set FindDeviceRow = lo.ListRows(c.Row - lo.HeaderRowRange.Row).Range
            Exit Function
        End If
    Next c
End Function

' Builds an InputBox prompt with a short list of parts in stock (kept under 250 characters).
Private Function PartPrompt(ByVal head As String) As String
    Dim lo As ListObject, rw As Range, s As String, ln As String
    s = head & vbLf & T("(~45~2B~44~27~4B P-0001 ~CC~27 ~41~42~37 1)")
    Set lo = GetLO("tblParts")
    If Not lo.DataBodyRange Is Nothing Then
        For Each rw In lo.DataBodyRange.Rows
            If Len(CStr(rw.Cells(1, PC_CODE).Value)) > 0 And Num(rw.Cells(1, PC_QTY).Value) > 0 Then
                ln = vbLf & rw.Cells(1, PC_CODE).Value & "  " & Left$(CStr(rw.Cells(1, PC_NAME).Value), 22) & " (" & rw.Cells(1, PC_QTY).Value & ")"
                If Len(s) + Len(ln) > 245 Then Exit For
                s = s & ln
            End If
        Next rw
    End If
    PartPrompt = s
End Function

'==============================================================================
'  CODES AND COUNTERS
'==============================================================================

' MK-yyyy-0001 with the Jalali year; the counter restarts every year and never reuses a code.
Private Function NextDeviceCode() As String
    Dim jy As Long, n As Long, k As Long, prefix As String, c As Range, lo As ListObject
    jy = JalaliYear(Date)
    prefix = "MK-" & jy & "-"
    If CLng(Num(NC("mk_cntDevYear").Value)) = jy Then n = CLng(Num(NC("mk_cntDev").Value)) Else n = 0
    Set lo = GetLO("tblDevices")
    If Not lo.DataBodyRange Is Nothing Then
        For Each c In lo.ListColumns(DC_CODE).DataBodyRange.Cells
            If Left$(CStr(c.Value), Len(prefix)) = prefix Then
                k = CLng(Val(Mid$(CStr(c.Value), Len(prefix) + 1)))
                If k > n Then n = k
            End If
        Next c
    End If
    n = n + 1
    NC("mk_cntDevYear").Value = jy
    NC("mk_cntDev").Value = n
    NextDeviceCode = prefix & Format$(n, "0000")
End Function

Private Function NextSeqCode(ByVal prefix As String, ByVal digits As Long, ByVal cntName As String, ByVal lo As ListObject, ByVal col As Long) As String
    Dim n As Long, k As Long, c As Range, s As String
    n = CLng(Num(NC(cntName).Value))
    If Not lo.DataBodyRange Is Nothing Then
        For Each c In lo.ListColumns(col).DataBodyRange.Cells
            s = UCase$(CStr(c.Value))
            If Left$(s, Len(prefix)) = prefix Then
                k = CLng(Val(Mid$(s, Len(prefix) + 1)))
                If k > n Then n = k
            End If
        Next c
    End If
    n = n + 1
    NC(cntName).Value = n
    NextSeqCode = prefix & Format$(n, String$(digits, "0"))
End Function

'==============================================================================
'  JALALI (SHAMSI) DATES
'==============================================================================

' Standard Gregorian -> Jalali conversion (correct for leap years).
Private Sub JalaliParts(ByVal d As Date, ByRef jy As Long, ByRef jm As Long, ByRef jd As Long)
    Dim gy As Long, gm As Long, gd As Long, gy2 As Long, days As Long, gdm As Variant
    gdm = Array(0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334)
    gy = Year(d): gm = Month(d): gd = Day(d)
    If gm > 2 Then gy2 = gy + 1 Else gy2 = gy
    days = 355666 + 365& * gy + (gy2 + 3) \ 4 - (gy2 + 99) \ 100 + (gy2 + 399) \ 400 + gd + gdm(gm - 1)
    jy = -1595 + 33 * (days \ 12053)
    days = days Mod 12053
    jy = jy + 4 * (days \ 1461)
    days = days Mod 1461
    If days > 365 Then
        jy = jy + (days - 1) \ 365
        days = (days - 1) Mod 365
    End If
    If days < 186 Then
        jm = 1 + days \ 31
        jd = 1 + days Mod 31
    Else
        jm = 7 + (days - 186) \ 30
        jd = 1 + (days - 186) Mod 30
    End If
End Sub

' Returns a Jalali date as yyyy/mm/dd. Also usable in cells: =JalaliDate(TODAY())
Public Function JalaliDate(Optional ByVal d As Variant) As String
    If IsMissing(d) Then d = Date
    If IsEmpty(d) Then Exit Function
    If Not IsDate(d) And Not IsNumeric(d) Then Exit Function
    If CDbl(CDate(d)) <= 0 Then Exit Function
    JalaliDate = GregorianToJalali(CDate(d))
End Function

Public Function GregorianToJalali(ByVal d As Date) As String
    Dim jy As Long, jm As Long, jd As Long
    JalaliParts d, jy, jm, jd
    GregorianToJalali = jy & "/" & Format$(jm, "00") & "/" & Format$(jd, "00")
End Function

' Current Jalali date and time from the system clock: yyyy/mm/dd HH:MM
Public Function JalaliNow() As String
    JalaliNow = GregorianToJalali(Date) & " " & HHMM()
End Function

Private Function JalaliYear(ByVal d As Date) As Long
    Dim jy As Long, jm As Long, jd As Long
    JalaliParts d, jy, jm, jd
    JalaliYear = jy
End Function

Private Function HHMM() As String
    Dim tm As Date
    tm = Now
    HHMM = Format$(Hour(tm), "00") & ":" & Format$(Minute(tm), "00")
End Function

'==============================================================================
'  TEXT, NUMBERS, PHONES
'==============================================================================

' Decodes the ASCII-encoded Persian strings of this module (see header comment).
Private Function T(ByVal s As String) As String
    Dim i As Long, n As Long, ch As String, out As String
    n = Len(s)
    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = "~" Then
            out = out & ChrW$(&H600& + CLng("&H" & Mid$(s, i + 1, 2)))
            i = i + 3
        ElseIf ch = "^" Then
            out = out & ChrW$(CLng("&H" & Mid$(s, i + 1, 4)))
            i = i + 5
        Else
            out = out & ch
            i = i + 1
        End If
    Loop
    T = out
End Function

Private Function ToLatinDigits(ByVal s As String) As String
    Dim i As Long
    For i = 0 To 9
        s = Replace(s, ChrW$(&H6F0 + i), CStr(i))
        s = Replace(s, ChrW$(&H660 + i), CStr(i))
    Next i
    ToLatinDigits = s
End Function

' Unifies Arabic/Persian Yeh and Kaf and digits so searches match however the text was typed.
Private Function NormalizeFa(ByVal s As String) As String
    s = Replace(s, ChrW$(&H64A), ChrW$(&H6CC))
    s = Replace(s, ChrW$(&H649), ChrW$(&H6CC))
    s = Replace(s, ChrW$(&H643), ChrW$(&H6A9))
    NormalizeFa = Trim$(ToLatinDigits(s))
End Function

Private Function CleanText(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    CleanText = Trim$(Replace(Replace(CStr(v), ChrW$(&H64A), ChrW$(&H6CC)), ChrW$(&H643), ChrW$(&H6A9)))
End Function

Private Function DigitsOnly(ByVal s As String) As String
    Dim i As Long, ch As String
    s = ToLatinDigits(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then DigitsOnly = DigitsOnly & ch
    Next i
End Function

' 0912 123 4567, +98912..., 98912..., 912... -> 09121234567
Private Function NormalizePhone(ByVal s As String) As String
    s = DigitsOnly(s)
    If Left$(s, 4) = "0098" Then s = "0" & Mid$(s, 5)
    If Left$(s, 2) = "98" And Len(s) = 12 Then s = "0" & Mid$(s, 3)
    If Left$(s, 1) = "9" And Len(s) = 10 Then s = "0" & s
    NormalizePhone = s
End Function

Private Function IsValidMobile(ByVal s As String) As Boolean
    IsValidMobile = (s Like "09#########")
End Function

Private Function Num(ByVal v As Variant) As Double
    Dim d As Double
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If CellAmount(v, d) Then Num = d
End Function

' Reads a money/quantity value typed with Latin or Persian digits, with or without separators.
Private Function CellAmount(ByVal v As Variant, ByRef d As Double) As Boolean
    Dim s As String, i As Long, ch As String, clean As String, dots As Long
    d = 0
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then CellAmount = True: Exit Function
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency, vbDecimal, vbByte
            d = CDbl(v): CellAmount = True: Exit Function
    End Select
    s = ToLatinDigits(CStr(v))
    s = Replace(s, ",", ""): s = Replace(s, ChrW$(&H66C), ""): s = Replace(s, ChrW$(&H60C), "")
    s = Replace(s, " ", ""): s = Replace(s, ChrW$(160), "")
    s = Replace(s, T("~2A~48~45~27~46"), ""): s = Replace(s, T("~31~CC~27~44"), "")
    If Len(s) = 0 Then CellAmount = True: Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            clean = clean & ch
        ElseIf ch = "." Or ch = ChrW$(&H66B) Or ch = "/" Then
            dots = dots + 1
            clean = clean & "."
        ElseIf ch = "-" And i = 1 Then
            clean = clean & ch
        Else
            Exit Function
        End If
    Next i
    If dots > 1 Then Exit Function
    d = Val(clean)
    CellAmount = True
End Function

Private Function BalanceText(ByVal x As Double) As String
    If x < 0 Then
        BalanceText = T("~45~34~2A~31~CC ") & Fmt(-x) & T(" ~2A~48~45~27~46 ~28~33~2A~27~46~A9~27~31 ~27~33~2A")
    Else
        BalanceText = T("~45~27~46~2F~47: ") & Fmt(x) & T(" ~2A~48~45~27~46")
    End If
End Function

Private Function Fmt(ByVal x As Double) As String
    Fmt = Format$(x, "#,##0")
End Function

'==============================================================================
'  DIALOGS (Unicode-safe)
'==============================================================================

Private Function Msg(ByVal text As String, Optional ByVal buttons As Long = vbInformation, Optional ByVal title As String = "") As Long
    If Len(title) = 0 Then title = T("~45~48~A9~48~46~CC~A9")
    ' &H80000 = right-aligned text, &H100000 = right-to-left reading
    Msg = MessageBoxW(Application.hWnd, StrPtr(text), StrPtr(title), buttons Or &H80000 Or &H100000)
End Function

' Returns the typed text, or False when the user pressed Cancel.
Private Function AskText(ByVal prompt As String, ByVal title As String, Optional ByVal dflt As String = "") As Variant
    Dim v As Variant
    v = Application.InputBox(Prompt:=Left$(prompt, 250), Title:=title, Default:=dflt, Type:=2)
    If VarType(v) = vbBoolean Then
        AskText = False
    Else
        AskText = Trim$(CStr(v))
    End If
End Function

' Returns a number (Double), or False when the user pressed Cancel.
Private Function AskNumber(ByVal prompt As String, ByVal title As String, ByVal dflt As Double) As Variant
    Dim v As Variant, d As Double
    Do
        v = AskText(prompt, title, CStr(dflt))
        If VarType(v) = vbBoolean Then AskNumber = False: Exit Function
        If CellAmount(CStr(v), d) Then AskNumber = d: Exit Function
        Msg T("~44~37~41~27~4B ~CC~A9 ~39~2F~2F ~45~39~2A~28~31 ~48~27~31~2F ~A9~46~CC~2F (~45~2B~44~27~4B 250000)."), vbExclamation
    Loop
End Function

' Shows a numbered list from a settings table and returns the chosen position (0 = cancel).
Private Function PickIndex(ByVal tblName As String, ByVal title As String, ByVal head As String, Optional ByVal dflt As Long = 1) As Long
    Dim items As Collection, i As Long, p As String, v As Variant, d As Double
    Set items = ListItems(tblName)
    If items.Count = 0 Then Exit Function
    p = head
    For i = 1 To items.Count
        p = p & vbLf & i & " - " & items(i)
    Next i
    If dflt < 1 Or dflt > items.Count Then dflt = 1
    Do
        v = AskText(p, title, CStr(dflt))
        If VarType(v) = vbBoolean Then Exit Function
        If CellAmount(CStr(v), d) Then
            If d >= 1 And d <= items.Count And d = Int(d) Then PickIndex = CLng(d): Exit Function
        End If
        Msg T("~CC~A9 ~39~2F~2F ~28~CC~46 1 ~48 ") & items.Count & T(" ~48~27~31~2F ~A9~46~CC~2F."), vbExclamation
    Loop
End Function

Private Function PickMethod() As String
    Dim i As Long
    i = PickIndex("tblPay", T("~31~48~34 ~7E~31~2F~27~2E~2A"), T("~31~48~34 ~7E~31~2F~27~2E~2A:"), 1)
    If i > 0 Then PickMethod = ListItem("tblPay", i)
End Function

Private Sub FormError(ByVal text As String, ByVal cellName As String)
    Msg text, vbExclamation, T("~27~37~44~27~39~27~2A ~46~27~42~35")
    On Error Resume Next
    NC(cellName).Parent.Activate
    NC(cellName).Select
End Sub

Private Sub FailWith(ByVal procName As String)
    Dim d As String
    d = Err.Number & " - " & Err.Description
    If Len(mStep) > 0 Then d = d & vbLf & "(step: " & mStep & ")"
    mStep = ""
    EndBusy
    Msg T("~2E~37~27~CC ~3A~CC~31~45~46~2A~38~31~47 ~2F~31 ") & procName & ":" & vbLf & d, vbCritical, T("~2E~37~27")
End Sub

'==============================================================================
'  WORKBOOK HELPERS
'==============================================================================

Private Sub BeginBusy(Optional ByVal manualCalc As Boolean = False)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    If manualCalc Then Application.Calculation = xlCalculationManual
    UnprotectAll
End Sub

Private Sub EndBusy()
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    ProtectAll
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Function ProtectedKeys() As Variant
    ProtectedKeys = Array("DASH", "INTAKE", "SEARCH", "DEV", "CUST", "ACC", "STOCK", "BUY", "TOOLS")
End Function

Private Sub UnprotectAll()
    Dim k As Variant
    On Error Resume Next
    For Each k In ProtectedKeys()
        Sh(CStr(k)).Unprotect PROTECT_PWD
    Next k
End Sub

Private Sub ProtectAll()
    Dim k As Variant, ws As Worksheet
    On Error Resume Next
    For Each k In ProtectedKeys()
        Set ws = Nothing
        Set ws = Sh(CStr(k))
        If Not ws Is Nothing Then
            ws.Protect Password:=PROTECT_PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True, AllowSorting:=True
            If k = "INTAKE" Then ws.EnableSelection = xlUnlockedCells Else ws.EnableSelection = xlNoRestrictions
        End If
    Next k
End Sub

Private Function Sh(ByVal key As String) As Worksheet
    Set Sh = ThisWorkbook.Worksheets(SN(key))
End Function

' Workbook-level named cell.
Private Function NC(ByVal nm As String) As Range
    Set NC = ThisWorkbook.Names(nm).RefersToRange
End Function

Private Sub NameCell(ByVal nm As String, ByVal rng As Range)
    ThisWorkbook.Names.Add Name:=nm, RefersTo:="='" & rng.Worksheet.Name & "'!" & rng.Address
End Sub

Private Function GetLO(ByVal tblName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tblName, vbTextCompare) = 0 Then Set GetLO = lo: Exit Function
        Next lo
    Next ws
    Err.Raise vbObjectError + 513, "Mokonic", "Table not found: " & tblName & " - run Setup_Mokonic."
End Function

' Returns a row to write into: reuses the single blank starter row, otherwise appends.
Private Function NewRow(ByVal lo As ListObject) As Range
    Dim r As Range
    If lo.ListRows.Count = 1 Then
        If Len(CStr(lo.DataBodyRange.Cells(1, 1).Value)) = 0 Then Set r = lo.ListRows(1).Range
    End If
    If r Is Nothing Then Set r = lo.ListRows.Add.Range
    ApplyRowLocks lo, r
    Set NewRow = r
End Function

Private Sub ApplyRowLocks(ByVal lo As ListObject, ByVal r As Range)
    Dim cols As Variant, i As Long
    r.Locked = True
    cols = UnlockedCols(lo.Name)
    For i = LBound(cols) To UBound(cols)
        r.Cells(1, cols(i)).Locked = False
    Next i
End Sub

Private Sub ApplyLockPattern(ByVal lo As ListObject)
    Dim i As Long
    For i = 1 To lo.ListRows.Count
        ApplyRowLocks lo, lo.ListRows(i).Range
    Next i
End Sub

' Columns the user may type into while the sheet is protected.
Private Function UnlockedCols(ByVal tblName As String) As Variant
    Select Case tblName
        Case "tblDevices"
            UnlockedCols = Array(DC_TYPE, DC_BRAND, DC_SERIAL, DC_PROBLEM, DC_ACCESS, DC_COND, DC_SHELF, DC_STATUS, _
                                 DC_TECH, DC_DIAG, DC_PARTS, DC_ESTIMATE, DC_PARTSCOST, DC_LABOR, DC_STAFF, DC_NOTES)
        Case "tblCustomers"
            UnlockedCols = Array(CC_NAME, CC_PHONE, CC_PHONE2, CC_ADDR, CC_NOTES)
        Case "tblAcc"
            UnlockedCols = Array(AC_TYPE, AC_CAT, AC_REF, AC_DESC, AC_AMOUNT, AC_METHOD)
        Case "tblParts"
            UnlockedCols = Array(PC_NAME, PC_CAT, PC_SPEC, PC_SHELF, PC_QTY, PC_UNIT, PC_MIN, PC_BUY, PC_SELL, PC_SUPP)
        Case "tblPurchases"
            UnlockedCols = Array(BC_NOTES)
        Case "tblTools"
            UnlockedCols = Array(TC_NAME, TC_CAT, TC_BRAND, TC_LOC, TC_COND, TC_PDATE, TC_PRICE, TC_HOLDER, TC_NOTES)
        Case Else
            UnlockedCols = Array()
    End Select
End Function

Private Function ListItems(ByVal tblName As String) As Collection
    Dim c As Range, col As New Collection, lo As ListObject
    Set lo = GetLO(tblName)
    If Not lo.DataBodyRange Is Nothing Then
        For Each c In lo.ListColumns(1).DataBodyRange.Cells
            If Len(Trim$(CStr(c.Value))) > 0 Then col.Add CStr(c.Value)
        Next c
    End If
    Set ListItems = col
End Function

Private Function ListItem(ByVal tblName As String, ByVal i As Long) As String
    Dim items As Collection
    Set items = ListItems(tblName)
    If i >= 1 And i <= items.Count Then ListItem = items(i)
End Function

Private Function ListIndexOf(ByVal tblName As String, ByVal s As String) As Long
    Dim items As Collection, i As Long
    Set items = ListItems(tblName)
    s = Trim$(s)
    For i = 1 To items.Count
        If items(i) = s Then ListIndexOf = i: Exit Function
    Next i
End Function

Private Function StatusName(ByVal i As Long) As String
    StatusName = ListItem("tblStatus", i)
End Function

Private Function StatusIndex(ByVal s As String) As Long
    StatusIndex = ListIndexOf("tblStatus", s)
End Function

' Structured reference to a whole column, e.g. tblDevices[Status]
Private Function SRc(ByVal lo As ListObject, ByVal c As Long) As String
    SRc = lo.Name & "[" & lo.ListColumns(c).Name & "]"
End Function

' Structured reference to this row, e.g. [@[Status]]
Private Function TRc(ByVal lo As ListObject, ByVal c As Long) As String
    TRc = "[@[" & lo.ListColumns(c).Name & "]]"
End Function

' Wraps a formula so it stays blank while the row's first column is empty.
Private Function IfCode(ByVal lo As ListObject, ByVal f As String) As String
    IfCode = "=IF(" & TRc(lo, 1) & "="""",""""," & f & ")"
End Function

Private Sub SetColFormula(ByVal lo As ListObject, ByVal c As Long, ByVal f As String)
    lo.ListColumns(c).DataBodyRange.Formula = f
End Sub

Private Sub PutText(ByVal cell As Range, ByVal s As String)
    cell.NumberFormat = "@"
    cell.Value = s
End Sub

'==============================================================================
'  FORMATTING HELPERS
'==============================================================================

' Creates a table at topLeft from headers (+ nData rows already written below them).
Private Function MakeTable(ByVal ws As Worksheet, ByVal topLeft As Range, ByVal tblName As String, ByVal headers As Variant, _
                           ByVal widths As Variant, ByVal styleName As String, ByVal nData As Long) As ListObject
    Dim n As Long, i As Long, lo As ListObject
    n = UBound(headers) - LBound(headers) + 1
    For i = 0 To n - 1
        topLeft.Offset(0, i).Value = headers(LBound(headers) + i)
        If i <= UBound(widths) Then topLeft.Offset(0, i).EntireColumn.ColumnWidth = widths(i)
    Next i
    If nData < 1 Then nData = 1
    Set lo = ws.ListObjects.Add(xlSrcRange, topLeft.Resize(nData + 1, n), , xlYes)
    lo.Name = tblName
    lo.TableStyle = styleName
    lo.ShowTableStyleRowStripes = True
    With lo.HeaderRowRange
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .RowHeight = 34
    End With
    lo.DataBodyRange.VerticalAlignment = xlCenter
    Set MakeTable = lo
End Function

Private Sub ColFmt(ByVal lo As ListObject, ByVal cols As Variant, Optional ByVal numFmt As String = "", _
                   Optional ByVal center As Boolean = False, Optional ByVal wrap As Boolean = False, Optional ByVal hide As Boolean = False)
    Dim i As Long, rng As Range
    For i = LBound(cols) To UBound(cols)
        Set rng = lo.ListColumns(cols(i)).DataBodyRange
        If Len(numFmt) > 0 Then rng.NumberFormat = numFmt
        If center Then rng.HorizontalAlignment = xlCenter
        If wrap Then rng.WrapText = True
        If hide Then rng.EntireColumn.Hidden = True
    Next i
End Sub

Private Sub AddListDV(ByVal rng As Range, ByVal listName As String, Optional ByVal strict As Boolean = True)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & listName
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = strict
        .ErrorTitle = T("~45~48~A9~48~46~CC~A9")
        .ErrorMessage = T("~44~37~41~27~4B ~CC~A9~CC ~27~32 ~AF~32~CC~46~47^200C~47~27~CC ~41~47~31~33~2A ~31~27 ~27~46~2A~2E~27~28 ~A9~46~CC~2F. ~41~47~31~33~2A^200C~47~27 ~2F~31 ~28~31~AF~47 ^00AB~2A~46~38~CC~45~27~2A^00BB ~42~27~28~44 ~48~CC~31~27~CC~34^200C~27~46~2F.")
    End With
End Sub

' Conditional format by formula. The formula is relative to the first cell of rng.
Private Sub AddCF(ByVal rng As Range, ByVal f As String, ByVal fillColor As Long, ByVal fontColor As Long)
    Dim fc As FormatCondition
    rng.Worksheet.Activate
    rng.Cells(1, 1).Select
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=f)
    fc.Interior.Color = fillColor
    fc.Font.Color = fontColor
    fc.StopIfTrue = False
End Sub

' Dark header band with the shop name and the sheet title on row 1.
Private Sub HeaderBand(ByVal ws As Worksheet, ByVal title As String, ByVal lastCol As Long)
    ws.Rows(1).RowHeight = 40
    ws.Rows(2).RowHeight = 32
    ws.Rows(3).RowHeight = 20
    With ws.Range(ws.Cells(1, 2), ws.Cells(1, lastCol))
        .Merge
        .Formula = "=mk_ShopName&""   |   " & title & """"
        .Interior.Color = CLR("navy")
        .Font.Color = vbWhite
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub BackButton(ByVal ws As Worksheet)
    AddBtn ws, ws.Range("B2").Left, ws.Range("B2").Top + 3, 112, 26, T("~28~27~32~AF~34~2A ~28~47 ~2F~27~34~28~48~31~2F"), "Go_Dashboard", CLR("navy"), , 9
End Sub

Private Function AddBtn(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                        ByVal caption As String, ByVal macro As String, ByVal fillColor As Long, _
                        Optional ByVal fontColor As Long = 16777215, Optional ByVal fontSize As Single = 10, _
                        Optional ByVal shpName As String = "") As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(5, x, y, w, h)          ' 5 = rounded rectangle
    If Len(shpName) > 0 Then shp.Name = shpName
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.Visible = 0
    On Error Resume Next
    shp.Adjustments.Item(1) = 0.25
    On Error GoTo 0
    With shp.TextFrame2
        .WordWrap = -1
        .VerticalAnchor = 3                              ' middle
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 1: .MarginBottom = 1
        .TextRange.Text = caption
        .TextRange.ParagraphFormat.Alignment = 2         ' center
        With .TextRange.Font
            .Name = FONT_NAME
            .NameComplexScript = FONT_NAME
            .Size = fontSize
            .Bold = -1
            .Fill.ForeColor.RGB = fontColor
        End With
    End With
    On Error Resume Next
    shp.TextFrame2.TextRange.ParagraphFormat.TextDirection = 2   ' right-to-left
    On Error GoTo 0
    shp.Placement = xlFreeFloating
    shp.OnAction = macro
    Set AddBtn = shp
End Function

Private Sub SectionTitle(ByVal rng As Range, ByVal text As String)
    With rng
        .Merge
        .Value = text
        .Interior.Color = CLR("teal")
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .RowHeight = 24
    End With
End Sub

Private Sub TableHead(ByVal rng As Range, ByVal heads As Variant)
    Dim i As Long
    For i = 0 To UBound(heads)
        rng.Cells(1, i + 1).Value = heads(i)
    Next i
    With rng
        .Interior.Color = CLR("navy2")
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .RowHeight = 22
    End With
End Sub

Private Sub BodyBox(ByVal rng As Range)
    With rng
        .Borders(xlInsideHorizontal).Color = RGB(220, 220, 220)
        .Borders(xlInsideVertical).Color = RGB(220, 220, 220)
        .BorderAround xlContinuous, xlThin, , RGB(180, 180, 180)
    End With
End Sub

Private Sub FormLabel(ByVal rng As Range, ByVal text As String)
    With rng
        If .Cells.Count > 1 Then .Merge
        .Value = text
        .Interior.Color = CLR("tealLight")
        .Font.Bold = True
        .Font.Color = CLR("navy")
        .WrapText = True
        .IndentLevel = 1
    End With
End Sub

Private Sub StyleInput(ByVal rng As Range, Optional ByVal editable As Boolean = True)
    With rng
        If .Cells.Count > 1 Then .Merge
        .Interior.Color = CLR("input")
        .BorderAround xlContinuous, xlThin, , RGB(170, 170, 170)
        .Font.Size = 11
        .Locked = Not editable
        .IndentLevel = 1
    End With
End Sub

Private Sub Note(ByVal rng As Range, ByVal text As String)
    With rng
        .Merge
        .Value = text
        .WrapText = True
        .Font.Size = 8.5
        .Font.Color = RGB(100, 100, 100)
        .VerticalAlignment = xlCenter
    End With
End Sub

' Settings key/value row with a workbook name on the value cell.
Private Sub KV(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal v As Variant, ByVal nm As String, ByVal asText As Boolean)
    FormLabel ws.Cells(r, 2), label
    If asText Then PutText ws.Cells(r, 3), CStr(v) Else ws.Cells(r, 3).Value = v
    StyleInput ws.Cells(r, 3)
    ws.Rows(r).RowHeight = 22
    NameCell nm, ws.Cells(r, 3)
End Sub

'==============================================================================
'  NAMES, HEADERS AND COLORS
'==============================================================================

Private Function SheetKeys() As Variant
    SheetKeys = Array("DASH", "INTAKE", "DEV", "SEARCH", "CUST", "ACC", "STOCK", "BUY", "TOOLS", "SHELF", "SET")
End Function

Private Function SN(ByVal key As String) As String
    Select Case key
        Case "DASH": SN = T("~2F~27~34~28~48~31~2F")
        Case "INTAKE": SN = T("~7E~30~CC~31~34")
        Case "DEV": SN = T("~2F~33~2A~AF~27~47^200C~47~27")
        Case "SEARCH": SN = T("~2C~33~2A~2C~48")
        Case "CUST": SN = T("~45~34~2A~31~CC~27~46")
        Case "ACC": SN = T("~2D~33~27~28~2F~27~31~CC")
        Case "STOCK": SN = T("~27~46~28~27~31")
        Case "BUY": SN = T("~2E~31~CC~2F")
        Case "TOOLS": SN = T("~27~28~32~27~31")
        Case "SHELF": SN = T("~42~41~33~47^200C~47~27")
        Case "SET": SN = T("~2A~46~38~CC~45~27~2A")
    End Select
End Function

Private Function SheetTitle(ByVal key As String) As String
    Select Case key
        Case "INTAKE": SheetTitle = T("~7E~30~CC~31~34 ~33~31~CC~39 ~2F~33~2A~AF~27~47")
        Case "DEV": SheetTitle = T("~2F~33~2A~AF~27~47^200C~47~27 ~48 ~48~36~39~CC~2A ~2A~39~45~CC~31")
        Case "SEARCH": SheetTitle = T("~2C~33~2A~2C~48~CC ~2F~33~2A~AF~27~47 ~48 ~45~34~2A~31~CC")
        Case "CUST": SheetTitle = T("~45~34~2A~31~CC~27~46")
        Case "ACC": SheetTitle = T("~2D~33~27~28~2F~27~31~CC - ~2F~31~22~45~2F ~48 ~47~32~CC~46~47")
        Case "STOCK": SheetTitle = T("~27~46~28~27~31 ~42~37~39~27~2A")
        Case "BUY": SheetTitle = T("~2E~31~CC~2F ~42~37~39~27~2A")
        Case "TOOLS": SheetTitle = T("~27~28~32~27~31 ~A9~27~31~AF~27~47")
        Case "SHELF": SheetTitle = T("~42~41~33~47^200C~47~27")
        Case "SET": SheetTitle = T("~2A~46~38~CC~45~27~2A ~48 ~41~47~31~33~2A^200C~47~27")
        Case Else: SheetTitle = T("~2F~27~34~28~48~31~2F")
    End Select
End Function

Private Function SheetTabColor(ByVal key As String) As Long
    Select Case key
        Case "DASH": SheetTabColor = CLR("navy")
        Case "INTAKE": SheetTabColor = CLR("green")
        Case "DEV": SheetTabColor = CLR("teal")
        Case "SEARCH": SheetTabColor = CLR("purple")
        Case "CUST": SheetTabColor = CLR("blue")
        Case "ACC": SheetTabColor = CLR("gold")
        Case "STOCK": SheetTabColor = CLR("orange")
        Case "BUY": SheetTabColor = CLR("brown")
        Case "TOOLS": SheetTabColor = CLR("gray")
        Case "SHELF": SheetTabColor = CLR("olive")
        Case Else: SheetTabColor = CLR("dark")
    End Select
End Function

Private Function DevHeaders() As Variant
    DevHeaders = Array(T("~A9~2F ~2F~33~2A~AF~27~47"), T("~2A~27~31~CC~2E ~7E~30~CC~31~34"), T("~33~27~39~2A"), T("~2A~27~31~CC~2E ~45~CC~44~27~2F~CC"), T("~A9~2F ~45~34~2A~31~CC"), T("~46~27~45 ~45~34~2A~31~CC"), _
        T("~2A~44~41~46"), T("~46~48~39 ~2F~33~2A~AF~27~47"), T("~28~31~46~2F / ~45~2F~44"), T("~33~31~CC~27~44"), T("~27~CC~31~27~2F ~27~39~44~27~45^200C~34~2F~47"), T("~44~48~27~32~45 ~47~45~31~27~47"), _
        T("~48~36~39~CC~2A ~38~27~47~31~CC"), T("~42~41~33~47"), T("~48~36~39~CC~2A"), T("~2A~A9~46~33~CC~46"), T("~39~CC~28^200C~CC~27~28~CC"), T("~42~37~39~27~2A ~45~35~31~41~CC"), _
        T("~28~31~22~48~31~2F ~47~32~CC~46~47"), T("~47~32~CC~46~47 ~42~37~39~27~2A"), T("~2F~33~2A~45~32~2F"), T("~2C~45~39 ~47~32~CC~46~47"), T("~45~28~44~3A ~7E~31~2F~27~2E~2A~CC"), T("~45~27~46~2F~47"), _
        T("~22~2E~31~CC~46 ~2A~3A~CC~CC~31 ~48~36~39~CC~2A"), T("~2A~27~31~CC~2E ~2A~2D~48~CC~44"), T("~AF~27~31~27~46~2A~CC ~2A~27"), T("~7E~30~CC~31~46~2F~47"), T("~2A~48~36~CC~2D~27~2A"), T("~2A~2D~48~CC~44 ~45~CC~44~27~2F~CC"))
End Function

Private Function CustHeaders() As Variant
    CustHeaders = Array(T("~A9~2F ~45~34~2A~31~CC"), T("~46~27~45 ~45~34~2A~31~CC"), T("~2A~44~41~46 ~47~45~31~27~47"), T("~2A~44~41~46 ~2F~48~45"), T("~22~2F~31~33"), _
        T("~27~48~44~CC~46 ~45~31~27~2C~39~47"), T("~2A~39~2F~27~2F ~2F~33~2A~AF~27~47^200C~47~27"), T("~2C~45~39 ~7E~31~2F~27~2E~2A~CC"), T("~2A~48~36~CC~2D~27~2A"))
End Function

Private Function AccHeaders() As Variant
    AccHeaders = Array(T("~2A~27~31~CC~2E"), T("~2A~27~31~CC~2E ~45~CC~44~27~2F~CC"), T("~46~48~39"), T("~2F~33~2A~47"), T("~A9~2F ~45~31~2A~28~37"), T("~34~31~2D"), T("~45~28~44~3A"), T("~31~48~34 ~7E~31~2F~27~2E~2A"), T("~45~27~47"))
End Function

Private Function PartHeaders() As Variant
    PartHeaders = Array(T("~A9~2F ~42~37~39~47"), T("~46~27~45 ~42~37~39~47"), T("~2F~33~2A~47"), T("~45~34~2E~35~27~2A / ~45~42~2F~27~31"), T("~42~41~33~47"), T("~45~48~2C~48~2F~CC"), _
        T("~48~27~2D~2F"), T("~2D~2F~27~42~44 ~45~48~2C~48~2F~CC"), T("~42~CC~45~2A ~2E~31~CC~2F"), T("~42~CC~45~2A ~41~31~48~34"), T("~27~31~32~34 ~45~48~2C~48~2F~CC"), T("~2A~23~45~CC~46^200C~A9~46~46~2F~47"))
End Function

Private Function BuyHeaders() As Variant
    BuyHeaders = Array(T("~A9~2F ~2E~31~CC~2F"), T("~2A~27~31~CC~2E"), T("~2A~23~45~CC~46^200C~A9~46~46~2F~47"), T("~A9~2F ~42~37~39~47"), T("~46~27~45 ~42~37~39~47"), T("~2A~39~2F~27~2F"), _
        T("~42~CC~45~2A ~48~27~2D~2F"), T("~2C~45~39"), T("~31~48~34 ~7E~31~2F~27~2E~2A"), T("~2A~48~36~CC~2D~27~2A"), T("~2A~27~31~CC~2E ~45~CC~44~27~2F~CC"))
End Function

Private Function ToolHeaders() As Variant
    ToolHeaders = Array(T("~A9~2F ~27~28~32~27~31"), T("~46~27~45 ~27~28~32~27~31"), T("~2F~33~2A~47"), T("~28~31~46~2F"), T("~45~2D~44 ~46~AF~47~2F~27~31~CC"), T("~48~36~39~CC~2A"), _
        T("~2A~27~31~CC~2E ~2E~31~CC~2F"), T("~42~CC~45~2A"), T("~2F~31 ~27~2E~2A~CC~27~31"), T("~2A~48~36~CC~2D~27~2A"))
End Function

Private Function ShelfHeaders() As Variant
    ShelfHeaders = Array(T("~A9~2F ~42~41~33~47"), T("~2A~48~36~CC~2D~27~2A"), T("~46~48~39"), T("~38~31~41~CC~2A"), T("~2F~33~2A~AF~27~47^200C~47~27~CC ~45~48~2C~48~2F"), T("~42~37~39~27~2A (~31~2F~CC~41)"), T("~27~28~32~27~31~47~27"), T("~38~31~41~CC~2A ~22~32~27~2F"))
End Function

Private Function CLR(ByVal key As String) As Long
    Select Case key
        Case "navy": CLR = RGB(22, 54, 92)
        Case "navy2": CLR = RGB(41, 84, 138)
        Case "teal": CLR = RGB(0, 121, 121)
        Case "tealLight": CLR = RGB(224, 242, 241)
        Case "green": CLR = RGB(46, 125, 50)
        Case "red": CLR = RGB(198, 40, 40)
        Case "orange": CLR = RGB(230, 115, 0)
        Case "purple": CLR = RGB(106, 27, 154)
        Case "blue": CLR = RGB(21, 101, 192)
        Case "gray": CLR = RGB(97, 97, 97)
        Case "gold": CLR = RGB(175, 130, 0)
        Case "brown": CLR = RGB(121, 85, 72)
        Case "olive": CLR = RGB(85, 107, 47)
        Case "dark": CLR = RGB(55, 71, 79)
        Case "input": CLR = RGB(255, 253, 231)
        Case Else: CLR = RGB(0, 0, 0)
    End Select
End Function

' Light fill per status (row coloring).
Private Function StatusFill(ByVal i As Long) As Long
    Select Case i
        Case 1: StatusFill = RGB(221, 235, 247)   ' received - blue
        Case 2: StatusFill = RGB(234, 224, 247)   ' diagnosing - light purple
        Case 3: StatusFill = RGB(255, 245, 194)   ' awaiting approval - yellow
        Case 4: StatusFill = RGB(252, 218, 190)   ' waiting for parts - orange
        Case 5: StatusFill = RGB(255, 229, 153)   ' in repair - amber
        Case 6: StatusFill = RGB(214, 236, 207)   ' ready - green
        Case 7: StatusFill = RGB(217, 217, 217)   ' delivered - gray
        Case 8: StatusFill = RGB(248, 210, 213)   ' unrepairable / returned - red
        Case Else: StatusFill = RGB(255, 255, 255)
    End Select
End Function

' Dark text per status (keeps colored rows readable).
Private Function StatusFont(ByVal i As Long) As Long
    Select Case i
        Case 1: StatusFont = RGB(31, 78, 121)
        Case 2: StatusFont = RGB(84, 48, 140)
        Case 3: StatusFont = RGB(110, 84, 0)
        Case 4: StatusFont = RGB(150, 70, 0)
        Case 5: StatusFont = RGB(120, 80, 0)
        Case 6: StatusFont = RGB(40, 90, 30)
        Case 7: StatusFont = RGB(80, 80, 80)
        Case 8: StatusFont = RGB(150, 0, 10)
        Case Else: StatusFont = RGB(0, 0, 0)
    End Select
End Function
