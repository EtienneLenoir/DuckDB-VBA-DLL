Attribute VB_Name = "mDuckNative"
Option Explicit
Option Private Module
#Const VERBOSE = False
Public m_singleton As cDuck
Public gErrDuckDB       As String
Public gDuckDllReady    As Boolean

'==============================================================================
' DUCK VBA DLL — DuckDB bridge for Excel/VBA
' Copyright (c) 2026 Etienne Lenoir
' SPDX-License-Identifier: GPL-3.0-only
' License  : GNU General Public License v3.0 (see LICENSE at repository root)
' Requires Excel 64-bit (VBA7) + duckdb.dll + duckdb_vba_bridge.dll
'==============================================================================

'===== WIN32 =====
Private Declare PtrSafe Function SetDllDirectoryW Lib "kernel32" (ByVal lpPathName As LongPtr) As Long
Private Declare PtrSafe Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As LongPtr) As LongPtr
'===== DuckDB Bridge (x64) =====
'Base
Public Declare PtrSafe Function DuckVba_OpenW Lib "duckdb_vba_bridge.dll" (ByVal pwszPath As LongPtr) As LongPtr
Public Declare PtrSafe Function DuckVba_Close Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr) As Long
Public Declare PtrSafe Function DuckVba_ExecW Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pwszSql As LongPtr) As Long
Public Declare PtrSafe Function DuckVba_OpenReadOnlyW Lib "duckdb_vba_bridge.dll" (ByVal pwszPath As LongPtr) As LongPtr
Public Declare PtrSafe Function Duck_LastErrorW Lib "duckdb_vba_bridge.dll" (ByVal pwszBuf As LongPtr, ByVal cch As Long) As Long
'Select Query Arr
Public Declare PtrSafe Function DuckVba_QueryToArrayFastV Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pwszSelect As LongPtr, ByRef v As Variant) As Long
Public Declare PtrSafe Function DuckVba_SelectToCsvW Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pwszSelect As LongPtr, ByVal pwszCsv As LongPtr) As Long
Public Declare PtrSafe Function DuckVba_SelectShapeW Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pwszSelect As LongPtr, ByRef outRows As Long, ByRef outCols As Long) As Long
Private Declare PtrSafe Function DuckVba_SelectFill2D_TypedV Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pwszSelect As LongPtr, ByRef v As Variant) As Long
Public Declare PtrSafe Function DuckVba_FrameFromValue Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pwszFrame As LongPtr, ByRef v As Variant, ByVal hasHeader As Long, ByVal makeTemp As Long) As Long
Public Declare PtrSafe Function DuckVba_ExecPreparedToArrayV Lib "duckdb_vba_bridge.dll" (ByVal ps As LongPtr, ByRef v As Variant) As Long
'Prepared
Public Declare PtrSafe Function DuckVba_PrepareW Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pwszSql As LongPtr) As LongPtr
Public Declare PtrSafe Function DuckVba_Finalize Lib "duckdb_vba_bridge.dll" (ByVal ps As LongPtr) As Long
Public Declare PtrSafe Function DuckVba_BindVarcharW Lib "duckdb_vba_bridge.dll" (ByVal ps As LongPtr, ByVal idx As Long, ByVal pwsz As LongPtr) As Long
Public Declare PtrSafe Function DuckVba_BindInt64 Lib "duckdb_vba_bridge.dll" (ByVal ps As LongPtr, ByVal idx As Long, ByVal v As LongLong) As Long
Public Declare PtrSafe Function DuckVba_BindDouble Lib "duckdb_vba_bridge.dll" (ByVal ps As LongPtr, ByVal idx As Long, ByVal v As Double) As Long
Public Declare PtrSafe Function DuckVba_ExecPrepared Lib "duckdb_vba_bridge.dll" (ByVal ps As LongPtr) As Long
'Appender / ingestion
Public Declare PtrSafe Function DuckVba_AppendArrayV Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pTable As LongPtr, ByRef v As Variant, ByVal hasHeader As Long) As Long
Public Declare PtrSafe Function DuckVba_AppendAdoRecordset Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pRecordset As LongPtr, ByVal pwszTable As LongPtr, ByVal createIfMissing As Long) As Long
Public Declare PtrSafe Function DuckVba_AppendAdoRecordsetFast Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pRecordset As LongPtr, ByVal pwszTable As LongPtr, ByVal createIfMissing As Long) As Long
'Extensions, Parquet, JSON
Public Declare PtrSafe Function DuckVba_LoadExtW Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pName As LongPtr) As Long
Public Declare PtrSafe Function DuckVba_CopyToParquetW Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pSel As LongPtr, ByVal pOut As LongPtr) As Long
'Temp list  / columns info
Public Declare PtrSafe Function DuckVba_CreateTempListV Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal tabname_w As LongPtr, ByRef keys As Variant, ByVal sqltype_w As LongPtr) As Long
Public Declare PtrSafe Function DuckVba_SelectWithTempList2V Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pTabName As LongPtr, ByRef keys As Variant, ByVal pSqlType As LongPtr, ByVal pSelectOrTable As LongPtr, ByVal pJoinCol As LongPtr, ByVal autoJoin As Long, ByRef vOut As Variant) As Long
Public Declare PtrSafe Function DuckVba_TableInfoV Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pwszSchemaFilter As LongPtr, ByRef v As Variant) As Long
Public Declare PtrSafe Function DuckVba_ColumnsInfoV Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pwszTable As LongPtr, ByRef v As Variant) As Long
Public Declare PtrSafe Function DuckVba_TableExistsW Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal table_path_w As LongPtr) As Long
Public Declare PtrSafe Function DuckVba_ColumnExistsW Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pwszTablePath As LongPtr, ByVal pwszColName As LongPtr) As Long
Public Declare PtrSafe Function DuckVba_RenameTableW Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pwszOldTable As LongPtr, ByVal pwszNewTable As LongPtr) As Long
Public Declare PtrSafe Function DuckVba_RenameColumnW Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pwszTable As LongPtr, ByVal pwszOldCol As LongPtr, ByVal pwszNewCol As LongPtr) As Long
Public Declare PtrSafe Function DuckVba_ScalarV Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pwszSelect As LongPtr, ByRef v As Variant) As Long
'CSV & Update
Public Declare PtrSafe Function DuckVba_ReadCsvToTableW Lib "duckdb_vba_bridge.dll" (ByVal handle As LongPtr, ByVal table_w As LongPtr, ByVal csv_path_w As LongPtr, ByVal create_if_missing As Long) As Long
Public Declare PtrSafe Function DuckVba_UpsertFromArrayV Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pwszTable As LongPtr, ByRef v As Variant, ByVal headerRow As Long, ByVal pwszKeyColsCsv As LongPtr) As Long
'Dict
Public Declare PtrSafe Function DuckVba_SelectToDictW Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pwszSelect As LongPtr, ByVal pwszKeyCol As LongPtr, ByVal pDict As LongPtr, ByVal clearFirst As Long, ByVal onDupMode As Long) As Long
Public Declare PtrSafe Function DuckVba_SelectToDictFlatW Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pwszSelect As LongPtr, ByVal pwszKeyCol As LongPtr, ByVal pwszValCol As LongPtr, ByVal pDict As LongPtr, ByVal clearFirst As Long, ByVal onDupMode As Long) As Long
Public Declare PtrSafe Function DuckVba_SelectToDictValsColsW Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pwszSelect As LongPtr, ByVal pwszKeyCol As LongPtr, ByVal pwszValueColsCsv As LongPtr, ByVal pDict As LongPtr, ByVal clearFirst As Long, ByVal onDupMode As Long) As Long
'Appender low-level (row-by-row)
Public Declare PtrSafe Function DuckVba_AppenderOpen Lib "duckdb_vba_bridge.dll" (ByVal h As LongPtr, ByVal pwszSchema As LongPtr, ByVal pwszTable As LongPtr) As LongPtr
Public Declare PtrSafe Function DuckVba_AppenderClose Lib "duckdb_vba_bridge.dll" (ByVal app As LongPtr) As Long
Public Declare PtrSafe Function DuckVba_AppenderBeginRow Lib "duckdb_vba_bridge.dll" (ByVal app As LongPtr) As Long
Public Declare PtrSafe Function DuckVba_AppenderEndRow Lib "duckdb_vba_bridge.dll" (ByVal app As LongPtr) As Long
Public Declare PtrSafe Function DuckVba_AppendNull Lib "duckdb_vba_bridge.dll" (ByVal app As LongPtr) As Long
Public Declare PtrSafe Function DuckVba_AppendBool Lib "duckdb_vba_bridge.dll" (ByVal app As LongPtr, ByVal b As Long) As Long
Public Declare PtrSafe Function DuckVba_AppendInt64 Lib "duckdb_vba_bridge.dll" (ByVal app As LongPtr, ByVal v As LongLong) As Long
Public Declare PtrSafe Function DuckVba_AppendDouble Lib "duckdb_vba_bridge.dll" (ByVal app As LongPtr, ByVal v As Double) As Long
Public Declare PtrSafe Function DuckVba_AppendVarcharW Lib "duckdb_vba_bridge.dll" (ByVal app As LongPtr, ByVal pwsz As LongPtr) As Long
Public Declare PtrSafe Function DuckVba_AppendDateYMD Lib "duckdb_vba_bridge.dll" (ByVal app As LongPtr, ByVal y As Long, ByVal m As Long, ByVal d As Long) As Long
Public Declare PtrSafe Function DuckVba_AppendTimestampYMDHMSms Lib "duckdb_vba_bridge.dll" (ByVal app As LongPtr, ByVal y As Long, ByVal m As Long, ByVal d As Long, ByVal hh As Long, ByVal mm As Long, ByVal ss As Long, ByVal ms As Long) As Long
Public Declare PtrSafe Function DuckVba_AppendBlob Lib "duckdb_vba_bridge.dll" (ByVal app As LongPtr, ByVal pData As LongPtr, ByVal n As LongLong) As Long

Public Sub EnsureDuckDll(ByVal basePath As String)

    If gDuckDllReady Then Exit Sub
    
    Call SetDllDirectoryW(StrPtr(basePath))
    Call LoadLibraryW(StrPtr(basePath & "\duckdb.dll"))
    Call LoadLibraryW(StrPtr(basePath & "\duckdb_vba_bridge.dll"))
    
    gDuckDllReady = True
    
End Sub

'--- helper : affiche un tableau dans la feuille 1 ---
Sub ShowArrayOnSheet(ByVal a As Variant)
    With ThisWorkbook.Worksheets(1)
        .Cells.Clear
        .Range("A1").Resize(UBound(a, 1), UBound(a, 2)).Value = a
    End With
End Sub

'Retourne une instance unique pratique pour l’appli
Public Property Get CurrentDuckDb() As cDuck
    If m_singleton Is Nothing Then
        Set m_singleton = New cDuck
        m_singleton.Init ThisWorkbook.Path
    End If
    Set CurrentDuckDb = m_singleton
End Property

Public Sub CloseCurrentDuckDb()
    On Error Resume Next
    If Not m_singleton Is Nothing Then
        m_singleton.CloseDuckDb
        Set m_singleton = Nothing
    End If
    On Error GoTo 0
End Sub


'Helper rapide : exécuter une action avec une session jetable
Public Sub UsingSession(ByVal dbPath As String, ByVal action As String)
    ' action = nom d’une Sub publique qui prend (ByRef db As cDuck)
    Dim db As New cDuck
    db.Init ThisWorkbook.Path
    db.OpenDuckDb dbPath
    Application.Run action, db  ' exécute Sub MonCode(db As cDuck)
    db.CloseDuckDb
End Sub

'==== Helpers  ====
Public Function Native_LastErrorText() As String
    Dim buf As String, n As Long
    buf = String$(2048, vbNullChar)
    n = Duck_LastErrorW(StrPtr(buf), Len(buf))
    If n > 0 Then
        Native_LastErrorText = Left$(buf, n)
    Else
        Native_LastErrorText = ""
    End If
End Function

Public Function Duck_LastErrorText() As String
    Dim buf As String, n As Long
    buf = String$(1024, vbNullChar)
    n = Duck_LastErrorW(StrPtr(buf), Len(buf))
    If n > 0 Then Duck_LastErrorText = Left$(buf, n) Else Duck_LastErrorText = ""
End Function

Public Function q(ByVal s As String) As String
    q = "'" & Replace(s, "'", "''") & "'"
End Function

Public Function TrimSQL(ByVal s As String) As String
    Dim q$: q = Trim$(s)
    If Right$(q, 1) = ";" Then q = Left$(q, Len(q) - 1)
    TrimSQL = q
End Function

Public Function SqlQ(ByVal s As String) As String
    SqlQ = "'" & Replace(s, "'", "''") & "'"
End Function

Public Sub ArrayToSheet(ByRef arr As Variant, ByVal Target As Worksheet, Optional ByVal topLeft As String = "A1", Optional BoolNotClear As Boolean)
    If IsEmpty(arr) Then Exit Sub
    
    If BoolNotClear = False Then
        Target.Cells.Clear
    End If
    Target.Range(topLeft).Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
End Sub

Public Sub EnsureFolderExists(ByVal filePath As String)
    On Error Resume Next
    Dim fso As Object, folder As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.GetParentFolderName(filePath)
    If Len(folder) > 0 Then
        If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    End If
End Sub

'Écrit un fichier texte UTF-8 (pour HTML)
Public Sub WriteTextUtf8(ByVal filePath As String, ByVal text As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 'adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText text
    stm.SaveToFile filePath, 2 'adSaveCreateOverWrite
    stm.Close
End Sub

'Map ADO type -> DuckDB SQL type
Private Function AdoToDuckType(ByVal adoType As Long, ByVal size As Long, ByVal prec As Long, ByVal scales As Long) As String
    
    'Quelques valeurs utiles (ADODB.DataTypeEnum)
    Const adBoolean = 11, adTinyInt = 16, adSmallInt = 2, adInteger = 3, adBigInt = 20
    Const adUnsignedTinyInt = 17, adSingle = 4, adDouble = 5, adCurrency = 6
    Const adDecimal = 14, adNumeric = 131
    Const adDate = 7, adDBTimeStamp = 135
    Const adGUID = 72, adVarWChar = 202, adVarChar = 200, adLongVarWChar = 203, adLongVarChar = 201
    Const adLongVarBinary = 205, adVarBinary = 204, adBinary = 128

    Select Case adoType
        Case adBoolean:            AdoToDuckType = "BOOLEAN"
        Case adTinyInt, adSmallInt, adInteger, adUnsignedTinyInt: AdoToDuckType = "INTEGER"
        Case adBigInt:             AdoToDuckType = "BIGINT"
        Case adSingle, adDouble:   AdoToDuckType = "DOUBLE"
        Case adCurrency:           AdoToDuckType = "DECIMAL(19,4)"
        Case adDecimal, adNumeric:
            If prec > 0 Then
                If scales < 0 Then scales = 0
                If prec > 38 Then prec = 38
                AdoToDuckType = "DECIMAL(" & prec & "," & scales & ")"
            Else
                AdoToDuckType = "DECIMAL(18,6)"
            End If
        Case adDate, adDBTimeStamp: AdoToDuckType = "TIMESTAMP"
        Case adGUID:                AdoToDuckType = "UUID"
        Case adVarWChar, adVarChar, adLongVarWChar, adLongVarChar: AdoToDuckType = "VARCHAR"
        Case adVarBinary, adBinary, adLongVarBinary: AdoToDuckType = "BLOB"
        Case Else:                 AdoToDuckType = "VARCHAR"
    End Select
End Function

Public Function PW(ByVal s As String) As LongPtr
    PW = StrPtr(s)
End Function

Public Function PNull() As LongPtr
    PNull = 0
End Function

' ========= PUBLIC API =========
'Affiche un Variant(2D) en mode tableau lisible dans la console (Debug.Print)
'
'Exemple d'appel :
'   DebugArray2DTable preview, "main.clients"
Public Sub DebugArray2DTable(ByRef a As Variant, Optional ByVal label As String = "")

    Dim rLo         As Long, rHi As Long, cLo As Long, cHi As Long, rows As Long, cols As Long
    Dim colWidths() As Long, cellText As String, R As Long, C As Long
    
    On Error GoTo Not2D
    
    If IsEmpty(a) Then
        Debug.Print "DebugArray2DTable "; label; ": (Empty)"
        Exit Sub
    End If
    
    rLo = LBound(a, 1): rHi = UBound(a, 1)
    cLo = LBound(a, 2): cHi = UBound(a, 2)
    rows = rHi - rLo + 1
    cols = cHi - cLo + 1
    
    '1) calcul largeur max par colonne
    ReDim colWidths(1 To cols)
    For C = cLo To cHi
        colWidths(C - cLo + 1) = 0
        For R = rLo To rHi
            cellText = SafeToString(a(R, C))
            If Len(cellText) > colWidths(C - cLo + 1) Then
                colWidths(C - cLo + 1) = Len(cellText)
            End If
        Next R
    Next C
    
    '2) impression header / meta
    Debug.Print String(60, "-")
    Debug.Print "DebugArray2DTable "; IIf(label <> "", "[" & label & "] ", ""); _
                "(rows=" & rows & ", cols=" & cols & ")"
    
    '3) ligne séparatrice ( +-----+------+ )
    Debug.Print BuildBorderLine(colWidths)
    
    '4) lignes de données
    For R = rLo To rHi
        Debug.Print BuildRowLine(a, R, cLo, cHi, colWidths)
        'Après la 1ère ligne, on remet une bordure pour séparer header du reste.
        If R = rLo Then
            Debug.Print BuildBorderLine(colWidths)
        End If
    Next R
    
    'dernière bordure
    Debug.Print BuildBorderLine(colWidths)
    Debug.Print String(60, "-")
    Exit Sub

Not2D:
    Debug.Print "DebugArray2DTable "; label; ": argument n'est pas un Variant(2D)."
    
End Sub

' ========= HELPERS PRIVÉS =========
'Construit une ligne de données du style :
'| PassengerId | Survived | Pclass |
Private Function BuildRowLine(ByRef a As Variant, ByVal R As Long, ByVal cLo As Long, ByVal cHi As Long, ByRef colWidths() As Long) As String

    Dim s As String, txt As String, C As Long, idx As Long
    
    s = "|"
    idx = 1
    For C = cLo To cHi
        txt = SafeToString(a(R, C))
        s = s & " " & PadRight(txt, colWidths(idx)) & " |"
        idx = idx + 1
    Next C
    BuildRowLine = s
    
End Function

' Construit une bordure du style :
' +------------+----------+--------+
Private Function BuildBorderLine(ByRef colWidths() As Long) As String
    Dim s As String, i As Long
    s = "+"
    For i = LBound(colWidths) To UBound(colWidths)
        s = s & String(colWidths(i) + 2, "-") & "+" '+2 pour les espaces autour du texte
    Next i
    BuildBorderLine = s
End Function

'PadRight("abc", 6) -> "abc   "
Private Function PadRight(ByVal txt As String, ByVal width As Long) As String
    Dim n As Long
    n = width - Len(txt)
    If n < 0 Then n = 0
    PadRight = txt & String(n, " ")
End Function

'SafeToString gère Null / Empty / Erreurs Excel etc.
Private Function SafeToString(ByVal v As Variant) As String
    If IsError(v) Then
        SafeToString = "#ERR"
    ElseIf IsNull(v) Then
        SafeToString = "NULL"
    ElseIf IsEmpty(v) Then
        SafeToString = ""
    Else
        SafeToString = CStr(v)
    End If
End Function

'convertion un Array() "jagged" en Variant(2D)
Public Function ToVariant2D(ByVal jag As Variant) As Variant

    Dim R As Long, C As Long, rows As Long, cols As Long
    rows = UBound(jag) - LBound(jag) + 1
    cols = UBound(jag(LBound(jag))) - LBound(jag(LBound(jag))) + 1
    Dim v(): ReDim v(1 To rows, 1 To cols)
    For R = 1 To rows
        For C = 1 To cols
            v(R, C) = jag(LBound(jag) + R - 1)(LBound(jag(LBound(jag))) + C - 1)
        Next C
    Next R
    ToVariant2D = v
    
End Function

'Variant(1D, 0-based) de nKeys ISINs "FR" + n digits (1..maxVal, répartis)
Public Function MakeKeyList1D(prefix As String, digits As Long, nKeys As Long, maxVal As Long) As Variant
    Dim stepv As Double: stepv = maxVal / nKeys
    Dim v() As Variant, i As Long, k As Long
    ReDim v(0 To nKeys - 1)
    For i = 0 To nKeys - 1
        k = 1 + CLng(i * stepv)         ' équirépartition
        If k > maxVal Then k = maxVal
        v(i) = prefix & Right$(String$(digits, "0") & CStr(k), digits)
    Next
    MakeKeyList1D = v
End Function

'transforme un Variant(1D) -> "'v1','v2',...'vn'" (pour IN(...))
Public Function JoinQuoted(ids As Variant) As String
    Dim i As Long, buf As String, sep As String
    For i = LBound(ids) To UBound(ids)
        buf = buf & sep & SQ(CStr(ids(i)))
        sep = ","
    Next
    JoinQuoted = buf
End Function

Public Function SQ(ByVal s As String) As String
    SQ = "'" & Replace(s, "'", "''") & "'"
End Function

Public Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If sh Is Nothing Then
        Set sh = ThisWorkbook.Worksheets.Add
        sh.name = sheetName
    End If
    Set GetOrCreateSheet = sh
End Function

Public Sub DumpVariant2D(ByRef v As Variant, ByVal targetSheet As String, ByVal topLeft As String)
    Dim sh As Worksheet
    Set sh = GetOrCreateSheet(targetSheet)
    ArrayToSheet v, sh, topLeft
    sh.Activate
    sh.Range(topLeft).Select
End Sub

'=== Compte  lignes  Parquet via read_parquet + COUNT(*) ======
Public Function ParquetRowCount(ByRef db As cDuck, ByVal p As String) As Double
    Dim v As Variant, sql As String
    sql = "SELECT COUNT(*)::UBIGINT AS n FROM read_parquet(" & SqlQ(p) & ");"
    v = db.QueryFast(sql)
    If IsEmpty(v) Or UBound(v, 1) < 2 Or UBound(v, 2) < 1 Then
        Err.Raise 5, , "COUNT(*) n'a rien renvoyé."
    End If
    ParquetRowCount = CDbl(v(2, 1))
End Function

'=== ParquetInfo ===================
Public Function Duck_ParquetInfo(ByRef db As cDuck, ByVal parquetPath As String) As Variant

    Dim p As String, sql As String, v As Variant, nRows As Double
    
    p = Replace(parquetPath, "\", "/")

    'Charger l’extension parquet (best-effort)
    On Error Resume Next
        db.TryLoadExt "parquet"
    On Error GoTo 0
    '1) Compte fiable (lit réellement les lignes)
    nRows = ParquetRowCount(db, p)
    '2) Chemin rapide : parquet_schema + constante nRows
    On Error GoTo fallback
    sql = _
        "SELECT s.name AS column_name," & vbCrLf & _
        "       s.type AS data_type," & vbCrLf & _
        "       " & CStr(nRows) & "::UBIGINT AS row_count" & vbCrLf & _
        "FROM parquet_schema(" & SqlQ(p) & ") AS s;"
    v = db.QueryFast(sql)
    Duck_ParquetInfo = v
    Exit Function

fallback:
    '3) Fallback universel : DESCRIBE + constante nRows
    sql = _
        "WITH n AS (SELECT " & CStr(nRows) & "::UBIGINT AS n) " & vbCrLf & _
        "SELECT d.column_name AS column_name," & vbCrLf & _
        "       d.column_type AS data_type," & vbCrLf & _
        "       n.n           AS row_count" & vbCrLf & _
        "FROM (DESCRIBE SELECT * FROM read_parquet(" & SqlQ(p) & ")) AS d, n;"
    v = db.QueryFast(sql)
    Duck_ParquetInfo = v
    
End Function

'=========================== HELPERS CSV===========================

Public Sub DuckDbReadImportCsv(ByVal duckPath As String, ByVal csvPath As String, ByVal tableName As String, _
    Optional ByVal delim As String = "auto", Optional ByVal replaceAll As Boolean = True, Optional ByVal displayPreview As Boolean = True, Optional ByVal displaySheet As Variant = 1)

    On Error GoTo Fail

    Dim db As New cDuck, ws As Worksheet, preview As Variant, p As String, qtbl As String, createSQL As String, copySQL As String, m As String

    '--- init
    db.Init ThisWorkbook.Path
    db.ErrorMode = 2  '2=LogOnly (debug via duckdb_errors.log), 1=MsgBox, 0=Raise
    db.OpenDuckDb duckPath

    p = Replace(csvPath, "\", "/")
    qtbl = db.QuoteIdent(tableName)
    m = LCase$(Trim$(delim))

    '--- Build CREATE (schema) + COPY (bulk)
    Select Case m
        Case "auto"
            createSQL = "SELECT * FROM read_csv_auto(" & SqlQ(p) & ", header=true, sample_size=-1, ignore_errors=true) LIMIT 0"
            copySQL = "COPY " & qtbl & " FROM " & SqlQ(p) & " (FORMAT CSV, HEADER, AUTO_DETECT true, SAMPLE_SIZE -1, IGNORE_ERRORS true)"
        Case ","
            createSQL = "SELECT * FROM read_csv_auto(" & SqlQ(p) & ", header=true, delim=',', sample_size=-1, ignore_errors=true) LIMIT 0"
            copySQL = "COPY " & qtbl & " FROM " & SqlQ(p) & " (FORMAT CSV, HEADER, DELIMITER ',', SAMPLE_SIZE -1, IGNORE_ERRORS true)"
        Case ";"
            createSQL = "SELECT * FROM read_csv_auto(" & SqlQ(p) & ", header=true, delim=';', sample_size=-1, ignore_errors=true) LIMIT 0"
            copySQL = "COPY " & qtbl & " FROM " & SqlQ(p) & " (FORMAT CSV, HEADER, DELIMITER ';', SAMPLE_SIZE -1, IGNORE_ERRORS true)"
        Case "\t", "tab"
            createSQL = "SELECT * FROM read_csv_auto(" & SqlQ(p) & ", header=true, delim='\t', sample_size=-1, ignore_errors=true) LIMIT 0"
            copySQL = "COPY " & qtbl & " FROM " & SqlQ(p) & " (FORMAT CSV, HEADER, DELIMITER '\t', SAMPLE_SIZE -1, IGNORE_ERRORS true)"
        Case Else
            Err.Raise 5, , "Delim inconnu: " & delim & " (auto, ',', ';', '\t')"
    End Select

    '--- (re)create table (robuste)
    If replaceAll Then
        db.Exec "DROP TABLE IF EXISTS " & qtbl & ";"
    Else
        'si on ne remplace pas, on garantit quand même que la table existe
        db.Exec "CREATE TABLE IF NOT EXISTS " & qtbl & " AS " & createSQL & ";"
        db.Exec "DELETE FROM " & qtbl & ";"
        GoTo DoCopy
    End If

    db.Exec "CREATE TABLE " & qtbl & " AS " & createSQL & ";"

DoCopy:
    db.Exec copySQL & ";"

    '--- Preview Excel
    If displayPreview Then
        If IsNumeric(displaySheet) Then
            Set ws = ThisWorkbook.Worksheets(CLng(displaySheet))
        Else
            Set ws = ThisWorkbook.Worksheets(CStr(displaySheet))
        End If
        preview = db.QueryFast("SELECT * FROM " & qtbl & " LIMIT 200;")
        ArrayToSheet preview, ws, "A1"
    End If

    MsgBox "Import + affichage OK : " & tableName, vbInformation

CleanExit:
    db.CloseDuckDb
    Exit Sub

Fail:
    MsgBox "ERREUR Import CSV->DuckDB: " & Err.Description & vbCrLf & Native_LastErrorText(), vbCritical
    Resume CleanExit
End Sub

'=========================== HELPERS ACCESS to DuckDB ===========================

Public Function TryImportViaOdbcQuery(ByRef db As cDuck, ByVal conn As String, _
                                      ByVal accessTable As String, ByVal duckTable As String) As Boolean
    On Error GoTo KO

    Dim sql As String
    sql = "CREATE OR REPLACE TABLE " & db.QuoteIdent(duckTable) & " AS " & _
          "SELECT * FROM odbc_query(" & _
          "  password   => '', " & _
          "  connection => " & SqlQ(conn) & ", " & _
          "  query      => " & SqlQ("SELECT * FROM [" & accessTable & "]") & _
          ");"

    db.Exec sql
    TryImportViaOdbcQuery = True
    Exit Function

KO:
    TryImportViaOdbcQuery = False
End Function

Public Function TryImportViaOdbcScan(ByRef db As cDuck, ByVal conn As String, _
                                     ByVal accessTable As String, ByVal duckTable As String) As Boolean
    On Error GoTo KO
    Dim sql As String
    'avec odbc_scan, en général on passe le nom sans []
    sql = "CREATE OR REPLACE TABLE " & db.QuoteIdent(duckTable) & " AS " & _
          "SELECT * FROM odbc_scan(" & _
          "  connection => " & SqlQ(conn) & ", " & _
          "  table_name  => " & SqlQ(accessTable) & _
          ");"
    db.Exec sql
    TryImportViaOdbcScan = True
    Exit Function
KO:
    TryImportViaOdbcScan = False
End Function

Public Function TryImportViaADO_Recordset(ByRef db As cDuck, ByVal accdbPath As String, _
                                          ByVal accessTable As String, ByVal duckTable As String) As Boolean
    On Error GoTo KO
    Dim cn As Object, rs As Object

    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accdbPath & ";Persist Security Info=False;"

    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 2 'adUseServer (plus adapté gros volumes)
    rs.Open "SELECT * FROM [" & accessTable & "];", cn, 0, 1, 1 'forward-only, read-only

    db.Exec "DROP TABLE IF EXISTS " & db.QuoteIdent(duckTable) & ";"
    db.AppendAdoRecordsetFast rs, duckTable, True

    rs.Close: cn.Close
    Set rs = Nothing: Set cn = Nothing
    TryImportViaADO_Recordset = True
    Exit Function

KO:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> 0 Then rs.Close
    If Not cn Is Nothing Then If cn.State <> 0 Then cn.Close
    Set rs = Nothing: Set cn = Nothing
    TryImportViaADO_Recordset = False
End Function

Public Function TryImportViaADO_Variant(ByRef db As cDuck, ByVal accdbPath As String, _
                                        ByVal accessTable As String, ByVal duckTable As String) As Boolean
    On Error GoTo KO
    Dim cn As Object, rs As Object, v As Variant

    Set cn = CreateObject("ADODB.Connection")
    cn.CursorLocation = 3 'adUseClient (utile pour GetRows)
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accdbPath & ";Persist Security Info=False;"

    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3
    rs.Open "SELECT * FROM [" & accessTable & "];", cn, 1, 1 'keyset, readonly

    db.Exec "DROP TABLE IF EXISTS " & db.QuoteIdent(duckTable) & ";"

    If rs.EOF Then
        db.Exec "CREATE TABLE " & db.QuoteIdent(duckTable) & "(dummy INT);"
    Else
        db.CreateTableFromRecordsetSchema duckTable, rs
        v = db.RecordsetToVariant2D(rs, True)
        db.AppendArray duckTable, v, True
    End If

    rs.Close: cn.Close
    Set rs = Nothing: Set cn = Nothing
    TryImportViaADO_Variant = True
    Exit Function

KO:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> 0 Then rs.Close
    If Not cn Is Nothing Then If cn.State <> 0 Then cn.Close
    Set rs = Nothing: Set cn = Nothing
    TryImportViaADO_Variant = False
    
End Function

'=== Copie Access -> DuckDB via ODBC (scan OU query) ===========================
' mode = "scan"  : accessSource = nom de table Access (ex: "Clients")
' mode = "query" : accessSource = soit nom de table ("Clients")
'                  soit SQL Access complet ("SELECT ... FROM ... WHERE ...")
Public Function CopyAccessToDuck_ODBC(ByRef db As cDuck, ByVal accdbPath As String, ByVal accessSource As String, ByVal duckDbPath As String, ByVal duckTable As String, _
                                Optional ByVal mode As String = "scan")

    On Error GoTo Fail

    Dim conn As String, acc As String, sql As String, lastErr As String, srcSql As String, m As String, t As New cHiPerfTimer, msImport As Double

    m = LCase$(Trim$(mode))

    t.Start

    '1) Extension DuckDB odbc / nanodbc requise
    If Not db.EnsureOdbcLoaded Then
        Err.Raise 5, , "Extension DuckDB ODBC/NanoODBC introuvable (fichiers extension manquants ou mal placés)."
    End If
    '2) Connexion ODBC Access (chemin en /)
    acc = Replace(accdbPath, "\", "/")
    conn = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & acc & ";Uid=Admin;Pwd=;"

    '3) Build SQL selon mode
    Select Case m
        Case "scan"
            sql = _
                "CREATE TABLE " & db.QuoteIdent(duckTable) & " AS " & _
                "SELECT * FROM odbc_scan(" & _
                "  connection => " & SqlQ(conn) & ", " & _
                "  table_name  => " & SqlQ(accessSource) & _
                ");"

        Case "query"
            sql = "CREATE OR REPLACE TABLE " & db.QuoteIdent(duckTable) & " AS " & _
                  "SELECT * FROM odbc_query(" & _
                  "  password   => '', " & _
                  "  connection => " & SqlQ(conn) & ", " & _
                  "  query      => " & SqlQ("SELECT * FROM [" & accessSource & "]") & _
                  ");"
        
        Case Else
            Err.Raise 5, , "Mode invalide: '" & mode & "'. Utilise 'scan' ou 'query'."
    End Select

    '4) Exécute (transaction + drop)
    db.BeginTx
        db.Exec "DROP TABLE IF EXISTS " & db.QuoteIdent(duckTable) & ";"
        db.Exec sql
    db.Commit
    
    lastErr = Native_LastErrorText()
    If Len(lastErr) > 0 Then Err.Raise 5, , "DuckDB: " & lastErr
    
    
#If VERBOSE Then
    Dim v As Variant
    'v = db.QueryFast("SELECT COUNT(*) AS n FROM " & db.QuoteIdent(duckTable) & ";")
    Debug.Print "Copie ODBC OK (" & m & "), lignes="; v(2, 1)
#End If

    db.CloseDuckDb
    msImport = t.StopMilliseconds
    Debug.Print "Done ! " & " | Temps : "; Format$(msImport, "0.000"); " ms"
    
    Exit Function

Fail:
    On Error Resume Next
    If db.handle <> 0 Then db.Rollback
    db.CloseDuckDb
    MsgBox "Erreur ODBC (" & m & "): " & Err.Description & _
           IIf(Len(Native_LastErrorText) > 0, vbCrLf & Native_LastErrorText, ""), vbExclamation
                       
End Function

'Tools : Crée une base Access + table de test avec données factices
Public Function CreateAccesDbSample(ByVal accdbPath As String, ByVal tableName As String, Optional ByVal rowCount As Long = 50) As Boolean

    On Error GoTo Fail

    Dim fso As Object, cat As Object, cn As Object, cmd As Object, folder As String, qt As String
    Dim i   As Long, isin As String, nc As String, px As Double, dt As Date
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.GetParentFolderName(accdbPath)
    If Len(folder) > 0 Then If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    If fso.FileExists(accdbPath) Then fso.DeleteFile accdbPath, True

    '1) Créer la base (ADOX)
    Set cat = CreateObject("ADOX.Catalog")
    cat.Create "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accdbPath & ";"
    Set cat = Nothing

    '2) Ouvrir une connexion (ADODB)
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accdbPath & ";Persist Security Info=False;"

    '3) Créer la table
    qt = "[" & tableName & "]"
    cn.Execute _
        "CREATE TABLE " & qt & " (" & _
        "  ID AUTOINCREMENT PRIMARY KEY, " & _
        "  [ISIN] TEXT(12), " & _
        "  [NumeroContrat] TEXT(30), " & _
        "  [Price] DOUBLE, " & _
        "  [Modified At] DATETIME" & _
        ");"

    'Index utiles (facultatif)
    cn.Execute "CREATE INDEX ix_" & tableName & "_isin ON " & qt & " ([ISIN]);"
    cn.Execute "CREATE INDEX ix_" & tableName & "_numc ON " & qt & " ([NumeroContrat]);"
    
    '4) Insérer des lignes de démo (paramètres = protection locale)
    Const adVarChar As Long = 200       '<-- VarChar (évite VarWChar avec ACE)
    Const adDouble As Long = 5
    Const adDBTimeStamp As Long = 135   '<-- IMPORTANT pour DATETIME Access
    Const adParamInput As Long = 1
    Const adCmdText As Long = 1
    
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    
    With cmd
        .CommandType = adCmdText
        .CommandText = "INSERT INTO " & qt & " ([ISIN],[NumeroContrat],[Price],[Modified At]) VALUES (?,?,?,?)"
        'tailles explicites pour les champs texte
        .Parameters.Append cmd.CreateParameter("p1", adVarChar, adParamInput, 12)
        .Parameters.Append cmd.CreateParameter("p2", adVarChar, adParamInput, 30)
        .Parameters.Append cmd.CreateParameter("p3", adDouble, adParamInput)
        .Parameters.Append cmd.CreateParameter("p4", adDBTimeStamp, adParamInput)
        .Prepared = True   '<-- garde le plan/typage stable entre itérations
        Randomize
        For i = 1 To rowCount
            isin = "FR" & Right$("0000000000" & CStr(i), 10)
            nc = "C-" & Right$("000" & CStr(i Mod 1000), 3)
            px = 80# + Rnd * 50#
            dt = Now - TimeSerial((i Mod 72) \ 6, (i * 7) Mod 60, 0)
        
            .Parameters(0).Value = isin
            .Parameters(1).Value = nc
            .Parameters(2).Value = px
            .Parameters(3).Value = dt
            .Execute
        Next
    End With
    
    cn.Close
    Set cmd = Nothing
    Set cn = Nothing

    CreateAccesDbSample = True
    Exit Function

Fail:
    On Error Resume Next
    If Not cmd Is Nothing Then Set cmd = Nothing
    If Not cn Is Nothing Then If cn.State <> 0 Then cn.Close
    Set cn = Nothing
    CreateAccesDbSample = False
    MsgBox "Create Access DB KO: " & Err.Description, vbExclamation
End Function

Public Function BuildRandomCsvFile(ByVal outCsvPath As String, Optional ByVal nRows As Long = 500000) As Boolean

    On Error GoTo Fail

    Dim fso As Object, folder As String, isin As String, line As String, t0 As Double, t1 As Double
    Dim d   As Date, price As Long, vol As Long, i As Long, f As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    '1) dossier
    folder = fso.GetParentFolderName(outCsvPath)
    If Len(folder) > 0 Then
        If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    End If

    '2) supprime si existe
    If Dir$(outCsvPath, vbNormal) <> "" Then Kill outCsvPath

    '3) écrit CSV
    f = FreeFile
    Open outCsvPath For Output As #f

    Randomize
    Print #f, "ISIN,Price,Volume,TradeDate"  'header

    t0 = Timer
    For i = 1 To nRows
        isin = RandomIsin_12()

        'entiers pour éviter souci séparateur décimal
        price = CLng(Rnd * 100000)     '0..99999
        vol = CLng(Rnd * 1000000)      '0..999999

        'date aléatoire sur 365 jours
        d = DateAdd("d", -CLng(Rnd * 365), Date)

        'date entre guillemets => plus robuste pour parsers
        line = """" & isin & """," & CStr(price) & "," & CStr(vol) & ",""" & Format$(d, "yyyy-mm-dd") & """"
        Print #f, line
    Next i

    Close #f

    t1 = Timer
    Debug.Print "CSV généré : "; outCsvPath
    Debug.Print "Lignes     : "; nRows
    Debug.Print "Temps      : "; Format$(t1 - t0, "0.00"); " sec"

    BuildRandomCsvFile = True
    Exit Function

Fail:
    On Error Resume Next
    If f <> 0 Then Close #f
    Debug.Print "BuildRandomCsvFile - Erreur: "; Err.Number; Err.Description
    BuildRandomCsvFile = False
    
End Function

Public Sub MakeSampleCsv(Optional ByVal outPath As String = "")

    Dim f As Integer, i As Long, n As Long, isin$, num$, px As Double, dt As Date, p As String

    On Error GoTo Fail

    If outPath = "" Then outPath = ThisWorkbook.Path & "\data.csv"
    p = outPath
    f = FreeFile
    Open p For Output As #f
    'Entêtes
    Print #f, "ISIN,NumeroContrat,Prix,ModifiedAt"
    
    Randomize
    n = 200
    For i = 1 To n
        isin = "FR" & Right$("0000000000" & CStr(i), 10)                'FR0000000001...
        num = "C-" & Right$("000" & CStr((i Mod 999) + 1), 3)           'C-001...
        px = Round(50 + Rnd * 200, 2)                                   '50..250
        dt = Now - Rnd * 30 - (Rnd * 86400#) / 86400#                   '~ derniers 30 jours
        ' CSV avec virgules ; timestamp en ISO
        'Print #f, isin & "," & num & "," & CStr(px) & "," & Format$(dt, "yyyy-mm-dd hh:nn:ss")
        
        Print #f, """" & isin & """,""" & num & """," & Replace$(Format$(px, "0.00"), ",", ".") & _
          ",""" & Format$(dt, "yyyy-mm-dd hh:nn:ss") & """"
    Next i

    Close #f
    MsgBox "Échantillon CSV créé : " & p, vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    If f <> 0 Then Close #f
    MsgBox "Création CSV KO: " & Err.Description, vbExclamation
End Sub

'ISIN-like aléatoire (12 chars alphanum)
Private Function RandomIsin_12() As String
    Const CHARS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    Dim i As Long, n As Long, s As String

    n = Len(CHARS)
    s = ""
    For i = 1 To 12
        s = s & Mid$(CHARS, 1 + Int(Rnd * n), 1)
    Next i

    RandomIsin_12 = s
End Function

Public Sub Debug_OdbcScan_Signature()

    On Error GoTo Fail

    Dim db As New cDuck, v As Variant, found As Boolean

    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"

    If Not db.EnsureOdbcLoaded Then
        MsgBox "Impossible de charger odbc/nanodbc", vbCritical
        GoTo Bye
    End If

    '========================
    ' 1) Essai duckdb_functions()
    '========================
    found = False
    On Error Resume Next
    v = db.QueryFast( _
        "SELECT * " & _
        "FROM duckdb_functions() " & _
        "WHERE lower(function_name) = 'odbc_scan';")
    On Error GoTo 0

    If Not IsEmpty(v) Then
        'QueryFast renvoie souvent une ligne d'entête => il faut au moins 2 lignes pour avoir 1 résultat
        If UBound(v, 1) >= 2 Then
            DebugArray2DTable v, "duckdb_functions() : odbc_scan"
            found = True
        End If
    End If

    'Si pas trouvé, on élargit à "%odbc%" pour voir si l'extension a bien enregistré quelque chose
    If Not found Then
        On Error Resume Next
        v = db.QueryFast( _
            "SELECT * " & _
            "FROM duckdb_functions() " & _
            "WHERE lower(function_name) LIKE '%odbc%';")
        On Error GoTo 0

        If Not IsEmpty(v) And UBound(v, 1) >= 2 Then
            DebugArray2DTable v, "duckdb_functions() : *odbc* (odbc_scan non trouvé)"
            found = True
        End If
    End If

    '========================
    ' 2) Fallback pragma_functions()
    '========================
    If Not found Then
        On Error Resume Next
        v = db.QueryFast( _
            "SELECT * " & _
            "FROM pragma_functions() " & _
            "WHERE lower(name) = 'odbc_scan';")
        On Error GoTo 0

        If Not IsEmpty(v) And UBound(v, 1) >= 2 Then
            DebugArray2DTable v, "pragma_functions() : odbc_scan"
            found = True
        End If

        If Not found Then
            On Error Resume Next
            v = db.QueryFast( _
                "SELECT * " & _
                "FROM pragma_functions() " & _
                "WHERE lower(name) LIKE '%odbc%';")
            On Error GoTo 0

            If Not IsEmpty(v) And UBound(v, 1) >= 2 Then
                DebugArray2DTable v, "pragma_functions() : *odbc* (odbc_scan non trouvé)"
                found = True
            End If
        End If
    End If
    If Not found Then
        Debug.Print "Aucune entrée 'odbc_scan' trouvée dans duckdb_functions() ou pragma_functions()."
        Debug.Print "=> Soit l'extension n'a pas enregistré la fonction, soit la vue système diffère."
        MsgBox "Aucune signature trouvée pour odbc_scan." & vbCrLf & _
               "Teste aussi duckdb_functions()/pragma_functions() sans filtre pour inspecter.", vbExclamation
    End If

Bye:
    db.CloseDuckDb
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur Debug_OdbcScan_Signature:" & vbCrLf & Err.Description & _
           IIf(Len(Native_LastErrorText) > 0, vbCrLf & Native_LastErrorText, ""), vbExclamation
End Sub

Public Function AccessTable_ToParquet(ByRef db As cDuck, ByVal accdbPath As String, ByVal tableName As String, ByVal outParquet As String, Optional ByVal threads As Long = 0) As Boolean

    On Error GoTo Fail

    Dim accNorm As String, outNorm As String, conn As String, sql As String, odbcErr As String

    Call EnsureFolderExists(outParquet)


    If threads > 0 Then
        db.Exec "PRAGMA threads=" & CStr(Application.WorksheetFunction.Max(1, threads)) & ";"
    End If

    Call db.TryLoadExt("parquet")

    accNorm = Replace(accdbPath, "\", "/")
    outNorm = Replace(outParquet, "\", "/")

    '========================
    ' 1) ODBC (nanoODBC)
    '========================
    If db.EnsureOdbcLoaded Then

        conn = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};" & _
               "Dbq=" & accNorm & ";" & _
               "Uid=Admin;Pwd=;"

        sql = "COPY (" & _
              "  SELECT * FROM odbc_scan(" & _
              "    connection => " & SqlQ(conn) & ", " & _
              "    table_name  => " & SqlQ(tableName) & _
              "  )" & _
              ") TO " & SqlQ(outNorm) & " (FORMAT PARQUET, COMPRESSION ZSTD);"

        'Debug.Print "SQL SENT (ODBC) = " & sql

        On Error GoTo OdbcFail
        db.Exec sql
        On Error GoTo Fail

        AccessTable_ToParquet = True
        GoTo Clean
    End If

OdbcFail:
    'On mémorise l'erreur ODBC mais on ne stoppe pas : fallback ADO
    odbcErr = Native_LastErrorText()
    Debug.Print "ODBC failed -> fallback ADO. DuckDB says: " & odbcErr
    Err.Clear
    On Error GoTo Fail

    '========================
    ' 2) Fallback ADO
    '========================
    Dim cn As Object, rs As Object

    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accdbPath & ";Persist Security Info=False;"

    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 2
    rs.Open "SELECT * FROM [" & tableName & "];", cn, 0, 1, 1

    db.Exec "DROP TABLE IF EXISTS __tmp__;"
    If rs.EOF Then
        db.Exec "CREATE TABLE __tmp__(dummy INT);"
    Else
        db.CreateTableFromRecordsetSchema "__tmp__", rs
        db.AppendAdoRecordset rs, "__tmp__", False
    End If

    rs.Close: cn.Close
    Set rs = Nothing: Set cn = Nothing

    sql = "COPY __tmp__ TO " & SqlQ(outNorm) & " (FORMAT PARQUET, COMPRESSION ZSTD);"
    Debug.Print "SQL SENT (ADO->COPY) = " & sql
    db.Exec sql

    AccessTable_ToParquet = True

Clean:
    On Error Resume Next
    db.CloseDuckDb
    Exit Function

Fail:
    On Error Resume Next
    db.CloseDuckDb
    AccessTable_ToParquet = False

    Dim msg As String
    msg = Err.Description
    Dim lastErr As String
    lastErr = Native_LastErrorText()
    If Len(lastErr) > 0 Then msg = msg & vbCrLf & "DuckDB: " & lastErr

    MsgBox "Access->Parquet KO:" & vbCrLf & msg, vbExclamation
End Function

'=========================== HELPERS PARQUET RESEARCH ===========================

'===========================================================
' Lire 1 ligne dans un Parquet via une clé (ISIN, etc.)
'   - Variant 2D avec headers + 1 ligne (si trouvé)
'   - Empty (si aucune ligne)
'===========================================================
Public Function ParquetRowByKey(ByRef db As cDuck, parquetPath As String, ByVal keyCol As String, ByVal keyValue As Variant, _
                        Optional ByVal openDuckDbPath As String = ":memory:") As Variant

    On Error GoTo Fail

    Dim p As String, sql As String, keyExpr As String, arr As Variant

    'Normaliser le chemin pour DuckDB
    p = Replace(parquetPath, "\", "/")

    'Valeur clé : string vs numeric
    If IsNumeric(keyValue) Then
        keyExpr = CStr(keyValue)
    ElseIf IsDate(keyValue) Then
        'Si tu veux gérer dates, adapte selon ton format (TIMESTAMP/DATE)
        keyExpr = SqlQ(CStr(keyValue))
    Else
        keyExpr = SqlQ(CStr(keyValue))
    End If
    'Requête : read_parquet + filtre sur colonne clé
    sql = "SELECT * " & _
          "FROM read_parquet(" & SqlQ(p) & ") " & _
          "WHERE " & db.QuoteIdent(keyCol) & " = " & keyExpr & " " & _
          "LIMIT 1;"

    arr = db.QueryFast(sql)
    db.CloseDuckDb
    'Si QueryFast renvoie juste les headers (1 ligne) ou vide -> pas trouvé
    If IsEmpty(arr) Then
        ParquetRowByKey = Empty
    ElseIf UBound(arr, 1) < 2 Then
        ParquetRowByKey = Empty
    Else
        ParquetRowByKey = arr
    End If

    Exit Function

Fail:
    On Error Resume Next
    db.CloseDuckDb
    ParquetRowByKey = Empty
    
End Function

'==============================================================================
' Lit un Parquet et retourne toutes les lignes dont keyCol est dans dictKeys.
' - dictKeys: Scripting.Dictionary contenant les ISIN en clés (values ignorées)
' - Retour: Variant 2D avec headers (ligne 1) + lignes trouvées
' Perf:
'   - 1 seule ouverture DuckDB
'   - 1 seule lecture read_parquet
'   - filtre par JOIN sur une table TEMP des clés
'==============================================================================
Public Function ParquetRowsByKeyDict(ByRef db As cDuck, ByVal parquetPath As String, ByVal keyCol As String, ByRef dictKeys As Object, _
            Optional ByVal keepInputOrder As Boolean = True) As Variant

    On Error GoTo Fail

    Dim keysArr As Variant, k As Variant, p As String, sql As String, i As Long, n As Long
    
    If dictKeys Is Nothing Then
        ParquetRowsByKeyDict = Empty
        Exit Function
    End If
    If dictKeys.Count = 0 Then
        ParquetRowsByKeyDict = Empty
        Exit Function
    End If
    '--- Build array des clés pour FrameFromValue (headers + n lignes)
    n = dictKeys.Count
    If keepInputOrder Then
        ReDim keysArr(1 To n + 1, 1 To 2)
        keysArr(1, 1) = keyCol
        keysArr(1, 2) = "__ord"
        i = 2
        For Each k In dictKeys.keys
            keysArr(i, 1) = CStr(k)
            keysArr(i, 2) = i - 2
            i = i + 1
        Next
    Else
        ReDim keysArr(1 To n + 1, 1 To 1)
        keysArr(1, 1) = keyCol
        i = 2
        For Each k In dictKeys.keys
            keysArr(i, 1) = CStr(k)
            i = i + 1
        Next
    End If
    '--- table TEMP des clés (très rapide)
    db.FrameFromValue "__keys", keysArr, True, True  'hasHeader=True, makeTemp=True
    '--- Normaliser chemin parquet
    p = Replace(parquetPath, "\", "/")

    '--- 1 requête: read_parquet + JOIN sur les clés
    If keepInputOrder Then
        sql = _
            "SELECT p.* " & _
            "FROM read_parquet(" & SqlQ(p) & ") p " & _
            "INNER JOIN __keys k " & _
            "ON p." & db.QuoteIdent(keyCol) & " = k." & db.QuoteIdent(keyCol) & " " & _
            "ORDER BY k.__ord;"
    Else
        sql = _
            "SELECT p.* " & _
            "FROM read_parquet(" & SqlQ(p) & ") p " & _
            "INNER JOIN __keys k " & _
            "ON p." & db.QuoteIdent(keyCol) & " = k." & db.QuoteIdent(keyCol) & ";"
    End If

    ParquetRowsByKeyDict = db.QueryFast(sql)

    db.CloseDuckDb
    Exit Function

Fail:
    On Error Resume Next
    db.CloseDuckDb
    ParquetRowsByKeyDict = Empty
    
End Function

'==============================================================================
' ParquetRowByKey_SelectCols
' - Lit une seule ligne depuis un Parquet via keyCol=keyValue
' - Retourne soit:
'     * un scalaire (si 1 colonne demandée)
'     * un Variant(2D) (headers + 1 ligne) si plusieurs colonnes
' - Ne ferme PAS la DB (l'appelant gère Open/Close)
'==============================================================================
Public Function ParquetRowByKey_SelectCols(ByRef db As cDuck, ByVal parquetPath As String, ByVal keyCol As String, _
                    ByVal keyValue As Variant, ParamArray cols() As Variant) As Variant

    On Error GoTo Fail

    Dim arr As Variant, i As Long, p As String, sql As String, keyExpr As String, selectList As String

    p = Replace(parquetPath, "\", "/")
    If IsNumeric(keyValue) Then
        keyExpr = CStr(keyValue)
    Else
        keyExpr = SqlQ(CStr(keyValue))
    End If

    If (Not Not cols) = 0 Or (UBound(cols) < LBound(cols)) Then
        selectList = "*"
    Else
        selectList = ""
        For i = LBound(cols) To UBound(cols)
            If Len(CStr(cols(i))) > 0 Then
                If Len(selectList) > 0 Then selectList = selectList & ", "
                selectList = selectList & "p." & db.QuoteIdent(CStr(cols(i)))
            End If
        Next i
        If Len(selectList) = 0 Then selectList = "*"
    End If

    sql = "SELECT " & selectList & " " & _
          "FROM read_parquet(" & SqlQ(p) & ") p " & _
          "WHERE p." & db.QuoteIdent(keyCol) & " = " & keyExpr & " " & _
          "LIMIT 1;"

    arr = db.QueryFast(sql)
    If IsEmpty(arr) Then
        ParquetRowByKey_SelectCols = Empty
        Exit Function
    End If
    If UBound(arr, 1) < 2 Then
        ParquetRowByKey_SelectCols = Empty
        Exit Function
    End If

    If selectList <> "*" Then
        Dim colCount As Long
        colCount = UBound(arr, 2)
        If colCount = 1 Then
            ParquetRowByKey_SelectCols = arr(2, 1)
            Exit Function
        End If
    End If

    ParquetRowByKey_SelectCols = arr
    Exit Function

Fail:
    ParquetRowByKey_SelectCols = Empty
End Function

'==============================================================================
' ParquetReadFiltersToArray
' Objectif:
'   Lire un fichier Parquet et retourner un Variant(2D) (headers + lignes)
'   en appliquant N filtres SQL (pushdown) + un ORDER BY optionnel.
'
' Pourquoi c'est rapide:
'   - DuckDB lit le Parquet en colonne (columnar scan)
'   - Les filtres WHERE sont "poussés" dans le scan (moins d'I/O)
'   - Le résultat arrive directement en Variant(2D) via QueryFast (pas de boucle VBA)
'
' Paramètres:
'   db          : connexion DuckDB déjà ouverte (ex: :memory:)
'   parquetPath : chemin Parquet (Windows ou déjà normalisé)
'   orderBySql  : "" ou "Price DESC" ou "ORDER BY Price DESC"
'   filters()   : ParamArray de conditions SQL combinées avec AND
'                 Ex: "Price > 100", "ISIN LIKE 'FR%'", "Name IS NOT NULL"
'
' Retour:
'   Variant(2D) avec headers (ligne 1) + lignes
'   Empty si erreur
'==============================================================================
Public Function ParquetReadFiltersToArray(ByRef db As cDuck, ByVal parquetPath As String, _
            ByVal orderBySql As String, ParamArray filters() As Variant) As Variant

    On Error GoTo Fail

    Dim p As String, sql As String, whereSql As String, i As Long, cond As String

    p = Replace(parquetPath, "\", "/")
    whereSql = ""
    If (Not Not filters) <> 0 Then
        For i = LBound(filters) To UBound(filters)
            cond = Trim$(CStr(filters(i)))
            If Len(cond) > 0 Then
                If Len(whereSql) > 0 Then whereSql = whereSql & " AND "
                whereSql = whereSql & "(" & cond & ")"
            End If
        Next i
    End If
    sql = "SELECT * FROM read_parquet(" & SqlQ(p) & ") p "
    If Len(whereSql) > 0 Then sql = sql & "WHERE " & whereSql & " "

    If Len(orderBySql) > 0 Then
        If UCase$(Left$(Trim$(orderBySql), 8)) = "ORDER BY" Then
            sql = sql & orderBySql & " "
        Else
            sql = sql & "ORDER BY " & orderBySql & " "
        End If
    End If

    sql = sql & ";"

    ParquetReadFiltersToArray = db.QueryFast(sql)
    Exit Function

Fail:
    ParquetReadFiltersToArray = Empty
End Function


