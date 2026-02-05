Attribute VB_Name = "Mod2DuckDb_Miniplot_Ext"
Option Explicit

Public Sub Demo_Extension_Miniplot_BarChart()

    'https://duckdb.org/community_extensions/extensions/miniplot
    'https://github.com/nkwork9999/miniplot
    
    Dim db As New cDuck, csvPath As String, duckPath As String, outHtml As String, sql As String, a As Variant, html As String

    On Error GoTo Fail

    csvPath = ThisWorkbook.Path & "\test.csv" ' < ' file csv exemple
    duckPath = ThisWorkbook.Path & "\DbDuckDb.duckdb" ' ou ":memory:"
    outHtml = ThisWorkbook.Path & "\miniplot_product_revenue.html"

    db.Init ThisWorkbook.Path
    db.ErrorMode = 2  '2=LogOnly, 1=MsgBox, 0=Raise
    db.OpenDuckDb duckPath

    db.Exec "LOAD miniplot;"

    ' --- A) Tentative: générer directement un fichier HTML (si supporté par miniplot)
    On Error Resume Next
    sql = "SELECT bar_chart(" & _
          "list(product)," & _
          "list(revenue)," & _
          "'Product Revenue'," & _
          "'" & Replace(outHtml, "'", "''") & "'" & _
          ") FROM read_csv_auto('" & Replace(csvPath, "'", "''") & "');"
    db.Exec sql
    If Err.Number = 0 Then
        On Error GoTo Fail
        ThisWorkbook.FollowHyperlink outHtml
        'MsgBox "OK : miniplot a généré " & vbCrLf & outHtml, vbInformation
        GoTo CleanExit
    End If

    ' --- B) Fallback: le chart est renvoyé en texte (HTML) -> on écrit le fichier nous-même
    Err.Clear
    On Error GoTo Fail

    sql = "SELECT bar_chart(" & _
          "list(product)," & _
          "list(revenue)," & _
          "'Product Revenue'" & _
          ") FROM read_csv_auto('" & Replace(csvPath, "'", "''") & "');"

    a = db.QueryFast(sql)

    ' a est un tableau 2D (headers + 1 ligne)
    ' La valeur HTML (ou un texte équivalent) est généralement en (2,1)
    If IsArray(a) Then
        If UBound(a, 1) >= 2 And UBound(a, 2) >= 1 Then
            html = CStr(a(2, 1))
        End If
    End If

    If Len(html) = 0 Then Err.Raise vbObjectError + 300, , "miniplot n'a rien renvoyé (HTML vide)."

    Call WriteTextUtf8(outHtml, html)
    ThisWorkbook.FollowHyperlink outHtml

    'MsgBox "OK : miniplot a généré (fallback) " & vbCrLf & outHtml, vbInformation

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    On Error Resume Next
    db.Rollback
    db.CloseDuckDb
    On Error GoTo 0
    MsgBox "Erreur: " & db.LastError, vbExclamation
    Resume CleanExit

End Sub

' =========================================================
' 1) BAR CHART (DuckDB)
' =========================================================
Public Sub Demo_Miniplot_BarChart_()

    Dim db As New cDuck, duckPath As String, outHtml As String, sql As String, a As Variant, html As String

    On Error GoTo Fail

    duckPath = ThisWorkbook.Path & "\DbDuckDb.duckdb"  'ou ":memory:"
    outHtml = ThisWorkbook.Path & "\miniplot_bar_from_duck.html"

    db.Init ThisWorkbook.Path
    db.ErrorMode = 2
    db.OpenDuckDb duckPath

    db.Exec "LOAD miniplot;"

    ' Données dans DuckDB
    db.Exec "CREATE OR REPLACE TEMP TABLE t_bar(product VARCHAR, revenue DOUBLE);"
    db.Exec "INSERT INTO t_bar VALUES " & _
            "('iPhone',450),('MacBook',380),('iPad',290),('AirPods',185),('Watch',160);"

    ' --- A) Tentative: générer directement un fichier HTML
    ' IMPORTANT: SqlQ(outHtml) est déjà quoted -> ne pas rajouter de quotes autour
    On Error Resume Next
    sql = "SELECT bar_chart(" & _
          "list(product ORDER BY revenue DESC)," & _
          "list(revenue ORDER BY revenue DESC)," & _
          "'Product Revenue (DuckDB)'," & _
          SqlQ(outHtml) & _
          ") FROM t_bar;"
    db.Exec sql

    If Err.Number = 0 Then
        On Error GoTo Fail
        ThisWorkbook.FollowHyperlink outHtml
        GoTo CleanExit
    End If

    ' --- B) Fallback: renvoi HTML -> on écrit le fichier
    Err.Clear
    On Error GoTo Fail

    sql = "SELECT bar_chart(" & _
          "list(product ORDER BY revenue DESC)," & _
          "list(revenue ORDER BY revenue DESC)," & _
          "'Product Revenue (DuckDB)'" & _
          ") FROM t_bar;"

    a = db.QueryFast(sql)

    If IsArray(a) Then
        If UBound(a, 1) >= 2 And UBound(a, 2) >= 1 Then html = CStr(a(2, 1))
    End If

    If Len(html) = 0 Then Err.Raise vbObjectError + 300, , "miniplot n'a rien renvoyé (HTML vide)."

    WriteTextUtf8 outHtml, html
    ThisWorkbook.FollowHyperlink outHtml

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    On Error Resume Next
    db.Rollback
    db.CloseDuckDb
    On Error GoTo 0
    MsgBox "Erreur: " & db.LastError, vbExclamation
    Resume CleanExit

End Sub

' =========================================================
' 2) Line Chart(DuckDB)
' =========================================================

Public Sub Demo_Miniplot_LineChart()

    Dim db As New cDuck, duckPath As String, outHtml As String
    Dim sql As String, a As Variant, html As String

    On Error GoTo Fail

    duckPath = ThisWorkbook.Path & "\DbDuckDb.duckdb"  'ou ":memory:"
    outHtml = ThisWorkbook.Path & "\miniplot_line_from_duck.html"

    db.Init ThisWorkbook.Path
    db.ErrorMode = 2
    db.OpenDuckDb duckPath

    db.Exec "LOAD miniplot;"

    ' Données dans DuckDB (série temporelle)
    db.Exec "CREATE OR REPLACE TEMP TABLE t_line(x DATE, y DOUBLE);"

    ' IMPORTANT: range() -> BIGINT, on cast en INTEGER pour DATE + INTEGER
    db.Exec "INSERT INTO t_line " & _
            "SELECT (DATE '2024-01-01' + CAST(i AS INTEGER)) AS x, " & _
            "       50 + 10*sin(CAST(i AS DOUBLE)/2.0) AS y " & _
            "FROM range(0, 30) tbl(i);"

    ' --- A) Tentative: générer directement un fichier HTML
    On Error Resume Next
    sql = "SELECT line_chart(" & _
          "list(CAST(x AS VARCHAR) ORDER BY x)," & _
          "list(y ORDER BY x)," & _
          "'Line Chart (DuckDB)'," & _
          SqlQ(outHtml) & _
          ") FROM t_line;"
    db.Exec sql

    If Err.Number = 0 Then
        On Error GoTo Fail
        ThisWorkbook.FollowHyperlink outHtml
        GoTo CleanExit
    End If

    ' --- B) Fallback: renvoi HTML -> on écrit le fichier
    Err.Clear
    On Error GoTo Fail

    sql = "SELECT line_chart(" & _
          "list(CAST(x AS VARCHAR) ORDER BY x)," & _
          "list(y ORDER BY x)," & _
          "'Line Chart (DuckDB)'" & _
          ") FROM t_line;"

    a = db.QueryFast(sql)

    If IsArray(a) Then
        If UBound(a, 1) >= 2 And UBound(a, 2) >= 1 Then html = CStr(a(2, 1))
    End If

    If Len(html) = 0 Then Err.Raise vbObjectError + 310, , "miniplot n'a rien renvoyé (HTML vide)."

    WriteTextUtf8 outHtml, html
    ThisWorkbook.FollowHyperlink outHtml

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    On Error Resume Next
    db.Rollback
    db.CloseDuckDb
    On Error GoTo 0
    MsgBox "Erreur: " & db.LastError, vbExclamation
    Resume CleanExit

End Sub

' =========================================================
' 3)  Scatter Chart (DuckDB)
' =========================================================

Public Sub Demo_Miniplot_ScatterChart()

    Dim db As New cDuck, duckPath As String, outHtml As String, sql As String, a As Variant, html As String

    On Error GoTo Fail

    duckPath = ThisWorkbook.Path & "\DbDuckDb.duckdb"  'ou ":memory:"
    outHtml = ThisWorkbook.Path & "\miniplot_scatter_from_duck.html"

    db.Init ThisWorkbook.Path
    db.ErrorMode = 2
    db.OpenDuckDb duckPath

    db.Exec "LOAD miniplot;"

    ' Données dans DuckDB (nuage de points avec un peu de bruit)
    db.Exec "CREATE OR REPLACE TEMP TABLE t_scatter(x DOUBLE, y DOUBLE);"
    db.Exec "INSERT INTO t_scatter " & _
            "SELECT CAST(i AS DOUBLE) AS x, " & _
            "       (CAST(i AS DOUBLE) * 1.8) + (random() - 0.5) * 8.0 AS y " & _
            "FROM range(1, 101) tbl(i);"

    ' --- A) Tentative: générer directement un fichier HTML
    On Error Resume Next
    sql = "SELECT scatter_chart(" & _
          "list(x ORDER BY x)," & _
          "list(y ORDER BY x)," & _
          "'Scatter Chart (DuckDB)'," & _
          SqlQ(outHtml) & _
          ") FROM t_scatter;"
    db.Exec sql

    If Err.Number = 0 Then
        On Error GoTo Fail
        ThisWorkbook.FollowHyperlink outHtml
        GoTo CleanExit
    End If

    ' --- B) Fallback: renvoi HTML -> on écrit le fichier
    Err.Clear
    On Error GoTo Fail

    sql = "SELECT scatter_chart(" & _
          "list(x ORDER BY x)," & _
          "list(y ORDER BY x)," & _
          "'Scatter Chart (DuckDB)'" & _
          ") FROM t_scatter;"

    a = db.QueryFast(sql)

    If IsArray(a) Then
        If UBound(a, 1) >= 2 And UBound(a, 2) >= 1 Then html = CStr(a(2, 1))
    End If

    If Len(html) = 0 Then Err.Raise vbObjectError + 320, , "miniplot n'a rien renvoyé (HTML vide)."

    WriteTextUtf8 outHtml, html
    ThisWorkbook.FollowHyperlink outHtml

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    On Error Resume Next
    db.Rollback
    db.CloseDuckDb
    On Error GoTo 0
    MsgBox "Erreur: " & db.LastError, vbExclamation
    Resume CleanExit

End Sub

Public Sub Demo_Miniplot_ScatterChart_ListValueStyle()

    Dim db As New cDuck, duckPath As String, outHtml As String, sql As String, a As Variant, html As String

    On Error GoTo Fail

    duckPath = ThisWorkbook.Path & "\DbDuckDb.duckdb"  'ou ":memory:"
    outHtml = ThisWorkbook.Path & "\miniplot_scatter_arraystyle.html"

    db.Init ThisWorkbook.Path
    db.ErrorMode = 2
    db.OpenDuckDb duckPath

    db.Exec "LOAD miniplot;"

    ' --- A) Tentative: fichier direct
    On Error Resume Next
    sql = "SELECT scatter_chart(" & _
          "list_value(1.0, 2.0, 3.0, 4.0, 5.0)," & _
          "list_value(2.5, 5.0, 7.5, 10.0, 12.5)," & _
          "'Correlation Analysis'," & _
          SqlQ(outHtml) & _
          ");"
    db.Exec sql

    If Err.Number = 0 Then
        On Error GoTo Fail
        ThisWorkbook.FollowHyperlink outHtml
        GoTo CleanExit
    End If

    ' --- B) Fallback HTML
    Err.Clear
    On Error GoTo Fail

    sql = "SELECT scatter_chart(" & _
          "list_value(1.0, 2.0, 3.0, 4.0, 5.0)," & _
          "list_value(2.5, 5.0, 7.5, 10.0, 12.5)," & _
          "'Correlation Analysis'" & _
          ");"

    a = db.QueryFast(sql)

    If IsArray(a) Then
        If UBound(a, 1) >= 2 And UBound(a, 2) >= 1 Then html = CStr(a(2, 1))
    End If

    If Len(html) = 0 Then Err.Raise vbObjectError + 341, , "miniplot n'a rien renvoyé (HTML vide)."

    WriteTextUtf8 outHtml, html
    ThisWorkbook.FollowHyperlink outHtml

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    On Error Resume Next
    db.Rollback
    db.CloseDuckDb
    On Error GoTo 0
    MsgBox "Erreur: " & db.LastError, vbExclamation
    Resume CleanExit

End Sub

' =========================================================
' 3)  AreaChart (DuckDB)
' =========================================================

Public Sub Demo_Miniplot_AreaChart()

    Dim db As New cDuck, duckPath As String, outHtml As String, sql As String, a As Variant, html As String

    On Error GoTo Fail

    duckPath = ThisWorkbook.Path & "\DbDuckDb.duckdb"  'ou ":memory:"
    outHtml = ThisWorkbook.Path & "\miniplot_area_from_duck.html"

    db.Init ThisWorkbook.Path
    db.ErrorMode = 2
    db.OpenDuckDb duckPath

    db.Exec "LOAD miniplot;"

    ' --- A) Tentative: générer directement un fichier HTML
    On Error Resume Next
    sql = "WITH s AS (" & _
          "  SELECT " & _
          "    (DATE '2024-01-01' + CAST(i AS INTEGER)) AS x," & _
          "    100 + 25*sin(CAST(i AS DOUBLE)/3.0) + CAST(i AS DOUBLE)*1.2 AS y " & _
          "  FROM range(0, 40) tbl(i)" & _
          ") " & _
          "SELECT area_chart(" & _
          "  list(CAST(x AS VARCHAR) ORDER BY x)," & _
          "  list(y ORDER BY x)," & _
          "  'Area Chart (DuckDB)'," & _
          "  " & SqlQ(outHtml) & _
          ") FROM s;"
    db.Exec sql

    If Err.Number = 0 Then
        On Error GoTo Fail
        ThisWorkbook.FollowHyperlink outHtml
        GoTo CleanExit
    End If

    ' --- B) Fallback: renvoi HTML -> on écrit le fichier
    Err.Clear
    On Error GoTo Fail

    sql = "WITH s AS (" & _
          "  SELECT " & _
          "    (DATE '2024-01-01' + CAST(i AS INTEGER)) AS x," & _
          "    100 + 25*sin(CAST(i AS DOUBLE)/3.0) + CAST(i AS DOUBLE)*1.2 AS y " & _
          "  FROM range(0, 40) tbl(i)" & _
          ") " & _
          "SELECT area_chart(" & _
          "  list(CAST(x AS VARCHAR) ORDER BY x)," & _
          "  list(y ORDER BY x)," & _
          "  'Area Chart (DuckDB)'" & _
          ") FROM s;"

    a = db.QueryFast(sql)

    If IsArray(a) Then
        If UBound(a, 1) >= 2 And UBound(a, 2) >= 1 Then html = CStr(a(2, 1))
    End If

    If Len(html) = 0 Then Err.Raise vbObjectError + 350, , "miniplot n'a rien renvoyé (HTML vide)."

    WriteTextUtf8 outHtml, html
    ThisWorkbook.FollowHyperlink outHtml

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    On Error Resume Next
    db.Rollback
    db.CloseDuckDb
    On Error GoTo 0
    MsgBox "Erreur: " & db.LastError, vbExclamation
    Resume CleanExit

End Sub

' =========================================================
' 5) Basic 3D Scatter (DuckDB)
' =========================================================

Public Sub Demo_Miniplot_Scatter3D_Basic()

    Dim db As New cDuck, duckPath As String, outHtml As String, sql As String, a As Variant, html As String

    On Error GoTo Fail

    duckPath = ThisWorkbook.Path & "\DbDuckDb.duckdb"  'ou ":memory:"
    outHtml = ThisWorkbook.Path & "\miniplot_scatter3d_from_duck.html"

    db.Init ThisWorkbook.Path
    db.ErrorMode = 2
    db.OpenDuckDb duckPath

    db.Exec "LOAD miniplot;"

    ' --- A) Tentative: générer directement un fichier HTML
    On Error Resume Next
    sql = "WITH s AS (" & _
          "  SELECT " & _
          "    CAST(i AS DOUBLE) AS x," & _
          "    (random() * 10.0) AS y," & _
          "    (CAST(i AS DOUBLE)/4.0 + random()*2.0) AS z " & _
          "  FROM range(1, 81) tbl(i)" & _
          ") " & _
          "SELECT scatter_3d_chart(" & _
          "  list(x ORDER BY x)," & _
          "  list(y ORDER BY x)," & _
          "  list(z ORDER BY x)," & _
          "  'Basic 3D Scatter (DuckDB)'," & _
          "  " & SqlQ(outHtml) & _
          ") FROM s;"
    db.Exec sql

    If Err.Number = 0 Then
        On Error GoTo Fail
        ThisWorkbook.FollowHyperlink outHtml
        GoTo CleanExit
    End If

    ' --- B) Fallback: renvoi HTML -> on écrit le fichier
    Err.Clear
    On Error GoTo Fail

    sql = "WITH s AS (" & _
          "  SELECT " & _
          "    CAST(i AS DOUBLE) AS x," & _
          "    (random() * 10.0) AS y," & _
          "    (CAST(i AS DOUBLE)/4.0 + random()*2.0) AS z " & _
          "  FROM range(1, 81) tbl(i)" & _
          ") " & _
          "SELECT scatter_3d_chart(" & _
          "  list(x ORDER BY x)," & _
          "  list(y ORDER BY x)," & _
          "  list(z ORDER BY x)," & _
          "  'Basic 3D Scatter (DuckDB)'" & _
          ") FROM s;"

    a = db.QueryFast(sql)

    If IsArray(a) Then
        If UBound(a, 1) >= 2 And UBound(a, 2) >= 1 Then html = CStr(a(2, 1))
    End If

    If Len(html) = 0 Then Err.Raise vbObjectError + 360, , "miniplot n'a rien renvoyé (HTML vide)."

    WriteTextUtf8 outHtml, html
    ThisWorkbook.FollowHyperlink outHtml

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    On Error Resume Next
    db.Rollback
    db.CloseDuckDb
    On Error GoTo 0
    MsgBox "Erreur: " & db.LastError, vbExclamation
    Resume CleanExit

End Sub

Public Sub Demo_Miniplot_Scatter3D_WithTimestamps()

    Dim db As New cDuck, duckPath As String, outHtml As String, sql As String, a As Variant, html As String

    On Error GoTo Fail

    duckPath = ThisWorkbook.Path & "\DbDuckDb.duckdb"  'ou ":memory:"
    outHtml = ThisWorkbook.Path & "\miniplot_scatter3d_timestamps_from_duck.html"

    db.Init ThisWorkbook.Path
    db.ErrorMode = 2
    db.OpenDuckDb duckPath

    db.Exec "LOAD miniplot;"

    ' --- A) Tentative: générer directement un fichier HTML
    On Error Resume Next
    sql = "WITH s AS (" & _
          "  SELECT " & _
          "    CAST(i AS DOUBLE) AS x," & _
          "    (random() * 10.0) AS y," & _
          "    (CAST(i AS DOUBLE)/4.0 + random()*2.0) AS z," & _
          "    (TIMESTAMP '2024-01-01 10:00:00' + CAST(i AS INTEGER) * INTERVAL 10 MINUTE) AS ts " & _
          "  FROM range(1, 61) tbl(i)" & _
          ") " & _
          "SELECT scatter_3d_chart(" & _
          "  list(x ORDER BY ts)," & _
          "  list(y ORDER BY ts)," & _
          "  list(z ORDER BY ts)," & _
          "  list(strftime(ts, '%Y-%m-%d %H:%M:%S') ORDER BY ts)," & _
          "  '3D Scatter with Timestamps (DuckDB)'," & _
          "  " & SqlQ(outHtml) & _
          ") FROM s;"
    db.Exec sql

    If Err.Number = 0 Then
        On Error GoTo Fail
        ThisWorkbook.FollowHyperlink outHtml
        GoTo CleanExit
    End If

    ' --- B) Fallback: renvoi HTML -> on écrit le fichier
    Err.Clear
    On Error GoTo Fail

    sql = "WITH s AS (" & _
          "  SELECT " & _
          "    CAST(i AS DOUBLE) AS x," & _
          "    (random() * 10.0) AS y," & _
          "    (CAST(i AS DOUBLE)/4.0 + random()*2.0) AS z," & _
          "    (TIMESTAMP '2024-01-01 10:00:00' + CAST(i AS INTEGER) * INTERVAL 10 MINUTE) AS ts " & _
          "  FROM range(1, 61) tbl(i)" & _
          ") " & _
          "SELECT scatter_3d_chart(" & _
          "  list(x ORDER BY ts)," & _
          "  list(y ORDER BY ts)," & _
          "  list(z ORDER BY ts)," & _
          "  list(strftime(ts, '%Y-%m-%d %H:%M:%S') ORDER BY ts)," & _
          "  '3D Scatter with Timestamps (DuckDB)'" & _
          ") FROM s;"

    a = db.QueryFast(sql)

    If IsArray(a) Then
        If UBound(a, 1) >= 2 And UBound(a, 2) >= 1 Then html = CStr(a(2, 1))
    End If

    If Len(html) = 0 Then Err.Raise vbObjectError + 370, , "miniplot n'a rien renvoyé (HTML vide)."

    WriteTextUtf8 outHtml, html
    ThisWorkbook.FollowHyperlink outHtml

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    On Error Resume Next
    db.Rollback
    db.CloseDuckDb
    On Error GoTo 0
    MsgBox "Erreur: " & db.LastError, vbExclamation
    Resume CleanExit

End Sub


