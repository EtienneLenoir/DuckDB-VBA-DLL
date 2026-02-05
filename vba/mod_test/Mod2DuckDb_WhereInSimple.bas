Attribute VB_Name = "Mod2DuckDb_WhereInSimple"
Option Explicit


'Filtres de type "WHERE x IN (liste)" avec creation table temp sans méthode SelectWithTempList
Public Sub Smoke_TempList()

    On Error GoTo Fail
    'session DuckDB en mémoire
    Dim db As New cDuck, ids As Variant, C As Variant, a As Variant
    
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"

    'Données de test
    db.Exec "CREATE TABLE T(isin TEXT, px DOUBLE, name TEXT);" & _
            "INSERT INTO T VALUES " & _
            "('FR0001', 101.2, 'TOTAL')," & _
            "('FR0002',  99.9, 'AIRBUS')," & _
            "('FR0003', 103.5, 'DANONE');"

    'Liste temporaire depuis un Variant(1D)
    ids = Array("FR0001", "FR0003")      ' 1D, 0-based — OK
    db.CreateTempList "tmp_ids", ids, "VARCHAR"

    '(optionnel) contrôle du nombre de clés
    C = db.QueryFast("SELECT COUNT(*) AS n FROM tmp_ids;")
    Debug.Print "tmp_ids rows = "; C(2, 1)

    '4) Sélection via la temp table
    a = db.QueryFast("SELECT isin, name, px " & _
                     "FROM T " & _
                     "WHERE isin IN (SELECT v FROM tmp_ids) " & _
                     "ORDER BY isin;")
    Debug.Print "rows sel = "; UBound(a, 1) - 1

    '5) Dump dans la feuille
    With ThisWorkbook.Worksheets(1)
        .Cells.Clear
        .Range("A1").Resize(UBound(a, 1), UBound(a, 2)).Value = a
    End With

    db.CloseDuckDb
    MsgBox "OK", vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    If Not db Is Nothing Then db.CloseDuckDb
    MsgBox Err.Description, vbExclamation
End Sub

'Benchmark WHERE x IN (liste) VS table temp
Public Sub Benchmark_IN_vs_Temp()

    On Error GoTo Fail
    
    Dim db  As New cDuck, oldCalc As XlCalculation, ps As LongPtr, oldScrUpd As Boolean, oldEvt As Boolean
    Dim ids As Variant, px As Double, t0 As Double, t1 As Double, msJOIN As Double, msIN As Double, a As Variant
    Dim nm  As String, i As Long, nJOIN As Long, nIN As Long, isin As String, sqlIN As String, sqlJoin As String, msg As String

    '--------- Paramètres ----------
    Const N_ROWS As Long = 100000     'lignes dans T
    Const N_KEYS As Long = 10000      'nb de clés pour IN/temp
    Const SHOW_SAMPLE As Boolean = False

    '--------- Setup Excel (perf) ----------
    oldCalc = Application.Calculation: Application.Calculation = xlCalculationManual
    oldScrUpd = Application.ScreenUpdating: Application.ScreenUpdating = False
    oldEvt = Application.EnableEvents: Application.EnableEvents = False

    '--------- Session DuckDB ----------
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"

    ' --------- Jeu de données ----------
    db.Exec "DROP TABLE IF EXISTS T;"
    db.Exec "CREATE TABLE T(isin TEXT, px DOUBLE, name TEXT);"

    'insertion rapide via prepared statement
    ps = db.Prepare("INSERT INTO T VALUES(?, ?, ?)")

    For i = 1 To N_ROWS
        isin = "FR" & Right$("0000000000" & CStr(i), 10)  ' FR0000000001, ...
        px = 50# + (i Mod 1000) / 10#
        nm = "NAME_" & CStr(i Mod 1000)

        db.PS_BindText ps, 1, isin
        db.PS_BindDouble ps, 2, px
        db.PS_BindText ps, 3, nm
        db.PS_Exec ps
    Next
    db.PS_CloseDuckDb ps

    ' --------- Construire la liste de clés de test ----------
    ids = MakeKeyList1D("FR", 10, N_KEYS, N_ROWS)  ' 1D (0-based)

    ' --------- Mesure 1 : WHERE IN (...) ----------
    sqlIN = "SELECT count(*) AS n FROM T WHERE isin IN (" & JoinQuoted(ids) & ")"

    t0 = Timer
    a = db.QueryFast(sqlIN)
    t1 = Timer
    nIN = CLng(a(2, 1))
    msIN = Round((t1 - t0) * 1000#, 1)

    ' --------- Mesure 2 : temp table + JOIN ----------
    db.CreateTempList "tmp_ids", ids, "VARCHAR"
    
    sqlJoin = "SELECT count(*) AS n FROM T JOIN tmp_ids ON T.isin = tmp_ids.v"

    t0 = Timer
    a = db.QueryFast(sqlJoin)
    t1 = Timer
    nJOIN = CLng(a(2, 1))
    msJOIN = Round((t1 - t0) * 1000#, 1)

    ' --------- Vérif + échantillon ----------
    If nIN <> nJOIN Then Err.Raise 5, , "Comptes différents: IN=" & nIN & " vs JOIN=" & nJOIN

    If SHOW_SAMPLE Then
        a = db.QueryFast( _
            "SELECT T.isin, T.name, T.px " & _
            "FROM T JOIN tmp_ids ON T.isin = tmp_ids.v " & _
            "ORDER BY T.isin LIMIT 10")
        With ThisWorkbook.Worksheets(1)
            .Cells.Clear
            .Range("A1").Resize(UBound(a, 1), UBound(a, 2)).Value = a
        End With
    End If

    ' --------- Résultat ----------
    msg = "N_ROWS = " & N_ROWS & vbCrLf & _
          "N_KEYS = " & N_KEYS & vbCrLf & vbCrLf & _
          "WHERE IN(...) : " & msIN & " ms  (n=" & nIN & ")" & vbCrLf & _
          "Temp + JOIN   : " & msJOIN & " ms (n=" & nJOIN & ")" & vbCrLf & vbCrLf & _
          IIf(msJOIN < msIN, "? Temp+JOIN plus rapide", "? IN(...) plus rapide sur ce run")
    MsgBox msg, vbInformation, "Benchmark DuckDB (IN vs Temp+JOIN)"

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScrUpd
    Application.EnableEvents = oldEvt
    Exit Sub
Fail:
    On Error Resume Next
    db.CloseDuckDb
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScrUpd
    Application.EnableEvents = oldEvt
    MsgBox "Erreur: " & Err.Description, vbExclamation
    
End Sub


