Attribute VB_Name = "Mod3DuckDb_2AccessToDuck_dll"
Option Explicit

'===============================================================================
' Module : ModTest3DuckDb_DbAcces
'
' Objectif
'   Démonstrations "end-to-end" Access <-> DuckDB en VBA (late binding),
'   avec focus sur l’ingestion rapide et la portabilité (pas de références requises).
'
' Contenu / Scénarios
'   1) Import_Recordset_Direct
'        - Ouvre Access via ADODB en late binding
'        - Récupère un Recordset (server-side, forward-only)
'        - Ingestion directe dans DuckDB via db.AppendAdoRecordset
'        - Affiche un extrait dans Excel + timer perf
'
'   2) CreateAccesDbSample (+ Make_Access_Sample)
'        - Crée un .accdb et une table de test (ADOX + ADODB)
'        - Insère des données de démo via ADODB.Command paramétré (prepared)
'
'   3) AccessTable_ToParquet_Fast (+ Export_Clients)
'        - Export Access -> Parquet
'        - Chemin rapide : DuckDB + odbc_scan/odbc_query (si extension dispo)
'        - Fallback : ADO -> table temp -> COPY PARQUET
'
'   4) ImportResultatAccessDansDuck_Test / CopyAccessTableToDuck_ODBC
'        - Import d’une requête Access complexe vers DuckDB (transaction)
'        - Variante ODBC (CREATE TABLE AS SELECT odbc_scan)
'
' Différences avec les autres modules
'   - Vs DbAcces2 : ici tout est late binding + ingestion directe Recordset (moins "VBA refs")
'   - Vs DbAccesExtension : ici on couvre ADO + ODBC + export Parquet, pas juste le test d’extension
'
' Pré-requis
'   - cDuck + mDuckNative (EnsureDuckDll, Native_LastErrorText, SqlQ, ArrayToSheet…)
'   - Access Database Engine installé si Provider ACE utilisé
'===============================================================================

'Crée une base Access de test avec données factices
Private Sub CreateDataBaseAccessSample()

    Dim ok As Boolean
    ok = CreateAccesDbSample(ThisWorkbook.Path & "\DbAccess.accdb", "Clients", 300)
    If ok Then MsgBox "Base et table créées avec succès.", vbInformation
    
End Sub

Public Sub TestAccessToDuckDb_AppendAdoRecordset()

    On Error GoTo Fail

    Dim db As New cDuck, v As Variant, cn As Object, rs As Object
    
    Dim t As New cHiPerfTimer
    t.Start
    
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ThisWorkbook.Path & "\cache.duckdb"

    '2) Recordset ADO (late binding)
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\DbAccess.accdb;"
    Set rs = CreateObject("ADODB.Recordset")
    'rs.CacheSize = 1000 '5000                     'Augmente cache, réduit aller retour
    rs.CursorLocation = 2                          'adUseServer (0 = Default, 2 = Server (+rapide) 3=coté Client)
    rs.Open "SELECT * FROM [TestAdo]", cn, 0, 1, 1 'adOpenForwardOnly = 0, adLockReadOnly = 1, adCmdText = 1

    '3) Push direct du Recordset -> DuckDB
    '(createIfMissing:=True => crée la table si elle n'existe pas)
    db.AppendAdoRecordset rs, "Access_Table_RS", True

    '4)Affiche un extrait
    v = db.QueryFast("select * FROM Access_Table_RS limit 200;")

    With ThisWorkbook.Worksheets(1)
        .Cells.Clear
        .Range("A1").Resize(UBound(v, 1), UBound(v, 2)).Value = v
        .Columns.AutoFit
    End With
    
    Debug.Print "Durée : "; Format$(t.StopMilliseconds, "0.000"); " ms"

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not cn Is Nothing Then cn.Close
    Set rs = Nothing: Set cn = Nothing
    db.CloseDuckDb
    Exit Sub
Fail:
    MsgBox "Erreur: " & Err.Description, vbExclamation
    Resume CleanExit
    
End Sub

'Version Fast
    'No malloc/free per VT_BSTR cell: reusable UTF-8 buffer + WideCharToMultiByte
    'Explicit transaction BEGIN/COMMIT (ROLLBACK on error)
    'Avoid VariantChangeType in hot path (CY direct; DECIMAL via VarR8FromDec when possible
    'treat VT_ERROR/VT_DISPATCH/VT_UNKNOWN as NULL fast-path; fallback still uses ChangeType)
    
Public Sub TestAccessToDuckDb_AppendAdoRecordsetFast()

    On Error GoTo Fail

    Dim db As New cDuck, v As Variant, cn As Object, rs As Object
    Dim t As New cHiPerfTimer
    t.Start

    db.Init ThisWorkbook.Path
    db.OpenDuckDb ThisWorkbook.Path & "\cache.duckdb"

    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\DbAccess.accdb;"
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 2
    rs.Open "SELECT * FROM [TestAdo]", cn, 0, 1, 1

    ' Fast version
    db.AppendAdoRecordsetFast rs, "Access_Table_RS_FAST", True

    v = db.QueryFast("select * FROM Access_Table_RS_FAST limit 200;")

    With ThisWorkbook.Worksheets(1)
        .Cells.Clear
        .Range("A1").Resize(UBound(v, 1), UBound(v, 2)).Value = v
        .Columns.AutoFit
    End With

    Debug.Print "FAST Durée : "; Format$(t.StopMilliseconds, "0.000"); " ms"

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not cn Is Nothing Then cn.Close
    Set rs = Nothing: Set cn = Nothing
    db.CloseDuckDb
    Exit Sub

Fail:
    MsgBox "Erreur: " & Err.Description, vbExclamation
    Resume CleanExit

End Sub



Public Sub Bench_AppendAdoRecordset_NormalVsFast()

    On Error GoTo Fail

    Dim db  As New cDuck, t As New cHiPerfTimer, cn As Object, rsN As Object, rsF As Object
    Dim msN As Double, msF As Double, sumN As Double, sumF As Double, i As Long, n As Long

    n = 20 'mets 50 si tu veux

    db.Init ThisWorkbook.Path
    db.OpenDuckDb ThisWorkbook.Path & "\cache.duckdb"

    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\DbAccess.accdb;"

    'Création 1 fois (hors bench)
    Set rsN = CreateObject("ADODB.Recordset")
    rsN.CursorLocation = 2
    rsN.Open "SELECT * FROM [TestAdo]", cn, 0, 1, 1
    db.AppendAdoRecordset rsN, "T_NORM", True
    rsN.Close

    Set rsF = CreateObject("ADODB.Recordset")
    rsF.CursorLocation = 2
    rsF.Open "SELECT * FROM [TestAdo]", cn, 0, 1, 1
    db.AppendAdoRecordsetFast rsF, "T_FAST", True
    rsF.Close

    ' Warm-up (non mesuré)
    db.Exec "DELETE FROM T_NORM;"
    Set rsN = CreateObject("ADODB.Recordset")
    rsN.CursorLocation = 2
    rsN.Open "SELECT * FROM [TestAdo]", cn, 0, 1, 1
    db.AppendAdoRecordset rsN, "T_NORM", False
    rsN.Close

    db.Exec "DELETE FROM T_FAST;"
    Set rsF = CreateObject("ADODB.Recordset")
    rsF.CursorLocation = 2
    rsF.Open "SELECT * FROM [TestAdo]", cn, 0, 1, 1
    db.AppendAdoRecordsetFast rsF, "T_FAST", False
    rsF.Close

    'Benchmark
    For i = 1 To n
        If (i Mod 2) = 1 Then 'Normal puis Fast
            db.Exec "DELETE FROM T_NORM;"
            Set rsN = CreateObject("ADODB.Recordset")
            rsN.CursorLocation = 2
            rsN.Open "SELECT * FROM [TestAdo]", cn, 0, 1, 1
            t.Start
            db.AppendAdoRecordset rsN, "T_NORM", False
            msN = t.StopMilliseconds
            rsN.Close

            db.Exec "DELETE FROM T_FAST;"
            Set rsF = CreateObject("ADODB.Recordset")
            rsF.CursorLocation = 2
            rsF.Open "SELECT * FROM [TestAdo]", cn, 0, 1, 1
            t.Start
            db.AppendAdoRecordsetFast rsF, "T_FAST", False
            msF = t.StopMilliseconds
            rsF.Close
        Else 'Fast puis Normal (pour compenser caches)
            db.Exec "DELETE FROM T_FAST;"
            Set rsF = CreateObject("ADODB.Recordset")
            rsF.CursorLocation = 2
            rsF.Open "SELECT * FROM [TestAdo]", cn, 0, 1, 1
            t.Start
            db.AppendAdoRecordsetFast rsF, "T_FAST", False
            msF = t.StopMilliseconds
            rsF.Close

            db.Exec "DELETE FROM T_NORM;"
            Set rsN = CreateObject("ADODB.Recordset")
            rsN.CursorLocation = 2
            rsN.Open "SELECT * FROM [TestAdo]", cn, 0, 1, 1
            t.Start
            db.AppendAdoRecordset rsN, "T_NORM", False
            msN = t.StopMilliseconds
            rsN.Close
        End If

        sumN = sumN + msN
        sumF = sumF + msF
        Debug.Print "Iter"; i; "Normal="; Format$(msN, "0.000"); " Fast="; Format$(msF, "0.000")
    Next i

    Debug.Print "AVG Normal ms:"; Format$(sumN / n, "0.000")
    Debug.Print "AVG Fast   ms:"; Format$(sumF / n, "0.000")
    Debug.Print "Gain x:"; Format$((sumN / n) / (sumF / n), "0.000")

CleanExit:
    On Error Resume Next
    If Not cn Is Nothing Then cn.Close
    Set cn = Nothing
    db.CloseDuckDb
    Exit Sub

Fail:
    MsgBox "Erreur: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Public Sub TestAccessToDuckDb_AppendArray()

    On Error GoTo Fail

    Dim db As New cDuck, cn As ADODB.Connection, rs As ADODB.Recordset, v As Variant, a As Variant
    
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ThisWorkbook.Path & "\cache.duckdb"

    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\DbAccess.accdb;"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT TOP 10 * FROM [Clients]", cn, adOpenKeyset, adLockReadOnly

    '1) Crée la table à partir du schéma ADO
    db.CreateTableFromRecordsetSchema "Access_Table", rs

    '2) Convertit en Variant(2D) et append
    v = db.RecordsetToVariant2D(rs, True)
    db.AppendArray "Access_Table", v, True

    '3) Affiche
    a = db.QueryFast("SELECT * FROM Access_Table")
    ArrayToSheet a, ThisWorkbook.Worksheets(1), "A1"

    rs.Close: cn.Close
    Set rs = Nothing: Set cn = Nothing
    db.CloseDuckDb

    MsgBox "ADO early binding OK (schéma + Variant)", vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> 0 Then rs.Close
    If Not cn Is Nothing Then If cn.State <> 0 Then cn.Close
    db.CloseDuckDb
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
    
End Sub





