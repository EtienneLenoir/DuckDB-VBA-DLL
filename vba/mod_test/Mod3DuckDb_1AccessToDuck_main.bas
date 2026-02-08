Attribute VB_Name = "Mod3DuckDb_1AccessToDuck_main"
Option Explicit

'===============================================================================
' Module : MsAccess_To_DuckDB_Import
'
' Objectif
'   Importer des données Microsoft Access (.accdb/.mdb) vers DuckDB depuis Excel/VBA.
'   Le module propose 4 chemins d’import, du plus tout se fait dans DuckDB;

'   Important : on parle de 2 méthodes différentes
'     A) ODBC (DuckDB lit Access via son extension odbc/nanodbc)
'       - (1) odbc_scan : copie table brute via extension DuckDB ODBC (simple/robuste).
'       - (2) odbc_query: requête Access complète via extension DuckDB ODBC (flexible, souvent rapide si filtre).
'     B) ADO  (VBA lit Access via ACE OLEDB, puis on pousse vers DuckDB via la DLL bridge)
'       - (3) AppendAdoRecordset : ADO -> ingestion directe via DLL bridge (pas besoin d’extension ODBC duckDB).
'       - (4) Variant2D + AppendArray  via DLL bridge (plus RAM/plus lent, pas besoin d’extension ODBC duckDB).
'
'------------------------------------------------------------------------------
' (1) ODBC / odbc_query : exécution d’un SQL “côté Access” (le plus rapide mais nécessite extension duckdb)
'   Principe :
'     DuckDB envoie une requête SQL via  extension duckdb(nanodbc) ; c’est Access/ACE qui interprète le SQL :
'       SELECT * FROM odbc_query(conn, 'SELECT ... FROM [MaTable] ...');
'
'   ? Avantages/ désavantages :
'     - Très flexible : requête Access complète possible :
'         * SELECT ... WHERE ...JOIN / GROUP BY / ORDER BY
'         * requête sur une Query Access (pas seulement une table)
'     - Souvent plus rapide si la requête filtre beaucoup : moins de données traversent ODBC.
'     - Extension DuckDB "odbc" ou "nanodbc" disponible.
'     - Driver ODBC Microsoft Access installé.
'
'------------------------------------------------------------------------------
' (2) ODBC / odbc_scan  : lecture table brute (scan)
    'https://duckdb.org/community_extensions/extensions/nanodbc
'    ou mettre extension ici %USERPROFILE%\.duckdb\extensions\v1.4.3\windows_amd64
'                (= C:\Users\<username>\.duckdb\extensions\v1.4.3\windows_amd64)
'
'   Principe :
'     DuckDB accède à Access via extension duckdb(nanodbc) et lit une TABLE en mode scan :
'       SELECT * FROM odbc_scan(connection => conn, table_name => 'MaTable');
'
'     Ici tu ne donnes PAS un SQL complet : tu donnes un nom de table (ou parfois une vue),
'     et DuckDB fait ensuite son SELECT côté DuckDB.
'
'   ? Avantages/ désavantages :
'     - parfait pour copier une table telle quelle.
'     - Moins flexible : pas de JOIN/WHERE complexes côté Access.
'     - pour filtrer/limiter, il faut le faire après côté DuckDB, ex :
'         CREATE TABLE t AS SELECT * FROM odbc_scan(...) WHERE ...;
'     - Pré-requis : Extension DuckDB "nanodbc" disponible
'     - Driver ODBC Microsoft Access (mdb/accdb) installé sur la machine.
'------------------------------------------------------------------------------
' (3) ADO Recordset -> DuckDB via DLL : AppendAdoRecordset (DLL bridge)
'   Principe :
'     VBA ouvre Access via ADODB (Provider ACE OLEDB), récupère un Recordset,
'     puis la DLL "duckdb_vba_bridge.dll" ingère directement ce Recordset dans DuckDB :
'       rs = ADODB.Recordset
'       db.AppendAdoRecordset rs, "MaTableDuck", createIfMissing:=True
'     Ici, la "connexion Access" n’est pas DuckDB->ODBC, mais VBA->ACE OLEDB,
'     et l’ingestion se fait via la DLL bridge (marshalling optimisé côté DLL).
'
'     - Ne dépend pas des extensions DuckDB odbc/nanodbc.

'------------------------------------------------------------------------------
' (4) ADO Recordset -> Variant(2D) -> AppendArray  (DLL bridge)
'   Principe
'     VBA lit Access via ADODB, convertit le Recordset en Variant(2D) (GetRows),
'     puis envoie l’array à DuckDB via AppendArray (SAFEARRAY) :
'       v = RecordsetToVariant2D(rs, withHeader:=True)
'       db.AppendArray "MaTableDuck", v, hasHeader:=True
'
'   ? Avantages/ désavantages :
'     - Fallback universel : utile si AppendAdoRecordset n’est pas dispo / problématique.
'     - Facile à déboguer (tu vois l’array).
'     - Plus de RAM ( on matérialises toutes les données en mémoire VBA).
'     - Souvent plus lent sur gros volumes (copie + conversions + SAFEARRAY).
'     - Recommandé surtout pour volumes modestes ou dernier recours.
'     - ACE OLEDB installé.
'
'===============================================================================

'Crée une base Access de test avec données factices
Private Sub CreateDataBaseAccessSample()

    Dim ok As Boolean
    ok = CreateAccesDbSample(ThisWorkbook.Path & "\DbAccess.accdb", "Clients", 300)
    'Call BuildRandomCsvFile(ThisWorkbook.Path & "\RandomData_TestAdo.csv", 500000)
    If ok Then MsgBox "Base et table créées avec succès.", vbInformation
    
End Sub

Public Sub TestMain_AccessToDuckDB()

    On Error GoTo Fail

    '==================== PARAMS (à adapter) ====================
    Dim accdbPath   As String, accessTable As String, duckDbPath As String, duckTable As String, t As New cHiPerfTimer, msImport As Double
    Dim db          As New cDuck, a As Variant, conn As String, accNorm As String, usedPath As String, strMethod As String, method As Long, ok As Boolean

    accdbPath = ThisWorkbook.Path & "\DbAccess.accdb"
    accessTable = "Clients" '"Clients"           'ex: "Clients" ou "TestAdo"
    duckDbPath = ThisWorkbook.Path & "\cache.duckdb"
    duckTable = "Access_Table"        'table créée/écrasée dans DuckDB

    '--- Choix de méthode :
    '    1 = ODBC odbc_query
    '    2 =  ODBC odbc_scan
    '    3 = ADO Recordset -> AppendAdoRecordsetFast (DLL)
    '    4 = ADO Variant2D -> AppendArray
    '    0 = Auto (tente 2 -> 1 -> 3 -> 4
    method = 2
    '============================================================

    db.Init ThisWorkbook.Path
    db.OpenDuckDb duckDbPath
    t.Start

    '--- prépare la connexion ODBC Access (chemin en /)
    accNorm = Replace(accdbPath, "\", "/")
    conn = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & accNorm & ";Uid=Admin;Pwd=;"

    '============================================================
    ' Exécution selon "method"
    '============================================================
    Select Case method

        Case 1
            strMethod = "ODBC odbc_query"
            If Not db.EnsureOdbcLoaded Then Err.Raise 5, , "ODBC: extension DuckDB odbc/nanodbc non dispo."
            ok = TryImportViaOdbcQuery(db, conn, accessTable, duckTable)
            usedPath = "ODBC / odbc_query"
            
        Case 2
            strMethod = "ODBC odbc_scan"
            If Not db.EnsureOdbcLoaded Then Err.Raise 5, , "ODBC: extension DuckDB odbc/nanodbc non dispo."
            ok = TryImportViaOdbcScan(db, conn, accessTable, duckTable)
            usedPath = "ODBC / odbc_scan"

        Case 3
            strMethod = "ADO AppendAdoRecordset"
            ok = TryImportViaADO_Recordset(db, accdbPath, accessTable, duckTable)
            usedPath = "ADO / AppendAdoRecordset"

        Case 4
            strMethod = "ADO Variant2D+AppendArray"
            ok = TryImportViaADO_Variant(db, accdbPath, accessTable, duckTable)
            usedPath = "ADO / Variant2D + AppendArray"

        Case Else
            'AUTO : 2 -> 1 -> 3 -> 4 (ordre logique)
            ok = False

            If db.EnsureOdbcLoaded Then
                If TryImportViaOdbcQuery(db, conn, accessTable, duckTable) Then
                    ok = True: usedPath = "ODBC / odbc_query"
                ElseIf TryImportViaOdbcScan(db, conn, accessTable, duckTable) Then
                    ok = True: usedPath = "ODBC / odbc_scan"
                End If
            End If

            If Not ok Then
                If TryImportViaADO_Recordset(db, accdbPath, accessTable, duckTable) Then
                    ok = True: usedPath = "ADO / AppendAdoRecordset"
                ElseIf TryImportViaADO_Variant(db, accdbPath, accessTable, duckTable) Then
                    ok = True: usedPath = "ADO / Variant2D + AppendArray"
                End If
            End If

    End Select

    If Not ok Then Err.Raise 5, , "Import KO : aucune méthode n'a abouti."
    msImport = t.StopMilliseconds
    
Show:
    a = db.QueryFast("SELECT * FROM " & db.QuoteIdent(duckTable) & " LIMIT 300;")
    ArrayToSheet a, ThisWorkbook.Worksheets(1), "A1"

    db.CloseDuckDb

    Debug.Print "Méthode "; strMethod; " | Temps : "; Format$(msImport, "0.000"); " ms"
                
    'MsgBox "Import Access -> DuckDB OK." & vbCrLf & _
           "Access=" & accessTable & vbCrLf & _
           "DuckDB=" & duckTable & vbCrLf & _
           "Méthode=" & usedPath & vbCrLf & _
           "Temps import=" & Format$(msImport, "0.000") & " ms", vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Erreur import: " & Err.Description & _
           IIf(Len(Native_LastErrorText) > 0, vbCrLf & Native_LastErrorText, ""), vbExclamation
End Sub

