Attribute VB_Name = "Mod3DuckDb_3Nanodbc_Ext"
Option Explicit

'===============================================================================
'               Extension nanodbc - DuckDB
'
'        https://duckdb.org/community_extensions/extensions/nanodbc
'       ou mettre extension ici  %USERPROFILE%\.duckdb\extensions\v1.4.3\windows_amd64
'
'   Copier des données Microsoft Access (.accdb/.mdb) vers DuckDB depuis Excel/VBA,
'   en s’appuyant sur l’extension **DuckDB nanoODBC** (extension "nanodbc").

'   Même si la voie nanoODBC (DuckDB -> ODBC) est souvent très rapide, le code C (DLL) n’est pas un plan B :
'   l’ingestion via AppendAdoRecordset est très optimisée côté DLL(C), et peut atteindre des performances
'   du même ordre (voire meilleures) selon la machine.
'
'   Autre avantage majeur :ne dépend pas de l’extension DuckDB nanoODBC (ni de ses fichiers .duckdb_extension à déployer).
'   S’appuie uniquement sur le provider ACE/OLEDB côté Windows et sur la DLL bridge déjà embarquée.
'   Résultat : intégration plus simple, moins de prérequis à installer/coller au bon endroit ((poste utilisateur, VDI, parc verrouillé)

'   DuckDB joue le rôle de client Access via ODBC :
'     - DuckDB charge l’extension nanodbc
'     - DuckDB se connecte au driver ODBC Microsoft Access
'     - DuckDB matérialise le résultat dans une table DuckDB
'
' Deux modes d’import (ODBC côté DuckDB) :
'
'   1) mode="query"  (recommandé pour requete SQL plus sophistiqué  / filtrage / jointures)
'      - DuckDB envoie une requête SQL au driver ODBC Access.
'      - Le SQL est interprété par Access/ACE (pas par DuckDB).
'      - Exemple :
'          CREATE OR REPLACE TABLE T AS
'          SELECT * FROM odbc_query(
'             password   => '',
'             connection => '<conn>',
'             query      => 'SELECT ... FROM [MaTable] WHERE ...');
'
'      Points forts :
'        * Très flexible (JOIN / WHERE / GROUP BY / ORDER BY côté Access)
'        * Souvent plus rapide si tu filtres beaucoup (moins de données transitent via ODBC)
'
'   2) mode="scan"   (recommandé pour copier une table brute telle quelle)
'      - DuckDB scanne une table Access par son nom.
'      - Exemple :
'          CREATE TABLE T AS
'          SELECT * FROM odbc_scan(
'             connection => '<conn>',
'             table_name => 'MaTable');
'      Points forts :
'        * Simple, robuste, idéal pour “copie brute”
'        * Peu de SQL Access à gérer
'
' Pré-requis
'   - Extension DuckDB **nanodbc** disponible et chargeable :
'       db.EnsureOdbcLoaded = True
'     (les fichiers d’extension doivent être présents au bon endroit / bonne version)
'   - Driver ODBC Microsoft Access installé :
'       "Microsoft Access Driver (*.mdb, *.accdb)"
'
' Différences avec les voies ADO/DLL (autres modules)
'   - Ici : tout passe par DuckDB + extension nanoODBC + driver ODBC.
'   - Ailleurs : VBA ouvre Access via ACE OLEDB (ADODB), puis la DLL bridge ingère
'     (AppendAdoRecordset ou AppendArray). Pas besoin d’extension nanodbc dans ce cas.
'===============================================================================

'Crée une base Access de test avec données factices
Private Sub CreateDataBaseAccessSample()

    Dim ok As Boolean
    ok = CreateAccesDbSample(ThisWorkbook.Path & "\DbAccess.accdb", "Clients", 300)
    If ok Then MsgBox "Base et table créées avec succès.", vbInformation
    
End Sub

Public Sub TestAccessToDuckDb_ExtScan()

    On Error GoTo Fail

    Dim db As New cDuck, v As Variant, accdbPath As String, duckPath As String, accessTable As String, duckTable As String

    accdbPath = ThisWorkbook.Path & "\DbAccess.accdb"
    duckPath = ThisWorkbook.Path & "\cache.duckdb"
    
    db.Init ThisWorkbook.Path
    db.OpenDuckDb duckPath

    accessTable = "Clients"
    duckTable = "Clients_ODBC"

    '1) Copie Access -> DuckDB (ODBC)
             '-> selection "scan" ou "query"
    Call CopyAccessToDuck_ODBC(db, accdbPath, accessTable, duckPath, duckTable, "scan")

    '2) Vérif rapide : relire dans DuckDB et afficher
    db.Init ThisWorkbook.Path
    db.OpenDuckDb duckPath
    'db.OpenDuckDb ":memory:"

    v = db.QueryFast("SELECT * FROM " & db.QuoteIdent(duckTable) & " LIMIT 200;")
    Call ArrayToSheet(v, ThisWorkbook.Worksheets(1), "A1")

    db.CloseDuckDb
    MsgBox "Test OK : " & duckTable & " affichée en Feuil1", vbInformation
    Exit Sub

Fail:
    On Error Resume Next
    db.CloseDuckDb
    MsgBox "Test_CopyAccess_Clients_ODBC KO:" & vbCrLf & Err.Description & _
           IIf(Len(Native_LastErrorText) > 0, vbCrLf & Native_LastErrorText, ""), vbExclamation
End Sub

Sub Test_ExportAccessTable_ToParquet()

    Dim db As New cDuck, ok As Boolean


    db.Init ThisWorkbook.Path
    db.OpenDuckDb ":memory:"
    
    ok = AccessTable_ToParquet(db, _
            ThisWorkbook.Path & "\DbAccess.accdb", _
            "TestAdo", _
            ThisWorkbook.Path & "\access_table.parquet", _
            4)

    If ok Then MsgBox "Export OK", vbInformation
End Sub

