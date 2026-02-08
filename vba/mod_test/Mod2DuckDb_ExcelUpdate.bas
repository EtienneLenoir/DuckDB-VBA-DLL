Attribute VB_Name = "Mod2DuckDb_ExcelUpdate"
Option Explicit

'===============================================================================
' Module : ExcelUpdate Function PushExcelToDuck
'
' Objectif :
'   Synchroniser une table DuckDB <-> une feuille Excel.
'
' Fonctions principales :
'   1) ReloadFromDuckToExcel
'        - Ouvre la base DuckDB (demo.duckdb)
'        - Lit la table cible (tableName)
'        - Dépose le résultat en feuille (A1)
'
'   2) PushExcelToDuck
'        - Lit la zone Excel à partir de A1 (CurrentRegion)
'        - La 1ère ligne DOIT contenir les en-têtes de colonnes EXACTS de DuckDB
'        - Upsert vers DuckDB via db.UpsertFromArray (UPDATE si clé existe, sinon INSERT)
'        - Relit la table et rafraîchit la feuille pour contrôle visuel
'
' Paramètres à adapter :
'   - duckPath  : chemin de la base (par défaut: ThisWorkbook.Path & "\demo.duckdb")
'   - tableName : nom de la table DuckDB (par défaut: "data")
'   - keyCols   : colonne(s) clé(s) pour l'upsert (ex: "ISIN" ou "id" ou "k1,k2")
'
' Notes
'   - PushExcelToDuck suppose que la table existe déjà dans DuckDB
'     (ou que ta DLL gère la création côté UpsertFromArray)
'   - Pour des gros volumes, éviter de rafraîchir toute la feuille si inutile
'===============================================================================

Public Sub ReloadFromDuckToExcel()

    'lancez 'sub Demo_Euronext_Csv' avant
    
    Dim db As New cDuck, preview As Variant, duckPath As String, tableName As String, msg As String

    On Error GoTo Fail

    duckPath = ThisWorkbook.Path & "\demo.duckdb"
    tableName = "ImportedCsv"

    db.Init ThisWorkbook.Path
    db.OpenDuckDb duckPath

    preview = db.QueryFast("SELECT * FROM " & tableName & ";")
    Call ArrayToSheet(preview, ThisWorkbook.Worksheets(1), "A1")

    MsgBox "Feuil1 rafraîchie depuis DuckDB.", vbInformation

CleanExit:
    db.CloseDuckDb
    Exit Sub
Fail:
    msg = "ERREUR ReloadFromDuckToExcel:" & vbCrLf & _
          Err.Description & vbCrLf & _
          "DLL says: " & Native_LastErrorText()
    MsgBox msg, vbCritical
    Resume CleanExit
    
End Sub

'===============================================================================
' Workflow (édition Excel -> synchronisation DuckDB)
'
' 1) Rafraîchis d'abord la feuille depuis DuckDB :
'       -> lance ReloadFromDuckToExcel
'
' 2) Modifie ensuite des valeurs directement dans la feuille (Feuil1) :
'       - ne change pas les en-têtes en ligne 1
'       - évite de casser le "bloc" (CurrentRegion) : lignes/colonnes vides au milieu
'
' 3) Pousse tes changements vers DuckDB (UPSERT) :
'       -> lance PushExcelToDuck
'
' 4) Contrôle / validation :
'       - la macro relit la table et réécrit la feuille
'       - tu peux aussi relancer ReloadFromDuckToExcel pour confirmer
'
' Notes:
'   - La/les colonne(s) clé(s) définies dans keyCols déterminent UPDATE vs INSERT.
'   - Si keyCols est mauvais/incomplet, tu risques des doublons ou des updates inattendus.
'===============================================================================

Public Sub PushExcelToDuck()

    Dim db          As New cDuck, ws As Worksheet, arr As Variant, preview As Variant, colsDb As Variant
    Dim duckPath    As String, tableName As String, keyCols As String, msg As String, i As Long

    On Error GoTo Fail

    '1) paramètres de base
    duckPath = ThisWorkbook.Path & "\demo.duckdb"
    tableName = "ImportedCsv"      ' la même table que tu remplis avec Test_ReadCsvToTable
    keyCols = "ISIN"                  ' <<< ADAPTE ICI : le ou les champs clé(s) pour faire l'UPDATE
                                    'ex: "id"
                                    'ex: "isin"
                                    'ex: "code_client,date_operation"
 
    db.Init ThisWorkbook.Path
    db.OpenDuckDb duckPath

    ' lis la zone de la feuille vers un Variant 2D
    '    IMPORTANT :
    '    - La 1ère ligne de la zone doit être les en-têtes de colonnes EXACTES (mêmes noms que dans DuckDB),
    '      par ex: id | nom | age | ville
    '    - Les lignes suivantes sont les données
    '    - CurrentRegion part de A1 et prend le bloc rempli
    Set ws = ThisWorkbook.Worksheets(1)  ' ou "Feuil1" si tu préfères le nom
    arr = ws.Range("A1").CurrentRegion.Value

    colsDb = db.QueryFast("SELECT column_name FROM information_schema.columns " & _
                          "WHERE lower(table_schema)=lower('main') AND lower(table_name)=lower('" & tableName & "') " & _
                          "ORDER BY ordinal_position;")
    Debug.Print "Cols DB:"
    For i = 2 To UBound(colsDb, 1) ' 1 = entête dans QueryFast
        Debug.Print " - "; CStr(colsDb(i, 1))
    Next
    'entête Excel
    Debug.Print "Cols Excel:"
    For i = LBound(arr, 2) To UBound(arr, 2)
        Debug.Print " - "; CStr(arr(1, i))
    Next

    'Call DebugArray2DTable(arr, "clients après import")

    ' Upsert vers DuckDB
    '    headerRow := 1 -> dit à la DLL : la ligne 1 contient les noms de colonnes
    '    keyCols    -> colonne(s) utilisée(s) pour savoir si on UPDATE ou INSERT
    '
    '    Effet :
    '      - si une ligne avec le même key existe déjà dans main.clients :
    '            UPDATE des autres colonnes avec les nouvelles valeurs de la feuille
    '      - sinon :
    '            INSERT d'une nouvelle ligne
    '
    '    Si tu as juste modifié une cellule sur la feuille (ex: tu changes l'adresse d'un client),
    '    cette nouvelle valeur va écraser l'ancienne en base pour cette ligne.
    Call db.UpsertFromArray(tableName, arr, 1, keyCols)

    preview = db.QueryFast("SELECT * FROM " & tableName & ";")
    Call ArrayToSheet(preview, ws, "A1")

    MsgBox "Mise à jour DuckDB OK + rafraîchissement Excel terminé.", vbInformation

CleanExit:
    db.CloseDuckDb
    Exit Sub
Fail:
    msg = "ERREUR PushExcelToDuck:" & vbCrLf & _
          Err.Description & vbCrLf & _
          "DLL says: " & Native_LastErrorText()
    MsgBox msg, vbCritical
    Resume CleanExit
End Sub

