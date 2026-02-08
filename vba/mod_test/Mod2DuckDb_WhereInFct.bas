Attribute VB_Name = "Mod2DuckDb_WhereInFct"
Option Explicit

'===============================================================================
' Module : Bench_TempList
'
' Objectif
'   Démonstrations + bench autour des "temp lists" DuckDB (table temporaire)
'   alimentées depuis VBA (Variant 1D), pour accélérer les filtres de type
'   "WHERE x IN (liste)" quand la liste devient grande.
'
' Idée
'   - Cas réel finance de marché : filtrer un univers d’instruments (ISIN, ticker,
'     FIGI...) sur des milliers / dizaines de milliers de clés.
'   - Plutôt que construire une énorme clause IN(...) (string longue, parse SQL,
'     compilation, etc.), on envoie la liste une fois à DuckDB sous forme de
'     table temporaire (tmp_ids) puis on requête via JOIN / IN sur (SELECT v ...).
'
' Fonctions utilisées (cDuck -> DLL)
'   - db.CreateTempList(tempName, keysVariant1D, sqlType)
'       -> DuckVba_CreateTempListV : crée tempName(v) et y insère les clés.
'   - db.QueryFast(sql)
'       -> DuckVba_QueryToArrayFastV : exécute un SELECT et retourne Variant(2D).
'   - db.SelectWithTempList(tempName, keysVariant1D, sqlType, selectOrTable, joinCol, autoJoin)
'       -> DuckVba_SelectWithTempList2V : crée la temp list puis exécute selon 2 modes :
'
'       MODE 1 : autoJoin = True   (simple & rapide)
'         - selectOrTable = nom de table (ou vue) cible, joinCol = colonne de jointure
'         - la DLL fabrique le SELECT automatiquement :  ... FROM <table> JOIN <tempName> ON <joinCol>=v
'
'       MODE 2 : autoJoin = False  (mode libre)
'         - selectOrTable = SQL complet écrit par toi
'         - ton SQL doit référencer la temp list : ... WHERE x IN (SELECT v FROM <tempName>)  (ou JOIN manuel)
'         - la DLL fournit seulement la table temporaire de clés, et exécute ton SQL tel quel.
'
'   - (Benchmark) Prepared statements
'       -> Prepare / Bind / Exec pour remplir rapidement une table de test.
'
' Notes
'   - Les temp tables sont en général plus stables et plus performantes dès que
'     la liste de clés est grande (p.ex. > quelques milliers), mais le gagnant
'     dépend du run, des tailles et du cache.
'   - Résultat du bench : compare WHERE IN(...) vs temp table + JOIN sur COUNT(*).
'===============================================================================

Sub Demo_CreateInsertDispaly() 'Demo Create + Insert + Display

    Dim db As New cDuck, a As Variant, ws As Worksheet, sDbPath As String

    'Base créée si introuvable
    sDbPath = ThisWorkbook.Path & "\Db_DuckDb_Exemple.duckdb"

    db.Init 'Chemin du DLL : par défaut celui du classeur
    db.OpenDuckDb sDbPath

    'Schéma + données de test
    db.Exec "DROP TABLE IF EXISTS Instruments;"
    db.Exec "CREATE TABLE Instruments (" & _
             "ISIN VARCHAR, " & _
             "Nom VARCHAR, " & _
             "Prix DOUBLE, " & _
             "ModifiedAt TIMESTAMP);"

    db.BeginTx
    db.Exec _
        "INSERT INTO Instruments VALUES " & _
        "('FR0000131104','TotalEnergies',64.25, NOW())," & _
        "('US0378331005','Apple',215.36, NOW())," & _
        "('US5949181045','Microsoft',420.58, NOW())," & _
        "('NL0000009355','Shell',32.14, NOW());"
    db.Commit

    Set ws = ThisWorkbook.Worksheets(1)

    a = db.QueryFast("SELECT * FROM Instruments ORDER BY ISIN;")

    ws.Cells.Clear
    If Not IsEmpty(a) Then
        ws.Range("A1").Resize(UBound(a, 1), UBound(a, 2)).Value = a
    End If

    db.CloseDuckDb
    MsgBox "OK : données écrites sur " & ws.name, vbInformation

End Sub

Public Sub Demo_SelectWithTempList_AutoJoin()

    'Test Fct Custom Méthode Fast "WHERE IN" : db.SelectWithTempList
        '1 arg: nom table temp
        '2 arg : Array
        '3 arg: sqlType="VARCHAR"
        '4 arg : Nom table
        '5 arg : Field Join
        '6 arg : True = Méthode Fast création par dll C d'une table temporaire pour jointure
              ': False = Méthode création de la liste auto "WHERE IN"

    Dim db As New cDuck, ws As Worksheet, keys As Variant, out As Variant
      
    '--- Session RO via cDuck (remplace h = Duck_OpenReadOnly(...)) ---
    db.Init ThisWorkbook.Path
    db.OpenReadOnly ThisWorkbook.Path & "\Db_DuckDb_Exemple.duckdb"

    '--- Clés : idem (peuvent venir d'une plage si tu veux) ---
    keys = Array("FR0000131104", "NL0000009355")

    ' --- Exécution : temp list + auto-join sur Instruments.ISIN ---
    out = db.SelectWithTempList( _
            "__tmp_list", keys, "VARCHAR", "Instruments", "ISIN", True) 'True = autoJoin (équiv. 1)

    ' --- Affichage ---
    Set ws = ThisWorkbook.Worksheets(1)
    Call ArrayToSheet(out, ws, "A1")
    MsgBox "Auto-join OK : " & (UBound(out, 1) - 1) & " lignes.", vbInformation

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
    Resume CleanExit
End Sub

Public Sub Demo_SelectWithTempList()

    'Test Fct Custom Méthode Fast "WHERE IN" : db.SelectWithTempList
        '1 arg: nom table temp
        '2 arg : Array
        '3 arg: sqlType="VARCHAR"
        '4 arg : Nom table
        '5 arg : Field Join
        '6 arg : True = Méthode Fast création par dll C d'une table temporaire pour jointure
              ': False = Méthode création de la liste auto "WHERE IN"

    Dim db As New cDuck, ws As Worksheet, keys As Variant, out As Variant, sql As String
    
    On Error GoTo Fail

    ' --- Session RO via cDuck (remplace h = Duck_OpenReadOnly(...)) ---
    db.Init ThisWorkbook.Path
    db.OpenReadOnly ThisWorkbook.Path & "\Db_DuckDb_Exemple.duckdb"

    ' --- Clés ---
    keys = Array("FR0000131104", "NL0000009355")

    ' --- SQL libre (réécriture auto des [identifiants] par la DLL) ---
    sql = "SELECT [ISIN],[Prix], " & _
          "strftime(ModifiedAt, '%Y-%m-%d %H:%M:%S') AS [Date] " & _
          "FROM Instruments WHERE [ISIN] IN (SELECT v FROM __tmp_free) " & _
          "ORDER BY [ISIN]"

    ' --- Temp list + exécution (free mode) ---
    out = db.SelectWithTempList( _
            "__tmp_free", _
            keys, _
            "VARCHAR", _
            sql, _
            "", _
            False)

    Set ws = ThisWorkbook.Worksheets(1)
    Call ArrayToSheet(out, ws, "A1")
    MsgBox "Mode libre OK : " & (UBound(out, 1) - 1) & " lignes.", vbInformation

CleanExit:
    On Error Resume Next
    db.CloseDuckDb
    Exit Sub

Fail:
    MsgBox "Erreur: " & Err.Description & vbCrLf & Native_LastErrorText(), vbExclamation
    Resume CleanExit
    
End Sub



