Attribute VB_Name = "Mod2DuckDb_PythonLike"
Option Explicit

'===============================================================================
' Demo : FrameFromValue "Python-like" (VBA array -> DuckDB table)
'
' Idée
'   Reproduire un workflow type Python/pandas, mais en VBA :
'   - exemple un Variant(2D) en mémoire (array dynamique)
'   - On le "matérialises" côté DuckDB via FrameFromValue
'   - ensuite requêtes SQL ultra-rapides dessus (JOIN/GROUP BY/etc.)
'
' Deux modes
'   1) makeTemp:=True  -> CREATE TEMP TABLE : table en mémoire (session courante)
'      - parfait pour traitements intermédiaires, itérations, "what-if", macros rapides
'      - zéro fichier, zéro I/O disque, pas de staging CSV
'
'   2) makeTemp:=False -> CREATE TABLE persistante : stockée dans un fichier .duckdb
'      - utile  pour réutiliser les données après fermeture/réouverture d’Excel
'
' Potentiel / cas d’usage
'   - "DataFrame" local : prototypage rapide, transformations SQL, checks qualité
'   - Pipeline hybride : Excel/VBA pour préparer, DuckDB pour calculer (OLAP)
'
' Note
'   - hasHeader:=True : la 1ère ligne de l’array contient les noms de colonnes.
'   - Attention aux types : DuckDB infère/convertit (numériques, textes, dates).
'===============================================================================

Sub Test_FrameFromValue_PythonLike()

    Dim db As New cDuck, v As Variant, a As Variant, dbPath As String
    '--- 1) On fabrique un petit array 2D en mémoire ---
    ReDim v(1 To 4, 1 To 3)
    
    'Header
    v(1, 1) = "id"
    v(1, 2) = "nom"
    v(1, 3) = "valeur"
    
    'Lignes
    v(2, 1) = 1: v(2, 2) = "Alice": v(2, 3) = 10.5
    v(3, 1) = 2: v(3, 2) = "Bob":   v(3, 3) = 20
    v(4, 1) = 3: v(4, 2) = "Chloé": v(4, 3) = 30
    
    db.Init ThisWorkbook.Path
    
    '===================================================
    ' CAS 1 : table TEMP (CREATE TEMP TABLE __frame_x)
    '===================================================
    db.OpenDuckDb ":memory:"
    
    'Crée une table TEMP à partir de l’array
    db.FrameFromValue "aa", v, True, True   ' hasHeader:=True, makeTemp:=True
    
    'On lit ce qu’il y a dedans
    a = db.QueryFast("SELECT * FROM aa ")
    DebugArray2DTable a, "TEMP __frame_x"
    
    If Not IsEmpty(a) Then
        With ThisWorkbook.Worksheets("Feuil1")
            .Cells.Clear
            .Range("A1").Resize(UBound(a, 1), UBound(a, 2)).Value = a
        End With
    End If
    
    db.CloseDuckDb
    
    '===================================================
    ' CAS 2 : table PERSISTANTE dans un fichier DuckDB
    '===================================================
    dbPath = ThisWorkbook.Path & "\frame_test.duckdb"
    
    '--- Première ouverture : création de la table ---
    db.OpenDuckDb dbPath
    'Crée une table PERSISTANTE FramePersist
    db.FrameFromValue "FramePersist", v, True, False   ' makeTemp:=False => CREATE TABLE
    
    a = db.QueryFast("SELECT * FROM FramePersist ORDER BY id;")
    DebugArray2DTable a, "FramePersist (session 1)"
    
    db.CloseDuckDb
    
    '--- Deuxième ouverture : on vérifie que la table existe toujours ---
    db.OpenDuckDb dbPath

    a = db.QueryFast("SELECT * FROM FramePersist ORDER BY id;")
    DebugArray2DTable a, "FramePersist (session 2, après réouverture)"
    
    'Optionnel : envoyer au Sheet1
    If Not IsEmpty(a) Then
        With ThisWorkbook.Worksheets(1)
            .Cells.Clear
            .Range("A1").Resize(UBound(a, 1), UBound(a, 2)).Value = a
            .Columns.AutoFit
        End With
    End If
    
    db.CloseDuckDb

End Sub
