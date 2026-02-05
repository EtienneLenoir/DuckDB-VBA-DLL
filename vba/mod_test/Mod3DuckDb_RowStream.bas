Attribute VB_Name = "Mod3DuckDb_RowStream"
Option Explicit
    
Sub IngestWithAppender()


'L 'utilisation d'un appender est la méthode la plus efficace pour charger des données dans DuckDB depuis l'interface C et est recommandée pour un chargement rapide.
'L 'appender est bien plus rapide que l'utilisation de requêtes préparées ou INSERT INTOde requêtes individuelles.

    'Appender tuyau d’ingestion en RAM + controle du typage, pour lecture streaming ligne par ligne
    'Autre cas utilise DuckVba_AppendArrayV ou COPY FROM read_csv_auto/read_parquet
        'Ex: ingérer de manière incrémentale
        'Ex
        'Do Until EOF(f)
            'DuckVba_AppenderBeginRow app
            'DuckVba_AppendInt64 app, ID
            'DuckVba_AppendVarcharW app, PW(name)
            'DuckVba_AppendDouble app, amount
            'DuckVba_AppenderEndRow app
        'Loop
        
    Dim db As New cDuck, app As LongPtr, a As Variant
    
    db.Init ThisWorkbook.Path
    db.OpenDuckDb ThisWorkbook.Path & "\demo.duckdb"

    ' 1) Crée la table destination (si besoin)
    db.Exec "DROP TABLE IF EXISTS main.people;"
    db.Exec "CREATE TABLE main.people(id BIGINT, name VARCHAR, birthday DATE, active BOOLEAN);"

    ' 2) Ouvre un appender sur main.people
    app = DuckVba_AppenderOpen(db.handle, PNull(), PW("people")) 'schema NULL = main
    If app = 0 Then Err.Raise 5, , "AppenderOpen: " & Duck_LastErrorText()

    ' 3) Transaction (fortement recommandé)
    db.Exec "BEGIN;"

    ' --------- Ligne 1 ----------
    If DuckVba_AppenderBeginRow(app) = 0 Then GoTo Fail
    Call DuckVba_AppendInt64(app, 1)                       'id
    Call DuckVba_AppendVarcharW(app, PW("Alice"))          'name
    Call DuckVba_AppendDateYMD(app, 1990, 5, 17)           'birthday
    Call DuckVba_AppendBool(app, 1)                        'active
    If DuckVba_AppenderEndRow(app) = 0 Then GoTo Fail

    ' --------- Ligne 2 ----------
    If DuckVba_AppenderBeginRow(app) = 0 Then GoTo Fail
    Call DuckVba_AppendInt64(app, 2)
    Call DuckVba_AppendVarcharW(app, PW("Bob"))
    Call DuckVba_AppendNull(app)
    Call DuckVba_AppendBool(app, 0)
    If DuckVba_AppenderEndRow(app) = 0 Then GoTo Fail

    ' 4) Commit + close
    db.Exec "COMMIT;"
    Call DuckVba_AppenderClose(app)

    ' 5) Vérifie
    a = db.QueryFast("SELECT * FROM main.people ORDER BY id;")
    ShowArrayOnSheet a
    db.CloseDuckDb
    Exit Sub

Fail:
    db.Exec "ROLLBACK;"
    If app <> 0 Then Call DuckVba_AppenderClose(app)
    Err.Raise 5, , Duck_LastErrorText()
End Sub

