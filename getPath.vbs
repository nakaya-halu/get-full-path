Option Explicit

    'ドラッグアンドドロップで取得したファイルパスを変数に入れる
    Dim GetPathArray
    Set GetPathArray = WScript.Arguments
    
    'ファイルシステムオブジェクト
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'イテレータ
    Dim pt

    'ファイルの数ぶんループする
    For Each pt in GetPathArray

        '取得したファイル名
        Dim FileName
        FileName = objFSO.GetFileName(pt)

        'ファイルのフルパスをInputBoxで表示させる
        pt = InputBox("ドラッグアンドドロップしたファイル名 " & FileName,"ファイルのフルパスを表示", pt)

    Next

    'オブジェクト変数をクリア
    Set objFSO = Nothing
