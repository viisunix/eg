'---
' プログラム: EGPのファイルからSASログを抽出します。
'       説明: EGPの拡張子をZIPに変更して、ZIPのフォルダからresult.logのファイルを抽出
'     作成者: viisunix
'
'      引数1: EGPのファイルパス
'      引数2: ログファイルのパス
'     実行例: cscript foo.vbs C:\temp\foo.egp C:\temp\foo.log
'
'---

Option Explicit
On Error Resume Next

'---
' 定数
'---

Const FOF_SILENT = &H4              ' 進捗ダイアログを表示しない
Const FOF_NOCONFIRMATION = &H10     ' 上書き確認ダイアログを表示しない
Const ForWriting = 2                ' テキストファイルのオープン
Const ForReading = 1                ' テキストファイルのオープン
Const TristateUseDefault = -2       ' Opens the file using the system default.
Const TristateTrue = -1             ' Opens the file as Unicode.
Const TristateFalse = 0             ' Opens the file as ASCII.
Const DebugFlag = True

'---
' 変数
'---

Dim objShell
Dim objFso
Dim objTs
Dim sZipFile
Dim sTempFolder
Dim sEgpFilePath
Dim sLogFilePath
Dim ErrCount
Dim WarCount
Dim sMsg

'---
' オブジェクト生成します。
'---

Set objShell = CreateObject("Shell.Application")
Set objFso = CreateObject("Scripting.FileSystemObject")

'---
' 引数をチェックします。
'--

If WScript.Arguments.Count <> 2 Then
    Call MsgBox("引数1にEGP、引数2にログファイルを指定してください。", vbOKOnly + vbExclamation, WScript.ScriptName)
    WScript.Quit(1)
End If

sEgpFilePath = WScript.Arguments.Item(0)
sLogFilePath = WScript.Arguments.Item(1)

If objFso.FileExists(sEgpFilePath) = False Then
    Call MsgBox("引数1で指定したファイルが存在しません。", vbOKOnly + vbExclamation, WScript.ScriptName)
    WScript.Quit(1)
End If

'---
'   デバッグ用のメッセージ出力
'---

Sub DebugMsg(msg)
    If DebugFlag = True Then
        Call MsgBox(msg, vbOKOnly + vbInformation, WScript.ScriptName)
    End If
End Sub


'---
'   ZIPファイルを指定したフォルダに解凍
'---

Sub Unzip(objShell, sFile, sFolder)
    Dim objFilesInZip
    Dim objFolder
    
    Set objFilesInZip = objShell.Namespace(sFile).Items
    If Err.Number <> 0 Then
        Exit Sub
    End If
    Set objFolder = objShell.Namespace(sFolder)
    If Err.Number <> 0 Then
        Exit Sub
    End If
    
    If (Not objFolder Is Nothing) Then
        objFolder.CopyHere objFilesInZip, FOF_NOCONFIRMATION + FOF_SILENT
    Else
 Err.Raise 432 ' オートメーションの操作中にファイル名またはクラス名を見つけられませんでした。
    End If

   Set objFilesInZip = Nothing
   Set objFolder = Nothing
End Sub

'---
'   フォルダを作成
'---

Sub CreateUnzipFolder(objFso, sFolder)
    objFso.CreateFolder sFolder
End Sub

'---
'   フォルダを削除
'---

Sub DeleteUnzipFolder(objFso, sFolder)
    If objFso.FolderExists(sFolder) = True Then
        objFso.DeleteFolder sFolder, True
    End If
End Sub

'---
'   テンポラリのフォルダのパスを作成
'---

Function CreateFolderPath(objFso, sFolder)
    Const TemporaryFolder = 2
    Dim objTempFolder
    
    Set objTempFolder = objFso.GetSpecialFolder(TemporaryFolder)
    CreateFolderPath = objFso.BuildPath(objTempFolder.Path, sFolder)
    Set objTempFolder = Nothing
End Function

'---
'   サブフォルダからresult.logを探して、objTsに出力
'---

Sub SearchLog(objFso, objTs, tmpFolderItems)
    Const FileName = "result.log"
    Dim objFolderItemsB
    Dim objItem
    Dim Stream
    
    For Each objItem in tmpFolderItems
    
        ' 取り出した物がファイルかフォルダかを判定
        If objItem.IsFolder Then
            ' フォルダであれば、再帰呼び出しでフォルダ階層を手繰ります。
            Set objFolderItemsB = objItem.GetFolder
            Call SearchLog(objFso, objTs, objFolderItemsB.Items())
        ElseIf objItem.Name = FileName Then
            ' ファイル名が一致したら、テキストを読み取りobjTSに出力します。
            Set Stream = CreateObject("ADODB.Stream")
            Stream.Charset = "UTF-8"
            Stream.Type = 2
            Stream.Open
            Stream.LoadFromFile(objItem.Path)
            objTs.Write(Stream.ReadText)
            Stream.Close
            Set Stream = Nothing
        End If
    
    Next
    
    Set objItem = Nothing
    Set objFolderItemsB = Nothing

End Sub

'---
'   ログファイルからERROR, WARNINGの件数をカウント
'---

Sub CountLog(objFso, sLogFile, byRef ErrCount, byRef WarCount)
    Const KeyError = "e ERROR"
    Const KeyWarning = "w WARNING"
    Dim objTs
    Dim sBuf

    On Error Goto 0

    ErrCount = 0
    WarCount = 0

    Set objTs = objFso.OpenTextFile(sLogFile, ForReading, False, TristateTrue)
    If Err.Number <> 0 Then
        Exit Sub
    End If

    Do Until objTs.AtEndOfLine = True
        sBuf = objTs.ReadLine
        If Left(sBuf, Len(KeyError)) = KeyError Then
            ErrCount = ErrCount + 1
        ElseIf Left(sBuf, Len(KeyWarning)) = KeyWarning Then
            WarCount = WarCount + 1
        End If
    Loop

    objTs.Close
    If Err.Number <> 0 Then
        Exit Sub
    End If

    Set objTs = Nothing

End Sub

'---
' エラーチェック
'---

Function CheckError(fnName)
    Checkerror = False
    
    Dim strmsg
    Dim errNum
    
    If Err.Number <> 0 Then
        strmsg = "Error #" & Hex(Err.Number) & vbCrLf & "In Function " & fnName & vbCrLf & Err.Description
        Call MsgBox(strmsg, vbOkOnly + vbCritical, WScript.ScriptName)
        Checkerror = True
    End If
         
End Function

'---
' ログファイルを開きます。
'---

Set objTs = objFso.CreateTextFile(sLogFilePath, True, True)
If CheckError("objFso.CreateTextFile") Then
    WScript.Quit(1)
End If

'---
' EGPの拡張子をZIPに変更してコピーします。
'---

sZipFile = CreateFolderPath(objFso, objFso.GetTempName & ".zip")
Call DebugMsg("EGPの拡張子をZIPに変えてコピー:" & sZipFile)
objFso.CopyFile sEgpFilePath, sZipFile
If CheckError("objFso.CopyFile") Then
    WScript.Quit(1)
End If

'---
' 解凍先のテンポラリのフォルダを作成します。
'---

sTempFolder = CreateFolderPath(objFso, objFso.GetTempName)
Call DebugMsg("テンポラリのフォルダを作成:" & sTempFolder)
Call CreateUnzipFolder(objFso, sTempFolder)
If CheckError("CreateUnzipFolder") Then
    WScript.Quit(1)
End If


'---
' ZIPファイルを解凍します。
'---

Call DebugMsg("ZIPファイルを解凍:" & sZipFile)
Call Unzip(objShell, sZipFile, sTempFolder)
If CheckError("Unzip") Then
    WScript.Quit(1)
End If


'---
' テンポラリフォルダからログファイル探してobjTSに出力します。
'---

Call DebugMsg("テンポラリのフォルダからログを収集:" & sTempFolder)
Call SearchLog(objFso, objTs, (objShell.NameSpace(sTempFolder)).Items)
If CheckError("SearchLog") Then
    WScript.Quit(1)
End If

'---
' ログファイルを閉じます。
'---

objTs.Close
If CheckError("objTs.Close") Then
    WScript.Quit(1)
End If

'---
' テンポラリのフォルダを削除します。
'---

Call DebugMsg("テンポラリのフォルダを削除:" & sTempFolder)
Call DeleteUnzipFolder(objFso, sTempFolder)
If CheckError("DeleteUnzipFolder") Then
    WScript.Quit(1)
End If

'---
' ZIPファイルを削除します。
'---

Call DebugMsg("ZIPファイルを削除:" & sZipFile)
objFso.DeleteFile sZipFile
If CheckError("objFso.DeleteFile") Then
    WScript.Quit(1)
End If

'---
' Error, Warningの件数を数えます。
'---

Call CountLog(objFso, sLogFilePath, ErrCount, WarCount)
If CheckError("objFso.DeleteFile") Then
    WScript.Quit(1)
End If
Call DebugMsg("ERROR件数:" & CStr(ErrCount) & " WARNING件数:" & CStr(WarCount))

'---
' オブジェクトを破棄します。
'---

Set objTs = Nothing
Set objFso = Nothing
Set objShell = Nothing

'---
' 終了
'---

WScript.Quit(0)


