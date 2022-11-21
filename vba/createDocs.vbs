Option Explicit

' 空白 そのまま
' ; 1行コメント
' | これ以降コメント

'' ########################メインルーチン###################################
Dim OUTPUT_TEXT_NAME, OUTPUT_HTML_NAME, THIS_SCRIPT_NAME, THIS_VERSION, OUTPUT_TEXT_NAME_EX
OUTPUT_TEXT_NAME = "__Files.txt"
OUTPUT_TEXT_NAME_EX = "___Files.txt"
OUTPUT_HTML_NAME = "__Files.html"
THIS_SCRIPT_NAME = WScript.ScriptName  ' スクリプト名の取得
THIS_VERSION = "0.0.9.0"

Dim MSG_ADD_FILES
Dim MSG_DEL_FILES
MSG_ADD_FILES = ";次のファイルは追加されました"
MSG_DEL_FILES = ";次のファイルは削除されています"

Dim ignoreFolders
ignoreFolders = Split("out,.vscode,obj,bin,.buildLog,packages", ",")
'' ignoreFolders = Split("", ",")


Dim XFC
Set XFC = new XFileInfoCreator

call Main()



Class XFiles
    Dim m_Files
    Dim FSO

    Private Sub Class_Initialize
        Set m_Files = CreateObject("Scripting.Dictionary")
        Set FSO = CreateObject("Scripting.FileSystemObject")
    End Sub

    Public Function Files()
        set Files = m_Files
    End Function

    Public Sub Clear()
        Set m_Files = CreateObject("Scripting.Dictionary")
    End Sub

    Public Sub Add(key, value)
        m_Files.Add key, value
    End Sub

    Public Function Count()
        Count = m_Files.Count
    End Function

    Public Sub Remove(key)
        m_Files.remove key
    End Sub

    Public Sub Load(filePath)
        Dim pathFile
        dim ppath
        Dim currentLine, tmpSplitLine, comment
        Dim firstChar
        Set m_Files = CreateObject("Scripting.Dictionary")

        If FSO.FileExists(filePath) = false Then
            exit Sub
        End If
        With CreateObject("ADODB.Stream")
            .Charset = "UTF-8"
            .Open
            .LoadFromFile filePath
            Do Until .EOS
                ' adReadAll ' -1
                ' 既定値。現在の位置から EOS マーカー方向に、すべてのバイトをストリームから読み取ります。
                'これは、バイナリ ストリーム (Type は adTypeBinary) に唯一有効な StreamReadEnum 値です。
                ' adReadLine' -2
                ' ストリームから次の行を読み取ります (LineSeparator プロパティで指定)。
                '[; ] relative path | comment | info
                dim a
                currentLine = .ReadText(-2)
                firstChar = left(trim(currentLine), 1)
                If trim(currentLine) = "" Then
                    set a = XFC.CreateXFileInfo(CreateObject("Scriptlet.TypeLib").GUID, currentLine, "", False, True)
                    m_Files.add a.FileID, a
                elseif  firstChar = ";" Then
                    currentLine = Mid(currentLine, 2)
                    set a = XFC.CreateXFileInfo(CreateObject("Scriptlet.TypeLib").GUID, currentLine, "", False, True)
                    m_Files.add a.FileID, a
                Else
                    tmpSplitLine = split((currentLine), "|")  ''''''''''''' 大文字小文字区別
                    pathFile = trim(tmpSplitLine(0))
                    comment = ""
                    If Ubound(tmpSplitLine) > 0 Then
                        comment = trim(tmpSplitLine(1))
                    End If
                    If m_Files.Exists(pathFile) Then
                        debugPrint 1, "file not exist", pathFile
                    Else
                        set a = XFC.CreateXFileInfo(pathFile, pathFile, comment, False, False)
                        m_Files.add a.FileID, a
                    End If
                End If
            Loop
            .Close
        End with
    End Sub

    Sub GetFilesInFolder(RootPath, Path)
        debugPrint  0, RootPath, Path
        dim a
        If FSO.GetFolder(Path).Attributes And (2) Then
            exit Sub
        End If
        '値	属性
        '0 	 標準ファイル
        '1 	 読み取り専用ファイル
        '2 	 隠しファイル
        '4 	 システムファイル
        '8 	 ディスクドライブボリュームラベル(取得のみ可能)
        '16 	 フォルダまたはディレクトリ(取得のみ可能)
        '32 	 アーカイブファイル
        '64 	 リンクまたはショートカット(取得のみ可能)
        '128 	 圧縮ファイル(取得のみ可能)


        Dim folder
        Set folder = FSO.GetFolder(Path)

        Dim filesInFolder
        Set filesInFolder = folder.Files

        Dim ppath
        Dim pathFile

        ppath = Mid(Path, len(RootPath) + 1)  
        If filesInFolder.Count = 0 Then
            pathFile = Mid((ppath & "\" & "."), 2) '最初の\を除く, フォルダにファイルがないと タイトルが出力されないから
            set a = XFC.CreateXFileInfo(pathFile, pathFile, "", False, False)
            m_Files.add a.FileID, a
        End If

        Dim fileObject
        For Each fileObject In filesInFolder
            pathFile = Mid((ppath & "\" & fileObject.Name), 2) '最初の\を除く
            If m_Files.Exists(pathFile) = False Then
                If pathFile = OUTPUT_TEXT_NAME or pathFile = OUTPUT_HTML_NAME or pathFile = THIS_SCRIPT_NAME Then
                    'no operation
                Else
                    set a = XFC.CreateXFileInfo(pathFile, pathFile, "", False, False)
                    m_Files.add a.FileID, a
                End If
            End If
        Next

        Dim subFolders, folderObject
        Set subFolders = folder.subFolders
        For Each folderObject In subFolders
            If isIgnoreDir(folderObject.Name) Then
                rem
            Else
                GetFilesInFolder RootPath, (Path & "\" & folderObject.Name)
            End If
        Next
    End Sub

    '' merged : information form folder
    Sub Merge(merged)
        Dim dicMerged
        Set dicMerged = CreateObject("Scripting.Dictionary")

        Dim FilesInFolder
        Set FilesInFolder = merged.Files

        dim i
        dim itemsInFile
        dim itemInFile
        itemsInFile = m_Files.Items
        'add delete flg to ListInFile
        For i = 0 to m_Files.Count - 1
            set itemInFile = itemsInFile(i)
            If FilesInFolder.Exists(itemInFile.FileID) Then
                '' FilesInFolderには新規のファイルだけにするので、存在するものは削除
                FilesInFolder.Remove itemInFile.FileID
            Else
                '' フォルダにないものは削除されている
                itemInFile.IsDelete = True
            End If
        next

        dim itemsInFolder
        dim itemInfolder
        itemsInFolder = FilesInFolder.Items
        dim previousFolder
        previousFolder = "djdkdkdlslslslsldkdkdkdkdkd"
        dim newFiles

        'add ListInFile to Merge List
        For i = 0 to m_Files.Count - 1
            set itemInFile = itemsInFile(i)
            if itemInFile.folder <> previousFolder Then
                debugPrint 1, "Change Folder", itemInFile.folder
                set newFiles = GetFilesAtFolder(FilesInFolder, itemInFile.folder)
                debugPrint 1, "GetFiles", newFiles.Count
                call AddFiles(dicMerged, newFiles)
                previousFolder = itemInFile.folder
            End If

            If dicMerged.Exists(itemInFile.FileID) Then
                debugPrint 1, "Merge Err", itemInFile.FileID
            Else
               dicMerged.Add itemInFile.FileID, itemInFile
            End If
        next

        '' add items not merged
        call AddFiles(dicMerged, FilesInFolder)

        set m_Files = dicMerged
    End Sub

    Sub AddFiles(dicTo, dicFrom)
        dim items
        dim item
        items = dicFrom.Items
        dim i
        For i = 0 to dicFrom.Count - 1
            dicTo.add items(i).FileID, items(i)
        next
    End Sub

    Function GetFilesAtFolder(dicFiles, folderName)
        Dim NewFilesAtFolder
        Set NewFilesAtFolder = CreateObject("Scripting.Dictionary")

        Dim newItems
        newItems = dicFiles.items
        Dim iNew
        Dim newItem
        For iNew = 0 To dicFiles.Count - 1
            set newItem = newItems(iNew)
            If newItem.folder = folderName Then
                NewFilesAtFolder.Add newItem.FileID, newItem
                dicFiles.Remove newItem.FileID
            End If
        Next
        Set GetFilesAtFolder = NewFilesAtFolder
    End Function

    Sub WriteToFile(filename)
        'Dim writeStream As ADODB.Stream
        'Microsoft ActiveX Data Objects 2.5 Libraryと
        ' WriteText str, 1 => add a newline
        ' WriteText str, 0 => add no newline
        Dim writeStream

        ' 文字コードを指定してファイルをオープン
        Set writeStream = CreateObject("ADODB.Stream")
        writeStream.Charset = "UTF-8"
        writeStream.Open

        '実際の中身の書き込み
        Dim i, items, flgComment
        items = m_Files.items
        For i = 0 to m_Files.Count -1
            flgComment = ""
            if items(i).IsDelete Then
                flgComment = ";"
            End if
            writeStream.WriteText flgComment & items(i).relativePath, 0
            if items(i).Comment <> "" Then
                writeStream.WriteText " | ", 0
                writeStream.WriteText items(i).Comment, 0
            End if
            writeStream.WriteText "", 1
        next
        ' ファイルに書き込み
        writeStream.SaveToFile filename, 2 'adSaveCreateOverWrite:2

        ' ファイルをクローズ
        writeStream.Close
        Set writeStream = Nothing
    End Sub

    Sub WriteToFileEx(filename)
        'Dim writeStream As ADODB.Stream
        'Microsoft ActiveX Data Objects 2.5 Libraryと
        ' WriteText str, 1 => add a newline
        ' WriteText str, 0 => add no newline
        Dim writeStream

        ' 文字コードを指定してファイルをオープン
        Set writeStream = CreateObject("ADODB.Stream")
        writeStream.Charset = "UTF-8"
        writeStream.Open

        '実際の中身の書き込み
        Dim i, items, flgComment
        items = m_Files.items
        For i = 0 to m_Files.Count -1
            flgComment = ""
            if items(i).IsDelete Then
                flgComment = ";"
            End if
            
            writeStream.WriteText flgComment & items(i).relativePath, 0
            if items(i).Comment <> "" Then
                writeStream.WriteText " | ", 0
                writeStream.WriteText items(i).Comment, 0
            End if
            writeStream.WriteText "", 1
        next
        ' ファイルに書き込み
        writeStream.SaveToFile filename, 2 'adSaveCreateOverWrite:2

        ' ファイルをクローズ
        writeStream.Close
        Set writeStream = Nothing
    End Sub

    Function ToDebugString()
        dim items
        items = m_Files.Items
        dim i
        debugPrint 1, "Files.Count", m_Files.Count
        For i = 0 to m_Files.Count - 1
            debugPrint 1, items(i).ToString(), ""
        next
    End Function
End Class

Class XFileInfoCreator
    Dim FSO
    Private Sub Class_Initialize
        Set FSO = CreateObject("Scripting.FileSystemObject")
    End Sub
    Function CreateXFileInfo(FileID, relativePath, Comment, IsDelete, IsComment)
        dim f
        set f = new XFileInfo
        f.FileID = FileID
        f.relativePath = relativePath
        f.Comment = Comment
        f.IsDelete = IsDelete
        f.IsComment = IsComment

        f.Folder = FSO.GetParentFolderName(relativePath)
        f.FileName = FSO.GetFileName(relativePath)

        Set CreateXFileInfo = f
    End Function
End Class

Class XFileInfo
  Public FileID
  Public Name
  Public relativePath
  Public Folder
  Public FileName
  Public Comment
  Public IsDelete
  Public IsComment
  Public Function ToString()
    ToString = Folder & ": " & FileName & ": " & Comment
  End Function
End Class

''///////////////////////////////////////////////////////////////


Sub Main()

    Dim TIME_DATE
    TIME_DATE = Now

    Dim RootPath
    RootPath = "."

    If WScript.Arguments.Count > 0 Then
        RootPath = WScript.Arguments(0)
    End If

    dim fileList
    set fileList = new XFiles
    '' test.Load(OUTPUT_TEXT_NAME)
    call fileList.Load(OUTPUT_TEXT_NAME)

    dim folderList
    set folderList = new XFiles
    call folderList.GetFilesInFolder(".",".")

    debugPrint 1, "fileList.Count", fileList.Count
    debugPrint 1, "folderList.Count", folderList.Count
    call fileList.Merge(folderList)
    debugPrint 1, "fileList.Count", fileList.Count
    debugPrint 1, "folderList.Count", folderList.Count

    call  fileList.WriteToFile(OUTPUT_TEXT_NAME)


    WScript.Quit


    Call GetFilesInFolder(RootPath, RootPath)

    WScript.Quit

    debugPrint 0, "FilesInFolders.Count", FilesInFolders.Count
    debugPrint 0, "DeleteFiles.Count", DeleteFiles.Count
    debugPrint 0, "ExistFiles.Count", ExistFiles.Count

    Call CheckList

    call AddNewFileToExist

    WriteToFile OUTPUT_TEXT_NAME, FilesInFolders, MSG_ADD_FILES & TIME_DATE , true
    WriteToFile OUTPUT_TEXT_NAME, DeleteFiles, MSG_DEL_FILES & TIME_DATE , false
    WriteToFile OUTPUT_TEXT_NAME, ExistFiles, "", false

    WriteToFileHTML FilesInFolders, DeleteFiles, ExistFiles

    debugPrint 0, "FilesInFolders.Count", FilesInFolders.Count
    debugPrint 0, "DeleteFiles.Count", DeleteFiles.Count
    debugPrint 0, "ExistFiles.Count", ExistFiles.Count

    InfoPrint 1,  "ファイルを出力しました", OUTPUT_TEXT_NAME
End Sub
WScript.Quit

'システムに存在するファイルを取得する
'フルパスで取得
'隠しフォルダは処理しない

Sub GetFilesInFolder(RootPath, Path)
    debugPrint  0, RootPath, Path
    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.GetFolder(Path).Attributes And (2) Then
        exit Sub
    End If
     '値	属性
     '0 	 標準ファイル
     '1 	 読み取り専用ファイル
     '2 	 隠しファイル
     '4 	 システムファイル
     '8 	 ディスクドライブボリュームラベル(取得のみ可能)
     '16 	 フォルダまたはディレクトリ(取得のみ可能)
     '32 	 アーカイブファイル
     '64 	 リンクまたはショートカット(取得のみ可能)
     '128 	 圧縮ファイル(取得のみ可能)


    Dim folder
    Set folder = FSO.GetFolder(Path)

    Dim files
    Set files = folder.Files

    Dim ppath
    Dim pathFile

    ppath = Mid(Path, len(RootPath) + 1)  
    If files.Count = 0 Then
        pathFile = Mid((ppath & "\" & "."), 2) '最初の\を除く, フォルダにファイルがないと タイトルが出力されないから
        FilesInFolders.add (pathFile), pathFile
    End If

    Dim fileObject
    For Each fileObject In files
        pathFile = Mid((ppath & "\" & fileObject.Name), 2) '最初の\を除く
        If FilesInFolders.Exists(pathFile) = False Then
            If pathFile = OUTPUT_TEXT_NAME or pathFile = OUTPUT_HTML_NAME or pathFile = THIS_SCRIPT_NAME Then
                'no operation
            Else
                dim f
                set f = new XFileInfo
                f.Folder = pathFile
                f.FileName = GetParentFolderName(pathFile)
                f.Comment = ""
                ''FilesInFolders.add (pathFile), pathFile
                FilesInFolders.add pathFile, f
                debugPrint 1, "XFileInfo", f.ToString
            End If
        End If
    Next

    Dim subFolders, folderObject
    Set subFolders = folder.subFolders
    For Each folderObject In subFolders
        If isIgnoreDir(folderObject.Name) Then
            rem
        Else
            GetFilesInFolder RootPath, (Path & "\" & folderObject.Name)
        End If
    Next
End Sub

'必要なファイルの存在確認（不必要なファイルはこの次にチェックする）
Sub CheckList()
    ' FilesInFolders 実際に存在してるファイルのコレクション
    ' ここで、FilesInFoldersは、新しく追加されたファイルのコレクションとなる

    Dim pathFile, ppath
    Dim currentLine, tmpSplitLine
    Dim firstChar

    Dim FS
    Set FS = CreateObject("Scripting.FileSystemObject")
    If FS.FileExists(OUTPUT_TEXT_NAME) = false Then
        exit Sub
    End If
    

    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile OUTPUT_TEXT_NAME
        Do Until .EOS
            ' adReadAll ' -1
            ' 既定値。現在の位置から EOS マーカー方向に、すべてのバイトをストリームから読み取ります。
            'これは、バイナリ ストリーム (Type は adTypeBinary) に唯一有効な StreamReadEnum 値です。
            ' adReadLine' -2
            ' ストリームから次の行を読み取ります (LineSeparator プロパティで指定)。
            currentLine = .ReadText(-2)
            firstChar = left(trim(currentLine), 1)
            If trim(currentLine) = "" or firstChar = ";" Then
                ExistFiles.add CreateObject("Scriptlet.TypeLib").GUID, currentLine
            Else
                tmpSplitLine = split((currentLine), "|")  ''''''''''''' 大文字小文字区別
                pathFile = trim(tmpSplitLine(0))

                If FilesInFolders.Exists(pathFile) Then
                    'FilesInFolders.Remove(pathFile)  'ここでは消さないで、最後にまとめて消す。*.txtにファイル名を重複させるため
                    If FilesExistsForDeleteAfter.Exists(pathFile) = false Then
                        FilesExistsForDeleteAfter.add pathFile, pathFile
                    End If
                    ExistFiles.add CreateObject("Scriptlet.TypeLib").GUID, currentLine
                Else
                    If DeleteFiles.Exists(currentLine) = False Then
                        DeleteFiles.add CreateObject("Scriptlet.TypeLib").GUID, ";削除済み:" & currentLine
                        ExistFiles.add CreateObject("Scriptlet.TypeLib").GUID, ";削除済み:" & currentLine
                    End If
                End If
            End If
        Loop
        .Close
    End with

    '存在していたファイルを、FilesInFoldersからまとめて削除する。
    Dim afile
    For each afile in FilesExistsForDeleteAfter
        FilesInFolders.Remove(afile)
    next

    ''FilesInFolders には追加されたファイルだけが残っている

End Sub

Sub AddNewFileToExist()
    Dim existItems
    existItems = ExistFiles.items
    Dim iExist
    Dim existItem
    Dim ppath
    Dim newFileObj
    For iExist = 0 To ExistFiles.Count - 1
        existItem = existItems(iExist)
        ppath = GetParentFolderName(existItem)
        set newFileObj = GetNewFilesAtFolder(ppath)
        debugPrint 1, "newFileObj", newFileObj.Count
    Next
End Sub

Function GetNewFilesAtFolder(folderName)
    debugPrint 1, "GetNewFilesAtFolder", folderName

    Dim AtFolders
    Set AtFolders = CreateObject("Scripting.Dictionary")

    Dim newItems
    newItems = FilesInFolders.items
    Dim iNew
    Dim newItem
    Dim ppath
    For iNew = 0 To FilesInFolders.Count - 1
        newItem = newItems(iNew)
        ppath = GetParentFolderName(newItem)
        If ppath = folderName Then
            AtFolders.Add newItem, newItem
            FilesInFolders.Remove newItem
        End If
    Next
    Set GetNewFilesAtFolder = AtFolders
End Function

Function isIgnoreDir(dirName)
    Dim s, e
    s = Lbound(ignoreFolders)
    e = Ubound(ignoreFolders)

    Dim i
    For i = s to e
        If lcase(ignoreFolders(i)) = lcase(dirName) Then
            isIgnoreDir = true
            Exit Function
        End If
    Next
    isIgnoreDir = False
    Exit Function
End Function

Sub debugPrint(flg, msg1, msg2)
    If flg = 0 Then
        Exit Sub
    End If
    call PrintCore(msg1, msg2)
End Sub

Sub InfoPrint(flg, msg1, msg2)
    If flg = 0 Then
        Exit Sub
    End If
    call PrintCore(msg1, msg2)
End Sub
Sub PrintCore(msg1, msg2)
    WScript.Echo msg1 & ": " & msg2
End Sub


Sub WriteToFile(filename, collection, headerMsg , isNew)
    'Dim writeStream As ADODB.Stream
    'Microsoft ActiveX Data Objects 2.5 Libraryと
    ' WriteText str, 1 => add a newline
    ' WriteText str, 0 => add no newline
    Dim writeStream

    ' 文字コードを指定してファイルをオープン
    Set writeStream = CreateObject("ADODB.Stream")
    writeStream.Charset = "UTF-8"
    writeStream.Open

    If isNew =false Then
        writeStream.LoadFromFile filename
        writeStream.Position = writeStream.Size
    End If

    If collection.Count > 0 and headerMsg <> "" Then
        writeStream.WriteText "", 1
        writeStream.WriteText  headerMsg, 1
        writeStream.WriteText "", 1
    End If

    '実際の中身の書き込み
    Dim i, items
    items = collection.items
    For i = 0 to collection.Count -1
    	debugPrint 0, headerMsg,  collection.Count
        writeStream.WriteText items(i), 1
    next

    If collection.Count > 0 and headerMsg <> ""Then
        writeStream.WriteText "",1
        writeStream.WriteText ";>>>>>>>>>>>>>>>>>>ここまで>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>",1
        writeStream.WriteText "",1
    End If

    ' ファイルに書き込み
    writeStream.SaveToFile filename, 2 'adSaveCreateOverWrite:2

    ' ファイルをクローズ
    writeStream.Close
    Set writeStream = Nothing
End Sub

Sub WriteToFileEx(filename, collection, headerMsg , isNew)
    'Dim writeStream As ADODB.Stream
    'Microsoft ActiveX Data Objects 2.5 Libraryと
    ' WriteText str, 1 => add a newline
    ' WriteText str, 0 => add no newline
    Dim writeStream

    ' 文字コードを指定してファイルをオープン
    Set writeStream = CreateObject("ADODB.Stream")
    writeStream.Charset = "UTF-8"
    writeStream.Open

    If isNew =false Then
        writeStream.LoadFromFile filename
        writeStream.Position = writeStream.Size
    End If

    If collection.Count > 0 and headerMsg <> "" Then
        writeStream.WriteText "", 1
        writeStream.WriteText  headerMsg, 1
        writeStream.WriteText "", 1
    End If

    '実際の中身の書き込み
    Dim i, items, folderName1, fileName1

    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")

    GetParentFolderName= FS.GetParentFolderName(fpath)
    Set FS = Nothing

    items = collection.items
    For i = 0 to collection.Count -1
    	debugPrint 0, headerMsg,  collection.Count

        folderName1 = FSO.GetParentFolderName(items(i))
        fileName1 = FSO.GetFileName((items(i)))
        writeStream.WriteText folderName & ": " & fileName , 1
    next

    If collection.Count > 0 and headerMsg <> ""Then
        writeStream.WriteText "",1
        writeStream.WriteText ";>>>>>>>>>>>>>>>>>>ここまで>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>",1
        writeStream.WriteText "",1
    End If

    ' ファイルに書き込み
    writeStream.SaveToFile filename, 2 'adSaveCreateOverWrite:2

    ' ファイルをクローズ
    writeStream.Close
    Set writeStream = Nothing
End Sub


''FilesInFolders, DeleteFiles, ExistFiles
Public Sub WriteToFileHTML(collection0, collection1, collection2)
'collection0 FilesInFolders(最初は、フォルダ内全ファイルだが、この時点では、新しく追加されたファイルのコレクションになっている。)
'collection1 DeleteFiles
'collection2 ExistFiles

    '## XML Object の作成
    Dim XMLobj
    Set XMLobj = CreateObject("MSXML2.DOMdocument")
    XMLobj.async = false

    '## ファイルの存在を確認し、存在する場合はXMLファイルを読み込む。存在しない場合はXMLオブジェクトの基本構造を作成する。
    XMLobj.loadXML("<!DOCTYPE HTML>")

    Dim HtmlElement
    Set HtmlElement = XMLobj.createElement("html")

    Dim HeaderElement
    Set HeaderElement = XMLobj.createElement("header")

    Dim charset
    Set charset = XMLobj.createElement("meta")
    Call charset.setAttribute("charset","utf-8")

    Dim styleElement
    Set styleElement = XMLobj.createElement("style")
    Call styleElement.setAttribute("type", "text/css")
    styleElement.appendChild(XMLobj.CreateTextNode("<!-- html, body {  background-color: #ccffff;  color: #000000;} "))
    styleElement.appendChild(XMLobj.CreateTextNode("html, body {  background-color: #ccffff;  color: #000000;}"))
    styleElement.appendChild(XMLobj.CreateTextNode("h1 {  padding-left: 0em;  background-color: #80ffff;  color: #000080;  font:bold;}"))
    styleElement.appendChild(XMLobj.CreateTextNode("h2 {  padding-left: 3em;  background-color: #80ffff;  color: #000080;  font:bold;}"))
    styleElement.appendChild(XMLobj.CreateTextNode("h3 {  padding-left: 6em;  background-color: #80ffff;  color: #000080;  font:bold;}"))
    styleElement.appendChild(XMLobj.CreateTextNode("h4 {  padding-left: 9em;  background-color: #80ffff;  color: #000080;  font:bold;}"))
    styleElement.appendChild(XMLobj.CreateTextNode("h5 {  padding-left: 12em;  background-color: #80ffff;  color: #000080;  font:bold;}"))
    styleElement.appendChild(XMLobj.CreateTextNode("h6 {  padding-left: 15em;  background-color: #80ffff;  color: #000080;  font:bold;}"))
    styleElement.appendChild(XMLobj.CreateTextNode("h7 {  padding-left: 18em;  background-color: #80ffff;  color: #000080;  font:bold;}"))
    styleElement.appendChild(XMLobj.CreateTextNode("h8 {  padding-left: 21em;  background-color: #80ffff;  color: #000080;  font:bold;}"))
    styleElement.appendChild(XMLobj.CreateTextNode("h9 {  padding-left: 24em;  background-color: #80ffff;  color: #000080;  font:bold;}"))
    styleElement.appendChild(XMLobj.CreateTextNode("h10 {  padding-left: 3em;  background-color: #80ffff;  color: #000080;  font:bold;}"))
    styleElement.appendChild(XMLobj.CreateTextNode(".spc {  text-align: right;}"))
    styleElement.appendChild(XMLobj.CreateTextNode("// -->"))

    HeaderElement.appendChild(charset)
    HeaderElement.appendChild(styleElement)

    Dim BodyElement
    Set BodyElement = XMLobj.createElement("body")

    Dim collections(2)
    Set collections(0) = collection0
    Set collections(1) = collection1
    Set collections(2) = collection2

    Dim i ,j, itemVal, splitVal, items, childNode, FilePath, FileComment, k
    Dim beforeParentPath
    beforeParentPath = "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"

    Dim beforIndentN
    beforIndentN = 0

    Dim CommentCollection
    Set CommentCollection = CreateObject("Scripting.Dictionary")

    For i = 0 to 2
        collections(i).add CreateObject("Scriptlet.TypeLib").GUID, " " ' からデータ追加、Itmeの次のインデックスを評価するため
        For j = 0 to collections(i).Count -1
            items = collections(i).items
            itemVal = items(j)
            If itemVal = "" Then
                itemVal = "|"
            End If


            splitVal = split(itemVal & "|" ,"|") '2個以上に分割されるようにする。
            debugPrint 0, "itemVal", splitVal(0)

            FilePath = splitVal(0)
            FileComment = splitVal(1)

            If left(trim(itemVal),1) = ";" Then
                '''フォルダタイトルをつけるときには、フォルダタイトルの後にコメントをつけるので、記憶しておく
                CommentCollection.add CreateObject("Scriptlet.TypeLib").GUID, itemVal
            Else
                ''フォルダタイトルの作成
                If (GetParentFolderName(FilePath) <> beforeParentPath) and (FilePath <> "") Then

                    Dim indentn
                    beforeParentPath = GetParentFolderName(FilePath)
                    indentn = len(beforeParentPath) - len(replace(beforeParentPath,"\",""))


                    Dim tagh
                    Set tagh = XMLobj.createElement("h" & (indentn + 1))


                    Set childNode = XMLobj.createElement("a")
                    Call childNode.setAttribute("href", "./" & beforeParentPath)
                    Call childNode.appendChild(XMLobj.CreateTextNode(beforeParentPath))

                    tagh.appendChild childNode
                    BodyElement.appendChild(tagh)

                End If
	            ''' コメント追加
	            Call OutPutCommentHtml(BodyElement, XMLobj, CommentCollection)
		        
		        
                Set childNode = XMLobj.createElement("a")
                Call childNode.setAttribute("href", FilePath)

                Dim linktitle
                linktitle = GetFileName(FilePath)

                If linktitle <> "." Then
                    'リンクタイトルが . の時は追加しない
                    If trim(FileComment) = "" Then
                        ' そのまま
                    Else
                        linktitle = linktitle &  "【" & FileComment & "】"
                    End If
                    
                    '' add files
                    childNode.appendChild(XMLobj.CreateTextNode(linktitle))
                    BodyElement.appendChild(childNode)
                    BodyElement.appendChild(XMLobj.createElement("br"))
                    BodyElement.appendChild(XMLobj.CreateTextNode(chr(10)))
                End If
            End If

            '/* p要素を作成し、テキストノードを追加 */
            'var tObj=document.createTextNode("引用元：");
            'var pObj=document.createElement("p");
            'pObj.appendChild(tObj);
            'pObj.appendChild(aObj);
        next

        ''' 残りのコメント追加
        Call OutPutCommentHtml(BodyElement, XMLobj, CommentCollection)
    next

    HtmlElement.appendChild(HeaderElement)
    HtmlElement.appendChild(BodyElement)
    XMLobj.appendChild(HtmlElement)
    XMLobj.save(OUTPUT_HTML_NAME)
End Sub


''FilesInFolders, DeleteFiles, ExistFiles
Public Sub WriteToMarkdown2(collection0, collection1, collection2)
    'collection0 FilesInFolders(最初は、フォルダ内全ファイルだが、この時点では、新しく追加されたファイルのコレクションになっている。)
    'collection1 DeleteFiles
    'collection2 ExistFiles
    
   

    
        Dim collections(2)
        Set collections(0) = collection0
        Set collections(1) = collection1
        Set collections(2) = collection2
    
        Dim i ,j, itemVal, splitVal, items, childNode, FilePath, FileComment, k
        Dim beforeParentPath
        beforeParentPath = "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
    
        Dim CommentCollection
        Set CommentCollection = CreateObject("Scripting.Dictionary")
    
        For i = 0 to 2
            collections(i).add CreateObject("Scriptlet.TypeLib").GUID, " " ' からデータ追加、Itmeの次のインデックスを評価するため
            For j = 0 to collections(i).Count -1
                items = collections(i).items
                itemVal = items(j)
                If itemVal = "" Then
                    itemVal = "|"
                End If
    
    
                splitVal = split(itemVal & "|" ,"|") '2個以上に分割されるようにする。
                debugPrint, "itemVal", splitVal(0)
    
                FilePath = splitVal(0)
                FileComment = splitVal(1)
    
                If left(trim(itemVal),1) = ";" Then
                    '''フォルダタイトルをつけるときには、フォルダタイトルの後にコメントをつけるので、記憶しておく
                    CommentCollection.add CreateObject("Scriptlet.TypeLib").GUID, itemVal
                Else
                    ''フォルダタイトルの作成
                    If (GetParentFolderName(FilePath) <> beforeParentPath) and (FilePath <> "") Then
    
                        Dim indentn
                        beforeParentPath = GetParentFolderName(FilePath)
                        indentn = len(beforeParentPath) - len(replace(beforeParentPath,"\",""))
    
                        Dim tagh
                        Set tagh = XMLobj.createElement("h" & (indentn + 1))
    
                        Set childNode = XMLobj.createElement("a")
                        Call childNode.setAttribute("href", "./" & beforeParentPath)
                        Call childNode.appendChild(XMLobj.CreateTextNode(beforeParentPath))
    
                        tagh.appendChild childNode
                        BodyElement.appendChild(tagh)
    
                    End If
                    ''' コメント追加
                    Call OutPutCommentHtml(BodyElement, XMLobj, CommentCollection)
                    
                    
                    Set childNode = XMLobj.createElement("a")
                    Call childNode.setAttribute("href", FilePath)
    
                    Dim linktitle
                    linktitle = GetFileName(FilePath)
                    If trim(FileComment) = "" Then
                        ' そのまま
                    Else
                        linktitle = linktitle &  "【" & FileComment & "】"
                    End If
    
                    childNode.appendChild(XMLobj.CreateTextNode(linktitle))
    
                    BodyElement.appendChild(childNode)
                    BodyElement.appendChild(XMLobj.createElement("br"))
                    BodyElement.appendChild(XMLobj.CreateTextNode(chr(10)))
                End If
    
                '/* p要素を作成し、テキストノードを追加 */
                'var tObj=document.createTextNode("引用元：");
                'var pObj=document.createElement("p");
                'pObj.appendChild(tObj);
                'pObj.appendChild(aObj);
            next
    
            ''' 残りのコメント追加
            Call OutPutCommentHtml(BodyElement, XMLobj, CommentCollection)
        next
    
        HtmlElement.appendChild(HeaderElement)
        HtmlElement.appendChild(BodyElement)
        XMLobj.appendChild(HtmlElement)
        XMLobj.save(OUTPUT_HTML_NAME)
End Sub


Sub WriteToMarkdown(filename, collection, headerMsg , isNew)
    'Dim writeStream As ADODB.Stream
    'Microsoft ActiveX Data Objects 2.5 Libraryと
    Dim writeStream

    ' 文字コードを指定してファイルをオープン
    Set writeStream = CreateObject("ADODB.Stream")
    writeStream.Charset = "UTF-8"
    writeStream.Open

    If isNew =false Then
        writeStream.LoadFromFile filename
        writeStream.Position = writeStream.Size
    End If

    If collection.Count > 0 and headerMsg <> ""Then
        writeStream.WriteText "",1
        writeStream.WriteText  headerMsg,1
        writeStream.WriteText "",1
    End If

    '実際の中身の書き込み
    Dim i, items
    items = collection.items
    For i = 0 to collection.Count -1
    	debugPrint 0, headerMsg, collection.Count
        writeStream.WriteText items(i) ,1
    next

    If collection.Count > 0 and headerMsg <> ""Then
        writeStream.WriteText "",1
        writeStream.WriteText ";>>>>>>>>>>>>>>>>>>ここまで>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>",1
        writeStream.WriteText "",1
    End If

    ' ファイルに書き込み
    writeStream.SaveToFile filename, 2 'adSaveCreateOverWrite:2

    ' ファイルをクローズ
    writeStream.Close
    Set writeStream = Nothing
End Sub


    
Public Sub OutPutCommentHtml(BodyElement, XMLobj, CommentCollection)
    Dim i, items
    items = CommentCollection.items
    For i = 0 to CommentCollection.Count - 1
        BodyElement.appendChild(XMLobj.CreateTextNode(items(i)))
        BodyElement.appendChild(XMLobj.createElement("br"))
        BodyElement.appendChild(XMLobj.CreateTextNode(chr(10)))
    next
    Set CommentCollection = CreateObject("Scripting.Dictionary")
End Sub


Public Function GetFileName(fpath)
    Dim FS
    Set FS = CreateObject("Scripting.FileSystemObject")
    GetFileName= FS.GetFileName(fpath)
    Set FS = Nothing
End Function

Public Function GetParentFolderName(fpath)
    Dim FS
    Set FS = CreateObject("Scripting.FileSystemObject")
    GetParentFolderName= FS.GetParentFolderName(fpath)
    Set FS = Nothing
End Function
