Option Explicit

Sub CSVから文字検索して行を抽出()

    Dim objCn As Object
    Dim objRS As Object
    Dim strSQL As String
    Dim csvPath As String
    Dim searchStr As String
    Dim searchStr2 As String
    Dim destSheet As Worksheet
    Dim lNextSheet As Long
    Dim wsCreation As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim columnName As String
    Dim columnName2 As String
    Dim columnName3 As String
    Dim startTime As Double
    Dim endTime As Double
    Dim elapsedMinutes As Integer
    Dim elapsedSeconds As Integer
    Dim fileName As String
    Dim byteCount As Integer

    On Error GoTo ErrorHandler
  
    
    ' 画面の更新をオフにして実行速度を向上させる
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' マクロの開始時間を取得
    startTime = Timer
 
     Dim obtainedFileName As String
    obtainedFileName = FindAndSetFilename()
    ' 「シート作成」からCSVファイルのパスと検索文字列を取得
    csvPath = ThisWorkbook.Worksheets("項目名参照").Range("C6") & "\" & obtainedFileName
    
    searchStr = ThisWorkbook.Worksheets("シート作成").Range("B16")
    searchStr2 = ThisWorkbook.Worksheets("シート作成").Range("B17")
    
    
    '「シート作成」でnull判断
If ThisWorkbook.Worksheets("シート作成").Range("B16") = "" And ThisWorkbook.Worksheets("シート作成").Range("B17") = "" Then
       ThisWorkbook.Worksheets("シート作成").Range("A16").Font.Color = RGB(255, 0, 0)
       ThisWorkbook.Worksheets("シート作成").Range("A17").Font.Color = RGB(255, 0, 0)
    
    MsgBox "検索するメールアドレスと検索するキーワードはいずれか入力してください。"
  Exit Sub
    
End If
Dim ws As Worksheet
Dim cell As Range
Dim columns() As String
Dim modifiedColumns As String
Dim k As Integer


Set ws = ThisWorkbook.Worksheets("項目名参照") ' 対象のシートを指定
lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row ' B列の最後の行番号を取得

ReDim columns(1 To lastRow - 7) ' B8から始まるので、7を引いて配列のサイズを定義

' B8からB62までのセルの値を配列に格納
'k = 1
'For Each cell In ws.Range("B8:B" & lastRow)
    'columns(k) = cell.Value
    'k = k + 1
'Next cell

' SQLクエリを動的に生成
'For k = LBound(columns) To UBound(columns)
    'modifiedColumns = modifiedColumns & "REPLACE(" & columns(k) & ", '　', ' ') AS " & columns(k)
    'If k < UBound(columns) Then modifiedColumns = modifiedColumns & ", "
'Next k


    ' CSVに接続する設定
    Set objCn = CreateObject("ADODB.Connection")
    Set objRS = CreateObject("ADODB.Recordset")

    objCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                Left(csvPath, InStrRev(csvPath, "\")) & _
                ";Extended Properties=""text; HDR=Yes; FMT=Delimited;"""
                
    
    ' 「シート作成」から列名を取得
    columnName = ThisWorkbook.Worksheets("項目名参照").Range("C1").Value
    columnName2 = ThisWorkbook.Worksheets("項目名参照").Range("C2").Value
    columnName3 = ThisWorkbook.Worksheets("項目名参照").Range("C3").Value
    
    ThisWorkbook.Worksheets("シート作成").Range("A16").Font.Color = RGB(0, 0, 0)
    ThisWorkbook.Worksheets("シート作成").Range("A17").Font.Color = RGB(0, 0, 0)
If ThisWorkbook.Worksheets("シート作成").Range("B16") = "" And ThisWorkbook.Worksheets("シート作成").Range("B17") <> "" Then
           'バイト数の計算
    byteCount = LenB(searchStr2)
    
    'バイト数チェック
    If byteCount < 3 Then
        MsgBox "2文字以上入力してください。"
        Exit Sub
    End If

    MsgBox "キーワード検索は時間をかかります。"
    If Not SheetExists("検索結果") Then
        Set destSheet = ThisWorkbook.Sheets.Add
        destSheet.Name = "検索結果"
    Else
        lNextSheet = 1
        Do While SheetExists("検索結果" & lNextSheet)
            lNextSheet = lNextSheet + 1
        Loop
        Set destSheet = ThisWorkbook.Sheets.Add
        destSheet.Name = "検索結果" & lNextSheet
    End If

        Set wsCreation = ThisWorkbook.Worksheets("項目名参照")
        lastRow = wsCreation.Cells(Rows.Count, 2).End(xlUp).Row
        Set dataRange = wsCreation.Range(wsCreation.Cells(8, 2), wsCreation.Cells(lastRow, 2))
        dataRange.Copy
        destSheet.Range("A1").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
                
        ' フォーマットの設定
        destSheet.Rows(1).AutoFilter
        destSheet.Cells.EntireRow.AutoFit
        destSheet.Cells.EntireColumn.AutoFit
        destSheet.Rows(1).Interior.Color = RGB(255, 255, 0)
        destSheet.UsedRange.Borders.LineStyle = xlContinuous

    Dim i As Integer
    ' 「項目名参照」シートのD8からD35までループ
    For i = 8 To 35
        Dim currentFile As String
        currentFile = ThisWorkbook.Worksheets("項目名参照").Range("C6") & "\" & ThisWorkbook.Worksheets("項目名参照").Cells(i, 4).Value ' D列の値を取得
        
        If currentFile <> "" Then
            ' SQLクエリの作成
        strSQL = "SELECT " & _
                 "REPLACE(INQUERYCONTENTS__C, '　', ' ')," & _
                 "REPLACE(DESCRIPTION, '　', ' ')," & _
                 "REPLACE(FORMNAME__C, '　', ' ')," & _
                 "*" & _
                 "FROM [" & Dir(currentFile) & "] WHERE " & _
                     " [" & columnName2 & "] LIKE '%" & searchStr2 & "%'"
            
            ' SQLクエリで文字列の検索
            objRS.Open strSQL, objCn
        Dim j As Integer
            
        ' 「検索結果」シートの最後の空白行を見つける
        Dim nextEmptyRow As Long
        nextEmptyRow = destSheet.Cells(destSheet.Rows.Count, 5).End(xlUp).Row + 1


            ' 最後の空白行にデータをコピー
        destSheet.Cells(nextEmptyRow, 1).CopyFromRecordset objRS


            ' 接続を閉じる
            objRS.Close
        End If
    Next i
        destSheet.columns("BB").Delete Shift:=xlToLeft
        destSheet.columns("R").Delete Shift:=xlToLeft
        destSheet.columns("AX").Delete Shift:=xlToLeft

     If WorksheetFunction.Subtotal(3, Range("E:E")) > 1 Then
        MsgBox "新規シート「" + destSheet.Name + "」を確認してください。"
   
    
        ' マクロの終了時間を取得して、経過時間を計算
    endTime = Timer
    elapsedMinutes = Int((endTime - startTime) / 60)
    elapsedSeconds = (endTime - startTime) Mod 60

    ' 経過時間を「シート作成」のD5に表示
    ThisWorkbook.Worksheets("シート作成").Range("B1").Value = elapsedMinutes & "分" & elapsedSeconds & "秒"
    ThisWorkbook.Worksheets("シート作成").Range("B2").Value = WorksheetFunction.Subtotal(3, Range("E:E")) - 1
    ' 設定を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ThisWorkbook.Worksheets("シート作成").Range("B2").Value = WorksheetFunction.Subtotal(3, Range("E:E")) - 1
     Exit Sub
   Else
           ' マクロの終了時間を取得して、経過時間を計算
    endTime = Timer
    elapsedMinutes = Int((endTime - startTime) / 60)
    elapsedSeconds = (endTime - startTime) Mod 60

    ' 経過時間を「シート作成」のD5に表示
    ThisWorkbook.Worksheets("シート作成").Range("B1").Value = elapsedMinutes & "分" & elapsedSeconds & "秒"
    ThisWorkbook.Worksheets("シート作成").Range("B2").Value = WorksheetFunction.Subtotal(3, Range("E:E")) - 1

    ' 設定を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

        MsgBox "一致するレコードが見つかりませんでした。"
        Exit Sub
    End If
    Exit Sub
End If



If ThisWorkbook.Worksheets("シート作成").Range("B16") <> "" And ThisWorkbook.Worksheets("シート作成").Range("B17") <> "" Then
   
    ' SQLクエリの作成
    strSQL = "SELECT " & _
           "REPLACE(INQUERYCONTENTS__C, '　', ' ')," & _
           "REPLACE(DESCRIPTION, '　', ' ')," & _
           "REPLACE(FORMNAME__C, '　', ' ')," & _
           "*" & _
           " FROM [" & Dir(csvPath) & "] WHERE " & _
             "([" & columnName & "] LIKE '%" & searchStr & "%'or [" & columnName3 & "] LIKE '%" & searchStr & "%')and [" & columnName2 & "] LIKE '%" & searchStr2 & "%'"
End If

If ThisWorkbook.Worksheets("シート作成").Range("B16") <> "" And ThisWorkbook.Worksheets("シート作成").Range("B17") = "" Then
    
    ' SQLクエリの作成
    strSQL = "SELECT " & _
    "REPLACE(INQUERYCONTENTS__C, '　', ' ')," & _
    "REPLACE(DESCRIPTION, '　', ' ')," & _
    "REPLACE(FORMNAME__C, '　', ' ')," & _
    "*" & _
    " FROM [" & Dir(csvPath) & "] WHERE " & _
             "[" & columnName & "] LIKE '%" & searchStr & "%'or [" & columnName3 & "] LIKE '%" & searchStr & "%'"
End If


    ' SQLクエリで文字列の検索
    objRS.Open strSQL, objCn

    ' 「検索結果」シートが存在するかチェックして新しいシートを作成
    If Not SheetExists("検索結果") Then
        Set destSheet = ThisWorkbook.Sheets.Add
        destSheet.Name = "検索結果"
    Else
        lNextSheet = 1
        Do While SheetExists("検索結果" & lNextSheet)
            lNextSheet = lNextSheet + 1
        Loop
        Set destSheet = ThisWorkbook.Sheets.Add
        destSheet.Name = "検索結果" & lNextSheet
    End If

    ' 「シート作成」が存在する場合、データをコピー
    If Not objRS.EOF And SheetExists("シート作成") Then
        Set wsCreation = ThisWorkbook.Worksheets("項目名参照")
        lastRow = wsCreation.Cells(Rows.Count, 2).End(xlUp).Row
        Set dataRange = wsCreation.Range(wsCreation.Cells(8, 2), wsCreation.Cells(lastRow, 2))
        dataRange.Copy
        destSheet.Range("A1").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        destSheet.Range("A2").CopyFromRecordset objRS
        destSheet.columns("BB").Delete Shift:=xlToLeft
        destSheet.columns("R").Delete Shift:=xlToLeft
        destSheet.columns("AX").Delete Shift:=xlToLeft
        
        ' フォーマットの設定
        destSheet.Rows(1).AutoFilter
        destSheet.Cells.EntireRow.AutoFit
        destSheet.Cells.EntireColumn.AutoFit
        destSheet.Rows(1).Interior.Color = RGB(255, 255, 0)
        destSheet.UsedRange.Borders.LineStyle = xlContinuous

        MsgBox "新規シート「" + destSheet.Name + "」を確認してください。"

    Else
        MsgBox "一致するレコードが見つかりませんでした。"
    End If

    ' 接続を閉じる
    objCn.Close
    Set objRS = Nothing
    Set objCn = Nothing

    ' マクロの終了時間を取得して、経過時間を計算
    endTime = Timer
    elapsedMinutes = Int((endTime - startTime) / 60)
    elapsedSeconds = (endTime - startTime) Mod 60

    ' 経過時間を「シート作成」のD5に表示
    ThisWorkbook.Worksheets("シート作成").Range("B1").Value = elapsedMinutes & "分" & elapsedSeconds & "秒"
    ThisWorkbook.Worksheets("シート作成").Range("B2").Value = WorksheetFunction.Subtotal(3, Range("E:E")) - 1

    ' 設定を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    On Error GoTo ErrorHandler
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description


End Sub

' シートが存在するか確認する関数
Function SheetExists(sheetName As String) As Boolean
    Dim sht As Worksheet
    On Error Resume Next
    Set sht = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If Not sht Is Nothing Then SheetExists = True
End Function




Function FindAndSetFilename() As String

    Dim searchStr As String
    searchStr = ThisWorkbook.Worksheets("シート作成").Range("B16").Value
    
    Dim firstChar As String
    Dim cellValue As String
    Dim foundMatch As Boolean
    Dim cell As Range
    Dim firstCharNormalized As String
    
    ' B16セルの先頭の文字を取得
    firstChar = Left(searchStr, 1)

    ' 全角文字や数字を半角に変換し、大文字を小文字に変換して正規化
    firstCharNormalized = LCase(Application.WorksheetFunction.Text(firstChar, "＠"))

    ' 先頭の文字が数字かどうかを判定（全角数字も考慮）
    If IsNumeric(firstCharNormalized) Or IsNumeric(Application.WorksheetFunction.Text(firstChar, "＠")) Then
        FindAndSetFilename = "0-9.csv"
    Else
        foundMatch = False
        For Each cell In ThisWorkbook.Worksheets("項目名参照").Range("D8:D35")
            cellValue = CStr(cell.Value)
            
            If LCase(Application.WorksheetFunction.Text(Left(cellValue, 1), "＠")) = firstCharNormalized Then
                FindAndSetFilename = cellValue
                foundMatch = True
                Exit For
            End If
        Next cell
        
        If Not foundMatch Then
            For Each cell In ThisWorkbook.Worksheets("項目名参照").Range("D8:D35")
                If cell.Value = "文字と数字以外.csv" Then
                    FindAndSetFilename = cellValue
                    foundMatch = True
                    Exit For
                End If
            Next cell
        End If

        If Not foundMatch Then
            MsgBox "ファイルを見つけませんでした"
        End If
    End If

End Function


# VBA
![image](https://github.com/lianghunan17/VBA/assets/50505315/98a84b6e-915c-4fde-835c-7f254eb05587)
![image](https://github.com/lianghunan17/VBA/assets/50505315/189bd7a6-a7ce-443f-926d-a8da5b7e6536)

