Option Explicit

Function loadFileToList(filePath)
'지정된 텍스트 파일을 1줄씩 읽어 배열로 반환한다.
    Dim line As String
    Dim fileNum As Integer
    Dim element As Variant
    Dim list() As String '동적 배열 선언
    Dim i As Long: i = 0
    
    fileNum = FreeFile()
    Open filePath For Input Access Read As #fileNum
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        If line <> "" Then
            ReDim Preserve list(i)
            list(i) = line
            i = i + 1
        End If
    Loop
    
    Close #fileNum
    loadFileToList = list
End Function

Function allDwgTextToList()
'DWG 파일 안의 모든 텍스트 데이터를 배열에 저장한다.
    Dim mSpaceObj As AcadObject
    Dim count As Integer: count = ThisDrawing.ModelSpace.count
    Dim i As Long
    Dim j As Long: j = 0
    Dim text As String
    Dim textList() As String '동적 배열 선언
    
    For i = 0 To count - 1
        Set mSpaceObj = ThisDrawing.ModelSpace.Item(i)
        On Error Resume Next '텍스트 속성이 없으면 그냥 무시
        text = mSpaceObj.TextString
        If text <> "" Then
            ReDim Preserve textList(j)
            textList(j) = text
            j = j + 1
        End If
    Next

    allDwgTextToList = textList
End Function

Sub listToFile(textList, filePath)
'배열을 받아, 그 배열의 내용을 지정된 파일에 기록한다.
    Dim fileNum As Integer
    Dim element As Variant
    
    fileNum = FreeFile()
    Open filePath For Append As #fileNum
    
    For Each element In textList
        Print #fileNum, Chr(9); element
    Next
    
    Write #fileNum, '빈 줄 추가
    Close #fileNum
    
End Sub

Sub main()
    '여기서 각 필요 파일 이름을 수동으로 지정
    Dim dwgFileListPath As String: dwgFileListPath = "Z:\WORK\DWGlist.txt"
    Dim resultFilePath As String: resultFilePath = "Z:\WORK\DWGtext.txt"
    
    Dim fileList() As String
    Dim element As Variant
    Dim fileNum As Integer
    Dim ReadOnly As Boolean
    
    '파일 목록 불러오기
    fileList = loadFileToList(dwgFileListPath)
    
    For Each element In fileList
        On Error Resume Next 'DWG 파일 열다가 에러가 나도 그냥 다음 파일로 넘어가기
        AutoCAD.Documents.Open element, ReadOnly = True
        
        '텍스트 내용이 있는 파일 이름 적어주기
        fileNum = FreeFile()
        Open resultFilePath For Append As #fileNum
        Write #fileNum, element
        Close #fileNum
        
        '파일을 실제로 열어, 그 중 텍스트 부분을 추출한다
        listToFile allDwgTextToList, resultFilePath
        
        AutoCAD.Documents.Close
    Next
End Sub
AutoCAD에서 VBA를 이용해서 여러 DWG 파일에서 한번에 텍스트 내용 추출하기 - 인터넷 / 소프트웨어 - 기글하드웨어 : https://gigglehd.com/gg/?mid=soft&document_srl=379801
