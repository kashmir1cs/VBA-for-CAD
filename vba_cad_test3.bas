Option Explicit


Public dwgList As Collection

Sub main()
'변수 설정
    Dim dwgName As Variant
    Dim i As Long: i = 0
    Dim readOnly As Boolean: readOnly = False
    Dim filesystem As Object
    Dim hostFolder As String
    Dim outputFolder As String
    Dim prefix As String
    prefix = "변환"
    Set dwgList = New Collection
    
    hostFolder = "C:\Users\User\Desktop\IC"
    outputFolder = "C:\"

'실행창에 진행사항 표시
    
    Debug.Print "ND → OD Size 변환 작업 시작"

'지정폴더 File 연다.
    Set filesystem = CreateObject("scripting.filesystemobject")
    doFolder filesystem.getfolder(hostFolder)
    
'2단계 시작 표시
    Debug.Print ""
    Debug.Print "변환작업 시작"
    
For Each dwgName In dwgList
    
    i = i + 1
    Debug.Print i & "/" & dwgList.Count
    
    On Error Resume Next
    'CAD도면 Open
    AutoCAD.Documents.Open dwgName
    '작성한 Procedure 호출
    Call text_change
    Call text_change2
    Call text_delete("delete text")
    Call text_replace("Old Text", "New Text")

    AutoCAD.Documents.Save
    Debug.Print "변환완료"
    AutoCAD.Documents.Close
    
Next

End Sub
Sub doFolder(folder)
'각 디렉터리와 서브 디렉터리를 재귀적으로 순회하기 위한 서브루틴.
    On Error Resume Next '오류 발생시 다음 폴더로 넘어감.
    
    Dim subFolder As Variant
    For Each subFolder In folder.subfolders
        doFolder subFolder
    Next
    
' 순회 과정에서 DWG 파일이 발견되면 경로 컬렉션에 경로를 추가한다.
    Dim file As Variant
    For Each file In folder.Files
    'Like 연산자를 사용하여 dwg가 들어있는 파일은 컬렉션에 추가함
        If LCase(file.Name) Like "*.dwg" Then
            dwgList.Add file.Path
            Debug.Print file.Path '파일 경로를 debug.print 함수를 이용해 화면에 출력
        End If
    Next
End Sub



Sub text()
    Dim mSpaceObj As AcadObject
    Dim item As Variant
    Dim text As String
    Dim result As New Collection '한번에 선언 및 생성 - 초기화
    For Each item In ThisDrawing.ModelSpace
        On Error Resume Next 
        'text 초기화
        text = ""
        Set mSpaceObj = item
        'text에서 find text를 찾으면 컬렉션 개체에 추가
        text = mSpaceObj.textString
        If InStr(text, "find text") Then
            result.Add text
        End If
    Next
    Dim i As Integer
     
    For i = 1 To result.Count
     Debug.Print i
     Debug.Print Chr(10) 'enter키 입력과 동일 
     Debug.Print result(i)
     Debug.Print Chr(10)
    Next
    
    
    
End Sub


Sub text_change()
    'old text -> new text로 변경하는 procedure
    Dim mSpaceObj As AcadObject
    Dim item As Variant
    Dim text As String
    Dim result As New Collection '한번에 선언 및 생성 - 초기화
    For Each item In ThisDrawing.ModelSpace
        '반복문 시작, item의 textString 속성 읽은 후 text 변수에 저장 
        On Error Resume Next '에러가 발생해도 다음으로 진행
        text = "" 'text 변수는 초기화 
        Set mSpaceObj = item
        text = mSpaceObj.textString
        ' if elseif 반복을 통해 필요한 텍스트로 변경
        If InStr(text, "old text1") Then
            
            item.textString = Replace(text, "old text1", "new text1")
            
        ElseIf InStr(text, "old text2") Then
            item.textString = Replace(text, "old text2", "new text2")
             
        End If


    Next
 
Sub text_change2()

Debug.Print "진행 Procedure 설명"
    Dim mSpaceObj As AcadObject
    Dim item As Variant
    Dim text As String
    Dim result As New Collection '한번에 선언 및 생성 - 초기화
    For Each item In ThisDrawing.ModelSpace
        On Error Resume Next
        text = ""
        Set mSpaceObj = item 'Modelspace상의 객체 item을 mSpaceObj에 할당
        
        text = mSpaceObj.textString 'mSpaceObj의 textString을 text에 할당
        If text = "old text1" Then
            
            item.textString = Replace(text, "old text1", "new text2" & Chr(34)) '특수 문자 ascii code값 이용하여 출력
        
        ElseIf InStr(text, "old text2") Then
            
            item.textString = Replace(text, "old text2", "new text2" & Chr(34))
            
        End If


    Next
 
End Sub

Sub text_delete(strText As String)
'Text 삭제
    Debug.Print strText & "삭제"
    Dim mSpaceObj As AcadObject
    Dim item As Variant
    Dim text As String
    Dim result As New Collection '한번에 선언 및 생성 - 초기화
    For Each item In ThisDrawing.ModelSpace
        On Error Resume Next
        text = ""
        Set mSpaceObj = item
            
        text = mSpaceObj.textString
        If text = strText Then
            
            item.textString = "" 'textString 삭제
            
       
            
        End If


    Next
 
End Sub

Sub text_replace(strSearch As String, strChange As String)
'Text 삭제 혹은 변경하는 함수
Debug.Print strSearch & "→" & strChange
    Dim mSpaceObj As AcadObject
    Dim item As Variant
    Dim text As String
    Dim result As New Collection '한번에 선언 및 생성 - 초기화
    For Each item In ThisDrawing.ModelSpace
        On Error Resume Next
        text = ""
        Set mSpaceObj = item
        text = mSpaceObj.textString
        
        If InStr(text, strSearch) Then
            
            item.textString = Replace(text, strSearch, strChange)
       
            
        End If


    Next
 
End Sub


Function dwgTextFind(findingText) As Collection
'DWG 파일 내의 모든 텍스트에서 원하는 부분이 있는지 찾아, _
'조건이 맞는 모든 텍스트를 모은 컬렉션을 반환한다.
    Dim mSpaceObj As AcadObject
    Dim item As Variant
    Dim text As String
    Dim result As New Collection '한번에 선언 및 생성 - 초기화
    
    For Each item In ThisDrawing.ModelSpace
        On Error Resume Next '텍스트 스트링이 없으면 그냥 지나감
        text = "" '초기화
        Set mSpaceObj = item
        text = mSpaceObj.textString
        If InStr(LCase(text), LCase(findingText)) Then '대소문자 구분하지 않음!
            result.Add text
        End If
    Next
    
    Set dwgTextFind = result
End Function

