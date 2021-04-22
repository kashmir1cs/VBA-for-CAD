Public excel As Object
Sub test()
 Dim excel As Object
 Set excel = GetObject(, "Excel.Application")
 MsgBox excel.ActiveSheet.Name
End Sub


Sub Macro1()
  Dim acad As Object      'AutoCAD개체를 넣어둘 변수
  Dim c As Object         'circle object를 넣어둘 변수
  Dim center(2) As Double '원의 중심점을 넣어둘 배열
  
  center(0) = 200  'x좌표
  center(1) = 200  'y좌표
  center(2) = 200  'z좌표
  
  Set acad = GetObject(, "AutoCAD.application")
  Set c = acad.ActiveDocument.ModelSpace.AddCircle(center, 2000)

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
        text = mSpaceObj.TextString
        If InStr(LCase(text), LCase(findingText)) Then '대소문자 구분하지 않음!
            result.Add text
        End If
    Next
    
    Set dwgTextFind = result
End Function



