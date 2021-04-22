Option Explicit

Sub dwgTextChange()
'DWG 파일 내의 모든 텍스트에서 원하는 부분이 있는지 찾아, _
'조건이 맞는 모든 텍스트를 모은 컬렉션을 반환한다.
    Dim mSpaceObj As AcadObject
    Dim item As Variant
    Dim text As String
    Dim findtext As String
        findtext = "100A"
        
    Dim result As New Collection '한번에 선언 및 생성 - 초기화
    
    For Each item In ThisDrawing.ModelSpace
        On Error Resume Next '텍스트 스트링이 없으면 그냥 지나감
        text = "" '초기화
        Set mSpaceObj = item
        text = mSpaceObj.TextString
        If text = "100A" Then '대소문자 구분하지 않음!
            text.TextString = "OD110"
        End If
    Next
    
    
End Sub

