Sub text_point이동()
'text 이동하는 프로시져
  Dim mSpaceObj As AcadObject
  Dim item As AcadEntity 'CAD 개체 선언
  Dim text As String
  Dim result As New Collection '한번에 선언 및 생성 - 초기화
  Dim pt As Variant '좌표 입력 받을 개체 선언
  Dim old_pt(0 To 2) As Double 'x,y,z 좌표 입력할 수 있도록 배열 선언
  Dim new_pt(0 To 2) As Double 'x,y,z 좌표 입력할 수 있도록 배열 선언
  Dim cadText As AcadText

  For Each item In ThisDrawing.ModelSpace
    text = "" 'text 변수는 초기화
    Set mSpaceObj = item
    If TypeOf item Is AcadText Then 'CAD Entity가 AcadText Class인 경우
        If item.TextString = "__ER-000" Then
            pt = item.InsertionPoint
            text = mSpaceObj.TextString
            '각 배열 원소에 x,y,z 값 입력
            old_pt(0) = pt(0)
            old_pt(1) = pt(1)
            old_pt(2) = pt(2)
            new_pt(0) = pt(0) + 7.4
            new_pt(1) = pt(1)
            new_pt(2) = pt(2)
            item.Height = 4
            item.ScaleFactor = 0.8 '폭비율 조정하기
            item.Move old_pt, new_pt
            Debug.Print text + "__ER-000 작업 완료"
        End If
    End If
    Next
End Sub
