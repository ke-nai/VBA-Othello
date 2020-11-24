# VBA-Othello
![Othello](https://user-images.githubusercontent.com/66747535/100058445-9ec52d00-2e6c-11eb-9166-b764d9970310.gif)
엑셀에서 VBA 매크로를 통해 실행할 수 있는 오델로(오셀로) 게임이다.

## 적용법
1. VBA 편집창에 들어간다.
2. 모듈이 아니라 적용할 시트의 코드 창에 아래의 코드를 모두 넣는다.
3. 매크로 직접 실행으로 Format 실행

## 코드
<details>
    <summary>코드보기</summary>

```
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Application.ScreenUpdating = False
    
    Cells(1, 1) = "=rows(" + Selection.Address + ")"
    m = Cells(1, 1) '행 크기
    Cells(1, 1) = "=columns(" + Selection.Address + ")"
    n = Cells(1, 1) '열 크기
    
    r = Selection.Row
    c = Selection.Column
    
    If m > 1 Or m > 1 Then '다중 선택 방지
    ElseIf Selection.Address = Cells(2, 11).Address Then '다시 시작
        Call Start
    ElseIf Selection.Address = Cells(4, 11).Address Then
        Cells(r, c) = 1 + Cells(4, 11) Mod 2
    ElseIf 1 < r And r < 10 And 1 < c And c < 10 And Selection = 0 Then '범위 내
        t = 0
        For i = 1 To 9
            If i <> 5 Then
                rr = Range(Cells(r - 1, c - 1), Cells(r + 1, c + 1)).Cells(i).Row
                cc = Range(Cells(r - 1, c - 1), Cells(r + 1, c + 1)).Cells(i).Column
                
                If check(i, rr, cc) = 2 Then '체크, 뒤집었음?
                    t = 1 '뒤집었음
                End If
            End If
        Next
        If t = 1 Then '뒤집은 거 있음(돌 놓을 수 있음)
            m = Cells(4, 11)
            Cells(r, c) = m '돌 놓기
            Cells(r, c).Font.Color = RGB(255 * (m - 1), 255 * (m - 1), 255 * (m - 1))
            Cells(4, 11) = 1 + m Mod 2 '턴 변경
        End If
    End If
    Cells(12, 1).Select
    
    Application.ScreenUpdating = True
End Sub

Function check(i, r, c) '방향, row, column
    m = Cells(4, 11) '내 턴 1흑 2백
    y = 1 + m Mod 2 '상대의 숫자
    
    '내 돌이나 빈칸일 때까지 찾아가서 확인
    '내 돌로 끝나면 중간의 상대 돌을 내 돌로 변경
    
    If 1 < r And r < 10 And 1 < c And c < 10 Then
        If Cells(r, c) = y Then '상대 돌
            rr = Range(Cells(r - 1, c - 1), Cells(r + 1, c + 1)).Cells(i).Row '같은 방향 확인
            cc = Range(Cells(r - 1, c - 1), Cells(r + 1, c + 1)).Cells(i).Column
            tmp = check(i, rr, cc) '다음 칸 확인
            If tmp = 0 Then '다음 칸이 빈칸
                check = 0 '이전 칸에 빈 칸임을 전달
            ElseIf tmp > 0 Then '다음 칸이 내 돌이거나 뒤집으라고 전달받음
                Cells(r, c) = m '뒤집기
                Cells(r, c).Font.Color = RGB(255 * (m - 1), 255 * (m - 1), 255 * (m - 1))
                check = 2 '뒤집으라고 전달
            End If
        ElseIf Cells(r, c) = m Then '내 돌
            check = 1 '내 돌이라고 전달
        Else '빈 칸
        check = 0 '빈칸이라고  전달
        End If
    Else '범위 밖
        check = 0 '빈칸이라고 전달
    End If
End Function

Sub Format()
    Application.ScreenUpdating = False

    Range("A1:XFD1048576").EntireRow.Clear
    Range("A1:XFD1048576").EntireColumn.Clear
    Range("A1:XFD1048576").EntireRow.Hidden = False
    Range("A1:XFD1048576").EntireColumn.Hidden = False
    Range("M11:XFD1048576").EntireRow.Hidden = True
    Range("M11:XFD1048576").EntireColumn.Hidden = True

    With Range(Cells(1, 1), Cells(10, 12))
        .ColumnWidth = 8
        .RowHeight = 55
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(0, 0, 0)
    End With
    
    Cells(1, 11).ColumnWidth = 20
    With Range(Cells(2, 2), Cells(9, 9))
        .Interior.Color = RGB(255, 189, 101)
        .Font.Size = 36
        .NumberFormatLocal = "●"
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With
    With Cells(2, 11)
        .Value = "다시 시작"
        .Interior.Color = RGB(255, 255, 255)
        .Font.Size = 15
        .Font.Bold = True
    End With
    With Cells(4, 11)
        .NumberFormatLocal = "[=1]차례 : 흑;[=2]차례 : 백;"
        .Interior.Color = RGB(255, 255, 255)
        .Font.Size = 15
        .Font.Bold = True
    End With
    With Cells(6, 11)
        .Value = "=""흑: ""&COUNTIF(B2:I9,1)&"", 백: ""&COUNTIF(B2:I9,2)"
        .Interior.Color = RGB(255, 255, 255)
        .Font.Size = 15
    End With
    
    Call Start

    Application.ScreenUpdating = True
End Sub

Function Start()
    With Range(Cells(2, 2), Cells(9, 9))
        .Value = ""
    End With
    
    Range("E5,F6").Value = 1
    Range("E5,F6").Font.Color = RGB(0, 0, 0)
    Range("E6,F5").Value = 2
    Range("E6,F5").Font.Color = RGB(255, 255, 255)
    Cells(4, 11) = 1 '흑부터
End Function
```
</details>
