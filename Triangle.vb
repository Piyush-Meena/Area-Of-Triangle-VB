Private Sub cmdClick_Click()
Dim h As Integer, b As Integer
Dim area As Double
h = Val(Text1.Text)
b = Val(Text2.Text)
area = (0.5) * h * b
Label3.Caption = "Area of Triangle " & area
End Sub
