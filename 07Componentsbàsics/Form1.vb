Public Class Form1
    Dim PI As Double = Math.PI
    Private Sub desactivarPaneles()

        'Paneles Desactivados
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
        Panel6.Visible = False
        Panel7.Visible = False
        Panel8.Visible = False
        Panel9.Visible = False
        Panel10.Visible = False
        Panel11.Visible = False
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = False
    End Sub
    Private Sub TreeView1_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        'Aqui activaremos y desactivaremos los paneles
        Select Case e.Node.Tag
            Case 1
                desactivarPaneles()
                Panel1.Visible = True
                Panel1.Dock = DockStyle.Fill
            Case 2
                desactivarPaneles()
                Panel2.Visible = True
                Panel2.Dock = DockStyle.Fill
            Case 3
                desactivarPaneles()
                Panel3.Visible = True
                Panel3.Dock = DockStyle.Fill
            Case 4
                desactivarPaneles()
                Panel4.Visible = True
                Panel4.Dock = DockStyle.Fill
            Case 5
                desactivarPaneles()
                Panel5.Visible = True
                Panel5.Dock = DockStyle.Fill
            Case 6
                desactivarPaneles()
                Panel6.Visible = True
                Panel6.Dock = DockStyle.Fill
            Case 7
                desactivarPaneles()
                Panel7.Visible = True
                Panel7.Dock = DockStyle.Fill
            Case 8
                desactivarPaneles()
                Panel8.Visible = True
                Panel8.Dock = DockStyle.Fill
            Case 9
                desactivarPaneles()
                Panel9.Visible = True
                Panel9.Dock = DockStyle.Fill
            Case 10
                desactivarPaneles()
                Panel10.Visible = True
                Panel10.Dock = DockStyle.Fill
            Case 11
                desactivarPaneles()
                Panel11.Visible = True
                Panel11.Dock = DockStyle.Fill
            Case 12
                desactivarPaneles()
                Panel12.Visible = True
                Panel12.Dock = DockStyle.Fill
            Case 13
                desactivarPaneles()
                Panel13.Visible = True
                Panel13.Dock = DockStyle.Fill
            Case 14
                desactivarPaneles()
                Panel14.Visible = True
                Panel14.Dock = DockStyle.Fill
        End Select
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'AREA DEL TRIANGULO
        Dim Base As Double
        Dim Altura As Double
        Dim Res As Double
        Try
            Base = CDbl(TextBox1.Text)
            Altura = CDbl(TextBox2.Text)
            Res = Base * Altura / 2
            MsgBox("Area del triangulo : " + CStr(Res))
        Catch ex As Exception
            MsgBox("Introduce números.")
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'AREA DEL TRIANGULO
        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'AREA DEL TRIANGULO
        Panel1.Visible = False
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'AREA DEL RECTANGULO
        Dim Radio As Double
        Dim Res As Double
        Try
            Radio = CDbl(TextBox4.Text)
            Res = PI * (Radio * Radio)
            MsgBox("Area del rectangulo : " + CStr(Res))
        Catch ex As Exception
            MsgBox("Introduce un número")
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'AREA DEL RECTANGULO
        TextBox4.Text = ""
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'AREA DEL RECTANGULO
        Panel2.Visible = False
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TreeView1.ExpandAll()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        'LONGITUT DE LA CIRCUMFERÊNCIA
        Dim Radio As Double
        Dim Res As Double
        Try
            Radio = CDbl(TextBox5.Text)
            Res = 2 * PI * Radio
            MsgBox("La longitud de la circunferencia : " + CStr(Res))
        Catch ex As Exception
            MsgBox("Introduce un número")
        End Try

    End Sub
    'LONGITUT DE LA CIRCUMFERÊNCIA
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        TextBox5.Text = ""
    End Sub
    'LONGITUT DE LA CIRCUMFERÊNCIA
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Panel3.Visible = False
    End Sub
    'AREA DEL RECTNGLE
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim Base As Double
        Dim Altura As Double
        Dim Res As Double

        Try
            Base = CDbl(TextBox6.Text)
            Altura = CDbl(TextBox3.Text)
            Res = Base * Altura
            MsgBox("Area del rectangle: " + CStr(Res))
        Catch ex As Exception
            MsgBox("Introduce un número")
        End Try
    End Sub
    'AREA DEL RECTNGLE
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TextBox6.Text = ""
        TextBox3.Text = ""
    End Sub
    'AREA DEL RECTNGLE
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Panel4.Visible = False
    End Sub
    'SUMA DE DOS NUMEROS
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Dim valor1 As Double
        Dim valor2 As Double
        Dim Res As Double
        Try
            valor1 = CDbl(TextBox8.Text)
            valor2 = CDbl(TextBox7.Text)
            Res = valor1 + valor2
            MsgBox("Resultado: " + CStr(Res))
        Catch ex As Exception
            MsgBox("Introduce números")
        End Try
    End Sub
    'SUMA DE DOS NUMEROS
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        TextBox8.Text = ""
        TextBox7.Text = ""
    End Sub
    'SUMA DE DOS NUMEROS
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Panel5.Visible = False
    End Sub
    'RESTA DE DOS NUMEROS
    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Dim valor1 As Double
        Dim valor2 As Double
        Dim Res As Double
        Try
            valor1 = CDbl(TextBox10.Text)
            valor2 = CDbl(TextBox9.Text)
            Res = valor1 - valor2
            MsgBox("Resultado: " + CStr(Res))
        Catch ex As Exception
            MsgBox("Introduce números")
        End Try

    End Sub
    'RESTA DE DOS NUMEROS
    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        TextBox10.Text = ""
        TextBox9.Text = ""
    End Sub
    'RESTA DE DOS NUMEROS
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Panel6.Visible = False
    End Sub
    'DIVISION DE DOS NUMEROS
    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        Dim valor1 As Double
        Dim valor2 As Double
        Dim Res As Double
        Try
            valor1 = CDbl(TextBox12.Text)
            valor2 = CDbl(TextBox11.Text)
            If (valor2 <> 0) Then
                Res = valor1 / valor2
                MsgBox("Resultado de la división: " + CStr(Res))
            Else
                MsgBox("En el segundo textbox no se puede introducir 0")
            End If
        Catch ex As Exception
            MsgBox("Introduce un número")
        End Try

    End Sub
    'DIVISION DE DOS NUMEROS
    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        TextBox12.Text = ""
        TextBox11.Text = ""
    End Sub
    'DIVISION DE DOS NUMEROS
    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Panel7.Visible = False
    End Sub
    'MULTIPLICAR DOS NUMEROS
    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        Dim valor1 As Double
        Dim valor2 As Double
        Dim Res As Double
        Try
            valor1 = CDbl(TextBox14.Text)
            valor2 = CDbl(TextBox13.Text)
            Res = valor1 * valor2
            MsgBox("Resultado de la multiplicación: " + CStr(Res))
        Catch ex As Exception
            MsgBox("Introduce un número")
        End Try
    End Sub
    'MULTIPLICAR DOS NUMEROS
    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        TextBox14.Text = ""
        TextBox13.Text = ""
    End Sub
    'MULTIPLICAR DOS NUMEROS
    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        Panel8.Visible = False
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        TextBox15.Text = StrReverse(TextBox16.Text)
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        TextBox15.Text = ""
        TextBox16.Text = ""
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        Panel9.Visible = False
    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        Dim frase As String
        Dim letras As Integer = 0
        Dim Vocales As Integer = 0
        Dim Consonantes As Integer = 0
        Dim Espacios As Integer = 0

        frase = TextBox18.Text
        letras = frase.Length

        For c = 0 To letras - 1
            Select Case frase.ElementAt(c)
                Case "a"
                    Vocales = Vocales + 1
                Case "e"
                    Vocales = Vocales + 1
                Case "i"
                    Vocales = Vocales + 1
                Case "o"
                    Vocales = Vocales + 1
                Case "u"
                    Vocales = Vocales + 1
                Case " "
                    Espacios = Espacios + 1
            End Select
        Next

        Consonantes = letras - Vocales - Espacios
        TextBox17.Text = Vocales
        TextBox19.Text = Consonantes

    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        TextBox18.Text = ""
        TextBox17.Text = ""
        TextBox19.Text = ""
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        Panel10.Visible = False
    End Sub

    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click
        Dim caracter1 As Char
        Dim caracter2 As Char
        Dim frase As String
        Dim contador As Integer = 0
        Dim Repeticiones As Integer = 0

        caracter1 = CChar(TextBox21.Text)
        caracter2 = CChar(TextBox20.Text)
        frase = TextBox22.Text

        While contador < frase.Length - 1
            If caracter1 = frase.ElementAt(contador) Then
                contador = contador + 1
                If caracter2 = frase.ElementAt(contador) Then
                    Repeticiones = Repeticiones + 1
                End If
            End If
            contador = contador + 1
        End While
        TextBox23.Text = Repeticiones
    End Sub

    Private Sub Button32_Click(sender As Object, e As EventArgs) Handles Button32.Click
        TextBox21.Text = ""
        TextBox20.Text = ""
        TextBox23.Text = ""
        TextBox22.Text = ""
    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
        Panel11.Visible = False
    End Sub





    Private Sub Button35_Click(sender As Object, e As EventArgs) Handles Button35.Click
        TextBox25.Text = ""
        TextBox24.Text = ""
    End Sub

    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Button34.Click
        Panel12.Visible = False
    End Sub

    Private Sub Button39_Click(sender As Object, e As EventArgs) Handles Button39.Click
        ListBox1.Items.Clear()
        Dim numero As Integer
        Dim actual As Integer = 1
        Dim anterior As Integer = 1
        Dim suma As Integer
        Try
            numero = CUInt(TextBox27.Text)
            ListBox1.Items.Add(1)
            ListBox1.Items.Add(1)
            For c = 2 To numero - 1
                suma = anterior + actual
                anterior = actual
                actual = suma
                ListBox1.Items.Add(suma)
            Next
        Catch
            MsgBox("Introduce un número entero")
        End Try



    End Sub

    Private Sub Button38_Click(sender As Object, e As EventArgs) Handles Button38.Click
        TextBox27.Text = ""
        ListBox1.Items.Clear()
    End Sub

    Private Sub Button37_Click(sender As Object, e As EventArgs) Handles Button37.Click
        Panel13.Visible = False
    End Sub

    Private Sub Button42_Click(sender As Object, e As EventArgs) Handles Button42.Click
        Dim numero As Integer
        Dim numeroMultiplicar As Integer
        Dim multiplicar As Integer
        Try
            numero = TextBox26.Text
            numeroMultiplicar = TextBox28.Text
            ListBox2.Items.Clear()

            For c = 0 To numeroMultiplicar
                multiplicar = numero * c
                ListBox2.Items.Add(multiplicar)
            Next
        Catch ex As Exception
            MsgBox("Introduce un número entero")
        End Try

    End Sub

    Private Sub Button41_Click(sender As Object, e As EventArgs) Handles Button41.Click
        TextBox26.Text = ""
        TextBox28.Text = ""
        ListBox2.Items.Clear()
    End Sub

    Private Sub Button40_Click(sender As Object, e As EventArgs) Handles Button40.Click
        Panel14.Visible = False
    End Sub

    Private Sub Inicializar_Click(sender As Object, e As EventArgs) Handles Inicializar.Click
        CType(ContextMenuStrip1.SourceControl, TextBox).Text = 0

    End Sub

    Private Sub Incrementar_Click(sender As Object, e As EventArgs) Handles Incrementar.Click
        Try
            CType(ContextMenuStrip1.SourceControl, TextBox).Text = CInt(CType(ContextMenuStrip1.SourceControl, TextBox).Text) + 1
        Catch ex As Exception
            MsgBox("Introduce un valor antes de sumarle 1")
        End Try

    End Sub

    Private Sub Disminuir_Click(sender As Object, e As EventArgs) Handles Disminuir.Click
        Try
            CType(ContextMenuStrip1.SourceControl, TextBox).Text = CInt(CType(ContextMenuStrip1.SourceControl, TextBox).Text) - 1
        Catch ex As Exception
            MsgBox("Introduce un valor antes de restarle 1")
        End Try
    End Sub

    Private Sub CopiarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopiarToolStripMenuItem.Click
        Dim info As String
        Try
            info = CType(ContextMenuStrip1.SourceControl, TextBox).Text
            My.Computer.Clipboard.SetText(info)
        Catch ex As Exception
            MsgBox("No has copiado nada")
        End Try

    End Sub

    Private Sub CortarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CortarToolStripMenuItem.Click
        Dim info As String
        Try
            info = CType(ContextMenuStrip1.SourceControl, TextBox).Text
            My.Computer.Clipboard.SetText(info)
            CType(ContextMenuStrip1.SourceControl, TextBox).Text = ""
        Catch ex As Exception
            MsgBox("No has cortado nada")
        End Try

    End Sub

    Private Sub PegarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PegarToolStripMenuItem.Click
        Dim info As String
        Try
            info = My.Computer.Clipboard.GetText
            CType(ContextMenuStrip1.SourceControl, TextBox).Text = info
        Catch ex As Exception
            MsgBox("No has pegado nada")
        End Try

    End Sub

    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles Button36.Click
        TextBox24.Text = Numalet.ToCardinal(TextBox25.Text)
    End Sub


End Class
