Imports System.IO
Imports System.Text
Imports System.Runtime.InteropServices


Public Class Form1
    Private ficheiro As FileInfo
    Dim old_dxfs(2, 0) As String

    Private Sub DataGridView1_DragEnter(sender As Object, e As DragEventArgs) Handles DataGridView1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub DataGridView1_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs) Handles DataGridView1.DragDrop
        Dim files() As String
        Dim ext As String
        Dim duplicados, nao_suportados As String
        Dim check_ext_res As Integer

        files = e.Data.GetData(DataFormats.FileDrop)
        duplicados = ""
        nao_suportados = ""

        For i = 0 To files.Length - 1
            ficheiro = New FileInfo(files(i))
            ext = Strings.Right(ficheiro.Extension, Strings.Len(ficheiro.Extension) - 1)

            check_ext_res = Check_ext(ext, ficheiro.Name)

            If check_ext_res = 1 Then
                If nao_suportados = "" Then
                    nao_suportados = ficheiro.Name
                Else
                    nao_suportados = nao_suportados & vbNewLine & ficheiro.Name
                End If
            ElseIf check_ext_res = 0 Then
                If Check_DG(ficheiro.Name) = False Then
                    If duplicados = "" Then
                        duplicados = ficheiro.Name
                    Else
                        duplicados = duplicados & vbNewLine & ficheiro.Name
                    End If
                Else
                    DataGridView1.Rows.Add(ficheiro.Name, ficheiro.FullName, ext)
                End If
            End If
        Next

        If duplicados <> "" Then MsgBox("Tentou adicionar o(s) seguinte(s) ficheiro(s) mais do que 1 vez:" & vbNewLine & duplicados, vbExclamation)
        If nao_suportados <> "" Then MsgBox("O(s) seguinte(s) ficheiro(s) não é(são) suportado(s):" & vbNewLine & nao_suportados, vbExclamation)

        DataGridView1.Columns("Extensao").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Obsoleto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Mover").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Application.DoEvents()
    End Sub

    Private Function Check_ext(ByVal extensao As String, ByVal ficheiro As String) As Integer '0 - OK | 1 - NÃO SUPORTADO | 2 - TXT
        Check_ext = 0

        If extensao.ToLower <> "dxf" And extensao.ToLower <> "pdf" And extensao.ToLower <> "easm" And extensao.ToLower <> "jpg" And extensao.ToLower <> "txt" Then
            Check_ext = 1
        ElseIf extensao.ToLower = "txt" Then
            Call txt_processing(Strings.Left(ficheiro, Strings.Len(ficheiro) - (Strings.Len(extensao) + 1)))
            Check_ext = 2
        End If

    End Function

    Private Function Check_DG(ByVal ficheiro As String) As Boolean
        Dim a As Integer

        Check_DG = True

        If DataGridView1.RowCount > 0 Then
            For a = 0 To DataGridView1.RowCount - 1
                If ficheiro = DataGridView1.Rows(a).Cells("Desenho").Value Then
                    Check_DG = False
                    Exit For
                End If
            Next
        End If
    End Function

    Private Sub Button_browse_Click(sender As Object, e As EventArgs) Handles Button_browse.Click
        SaveFileDialog1.FileName = "Selecione a pasta do projeto na produção"
        If TextBox_pasta_prod.Text = "" Then
            SaveFileDialog1.InitialDirectory = "G:\producao\" & Date.Now.Year.ToString & "\Clientes"
        Else
            SaveFileDialog1.InitialDirectory = TextBox_pasta_prod.Text
        End If
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            TextBox_pasta_prod.Text = Path.GetDirectoryName(SaveFileDialog1.FileName)
        End If
    End Sub

    Private Sub Button_sair_Click(sender As Object, e As EventArgs) Handles Button_sair.Click
        End
    End Sub

    Private Sub Button_Limpar_Click(sender As Object, e As EventArgs) Handles Button_Limpar.Click
        DataGridView1.Rows.Clear()
        TextBox_pasta_prod.Text = Nothing
        Array.Clear(old_dxfs, 0, old_dxfs.Length)
        ReDim old_dxfs(2, 0)
    End Sub

    Private Sub Button_ok_Click(sender As Object, e As EventArgs) Handles Button_ok.Click
        Dim pasta_prod, origem, ficheiro, extensao, destino, folder1 As String
        Dim case_code, res_erros, res_avisos, res_aux As Integer
        Dim progress As Double

        Call clean_DG(DataGridView1)

        If DataGridView1.RowCount = 0 Then
            MsgBox("Não adicionou nenhum ficheiro.", vbExclamation)
            Exit Sub
        ElseIf TextBox_pasta_prod.Text = "" Then
            MsgBox("Não selecionou a pasta na produção.", vbExclamation)
            Exit Sub
        End If

        pasta_prod = TextBox_pasta_prod.Text
        If Strings.Right(pasta_prod, 1) <> "\" Then pasta_prod = pasta_prod & "\"
        folder1 = Strings.Left(pasta_prod, Strings.Len(pasta_prod) - 1)

        'backup_folder(folder1)

        Button_browse.Enabled = False
        Button_ok.Enabled = False
        Button_sair.Enabled = False
        Button_Limpar.Enabled = False

        case_code = verificar_pastas(pasta_prod)

        res_erros = 0
        res_avisos = 0

        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = DataGridView1.RowCount

        Label2.Left = (Me.Width \ 2) - (Label2.Width \ 2)
        ProgressBar1.Visible = True

        Select Case case_code '0 - sem pastas; 1 - projeto; 2 - encomenda
            Case 0
                MsgBox("Erro nas pastas", vbCritical)
                Exit Sub

            Case 1
                For i = 0 To DataGridView1.RowCount - 1
                    DataGridView1.CurrentCell = DataGridView1.Rows(i).Cells(0)
                    extensao = DataGridView1.Rows(i).Cells("Extensao").Value
                    origem = Strings.Left(DataGridView1.Rows(i).Cells("Localizacao").Value, Strings.Len(DataGridView1.Rows(i).Cells("Localizacao").Value) - Strings.Len(DataGridView1.Rows(i).Cells("Desenho").Value))
                    If extensao.ToLower = "dxf" Then
                        destino = pasta_prod & "DXF\"
                    ElseIf extensao.ToLower = "pdf" Then
                        destino = pasta_prod & "PDF\"
                    ElseIf extensao.ToLower = "easm" Or extensao.ToLower = "eprt" Then
                        destino = pasta_prod & "3D\"
                    ElseIf extensao.ToLower = "jpg" Then
                        destino = pasta_prod & "ARTIGOS\"
                    Else
                        introduzir_dados(i, "Obsoleto", "AVISO")
                        change_color(i, "Obsoleto", 1)
                        introduzir_dados(i, "Mover", "Extensão não suportada")
                        change_color(i, "Mover", 1)
                        res_avisos = res_avisos + 1
                        Continue For
                    End If

                    ficheiro = Strings.Left(DataGridView1.Rows(i).Cells("Desenho").Value, Strings.Len(DataGridView1.Rows(i).Cells("Desenho").Value) - Strings.Len(extensao) - 1)
                    res_aux = mover_ficheiro(origem, destino, ficheiro, extensao, i)
                    If res_aux = 1 Then '0- Ok | 1 - Erro | 2 - Aviso
                        res_erros = res_erros + 1
                    ElseIf res_aux = 2 Then
                        res_avisos = res_avisos + 1
                    End If

                    ProgressBar1.Value = i + 1
                    progress = ((((i + 1) / ProgressBar1.Maximum) * 100) \ 1)
                    Label2.Text = progress & "%"
                Next

                Call msg_erro(res_erros, res_avisos)

                DataGridToCSV(DataGridView1, ";", Strings.Right(folder1, Strings.Len(folder1) - folder1.LastIndexOf("\")))

            Case 2
                destino = pasta_prod
                For i = 0 To DataGridView1.RowCount - 1
                    DataGridView1.CurrentCell = DataGridView1.Rows(i).Cells(0)
                    extensao = DataGridView1.Rows(i).Cells("Extensao").Value
                    If extensao.ToLower <> "dxf" And extensao.ToLower <> "pdf" And extensao.ToLower <> "easm" And extensao.ToLower <> "eprt" And extensao.ToLower <> "jpg" Then
                        introduzir_dados(i, "Obsoleto", "AVISO")
                        change_color(i, "Obsoleto", 1)
                        introduzir_dados(i, "Mover", "Extensão não suportada")
                        change_color(i, "Mover", 1)
                        res_avisos = res_avisos + 1
                        Continue For
                    End If
                    origem = Strings.Left(DataGridView1.Rows(i).Cells("Localizacao").Value, Strings.Len(DataGridView1.Rows(i).Cells("Localizacao").Value) - Strings.Len(DataGridView1.Rows(i).Cells("Desenho").Value))
                    ficheiro = Strings.Left(DataGridView1.Rows(i).Cells("Desenho").Value, Strings.Len(DataGridView1.Rows(i).Cells("Desenho").Value) - Strings.Len(extensao) - 1)
                    res_aux = mover_ficheiro(origem, destino, ficheiro, extensao, i)
                    If res_aux = 1 Then '0- Ok | 1 - Erro | 2 - Aviso
                        res_erros = res_erros + 1
                    ElseIf res_aux = 2 Then
                        res_avisos = res_avisos + 1
                    End If

                    ProgressBar1.Value = i + 1
                    progress = ((((i + 1) / ProgressBar1.Maximum) * 100) \ 1)
                    Label2.Text = progress & "%"
                Next

                Call msg_erro(res_erros, res_avisos)

                DataGridToCSV(DataGridView1, ";", Strings.Right(folder1, Strings.Len(folder1) - folder1.LastIndexOf("\")))

        End Select
        DataGridView1.ClearSelection()
        ProgressBar1.Visible = False
        Label2.Text = ""
        Button_browse.Enabled = True
        Button_ok.Enabled = True
        Button_sair.Enabled = True
        Button_Limpar.Enabled = True
    End Sub

    Sub msg_erro(erros As Integer, avisos As Integer)
        If avisos = 0 Then
            If erros = 0 Then
                MsgBox("Concluido com sucesso")
            ElseIf erros = 1 Then
                MsgBox("Concluido com 1 erro.", vbCritical)
            Else
                MsgBox("Concluido com " & erros & " erros.", vbCritical)
            End If
        ElseIf erros = 0 Then
            If avisos = 1 Then
                MsgBox("Concluido com 1 aviso.", vbExclamation)
            Else
                MsgBox("Concluido com " & avisos & " avisos.", vbExclamation)
            End If
        Else
            If avisos = 1 And erros = 1 Then
                MsgBox("Concluido com 1 erro e 1 aviso.", vbCritical)
            ElseIf avisos = 1 Then
                MsgBox("Concluido com " & erros & " erros e 1 aviso.", vbCritical)
            ElseIf erros = 1 Then
                MsgBox("Concluido com 1 erro e " & avisos & " avisos.", vbCritical)
            Else
                MsgBox("Concluido com " & erros & " erros e " & avisos & " avisos.", vbCritical)
            End If
        End If
    End Sub

    ''' <summary>
    ''' VERIFICA SE AS PASTAS "OBSOLETO" EXISTEM
    ''' CASO NÃO EXISTAM, SÃO CRIADAS
    ''' </summary>
    ''' <param name="pasta"></param>
    ''' <returns>0 - sem pastas; 1 - projeto; 2 - encomenda</returns>
    Function verificar_pastas(ByVal pasta As String) As Integer '0 - sem pastas; 1 - projeto; 2 - encomenda
        Dim edrw, pdf, dxf, edrw_obsoleto, pdf_obsoleto, dxf_obsoleto, obsoleto, artigos As Boolean
        Dim answer As Integer
        Dim txt, txt_obsoleto As String

        txt = ""
        txt_obsoleto = ""
        verificar_pastas = 1

        If My.Computer.FileSystem.GetDirectoryInfo(pasta).Exists = False Then
            MsgBox("O caminho introduzido não é válido", MsgBoxStyle.Critical)
            verificar_pastas = 0
        Else
            edrw = verificar_subpasta(pasta & "3D\")
            If edrw = False Then txt = txt & vbNewLine & "3D"

            pdf = verificar_subpasta(pasta & "PDF\")
            If pdf = False Then txt = txt & vbNewLine & "PDF"

            dxf = verificar_subpasta(pasta & "DXF\")
            If dxf = False Then txt = txt & vbNewLine & "DXF"

            artigos = verificar_subpasta(pasta & "ARTIGOS\")
            If artigos = False Then txt = txt & vbNewLine & "ARTIGOS"

            obsoleto = verificar_subpasta(pasta & "OBSOLETO\")

            If txt <> "" Then
                answer = MsgBox("As seguintes pastas não existem:" & txt & vbNewLine & vbNewLine & "Criar pastas?", MsgBoxStyle.YesNo)
                If answer = 6 Then 'YES
                    If edrw = False Then
                        MkDir(pasta & "3D\")
                        MkDir(pasta & "3D\OBSOLETO\")
                    End If
                    If pdf = False Then
                        MkDir(pasta & "PDF\")
                        MkDir(pasta & "PDF\OBSOLETO\")
                    End If
                    If dxf = False Then
                        MkDir(pasta & "DXF\")
                        MkDir(pasta & "DXF\OBSOLETO\")
                    End If
                    If artigos = False Then
                        MkDir(pasta & "ARTIGOS\")
                    End If

                    edrw_obsoleto = verificar_subpasta(pasta & "3D\OBSOLETO\")
                    If edrw_obsoleto = False Then txt_obsoleto = txt_obsoleto & vbNewLine & "3D"

                    pdf_obsoleto = verificar_subpasta(pasta & "PDF\OBSOLETO\")
                    If pdf_obsoleto = False Then txt_obsoleto = txt_obsoleto & vbNewLine & "PDF"

                    dxf_obsoleto = verificar_subpasta(pasta & "DXF\OBSOLETO\")
                    If dxf_obsoleto = False Then txt_obsoleto = txt_obsoleto & vbNewLine & "DXF"

                    If txt_obsoleto <> "" Then
                        answer = MsgBox("As seguintes pastas OBSOLETO não existem:" & txt_obsoleto & vbNewLine & vbNewLine & "Criar pastas?", MsgBoxStyle.YesNo)
                        If answer = 6 Then 'YES OBSOLETOS
                            If edrw_obsoleto = False Then MkDir(pasta & "3D\OBSOLETO\")
                            If pdf_obsoleto = False Then MkDir(pasta & "PDF\OBSOLETO\")
                            If dxf_obsoleto = False Then MkDir(pasta & "DXF\OBSOLETO\")
                        Else 'NO OBSOLETOS
                            verificar_pastas = 0
                        End If
                    End If
                Else 'NO
                    verificar_pastas = 2
                    If obsoleto = False Then
                        answer = MsgBox("A pasta OBSOLETO não existe." & vbNewLine & vbNewLine & "Criar pasta?", MsgBoxStyle.YesNo)
                        If answer = 6 Then 'yes
                            MkDir(pasta & "OBSOLETO\")
                        Else
                            verificar_pastas = 0
                        End If
                    End If
                End If
            Else
                edrw_obsoleto = verificar_subpasta(pasta & "3D\OBSOLETO\")
                If edrw_obsoleto = False Then txt_obsoleto = txt_obsoleto & vbNewLine & "3D"

                pdf_obsoleto = verificar_subpasta(pasta & "PDF\OBSOLETO\")
                If pdf_obsoleto = False Then txt_obsoleto = txt_obsoleto & vbNewLine & "PDF"

                dxf_obsoleto = verificar_subpasta(pasta & "DXF\OBSOLETO\")
                If dxf_obsoleto = False Then txt_obsoleto = txt_obsoleto & vbNewLine & "DXF"

                If txt_obsoleto <> "" Then
                    answer = MsgBox("As seguintes pastas OBSOLETO não existem:" & txt_obsoleto & vbNewLine & vbNewLine & "Criar pastas?", MsgBoxStyle.YesNo)
                    If answer = 6 Then 'YES OBSOLETOS
                        If edrw_obsoleto = False Then
                            MkDir(pasta & "3D\OBSOLETO\")
                        End If
                        If pdf_obsoleto = False Then
                            MkDir(pasta & "PDF\OBSOLETO\")
                        End If
                        If dxf_obsoleto = False Then
                            MkDir(pasta & "DXF\OBSOLETO\")
                        End If
                    Else 'NO OBSOLETOS
                        verificar_pastas = 0
                    End If
                End If
            End If
        End If
    End Function

    Function verificar_subpasta(ByVal subpasta As String) As Boolean
        If My.Computer.FileSystem.GetDirectoryInfo(subpasta).Exists = False Then
            verificar_subpasta = False
        Else
            verificar_subpasta = True
        End If
    End Function

    ''' <summary>
    ''' MOVE O FICHEIRO PARA A PASTA DE DESTINO
    ''' </summary>
    ''' <param name="origem"></param>
    ''' <param name="destino"></param>
    ''' <param name="nome_ficheiro"></param>
    ''' <param name="extensao_ficheiro"></param>
    ''' <param name="linha"></param>
    ''' <returns>0- Ok | 1 - Erro | 2 - Aviso</returns>
    Function mover_ficheiro(ByVal origem As String, ByVal destino As String, ByVal nome_ficheiro As String, extensao_ficheiro As String, linha As Integer) As Integer
        '0- Ok | 1 - Erro | 2 - Aviso

        Dim data, filename, new_filename, pasta_obsoleto, DestFile, DestFileFull, txtFile, GEOname As String
        Dim DestFileInUse, OrigFileInUse, GEOexists As Boolean
        Dim res_check_old_dxf As Integer

        data = Date.Now.ToString("yyyy-MM-dd")
        filename = nome_ficheiro & "." & extensao_ficheiro
        DestFile = ""
        txtFile = ""
        GEOexists = False

        If extensao_ficheiro.ToLower = "jpg" And File.Exists(destino & filename) = True Then
            My.Computer.FileSystem.DeleteFile(destino & filename)
        End If

        pasta_obsoleto = destino & "OBSOLETO\"

        If extensao_ficheiro.ToLower = "dxf" Then
            res_check_old_dxf = search_old_dxfs(nome_ficheiro)
            If res_check_old_dxf = -1 Then
                DestFile = nome_ficheiro
            Else
                DestFile = old_dxfs(0, res_check_old_dxf)
                txtFile = old_dxfs(2, res_check_old_dxf)
                introduzir_dados(linha, "NOME_ANTIGO", nome_ficheiro & " (" & DestFile & ")")
            End If

            GEOname = DestFile
            If File.Exists(destino & GEOname & ".GEO") = True Then
                GEOexists = True
            ElseIf File.Exists(destino & nome_ficheiro & ".GEO") = True Then
                GEOname = nome_ficheiro
                GEOexists = True
            End If

            If GEOexists = True Then
                If IsFileOpen(destino & GEOname & ".GEO") = True Then
                    introduzir_dados(linha, "Obsoleto", "AVISO")
                    change_color(linha, "Obsoleto", 1)
                    introduzir_dados(linha, "Mover", "Não foi possivél mover ficheiro GEO antigo")
                    change_color(linha, "Mover", 1)
                    mover_ficheiro = 2
                Else
                    My.Computer.FileSystem.MoveFile(destino & GEOname & ".GEO", pasta_obsoleto & GEOname & "_OBSOLETO_" & data & "." & ".GEO")
                End If
            End If
        End If

        If DestFile <> "" Then
            DestFileFull = DestFile & "." & extensao_ficheiro
            If File.Exists(destino & DestFileFull) = False Then
                DestFileFull = filename
                DestFile = nome_ficheiro
            End If
        Else
            DestFileFull = filename
            DestFile = nome_ficheiro
        End If

        DestFileInUse = IsFileOpen(destino & DestFileFull)
        OrigFileInUse = IsFileOpen(origem & filename)

        If DestFileInUse = True Then
            introduzir_dados(linha, "Obsoleto", "ERRO")
            change_color(linha, "Obsoleto", 0)
            introduzir_dados(linha, "Mover", "Ficheiro antigo em uso")
            change_color(linha, "Mover", 0)
            mover_ficheiro = 1
        ElseIf File.Exists(origem & filename) = False Then
            introduzir_dados(linha, "Obsoleto", "ERRO")
            change_color(linha, "Obsoleto", 0)
            introduzir_dados(linha, "Mover", "Ficheiro novo não existe")
            change_color(linha, "Mover", 0)
            mover_ficheiro = 1
        ElseIf OrigFileInUse = True Then
            introduzir_dados(linha, "Obsoleto", "ERRO")
            change_color(linha, "Obsoleto", 0)
            introduzir_dados(linha, "Mover", "Ficheiro novo em uso")
            change_color(linha, "Mover", 0)
            mover_ficheiro = 1
        Else
            If File.Exists(destino & DestFileFull) = True Then
                new_filename = DestFile & "_OBSOLETO_" & data & "." & extensao_ficheiro
                If File.Exists(pasta_obsoleto & new_filename) = False Then
                    My.Computer.FileSystem.MoveFile(destino & DestFileFull, pasta_obsoleto & new_filename)
                    If filename <> DestFileFull And File.Exists(destino & filename) = True Then
                        new_filename = filename & "_OBSOLETO_" & data & "." & extensao_ficheiro
                        My.Computer.FileSystem.MoveFile(destino & filename, pasta_obsoleto & new_filename)
                    End If
                    introduzir_dados(linha, "Obsoleto", "SUCESSO")
                    change_color(linha, "Obsoleto", 2)
                    My.Computer.FileSystem.MoveFile(origem & filename, destino & filename)
                    introduzir_dados(linha, "Mover", "SUCESSO")
                    change_color(linha, "Mover", 2)
                    mover_ficheiro = 0
                    If txtFile <> "" And File.Exists(origem & txtFile & ".txt") Then My.Computer.FileSystem.DeleteFile(origem & txtFile & ".txt")
                Else
                    introduzir_dados(linha, "Obsoleto", "ERRO")
                    change_color(linha, "Obsoleto", 0)
                    introduzir_dados(linha, "Mover", "Obsoleto já existe")
                    change_color(linha, "Mover", 0)
                    mover_ficheiro = 1
                End If
            Else
                If extensao_ficheiro.ToLower <> "jpg" Then
                    introduzir_dados(linha, "Obsoleto", "NÃO EXISTE")
                    change_color(linha, "Obsoleto", 2)
                Else
                    introduzir_dados(linha, "Obsoleto", " ")
                End If
                My.Computer.FileSystem.MoveFile(origem & filename, destino & filename)
                introduzir_dados(linha, "Mover", "SUCESSO")
                change_color(linha, "Mover", 2)
                mover_ficheiro = 0
                If txtFile <> "" Then My.Computer.FileSystem.DeleteFile(origem & txtFile & ".txt")
            End If
        End If
        Application.DoEvents()
    End Function

    Sub introduzir_dados(linha As Integer, coluna As String, txt As String)
        DataGridView1.Rows(linha).Cells(coluna).Value = txt
    End Sub

    Sub backup_folder(ByVal folder As String)
        Dim data, backup_folder, zip_name As String

        data = Date.Now.ToString("yyyy-MM-dd_HH-mm")
        backup_folder = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        zip_name = Strings.Right(folder, Strings.Len(folder) - folder.LastIndexOf("\")) & "_" & data & ".zip"

        If File.Exists(backup_folder & zip_name) = True Then
            zip_name = Strings.Right(folder, Strings.Len(folder) - folder.LastIndexOf("\")) & "_" & Date.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".zip"
        End If

        System.IO.Compression.ZipFile.CreateFromDirectory(folder & "\", backup_folder & zip_name)

    End Sub

    ''' <summary>
    ''' MUDA A COR DO TEXTO NA DATAGRIDVIEW
    ''' </summary>
    ''' <param name="linha"></param>
    ''' <param name="coluna"></param>
    ''' <param name="tipo">0 - Erro | 1 - Aviso | 2 - Sucesso</param>
    Sub change_color(linha As Integer, coluna As String, tipo As Integer)
        'tipo: 0 - Erro; 1 - Aviso; 2 - Sucesso

        If tipo = 0 Then DataGridView1.Rows(linha).Cells(coluna).Style.ForeColor = Color.Red
        If tipo = 1 Then DataGridView1.Rows(linha).Cells(coluna).Style.ForeColor = Color.Gold
        If tipo = 2 Then DataGridView1.Rows(linha).Cells(coluna).Style.ForeColor = Color.ForestGreen
    End Sub

    Private Sub DataGridToCSV(dt As DataGridView, ByVal Qualifier As String, ByVal name As String)
        Dim sair_while As Integer = 0
        Dim t As Integer = 0

        Dim Directory As String = My.Computer.FileSystem.SpecialDirectories.Desktop
        'System.IO.Directory.CreateDirectory(TempDirectory)
        Dim data As String = Date.Now.ToString("yyyy-MM-dd_HH-mm")

        Dim file_name As String = name & "_" & data & ".csv"

        While sair_while = 0
            t = t + 1
            If File.Exists(Directory & file_name) = True Then
                file_name = name & "_" & data & "_" & t & ".csv"
            Else
                sair_while = 1
            End If
        End While
        'Dim oWrite As System.IO.StreamWriter
        Dim oWrite As New System.IO.StreamWriter(Directory & file_name, False, System.Text.Encoding.GetEncoding("iso-8859-1"))
        'oWrite = IO.File.CreateText(TempDirectory & file)
        Dim colunas = New String() {"Desenho", "Extensao", "Obsoleto", "Mover", "NOME_ANTIGO"}

        Dim CSV As StringBuilder = New StringBuilder()

        Dim i As Integer = 1
        Dim aux As Object
        Dim CSVHeader As StringBuilder = New StringBuilder()
        For Each c As String In colunas
            If i = 1 Then
                CSVHeader.Append(dt.Columns(c).HeaderText.ToString())
            Else
                aux = Qualifier & dt.Columns(c).HeaderText
                CSVHeader.Append(aux.ToString())
            End If
            i += 1
        Next

        CSV.AppendLine(CSVHeader.ToString())
        'oWrite.WriteLine(CSVHeader.ToString())
        'oWrite.Flush()

        For r As Integer = 0 To dt.Rows.Count - 1

            Dim CSVLine As StringBuilder = New StringBuilder()
            'Dim s As String = ""
            For c As Integer = 0 To colunas.Count - 1
                If c = 0 Then
                    CSVLine.Append(Strings.Left(dt.Rows(r).Cells(colunas(c)).Value.ToString(), Len(dt.Rows(r).Cells(colunas(c)).Value.ToString()) -
                                                                                              (Len(dt.Rows(r).Cells("Extensao").Value.ToString()) + 1)))
                    's = s & dt.Rows(r).Cells(colunas(c)).Value.ToString()
                Else
                    If dt.Rows(r).Cells(colunas(c)).Value <> Nothing Then
                        CSVLine.Append(Qualifier & dt.Rows(r).Cells(colunas(c)).Value.ToString())
                        's = s & Qualifier & dt.Rows(r).Cells(colunas(c)).Value.ToString()
                    Else
                        CSVLine.Append(Qualifier & " ")
                    End If
                End If

            Next
            'oWrite.WriteLine(s)
            'oWrite.Flush()
            CSV.AppendLine(CSVLine.ToString())
            CSVLine.Clear()
        Next

        oWrite.Write(CSV.ToString())

        oWrite.Close()
        oWrite = Nothing

        'System.Diagnostics.Process.Start(TempDirectory & "\" & file_name)

        GC.Collect()

    End Sub

    Function IsFileOpen(ByVal file As String) As Boolean
        Dim stream As FileStream = Nothing
        Dim finfo As FileInfo

        finfo = New FileInfo(file)
        IsFileOpen = False
        Try
            stream = finfo.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None)
            stream.Close()
        Catch ex As Exception

            If TypeOf ex Is IOException AndAlso IsFileLocked(ex) Then
                IsFileOpen = True
            End If
        End Try
    End Function

    Private Shared Function IsFileLocked(exception As Exception) As Boolean
        Dim errorCode As Integer = Marshal.GetHRForException(exception) And ((1 << 16) - 1)
        Return errorCode = 32 OrElse errorCode = 33
    End Function

    Private Sub clean_DG(dt As DataGridView)
        Dim r, count As Integer
        Dim Cel_Val, des As String

        count = dt.Rows.Count - 1

        For r = count To 0 Step -1
            Cel_Val = dt.Rows(r).Cells("Mover").Value
            des = dt.Rows(r).Cells("Desenho").Value

            If Cel_Val = "SUCESSO" Then dt.Rows.Remove(dt.Rows(r))

            Application.DoEvents()
        Next r

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Label2.BackColor = Color.Transparent
        Label2.Text = ""
        '    Label2.BringToFront()
    End Sub

    Private Sub txt_processing(ByVal txt_name As String)
        Dim old_name, new_name As String
        Dim array_len As Integer

        new_name = Strings.Right(txt_name, Strings.Len(txt_name) - (txt_name.LastIndexOf(" ") + 1))
        old_name = Strings.Left(txt_name, Strings.Len(txt_name) - (Strings.Len(new_name) + 1))
        old_name = Strings.Left(old_name, old_name.LastIndexOf(" "))

        array_len = old_dxfs.GetLength(1)

        ReDim Preserve old_dxfs(2, array_len)
        old_dxfs(0, array_len - 1) = old_name
        old_dxfs(1, array_len - 1) = new_name
        old_dxfs(2, array_len - 1) = txt_name
    End Sub

    Private Function search_old_dxfs(ByVal dxf_name As String) As Integer
        Dim i As Integer

        search_old_dxfs = -1

        For i = 0 To old_dxfs.GetLength(1) - 2
            If old_dxfs(1, i) = dxf_name Then
                search_old_dxfs = i
                Exit For
            End If
        Next
    End Function


End Class
