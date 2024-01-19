Imports System
Imports Excel = Microsoft.Office.Interop.Excel

'public variables
Public Class Var
    Public Shared Directory As String = IO.Path.Combine(My.Application.Info.DirectoryPath, "Save.xlsm")
    Public Shared Sheet As String = "Patients"
    Public Shared max As Integer = 1000
End Class

'form 1
Public Class Form1

    'dim excel applications as variables
    Dim APP As New Excel.Application
    Dim worksheet As Excel.Worksheet
    Dim workbook As Excel.Workbook

    'load variables with directory and sheet number
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        workbook = APP.Workbooks.Open(Var.Directory)
        worksheet = workbook.Worksheets(Var.Sheet)
        txtrow.Text = 1
    End Sub

    'add new patient
    Private Sub Btnnew_Click(sender As Object, e As EventArgs) Handles btnnew.Click
        'dim variables
        workbook = APP.Workbooks.Open(Var.Directory)
        worksheet = workbook.Worksheets(Var.Sheet)
        Dim rownum As Integer = 1
        Dim idd As String = worksheet.Cells(rownum, 2).Value

        'for loop to search for the last empty cell in row 2
        For id As Integer = 1 To Var.max
            If idd = "" Then
                idd = worksheet.Cells(rownum, 2).Value
                txtrow.Text = rownum
                txtdoa.Text = ""
                txtmrn.Text = ""
                txtage.Text = ""
            Else
                rownum = rownum + 1
                Console.WriteLine(rownum)
                idd = worksheet.Cells(rownum, 2).Value
            End If
        Next
        workbook.Close()

        'Textboxes empty
        txtdoa.Text = ""
        txtmrn.Text = ""
        txtage.Text = ""
        txtadmiss.Text = ""
        txtcrit.Text = ""
        txtdiag1.Text = ""
        txtdiag2.Text = ""
        txtdiag3.Text = ""
        txttotalpa.Text = ""
        radadenop.Checked = False
        radadenon.Checked = False
        radbcsp.Checked = False
        radbcsn.Checked = False
        radcovn.Checked = False
        radcovp.Checked = False
        radinfluan.Checked = False
        radinfluap.Checked = False
        radmycopn.Checked = False
        radmycopp.Checked = False
        radrotan.Checked = False
        radrotap.Checked = False
        radrstn.Checked = False
        radrstp.Checked = False
        radrsvn.Checked = False
        radrsvp.Checked = False
        radscsn.Checked = False
        radscsp.Checked = False
        raducsn.Checked = False
        raducsp.Checked = False
    End Sub

    'closing button
    Private Sub btnclose_Click(sender As Object, e As EventArgs) Handles btnclose.Click
        End
    End Sub

    'save button
    Private Sub btnsav_Click(sender As Object, e As EventArgs) Handles btnsav.Click
        'dim variables
        workbook = APP.Workbooks.Open(Var.Directory)
        worksheet = workbook.Worksheets(Var.Sheet)
        Dim rownum As Integer = txtrow.Text
        Dim Neg As String = "NEG"
        Dim Pos As String = "POS"

        'save as the data in the form
        If txtmrn.Text = "" Then
            MsgBox("Please Enter MRN")
        Else
            worksheet.Cells(rownum, 1).Value = txtdoa.Text
            worksheet.Cells(rownum, 2).Value = txtmrn.Text
            worksheet.Cells(rownum, 3).Value = txtage.Text
            worksheet.Cells(rownum, 4).Value = txtdiag1.Text
            worksheet.Cells(rownum, 5).Value = txtdiag2.Text
            worksheet.Cells(rownum, 6).Value = txtdiag3.Text

            If radcovp.Checked Then
                worksheet.Cells(rownum, 7).Value = Pos
            ElseIf radcovn.Checked Then
                worksheet.Cells(rownum, 7).Value = Neg
            End If

            worksheet.Cells(rownum, 8).Value = txtcrit.Text
            worksheet.Cells(rownum, 9).Value = txtadmiss.Text
            worksheet.Cells(rownum, 10).Value = txttotalpa.Text

            If radrotap.Checked Then
                worksheet.Cells(rownum, 11).Value = Pos
            ElseIf radrotan.Checked Then
                worksheet.Cells(rownum, 11).Value = Neg
            End If

            If radadenop.Checked Then
                worksheet.Cells(rownum, 12).Value = Pos
            ElseIf radadenon.Checked Then
                worksheet.Cells(rownum, 12).Value = Neg
            End If

            If radrsvp.Checked Then
                worksheet.Cells(rownum, 13).Value = Pos
            ElseIf radrsvn.Checked Then
                worksheet.Cells(rownum, 13).Value = Neg
            End If

            If radrstp.Checked Then
                worksheet.Cells(rownum, 14).Value = Pos
            ElseIf radrstn.Checked Then
                worksheet.Cells(rownum, 14).Value = Neg
            End If

            If radinfluap.Checked Then
                worksheet.Cells(rownum, 15).Value = Pos
            ElseIf radinfluan.Checked Then
                worksheet.Cells(rownum, 15).Value = Neg
            End If

            If radbcsp.Checked Then
                worksheet.Cells(rownum, 16).Value = Pos
            ElseIf radbcsn.Checked Then
                worksheet.Cells(rownum, 16).Value = Neg
            End If

            If raducsp.Checked Then
                worksheet.Cells(rownum, 17).Value = Pos
            ElseIf raducsn.Checked Then
                worksheet.Cells(rownum, 17).Value = Neg
            End If

            If radscsp.Checked Then
                worksheet.Cells(rownum, 18).Value = Pos
            ElseIf radscsn.Checked Then
                worksheet.Cells(rownum, 18).Value = Neg
            End If

            If radmycopp.Checked Then
                worksheet.Cells(rownum, 19).Value = Pos
            ElseIf radmycopn.Checked Then
                worksheet.Cells(rownum, 19).Value = Neg
            End If

            workbook.Save()
            workbook.Close()
        End If

    End Sub

    'load existing patient
    Private Sub btnload_Click(sender As Object, e As EventArgs) Handles btnload.Click

        'dim variables
        workbook = APP.Workbooks.Open(Var.Directory)
        worksheet = workbook.Worksheets(Var.Sheet)
        Dim Idnum As String = InputBox("Insert row number:")
        Dim rownum As Integer = 1
        Dim idd As String = worksheet.Cells(rownum, 2).Value

        'for loop to find the written number
        For id As Integer = 1 To Var.max
            If idd = Idnum Then
                idd = worksheet.Cells(rownum, 2).Value
                txtrow.Text = rownum

                txtdoa.Text = worksheet.Cells(rownum, 1).Value
                txtmrn.Text = worksheet.Cells(rownum, 2).Value
                txtage.Text = worksheet.Cells(rownum, 3).Value
                txtdiag1.Text = worksheet.Cells(rownum, 4).Value
                txtdiag2.Text = worksheet.Cells(rownum, 5).Value
                txtdiag3.Text = worksheet.Cells(rownum, 6).Value

                'covid test result
                If worksheet.Cells(rownum, 7).Value = "POS" Then
                    radcovp.Checked = True
                ElseIf worksheet.Cells(rownum, 7).Value = "NEG" Then
                    radcovn.Checked = True
                End If

                txtcrit.Text = worksheet.Cells(rownum, 8).Value
                txtadmiss.Text = worksheet.Cells(rownum, 9).Value
                txttotalpa.Text = worksheet.Cells(rownum, 10).Value

                If worksheet.Cells(rownum, 11).Value = "POS" Then
                    radrotap.Checked = True
                ElseIf worksheet.Cells(rownum, 11).Value = "NEG" Then
                    radrotan.Checked = True
                End If

                If worksheet.Cells(rownum, 12).Value = "POS" Then
                    radadenop.Checked = True
                ElseIf worksheet.Cells(rownum, 12).Value = "NEG" Then
                    radadenon.Checked = True
                End If

                If worksheet.Cells(rownum, 13).Value = "POS" Then
                    radrsvp.Checked = True
                ElseIf worksheet.Cells(rownum, 13).Value = "NEG" Then
                    radrsvn.Checked = True
                End If

                If worksheet.Cells(rownum, 14).Value = "POS" Then
                    radrstp.Checked = True
                ElseIf worksheet.Cells(rownum, 14).Value = "NEG" Then
                    radrstn.Checked = True
                End If

                If worksheet.Cells(rownum, 15).Value = "POS" Then
                    radinfluap.Checked = True
                ElseIf worksheet.Cells(rownum, 15).Value = "NEG" Then
                    radinfluan.Checked = True
                End If

                If worksheet.Cells(rownum, 16).Value = "POS" Then
                    radbcsp.Checked = True
                ElseIf worksheet.Cells(rownum, 16).Value = "NEG" Then
                    radbcsn.Checked = True
                End If

                If worksheet.Cells(rownum, 17).Value = "POS" Then
                    raducsp.Checked = True
                ElseIf worksheet.Cells(rownum, 17).Value = "NEG" Then
                    raducsn.Checked = True
                End If

                If worksheet.Cells(rownum, 18).Value = "POS" Then
                    radscsp.Checked = True
                ElseIf worksheet.Cells(rownum, 18).Value = "NEG" Then
                    radscsn.Checked = True
                End If

                If worksheet.Cells(rownum, 19).Value = "POS" Then
                    radmycopp.Checked = True
                ElseIf worksheet.Cells(rownum, 19).Value = "NEG" Then
                    radmycopn.Checked = True
                End If
                rownum = rownum + 1
                idd = worksheet.Cells(rownum, 2).Value
            Else
                rownum = rownum + 1
                idd = worksheet.Cells(rownum, 2).Value
            End If
        Next
    End Sub
End Class