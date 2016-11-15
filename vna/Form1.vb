'Copyright ©2016 Copper Mountain Technologies
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
' and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
' MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR
' ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH
' THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

'Imports Instr = R54Lib
'Imports Instr = TR1300Lib
Imports Instr = S2VNALib
'Imports Instr = S4VNALib

Public Class Form1

    'Private vna As Instr.RVNA
    'Private vna As Instr.TRVNA
    Private vna As Instr.S2VNA
    'Private vna As Instr.S4VNA

    Private vnaFdata(100) As Double
    Private vnaMdata(100) As Double
    Private Const numPoints As Integer = 7
    Private sw As Stopwatch

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'vna = New Instr.RVNA
        'vna = New Instr.TRVNA
        vna = New Instr.S2VNA
        'vna = New Instr.S4VNA

        sw = New Stopwatch

        sw.Reset()
        sw.Start()

        ' wait until vna has initialized
        ' exit if initialization takes longer than 25 seconds
        If Not vna.Ready Then
            Do While Not vna.Ready
                If sw.ElapsedMilliseconds >= 25000 Then
                    MsgBox("VNA Not Responding", CType(MsgBoxStyle.Critical + MsgBoxStyle.SystemModal, MsgBoxStyle), "VNA Automation Test")
                    sw.Stop()
                    Exit Sub
                End If
            Loop
            sw.Stop()
            MsgBox(vna.NAME, CType(MsgBoxStyle.Information + MsgBoxStyle.SystemModal, MsgBoxStyle), "VNA Automation Test")
        Else
            sw.Stop()
            MsgBox(vna.NAME, CType(MsgBoxStyle.Information + MsgBoxStyle.SystemModal, MsgBoxStyle), "VNA Automation Test")
        End If

        vna.SCPI.SYSTem.PRESet()
        vna.SCPI.SENSe(1).FREQuency.STARt = 400000000
        vna.SCPI.SENSe(1).FREQuency.STOP = 420000000
        vna.SCPI.SENSe(1).SWEep.POINts = numPoints
        vna.SCPI.SENSe(1).BANDwidth.RESolution = 20
        vna.SCPI.CALCulate(1).PARameter(1).DEFine = "S21"
        vna.SCPI.DISPlay.WINDow(1).TRACe(1).Y.SCALe.RPOSition = 10
        vna.SCPI.CALCulate(1).SELected.FORMat = "MLOG"
        vnaMdata = CType(vna.SCPI.CALCulate(1).SELected.DATA.FDATa, Double())
        vnaFdata = CType(vna.SCPI.SENSe(1).FREQuency.DATA, Double())
        vna.SCPI.CALCulate(1).SELected.MARKer(numPoints).ACTivate()

    End Sub
End Class