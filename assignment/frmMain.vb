Public Class frmMain

    Private Sub ThêmSảnPhẩmToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ThêmSảnPhẩmToolStripMenuItem.Click
        frmCapnhatsanpham.Show()
    End Sub

    Private Sub XemKhácHàngToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles XemKhácHàngToolStripMenuItem.Click
        frmKH.Show()
    End Sub

    Private Sub ChỉnhSữaKHToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ChỉnhSữaKHToolStripMenuItem.Click
        frmDieuchinhKH.Show()
    End Sub

    Private Sub XemSảnPhẩmToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles XemSảnPhẩmToolStripMenuItem.Click
        frmXemsanpham.Show()
    End Sub

    Private Sub btnNhapthongtinKH_Button(sender As Object, e As EventArgs) Handles btnNhapthongtinKH.Click
        frmDieuchinhKH.Show()
    End Sub

    Private Sub btnXemKH_Button(sender As Object, e As EventArgs) Handles btnXemKH.Click
        frmKH.Show()
    End Sub

    Private Sub btnNhapthongtinSP_Button(sender As Object, e As EventArgs) Handles btnNhapthongtinSP.Click
        frmCapnhatsanpham.Show()
    End Sub

    Private Sub btnXemSP_Button(sender As Object, e As EventArgs) Handles btnXemSP.Click
        frmXemsanpham.Show()
    End Sub



    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub HóaĐơnToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HóaĐơnToolStripMenuItem.Click
        frmHoadon.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        frmHoadon.Show()
    End Sub

    Private Sub ThoátỨngDụngToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ThoátỨngDụngToolStripMenuItem.Click
        Close()
    End Sub
End Class
