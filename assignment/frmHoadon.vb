Imports System.Data.SqlClient
Imports System.Data.DataTable

Public Class frmHoadon
    Dim db As New DataTable
    Dim chuoiketnoi As String = "workstation id=xuongtcps03100.mssql.somee.com;packet size=4096;user id=cuxuong000_SQLLogin_1;pwd=97tgood5o3;data source=xuongtcps03100.mssql.somee.com;persist security info=False;initial catalog=xuongtcps03100"
    Dim conn As SqlConnection = New SqlConnection(chuoiketnoi)

    Private Sub frmHoadon_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim refesh As SqlDataAdapter = New SqlDataAdapter("select MAHD as 'Mã hóa đơn' ,MAKH as 'Tên khách hàng', NGUOIXUATHD as 'Người xuất hóa đơn', NGAYHD as 'Ngày hóa đơn', TRIGIA as 'Trị giá' from HOADON", conn)
        db.Clear()
        refesh.Fill(db)
        dgvHoadon.DataSource = db.DefaultView
    End Sub

    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        Dim connect As SqlConnection = New SqlConnection(chuoiketnoi)
        connect.Open()
        Dim xem As SqlDataAdapter = New SqlDataAdapter("select MAHD as 'Mã hóa đơn' ,MAKH as 'Tên khách hàng', NGUOIXUATHD as 'Người xuất hóa đơn', NGAYHD as 'Ngày hóa đơn', TRIGIA as 'Trị giá' from HOADON where MAHD=N'" & txtMahd.Text & "'", connect)
        Try
            If txtMahd.Text = "" Then
                MessageBox.Show("Bạn cần nhập mã hóa đơn", "Nhập thiếu", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
            Else
                db.Clear()
                dgvHoadon.DataSource = Nothing
                xem.Fill(db)
                If db.Rows.Count > 0 Then
                    dgvHoadon.DataSource = db.DefaultView
                    txtMahd.Text = Nothing
                    btnXoa.Enabled = True
                Else
                    MessageBox.Show("Không tìm thấy")
                    txtMahd.Text = Nothing
                End If
            End If
            connect.Close()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub dgvHoadon_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvHoadon.CellContentClick
        Dim click As Integer = dgvHoadon.CurrentCell.RowIndex
        txtMahd.Text = dgvHoadon.Item(0, click).Value
        txtMakh.Text = dgvHoadon.Item(1, click).Value
        txtNguoixuathd.Text = dgvHoadon.Item(2, click).Value
        txtNgayhd.Text = dgvHoadon.Item(3, click).Value
        txtTrigia.Text = dgvHoadon.Item(4, click).Value
    End Sub

    Private Sub btnThem_Click(sender As Object, e As EventArgs) Handles btnThem.Click
        Dim conn As SqlConnection = New SqlConnection(chuoiketnoi)
        Dim query As String = "insert into HOADON values(@MAHD,@MAKH,@NGUOIXUATHD,@NGAYHD,@TRIGIA)"
        Dim save As SqlCommand = New SqlCommand(query, conn)
        conn.Open()
        Try
            txtMahd.Focus()
            If txtMahd.Text = "" Then
                MessageBox.Show("Bạn chưa nhập mã hóa đơn", "Nhập thiếu", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
            Else
                txtMahd.Focus()
                If txtMakh.Text = "" Then
                    MessageBox.Show("Bạn chưa nhập mã khách hàng", "Nhập thiếu", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
                Else
                    txtMakh.Focus()
                    If txtNguoixuathd.Text = "" Then
                        MessageBox.Show("Bạn chưa nhập người xuất hóa đơn", "Nhập thiếu", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
                    Else
                        txtNguoixuathd.Focus()
                        If txtNgayhd.Text = "" Then
                            MessageBox.Show("Bạn chưa nhập ngày hóa đơn", "Nhập thiếu", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
                        Else
                            txtNgayhd.Focus()
                            If txtTrigia.Text = "" Then
                                MessageBox.Show("Bạn chưa nhập trị giá", "Nhập thiếu", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
                                txtTrigia.Focus()
                            Else
                                save.Parameters.AddWithValue("@MAHD", txtMahd.Text)
                                save.Parameters.AddWithValue("@MAKH", txtMakh.Text)
                                save.Parameters.AddWithValue("@NGUOIXUATHD", txtNguoixuathd.Text)
                                save.Parameters.AddWithValue("@NGAYHD", txtNgayhd.Text)
                                save.Parameters.AddWithValue("@TRIGIA", txtTrigia.Text)
                            save.ExecuteNonQuery()
                            MessageBox.Show("Lưu thành công")
                            'Sau khi nhập thành công, tự động làm mới các khung textbox, combox và date....
                                txtMahd.Text = Nothing
                                txtMakh.Text = Nothing
                                txtNguoixuathd.Text = Nothing
                                txtNgayhd.Text = Nothing
                                txtTrigia.Text = Nothing
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception  'Ngược lại báo lỗi
            MessageBox.Show("Không được trùng mã hóa đơn", "Lỗi", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error)
        End Try

        'Làm mới lại bảng sau khi lưu thành công
        Dim refesh As SqlDataAdapter = New SqlDataAdapter("select MAHD as 'Mã hóa đơn' ,MAKH as 'Tên khách hàng', NGUOIXUATHD as 'Người xuất hóa đơn', NGAYHD as 'Ngày hóa đơn', TRIGIA as 'Trị giá' from HOADON", conn)
        db.Clear()
        refesh.Fill(db)
        dgvHoadon.DataSource = db.DefaultView
    End Sub

    Private Sub btnXoa_Click(sender As Object, e As EventArgs) Handles btnXoa.Click
        Dim delquery As String = "delete from HOADON where MAHD=@MAHD"
        Dim delete As SqlCommand = New SqlCommand(delquery, conn)
        Dim resulft As DialogResult = MessageBox.Show("Bạn muốn xóa không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        conn.Open()
        Try
            If txtMahd.Text = "" Then
                MessageBox.Show("Bạn cần nhập mã hóa đơn", "Nhập thiếu", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
                txtMahd.Focus()
            Else
                If resulft = Windows.Forms.DialogResult.Yes Then
                    delete.Parameters.AddWithValue("@MAHD", txtMahd.Text)
                    delete.ExecuteNonQuery()
                    conn.Close()
                    MessageBox.Show("Xóa thành công")
                    'Sau khi xóa thành công, tự động làm mới các khung textbox, combox và date....
                    txtMahd.Text = Nothing
                    txtMakh.Text = Nothing
                    txtNguoixuathd.Text = Nothing
                    txtNgayhd.Text = Nothing
                    txtTrigia.Text = Nothing
                    txtMahd.Focus()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Nhập đúng mã hóa đơn cần xóa", "Lỗi", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error)
        End Try

        'làm mới bảng
        db.Clear()
        dgvHoadon.DataSource = db
        dgvHoadon.DataSource = Nothing
        LoadData()
    End Sub

    Private Sub LoadData()
        Dim conn As SqlConnection = New SqlConnection(chuoiketnoi)
        Dim load As SqlDataAdapter = New SqlDataAdapter("select MAHD as 'Mã hóa đơn' ,MAKH as 'Tên khách hàng', NGUOIXUATHD as 'Người xuất hóa đơn', NGAYHD as 'Ngày hóa đơn', TRIGIA as 'Trị giá' from HOADON", conn)
        conn.Open()
        load.Fill(db)
        dgvHoadon.DataSource = db.DefaultView
    End Sub

    Private Sub btnCapnhat_Click(sender As Object, e As EventArgs) Handles btnCapnhat.Click
        Dim conn As SqlConnection = New SqlConnection(chuoiketnoi)
        Dim updatequery As String = "update HOADON set MAHD=@MAHD, MAKH=@MAKH, NGUOIXUATHD=@NGUOIXUATHD, NGAYHD=@NGAYHD, TRIGIA=@TRIGIA where MAHD=@MAHD"
        Dim addupdate As SqlCommand = New SqlCommand(updatequery, conn)
        conn.Open()
        Try
            txtMahd.Focus()
            If txtMahd.Text = "" Then
                MessageBox.Show("Bạn chưa nhập mã hóa đơn", "Nhập thiếu", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
            Else
                txtMahd.Focus()
                If txtMakh.Text = "" Then
                    MessageBox.Show("Bạn chưa nhập mã khách hàng", "Nhập thiếu", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
                Else
                    txtMakh.Focus()
                    If txtNguoixuathd.Text = "" Then
                        MessageBox.Show("Bạn chưa nhập người xuất hóa đơn", "Nhập thiếu", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
                    Else
                        txtNguoixuathd.Focus()
                        If txtNgayhd.Text = "" Then
                            MessageBox.Show("Bạn chưa nhập ngày hóa đơn", "Nhập thiếu", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
                        Else
                            txtNgayhd.Focus()
                            If txtTrigia.Text = "" Then
                                MessageBox.Show("Bạn chưa nhập trị giá", "Nhập thiếu", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
                            Else
                                addupdate.Parameters.AddWithValue("@MAHD", txtMahd.Text)
                                addupdate.Parameters.AddWithValue("@MAKH", txtMakh.Text)
                                addupdate.Parameters.AddWithValue("@NGUOIXUATHD", txtNguoixuathd.Text)
                                addupdate.Parameters.AddWithValue("@NGAYHD", txtNgayhd.Text)
                                addupdate.Parameters.AddWithValue("@TRIGIA", txtTrigia.Text)
                                addupdate.ExecuteNonQuery()
                                conn.Close() 'đóng kết nối
                                MessageBox.Show("Cập nhật thành công")
                                txtMahd.Text = Nothing
                                txtMakh.Text = Nothing
                                txtNguoixuathd.Text = Nothing
                                txtNgayhd.Text = Nothing
                                txtTrigia.Text = Nothing
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Không thành công", "Lỗi", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error)
        End Try

        'Sau khi cập nhật xong sẽ tự làm mới lại bảng
        db.Clear()
        dgvHoadon.DataSource = db
        dgvHoadon.DataSource = Nothing
        LoadData()
    End Sub
End Class