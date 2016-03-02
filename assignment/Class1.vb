Imports System.Data.SqlClient
Imports System.Data
Public Class Class1
    Public Function Loadkhachang() As DataSet
        Dim chuoiketnoi As String = "workstation id=xuongtcps03100.mssql.somee.com;packet size=4096;user id=cuxuong000_SQLLogin_1;pwd=97tgood5o3;data source=xuongtcps03100.mssql.somee.com;persist security info=False;initial catalog=xuongtcps03100"
        Dim conn As SqlConnection = New SqlConnection(chuoiketnoi)
        Dim LoadKH As New SqlDataAdapter("select MAKH as 'Mã khách hàng' ,TENKH as 'Tên khách hàng', DIACHI as 'Địa chỉ', SDT as 'Số điện thoại', EMAIL from KHACHHANG", conn)
        Dim db As New DataSet
        conn.Open()
        LoadKH.Fill(db)
        conn.Close()
        Return db
    End Function
    Public Function Loadsanpham() As DataSet
        Dim chuoiketnoi As String = "workstation id=xuongtcps03100.mssql.somee.com;packet size=4096;user id=cuxuong000_SQLLogin_1;pwd=97tgood5o3;data source=xuongtcps03100.mssql.somee.com;persist security info=False;initial catalog=xuongtcps03100"
        Dim conn As SqlConnection = New SqlConnection(chuoiketnoi)
        Dim LoadSP As New SqlDataAdapter("select SANPHAM.MASP as 'Mã sản phẩm',SANPHAM.TENSP as 'Tên sản phẩm', LOAISANPHAM.MALOAI as 'Mã Loại', LOAISANPHAM.TENLOAI as 'Tên Loại',SANPHAM.SOLUONG as 'Số lượng' from SANPHAM inner join LOAISANPHAM on SANPHAM.MASP = LOAISANPHAM.MASP", conn)
        Dim db As New DataSet
        conn.Open()
        LoadSP.Fill(db)
        conn.Close()
        Return db
    End Function
    Public Function Loadhoadon() As DataSet
        Dim chuoiketnoi As String = "workstation id=xuongtcps03100.mssql.somee.com;packet size=4096;user id=cuxuong000_SQLLogin_1;pwd=97tgood5o3;data source=xuongtcps03100.mssql.somee.com;persist security info=False;initial catalog=xuongtcps03100"
        Dim conn As SqlConnection = New SqlConnection(chuoiketnoi)
        Dim LoadKH As New SqlDataAdapter("select MAHD as 'Mã hóa đơn' ,MAKH as 'Mã khách hàng', NGUOIXUATHD as 'Người xuất hóa đơn', NGAYHD as 'Ngày hóa đơn', TRIGIA as 'Trị giá' from HOADON", conn)
        Dim db As New DataSet
        conn.Open()
        LoadKH.Fill(db)
        conn.Close()
        Return db
    End Function
End Class
