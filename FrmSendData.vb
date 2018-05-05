Imports System.Text
Imports System.Net
Imports System
Imports System.Data
Imports System.Collections
Imports System.Configuration
Imports System.Threading
Imports System.Globalization
Imports System.Data.OleDb
Imports System.IO
Imports System.Data.SqlClient

Public Class FrmSendData

    Private Sub FrmSendData_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' text_line_ac_loss()
    End Sub
    Private Sub LineSendMessage_Click(sender As Object, e As EventArgs) Handles LineSendMessage.Click

        'If Date.Now.Hour.ToString = 14 And Date.Now.Minute.ToString = 7 Then
        'text_line_srtu()
        'text_line_frtu2()
        'text_line_frtu1()
        'text_line_frtu_percent()
        'End If


        'line_notify_ffg()
        text_line_sa_record_1()
        'text_line_sa_record_2()
        'text_line_sa_record_3()


        'text_line_frtu2_up()
        'text_line_frtu1_up()
        'text_line_frtu2_down()
        'text_line_frtu1_down()
        'text_line_frtu_percent()
        'text_line_frtu_percent_all()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try

       

        text_line_srtu()
        text_line_ac_loss()
        'text_line_frtu()
        line_notify_ffg()
            line_notify_ffg_s2()
            line_notify_sys_scada()
        If Date.Now.Hour.ToString = 8 And Date.Now.Minute.ToString = 0 Then

            text_line_frtu2_up()
            text_line_frtu1_up()
            text_line_frtu2_down()
            text_line_frtu1_down()
            text_line_frtu_percent()
            text_line_frtu_percent_all()
        End If
        If Date.Now.Hour.ToString = 16 And Date.Now.Minute.ToString = 0 Then
            text_line_sa_record_1()
            text_line_sa_record_2()
            text_line_sa_record_3()
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub text_line_sa_record_1()
        ''stickerPackageId=1&stickerId=1 //stick=https://devdocs.line.me/files/sticker_list.pdf
        ''&imageThumbnail=URL.png&imageFullsize=URL" //Photo
        ''Getline api token = https://notify-bot.line.me/my/
        ' Dim postData = String.Format("message=" & "สำหรับ Up Down งานสถานีรอสักพักนะค่ะ ขอเวานู๋ upgrade ตัวเองแพพ ")

        Dim str0 As String = " &stickerPackageId=2&stickerId=34"
        Dim str1 As String = "SA Record ผบอ.กบษ(น๓) บำรุงรักษาดังนี้" & vbCrLf
        ' Dim str2() As String = select_sarecord_1()

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= SA_System")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        'Dim Dr5 As DataRow

        Dim where_str As String = "SELECT [office],[op_id],[location],[date_operate],[operation],[remark],[status_work],[pmcm_id],[id_type_frtu],[date_update],[type_frtu],[dbname] FROM [SA_System].[sa].[View_sa_line] where office =  'ผบอ.กบษ.น.3' AND date_update >= '" & Date.Now.AddDays(-1).ToString("MM/dd/yyyy") & "' order by [type_frtu],date_operate asc"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "select_sarecord_1")
        Dt5 = Ds5.Tables("select_sarecord_1")

        Dim Str As String

        If Dt5.Rows.Count > 0 Then
            For i As Integer = 0 To Dt5.Rows.Count - 1

                Str = Dt5.Rows(i).Item(10) & " " & Dt5.Rows(i).Item(1) & vbCrLf & Dt5.Rows(i).Item(2) & " " & Date.Parse(Dt5.Rows(i).Item(3)).ToString("dd/MM/yyyy") & vbCrLf & Dt5.Rows(i).Item(4) & " " & Dt5.Rows(i).Item(5) & "  สถานะการปฏิบัติงาน : " & Dt5.Rows(i).Item(6)
                Str = Str & vbCrLf & "หมายเหตุ : " & Dt5.Rows(i).Item(5)
                If Mid(Dt5.Rows(i).Item(10).ToString, 1, 4) = "FRTU" Then
                    Dim dbname1 As String = Mid(Dt5.Rows(i).Item(11).ToString, 1, 5) & Mid(Dt5.Rows(i).Item(11).ToString, 7, 3) & "F"

                    Dim Ds As New DataSet
                    Dim Dt As DataTable
                    Dim where_str2 As String = "SELECT status,time_commu,dbname, desc_name, location,  type_frtu FROM scada.View_up_down_diff where dbname = '" & dbname1 & "' "
                    Dim Adpt As New SqlClient.SqlDataAdapter(where_str2, StrCon5)
                    Adpt.Fill(Ds, "select_sarecord_frtu")
                    Dt = Ds.Tables("select_sarecord_frtu")
                    If Dt.Rows.Count > 0 Then
                        Str = Str & vbCrLf & "สถานะจากระบบ SCADA : DB_name = " & Dt.Rows(0).Item(2) & " " & Dt.Rows(0).Item(0) & " " & Date.Parse(Dt.Rows(0).Item(1)).ToString("dd/MM/yyyy :hh:mm") & vbCrLf
                    Else
                        Str = Str & vbCrLf & "ค้นหาสถานะจากระบบ SCADA ไม่พบ : DB_name = " & Mid(Dt5.Rows(i).Item(11).ToString, 7, 3) & "F" & vbCrLf
                    End If
                Else
                    Str = Str & vbCrLf
                End If


                Dim Ds6 As New DataSet
                Dim Dt6 As DataTable
                Dim where_str6 As String = "SELECT [pmcm_id],[id],[Emp_id],[Names],[Position] FROM [SA_System].[sa].[View_name_name_list_pmcm_id] where pmcm_id = " & Dt5.Rows(i).Item(7)
                Dim Adpt6 As New SqlClient.SqlDataAdapter(where_str6, StrCon5)
                Adpt6.Fill(Ds6, "select_sarecord_6")
                Dt6 = Ds6.Tables("select_sarecord_6")
                Dim str6 As String = vbCrLf & "มีผู้ดำเดินการจำนวน " & Dt6.Rows.Count.ToString & " ท่าน คือ" & vbCrLf
                For ii As Integer = 0 To Dt6.Rows.Count - 1
                    str6 = str6 & Dt6.Rows(ii).Item(3).ToString & "  " & Dt6.Rows(ii).Item(4).ToString & vbCrLf

                Next

               






                Dim Ds7 As New DataSet
                Dim Dt7 As DataTable
                Dim where_str7 As String = "SELECT [damage_name],[Correction],[Cause] FROM [SA_System].[sa].[View_damage_list] where pmcm_id = " & Dt5.Rows(i).Item(7)
                Dim Adpt7 As New SqlClient.SqlDataAdapter(where_str7, StrCon5)
                Adpt7.Fill(Ds7, "select_sarecord_7")
                Dt7 = Ds7.Tables("select_sarecord_7")
                Dim str7 As String
                If Dt7.Rows.Count > 0 Then

                    str7 = vbCrLf & "มีอาการชำรุดจำนวน " & Dt7.Rows.Count.ToString & " อาการ คือ"
                    Dim damage_list As String
                    For ii As Integer = 0 To Dt7.Rows.Count - 1
                        damage_list = vbCrLf & ii + 1 & " : " & Dt7.Rows(ii).Item(0).ToString
                        damage_list = damage_list & vbCrLf & "วิธีการแก้ไข :" & Dt7.Rows(ii).Item(1).ToString
                        damage_list = damage_list & vbCrLf & "สาเหตุ :  " & Dt7.Rows(ii).Item(2).ToString & vbCrLf
                        str7 = str7 & damage_list
                    Next

                Else
                    str7 = vbCrLf & "มีอาการชำรุดจำนวน " & Dt7.Rows.Count.ToString & " อาการ" & vbCrLf

                End If



                Str = str1 & Str & str6 & str7





                to_line_sa(Str)
            Next


        Else

            Dim str_no As String = "วันที่ " & Date.Now.AddDays(-1).ToString("dd/MM/yyyy").ToString & " ผบอ.กบษ.(น๓) ไม่มีการลงข้อมูลใน Sa Record"
            to_line_sa(str_no)
        End If


        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''






       
    End Sub
    Private Sub text_line_sa_record_2()
        ''stickerPackageId=1&stickerId=1 //stick=https://devdocs.line.me/files/sticker_list.pdf
        ''&imageThumbnail=URL.png&imageFullsize=URL" //Photo
        ''Getline api token = https://notify-bot.line.me/my/
        ' Dim postData = String.Format("message=" & "สำหรับ Up Down งานสถานีรอสักพักนะค่ะ ขอเวานู๋ upgrade ตัวเองแพพ ")

        Dim str0 As String = " &stickerPackageId=2&stickerId=34"
        Dim str1 As String = "SA Record ผรล.กบษ(น๓) บำรุงรักษาดังนี้" & vbCrLf
        ' Dim str2() As String = select_sarecord_1()

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= SA_System")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        'Dim Dr5 As DataRow

        Dim where_str As String = "SELECT [office],[op_id],[location],[date_operate],[operation],[remark],[status_work],[pmcm_id],[id_type_frtu],[date_update],[type_frtu],[dbname] FROM [SA_System].[sa].[View_sa_line] where office =  'ผรล.กบษ.น.3' AND date_update >= '" & Date.Now.AddDays(-1).ToString("MM/dd/yyyy") & "' order by [type_frtu],date_operate asc"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "select_sarecord_3")
        Dt5 = Ds5.Tables("select_sarecord_3")

        Dim Str As String

        If Dt5.Rows.Count > 0 Then
            For i As Integer = 0 To Dt5.Rows.Count - 1

                Str = Dt5.Rows(i).Item(10) & " " & Dt5.Rows(i).Item(1) & vbCrLf & Dt5.Rows(i).Item(2) & " " & Date.Parse(Dt5.Rows(i).Item(3)).ToString("dd/MM/yyyy") & vbCrLf & Dt5.Rows(i).Item(4) & " " & Dt5.Rows(i).Item(5) & "  สถานะการปฏิบัติงาน : " & Dt5.Rows(i).Item(6)

                Dim Ds6 As New DataSet
                Dim Dt6 As DataTable

                Dim where_str6 As String = "SELECT [pmcm_id],[id],[Emp_id],[Names],[Position] FROM [SA_System].[sa].[View_name_name_list_pmcm_id] where pmcm_id = " & Dt5.Rows(i).Item(7)
                Dim Adpt6 As New SqlClient.SqlDataAdapter(where_str6, StrCon5)
                Adpt6.Fill(Ds6, "select_sarecord_6")
                Dt6 = Ds6.Tables("select_sarecord_6")
                Dim str6 As String = vbCrLf & "มีผู้ดำเดินการจำนวน " & Dt6.Rows.Count.ToString & " ท่าน คือ" & vbCrLf
                For ii As Integer = 0 To Dt6.Rows.Count - 1
                    str6 = str6 & Dt6.Rows(ii).Item(3).ToString & "  " & Dt6.Rows(ii).Item(4).ToString & vbCrLf

                Next
                Str = str1 & Str & str6
                to_line_sa(Str)
            Next

           
        Else
            Dim str_no As String = "วันที่ " & Date.Now.AddDays(-1).ToString("dd/MM/yyyy").ToString & " ผรล.กบษ.(น๓) ไม่มีการลงข้อมูลใน Sa Record"
            to_line_sa(str_no)

        End If


        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''







    End Sub
    Private Sub text_line_sa_record_3()
        ''stickerPackageId=1&stickerId=1 //stick=https://devdocs.line.me/files/sticker_list.pdf
        ''&imageThumbnail=URL.png&imageFullsize=URL" //Photo
        ''Getline api token = https://notify-bot.line.me/my/
        ' Dim postData = String.Format("message=" & "สำหรับ Up Down งานสถานีรอสักพักนะค่ะ ขอเวานู๋ upgrade ตัวเองแพพ ")

        Dim str0 As String = " &stickerPackageId=2&stickerId=34"
        Dim str1 As String = "SA Record ผบฟ.กบษ(น๓) บำรุงรักษาดังนี้" & vbCrLf
        ' Dim str2() As String = select_sarecord_1()

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= SA_System")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        'Dim Dr5 As DataRow

        Dim where_str As String = "SELECT [office],[op_id],[location],[date_operate],[operation],[remark],[status_work],[pmcm_id],[id_type_frtu],[date_update],[type_frtu],[dbname] FROM [SA_System].[sa].[View_sa_line] where office =  'ผบฟ.กบษ.น.3' AND date_update >= '" & Date.Now.AddDays(-1).ToString("MM/dd/yyyy") & "' order by [type_frtu],date_operate asc"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "select_sarecord_3")
        Dt5 = Ds5.Tables("select_sarecord_3")

        Dim Str As String

        If Dt5.Rows.Count > 0 Then
            For i As Integer = 0 To Dt5.Rows.Count - 1

                Str = Dt5.Rows(i).Item(10) & " " & Dt5.Rows(i).Item(1) & vbCrLf & Dt5.Rows(i).Item(2) & " " & Date.Parse(Dt5.Rows(i).Item(3)).ToString("dd/MM/yyyy") & vbCrLf & Dt5.Rows(i).Item(4) & " " & Dt5.Rows(i).Item(5) & "  สถานะการปฏิบัติงาน : " & Dt5.Rows(i).Item(6)

                Dim Ds6 As New DataSet
                Dim Dt6 As DataTable

                Dim where_str6 As String = "SELECT [pmcm_id],[id],[Emp_id],[Names],[Position] FROM [SA_System].[sa].[View_name_name_list_pmcm_id] where pmcm_id = " & Dt5.Rows(i).Item(7)
                Dim Adpt6 As New SqlClient.SqlDataAdapter(where_str6, StrCon5)
                Adpt6.Fill(Ds6, "select_sarecord_6")
                Dt6 = Ds6.Tables("select_sarecord_6")
                Dim str6 As String = vbCrLf & "มีผู้ดำเดินการจำนวน " & Dt6.Rows.Count.ToString & " ท่าน คือ" & vbCrLf
                For ii As Integer = 0 To Dt6.Rows.Count - 1
                    str6 = str6 & Dt6.Rows(ii).Item(3).ToString & "  " & Dt6.Rows(ii).Item(4).ToString & vbCrLf

                Next
                Str = str1 & Str & str6
                to_line_sa(Str)
            Next


        Else

            Dim str_no As String = "วันที่ " & Date.Now.AddDays(-1).ToString("dd/MM/yyyy").ToString & " ผบฟ.กบษ.(น๓) ไม่มีการลงข้อมูลใน Sa Record"
            to_line_sa(str_no)
        End If


        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''







    End Sub
    Private Sub to_line_sa(ByVal str_sub As String)
        Dim requestData = DirectCast(WebRequest.Create("https://notify-api.line.me/api/notify"), HttpWebRequest)

        Dim postData = String.Format("message=" & str_sub)
        Dim data = Encoding.UTF8.GetBytes(postData)
        requestData.Method = "POST"
        requestData.ContentType = "application/x-www-form-urlencoded"
        requestData.ContentLength = data.Length

        Dim StrCon55 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= staging")
        Dim Ds55 As New DataSet
        Dim Dt55 As DataTable
        Dim where_str55 As String = "SELECT token FROM line.line_token where group_name  = 'งานscada'"
        Dim Adpt55 As New SqlClient.SqlDataAdapter(where_str55, StrCon55)
        Adpt55.Fill(Ds55, "select_sarecord_1")
        Dt55 = Ds55.Tables("select_sarecord_1")
        requestData.Headers.Add("Authorization", Dt55.Rows(0).Item(0).ToString)

        Using stream = requestData.GetRequestStream()
            stream.Write(data, 0, data.Length)
        End Using
        Dim response = DirectCast(requestData.GetResponse(), HttpWebResponse)
        Dim responseString = New StreamReader(response.GetResponseStream()).ReadToEnd()

    End Sub


    Private Sub text_line_sa_record_22()
        ''stickerPackageId=1&stickerId=1 //stick=https://devdocs.line.me/files/sticker_list.pdf
        ''&imageThumbnail=URL.png&imageFullsize=URL" //Photo
        ''Getline api token = https://notify-bot.line.me/my/
        Dim requestData = DirectCast(WebRequest.Create("https://notify-api.line.me/api/notify"), HttpWebRequest)
        ' Dim postData = String.Format("message=" & "สำหรับ Up Down งานสถานีรอสักพักนะค่ะ ขอเวานู๋ upgrade ตัวเองแพพ ")

        Dim str As String = " &stickerPackageId=2&stickerId=34"
        Dim str1 As String = "SA Record ผรล.กบษ(น๓) บำรุงรักษาดังนี้" & vbCrLf
        Dim str2 As String = select_sarecord_22()
        Dim postData = String.Format("message=" & str1 & str2)
        Dim data = Encoding.UTF8.GetBytes(postData)
        requestData.Method = "POST"
        requestData.ContentType = "application/x-www-form-urlencoded"
        requestData.ContentLength = data.Length

        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= staging")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        Dim where_str As String = "SELECT token FROM line.line_token where group_name  = 'งานscada'"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "select_sarecord_1")
        Dt5 = Ds5.Tables("select_sarecord_1")
        requestData.Headers.Add("Authorization", Dt5.Rows(0).Item(0).ToString)

        Using stream = requestData.GetRequestStream()
            stream.Write(data, 0, data.Length)
        End Using
        Dim response = DirectCast(requestData.GetResponse(), HttpWebResponse)
        Dim responseString = New StreamReader(response.GetResponseStream()).ReadToEnd()

    End Sub
    Private Sub text_line_frtu_percent_all()
        ''stickerPackageId=1&stickerId=1 //stick=https://devdocs.line.me/files/sticker_list.pdf
        ''&imageThumbnail=URL.png&imageFullsize=URL" //Photo
        ''Getline api token = https://notify-bot.line.me/my/
        Dim requestData = DirectCast(WebRequest.Create("https://notify-api.line.me/api/notify"), HttpWebRequest)
        ' Dim postData = String.Format("message=" & "สำหรับ Up Down งานสถานีรอสักพักนะค่ะ ขอเวานู๋ upgrade ตัวเองแพพ ")

        Dim str As String = " &stickerPackageId=2&stickerId=34"
        Dim str1 As String = "วันที่ " & Date.Now.ToString("dd/MM/yyyy") & " มีเปอร์เซ็นความเพร้อมใช้งานดังนี้ " & vbCrLf
        Dim str2 As String = select_all_frtu_percent()
        Dim postData = String.Format("message=" & str1 & str2)
        Dim data = Encoding.UTF8.GetBytes(postData)
        requestData.Method = "POST"
        requestData.ContentType = "application/x-www-form-urlencoded"
        requestData.ContentLength = data.Length

        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= staging")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        Dim where_str As String = "SELECT token FROM line.line_token where group_name  = 'FRTU'"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "text_line_frtu_percent")
        Dt5 = Ds5.Tables("text_line_frtu_percent")
        requestData.Headers.Add("Authorization", Dt5.Rows(0).Item(0).ToString)

        Using stream = requestData.GetRequestStream()
            stream.Write(data, 0, data.Length)
        End Using
        Dim response = DirectCast(requestData.GetResponse(), HttpWebResponse)
        Dim responseString = New StreamReader(response.GetResponseStream()).ReadToEnd()

    End Sub
    Private Sub text_line_frtu_percent()
        ''stickerPackageId=1&stickerId=1 //stick=https://devdocs.line.me/files/sticker_list.pdf
        ''&imageThumbnail=URL.png&imageFullsize=URL" //Photo
        ''Getline api token = https://notify-bot.line.me/my/
        Dim requestData = DirectCast(WebRequest.Create("https://notify-api.line.me/api/notify"), HttpWebRequest)
        ' Dim postData = String.Format("message=" & "สำหรับ Up Down งานสถานีรอสักพักนะค่ะ ขอเวานู๋ upgrade ตัวเองแพพ ")

        Dim str As String = " &stickerPackageId=2&stickerId=34"
        Dim str1 As String = "วันที่ " & Date.Now.ToString("dd/MM/yyyy") & " มีเปอร์เซ็นความเพร้อมใช้งานดังนี้ " & vbCrLf
        Dim str2 As String = select_down_frtu_percent()
        Dim postData = String.Format("message=" & str1 & str2)
        Dim data = Encoding.UTF8.GetBytes(postData)
        requestData.Method = "POST"
        requestData.ContentType = "application/x-www-form-urlencoded"
        requestData.ContentLength = data.Length

        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= staging")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        Dim where_str As String = "SELECT token FROM line.line_token where group_name  = 'FRTU'"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "text_line_frtu_percent")
        Dt5 = Ds5.Tables("text_line_frtu_percent")
        requestData.Headers.Add("Authorization", Dt5.Rows(0).Item(0).ToString)

        Using stream = requestData.GetRequestStream()
            stream.Write(data, 0, data.Length)
        End Using
        Dim response = DirectCast(requestData.GetResponse(), HttpWebResponse)
        Dim responseString = New StreamReader(response.GetResponseStream()).ReadToEnd()

    End Sub
    Private Sub text_line_frtu2_up()
        ''stickerPackageId=1&stickerId=1 //stick=https://devdocs.line.me/files/sticker_list.pdf
        ''&imageThumbnail=URL.png&imageFullsize=URL" //Photo
        ''Getline api token = https://notify-bot.line.me/my/
        Dim requestData = DirectCast(WebRequest.Create("https://notify-api.line.me/api/notify"), HttpWebRequest)
        ' Dim postData = String.Format("message=" & "สำหรับ Up Down งานสถานีรอสักพักนะค่ะ ขอเวานู๋ upgrade ตัวเองแพพ ")

        Dim str As String = " &stickerPackageId=2&stickerId=34"
        Dim str1 As String = "วันที่ " & Date.Now.AddDays(-2).ToString("dd/MM/yyyy") & " มี FRTU Up ดังนี้ค่ะ" & vbCrLf
        Dim str2 As String = select_down_frtu_2_up()
        Dim postData = String.Format("message=" & str1 & str2)
        Dim data = Encoding.UTF8.GetBytes(postData)
        requestData.Method = "POST"
        requestData.ContentType = "application/x-www-form-urlencoded"
        requestData.ContentLength = data.Length

        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= staging")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        Dim where_str As String = "SELECT token FROM line.line_token where group_name  = 'FRTU'"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "text_line_ac_loss")
        Dt5 = Ds5.Tables("text_line_ac_loss")
        requestData.Headers.Add("Authorization", Dt5.Rows(0).Item(0).ToString)

        Using stream = requestData.GetRequestStream()
            stream.Write(data, 0, data.Length)
        End Using
        Dim response = DirectCast(requestData.GetResponse(), HttpWebResponse)
        Dim responseString = New StreamReader(response.GetResponseStream()).ReadToEnd()

    End Sub
    Private Sub text_line_frtu2_down()
        ''stickerPackageId=1&stickerId=1 //stick=https://devdocs.line.me/files/sticker_list.pdf
        ''&imageThumbnail=URL.png&imageFullsize=URL" //Photo
        ''Getline api token = https://notify-bot.line.me/my/
        Dim requestData = DirectCast(WebRequest.Create("https://notify-api.line.me/api/notify"), HttpWebRequest)
        ' Dim postData = String.Format("message=" & "สำหรับ Up Down งานสถานีรอสักพักนะค่ะ ขอเวานู๋ upgrade ตัวเองแพพ ")

        Dim str As String = " &stickerPackageId=2&stickerId=34"
        Dim str1 As String = "วันที่ " & Date.Now.AddDays(-2).ToString("dd/MM/yyyy") & " มี FRTU Down ดังนี้ค่ะ" & vbCrLf
        Dim str2 As String = select_down_frtu_2_down()
        Dim postData = String.Format("message=" & str1 & str2)
        Dim data = Encoding.UTF8.GetBytes(postData)
        requestData.Method = "POST"
        requestData.ContentType = "application/x-www-form-urlencoded"
        requestData.ContentLength = data.Length

        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= staging")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        Dim where_str As String = "SELECT token FROM line.line_token where group_name  = 'FRTU'"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "text_line_ac_loss")
        Dt5 = Ds5.Tables("text_line_ac_loss")
        requestData.Headers.Add("Authorization", Dt5.Rows(0).Item(0).ToString)

        Using stream = requestData.GetRequestStream()
            stream.Write(data, 0, data.Length)
        End Using
        Dim response = DirectCast(requestData.GetResponse(), HttpWebResponse)
        Dim responseString = New StreamReader(response.GetResponseStream()).ReadToEnd()

    End Sub
    Private Sub text_line_frtu1_up()
        ''stickerPackageId=1&stickerId=1 //stick=https://devdocs.line.me/files/sticker_list.pdf
        ''&imageThumbnail=URL.png&imageFullsize=URL" //Photo
        ''Getline api token = https://notify-bot.line.me/my/
        Dim requestData = DirectCast(WebRequest.Create("https://notify-api.line.me/api/notify"), HttpWebRequest)
        ' Dim postData = String.Format("message=" & "สำหรับ Up Down งานสถานีรอสักพักนะค่ะ ขอเวานู๋ upgrade ตัวเองแพพ ")

        Dim str As String = " &stickerPackageId=2&stickerId=34"
        Dim str1 As String = "และเมื่อวานนี้ วันที่ " & Date.Now.AddDays(-1).ToString("dd/MM/yyyy") & " มี FRTU Up ดังนี้ค่ะ" & vbCrLf
        Dim str2 As String = select_down_frtu_1_up()
        Dim postData = String.Format("message=" & str1 & str2)
        Dim data = Encoding.UTF8.GetBytes(postData)
        requestData.Method = "POST"
        requestData.ContentType = "application/x-www-form-urlencoded"
        requestData.ContentLength = data.Length



        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= staging")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        Dim where_str As String = "SELECT token FROM line.line_token where group_name  = 'FRTU'"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "text_line_ac_loss")
        Dt5 = Ds5.Tables("text_line_ac_loss")
        requestData.Headers.Add("Authorization", Dt5.Rows(0).Item(0).ToString)



        'requestData.Headers.Add("Authorization", "Bearer QXFUKokFhoxXxdSKJKoB43WKaOwvWpxjEySN67gcZ1x") 'งานscada
        'requestData.Headers.Add("Authorization", "Bearer d4q1dKt6OfE5aLN1TtkORtK1tFbqzG8eOjwMRtR7pZa") 'ate

        'requestData.Headers.Add("Authorization", "Bearer Z8z0BojdHdBV1Slov6S2fbvf8f2snWnqDdXLwq47cit") 'ผรศ

        'requestData.Headers.Add("Authorization", "Bearer RIdDz9aS1wB3rfB706crpiUwozZuhtoatrcU6P7DkCf") 'test
        Using stream = requestData.GetRequestStream()
            stream.Write(data, 0, data.Length)
        End Using
        Dim response = DirectCast(requestData.GetResponse(), HttpWebResponse)
        Dim responseString = New StreamReader(response.GetResponseStream()).ReadToEnd()
    End Sub
    Private Sub text_line_frtu1_down()
        ''stickerPackageId=1&stickerId=1 //stick=https://devdocs.line.me/files/sticker_list.pdf
        ''&imageThumbnail=URL.png&imageFullsize=URL" //Photo
        ''Getline api token = https://notify-bot.line.me/my/
        Dim requestData = DirectCast(WebRequest.Create("https://notify-api.line.me/api/notify"), HttpWebRequest)
        ' Dim postData = String.Format("message=" & "สำหรับ Up Down งานสถานีรอสักพักนะค่ะ ขอเวานู๋ upgrade ตัวเองแพพ ")

        Dim str As String = " &stickerPackageId=2&stickerId=34"
        Dim str1 As String = "และเมื่อวานนี้ วันที่ " & Date.Now.AddDays(-1).ToString("dd/MM/yyyy") & " มี FRTU Down ดังนี้ค่ะ" & vbCrLf
        Dim str2 As String = select_down_frtu_1_Down()
        Dim postData = String.Format("message=" & str1 & str2)
        Dim data = Encoding.UTF8.GetBytes(postData)
        requestData.Method = "POST"
        requestData.ContentType = "application/x-www-form-urlencoded"
        requestData.ContentLength = data.Length



        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= staging")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        Dim where_str As String = "SELECT token FROM line.line_token where group_name  = 'FRTU'"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "text_line_ac_loss")
        Dt5 = Ds5.Tables("text_line_ac_loss")
        requestData.Headers.Add("Authorization", Dt5.Rows(0).Item(0).ToString)



        'requestData.Headers.Add("Authorization", "Bearer QXFUKokFhoxXxdSKJKoB43WKaOwvWpxjEySN67gcZ1x") 'งานscada
        'requestData.Headers.Add("Authorization", "Bearer d4q1dKt6OfE5aLN1TtkORtK1tFbqzG8eOjwMRtR7pZa") 'ate

        'requestData.Headers.Add("Authorization", "Bearer Z8z0BojdHdBV1Slov6S2fbvf8f2snWnqDdXLwq47cit") 'ผรศ

        'requestData.Headers.Add("Authorization", "Bearer RIdDz9aS1wB3rfB706crpiUwozZuhtoatrcU6P7DkCf") 'test
        Using stream = requestData.GetRequestStream()
            stream.Write(data, 0, data.Length)
        End Using
        Dim response = DirectCast(requestData.GetResponse(), HttpWebResponse)
        Dim responseString = New StreamReader(response.GetResponseStream()).ReadToEnd()
    End Sub
    Private Sub text_line_ac_loss()
        Dim requestData = DirectCast(WebRequest.Create("https://notify-api.line.me/api/notify"), HttpWebRequest)
        Dim str As String = " &stickerPackageId=1&stickerId=139"
        Dim str2 As String = ac_loss()
        If str2 <> "0" Then
            Dim postData = String.Format("message=" & vbCrLf & str2)
            Dim data = Encoding.UTF8.GetBytes(postData)
            requestData.Method = "POST"
            requestData.ContentType = "application/x-www-form-urlencoded"
            requestData.ContentLength = data.Length
            Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= staging")
            Dim Ds5 As New DataSet
            Dim Dt5 As DataTable
            Dim where_str As String = "SELECT token FROM line.line_token where group_name  = 'งานสนับสนุนการจ่ายไฟ'"
            Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
            Adpt5.Fill(Ds5, "text_line_ac_loss")
            Dt5 = Ds5.Tables("text_line_ac_loss")
            requestData.Headers.Add("Authorization", Dt5.Rows(0).Item(0).ToString)
            Using stream = requestData.GetRequestStream()
                stream.Write(data, 0, data.Length)
            End Using
            Dim response = DirectCast(requestData.GetResponse(), HttpWebResponse)
            Dim responseString = New StreamReader(response.GetResponseStream()).ReadToEnd()
        End If
    End Sub
    Private Sub text_line_srtu()
        ''stickerPackageId=1&stickerId=1 //stick=https://devdocs.line.me/files/sticker_list.pdf
        ''&imageThumbnail=URL.png&imageFullsize=URL" //Photo
        ''Getline api token = https://notify-bot.line.me/my/
        Dim requestData = DirectCast(WebRequest.Create("https://notify-api.line.me/api/notify"), HttpWebRequest)
        ' Dim postData = String.Format("message=" & "สำหรับ Up Down งานสถานีรอสักพักนะค่ะ ขอเวานู๋ upgrade ตัวเองแพพ ")

        Dim str As String = " &stickerPackageId=1&stickerId=139"
        'Dim str1 As String = "สวัสดีค่ะเมื่อสักครู่มีเหตุการณ์เกี่ยวกับระบบสื่อสารของสถานีไฟฟ้าดังนี้ค่ะ" & vbCrLf
        Dim str2 As String = select_down_srtu()

        If str2 <> "0" Then
            Dim postData = String.Format("message=" & vbCrLf & str2)
            Dim data = Encoding.UTF8.GetBytes(postData)
            requestData.Method = "POST"
            requestData.ContentType = "application/x-www-form-urlencoded"
            requestData.ContentLength = data.Length
            Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= staging")
            Dim Ds5 As New DataSet
            Dim Dt5 As DataTable
            'Dim Dr5 As DataRow

            Dim where_str As String = "SELECT token FROM line.line_token where group_name  = 'งานscada'"
            Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
            Adpt5.Fill(Ds5, "select_down_frtu")
            Dt5 = Ds5.Tables("select_down_frtu")

            requestData.Headers.Add("Authorization", Dt5.Rows(0).Item(0).ToString) 'งานscada
            'requestData.Headers.Add("Authorization", "Bearer QXFUKokFhoxXxdSKJKoB43WKaOwvWpxjEySN67gcZ1x") 'งานscada
            'requestData.Headers.Add("Authorization", "Bearer d4q1dKt6OfE5aLN1TtkORtK1tFbqzG8eOjwMRtR7pZa") 'ate

            'requestData.Headers.Add("Authorization", "Bearer Z8z0BojdHdBV1Slov6S2fbvf8f2snWnqDdXLwq47cit") 'ผรศ

            'requestData.Headers.Add("Authorization", "Bearer RIdDz9aS1wB3rfB706crpiUwozZuhtoatrcU6P7DkCf") 'test

            'requestData.Headers.Add("Authorization", "Bearer  WYfndMvSwPP7BZLUB0tyMyPDmRthqRr9Vnbn3EeO4Uz") 'ทีม F
            Using stream = requestData.GetRequestStream()
                stream.Write(data, 0, data.Length)
            End Using
            Dim response = DirectCast(requestData.GetResponse(), HttpWebResponse)
            Dim responseString = New StreamReader(response.GetResponseStream()).ReadToEnd()
        End If

    End Sub
    Private Sub text_line_frtu()
        ''stickerPackageId=1&stickerId=1 //stick=https://devdocs.line.me/files/sticker_list.pdf
        ''&imageThumbnail=URL.png&imageFullsize=URL" //Photo
        ''Getline api token = https://notify-bot.line.me/my/
        Dim requestData = DirectCast(WebRequest.Create("https://notify-api.line.me/api/notify"), HttpWebRequest)

        Dim str As String = " &stickerPackageId=1&stickerId=139"
        Dim str1 As String = "สวัสดีค่ะเมื่อสักครู่มีเหตุการณ์เกี่ยวกับระบบสื่อสารของ FRTU ดังนี้ค่ะ" & vbCrLf
        Dim str2 As String = select_down_frtu1()

        If str2 <> "0" Then
            'Dim postData = String.Format("message=" & "&Test 1")

            Dim postData = String.Format("message=" & str2)
            Dim data = Encoding.UTF8.GetBytes(postData)
            requestData.Method = "POST"
            requestData.ContentType = "application/x-www-form-urlencoded"
            requestData.ContentLength = data.Length
            requestData.Headers.Add("Authorization", "Bearer WYfndMvSwPP7BZLUB0tyMyPDmRthqRr9Vnbn3EeO4Uz") 'ทีม F
            Using stream = requestData.GetRequestStream()
                stream.Write(data, 0, data.Length)
            End Using
            Dim response = DirectCast(requestData.GetResponse(), HttpWebResponse)
            Dim responseString = New StreamReader(response.GetResponseStream()).ReadToEnd()
        End If

    End Sub
    Function select_all_frtu_percent()
        Try

       
        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= SA_System")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        'Dim Dr5 As DataRow
        Dim str As String = ""
        Dim where_str As String = "SELECT * FROM [SA_System].[scada].[View_percent_commu_all] order by status DESC"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "select_all_frtu_percent")
        Dt5 = Ds5.Tables("select_all_frtu_percent")

        str = "SCADA have " & Dt5.Rows(0).Item(1) & " FRTU " & vbCrLf
        str = str & "Up      is " & Dt5.Rows(0).Item(0) & " = " & Dt5.Rows(0).Item(2) & " percent" & vbCrLf
        str = str & "Down    is " & Dt5.Rows(1).Item(0) & " = " & Dt5.Rows(1).Item(2) & " percent" & vbCrLf
        str = str & "Disable is " & Dt5.Rows(2).Item(0) & " =" & Dt5.Rows(2).Item(2) & " percent" & vbCrLf




        Return str
        Catch ex As Exception

        End Try
    End Function
    Function select_down_frtu_percent()
        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= SA_System")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        'Dim Dr5 As DataRow
        Dim str As String = ""
        Dim where_str As String = "SELECT * FROM [SA_System].[scada].[View_percent_commu_rcs] where status = 'Up'"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "select_down_frtu_percent")
        Dt5 = Ds5.Tables("select_down_frtu_percent")

        str = "RCS " & Dt5.Rows(0).Item(2) & " %" & vbCrLf

        where_str = "SELECT * FROM [SA_System].[scada].[View_percent_commu_rec] where status = 'Up'"
        Dim Adpt6 As New SqlClient.SqlDataAdapter(where_str, StrCon5)

        Adpt6.Fill(Ds5, "select_down_frtu_percent1")
        Dt5 = Ds5.Tables("select_down_frtu_percent1")
        str = str & "Recloser " & Dt5.Rows(0).Item(2) & " %" & vbCrLf


        where_str = "SELECT * FROM [SA_System].[scada].[View_percent_commu_avr] where status = 'Up'"
        Dim Adpt7 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt7.Fill(Ds5, "select_down_frtu_percent2")
        Dt5 = Ds5.Tables("select_down_frtu_percent2")
        str = str & "AVR " & Dt5.Rows(0).Item(2) & " %" & vbCrLf


        Return str

    End Function

    Function select_down_frtu_2_up()
        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= SA_System")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        'Dim Dr5 As DataRow
        Dim str As String = ""
        Dim where_str As String = "SELECT time_commu,dbname, desc_name, location, status, type_frtu FROM scada.View_up_down_diff where time_commu >= '" & Date.Now.AddDays(-2).ToString("MM/dd/yyyy") & "' AND time_commu <= '" & Date.Now.AddDays(-1).ToString("MM/dd/yyyy") & "' AND status = 'Up' order by time_commu asc"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "select_down_frtu")
        Dt5 = Ds5.Tables("select_down_frtu")

        If Dt5.Rows.Count > 0 Then
            For i As Integer = 0 To Dt5.Rows.Count - 1

                str = str & Date.Parse(Dt5.Rows(i).Item(0)).ToString("HH:mm:ss") & " " & Dt5.Rows(i).Item(1) & " " & Dt5.Rows(i).Item(2) & " " & Dt5.Rows(i).Item(3) & vbCrLf

            Next
        Else
            str = Not Nothing
        End If

        Return str

    End Function

    Function select_down_frtu_2_down()
        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= SA_System")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        'Dim Dr5 As DataRow
        Dim str As String = ""
        Dim where_str As String = "SELECT time_commu,dbname, desc_name, location, status, type_frtu FROM scada.View_up_down_diff where time_commu >= '" & Date.Now.AddDays(-2).ToString("MM/dd/yyyy") & "' AND time_commu <= '" & Date.Now.AddDays(-1).ToString("MM/dd/yyyy") & "' AND status = 'Down' order by time_commu asc"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "select_down_frtu")
        Dt5 = Ds5.Tables("select_down_frtu")

        If Dt5.Rows.Count > 0 Then
            For i As Integer = 0 To Dt5.Rows.Count - 1

                str = str & Date.Parse(Dt5.Rows(i).Item(0)).ToString("HH:mm:ss") & " " & Dt5.Rows(i).Item(1) & " " & Dt5.Rows(i).Item(2) & " " & Dt5.Rows(i).Item(3) & vbCrLf

            Next
        Else
            str = Not Nothing
        End If

        Return str

    End Function

    Function select_down_frtu_1_up()
        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= SA_System")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        'Dim Dr5 As DataRow
        Dim str As String = ""
        Dim where_str As String = "SELECT time_commu,dbname, desc_name, location, status, type_frtu FROM scada.View_up_down_diff where time_commu >= '" & Date.Now.AddDays(-1).ToString("MM/dd/yyyy") & "' AND time_commu <= '" & Date.Now.AddHours(-1).ToString("MM/dd/yyyy") & "' AND status = 'Up' order by time_commu asc"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "select_down_frtu")
        Dt5 = Ds5.Tables("select_down_frtu")

        If Dt5.Rows.Count > 0 Then
            For i As Integer = 0 To Dt5.Rows.Count - 1

                str = str & Date.Parse(Dt5.Rows(i).Item(0)).ToString("HH:mm:ss") & " " & Dt5.Rows(i).Item(1) & " " & Dt5.Rows(i).Item(2) & " " & Dt5.Rows(i).Item(3) & vbCrLf
            Next
        Else
            str = "0"
        End If

        Return str

    End Function
    Function select_down_frtu_1_Down()
        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= SA_System")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        'Dim Dr5 As DataRow
        Dim str As String = ""
        Dim where_str As String = "SELECT time_commu,dbname, desc_name, location, status, type_frtu FROM scada.View_up_down_diff where time_commu >= '" & Date.Now.AddDays(-1).ToString("MM/dd/yyyy") & "' AND time_commu <= '" & Date.Now.AddHours(-1).ToString("MM/dd/yyyy") & "' AND status = 'Down' order by time_commu asc"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "select_down_frtu")
        Dt5 = Ds5.Tables("select_down_frtu")

        If Dt5.Rows.Count > 0 Then
            For i As Integer = 0 To Dt5.Rows.Count - 1

                str = str & Date.Parse(Dt5.Rows(i).Item(0)).ToString("HH:mm:ss") & " " & Dt5.Rows(i).Item(1) & " " & Dt5.Rows(i).Item(2) & " " & Dt5.Rows(i).Item(3) & vbCrLf
            Next
        Else
            str = "0"
        End If

        Return str

    End Function
    Function select_sarecord_1()
        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= SA_System")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        'Dim Dr5 As DataRow

        Dim where_str As String = "SELECT [office],[op_id],[location],[date_operate],[operation],[remark],[status_work],[pmcm_id],[id_type_frtu],[date_update],[type_frtu],[dbname] FROM [SA_System].[sa].[View_sa_line] where office =  'ผบอ.กบษ.น.3' AND date_update >= '" & Date.Now.AddDays(-2).ToString("MM/dd/yyyy") & "' order by [type_frtu],date_operate asc"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "select_sarecord_1")
        Dt5 = Ds5.Tables("select_sarecord_1")


        If Dt5.Rows.Count > 0 Then

            For i As Integer = 0 To Dt5.Rows.Count - 1
                Dim Str(Dt5.Rows.Count - 1) As String
                Str(i) = Dt5.Rows(i).Item(10) & " " & Dt5.Rows(i).Item(1) & " " & Dt5.Rows(i).Item(2) & " " & Date.Parse(Dt5.Rows(i).Item(3)).ToString("dd/MM/yyyy") & " " & Dt5.Rows(i).Item(4) & " " & Dt5.Rows(i).Item(5) & "  สถานะการปฏิบัติงาน : " & Dt5.Rows(i).Item(6)

                If Mid(Dt5.Rows(i).Item(10).ToString, 1, 4) = "FRTU" Then
                    Dim dbname1 As String = Mid(Dt5.Rows(i).Item(11).ToString, 1, 5) & Mid(Dt5.Rows(i).Item(11).ToString, 7, 3) & "F"

                    Dim Ds As New DataSet
                    Dim Dt As DataTable
                    Dim where_str2 As String = "SELECT status,time_commu,dbname, desc_name, location,  type_frtu FROM scada.View_up_down_diff where dbname = '" & dbname1 & "' "
                    Dim Adpt As New SqlClient.SqlDataAdapter(where_str2, StrCon5)
                    Adpt.Fill(Ds, "select_sarecord_frtu")
                    Dt = Ds.Tables("select_sarecord_frtu")
                    If Dt.Rows.Count > 0 Then
                        Str(i) = Str(i) & vbCrLf & "สถานะจากระบบ SCADA : DB_name = " & Dt.Rows(0).Item(2) & " " & Dt.Rows(0).Item(0) & " " & Date.Parse(Dt.Rows(0).Item(1)).ToString("dd/MM/yyyy :hh:mm") & vbCrLf
                    Else
                        Str(i) = Str(i) & "ไม่มีสถานะจากระบบ SCADA : DB_name = " & Mid(Dt5.Rows(i).Item(11).ToString, 7, 3) & "F"
                    End If
                Else
                    Str(i) = Str(i) & vbCrLf
                End If
                Dim Ds6 As New DataSet
                Dim Dt6 As DataTable

                Dim where_str6 As String = "SELECT [pmcm_id],[id],[Emp_id],[Names],[Position] FROM [SA_System].[sa].[View_name_name_list_pmcm_id] where pmcm_id = " & Dt5.Rows(i).Item(7)
                Dim Adpt6 As New SqlClient.SqlDataAdapter(where_str6, StrCon5)
                Adpt6.Fill(Ds6, "select_sarecord_6")
                Dt6 = Ds6.Tables("select_sarecord_6")
                Dim str6 As String = "มีผู้ดำเดินการจำนวน " & Dt6.Rows.Count.ToString & " ท่าน คือ" & vbCrLf
                For ii As Integer = 0 To Dt6.Rows.Count - 1
                    str6 = str6 & Dt6.Rows(ii).Item(3).ToString & "  " & Dt6.Rows(ii).Item(4).ToString & vbCrLf

                Next
                Str(i) = Str(i) & str6
            Next
            Return Str(10)
        Else
            Return 0
        End If



    End Function
    Function select_sarecord_22()
        Try

     
        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= SA_System")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        'Dim Dr5 As DataRow
        Dim str As String = ""
        Dim where_str As String = "SELECT [office],[op_id],[location],[date_operate],[operation],[remark],[status_work],[pmcm_id],[id_type_frtu],[date_update],[type_frtu],[dbname] FROM [SA_System].[sa].[View_sa_line] where office =  'ผรล.กบษ.น.3' AND date_update >= '" & Date.Now.AddDays(-1).ToString("MM/dd/yyyy") & "' order by office,date_operate asc"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "select_sarecord_1")
        Dt5 = Ds5.Tables("select_sarecord_1")

        If Dt5.Rows.Count > 0 Then
            For i As Integer = 0 To Dt5.Rows.Count - 1

                str = str & " " & Dt5.Rows(i).Item(10) & " " & Dt5.Rows(i).Item(1) & " " & Dt5.Rows(i).Item(2) & " " & Date.Parse(Dt5.Rows(i).Item(3)).ToString("dd/MM/yyyy") & " " & Dt5.Rows(i).Item(4) & " " & Dt5.Rows(i).Item(5) & vbCrLf & "สถานะการปฏิบัตืงาน : " & Dt5.Rows(i).Item(6)

                If Mid(Dt5.Rows(i).Item(10).ToString, 0, 4) = "FRTU" Then
                    Dim dbname1 As String = Mid(Dt5.Rows(i).Item(11).ToString, 0, 5) & Mid(Dt5.Rows(i).Item(11).ToString, 6, 3) & "F"

                    Dim Ds As New DataSet
                    Dim Dt As DataTable
                    Dim where_str2 As String = "SELECT status,time_commu,dbname, desc_name, location,  type_frtu FROM scada.View_up_down_diff where dbname = '" & dbname1 & "' "
                    Dim Adpt As New SqlClient.SqlDataAdapter(where_str, StrCon5)
                    Adpt.Fill(Ds, "select_sarecord_frtu")
                    Dt = Ds.Tables("select_sarecord_frtu")
                    str = str & vbCrLf & "สถานะจากระบบ SCADA : " & Dt.Rows(0).Item(0) & " " & Dt.Rows(0).Item(1).ToString("dd/MM/yyyy") & vbCrLf & vbCrLf
                Else
                    str = str & vbCrLf
                End If

                Dim Ds6 As New DataSet
                Dim Dt6 As DataTable

                Dim where_str6 As String = "SELECT [pmcm_id],[id],[Emp_id],[Names],[Position] FROM [SA_System].[sa].[View_name_name_list_pmcm_id] where pmcm_id = " & Dt5.Rows(i).Item(7)
                Dim Adpt6 As New SqlClient.SqlDataAdapter(where_str6, StrCon5)
                Adpt6.Fill(Ds6, "select_sarecord_6")
                Dt6 = Ds6.Tables("select_sarecord_6")
                Dim str6 As String = "มีผู้ดำเดินการจำนวน " & Dt6.Rows.Count.ToString & " ท่าน คือ" & vbCrLf
                For ii As Integer = 0 To Dt6.Rows.Count - 1
                    str6 = str6 & Dt6.Rows(ii).Item(3).ToString & "  " & Dt6.Rows(ii).Item(4).ToString & vbCrLf

                Next
                str = str & str6
            Next
        Else
            str = "0"
        End If

        Return str
        Catch ex As Exception

        End Try
    End Function
    Function ac_loss()
        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= SA_System")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        Dim objCmd As SqlCommand
        Dim strSQL As String
        'Dim Dr5 As DataRow
        Dim str As String = ""
        Dim where_str As String = "SELECT  [event] FROM [SA_System].[scada].[scada_line] where id_num = 3 order by time_commu asc"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "ac_loss")
        Dt5 = Ds5.Tables("ac_loss")

        If Dt5.Rows.Count > 0 Then
            For i As Integer = 0 To Dt5.Rows.Count - 1

                str = str & Dt5.Rows(i).Item(0).ToString & vbCrLf
                strSQL = "DELETE FROM [SA_System].[scada].[scada_line]  where event = '" & Dt5.Rows(i).Item(0).ToString & "'"
                objCmd = New SqlCommand(strSQL, StrCon5)
                StrCon5.Open()
                objCmd.ExecuteNonQuery()
                StrCon5.Close()
            Next
        Else

            str = "0"
        End If

        Return str

    End Function

    Function line_ffg1()
        Try


        
        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= SA_System")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        Dim objCmd As SqlCommand
        Dim strSQL As String
        'Dim Dr5 As DataRow
        Dim str As String = ""
        Dim where_str As String = "SELECT [event],[office_pea] FROM [SA_System].[scada].[View_line_ffg] order by time_commu asc"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "line_ffg1")
        Dt5 = Ds5.Tables("line_ffg1")

        If Dt5.Rows.Count > 0 Then
            For i As Integer = 0 To Dt5.Rows.Count - 1

                str = str & Dt5.Rows(i).Item(0).ToString & vbCrLf
                strSQL = "DELETE FROM [SA_System].[scada].[scada_line]  where event = '" & Dt5.Rows(i).Item(0).ToString & "'"
                objCmd = New SqlCommand(strSQL, StrCon5)
                StrCon5.Open()
                objCmd.ExecuteNonQuery()
                StrCon5.Close()
            Next
        Else

            str = "0"
        End If

        Return str
        Catch ex As Exception

        End Try
    End Function
    Sub line_notify_ffg()
        Try


            Dim request = DirectCast(WebRequest.Create("https://notify-api.line.me/api/notify"), HttpWebRequest)
            ' Dim Str As String = line_ffg1()

            Dim StrCon As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= SA_System")
            Dim Ds As New DataSet
            Dim Dt As DataTable
            Dim objCmd As SqlCommand
            Dim strSQL As String
            'Dim Dr5 As DataRow
            Dim Str As String = ""
            Dim office_pea As String = ""
            Dim where_str As String = "SELECT [event],[office_pea] FROM [SA_System].[scada].[View_line_ffg] order by time_commu asc"
            Dim Adpt As New SqlClient.SqlDataAdapter(where_str, StrCon)
            Adpt.Fill(Ds, "line_ffg1")
            Dt = Ds.Tables("line_ffg1")

            If Dt.Rows.Count > 0 Then
                StrCon.Open()
                For i As Integer = 0 To Dt.Rows.Count - 1

                    Str = Str & Dt.Rows(i).Item(0).ToString & vbCrLf
                    office_pea = Dt.Rows(i).Item(1).ToString
                    strSQL = "DELETE FROM [SA_System].[scada].[scada_line]  where event = '" & Dt.Rows(i).Item(0).ToString & "'"
                    objCmd = New SqlCommand(strSQL, StrCon)

                    objCmd.ExecuteNonQuery()

                Next
                StrCon.Close()
            Else

                Str = "0"
                office_pea = "0"
            End If


            If Str <> "0" Then

                Dim postData = String.Format("message=" & Str)
                Dim data = Encoding.UTF8.GetBytes(postData)

                request.Method = "POST"
                request.ContentType = "application/x-www-form-urlencoded"
                request.ContentLength = data.Length
                'request.Headers.Add("Authorization", "Bearer gdwshgUNuBBsry5lift6D6MXB35crHY24qt8q6YBBn5") 'ffg

                Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= staging")
                Dim Ds5 As New DataSet
                Dim Dt5 As DataTable
                Dim where_str5 As String = "SELECT token FROM line.line_token where group_name  = '" & office_pea & "'"
                Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str5, StrCon5)
                Adpt5.Fill(Ds5, "select_down_frtu")
                Dt5 = Ds5.Tables("select_down_frtu")


                If Dt5.Rows.Count > 0 Then
                    request.Headers.Add("Authorization", Dt5.Rows(0).Item(0).ToString) 'ffg

                    Using stream = request.GetRequestStream()
                        stream.Write(data, 0, data.Length)
                    End Using
                    Dim response = DirectCast(request.GetResponse(), HttpWebResponse)
                    Dim responseString = New StreamReader(response.GetResponseStream()).ReadToEnd()
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub
    Sub line_notify_ffg_s2()
        Try

            Dim request = DirectCast(WebRequest.Create("https://notify-api.line.me/api/notify"), HttpWebRequest)
            ' Dim Str As String = line_ffg1()

            Dim StrCon As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= SA_System")
            Dim Ds As New DataSet
            Dim Dt As DataTable
            Dim objCmd As SqlCommand
            Dim strSQL As String
            'Dim Dr5 As DataRow
            Dim Str As String = ""
            Dim office_pea As String = ""
            Dim where_str As String = "SELECT [event],[office_pea] FROM [SA_System].[scada].[scada_line] where id_num = '42' order by time_commu asc"
            Dim Adpt As New SqlClient.SqlDataAdapter(where_str, StrCon)
            Adpt.Fill(Ds, "line_ffg1")
            Dt = Ds.Tables("line_ffg1")

            If Dt.Rows.Count > 0 Then
                StrCon.Open()
                For i As Integer = 0 To Dt.Rows.Count - 1

                    Str = Str & Dt.Rows(i).Item(0).ToString & vbCrLf
                    office_pea = Dt.Rows(i).Item(1).ToString
                    strSQL = "DELETE FROM [SA_System].[scada].[scada_line]  where event = '" & Dt.Rows(i).Item(0).ToString & "'"
                    objCmd = New SqlCommand(strSQL, StrCon)

                    objCmd.ExecuteNonQuery()

                Next
                StrCon.Close()
            Else

                Str = "0"
                office_pea = "0"
            End If


            If Str <> "0" Then

                Dim postData = String.Format("message=" & Str)
                Dim data = Encoding.UTF8.GetBytes(postData)

                request.Method = "POST"
                request.ContentType = "application/x-www-form-urlencoded"
                request.ContentLength = data.Length
                'request.Headers.Add("Authorization", "Bearer gdwshgUNuBBsry5lift6D6MXB35crHY24qt8q6YBBn5") 'ffg

                Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= staging")
                Dim Ds5 As New DataSet
                Dim Dt5 As DataTable
                Dim where_str5 As String = "SELECT token FROM line.line_token where group_name  = 'ffg_s2'"
                Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str5, StrCon5)
                Adpt5.Fill(Ds5, "line_token")
                Dt5 = Ds5.Tables("line_token")


                If Dt5.Rows.Count > 0 Then
                    request.Headers.Add("Authorization", Dt5.Rows(0).Item(0).ToString) 'ffg

                    Using stream = request.GetRequestStream()
                        stream.Write(data, 0, data.Length)
                    End Using
                    Dim response = DirectCast(request.GetResponse(), HttpWebResponse)
                    Dim responseString = New StreamReader(response.GetResponseStream()).ReadToEnd()
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub
    Sub line_notify_sys_scada()
        Try

            Dim request = DirectCast(WebRequest.Create("https://notify-api.line.me/api/notify"), HttpWebRequest)
            ' Dim Str As String = line_ffg1()

            Dim StrCon As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= SA_System")
            Dim Ds As New DataSet
            Dim Dt As DataTable
            Dim objCmd As SqlCommand
            Dim strSQL As String
            'Dim Dr5 As DataRow
            Dim Str As String = ""
            Dim office_pea As String = ""
            Dim where_str As String = "SELECT [event],[office_pea] FROM [SA_System].[scada].[scada_line] where id_num = '99' order by time_commu asc"
            Dim Adpt As New SqlClient.SqlDataAdapter(where_str, StrCon)
            Adpt.Fill(Ds, "line_ffg1")
            Dt = Ds.Tables("line_ffg1")

            If Dt.Rows.Count > 0 Then
                StrCon.Open()
                For i As Integer = 0 To Dt.Rows.Count - 1

                    Str = Str & Dt.Rows(i).Item(0).ToString & vbCrLf
                    office_pea = Dt.Rows(i).Item(1).ToString
                    strSQL = "DELETE FROM [SA_System].[scada].[scada_line]  where event = '" & Dt.Rows(i).Item(0).ToString & "'"
                    objCmd = New SqlCommand(strSQL, StrCon)

                    objCmd.ExecuteNonQuery()

                Next
                StrCon.Close()
            Else

                Str = "0"
                office_pea = "0"
            End If


            If Str <> "0" Then

                Dim postData = String.Format("message=" & Str)
                Dim data = Encoding.UTF8.GetBytes(postData)

                request.Method = "POST"
                request.ContentType = "application/x-www-form-urlencoded"
                request.ContentLength = data.Length
                'request.Headers.Add("Authorization", "Bearer gdwshgUNuBBsry5lift6D6MXB35crHY24qt8q6YBBn5") 'ffg

                Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= staging")
                Dim Ds5 As New DataSet
                Dim Dt5 As DataTable
                Dim where_str5 As String = "SELECT token FROM line.line_token where group_name  = 'ผรศ'"
                Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str5, StrCon5)
                Adpt5.Fill(Ds5, "line_token")
                Dt5 = Ds5.Tables("line_token")


                If Dt5.Rows.Count > 0 Then
                    request.Headers.Add("Authorization", Dt5.Rows(0).Item(0).ToString) 'ผรศ

                    Using stream = request.GetRequestStream()
                        stream.Write(data, 0, data.Length)
                    End Using
                    Dim response = DirectCast(request.GetResponse(), HttpWebResponse)
                    Dim responseString = New StreamReader(response.GetResponseStream()).ReadToEnd()
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub
    Function select_down_srtu()
        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= SA_System")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        Dim objCmd As SqlCommand
        Dim strSQL As String
        'Dim Dr5 As DataRow
        Dim str As String = ""
        Dim where_str As String = "SELECT  [event] FROM [SA_System].[scada].[scada_line] where id_num = 1 order by time_commu asc"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "select_down_srtu")
        Dt5 = Ds5.Tables("select_down_srtu")

        If Dt5.Rows.Count > 0 Then
            For i As Integer = 0 To Dt5.Rows.Count - 1

                str = str & Dt5.Rows(i).Item(0).ToString & vbCrLf
                strSQL = "DELETE FROM [SA_System].[scada].[scada_line]  where event = '" & Dt5.Rows(i).Item(0).ToString & "'"
                objCmd = New SqlCommand(strSQL, StrCon5)
                StrCon5.Open()
                objCmd.ExecuteNonQuery()
                StrCon5.Close()
            Next
        Else

            str = "0"
        End If

        Return str

    End Function
    Function select_down_frtu1()
        Dim StrCon5 As New SqlClient.SqlConnection("Server=172.30.203.155; uid=sa;pwd=1234; database= SA_System")
        Dim Ds5 As New DataSet
        Dim Dt5 As DataTable
        'Dim Dr5 As DataRow
        Dim str As String = ""
        Dim where_str As String = "SELECT  [event] FROM [SA_System].[scada].[scada_line] where id_num = 2 order by time_commu asc"
        Dim Adpt5 As New SqlClient.SqlDataAdapter(where_str, StrCon5)
        Adpt5.Fill(Ds5, "select_down_frtu")
        Dt5 = Ds5.Tables("select_down_frtu")

        If Dt5.Rows.Count > 0 Then
            str = vbCrLf
            For i As Integer = 0 To Dt5.Rows.Count - 1

                str = str & Dt5.Rows(i).Item(0) & vbCrLf
            Next

            Dim objCmd As SqlCommand
            Dim strSQL As String
            ''strConnString = "Server=172.30.203.155;Uid=sa;PASSWORD=1234;database=SA_System;Max Pool Size=400;Connect Timeout=600;"
            'objConn.ConnectionString = "Server=172.30.203.155; uid=sa;pwd=1234; database= SA_System"
            'objConn.Open()
            'If SCADA_Status_program_check_update(program_name) Then
            strSQL = "DELETE FROM [SA_System].[scada].[scada_line]  where id_num = 2 "

            'Else
            'strSQL = "INSERT INTO scada.program_status (region,program_run,status,datetime_run )VALUES ('" & set_area.area & "','" & program_name & "','" & status_program & "','" & Date.Now & "')"


            objCmd = New SqlCommand(strSQL, StrCon5)
            StrCon5.Open()
            objCmd.ExecuteNonQuery()
            StrCon5.Close()


        Else

            str = "0"

        End If





        Return str

    End Function



End Class
