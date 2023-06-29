Public Class frmMain
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal DSN As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _strDSN = DSN
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnDone As System.Windows.Forms.Button
    Friend WithEvents txtSUSER_NAME As System.Windows.Forms.TextBox
    Friend WithEvents txtSUSER_SNAME As System.Windows.Forms.TextBox
    Friend WithEvents txtUSER_NAME As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtServername As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtDB_Name As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtHost_Name As System.Windows.Forms.TextBox
    Friend WithEvents txtUSER As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtSUSER_NAME = New System.Windows.Forms.TextBox
        Me.txtSUSER_SNAME = New System.Windows.Forms.TextBox
        Me.txtUSER_NAME = New System.Windows.Forms.TextBox
        Me.txtUSER = New System.Windows.Forms.TextBox
        Me.btnDone = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtServername = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtDB_Name = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtHost_Name = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(14, 97)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 23)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "SUSER_NAME"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(14, 121)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 23)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "SUSER_SNAME"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(14, 145)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 23)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "USER_NAME"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(14, 169)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 23)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "USER"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtSUSER_NAME
        '
        Me.txtSUSER_NAME.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSUSER_NAME.Location = New System.Drawing.Point(126, 97)
        Me.txtSUSER_NAME.Name = "txtSUSER_NAME"
        Me.txtSUSER_NAME.ReadOnly = True
        Me.txtSUSER_NAME.Size = New System.Drawing.Size(200, 20)
        Me.txtSUSER_NAME.TabIndex = 7
        '
        'txtSUSER_SNAME
        '
        Me.txtSUSER_SNAME.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSUSER_SNAME.Location = New System.Drawing.Point(126, 121)
        Me.txtSUSER_SNAME.Name = "txtSUSER_SNAME"
        Me.txtSUSER_SNAME.ReadOnly = True
        Me.txtSUSER_SNAME.Size = New System.Drawing.Size(200, 20)
        Me.txtSUSER_SNAME.TabIndex = 9
        '
        'txtUSER_NAME
        '
        Me.txtUSER_NAME.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUSER_NAME.Location = New System.Drawing.Point(126, 145)
        Me.txtUSER_NAME.Name = "txtUSER_NAME"
        Me.txtUSER_NAME.ReadOnly = True
        Me.txtUSER_NAME.Size = New System.Drawing.Size(200, 20)
        Me.txtUSER_NAME.TabIndex = 11
        '
        'txtUSER
        '
        Me.txtUSER.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUSER.Location = New System.Drawing.Point(126, 169)
        Me.txtUSER.Name = "txtUSER"
        Me.txtUSER.ReadOnly = True
        Me.txtUSER.Size = New System.Drawing.Size(200, 20)
        Me.txtUSER.TabIndex = 13
        '
        'btnDone
        '
        Me.btnDone.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDone.Location = New System.Drawing.Point(134, 209)
        Me.btnDone.Name = "btnDone"
        Me.btnDone.Size = New System.Drawing.Size(75, 23)
        Me.btnDone.TabIndex = 14
        Me.btnDone.Text = "Done"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(14, 12)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(100, 23)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "ServerName"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtServername
        '
        Me.txtServername.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtServername.Location = New System.Drawing.Point(126, 12)
        Me.txtServername.Name = "txtServername"
        Me.txtServername.ReadOnly = True
        Me.txtServername.Size = New System.Drawing.Size(200, 20)
        Me.txtServername.TabIndex = 1
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(14, 38)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(100, 23)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "Default DB"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtDB_Name
        '
        Me.txtDB_Name.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDB_Name.Location = New System.Drawing.Point(126, 38)
        Me.txtDB_Name.Name = "txtDB_Name"
        Me.txtDB_Name.ReadOnly = True
        Me.txtDB_Name.Size = New System.Drawing.Size(200, 20)
        Me.txtDB_Name.TabIndex = 3
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(14, 64)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(100, 23)
        Me.Label7.TabIndex = 4
        Me.Label7.Text = "HOST_NAME"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtHost_Name
        '
        Me.txtHost_Name.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtHost_Name.Location = New System.Drawing.Point(126, 64)
        Me.txtHost_Name.Name = "txtHost_Name"
        Me.txtHost_Name.ReadOnly = True
        Me.txtHost_Name.Size = New System.Drawing.Size(200, 20)
        Me.txtHost_Name.TabIndex = 5
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(360, 255)
        Me.Controls.Add(Me.btnDone)
        Me.Controls.Add(Me.txtUSER)
        Me.Controls.Add(Me.txtUSER_NAME)
        Me.Controls.Add(Me.txtSUSER_SNAME)
        Me.Controls.Add(Me.txtHost_Name)
        Me.Controls.Add(Me.txtDB_Name)
        Me.Controls.Add(Me.txtServername)
        Me.Controls.Add(Me.txtSUSER_NAME)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMain"
        Me.Text = "SQL Who Am I"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private _strDSN As String
#End Region

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim cn As SqlClient.SqlConnection
        Dim cmd As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim strSQL As String

        Try
            strSQL = "SELECT @@ServerName AS ServerName, DB_Name() AS DB_Name, HOST_NAME() AS Host_Name, SUSER_NAME() AS SUSERNAME, SUSER_SNAME() AS SUSERSNAME, USER_NAME() AS USERNAME, USER AS JustUser, SUSER_SID() AS SUSERSID"
            cn = New SqlClient.SqlConnection(_strDSN)
            cn.Open()

            cmd = cn.CreateCommand
            cmd.CommandText = strSQL

            dr = cmd.ExecuteReader
            If dr.Read Then
                txtServername.Text = DBNotNull.DBNotNullStr(dr.Item("ServerName"))
                txtDB_Name.Text = DBNotNull.DBNotNullStr(dr.Item("DB_Name"))
                txtHost_Name.Text = DBNotNull.DBNotNullStr(dr.Item("Host_Name"))
                txtSUSER_NAME.Text = DBNotNull.DBNotNullStr(dr.Item("SUSERNAME"))
                txtSUSER_SNAME.Text = DBNotNull.DBNotNullStr(dr.Item("SUSERSNAME"))
                txtUSER_NAME.Text = DBNotNull.DBNotNullStr(dr.Item("USERNAME"))
                txtUSER.Text = DBNotNull.DBNotNullStr(dr.Item("JustUser"))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, , "Error")
        End Try
    End Sub

    Private Sub btnDone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDone.Click
        End
    End Sub
End Class
