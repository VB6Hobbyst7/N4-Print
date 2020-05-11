Imports System.Drawing.Printing
Imports System.IO
Imports System.Net
Imports Newtonsoft.Json
'Imports ServiceStack.Redis
Imports StackExchange.Redis

'Imports System.Net
Imports System.Web.Script.Serialization
Imports System.Windows.Forms
Imports System.Threading.Channels
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Xml

Imports Program04.N4JsonDocument
Imports System.Collections.Generic
Imports Newtonsoft.Json.Linq
Imports System.Globalization

Public Class Form1
    Inherits System.Windows.Forms.Form

    ' Constant variable holding the Printer name.
    'BIXOLON SRP-330II
    'CutePDF Writer
    Private BARCODE_SERVER As String = "http://10.24.50.91:8010/"
    Private PRINTER_NAME As String = "EPSON TM-T82"
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents lcb1_logo As PictureBox
    Friend WithEvents iso_logo As PictureBox

    ' Variables/Objects.
    Private WithEvents pdPrint As PrintDocument
    Private WithEvents pdEirPrint As PrintDocument

    'Dim redisClient As RedisClient
    Dim redis As ConnectionMultiplexer
    Friend WithEvents lblMsg As Label
    Dim db As IDatabase

    Dim objCurrentPrintingJson As Object
    Friend WithEvents txtLicense As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents btnReprint As Button
    Friend WithEvents lblFile As Label
    Friend WithEvents btnFile As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents btnFilePrint As Button
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents txtDamageMessage As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents txtContainer As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents btnSaveDamage As Button
    Friend WithEvents lblStatus As Label
    Friend WithEvents txtLicensePlate As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents lcmt_logo As PictureBox
    Friend WithEvents lblPrinter As Label
    Friend WithEvents lblDatabase As Label
    Dim vDocumentType As String


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents pbImage As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lcmt_logo = New System.Windows.Forms.PictureBox()
        Me.btnFilePrint = New System.Windows.Forms.Button()
        Me.lblFile = New System.Windows.Forms.Label()
        Me.btnFile = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.iso_logo = New System.Windows.Forms.PictureBox()
        Me.btnReprint = New System.Windows.Forms.Button()
        Me.lcb1_logo = New System.Windows.Forms.PictureBox()
        Me.txtLicense = New System.Windows.Forms.TextBox()
        Me.pbImage = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblMsg = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.txtLicensePlate = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.btnSaveDamage = New System.Windows.Forms.Button()
        Me.txtDamageMessage = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtContainer = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblPrinter = New System.Windows.Forms.Label()
        Me.lblDatabase = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        CType(Me.lcmt_logo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.iso_logo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lcb1_logo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbImage, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(284, 250)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 28)
        Me.cmdClose.TabIndex = 3
        Me.cmdClose.Text = "Close"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lcmt_logo)
        Me.GroupBox1.Controls.Add(Me.btnFilePrint)
        Me.GroupBox1.Controls.Add(Me.lblFile)
        Me.GroupBox1.Controls.Add(Me.btnFile)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.iso_logo)
        Me.GroupBox1.Controls.Add(Me.btnReprint)
        Me.GroupBox1.Controls.Add(Me.lcb1_logo)
        Me.GroupBox1.Controls.Add(Me.txtLicense)
        Me.GroupBox1.Controls.Add(Me.pbImage)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.lblMsg)
        Me.GroupBox1.Controls.Add(Me.PictureBox1)
        Me.GroupBox1.Location = New System.Drawing.Point(9, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(354, 106)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Client : "
        '
        'lcmt_logo
        '
        Me.lcmt_logo.Image = CType(resources.GetObject("lcmt_logo.Image"), System.Drawing.Image)
        Me.lcmt_logo.Location = New System.Drawing.Point(311, 2)
        Me.lcmt_logo.Name = "lcmt_logo"
        Me.lcmt_logo.Size = New System.Drawing.Size(29, 33)
        Me.lcmt_logo.TabIndex = 15
        Me.lcmt_logo.TabStop = False
        Me.lcmt_logo.Visible = False
        '
        'btnFilePrint
        '
        Me.btnFilePrint.Location = New System.Drawing.Point(263, 72)
        Me.btnFilePrint.Name = "btnFilePrint"
        Me.btnFilePrint.Size = New System.Drawing.Size(87, 28)
        Me.btnFilePrint.TabIndex = 14
        Me.btnFilePrint.Text = "Print File"
        Me.btnFilePrint.UseVisualStyleBackColor = True
        '
        'lblFile
        '
        Me.lblFile.AutoSize = True
        Me.lblFile.Location = New System.Drawing.Point(95, 82)
        Me.lblFile.Name = "lblFile"
        Me.lblFile.Size = New System.Drawing.Size(16, 13)
        Me.lblFile.TabIndex = 13
        Me.lblFile.Text = "..."
        '
        'btnFile
        '
        Me.btnFile.Location = New System.Drawing.Point(62, 77)
        Me.btnFile.Name = "btnFile"
        Me.btnFile.Size = New System.Drawing.Size(27, 23)
        Me.btnFile.TabIndex = 12
        Me.btnFile.Text = "..."
        Me.btnFile.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 82)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(29, 13)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "File :"
        '
        'iso_logo
        '
        Me.iso_logo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.iso_logo.Image = CType(resources.GetObject("iso_logo.Image"), System.Drawing.Image)
        Me.iso_logo.Location = New System.Drawing.Point(279, -15)
        Me.iso_logo.Name = "iso_logo"
        Me.iso_logo.Size = New System.Drawing.Size(39, 50)
        Me.iso_logo.TabIndex = 7
        Me.iso_logo.TabStop = False
        Me.iso_logo.Visible = False
        '
        'btnReprint
        '
        Me.btnReprint.Location = New System.Drawing.Point(263, 41)
        Me.btnReprint.Name = "btnReprint"
        Me.btnReprint.Size = New System.Drawing.Size(87, 27)
        Me.btnReprint.TabIndex = 10
        Me.btnReprint.Text = "&Reprint"
        Me.btnReprint.UseVisualStyleBackColor = True
        '
        'lcb1_logo
        '
        Me.lcb1_logo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.lcb1_logo.Image = CType(resources.GetObject("lcb1_logo.Image"), System.Drawing.Image)
        Me.lcb1_logo.Location = New System.Drawing.Point(226, -15)
        Me.lcb1_logo.Name = "lcb1_logo"
        Me.lcb1_logo.Size = New System.Drawing.Size(79, 83)
        Me.lcb1_logo.TabIndex = 6
        Me.lcb1_logo.TabStop = False
        Me.lcb1_logo.Visible = False
        '
        'txtLicense
        '
        Me.txtLicense.Location = New System.Drawing.Point(62, 41)
        Me.txtLicense.Name = "txtLicense"
        Me.txtLicense.Size = New System.Drawing.Size(144, 20)
        Me.txtLicense.TabIndex = 9
        Me.txtLicense.TabStop = False
        '
        'pbImage
        '
        Me.pbImage.Image = CType(resources.GetObject("pbImage.Image"), System.Drawing.Image)
        Me.pbImage.Location = New System.Drawing.Point(103, -5)
        Me.pbImage.Name = "pbImage"
        Me.pbImage.Size = New System.Drawing.Size(72, 40)
        Me.pbImage.TabIndex = 4
        Me.pbImage.TabStop = False
        Me.pbImage.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 44)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(50, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "License :"
        '
        'lblMsg
        '
        Me.lblMsg.AutoSize = True
        Me.lblMsg.Location = New System.Drawing.Point(6, 17)
        Me.lblMsg.Name = "lblMsg"
        Me.lblMsg.Size = New System.Drawing.Size(101, 13)
        Me.lblMsg.TabIndex = 7
        Me.lblMsg.Text = "Response Massage"
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PictureBox1.Location = New System.Drawing.Point(181, -5)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(39, 50)
        Me.PictureBox1.TabIndex = 5
        Me.PictureBox1.TabStop = False
        Me.PictureBox1.Visible = False
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtLicensePlate)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.lblStatus)
        Me.GroupBox2.Controls.Add(Me.btnSaveDamage)
        Me.GroupBox2.Controls.Add(Me.txtDamageMessage)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.txtContainer)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Location = New System.Drawing.Point(9, 124)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(353, 121)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Damage Information"
        '
        'txtLicensePlate
        '
        Me.txtLicensePlate.Location = New System.Drawing.Point(263, 22)
        Me.txtLicensePlate.Name = "txtLicensePlate"
        Me.txtLicensePlate.Size = New System.Drawing.Size(55, 20)
        Me.txtLicensePlate.TabIndex = 11
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(188, 25)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(77, 13)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "License Plate :"
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(6, 95)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(16, 13)
        Me.lblStatus.TabIndex = 16
        Me.lblStatus.Text = "..."
        '
        'btnSaveDamage
        '
        Me.btnSaveDamage.Enabled = False
        Me.btnSaveDamage.Location = New System.Drawing.Point(263, 62)
        Me.btnSaveDamage.Name = "btnSaveDamage"
        Me.btnSaveDamage.Size = New System.Drawing.Size(87, 28)
        Me.btnSaveDamage.TabIndex = 13
        Me.btnSaveDamage.Text = "Save"
        Me.btnSaveDamage.UseVisualStyleBackColor = True
        '
        'txtDamageMessage
        '
        Me.txtDamageMessage.Location = New System.Drawing.Point(70, 48)
        Me.txtDamageMessage.Multiline = True
        Me.txtDamageMessage.Name = "txtDamageMessage"
        Me.txtDamageMessage.Size = New System.Drawing.Size(187, 42)
        Me.txtDamageMessage.TabIndex = 12
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 51)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Detail :"
        '
        'txtContainer
        '
        Me.txtContainer.Location = New System.Drawing.Point(70, 22)
        Me.txtContainer.Name = "txtContainer"
        Me.txtContainer.Size = New System.Drawing.Size(112, 20)
        Me.txtContainer.TabIndex = 10
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 25)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(58, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Container :"
        '
        'lblPrinter
        '
        Me.lblPrinter.AutoSize = True
        Me.lblPrinter.Location = New System.Drawing.Point(12, 250)
        Me.lblPrinter.Name = "lblPrinter"
        Me.lblPrinter.Size = New System.Drawing.Size(43, 13)
        Me.lblPrinter.TabIndex = 10
        Me.lblPrinter.Text = "Printer :"
        '
        'lblDatabase
        '
        Me.lblDatabase.AutoSize = True
        Me.lblDatabase.Location = New System.Drawing.Point(12, 265)
        Me.lblDatabase.Name = "lblDatabase"
        Me.lblDatabase.Size = New System.Drawing.Size(59, 13)
        Me.lblDatabase.TabIndex = 11
        Me.lblDatabase.Text = "Database :"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(367, 279)
        Me.Controls.Add(Me.lblDatabase)
        Me.Controls.Add(Me.lblPrinter)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "N4-Print"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.lcmt_logo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.iso_logo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lcb1_logo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbImage, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Public Function MMToHP(ByVal pMM As Integer) As Integer 'HP : hundredth of an inch

        Return Convert.ToInt32(pMM * 100 / 25.4)

    End Function



    ' The event handler function when pdPrint.Print is called.
    ' This is where the actual printing of sample data to the printer is made.
    Private Sub pdPrint_Print(ByVal sender As System.Object, ByVal e As PrintPageEventArgs) Handles pdPrint.PrintPage
        Dim x, y, lineOffset As Integer
        'Initial value from JsonObject
        Dim vPrintDate As String = objCurrentPrintingJson("start")
        Dim vLicence As String = objCurrentPrintingJson("license")
        Dim vNote As String = ""

        '--Added by Chutchai on Dec 10,2019 
        'To support Note key
        Dim key As String = ""
        For Each key In objCurrentPrintingJson.keys
            If key = "note" Then
                vNote = objCurrentPrintingJson("note")
                Exit For
            End If
        Next
        '-----------------------------

        If vDocumentType = "EIR" Then
            Dim objContainer As Object
            Dim vYStart As Integer = 10
            For Each objContainer In objCurrentPrintingJson("containers")
                EirPrint(objContainer, vLicence, objCurrentPrintingJson("company"), e)
                Exit For
            Next

            Exit Sub
        End If

        If vDocumentType = "TID" Then
            ' Instantiate font objects used in printing.
            Dim printFont As New Font("Arial", 8, FontStyle.Regular, GraphicsUnit.Point) 'Substituted to FontA Font
            Dim headFont As New Font("Arial", 16, FontStyle.Bold, GraphicsUnit.Point) 'Substituted to FontA Font
            Dim thaiFont As New Font("Microsoft San Serif", 8, FontStyle.Regular, GraphicsUnit.Point) 'Substituted to FontA Font



            e.Graphics.PageUnit = GraphicsUnit.Point
            x = 1
            y = 1
            e.Graphics.DrawString("TID", headFont, Brushes.Black, x, y)
            lineOffset = printFont.GetHeight(e.Graphics)

            y += lineOffset + 10
            x = 1
            e.Graphics.DrawString(vPrintDate, printFont, Brushes.Black, x, y)
            y += 10 : x = 1
            e.Graphics.DrawString(My.Resources.thai.licence + ": " & vLicence, thaiFont, Brushes.Black, x, y)
            y += 10 : x = 1
            '---Added by Chutchai on Dec 10,2019 -- to add Note message.
            e.Graphics.DrawString(My.Resources.thai.note + ": " & vNote, thaiFont, Brushes.Black, x, y)
            lineOffset = thaiFont.GetHeight(e.Graphics)
            y += lineOffset : x = 1
            e.Graphics.DrawString("___________________________________________", printFont, Brushes.Black, x, y)

            Dim positionFont As New Font("Arial", 32, FontStyle.Bold, GraphicsUnit.Point) 'Substituted to FontA Font
            Dim containerFont As New Font("Arial", 12, FontStyle.Regular, GraphicsUnit.Point) 'Substituted to FontA Font
            Dim sealFont As New Font("Arial", 8, FontStyle.Regular, GraphicsUnit.Point) 'Substituted to FontA Font
            Dim xSeal As Integer = 120



            '--Print  COntainer
            'y += 10 : x = 10
            lineOffset = printFont.GetHeight(e.Graphics)
            Dim objContainer As Object
            Dim vTransType As String = ""
            Dim Exportwords As Regex = New Regex("RE|RM|Dray in")
            Dim Importwords As Regex = New Regex("DI|DM|Dray out")
            For Each objContainer In objCurrentPrintingJson("containers")
                'lineOffset = positionFont.GetHeight(e.Graphics)
                y += lineOffset : x = 1
                If Exportwords.IsMatch(objContainer("trans_type").ToString) Then
                    vTransType = My.Resources.thai.liftoff 'ส่งตู้
                Else
                    vTransType = My.Resources.thai.lifton 'รับตู้
                End If
                e.Graphics.DrawString(vTransType + ": " & objContainer("number"), containerFont, Brushes.Black, x, y)
                lineOffset = containerFont.GetHeight(e.Graphics)
                y += lineOffset : x = 1
                e.Graphics.DrawString(objContainer("position"), positionFont, Brushes.Black, x, y)
                e.Graphics.DrawString("Seal  :" & objContainer("seal1"), sealFont, Brushes.Black, xSeal, y + 5)
                e.Graphics.DrawString("Damage: " & objContainer("damage"), sealFont, Brushes.Black, xSeal, y + sealFont.GetHeight(e.Graphics) + 10)
                lineOffset = positionFont.GetHeight(e.Graphics)
            Next

            Dim url As String = BARCODE_SERVER & vLicence
            'Dim tClient As WebClient = New WebClient
            'Dim tImage As Bitmap = Bitmap.FromStream(New MemoryStream(tClient.DownloadData(url)))
            'PictureBox1.Image = tImage

            'Edit by Chutchai on May 11,2020
            'Change method to get QR code
            PictureBox1.Image = getQRCode(url)
            '---------------------------------
            lineOffset = positionFont.GetHeight(e.Graphics)
            y += lineOffset
            x = 1
            'Added by Chutchai on May 11,2020
            'if No QR image returned , not print.
            If Not PictureBox1.Image Is Nothing Then
                e.Graphics.DrawImage(PictureBox1.Image, x, y, 160, 160)
            End If
            '-------------------------------------
            y += 161
                e.Graphics.DrawString("___________________________________________", printFont, Brushes.Black, x, y)
                'Print Notice

                lineOffset = printFont.GetHeight(e.Graphics)
                y = y + lineOffset
                Dim noticeFont As New Font("Microsoft San Serif", 12, FontStyle.Bold, GraphicsUnit.Point)
                e.Graphics.DrawString(My.Resources.thai.special_text, noticeFont, Brushes.Black, x, y)

                lineOffset = noticeFont.GetHeight(e.Graphics)
                y = y + lineOffset
                'to wordwrap text, you need to specify a layout rectangle
                noticeFont = New Font("Microsoft San Serif", 10, FontStyle.Regular, GraphicsUnit.Point)
                lineOffset = printFont.GetHeight(e.Graphics)
                Dim strPolicy As String = My.Resources.thai.notice
                e.Graphics.DrawString(strPolicy,
                                  noticeFont, Brushes.Black, New Rectangle(x, y + lineOffset, 200, 200),
                                  StringFormat.GenericTypographic)

            End If


            ' Indicate that no more data to print, and the Print Document can now send the print data to the spooler.
            e.HasMorePages = False
    End Sub

    Function getQRCode(vUrl As String) As Bitmap
        Dim vImage As Bitmap
        Try
            Dim tClient As New System.Net.WebClient
            'tClient.Credentials = New System.Net.NetworkCredential()
            Dim ImageInBytes() As Byte = tClient.DownloadData(vUrl)
            Dim ImageStream As New IO.MemoryStream(ImageInBytes)
            vImage = Bitmap.FromStream(ImageStream)
            tClient.Dispose()
        Catch ex As Exception
            vImage = Nothing
        End Try
        Return vImage
    End Function


    Private Sub EirPrint(container As Object, license As String, company As String,
                         ByVal e As PrintPageEventArgs)
        Dim x, y, lineOffset As Integer
        'Dim cultureInfo As New System.Globalization.CultureInfo("en-US")
        'Dim info As TextInfo = CultureInfo.InvariantCulture.TextInfo

        Dim info As TextInfo = New CultureInfo("en-US", False).TextInfo


        'pdPrint.DefaultPageSettings.PaperSize.Height = MMToHP(250)
        ' Instantiate font objects used in printing.
        Dim printFont As New Font("Arial", 8, FontStyle.Regular, GraphicsUnit.Point) 'Substituted to FontA Font
        Dim headFont As New Font("Arial", 10, FontStyle.Bold, GraphicsUnit.Point) 'Substituted to FontA Font
        'Dim thaiFont As New Font("Microsoft San Serif", 8, FontStyle.Regular, GraphicsUnit.Point) 'Substituted to FontA Font

        e.Graphics.PageUnit = GraphicsUnit.Point

        Dim x_end As Integer
        Dim x_half As Integer
        Dim y_start As Integer
        Dim y_end As Integer
        x = 5
        x_end = 220
        x_half = 110
        y = 10
        'Company logo
        If container("terminal") = "B1" Then
            e.Graphics.DrawImage(lcb1_logo.Image, 10, y, 70, 50)
        Else
            e.Graphics.DrawImage(lcmt_logo.Image, 10, y, 70, 50)
        End If

        x = 150
        e.Graphics.DrawImage(iso_logo.Image, x_half, y, 100, 60)
        'EIR Topic
        Dim thaiFont As New Font("Microsoft San Serif", 8, FontStyle.Regular, GraphicsUnit.Point) 'Substituted to FontA Font
        Dim thaiEIRFont As New Font("Microsoft San Serif", 14, FontStyle.Bold, GraphicsUnit.Point) 'Substituted to FontA Font

        y += 60 : x = 10

        e.Graphics.DrawString(My.Resources.thai.eir, thaiEIRFont, Brushes.Black, x, y)
        lineOffset = thaiEIRFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawString("EQUIPMENT INTERCHANGE RECEIPT", headFont, Brushes.Black, x, y)
        lineOffset = headFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawLine(New Pen(Color.Black, 1), New Point(x, y), New Point(x_end, y))
        lineOffset = 5

        '--First Row
        y += lineOffset : x = 10
        y_start = y
        e.Graphics.DrawString(My.Resources.thai.order_date + " Order date", printFont, Brushes.Black, x, y)
        e.Graphics.DrawString(My.Resources.thai.shpping_line + " Shipping line", printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawString(container("created"), printFont, Brushes.Black, x, y)
        e.Graphics.DrawString(container("line"), printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawLine(New Pen(Color.Black, 0), New Point(x, y), New Point(x_end, y))
        '----------------------------
        '--Second Row
        y += 1 : x = 10
        e.Graphics.DrawString(My.Resources.thai.container_no + " Container", printFont, Brushes.Black, x, y)
        e.Graphics.DrawString(My.Resources.thai.imo + " IMO / UN NO.", printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawString(container("number"), printFont, Brushes.Black, x, y)
        e.Graphics.DrawString(container("dg"), printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawLine(New Pen(Color.Black, 0), New Point(x, y), New Point(x_end, y))
        '----------------------------

        '--Third Row
        Dim voy_text As String
        voy_text = My.Resources.thai.vessel & "/" & My.Resources.thai.voy &
                    "Vessel/Voy"
        y += 1 : x = 10
        e.Graphics.DrawString(voy_text, printFont, Brushes.Black, x, y)
        e.Graphics.DrawString(My.Resources.thai.move + " Move", printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        'e.Graphics.DrawString(container("vessel_name"), printFont, Brushes.Black, x, y)
        'Modify on Feb 24,2020 - To add Voy-In

        If container("trans_type") = "DI" Or container("trans_type") = "DM" Then
            'Import
            e.Graphics.DrawString(LCase(container("vessel_name")) & "/" & container("voy_in"), printFont, Brushes.Black, x, y)
        Else
            'Export
            e.Graphics.DrawString(LCase(container("vessel_name")) & "/" & container("voy_out"), printFont, Brushes.Black, x, y)
        End If

        'x_half
        e.Graphics.DrawString(container("freightkind"), printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawLine(New Pen(Color.Black, 0), New Point(x, y), New Point(x_end, y))
        '----------------------------

        '--Fourt Row
        Dim temp_text As String
        temp_text = My.Resources.thai.temperature & " Temperature"
        y += 1 : x = 10
        e.Graphics.DrawString(temp_text, printFont, Brushes.Black, x, y)
        e.Graphics.DrawString(My.Resources.thai.pod + " POD", printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawString(container("temperature"), printFont, Brushes.Black, x, y)
        e.Graphics.DrawString(container("pod"), printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawLine(New Pen(Color.Black, 0), New Point(x, y), New Point(x_end, y))
        '----------------------------

        '--Fift Row
        y += 1 : x = 10
        e.Graphics.DrawString(My.Resources.thai.container_type & " Size /Type", printFont, Brushes.Black, x, y)
        e.Graphics.DrawString(My.Resources.thai.date_ + " Date", printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawString(container("iso_text"), printFont, Brushes.Black, x, y)
        e.Graphics.DrawString(container("created"), printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawLine(New Pen(Color.Black, 0), New Point(x, y), New Point(x_end, y))
        '----------------------------

        '--Six Row
        y += 1 : x = 10
        e.Graphics.DrawString(My.Resources.thai.iso_code & " ISO code", printFont, Brushes.Black, x, y)
        e.Graphics.DrawString(My.Resources.thai.license + " Plate No", printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawString(container("iso_code"), printFont, Brushes.Black, x, y)
        e.Graphics.DrawString(license, printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawLine(New Pen(Color.Black, 0), New Point(x, y), New Point(x_end, y))
        '----------------------------

        '--Seven Row
        y += 1 : x = 10
        e.Graphics.DrawString(My.Resources.thai.truck_company & " Truck Company", printFont, Brushes.Black, x, y)
        e.Graphics.DrawString(My.Resources.thai.consignee + " Consignee", printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawString(LCase(company), printFont, Brushes.Black, x, y)
        e.Graphics.DrawString("", printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawLine(New Pen(Color.Black, 0), New Point(x, y), New Point(x_end, y))
        '----------------------------

        '--Eight Row
        y += 1 : x = 10
        e.Graphics.DrawString(My.Resources.thai.booking_no & " Booking No", printFont, Brushes.Black, x, y)
        e.Graphics.DrawString(My.Resources.thai.seal_no + " Seal No", printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawString("", printFont, Brushes.Black, x, y)
        e.Graphics.DrawString(container("seal1"), printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawLine(New Pen(Color.Black, 0), New Point(x, y), New Point(x_end, y))
        '----------------------------

        '--Nine Row
        y += 1 : x = 10
        e.Graphics.DrawString(My.Resources.thai.weight & " Gross weight", printFont, Brushes.Black, x, y)
        e.Graphics.DrawString(My.Resources.thai.exception + " Exception", printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawString(container("gross_weight"), printFont, Brushes.Black, x, y)
        e.Graphics.DrawString("", printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawLine(New Pen(Color.Black, 0), New Point(x, y), New Point(x_end, y))
        '----------------------------

        '--Ten Row
        y += 1 : x = 10
        e.Graphics.DrawString(My.Resources.thai.genset_no & " Genset No", printFont, Brushes.Black, x, y)
        e.Graphics.DrawString(My.Resources.thai.damage + " Damage", printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawString("", printFont, Brushes.Black, x, y)
        e.Graphics.DrawString(container("damage"), printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawLine(New Pen(Color.Black, 0), New Point(x, y), New Point(x_end, y))
        '----------------------------

        '--Twelth Row
        y += 1 : x = 10
        e.Graphics.DrawString(My.Resources.thai.checker & " Checker", printFont, Brushes.Black, x, y)
        e.Graphics.DrawString(My.Resources.thai.date_ + " Date", printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawString(container("creator"), printFont, Brushes.Black, x, y)
        e.Graphics.DrawString(container("created"), printFont, Brushes.Black, x_half, y)
        lineOffset = printFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawLine(New Pen(Color.Black, 0), New Point(x, y), New Point(x_end, y))
        '----------------------------
        y_end = y

        '--Eleven Row
        y += 1 : x = 10
        e.Graphics.DrawString(My.Resources.thai.remark & " Remark", printFont, Brushes.Black, x, y)
        lineOffset = printFont.GetHeight(e.Graphics)

        'Get damage from Database
        Dim strKey As String = container("number") & ":" & license & ":damage"
        Dim strDamage As String = db.StringGet(strKey)
        Dim strDamageLine1 As String = ""
        Dim strDamageLine2 As String = ""
        If strDamage <> "" Then
            Dim str As String() = strDamage.Split(vbCrLf)
            strDamageLine1 = str(0).Replace(vbCrLf, "")
            If str.Length > 1 Then
                strDamageLine2 = str(1).Replace(vbLf, "")
            End If
        End If
        '-------------------------------------

        y += lineOffset : x = 10
        e.Graphics.DrawString(strDamageLine1, thaiFont, Brushes.Black, x, y)
        lineOffset = thaiFont.GetHeight(e.Graphics)

        y += lineOffset : x = 10
        e.Graphics.DrawString(strDamageLine2, thaiFont, Brushes.Black, x, y)
        lineOffset = thaiFont.GetHeight(e.Graphics)
        y += lineOffset : x = 10
        e.Graphics.DrawLine(New Pen(Color.Black, 0), New Point(x, y), New Point(x_end, y))
        '----------------------------


        'Hale vertical Line
        e.Graphics.DrawLine(New Pen(Color.Black, 0), New Point(x_half, y_start), New Point(x_half, y_end))
        'Notice Message
        y += 1 : x = 10

        'to wordwrap text, you need to specify a layout rectangle
        Dim strPolicy As String = "ข้าพเจ้าผู้ลงนามในเอกสารฉบับนี้ คือผู้ยอมรับว่าตู้สินค้าได้รับการตรวจสภาพ " &
                                "และสินค้าอยู่ในสภาพเรียบร้อย ยกเว้นเฉพาะรายการหมายเหตุความเสียหายที่ระบุเท่านั้น " &
                                "(THE UNDERSIGNED DECLARES THAT THE CONTAINER HAS BEEN INSPECTED AND ACCEPTED IN GOOD ORDER AND CONDITION " &
                                "EXCEPT NOTED DAMAGE CODES.)"
        e.Graphics.DrawString(strPolicy,
                              Me.Font, Brushes.Black, New Rectangle(x, y, 200, 200),
                              StringFormat.GenericTypographic)
        y = y + 70
        e.Graphics.DrawString("OPS 10 Rev. 04", thaiFont, Brushes.Black, x_end - 60, y)
        'Dim format As StringFormat = New StringFormat(StringFormatFlags.DirectionRightToLeft)
        'e.Graphics.DrawString("OPS 10 Rev. 04   ",
        '                      Me.Font, Brushes.Black, New Rectangle(x, y, 200, 12),
        '                      format)


        e.HasMorePages = False

    End Sub

    ' The executed function when the Close button is clicked.
    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Close()
    End Sub



    Sub print_document(license As String, Optional jsonStr As String = "")
        On Error GoTo HasError

        If jsonStr = "" Then
            If license = "" Then
                Exit Sub
            End If
            'Check License plate is Exist.
            jsonStr = db.StringGet(license)
            If jsonStr Is Nothing Then
                setText("Not found license : " & license)
                Exit Sub
            End If
        End If

        objCurrentPrintingJson = getJsonObject(jsonStr)
        'Dim test As String = objCurrentPrintingJson("license")

        vDocumentType = objCurrentPrintingJson("document")

        pdPrint = New PrintDocument
        ' Change the printer to the indicated printer
        pdPrint.PrinterSettings.PrinterName = PRINTER_NAME

        With pdPrint
            .PrinterSettings.PrinterName = PRINTER_NAME
            .DefaultPageSettings.PaperSize = New Printing.PaperSize("TM82", MMToHP(80),
                                                                                MMToHP(170)) '150
            .DefaultPageSettings.Margins = New System.Drawing.Printing.Margins(0, 0, 0, 0)
            .DefaultPageSettings.PrinterResolution.X = 204
            .DefaultPageSettings.PrinterResolution.Y = 204S
        End With





        If pdPrint.PrinterSettings.IsValid Then
            pdPrint.DocumentName = "TID priting"
            ' Start printing
            If vDocumentType = "TID" Then
                pdPrint.Print()
            End If
            If vDocumentType = "EIR" Then
                Dim objAllSingleContainer As Object
                objAllSingleContainer = getJsonObject(jsonStr)
                objCurrentPrintingJson.remove("containers")
                Dim c As Object
                'Dim strContainer As String
                For Each c In objAllSingleContainer("containers")
                    'filter only containern
                    Dim ContainerList As New List(Of Object)()
                    ContainerList.Add(c)
                    objCurrentPrintingJson("containers") = ContainerList
                    pdPrint.Print()
                    objCurrentPrintingJson.remove("containers")
                Next
            End If


        Else
            MessageBox.Show("Printer is not available.", "Program04", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
        Exit Sub

HasError:
        'MsgBox("Error on print_document : " & Err.Description)
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        'get Current IP address.
        Dim strHostName As String
        Dim GetIPv4Address As String = ""
        strHostName = System.Net.Dns.GetHostName()
        Dim iphe As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(strHostName)

        For Each ipheal As System.Net.IPAddress In iphe.AddressList
            If ipheal.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                GetIPv4Address = ipheal.ToString()
            End If
        Next

        GroupBox1.Text = GroupBox1.Text & " - " & strHostName & "(" & GetIPv4Address & ")"


        Dim versionNumber As Version
        versionNumber = Assembly.GetExecutingAssembly().GetName().Version
        Me.Text = Me.Text & " Version :" & versionNumber.Major & "." & versionNumber.Minor & "." & versionNumber.Build


        'Reading system.ini file
        Dim fName As String = ("system.ini")              'path to text file
        Dim My_Ini As New Ini(fName)

        Dim vPrinterSetting As String = ""
        Dim vServer As String = ""
        vPrinterSetting = My_Ini.GetValue("Setting", "printer")
        vServer = My_Ini.GetValue("Setting", "server")

        'Added by Chutchai on Apr 27,2020
        'BARCODE_SERVER
        Dim vBarcodeServer As String = ""
        vBarcodeServer = My_Ini.GetValue("Barcode", "server")
        If vServer = "" Then
            BARCODE_SERVER = vBarcodeServer
        End If
        '-End Barcode

        Dim SERVER_NAME As String = ""
        If vServer = "" Then
            SERVER_NAME = "127.0.0.1:6379"
        Else
            SERVER_NAME = vServer
        End If
        lblDatabase.Text = lblDatabase.Text & SERVER_NAME

        If vPrinterSetting = "" Then
            Dim objprinterName As PrinterSettings = New PrinterSettings()
            PRINTER_NAME = objprinterName.PrinterName
        Else
            PRINTER_NAME = vPrinterSetting
        End If
        lblPrinter.Text = lblPrinter.Text & PRINTER_NAME


        redis = ConnectionMultiplexer.Connect(SERVER_NAME)
        db = redis.GetDatabase()

        Dim subs As ISubscriber
        subs = redis.GetSubscriber()
        subs.Subscribe(strHostName.ToLower, New Action(Of RedisChannel, RedisValue)(AddressOf OnMessage))

    End Sub

    Private Sub OnMessage(ByVal channel As String, ByVal value As String)
        setText(value)
        print_document(value)
        setLicenseText(value)
    End Sub


    Private Sub setText(ByVal vStatus As String)
        'lblTested.Text = DataGridView2.RowCount.ToString
        If lblMsg.InvokeRequired Then
            lblMsg.Invoke(New Action(Of String)(AddressOf setText), vStatus)
            'lblTested.BackColor = Color.LightGreen
        Else
            lblMsg.Text = vStatus & " - " & Now
            lblMsg.ForeColor = Color.Blue
            'lblPassed.BackColor = Color.LightGreen
        End If
    End Sub

    Private Sub setLicenseText(ByVal vLicense As String)
        'lblTested.Text = DataGridView2.RowCount.ToString
        If txtLicense.InvokeRequired Then
            txtLicense.Invoke(New Action(Of String)(AddressOf setLicenseText), vLicense)
            'lblTested.BackColor = Color.LightGreen
        Else
            txtLicense.Text = vLicense
            'lblPassed.BackColor = Color.LightGreen
        End If
    End Sub

    'Public Sub Publish(ByVal channel As String, ByVal value As String)
    '    SourceClient.Publish(channel, Helper.GetBytes(value))
    'End Sub

    'Sub receiveMsg(channe As RedisChannel, msg As RedisValue)
    '    Console.Write(msg)
    'End Sub

    Public Shared Function getJsonObject(ByVal json As String) As Object
        Dim jss As New JavaScriptSerializer()
        'Dim client As WebClient = New WebClient()
        'Dim json As String = client.DownloadString(address)
        'jss = JavaScriptSerializer
        Dim data As Object = jss.Deserialize(Of Object)(json)
        Return data
    End Function

    Private Sub btnReprint_Click(sender As Object, e As EventArgs) Handles btnReprint.Click
        print_document(txtLicense.Text.Trim.ToUpper)
    End Sub



    Private Sub txtLicense_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtLicense.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            btnReprint_Click(sender, e)
        End If
    End Sub

    Private Sub txtLicense_TextChanged(sender As Object, e As EventArgs) Handles txtLicense.TextChanged

    End Sub

    Private Sub btnFile_Click(sender As Object, e As EventArgs) Handles btnFile.Click
        ' Call ShowDialog.
        Dim result As DialogResult = OpenFileDialog1.ShowDialog()

        ' Test result.
        If result = DialogResult.OK Then
            ' Do something.
            lblFile.Text = OpenFileDialog1.SafeFileName
        End If
    End Sub

    Private Sub btnFilePrint_Click(sender As Object, e As EventArgs) Handles btnFilePrint.Click
        Dim doc As XmlDocument
        doc = New XmlDocument
        doc.Load(OpenFileDialog1.FileName)
        Dim strJson As String
        strJson = makeJson(doc)


        'Dim jo As JObject = JObject.Parse(strJson)
        ''jo("containers").Last.Remove()
        'vDocumentType = jo("document")
        'Dim j As JObject
        'If vDocumentType = "EIR" Then
        '    For Each j In jo("containers")
        '        print_document("", j.ToString)
        '    Next

        'End If


        print_document("", strJson)

    End Sub
    Function makeJson(doc As XmlDocument) As String
        'Dim nodeList As XmlNodeList
        Dim root As XmlNode
        Dim body As XmlNode
        root = doc.DocumentElement
        Dim nsm As Xml.XmlNamespaceManager = New Xml.XmlNamespaceManager(doc.NameTable)
        nsm.AddNamespace("argo", "http://www.navis.com/argo")


        body = root.SelectSingleNode("argo:docBody/argo:truckVisit", nsm)
        Dim document_type, terminal As String
        Dim license_number, truck_company_code, truck_company_name, truck_start As String

        document_type = root.SelectSingleNode("argo:docDescription/docName", nsm).InnerText
        'printer = root.SelectSingleNode("argo:docDescription/ipAddress", nsm).InnerText


        license_number = root.SelectSingleNode("argo:docBody/argo:truckVisit/tvdtlsLicNbr", nsm).InnerText
        truck_company_code = root.SelectSingleNode("argo:docBody/argo:truckVisit/tvdtlsTrkCompany", nsm).InnerText
        truck_company_name = root.SelectSingleNode("argo:docBody/argo:truckVisit/tvdtlsTrkCompanyName", nsm).InnerText
        truck_start = root.SelectSingleNode("argo:docBody/argo:truckVisit/tvdtlsTrkStartTime", nsm).InnerText

        Dim nodeContainers As XmlNodeList
        nodeContainers = root.SelectNodes("argo:docBody/argo:trkTransaction", nsm)
        Dim node As XmlNode
        Dim container, transtype, line, iso, container_type, container_length, container_height,
            iso_text, vessel_code, vessel_name, voy_in, voy_out, freightkind, pod, gross_weight,
            created_date, created_user, changed_date, changed_user, is_damage, position, damage,
            category, seal1, seal2, new_position, temperature, dg As String

        'Create Json document
        Dim json As New N4JsonDocument
        With json
            .company = truck_company_name
            .company_code = truck_company_code
            .document = document_type
            .license = license_number
            .printer = "" 'printer
            .terminal = ""
            .start = truck_start
        End With

        For Each node In nodeContainers

            container = node.SelectSingleNode("tranCtrNbr", nsm).InnerText
            'In case Swap or Replace new container
            If Not node.SelectSingleNode("tranCtrNbrAssigned", nsm) Is Nothing Then
                If node.SelectSingleNode("tranCtrNbrAssigned", nsm).InnerText <> "" Then
                    container = node.SelectSingleNode("tranCtrNbrAssigned", nsm).InnerText
                End If
            End If


            transtype = node.SelectSingleNode("tranSubType", nsm).InnerText
            line = node.SelectSingleNode("tranLineId", nsm).InnerText
            iso = node.SelectSingleNode("tranCtrTypeId", nsm).InnerText

            If Not node.SelectSingleNode("tranEqoEqIsoGroup", nsm) Is Nothing Then
                container_type = node.SelectSingleNode("tranEqoEqIsoGroup", nsm).InnerText
            Else
                container_type = "DV"
            End If
            If Not node.SelectSingleNode("tranEqoEqLength", nsm) Is Nothing Then
                container_length = node.SelectSingleNode("tranEqoEqLength", nsm).InnerText.Replace("NOM", "")
            Else
                container_length = "40"
            End If
            If Not node.SelectSingleNode("tranEqoEqHeight", nsm) Is Nothing Then
                container_height = node.SelectSingleNode("tranEqoEqHeight", nsm).InnerText.Replace("NOM", "")
            Else
                container_height = "86"
            End If
            ''%s\' %s %s' % (container_length,int(container_height)/10,container_type)
            iso_text = container_length & "' " & Str(Val(container_height) / 10) & " " & container_type


            If Not node.SelectSingleNode("argo:tranCarrierVisit", nsm) Is Nothing Then
                vessel_code = node.SelectSingleNode("argo:tranCarrierVisit/cvId", nsm).InnerText
                vessel_name = node.SelectSingleNode("argo:tranCarrierVisit/cvCvdCarrierVehicleName", nsm).InnerText
                voy_in = node.SelectSingleNode("argo:tranCarrierVisit/cvCvdCarrierIbVygNbr", nsm).InnerText
                voy_out = node.SelectSingleNode("argo:tranCarrierVisit/cvCvdCarrierObVygNbr", nsm).InnerText
            Else
                vessel_code = ""
                vessel_name = ""
                voy_in = ""
                voy_out = ""
            End If

            freightkind = node.SelectSingleNode("tranCtrFreightKind", nsm).InnerText

            If Not node.SelectSingleNode("argo:tranDischargePoint1/pointId", nsm) Is Nothing Then
                pod = node.SelectSingleNode("argo:tranDischargePoint1/pointId", nsm).InnerText
            Else
                pod = ""
            End If

            gross_weight = node.SelectSingleNode("tranCtrGrossWeight", nsm).InnerText

            created_date = node.SelectSingleNode("tranCreated", nsm).InnerText
            created_user = node.SelectSingleNode("tranCreator", nsm).InnerText
            changed_date = node.SelectSingleNode("tranChanged", nsm).InnerText
            changed_user = node.SelectSingleNode("tranChanger", nsm).InnerText
            If Not node.SelectSingleNode("tranCtrIsDamaged", nsm) Is Nothing Then
                is_damage = node.SelectSingleNode("tranCtrIsDamaged", nsm).InnerText
            Else
                is_damage = "false"
            End If

            position = node.SelectSingleNode("tranFlexString02", nsm).InnerText

            new_position = Mid(position, 1, 3) & "-" & Mid(position, 4, 2)   '%s-%s' % (position[:3],position[3:5])

            If Not node.SelectSingleNode("argo:tranCtrDmg/dmgitemTypeDescription", nsm) Is Nothing Then
                damage = node.SelectSingleNode("argo:tranCtrDmg/dmgitemTypeDescription", nsm).InnerText
            Else
                damage = ""
            End If
            category = node.SelectSingleNode("tranUnitCategory", nsm).InnerText.Replace("UnitCategoryEnum", "")
            If Not node.SelectSingleNode("tranSealNbr1", nsm) Is Nothing Then
                seal1 = node.SelectSingleNode("tranSealNbr1", nsm).InnerText
            Else
                seal1 = ""
            End If

            If Not node.SelectSingleNode("tranSealNbr2", nsm) Is Nothing Then
                seal2 = node.SelectSingleNode("tranSealNbr2", nsm).InnerText
            Else
                seal2 = ""
            End If

            If Not node.SelectSingleNode("tranUnitFlexString01", nsm) Is Nothing Then
                terminal = node.SelectSingleNode("tranUnitFlexString01", nsm).InnerText
            Else
                terminal = ""
            End If

            temperature = ""
            If Not node.SelectSingleNode("tranTempRequired", nsm) Is Nothing Then
                temperature = node.SelectSingleNode("tranTempRequired", nsm).InnerText
            End If

            dg = ""
            'argo:tranHazard
            If Not node.SelectSingleNode("argo:tranHazard/hzrdiImdgCode", nsm) Is Nothing Then
                dg = node.SelectSingleNode("argo:tranHazard/hzrdiImdgCode", nsm).InnerText
            Else
                dg = ""
            End If

            If Not node.SelectSingleNode("argo:tranHazard/hzrdiUNnum", nsm) Is Nothing Then
                dg = dg & "/" & node.SelectSingleNode("argo:tranHazard/hzrdiUNnum", nsm).InnerText
            End If


            Dim con As New Container
            With con
                .number = container
                .category = category
                .changer = changed_user
                .changed = changed_date
                .creator = created_user
                .created = created_date
                .damage = damage
                .freightkind = freightkind
                .gross_weight = gross_weight
                .iso_code = iso
                .iso_text = iso_text
                .line = line
                .pod = pod
                .position = new_position
                .seal1 = seal1
                .seal2 = seal2
                .trans_type = transtype
                .vessel_code = vessel_code
                .vessel_name = vessel_name
                .voy_in = voy_in
                .voy_out = voy_out
                .temperature = temperature
                .terminal = terminal
                .dg = dg
            End With
            json.terminal = terminal
            json.containers.Add(con)
        Next
        Dim output As String
        output = JsonConvert.SerializeObject(json)
        Return output

    End Function

    Private Sub btnSaveDamage_Click(sender As Object, e As EventArgs) Handles btnSaveDamage.Click

        Try


            Dim ttl As TimeSpan = New TimeSpan(365 * 2, 0, 0, 0, 0) ' 60 * 60 * 24 * 365 * 2 '2 years
            Dim key As String = txtContainer.Text.Trim.ToUpper & ":" & txtLicensePlate.Text.Trim.ToUpper &
                                ":damage"
            Dim value As String = txtDamageMessage.Text.Trim
            db.StringSet(key, value, ttl)
            lblStatus.Text = key & "  Save to database successful."
            lblStatus.ForeColor = Color.Blue
            txtContainer.Text = ""
            txtDamageMessage.Text = ""
            txtLicensePlate.Text = ""
            btnSaveDamage.Enabled = False
        Catch ex As Exception
            lblStatus.Text = "  Save to database failed."
            MsgBox(lblStatus.Text, MsgBoxStyle.Critical, "Unable to save to database")
            lblStatus.ForeColor = Color.Red
        End Try
    End Sub

    Private Sub txtContainer_TextChanged(sender As Object, e As EventArgs) Handles txtContainer.TextChanged
        If txtContainer.Text <> "" And txtDamageMessage.Text <> "" And txtLicensePlate.Text <> "" Then
            btnSaveDamage.Enabled = True
        Else
            btnSaveDamage.Enabled = False
        End If
    End Sub

    Private Sub txtDamageMessage_TextChanged(sender As Object, e As EventArgs) Handles txtDamageMessage.TextChanged
        If txtContainer.Text <> "" And txtDamageMessage.Text <> "" And txtLicensePlate.Text <> "" Then
            btnSaveDamage.Enabled = True
        Else
            btnSaveDamage.Enabled = False
        End If
    End Sub

    Private Sub txtLicensePlate_TextChanged(sender As Object, e As EventArgs) Handles txtLicensePlate.TextChanged
        If txtContainer.Text <> "" And txtDamageMessage.Text <> "" And txtLicensePlate.Text <> "" Then
            btnSaveDamage.Enabled = True
        Else
            btnSaveDamage.Enabled = False
        End If
    End Sub
End Class


Public Class Ini

    Private _Sections As New Dictionary(Of String, Dictionary(Of String, String))
    Private _FileName As String
    ''' <summary>
    ''' </summary>
    ''' <param name="IniFileName">Drive,Path and Filname for the inifile</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal IniFileName As String)

        Dim Rd As StreamReader
        Dim Content As String
        Dim Lines() As String
        Dim Line As String
        Dim Key As String
        Dim Value As String
        Dim SectionValues As Dictionary(Of String, String) = Nothing
        Dim Name As String

        _FileName = IniFileName

        'check if the file exists
        If Not File.Exists(IniFileName) Then
            Throw New FileLoadException(String.Format("The file {0} is not found", IniFileName))
        Else
            'Read the file if present.
            Rd = New StreamReader(IniFileName)
            Content = Rd.ReadToEnd
            'Split It into lines
            Lines = Content.Split(vbCrLf)

            'Place the content in an object sructure
            For Each Line In Lines

                'Trim the line
                Line = Line.Trim
                If Line.Length <= 2 OrElse Line.Substring(0, 1) = "'" OrElse Line.Substring(0, 3).ToUpper = "REM" Then
                    'There's no valid data or it's commented out... Do nothing 
                ElseIf Line.IndexOf("[") = 0 AndAlso Line.IndexOf("]") = Line.Length - 1 Then
                    'We hit a section
                    Name = Line.Replace("]", String.Empty).Replace("[", String.Empty).Trim.ToUpper
                    SectionValues = New Dictionary(Of String, String)
                    _Sections.Add(Name.ToUpper, SectionValues)

                    'An = character as the firstcharacter is an invalid line... let's be relaxed an just ignore it.
                ElseIf Line.IndexOf("=") > 0 AndAlso SectionValues IsNot Nothing Then
                    'We hit a value line , empty line or out commented line
                    'we don't use split as that character could be part of the value as well
                    Key = Line.Substring(0, Line.IndexOf("=")).Trim
                    If Line.IndexOf("=") = Line.Length - 1 Then
                        Value = String.Empty
                    Else
                        Value = Line.Substring(Line.IndexOf("=") + 1, Line.Length - (Line.IndexOf("=") + 1)).Trim
                    End If
                    'Add the valu to 
                    SectionValues.Add(Key.ToUpper, Value)
                End If
            Next

            Rd.Close()
            Rd.Dispose()
            Rd = Nothing

        End If
    End Sub

    Public Function GetValue(ByVal Section As String, ByVal Name As String) As String

        If _Sections.ContainsKey(Section.ToUpper) Then
            Dim SectionValues As Dictionary(Of String, String) = Nothing
            SectionValues = _Sections(Section.ToUpper)
            If SectionValues.ContainsKey(Name.ToUpper) Then
                Return SectionValues(Name.ToUpper)
            End If
        End If

        Return Nothing 'if preferred return String.empty here

    End Function

    'Public Function SetValue(ByVal Section As String, ByVal Name As String, ByVal Value As String, Optional ByVal Save As Boolean = False) As Boolean
    '    Dim SectionValues As Dictionary(Of String, String) = Nothing
    '    Name = Name.ToUpper.Trim
    '    Section = Section.ToUpper.Trim
    '    If _Sections.ContainsKey(Section) Then
    '        SectionValues = _Sections(Section)
    '        If SectionValues.ContainsKey(Name) Then
    '            SectionValues.Remove(Name)
    '        End If
    '        SectionValues.Add(Name, Value)
    '    Else
    '        SectionValues = New Dictionary(Of String, String)
    '        _Sections.Add(Section, SectionValues)
    '        SectionValues.Add(Name, Value)
    '    End If

    '    If Save Then
    '        Return SaveIniFile()
    '    Else
    '        Return True
    '    End If

    'End Function


    'Public Function SaveIniFile() As Boolean

    '    Dim Rw As StreamWriter
    '    Dim SectionPair As KeyValuePair(Of String, Dictionary(Of String, String))
    '    Dim ValuePair As KeyValuePair(Of String, String)

    '    Dim Pth As String = Path.GetDirectoryName(_FileName)

    '    If Directory.Exists(Pth) Then
    '        Rw = New StreamWriter(_FileName, False)
    '        For Each SectionPair In _Sections
    '            Rw.WriteLine("[" & SectionPair.Key & "]")
    '            If SectionPair.Value IsNot Nothing Then
    '                For Each ValuePair In SectionPair.Value
    '                    Rw.WriteLine(ValuePair.Key & "=" & ValuePair.Value)
    '                Next
    '            End If
    '        Next
    '        Rw.WriteLine("")
    '        Rw.Flush()
    '        Rw.Close()
    '        Rw.Dispose()
    '        Rw = Nothing
    '        SaveIniFile = True
    '    End If

    'End Function

    'Function DeleteValue(ByVal Section As String, ByVal Name As String, Optional ByVal Save As Boolean = False) As Boolean

    '    Dim SectionValues As Dictionary(Of String, String) = Nothing

    '    Name = Name.ToUpper.Trim
    '    Section = Section.ToUpper.Trim
    '    If _Sections.ContainsKey(Section) Then
    '        SectionValues = _Sections(Section)
    '        If SectionValues.ContainsKey(Name) Then
    '            SectionValues.Remove(Name)
    '        End If
    '    End If

    '    If Save Then
    '        Return SaveIniFile()
    '    Else
    '        Return True
    '    End If

    'End Function

    'Function DeleteSection(ByVal Section As String, Optional ByVal Save As Boolean = False) As Boolean

    '    Dim SectionValues As Dictionary(Of String, String) = Nothing

    '    Section = Section.ToUpper.Trim
    '    If _Sections.ContainsKey(Section) Then
    '        _Sections.Remove(Section)
    '    End If

    '    If Save Then
    '        Return SaveIniFile()
    '    Else
    '        Return True
    '    End If

    'End Function


End Class