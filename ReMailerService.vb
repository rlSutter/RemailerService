Option Explicit On 
Option Strict On

Imports System.ServiceProcess
Imports System.Xml
Imports System.Text
Imports System.IO
Imports System.Diagnostics
Imports System.Data
Imports System.Web
Imports System.Web.Mail
Imports Microsoft.VisualBasic.ControlChars
Imports System.Timers
Imports System.Net
Imports System.Data.SqlClient
Imports System.Collections
Imports System.Reflection
Imports System.Configuration
Imports Microsoft.Win32
Imports System.Text.RegularExpressions
Imports log4net

Public Class ReMailerService
    Inherits System.ServiceProcess.ServiceBase
    Private t As Timer
    Private log As EventLog

    ' Declare parameters
    Private MyInterval, ExInternal As String        ' Timer interval
    Private UserName As String          ' Database username
    Private PassWord As String          ' Database password
    Private DBName As String            ' Database name
    Private DBServer As String          ' Database server name
    Private Debug As String             ' Debug flag - "Y" or "N"
    Private Logging As String           ' Logging flag - "Y" or "N"

    ' Service variables
    Private cloudsvc_farm As String

    ' Registry variables
    Private regKey As RegistryKey
    Private regSubKey As RegistryKey

    ' Misc variables
    Private NoInterval As Double
    Private InProcess As Boolean
    Private MachineName As String

    ' Logging declarations
    Private nagios As StreamWriter
    Private logfile, nagiosfile, nagiosmsg As String
    Private errmsg As String
    Private myeventlog As log4net.ILog
    Private mydebuglog As log4net.ILog

    ' Database declarations
    Private con As SqlConnection
    Private ucon As SqlConnection
    Private cmd As SqlCommand
    Private ucmd As SqlCommand
    Private dr As SqlDataReader
    Private SqlS As String
    Private ConnS As String
    Private returnv As Integer

#Region " Component Designer generated code "

    Public Sub New()
        MyBase.New()

        ' This call is required by the Component Designer.
        InitializeComponent()
        'Diagnostics.EventLog.WriteEntry("ReMailerService", "Called InitializeComponent")

        ' Add any initialization after the InitializeComponent() call

    End Sub

    'UserService overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' The main entry point for the process
    <MTAThread()> _
    Shared Sub Main()
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        ' More than one NT Service may run within the same process. To add
        ' another service to this process, change the following line to
        ' create a second service object. For example,
        '
        '
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New ReMailerService()}
        'Diagnostics.EventLog.WriteEntry("ReMailerService", "Called System.ServiceProcess.ServiceBase")

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    ' NOTE: The following procedure is required by the Component Designer
    ' It can be modified using the Component Designer.  
    ' Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        '
        'ReMailerService
        '
        Me.ServiceName = "ReMailerService"

    End Sub

#End Region

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' This sub is executed by the service on startup
        ' In this instance, it creates a time that executes the interval that is specified
        ' The timer executes an event when it fires.  That event is where the interval
        ' code should be placed.

        Dim objReg As New CRegistry
        Dim RegValue As Object
        Dim TestKey As RegistryKey

        ' Declare logging variables
        Try
            myeventlog = log4net.LogManager.GetLogger("EventLog")            
            mydebuglog = log4net.LogManager.GetLogger("DebugLog")
        Catch ex As Exception
            'Diagnostics.EventLog.WriteEntry("ReMailerService", "Unable to open log4net: " & ex.Message)
        End Try

        'Diagnostics.EventLog.WriteEntry("ReMailerService", "Opening registry")
        Try
            ' ============================================
            ' If the key can be created, then do so and set default values
            TestKey = Registry.LocalMachine.OpenSubKey("Software\ReMailerService")
            If TestKey Is Nothing Then
                objReg.CreateSubKey(objReg.HKeyLocalMachine, "Software\ReMailerService")
                objReg.WriteValue(objReg.HKeyLocalMachine, "Software\ReMailerService", "MyInterval", "60")
                objReg.WriteValue(objReg.HKeyLocalMachine, "Software\ReMailerService", "UserName", "[DB_USERNAME]")
                objReg.WriteValue(objReg.HKeyLocalMachine, "Software\ReMailerService", "Password", "[DB_PASSWORD]")
                objReg.WriteValue(objReg.HKeyLocalMachine, "Software\ReMailerService", "DBName", "[DATABASE_NAME]")
                objReg.WriteValue(objReg.HKeyLocalMachine, "Software\ReMailerService", "DBServer", "")
                objReg.WriteValue(objReg.HKeyLocalMachine, "Software\ReMailerService", "Debug", "Y")
                objReg.WriteValue(objReg.HKeyLocalMachine, "Software\ReMailerService", "Logging", "Y")
            End If

            ' ============================================
            ' Read parameters from Registry
            RegValue = ""
            objReg.ReadValue(objReg.HKeyLocalMachine, "Software\ReMailerService", "MyInterval", RegValue)
            MyInterval = RegValue.ToString
            NoInterval = Val(MyInterval) * 1000
            objReg.ReadValue(objReg.HKeyLocalMachine, "Software\ReMailerService", "UserName", RegValue)
            UserName = RegValue.ToString
            objReg.ReadValue(objReg.HKeyLocalMachine, "Software\ReMailerService", "Password", RegValue)
            PassWord = RegValue.ToString
            objReg.ReadValue(objReg.HKeyLocalMachine, "Software\ReMailerService", "DBName", RegValue)
            DBName = RegValue.ToString
            objReg.ReadValue(objReg.HKeyLocalMachine, "Software\ReMailerService", "DBServer", RegValue)
            DBServer = RegValue.ToString
            objReg.ReadValue(objReg.HKeyLocalMachine, "Software\ReMailerService", "Debug", RegValue)
            Debug = RegValue.ToString
            objReg.ReadValue(objReg.HKeyLocalMachine, "Software\ReMailerService", "Logging", RegValue)
            Logging = RegValue.ToString

            ' ============================================
            ' Get app.config data
            cloudsvc_farm = System.Configuration.ConfigurationManager.AppSettings.Get("cloudsvc_farm")
            If cloudsvc_farm = "" Then cloudsvc_farm = "[DEFAULT_PROXY_IP]"

            ' ============================================
            ' Setup timer
            t = New Timer()
            AddHandler t.Elapsed, AddressOf TimerFired
            With t
                .Interval = NoInterval
                .AutoReset = True
                .Enabled = True
                .Start()
            End With
            InProcess = False

            ' ============================================
            ' Log to event viewer log some parameters for debugging
            Dim myAssemblyName As New AssemblyName()
            Dim version As System.Version = Assembly.GetExecutingAssembly().GetName().Version
            Dim dispversion As String
            dispversion = Str(version.Major).Trim & "." & Str(version.Minor).Trim & "." & Str(version.Build).Trim & "." & Str(version.Revision).Trim
            'Diagnostics.EventLog.WriteEntry("ReMailerService", "Service version " & dispversion & " Starting")

            ' ============================================
            ' Open debug log file if applicable
            If Debug = "Y" Or Logging = "Y" Then
                Dim path As String
                path = "C:\Logs\"
                logfile = path & "ReMailerService.log"
                log4net.GlobalContext.Properties("LogFileName") = logfile
                log4net.Config.XmlConfigurator.Configure()
                mydebuglog.Debug("----------------------------------")
                mydebuglog.Debug("ReMailerService Trace Log Started " & Format(Now))
                If Debug = "Y" Then
                    mydebuglog.Debug("PARAMETERS")
                    mydebuglog.Debug("  Debug: " & Debug)
                    mydebuglog.Debug("  Logging: " & Logging)
                    mydebuglog.Debug("  cloudsvc_farm: " & cloudsvc_farm)
                    mydebuglog.Debug("  MyInterval: " & MyInterval)
                    mydebuglog.Debug("  UserName: " & UserName)
                    mydebuglog.Debug("  PassWord: " & PassWord)
                    mydebuglog.Debug("  DBName: " & DBName)
                    mydebuglog.Debug("  DBServer: " & DBServer & vbCrLf)
                End If
            End If

            ' Log to syslog
            myeventlog.Info("ReMailerService : Service Starting")

            ' ============================================
            ' Open Nagious result string
            nagiosmsg = ""
            errmsg = "No Error"
            nagiosfile = "C:\ReMailerService.nagios"
            If File.Exists(nagiosfile) Then
                File.Delete(nagiosfile)
            End If
            nagios = File.CreateText(nagiosfile)

            ' ============================================
            ' Open database connections
            ConnS = "server=" & DBServer & ";uid=" & UserName & ";pwd=" & PassWord & ";database=" & DBName
            errmsg = OpenDBConnection(ConnS, con, cmd)
            If errmsg <> "" Then
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Unable to open database connection with " & errmsg)
                myeventlog.Error("ReMailerService : Unable to open database connection with " & errmsg)
                OnStop()
            End If
            If Debug = "Y" Then
                mydebuglog.Debug("Opened connection to con with " & ConnS)
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Opened connection to con with " & ConnS)
            End If

            errmsg = OpenDBConnection(ConnS, ucon, ucmd)
            If errmsg <> "" Then
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Unable to open database connection with " & errmsg)
                myeventlog.Error("ReMailerService : Unable to open database connection with " & errmsg)
                OnStop()
            End If
            If Debug = "Y" Then mydebuglog.Debug("Opened connection to ucon")

            ' ============================================
            ' Get machine name - determines process information
            MachineName = Trim(Left(System.Environment.MachineName.ToString, 25))
            If Debug = "Y" Then
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Running on " & MachineName)
                mydebuglog.Debug(vbCrLf & "Running on " & MachineName)
            End If

            ' ============================================
            ' Purge lock table of any outstanding work - need to clean slate this
            SqlS = "DELETE FROM [DATABASE_NAME].dbo.MSGS_LOCK " & _
                "WHERE MACHINE='" & MachineName & "'"
            If Debug = "Y" Then
                mydebuglog.Debug(vbCrLf & "QUERY: Remove extraneous lock table entries: " & vbCrLf & SqlS & vbCrLf)
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Query: " & SqlS)
            End If
            Try
                ucmd.CommandText = SqlS
                returnv = ucmd.ExecuteNonQuery()
                If returnv = 0 Then
                    mydebuglog.Debug("No extraneous lock entries to remove")
                    'Diagnostics.EventLog.WriteEntry("ReMailerService", "No extraneous lock entries to remove")
                End If
            Catch ex As Exception
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Unable to Remove extraneous lock table entries " & ex.Message)
                myeventlog.Error("ReMailerService : Unable to Remove extraneous lock table entries " & ex.Message)
            End Try

        Catch obug As Exception
            LogEvent(obug.Message)
            myeventlog.Error("ReMailerService : Fatal error " & obug.Message)
            Throw obug
        End Try
    End Sub

    Protected Overrides Sub OnStop()
        ' This sub is executed when the service stops
        Try
            ' ============================================
            ' Close database connections
            errmsg = CloseDBConnection(con, cmd, dr)
            If errmsg <> "" Then
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Unable to close con database connection with " & errmsg)
                myeventlog.Error("ReMailerService : Unable to close con database connection with " & errmsg)
            End If
            If Debug = "Y" Then mydebuglog.Debug("Closed connection to con")

            errmsg = CloseDBConnection(ucon, ucmd, dr)
            If errmsg <> "" Then
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Unable to close ucon database connection with " & errmsg)
                myeventlog.Error("ReMailerService : Unable to close ucon database connection with " & errmsg)
            End If
            If Debug = "Y" Then mydebuglog.Debug("Closed connection to ucon")

            'Try
            'If Not con Is Nothing Then con.Close()
            'If Not ucon Is Nothing Then ucon.Close()
            'If Not dr Is Nothing Then dr.Close()
            'con.Dispose()
            'con = Nothing
            'ucon.Dispose()
            'ucon = Nothing
            'dr = Nothing
            'cmd.Dispose()
            'cmd = Nothing
            'ucmd.Dispose()
            'ucmd = Nothing
            'Catch ex As Exception
            'Diagnostics.EventLog.WriteEntry("ReMailerService", "Unable to close database connection with " & ex.Message)
            'myeventlog.Error("ReMailerService : Unable to close database connection with " & ex.Message)
            'End Try

            ' ============================================
            ' Close diagnostic logging
            If Debug = "Y" Or Logging = "Y" Then
                mydebuglog.Debug(vbCrLf & "Service Results: " & Trim(errmsg))
                mydebuglog.Debug("ReMailerService Trace Log Ended " & Format(Now))
                mydebuglog.Debug("----------------------------------")
            End If

            ' ============================================
            ' Close Nagios string
            If Not nagios Is Nothing Then
                If errmsg = "No Error" Or errmsg = "No messages found to process" Then
                    nagiosmsg = "Success on " & Format(Now)
                Else
                    nagiosmsg = "Failure on " & Format(Now) & " " & errmsg
                End If
                nagios.WriteLine(Trim(nagiosmsg))
                nagios.Flush()
                nagios.Close()
            End If

            ' ============================================
            ' Stop timer
            'If Debug = "Y" Then Diagnostics.EventLog.WriteEntry("ReMailerService", "Service Stopping: " & EventLogEntryType.Information)
            myeventlog.Info("ReMailerService : Service Stopping " & EventLogEntryType.Information)
            If Not t Is Nothing Then
                Try
                    t.Stop()
                    t.Dispose()
                Catch ex As Exception
                    'Diagnostics.EventLog.WriteEntry("ReMailerService", "Unable to stop timer")
                    myeventlog.Error("ReMailerService : Unable to stop timer with " & ex.Message)
                End Try
            End If
        Catch obug As Exception
            'Diagnostics.EventLog.WriteEntry("ReMailerService", obug.Message)
            myeventlog.Error("ReMailerService : Fatal error " & obug.Message)
        End Try
    End Sub

    Private Sub TimerFired(ByVal sender As Object, ByVal e As ElapsedEventArgs)
        ' The timer fired.  Call CheckEmailQueue
        If Not InProcess Then CheckEmailQueue()
    End Sub

    Private Sub CheckEmailQueue()
        ' This function does the following:
        '   1. Locates any email messages to be sent in the scanner.MESSAGES table
        '   2. Makes a call to the SendMail web service for each message
        ' Only execute if InProcess is set to False

        ' Web service declarations
        Dim response As HttpWebResponse = Nothing
        Dim Completed As String = ""
        Dim LineNumber As Integer
        'Dim SendEmailWS As New com.certegrity.cloudsvc.Service
        Dim http As New simplehttp()
        Dim results As String

        ' General declarations
        Dim i As Integer
        Dim NumInterval As Integer

        ' Email variable declarations
        Dim mSEND_TO As String
        Dim mSEND_FROM As String
        Dim mSUBJECT As String
        Dim mBODY As String
        Dim mATTACHMENT As String
        Dim mCC As String
        Dim mBCC As String
        Dim mHTML As String
        Dim mTO_ID As String
        Dim mATTACH_DOC_ID As String
        Dim mATTACH_TYPE As String
        Dim mFROM_NM As String
        Dim mFROM_ID As String
        Dim mMS_IDENT As String
        Dim uMS_IDENT(500, 3) As String

        ' XML declarations
        Dim wp As String
        Try
            LineNumber = 0
            InProcess = True

            ' ============================================
            ' Get the messages to process 
            If Val(MyInterval) > 0 Then
                ExInternal = Trim(Str(Val(MyInterval) * 1))
            Else
                ExInternal = "5"
            End If
            NumInterval = (CInt(Val(MyInterval)) * 2) + 1
            ReDim uMS_IDENT(NumInterval, 3)            ' resize and reinitialize array
            If Debug = "Y" Then
                mydebuglog.Debug(vbCrLf & "Checking for " & ExInternal.ToString & " new messages at " & Format(Now) & " on " & MachineName)
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Checking for new messages at " & Format(Now))
            End If

            ' ============================================
            ' Check to see if we have records in the lock table already
            Dim CurRecCnt As Integer
            CurRecCnt = 0
            SqlS = "SELECT COUNT(*) FROM [DATABASE_NAME].dbo.MSGS_LOCK WHERE MACHINE='" & MachineName & "'"
            If Debug = "Y" Then
                mydebuglog.Debug(vbCrLf & "QUERY: Get count of email messages in lock table: " & vbCrLf & SqlS & vbCrLf)
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Query: " & SqlS)
            End If
            Try
                cmd = New SqlCommand(SqlS, con)
                dr = cmd.ExecuteReader()
                While dr.Read()
                    CurRecCnt = CInt(dr(0))
                End While
                dr.Close()
            Catch ex As Exception
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Unable to query database")
                myeventlog.Error("ReMailerService : Unable to query database with " & ex.Message)

                ' Close database connections
                errmsg = CloseDBConnection(con, cmd, dr)
                errmsg = CloseDBConnection(ucon, ucmd, dr)

                ' Open database connections
                errmsg = OpenDBConnection(ConnS, con, cmd)
                errmsg = OpenDBConnection(ConnS, ucon, ucmd)

                GoTo CloseOut
            End Try
            If Debug = "Y" Then
                mydebuglog.Debug("   ...Lock table message count: " & CurRecCnt.ToString & vbCrLf)
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Lock table message count: " & CurRecCnt.ToString)
            End If

            ' ============================================
            ' Copy records to process to the lock table if none are currently in the table for this machine
            If CurRecCnt = 0 Then
                SqlS = "INSERT [DATABASE_NAME].dbo.MSGS_LOCK " & _
                    "(SEND_TO, SEND_FROM, SUBJECT, BODY, SENT_FLG, CREATED, SENT, " & _
                    "ATTACHMENT, CC, BCC, HTML, TO_ID, SRC_TYPE, DEFER_UNTIL, EXPIRES, " & _
                    "ATTACH_DOC_ID, READ_FLG, READ_DT, FROM_NM, FROM_ID, SRC_ID, MS_IDENT, MACHINE, ATTACH_TYPE, ATTEMPTS) " & _
                    "SELECT TOP " & ExInternal & " SEND_TO, SEND_FROM, SUBJECT, BODY, SENT_FLG, CREATED, SENT, " & _
                    "ATTACHMENT, CC, BCC, HTML, TO_ID, SRC_TYPE, DEFER_UNTIL, EXPIRES, " & _
                    "ATTACH_DOC_ID, READ_FLG, READ_DT, FROM_NM, FROM_ID, SRC_ID, MS_IDENT,'" & MachineName & "', ATTACH_TYPE, 1 " & _
                    "FROM [DATABASE_NAME].dbo.MESSAGES M " & _
                    "WHERE SENT_FLG='N' " & _
                    "AND (DEFER_UNTIL IS NULL OR GETDATE()>DEFER_UNTIL) " & _
                    "AND BODY IS NOT NULL " & _
                    "AND SEND_TO IS NOT NULL AND SEND_TO<>'' " & _
                    "AND SEND_FROM IS NOT NULL AND SEND_FROM<>'' " & _
                    "AND NOT EXISTS (" & _
                    "SELECT MS_IDENT FROM [DATABASE_NAME].dbo.MSGS_LOCK WHERE MS_IDENT=M.MS_IDENT" & _
                    ") " & _
                    "ORDER BY CREATED DESC"
                If Debug = "Y" Then
                    mydebuglog.Debug(vbCrLf & "QUERY: Select email messages to process into lock table: " & vbCrLf & SqlS & vbCrLf)
                    'Diagnostics.EventLog.WriteEntry("ReMailerService", "Query: " & SqlS)
                End If
                Try
                    ucmd.CommandText = SqlS
                    returnv = ucmd.ExecuteNonQuery()
                    If returnv = 0 Then
                        'mydebuglog.Debug("No messages to process")
                        'Diagnostics.EventLog.WriteEntry("ReMailerService", "No messages to process")
                        GoTo CloseOut
                    End If
                Catch ex As Exception
                    'Diagnostics.EventLog.WriteEntry("ReMailerService", "Unable to Select email messages to process into lock table")
                    myeventlog.Error("ReMailerService : Unable to select messages to process, " & ex.Message)
                    GoTo CloseOut
                End Try
            End If

            ' ============================================
            ' Query the lock table for work to do
            SqlS = "SELECT SEND_TO, SEND_FROM, SUBJECT, BODY, SENT_FLG, CREATED, SENT, " & _
            "ATTACHMENT, CC, BCC, HTML, TO_ID, SRC_TYPE, DEFER_UNTIL, EXPIRES, " & _
            "ATTACH_DOC_ID, READ_FLG, READ_DT, FROM_NM, FROM_ID, SRC_ID, MS_IDENT, ATTEMPTS, ATTACH_TYPE " & _
            "FROM [DATABASE_NAME].dbo.MSGS_LOCK WHERE MACHINE='" & MachineName & "'"
            If Debug = "Y" Then
                mydebuglog.Debug(vbCrLf & "QUERY: Get email messages to process from lock table: " & vbCrLf & SqlS & vbCrLf)
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Query: " & SqlS)
            End If
            Try
                cmd = New SqlCommand(SqlS, con)
                dr = cmd.ExecuteReader()
            Catch ex As Exception
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Unable to query database")
                myeventlog.Error("ReMailerService : Unable to query database, " & ex.Message)
                GoTo CloseOut
            End Try
            If dr Is Nothing Then
                errmsg = "Database access error"
                GoTo CloseOut
            End If

            ' ============================================
            ' Go through records
            While dr.Read()
                LineNumber = LineNumber + 1

                ' -----
                ' Get the specifics of the message found
                mSEND_TO = Trim(dr(0).ToString)
                mSEND_TO = CleanString(mSEND_TO)
                mSEND_FROM = Trim(dr(1).ToString)
                mSEND_FROM = CleanString(mSEND_FROM)
                mSUBJECT = Trim(dr(2).ToString)
                mBODY = Trim(dr(3).ToString)
                mATTACHMENT = dr(7).ToString
                If InStr(mATTACHMENT, "siebeldb\") > 0 Then
                    mATTACHMENT = mATTACHMENT.Replace("siebeldb\", "siebel\")
                End If
                mCC = Trim(dr(8).ToString)
                mBCC = Trim(dr(9).ToString)
                mHTML = dr(10).ToString
                If mHTML = "Y" Then
                    mHTML = "HTML"
                    mBODY = "<![CDATA[" & mBODY & " ]]>"  ' Reformat the body tag appropriately
                Else
                    mHTML = ""
                End If
                mTO_ID = Trim(dr(11).ToString)
                mATTACH_DOC_ID = dr(15).ToString
                mATTACH_TYPE = dr(23).ToString
                If Debug = "Y" Then
                    mydebuglog.Debug("Message # " & LineNumber.ToString & " Started")
                    mydebuglog.Debug("  > mSEND_FROM: " & mSEND_FROM)
                    mydebuglog.Debug("  > mSEND_TO: " & mSEND_TO)
                    mydebuglog.Debug("  > mSUBJECT: " & mSUBJECT)
                    mydebuglog.Debug("  > mTO_ID: " & mTO_ID)
                    mydebuglog.Debug("  > mATTACH_DOC_ID: " & mATTACH_DOC_ID)
                    mydebuglog.Debug("  > mATTACH_TYPE: " & mATTACH_TYPE)
                End If
                If mATTACH_TYPE <> "dms" And mATTACH_TYPE <> "reports" Then mATTACH_TYPE = ""
                mFROM_NM = dr(18).ToString
                mFROM_ID = Trim(dr(19).ToString)
                mMS_IDENT = Trim(dr(21).ToString)
                uMS_IDENT(LineNumber, 3) = dr(22).ToString      ' Store number of attempts

                ' -----
                ' Create XML document for message
                wp = "<EMailMessageList><EMailMessage><debug>" & Debug & "</debug>"
                wp = wp & "<database>U</database>"
                wp = wp & "<Id>" & mMS_IDENT & "</Id>"
                wp = wp & "<SourceId></SourceId>"
                wp = wp & "<From>" & mSEND_FROM & "</From>"
                wp = wp & "<FromId>" & mFROM_ID & "</FromId>"
                wp = wp & "<FromName>" & mFROM_NM & "</FromName>"
                wp = wp & "<To>" & mSEND_TO & "</To>"
                wp = wp & "<ToId>" & mTO_ID & "</ToId>"
                wp = wp & "<Cc>" & mCC & "</Cc>"
                wp = wp & "<Bcc>" & mBCC & "</Bcc>"
                wp = wp & "<ReplyTo>" & mSEND_FROM & "</ReplyTo>"
                wp = wp & "<Subject>" & mSUBJECT & "</Subject>"
                wp = wp & "<Body></Body>"
                wp = wp & "<Format>" & mHTML & "</Format>"
                wp = wp & "<AttachmentList>"
                If mATTACH_DOC_ID = "" Or mATTACH_TYPE = "" Then
                    wp = wp & "<Attachment>" & mATTACHMENT & "</Attachment>"
                Else
                    wp = wp & "<Attachment AttachType=""" & mATTACH_TYPE & """ AttachId=""" & mATTACH_DOC_ID & """></Attachment>"
                End If
                wp = wp & "</AttachmentList></EMailMessage></EMailMessageList>"
                If Debug = "Y" Then mydebuglog.Debug("  > wp " & wp)

                ' -----
                ' Send the web request
                Try
                    'Completed = SendEmailWS.SendMail(wp).ToString
                    'If Completed = "False" Then Completed = "Error"
                    wp = "http://[WEB_SERVICE_DOMAIN]/basic/service.asmx/SendMail?sXML=" & HttpUtility.UrlEncode(wp)
                    If Debug = "Y" Then
                        mydebuglog.Debug("Msg #: " & LineNumber.ToString & "  Sending: " & wp)
                        'Diagnostics.EventLog.WriteEntry("ReMailerService", "Msg #: " & LineNumber.ToString & "   Sending: " & wp)
                    End If
                    results = http.geturl(wp, cloudsvc_farm, 80, "", "")
                    If Debug = "Y" Then mydebuglog.Debug("  > results " & results)
                    If InStr(results, "False") > 0 Then
                        Completed = "Error"
                    Else
                        Completed = "Success"
                    End If
                Catch ex As Exception
                    Completed = "Error"
                    myeventlog.Error("ReMailerService : Error executing SendMail web service, " & ex.Message)
                End Try

                ' -----
                ' Debug output of response
                If Debug = "Y" Or Logging = "Y" Then
                    mydebuglog.Debug("Message # " & Str(LineNumber) & " ID: " & mMS_IDENT & "  Sent: " & Completed & vbCrLf)
                    'Diagnostics.EventLog.WriteEntry("ReMailerService", "Message # " & Str(LineNumber) & "  ID: " & mMS_IDENT & "  Sent: " & Completed)
                End If
                myeventlog.Info("ReMailerService : Message ID: " & mMS_IDENT & " sent to " & mSEND_TO & ". Sent: " & Completed)

                ' -----
                ' Remove record from the lock table if sent
                uMS_IDENT(LineNumber, 1) = mMS_IDENT            ' Store identity
                uMS_IDENT(LineNumber, 2) = Completed            ' Store status
            End While
            dr.Close()

            ' ============================================
            ' Update the lock table to remove sent messages
            If Debug = "Y" Then
                mydebuglog.Debug(vbCrLf & "QUERY: Msgs Processed: " & LineNumber.ToString)
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Msgs Processed: " & LineNumber.ToString)
            End If
            For i = 1 To LineNumber
                If Debug = "Y" Then
                    mydebuglog.Debug(vbCrLf & "Checking message to remove: " & uMS_IDENT(i, 1))
                End If
                If uMS_IDENT(i, 2) <> "Error" Then
                    SqlS = "DELETE FROM [DATABASE_NAME].dbo.MSGS_LOCK WHERE MS_IDENT=" & uMS_IDENT(i, 1)
                    If Debug = "Y" Then
                        mydebuglog.Debug(" .. QUERY: Removing processed msg from lock table: " & vbCrLf & SqlS)
                        'Diagnostics.EventLog.WriteEntry("ReMailerService", "Query: " & SqlS)
                    End If
                    Try
                        ucmd.CommandText = SqlS
                        returnv = ucmd.ExecuteNonQuery()
                    Catch ex As Exception
                        'Diagnostics.EventLog.WriteEntry("ReMailerService", "Unable to Remove processed msg from lock table " & ex.Message)
                        myeventlog.Error("ReMailerService : Unable to remove processed message, " & ex.Message)
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Unable to Remove processed msg from lock table " & ex.Message)
                    End Try
                Else
                    SqlS = "UPDATE [DATABASE_NAME].dbo.MSGS_LOCK SET ATTEMPTS=ATTEMPTS+1 WHERE MS_IDENT=" & uMS_IDENT(i, 1)
                    If Debug = "Y" Then
                        mydebuglog.Debug(" .. QUERY: Updating attempts in lock table: " & vbCrLf & SqlS)
                        'Diagnostics.EventLog.WriteEntry("ReMailerService", "Query: " & SqlS)
                    End If
                    Try
                        ucmd.CommandText = SqlS
                        returnv = ucmd.ExecuteNonQuery()
                    Catch ex As Exception
                        'Diagnostics.EventLog.WriteEntry("ReMailerService", "Unable to Remove processed msg from lock table " & ex.Message)
                        myeventlog.Error("ReMailerService : Unable to remove processed message, " & ex.Message)
                        If Debug = "Y" Then mydebuglog.Debug(vbCrLf & "Unable to Remove processed msg from lock table " & ex.Message)
                    End Try                
                End If
            Next

            ' ============================================
            ' Move messages that can't be processed to the "dead" table
            SqlS = "INSERT [DATABASE_NAME].dbo.MSGS_DEAD " & _
                "(SEND_TO, SEND_FROM, SUBJECT, BODY, SENT_FLG, CREATED, SENT, " & _
                "ATTACHMENT, CC, BCC, HTML, TO_ID, SRC_TYPE, DEFER_UNTIL, EXPIRES, " & _
                "ATTACH_DOC_ID, READ_FLG, READ_DT, FROM_NM, FROM_ID, SRC_ID, MS_IDENT, MACHINE, ATTACH_TYPE) " & _
                "SELECT SEND_TO, SEND_FROM, SUBJECT, BODY, SENT_FLG, CREATED, SENT, " & _
                "ATTACHMENT, CC, BCC, HTML, TO_ID, SRC_TYPE, DEFER_UNTIL, EXPIRES, " & _
                "ATTACH_DOC_ID, READ_FLG, READ_DT, FROM_NM, FROM_ID, SRC_ID, MS_IDENT, MACHINE, ATTACH_TYPE " & _
                "FROM [DATABASE_NAME].dbo.MSGS_LOCK ML " & _
                "WHERE MACHINE='" & MachineName & "' AND ATTEMPTS>3 AND NOT EXISTS (" & _
                "SELECT MS_IDENT FROM [DATABASE_NAME].dbo.MSGS_DEAD WHERE MS_IDENT=ML.MS_IDENT)"
            If Debug = "Y" Then
                mydebuglog.Debug(vbCrLf & "QUERY: Move unprocessed records to dead table: " & vbCrLf & SqlS)
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Query: " & SqlS)
            End If
            Try
                ucmd.CommandText = SqlS
                returnv = ucmd.ExecuteNonQuery()
                If returnv = 0 Then
                    If Debug = "Y" Then mydebuglog.Debug(" .. No unprocessed records to move to the dead" & vbCrLf)
                    'Diagnostics.EventLog.WriteEntry("ReMailerService", "No unprocessed records to move to the dead table")
                End If
            Catch ex As Exception
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Unable to Move unprocessed records to dead table " & ex.Message)
                myeventlog.Error("ReMailerService : Unable to move processed message, " & ex.Message)
                If Debug = "Y" Then mydebuglog.Debug(" .. Unable to Move unprocessed records to dead table " & ex.Message & vbCrLf)
            End Try

            ' ============================================
            ' Update records in MESSAGES that can't be sent
            SqlS = "UPDATE [DATABASE_NAME].dbo.MESSAGES SET SENT_FLG='E' " & _
            "WHERE MS_IDENT IN (SELECT MS_IDENT FROM [DATABASE_NAME].dbo.MSGS_LOCK " & _
            "WHERE MACHINE='" & MachineName & "' AND ATTEMPTS>3)"
            If Debug = "Y" Then
                mydebuglog.Debug("QUERY: Mark messages that can't be sent as errors: " & vbCrLf & SqlS & vbCrLf)
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Query: " & SqlS)
            End If
            Try
                ucmd.CommandText = SqlS
                returnv = ucmd.ExecuteNonQuery()
            Catch ex As Exception
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Unable to Mark messages that can't be sent as errors " & ex.Message)
                myeventlog.Error("ReMailerService : Unable to mark messages that can't be send, " & ex.Message)
                If Debug = "Y" Then mydebuglog.Debug(" .. Unable to Mark messages that can't be sent as errors " & ex.Message)
            End Try

            ' ============================================
            ' Remove dead messages for this machine from the lock table
            SqlS = "DELETE FROM [DATABASE_NAME].dbo.MSGS_LOCK WHERE MACHINE='" & MachineName & "'  AND ATTEMPTS>3"
            If Debug = "Y" Then
                mydebuglog.Debug("QUERY: Remove dead messages from the lock table: " & vbCrLf & SqlS & vbCrLf)
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Query: " & SqlS)
            End If
            Try
                ucmd.CommandText = SqlS
                returnv = ucmd.ExecuteNonQuery()
            Catch ex As Exception
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Unable to Remove dead messages from the lock table " & ex.Message)
                myeventlog.Error("ReMailerService : Unable to remove dead messages, " & ex.Message)
                If Debug = "Y" Then mydebuglog.Debug(" .. Unable to Remove dead messages from the lock table " & ex.Message)
            End Try

CloseOut:
            ' ============================================
            ' Close objects
            Try
                'SendEmailWS = Nothing
                uMS_IDENT = Nothing
            Catch ex As Exception
                'Diagnostics.EventLog.WriteEntry("ReMailerService", "Unable to close database connection")
                myeventlog.Error("ReMailerService : Unable to close web service object, " & ex.Message)
            End Try
            InProcess = False

        Catch wex As WebException
            If Not wex.Response Is Nothing Then
                Dim errorResponse As HttpWebResponse = Nothing
                Try
                    errorResponse = DirectCast(wex.Response, HttpWebResponse)
                    ' Save error description for log
                    errmsg = errorResponse.StatusDescription
                    'Diagnostics.EventLog.WriteEntry("ReMailerService", "WEB SERVICE ERROR RECEIVED:  " & errmsg)
                    myeventlog.Error("ReMailerService : Web service error, " & errmsg)
                Finally
                    If Not errorResponse Is Nothing Then errorResponse.Close()
                End Try
            End If
            'Diagnostics.EventLog.WriteEntry("ReMailerService", errmsg)

            'SendEmailWS = Nothing
            InProcess = False
        End Try

    End Sub

    Public Function EmailAddressCheck(ByVal emailAddress As String) As Boolean
        ' Validate email address

        Dim pattern As String = "^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
        Dim emailAddressMatch As Match = Regex.Match(emailAddress, pattern)
        If emailAddressMatch.Success Then
            EmailAddressCheck = True
        Else
            EmailAddressCheck = False
        End If

    End Function

    Public Shared Sub LogEvent(ByVal sMessage As String)
        ' Write error into the event viewer
        Try
            Dim oEventLog As EventLog = New EventLog("Application")
            If Not Diagnostics.EventLog.SourceExists("ReMailerService") Then
                Diagnostics.EventLog.CreateEventSource("ReMailerService", "Application")
            End If
            'Diagnostics.EventLog.WriteEntry("ReMailerService", sMessage, System.Diagnostics.EventLogEntryType.Error)
        Catch e As Exception
        End Try
    End Sub

    Private Sub AddXMLChild(ByVal xmldoc As XmlDocument, ByVal root As XmlElement, _
        ByVal childname As String, ByVal childvalue As String)
        Dim resultsItem As System.Xml.XmlElement

        resultsItem = xmldoc.CreateElement(childname)
        resultsItem.InnerText = childvalue
        root.AppendChild(resultsItem)
    End Sub

    Function ChkString(ByVal Instring As String) As String
        ' Generic function to create a string that can be used in a SQL INSERT statement
        Dim temp, outstring As String
        Dim i As Integer
        temp = Instring
        outstring = ""
        For i = 1 To Len(temp$)
            If Mid(temp, i, 1) = "'" Then
                outstring = outstring & "''"
            Else
                outstring = outstring & Mid(temp, i, 1)
            End If
        Next
        ChkString = outstring
    End Function

    Function CleanString(ByVal Instring As String) As String
        ' Replaces spaces with "+" signs in key fields
        Dim temp As String
        Dim outstring, tocheck As String
        Dim i As Integer

        If Len(Instring) = 0 Or Instring Is Nothing Then
            CleanString = ""
            Exit Function
        End If
        temp = Instring.ToString
        outstring = ""
        For i = 1 To Len(temp)
            tocheck = Mid(temp, i, 1)
            If Asc(tocheck) > 31 And Asc(tocheck) < 127 Then
                Select Case Asc(tocheck)
                    Case 34     ' "
                    Case 38     ' &
                    Case Else
                        outstring = outstring & tocheck
                End Select
            End If
        Next
        CleanString = outstring
    End Function

    Function ClearString(ByVal Instring As String) As String
        ' Generic function to remove CRLFs from a string
        Dim temp, outstring, tstr As String
        Dim i As Integer
        temp = Instring
        outstring = ""
        For i = 1 To Len(temp$)
            tstr = Mid(temp, i, 1)
            Select Case tstr
                Case Chr(9)
                    outstring = outstring
                Case Chr(10)
                    outstring = outstring
                Case Chr(13)
                    outstring = outstring
                Case Else
                    outstring = outstring & tstr
            End Select
        Next
        ClearString = outstring
    End Function

    ' =================================================
    ' DATABASE FUNCTIONS
    Public Function OpenDBConnection(ByVal ConnS As String, ByRef con As SqlConnection, ByRef cmd As SqlCommand) As String
        ' Function to open a database connection with extreme error-handling
        ' Returns an error message if unable to open the connection
        Dim SqlS As String
        SqlS = ""
        OpenDBConnection = ""

        Try
            con = New SqlConnection(ConnS)
            con.Open()
            If Not con Is Nothing Then
                Try
                    cmd = New SqlCommand(SqlS, con)
                    cmd.CommandTimeout = 300
                Catch ex2 As Exception
                    OpenDBConnection = "Error opening the command string: " & ex2.ToString
                End Try
            End If
        Catch ex As Exception
            If con.State <> Data.ConnectionState.Closed Then con.Dispose()
            ConnS = ConnS & ";Pooling=false"
            Try
                con = New SqlConnection(ConnS)
                con.Open()
                If Not con Is Nothing Then
                    Try
                        cmd = New SqlCommand(SqlS, con)
                        cmd.CommandTimeout = 300
                    Catch ex2 As Exception
                        OpenDBConnection = "Error opening the command string: " & ex2.ToString
                    End Try
                End If
            Catch ex2 As Exception
                OpenDBConnection = "Unable to open database connection for connection string: " & ConnS & vbCrLf & "Windows error: " & vbCrLf & ex2.ToString & vbCrLf
            End Try
        End Try

    End Function

    Public Function CloseDBConnection(ByRef con As SqlConnection, ByRef cmd As SqlCommand, ByRef dr As SqlDataReader) As String
        ' This function closes a database connection safely
        Dim ErrMsg As String
        ErrMsg = ""

        ' Handle datareader
        Try
            dr.Close()
        Catch ex As Exception
        End Try
        Try
            dr = Nothing
        Catch ex As Exception
        End Try

        ' Handle command
        Try
            cmd.Dispose()
        Catch ex As Exception
        End Try
        Try
            cmd = Nothing
        Catch ex As Exception
        End Try

        ' Handle connection
        Try
            con.Close()
        Catch ex As Exception
        End Try
        Try
            SqlConnection.ClearPool(con)
        Catch ex As Exception
        End Try
        Try
            con.Dispose()
        Catch ex As Exception
        End Try
        Try
            con = Nothing
        Catch ex As Exception
        End Try

        ' Exit
        Return ErrMsg
    End Function

    ' =================================================
    ' HTTP PROXY CLASS
    Class simplehttp
        Public Function geturl(ByVal url As String, ByVal proxyip As String, ByVal port As Integer, ByVal proxylogin As String, ByVal proxypassword As String) As String
            Dim resp As HttpWebResponse
            Dim req As HttpWebRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
            req.UserAgent = "Mozilla/5.0?"
            req.AllowAutoRedirect = True
            req.ReadWriteTimeout = 5000
            req.CookieContainer = New CookieContainer()
            req.Referer = ""
            req.Headers.[Set]("Accept-Language", "en,en-us")
            Dim stream_in As StreamReader

            Dim proxy As New WebProxy(proxyip, port)
            'if proxylogin is an empty string then don t use proxy credentials (open proxy)
            If proxylogin = "" Then
                proxy.Credentials = New NetworkCredential(proxylogin, proxypassword)
            End If
            req.Proxy = proxy

            Dim response As String = ""
            Try
                resp = DirectCast(req.GetResponse(), HttpWebResponse)
                stream_in = New StreamReader(resp.GetResponseStream())
                response = stream_in.ReadToEnd()
                stream_in.Close()
            Catch ex As Exception
            End Try
            Return response
        End Function

        Public Function getposturl(ByVal url As String, ByVal postdata As String, ByVal proxyip As String, ByVal port As Short, ByVal proxylogin As String, ByVal proxypassword As String) As String
            Dim resp As HttpWebResponse
            Dim req As HttpWebRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
            req.UserAgent = "Mozilla/5.0?"
            req.AllowAutoRedirect = True
            req.ReadWriteTimeout = 5000
            req.CookieContainer = New CookieContainer()
            req.Method = "POST"
            req.ContentType = "application/x-www-form-urlencoded"
            req.ContentLength = postdata.Length
            req.Referer = ""

            Dim proxy As New WebProxy(proxyip, port)
            'if proxylogin is an empty string then don t use proxy credentials (open proxy)
            If proxylogin = "" Then
                proxy.Credentials = New NetworkCredential(proxylogin, proxypassword)
            End If
            req.Proxy = proxy

            Dim stream_out As New StreamWriter(req.GetRequestStream(), System.Text.Encoding.ASCII)
            stream_out.Write(postdata)
            stream_out.Close()
            Dim response As String = ""

            Try
                resp = DirectCast(req.GetResponse(), HttpWebResponse)
                Dim resStream As Stream = resp.GetResponseStream()
                Dim stream_in As New StreamReader(req.GetResponse().GetResponseStream())
                response = stream_in.ReadToEnd()
                stream_in.Close()
            Catch ex As Exception
            End Try
            Return response
        End Function
    End Class

End Class
