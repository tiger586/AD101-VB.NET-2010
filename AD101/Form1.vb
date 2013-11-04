Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

'聲明：

Public Class Form1

  Public Structure AD101DEVICEPARAMETER
    Public nRingOn As Integer
    Public nRingOff As Integer
    Public nHookOn As Integer
    Public nHookOff As Integer
    Public nStopCID As Integer
    Public nNoLine As Integer
    ' Add this parameter in new AD101(MCU Version is 6.0)
  End Structure

  <DllImport("AD101Device.dll", EntryPoint:="AD101_InitDevice")> _
  Public Shared Function AD101_InitDevice(ByVal hWnd As Integer) As Integer
  End Function

  <DllImport("AD101Device.dll", EntryPoint:="AD101_GetDevice")> _
  Public Shared Function AD101_GetDevice() As Integer
  End Function

  <DllImport("AD101Device.dll", EntryPoint:="AD101_GetCPUVersion", CharSet:=CharSet.Ansi)> _
  Public Shared Function AD101_GetCPUVersion(ByVal nLine As Integer, ByVal szCPUVersion As StringBuilder) As Integer
  End Function

  ' Start reading cpu id of device 
  <DllImport("AD101Device.dll", EntryPoint:="AD101_ReadCPUID")> _
  Public Shared Function AD101_ReadCPUID(ByVal nLine As Integer) As Integer
  End Function

  ' Get readed cpu id of device 
  <DllImport("AD101Device.dll", EntryPoint:="AD101_GetCPUID", CharSet:=CharSet.Ansi)> _
  Public Shared Function AD101_GetCPUID(ByVal nLine As Integer, ByVal szCPUID As StringBuilder) As Integer
  End Function



  ' Get caller id number  
  <DllImport("AD101Device.dll", EntryPoint:="AD101_GetCallerID", CharSet:=CharSet.Ansi)> _
  Public Shared Function AD101_GetCallerID(ByVal nLine As Integer, ByVal szCallerIDBuffer As StringBuilder, ByVal szName As StringBuilder, ByVal szTime As StringBuilder) As Integer
  End Function

  ' Get dialed number 
  <DllImport("AD101Device.dll", EntryPoint:="AD101_GetDialDigit", CharSet:=CharSet.Ansi)> _
  Public Shared Function AD101_GetDialDigit(ByVal nLine As Integer, ByVal szDialDigitBuffer As StringBuilder) As Integer
  End Function


  ' Get collateral phone dialed number 
  <DllImport("AD101Device.dll", EntryPoint:="AD101_GetCollateralDialDigit", CharSet:=CharSet.Ansi)> _
  Public Shared Function AD101_GetCollateralDialDigit(ByVal nLine As Integer, ByVal szDialDigitBuffer As StringBuilder) As Integer
  End Function



  ' Get last line state 
  <DllImport("AD101Device.dll", EntryPoint:="AD101_GetState")> _
  Public Shared Function AD101_GetState(ByVal nLine As Integer) As Integer
  End Function

  ' Get ring count
  <DllImport("AD101Device.dll", EntryPoint:="AD101_GetRingIndex")> _
  Public Shared Function AD101_GetRingIndex(ByVal nLine As Integer) As Integer
  End Function

  ' Get talking time
  <DllImport("AD101Device.dll", EntryPoint:="AD101_GetTalkTime")> _
  Public Shared Function AD101_GetTalkTime(ByVal nLine As Integer) As Integer
  End Function


  <DllImport("AD101Device.dll", EntryPoint:="AD101_GetParameter")> _
  Public Shared Function AD101_GetParameter(ByVal nLine As Integer, ByRef tagParameter As AD101DEVICEPARAMETER) As Integer
  End Function

  <DllImport("AD101Device.dll", EntryPoint:="AD101_ReadParameter")> _
  Public Shared Function AD101_ReadParameter(ByVal nLine As Integer) As Integer
  End Function

  ' Set systematic parameter  
  <DllImport("AD101Device.dll", EntryPoint:="AD101_SetParameter")> _
  Public Shared Function AD101_SetParameter(ByVal nLine As Integer, ByRef tagParameter As AD101DEVICEPARAMETER) As Integer
  End Function

  ' Free devices 

  <DllImport("AD101Device.dll", EntryPoint:="AD101_FreeDevice")> _
  Public Shared Sub AD101_FreeDevice()
  End Sub

  ' Get current AD101 device count
  <DllImport("AD101Device.dll", EntryPoint:="AD101_GetCurDevCount")> _
  Public Shared Function AD101_GetCurDevCount() As Integer
  End Function

  ' Change handle of window that uses to receive message
  <DllImport("AD101Device.dll", EntryPoint:="AD101_ChangeWindowHandle")> _
  Public Shared Function AD101_ChangeWindowHandle(ByVal hWnd As Integer) As Integer
  End Function

  ' Show or don't show collateral phone dialed number
  <DllImport("AD101Device.dll", EntryPoint:="AD101_ShowCollateralPhoneDialed")> _
  Public Shared Sub AD101_ShowCollateralPhoneDialed(ByVal bShow As Boolean)
  End Sub

  ' Control led 
  <DllImport("AD101Device.dll", EntryPoint:="AD101_SetLED")> _
  Public Shared Function AD101_SetLED(ByVal nLine As Integer, ByVal enumLed As Integer) As Integer
  End Function


  ' Control line connected with ad101 device to busy or idel
  <DllImport("AD101Device.dll", EntryPoint:="AD101_SetBusy")> _
  Public Shared Function AD101_SetBusy(ByVal nLine As Integer, ByVal enumLineBusy As Integer) As Integer
  End Function

  ' Set line to start talking than start timer
  <DllImport("AD101Device.dll", EntryPoint:="AD101_SetLineStartTalk")> _
  Public Shared Function AD101_SetLineStartTalk(ByVal nLine As Integer) As Integer
  End Function


  ' Set time to start talking after dialed number 
  <DllImport("AD101Device.dll", EntryPoint:="AD101_SetDialOutStartTalkingTime")> _
  Public Shared Function AD101_SetDialOutStartTalkingTime(ByVal nSecond As Integer) As Integer
  End Function


  ' Set ring end time
  <DllImport("AD101Device.dll", EntryPoint:="AD101_SetRingOffTime")> _
  Public Shared Function AD101_SetRingOffTime(ByVal nSecond As Integer) As Integer
  End Function



  Public Const MCU_BACKID As Integer = &H7
  ' Return Device ID
  Public Const MCU_BACKSTATE As Integer = &H8
  ' Return Device State
  Public Const MCU_BACKCID As Integer = &H9
  ' Return Device CallerID
  Public Const MCU_BACKDIGIT As Integer = &HA
  ' Return Device Dial Digit
  Public Const MCU_BACKDEVICE As Integer = &HB
  ' Return Device Back Device ID
  Public Const MCU_BACKPARAM As Integer = &HC
  ' Return Device Paramter
  Public Const MCU_BACKCPUID As Integer = &HD
  ' Return Device CPU ID
  Public Const MCU_BACKCOLLATERAL As Integer = &HE
  ' Return Collateral phone dialed
  Public Const MCU_BACKDISABLE As Integer = &HFF
  ' Return Device Init
  Public Const MCU_BACKENABLE As Integer = &HEE
  Public Const MCU_BACKMISSED As Integer = &HAA
  ' Missed call 
  Public Const MCU_BACKTALK As Integer = &HBB
  ' Start Talk
  ' LED Status 
  Private Enum LEDTYPE
    LED_CLOSE = 1
    LED_RED
    LED_GREEN
    LED_YELLOW
    LED_REDSLOW
    LED_GREENSLOW
    LED_YELLOWSLOW
    LED_REDQUICK
    LED_GREENQUICK
    LED_YELLOWQUICK
  End Enum

  ' Line Status 
  Private Enum ENUMLINEBUSY
    LINEBUSY = 0
    LINEFREE
  End Enum


  Public Const HKONSTATEPRA As Integer = &H1
  ' hook on pr+  HOOKON_PRA
  Public Const HKONSTATEPRB As Integer = &H2
  ' hook on pr-  HOOKON_PRR
  Public Const HKONSTATENOPR As Integer = &H3
  ' have pr  HAVE_PR
  Public Const HKOFFSTATEPRA As Integer = &H4
  ' hook off pr+  HOOKOFF_PRA
  Public Const HKOFFSTATEPRB As Integer = &H5
  ' hook off pr-  HOOKOFF_PRR
  Public Const NO_LINE As Integer = &H6
  ' no line  NULL_LINE
  Public Const RINGONSTATE As Integer = &H7
  ' ring on  RING_ON
  Public Const RINGOFFSTATE As Integer = &H8
  ' ring off RING_OFF
  Public Const NOHKPRA As Integer = &H9
  ' NOHOOKPRA= 0x09, // no hook pr+
  Public Const NOHKPRB As Integer = &HA
  ' NOHOOKPRR= 0x0a, // no hook pr-
  Public Const NOHKNOPR As Integer = &HB
  ' NOHOOKNPR= 0x0b, // no hook no pr
  Public Const WM_USBLINEMSG As Integer = 1024 + 180



  Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    listView1.Columns.Clear()
    listView1.Columns.Add("Channel No.", 90, HorizontalAlignment.Left)
    listView1.Columns.Add("Device State", 80, HorizontalAlignment.Left)
    listView1.Columns.Add("CPU Version", 120, HorizontalAlignment.Left)
    listView1.Columns.Add("Line State", 120, HorizontalAlignment.Left)
    listView1.Columns.Add("Caller ID", 100, HorizontalAlignment.Left)
    listView1.Columns.Add("Dialed Number", 100, HorizontalAlignment.Left)
    listView1.Columns.Add("Collateral Dialed", 110, HorizontalAlignment.Left)
    listView1.Columns.Add("Talk Time", 100, HorizontalAlignment.Left)
    listView1.Columns.Add("CPU ID", 100, HorizontalAlignment.Left)


    listView1.Items.Add("Channel 1")
    listView1.Items.Add("Channel 2")
    listView1.Items.Add("Channel 3")
    listView1.Items.Add("Channel 4")

    ComboBox1.SelectedIndex = 0
    For i As Integer = 0 To 3
      listView1.Items(i).SubItems.Add("")
      listView1.Items(i).SubItems.Add("")
      listView1.Items(i).SubItems.Add("")
      listView1.Items(i).SubItems.Add("")
      listView1.Items(i).SubItems.Add("")
      listView1.Items(i).SubItems.Add("")
      listView1.Items(i).SubItems.Add("")
      listView1.Items(i).SubItems.Add("")
      listView1.Items(i).SubItems(1).Text = "Disable"
    Next

    If AD101_InitDevice(Handle.ToInt32()) = 0 Then
      Return
    End If

    AD101_GetDevice()

    textBox7.Text = "3"
    textBox8.Text = "7"


    AD101_SetDialOutStartTalkingTime(3)
    AD101_SetRingOffTime(7)

    checkBoxSHOWPHONENUMBER.Checked = True
    checkBoxSHOWNUMBER.Checked = True


  End Sub

  Private Sub OnDeviceMsg(ByVal wParam As IntPtr, ByVal Lparam As IntPtr)
    Dim nMsg As New Integer()
    Dim nLine As New Integer()

    nMsg = wParam.ToInt32() Mod 65536
    nLine = wParam.ToInt32() \ 65536

    Select Case nMsg
      Case MCU_BACKDISABLE
        listView1.Items(nLine).SubItems(1).Text = "Disable"
        listView1.Items(nLine).SubItems(2).Text = ""
        listView1.Items(nLine).SubItems(3).Text = ""
        listView1.Items(nLine).SubItems(4).Text = ""
        listView1.Items(nLine).SubItems(5).Text = ""
        listView1.Items(nLine).SubItems(6).Text = ""
        listView1.Items(nLine).SubItems(7).Text = ""
        listView1.Items(nLine).SubItems(8).Text = ""

        Exit Select

      Case MCU_BACKENABLE
        listView1.Items(nLine).SubItems(1).Text = "Enable"
        Exit Select

      Case MCU_BACKID
        If True Then
          Dim szCPUVersion As New StringBuilder(32)

          listView1.Items(nLine).SubItems(1).Text = "Enable"

          AD101_GetCPUVersion(nLine, szCPUVersion)

          listView1.Items(nLine).SubItems(2).Text = szCPUVersion.ToString()

        End If
        Exit Select

      Case MCU_BACKCID
        If True Then
          Dim szCallerID As New StringBuilder(128)
          Dim szName As New StringBuilder(128)
          Dim szTime As New StringBuilder(128)

          Dim nLen As Integer = AD101_GetCallerID(nLine, szCallerID, szName, szTime)
          listView1.Items(nLine).SubItems(4).Text = szCallerID.ToString()
        End If
        Exit Select

      Case MCU_BACKSTATE
        If True Then
          Select Case Lparam.ToInt32()
            Case HKONSTATEPRA
              listView1.Items(nLine).SubItems(3).Text = "HOOK ON PR+"

              Exit Select

            Case HKONSTATEPRB
              listView1.Items(nLine).SubItems(3).Text = "HOOK ON PR-"
              Exit Select

            Case HKONSTATENOPR
              listView1.Items(nLine).SubItems(3).Text = "HOOK ON NOPR"
              Exit Select

            Case HKOFFSTATEPRA
              If True Then
                listView1.Items(nLine).SubItems(3).Text = "HOOK OFF PR+"

                Dim szCallerID As New StringBuilder(128)
                Dim szName As New StringBuilder(128)
                Dim szTime As New StringBuilder(128)

                If AD101_GetCallerID(nLine, szCallerID, szName, szTime) < 1 OrElse AD101_GetRingIndex(nLine) < 1 Then
                  listView1.Items(nLine).SubItems(4).Text = ""
                End If

                listView1.Items(nLine).SubItems(5).Text = ""
                listView1.Items(nLine).SubItems(6).Text = ""
                listView1.Items(nLine).SubItems(7).Text = ""

              End If
              Exit Select

            Case HKOFFSTATEPRB
              If True Then
                listView1.Items(nLine).SubItems(3).Text = "HOOK OFF PR-"

                Dim szCallerID As New StringBuilder(128)
                Dim szName As New StringBuilder(128)
                Dim szTime As New StringBuilder(128)

                If AD101_GetCallerID(nLine, szCallerID, szName, szTime) < 1 OrElse AD101_GetRingIndex(nLine) < 1 Then
                  listView1.Items(nLine).SubItems(4).Text = ""
                End If

                listView1.Items(nLine).SubItems(5).Text = ""
                listView1.Items(nLine).SubItems(6).Text = ""
                listView1.Items(nLine).SubItems(7).Text = ""
              End If
              Exit Select

            Case NO_LINE
              If True Then
                listView1.Items(nLine).SubItems(3).Text = "No Line"

                Dim szCallerID As New StringBuilder(128)
                Dim szName As New StringBuilder(128)
                Dim szTime As New StringBuilder(128)

                If AD101_GetCallerID(nLine, szCallerID, szName, szTime) < 1 OrElse AD101_GetRingIndex(nLine) < 1 Then
                  listView1.Items(nLine).SubItems(4).Text = ""
                End If

                listView1.Items(nLine).SubItems(5).Text = ""
                listView1.Items(nLine).SubItems(6).Text = ""
                listView1.Items(nLine).SubItems(7).Text = ""
              End If
              Exit Select

            Case RINGONSTATE
              If True Then
                Dim szRing As String = "Ring:" & String.Format("{0:D2}", AD101_GetRingIndex(nLine))

                listView1.Items(nLine).SubItems(3).Text = szRing
              End If
              Exit Select

            Case RINGOFFSTATE
              listView1.Items(nLine).SubItems(3).Text = "Ring Off"
              Exit Select

            Case NOHKPRA
              listView1.Items(nLine).SubItems(3).Text = "NO HOOK PR+"

              Exit Select

            Case NOHKPRB
              listView1.Items(nLine).SubItems(3).Text = "NO HOOK PR-"

              Exit Select

            Case NOHKNOPR
              If True Then
                listView1.Items(nLine).SubItems(3).Text = "NO HOOK NOPR"

                Dim szCallerID As New StringBuilder(128)
                Dim szName As New StringBuilder(128)
                Dim szTime As New StringBuilder(128)

                If AD101_GetCallerID(nLine, szCallerID, szName, szTime) < 1 OrElse AD101_GetRingIndex(nLine) < 1 Then
                  listView1.Items(nLine).SubItems(4).Text = ""
                End If

                listView1.Items(nLine).SubItems(5).Text = ""
                listView1.Items(nLine).SubItems(6).Text = ""
                listView1.Items(nLine).SubItems(7).Text = ""
              End If
              Exit Select
            Case Else


              Exit Select
          End Select
        End If
        Exit Select

      Case MCU_BACKDIGIT
        If True Then
          If checkBoxSHOWPHONENUMBER.Checked = True Then
            Dim szDialDigit As New StringBuilder(128)

            AD101_GetDialDigit(nLine, szDialDigit)


            listView1.Items(nLine).SubItems(5).Text = szDialDigit.ToString()
          End If

        End If
        Exit Select


      Case MCU_BACKCOLLATERAL
        If True Then
          Dim szDialDigit As New StringBuilder(128)

          AD101_GetCollateralDialDigit(nLine, szDialDigit)

          listView1.Items(nLine).SubItems(6).Text = szDialDigit.ToString()
        End If
        Exit Select

      Case MCU_BACKDEVICE
        If True Then
          Dim szCPUVersion As New StringBuilder(32)

          listView1.Items(nLine).SubItems(1).Text = "Enable"

          AD101_GetCPUVersion(nLine, szCPUVersion)

          listView1.Items(nLine).SubItems(2).Text = szCPUVersion.ToString()

        End If
        Exit Select

      Case MCU_BACKPARAM
        If True Then
          Dim tagParameter As New AD101DEVICEPARAMETER()

          AD101_GetParameter(nLine, tagParameter)
          textBox1.Text = tagParameter.nRingOn.ToString()
          textBox2.Text = tagParameter.nRingOff.ToString()
          textBox3.Text = tagParameter.nHookOn.ToString()
          textBox4.Text = tagParameter.nHookOff.ToString()
          textBox5.Text = tagParameter.nStopCID.ToString()
          textBox6.Text = tagParameter.nNoLine.ToString()
        End If
        Exit Select

      Case MCU_BACKCPUID
        If True Then
          Dim szCPUID As New StringBuilder(4)

          AD101_GetCPUID(nLine, szCPUID)

          listView1.Items(nLine).SubItems(8).Text = szCPUID.ToString()

        End If
        Exit Select

      Case MCU_BACKMISSED
        If True Then
          listView1.Items(nLine).SubItems(3).Text = "Missed Call"
        End If
        Exit Select

      Case MCU_BACKTALK
        If True Then
          Dim strTalk As String
          strTalk = String.Format("{0:D2}", Lparam) & "S"

          listView1.Items(nLine).SubItems(7).Text = strTalk

        End If
        Exit Select
      Case Else

        Exit Select
    End Select
  End Sub

  Protected Overrides Sub DefWndProc(ByRef m As System.Windows.Forms.Message)
    Select Case m.Msg
      Case WM_USBLINEMSG
        '處理消息　
        OnDeviceMsg(m.WParam, m.LParam)
        Exit Select
      Case Else
        MyBase.DefWndProc(m)
        '調用基類函數處理非自定義消息。　　　
        Exit Select
    End Select
  End Sub

  Private Sub radioButton1_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles RadioButton1.CheckedChanged
    AD101_SetLED(ComboBox1.SelectedIndex, 1)
  End Sub

  Private Sub radioButton2_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles RadioButton2.CheckedChanged
    AD101_SetLED(ComboBox1.SelectedIndex, 2)
  End Sub

  Private Sub radioButton3_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles RadioButton3.CheckedChanged
    AD101_SetLED(ComboBox1.SelectedIndex, 5)
  End Sub

  Private Sub radioButton4_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles RadioButton4.CheckedChanged
    AD101_SetLED(ComboBox1.SelectedIndex, 8)
  End Sub

  Private Sub radioButton5_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles RadioButton5.CheckedChanged
    AD101_SetLED(ComboBox1.SelectedIndex, 3)
  End Sub

  Private Sub radioButton6_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles RadioButton6.CheckedChanged
    AD101_SetLED(ComboBox1.SelectedIndex, 6)
  End Sub

  Private Sub radioButton7_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles RadioButton7.CheckedChanged
    AD101_SetLED(ComboBox1.SelectedIndex, 9)
  End Sub

  Private Sub radioButton8_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles RadioButton8.CheckedChanged
    AD101_SetLED(ComboBox1.SelectedIndex, 4)
  End Sub

  Private Sub radioButton9_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles RadioButton9.CheckedChanged
    AD101_SetLED(ComboBox1.SelectedIndex, 7)
  End Sub

  Private Sub radioButton10_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles RadioButton10.CheckedChanged
    AD101_SetLED(ComboBox1.SelectedIndex, 10)
  End Sub

  Private Sub BtnLineBusy_Click(ByVal sender As Object, ByVal e As EventArgs) Handles BtnLineBusy.Click
    'Private Sub BtnLineBusy_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnLineBusy.Click
    AD101_SetBusy(ComboBox1.SelectedIndex, 0)
  End Sub

  Private Sub BtnLineFree_Click(ByVal sender As Object, ByVal e As EventArgs) Handles BtnLineFree.Click
    AD101_SetBusy(ComboBox1.SelectedIndex, 1)
  End Sub

  Private Sub BtnSet_Click(ByVal sender As Object, ByVal e As EventArgs) Handles BtnSet.Click
    AD101_SetDialOutStartTalkingTime(Convert.ToInt32(TextBox7.Text))
    AD101_SetRingOffTime(Convert.ToInt32(TextBox8.Text))
  End Sub

  Private Sub checkBoxSHOWNUMBER_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles checkBoxSHOWNUMBER.CheckedChanged
    AD101_ShowCollateralPhoneDialed(checkBoxSHOWPHONENUMBER.Checked)
  End Sub

  Private Sub Read_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Read.Click
    AD101_ReadParameter(ComboBox1.SelectedIndex)

    TextBox1.Text = ""
    TextBox2.Text = ""
    TextBox3.Text = ""
    TextBox4.Text = ""
    TextBox5.Text = ""
    TextBox6.Text = ""
  End Sub

  Private Sub Write_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Write.Click
    Dim tagParameter As New AD101DEVICEPARAMETER()

    tagParameter.nRingOn = Convert.ToInt32(TextBox1.Text)
    tagParameter.nRingOff = Convert.ToInt32(TextBox2.Text)
    tagParameter.nHookOn = Convert.ToInt32(TextBox3.Text)
    tagParameter.nHookOff = Convert.ToInt32(TextBox4.Text)
    tagParameter.nStopCID = Convert.ToInt32(TextBox5.Text)
    tagParameter.nNoLine = Convert.ToInt32(TextBox6.Text)

    If AD101_SetParameter(ComboBox1.SelectedIndex, tagParameter) = 0 Then
      MessageBox.Show("Set parameters failed!")
    End If
  End Sub


  Private Sub Form1_FormClosed(ByVal sender As Object, ByVal e As FormClosedEventArgs)
    AD101_FreeDevice()
  End Sub

  Public Sub New()

    ' 此為 Windows Form 設計工具所需的呼叫。
    InitializeComponent()

    ' 在 InitializeComponent() 呼叫之後加入任何初始設定。

  End Sub
End Class
