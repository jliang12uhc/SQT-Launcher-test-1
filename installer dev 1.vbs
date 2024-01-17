Option Explicit


Class ProgressBar 
    Private m_PercentComplete 
    Private m_CurrentStep 
    Private m_ProgressBar 
    Private m_Title 
    Private m_Text 
    Private m_StatusBarText 
    Private n
  
    'Initialize defaults 
    Private Sub ProgessBar_Initialize 
        m_PercentComplete = 0 
        m_CurrentStep = 0 
        m_Title = "Progress" 
        m_Text = "" 
    End Sub 
  
    Public Function SetTitle(pTitle) 
        m_Title = pTitle 
    End Function 
  
    Public Function SetText(pText) 
        m_Text = pText 
    End Function 
  
    Public Function Update(percentComplete) 
        m_PercentComplete = percentComplete 
        UpdateProgressBar() 
    End Function 
  
    Public Function Show() 
        Set m_ProgressBar = CreateObject("InternetExplorer.Application") 
        'in code, the colon acts as a line feed 
        m_ProgressBar.navigate2 "about:blank" : m_ProgressBar.width = 560 : m_ProgressBar.height = 60 : m_ProgressBar.toolbar = false : m_ProgressBar.menubar = false : m_ProgressBar.statusbar = false : m_ProgressBar.visible = True 
        m_ProgressBar.document.write "<body Scroll=no style='margin:0px;padding:0px;'><div style='text-align:center;'><span name='pc' id='pc'>0</span></div>" 
        m_ProgressBar.document.write "<div id='statusbar' name='statusbar' style='border:1px solid blue;line-height:10px;height:10px;color:blue;'></div>" 
        m_ProgressBar.document.write "<div style='text-align:center'><span id='text' name='text'></span></div>" 
    End Function 
  
    Public Function Close() 
        m_ProgressBar.quit 
        set m_ProgressBar = Nothing 
    End Function 
  
    Private Function UpdateProgressBar() 
        If m_PercentComplete = 0 Then 
            m_StatusBarText = "" 
        End If 
        For n = m_CurrentStep to m_PercentComplete - 1 
            m_StatusBarText = m_StatusBarText & "|" 
            m_ProgressBar.Document.GetElementById("statusbar").InnerHtml = m_StatusBarText 
            m_ProgressBar.Document.title = n & "% Complete : " & m_Title 
            m_ProgressBar.Document.GetElementById("pc").InnerHtml = n & "% Complete : " & m_Title 
            wscript.sleep 10 
        Next 
        m_ProgressBar.Document.GetElementById("statusbar").InnerHtml = m_StatusBarText 
        m_ProgressBar.Document.title = m_PercentComplete & "% Complete : " & m_Title 
        m_ProgressBar.Document.GetElementById("pc").InnerHtml = m_PercentComplete & "% Complete : " & m_Title 
        m_ProgressBar.Document.GetElementById("text").InnerHtml = m_Text 
        m_CurrentStep = m_PercentComplete 
    End Function 
  
 End Class 

Dim OBJSysInfo, OBJUser, STRUser
dim boolContinue
Dim pb 
Dim percentComplete 
Dim sqtVersionNumber

'Setup the initial progress bar
Set pb = New ProgressBar 

boolContinue=True

percentComplete = 0
'pb.SetTitle("Step 1 of 10")
'pb.SetText("Checking for updates...")
'pb.Show()

wscript.echo "Hello World!!"
Dim fso
Dim stdout
Dim stderr 


If wscript.Arguments.Count > 0 Then
    sqtVersionNumber = Wscript.arguments(0)    
else
'    pb.SetText("An error occurred during update. Please contact the SQT Administrators for assistance. ERROR: No arguments were passed to Installer")
'    pb.update(20)
end if

'pb.SetText("Sleeping...")
'pb.Update(20)
WScript.Sleep(5000)

'pb.Close()
wscript.quit