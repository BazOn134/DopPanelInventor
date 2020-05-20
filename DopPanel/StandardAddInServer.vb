Imports Inventor
Imports System.Runtime.InteropServices
'Imports Microsoft.Win32
Imports System.Windows.Forms
'����� ��������

Namespace InventorAddIn1
    <ProgIdAttribute("InventorAddIn1.StandardAddInServer"), _
    GuidAttribute("b982d25d-0e76-4273-a67d-5c7f0fa60037")> _
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer
        '=================================================================================================================
        ' Inventor application object.
        'Private m_inventorApplication As Inventor.Application
        '��������� ������ ����� ��������� (��� �������) ���� ��� ����
        'Public m_inventorApplication As Inventor.Application
        '������� ���������� ���������� ����
        Private WithEvents m_Button As ButtonDefinition
        Private WithEvents m_InvAppEvent As ApplicationEvents
        '==========================================================================

#Region "ApplicationAddInServer Members"

        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate
            ' Initialize AddIn members.
            m_inventorApplication = addInSiteObject.Application
            m_InvAppEvent = m_inventorApplication.ApplicationEvents
            ' TODO:  Add ApplicationAddInServer.Activate implementation.
            ' e.g. event initialization, command creation etc.
            ' ������� ����
            m_Window = m_inventorApplication.UserInterfaceManager.DockableWindows.Add("{b982d25d-0e76-4273-a67d-5c7f0fa60037}", _
                                                                                      "DopPanel", "DopPanel")
            '������� �����
            m_Form = New UserControl1
            Dim sv_Form As Form
            sv_Form = New Form
            '������ ����� �� ����
            m_Window.AddChild(m_Form.Handle)
            '��������� ������������ ���� ������ � ����� ������ (���� ����������)
            'm_Window.DisabledDockingStates = DockingStateEnum.kDockTop + DockingStateEnum.kDockBottom
            '���������� ���� �����
            m_Window.DockingState = DockingStateEnum.kDockLeft
            ' ������� ������� ����������
            m_Button = m_inventorApplication.CommandManager.ControlDefinitions.AddButtonDefinition("DopPanel", _
                 "InventorAddInmyPanelBtn", CommandTypesEnum.kNonShapeEditCmdType, "{b982d25d-0e76-4273-a67d-5c7f0fa60037}")
            ' ��������� ������� ������ �������
            m_Button.OverrideShortcut = "Ctrl+2"
        End Sub

        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate
            ' TODO:  Add ApplicationAddInServer.Deactivate implementation
            ' Release objects.
            ' AddIn �����������, ��������� ����� � ������
            m_Window.Delete()
            m_Window = Nothing
            m_Button = Nothing
            m_inventorApplication = Nothing
            System.GC.WaitForPendingFinalizers()
            System.GC.Collect()
        End Sub

        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation

            ' This property is provided to allow the AddIn to expose an API 
            ' of its own to other programs. Typically, this  would be done by
            ' implementing the AddIn's API interface in a class and returning 
            ' that class object through this property.

            Get
                Return Nothing
            End Get

        End Property

        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand

            ' Note:this method is now obsolete, you should use the 
            ' ControlDefinition functionality for implementing commands.

        End Sub

        Private Sub m_Button_OnExecute(ByVal Context As Inventor.NameValueMap) Handles m_Button.OnExecute
            m_Window.Visible = Not m_Window.Visible
        End Sub
#End Region

        '=================================================================================================================

        Private Sub m_InvAppEvent_OnActivateDocument(ByVal DocumentObject As Inventor._Document, ByVal BeforeOrAfter As Inventor.EventTimingEnum, ByVal Context As Inventor.NameValueMap, ByRef HandlingCode As Inventor.HandlingCodeEnum) Handles m_InvAppEvent.OnActivateDocument
            If BeforeOrAfter = EventTimingEnum.kAfter Then m_Form.Obrabotka(m_inventorApplication.ActiveEditDocument.FullFileName, False)
            '    MessageBox.Show(m_inventorApplication.ActiveEditDocument.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value & "   ���������")
            '    MessageBox.Show(m_inventorApplication.ActiveEditDocument.DisplayName & "   DisplayName ���������")
            '    MessageBox.Show(m_inventorApplication.ActiveEditDocument.FullFileName & "   FullFileName ���������")
            'End If
        End Sub

    End Class
End Namespace

