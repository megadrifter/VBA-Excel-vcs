Attribute VB_Name = "sents_VerControl"


Sub sents_Version_control_v0()
Attribute sents_Version_control_v0.VB_Description = "��������, ����, ���������, ���������, �������� �����, ����������"
Attribute sents_Version_control_v0.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Version_control_v0 ������
' ��������, ����, ���������, ���������, �������� �����, ����������
'
Debug.Print "��������: " & ActiveWorkbook.title
Debug.Print "����: " & ActiveWorkbook.Subject
'Debug.Print "���������: " & ActiveWorkbook.Title
Debug.Print "���������: " & ActiveWorkbook.BuiltinDocumentProperties("Content status")
Debug.Print "�������� �����: " & ActiveWorkbook.Keywords
Debug.Print "����������: " & ActiveWorkbook.Comments
Debug.Print "��� �����: " & ActiveWorkbook.Name
Debug.Print "���� � �����: " & ActiveWorkbook.Path;
' ����� Author
'Debug.Print "���������: " & ActiveWorkbook.Title
'Debug.Print "���������: " & ActiveWorkbook.Title


End Sub
Sub sents_Save_new_version()
Attribute sents_Save_new_version.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������2 ������
'
' get original path and name
sents_subfolder_name = "S_versions"

file_path = ActiveWorkbook.Path
file_name = ActiveWorkbook.Name
file_title = ActiveWorkbook.title

'check existing subfolder sents_subfolder_name

'check existing title property

'create new file name: format(date, "yyyy-mm-dd_") & title & "_v" & version
 
 
 
 '   ChDir "C:\_data\Projects\�������� ���������"
  '  ActiveWorkbook.SaveAs Filename:= _
  '      "C:\_data\Projects\�������� ���������\2016-06-06-Sales-activities1.xlsx", _
  '      FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        

    'ActiveWorkbook.Save
    Windows("2016-06-03_Intraservice_export.xls").Activate
    Windows("2016-06-06-Sales-activities1.xlsx").Activate
    ActiveWorkbook.Save
End Sub

Function sents_GetVersion(Optional show)
Debug.Print
'ActiveWorkbook.CustomDocumentProperties
ver = "not set"
'For Each cdp In ActiveWorkbook.CustomDocumentProperties
'  If cdp.Name = "S_version" Then ver = cdp.Value
'Next cdp
On Error Resume Next
ver = ActiveWorkbook.CustomDocumentProperties("S_version").Value

Debug.Print ver ' print acquired veersion number
sents_GetVersion = ver
If show Then MsgBox ver
End Function

Sub sents_SetVersion(Optional ver)


If (IsMissing(ver)) Then
On Error GoTo Quit
        ver = CInt(InputBox("������� ����� ������ (���� ����� �����)", "������?", ActiveWorkbook.CustomDocumentProperties("S_version").Value))
End If


On Error GoTo NotFound
ActiveWorkbook.CustomDocumentProperties("S_version").Value = ver
    found = True
Exit Sub


NotFound: ' create custom property
  ActiveWorkbook.CustomDocumentProperties.Add Name:="S_version", LinkToContent:=False, Value:=ver
  
Quit:
End Sub

Sub sents_SetTitle(Optional s_title)

found = False

' look in the document
On Error GoTo NotFound
old_title = ActiveWorkbook.CustomDocumentProperties("S_title").Value
found = True
Exit Sub

NotFound: ' create custom property
If Not found Then
ActiveWorkbook.CustomDocumentProperties.Add Name:="S_title", LinkToContent:=False, Value:=s_title
End If

'if no argument is given
If (IsMissing(s_title)) Then
On Error GoTo Quit
        s_title = InputBox("������� ��� ���������. ������ ���: " & old_title & ", ��� �����: " & ActiveWorkbook.Name, "���?", ActiveWorkbook.Name)
        
End If

ActiveWorkbook.CustomDocumentProperties("S_title").Value = s_title
  
Quit:
End Sub


