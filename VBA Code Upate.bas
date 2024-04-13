Attribute VB_Name = "ģ��5"
Sub ImportAndCompareData()
Const procName As String = "ImportAndCompareData"
On Error GoTo ErrorHandler

        ' �жϲ���ϵͳ
        Dim isWindows As Boolean
        Dim fd As fileDialog
        Dim sourceWb As Workbook
        Dim selectedFile As String
 
    
        ' Detect operating system
        operatingSystem = DetectOS()
        ' Notify user of the operating system
        MsgBox "Operation system is " & operatingSystem, vbInformation
    
        ' Open file dialog based on the operating system
        If operatingSystem = "Windows" Then
            selectedFile = OpenFileDialogWindows() ' Function to open file dialog on Windows
        Else
            selectedFile = OpenFileDialogMac() ' Function to open file dialog on Mac
        End If

        ' Check if a file was selected
        If selectedFile = "" Then Exit Sub
    
        ' Open source workbook
        Set sourceWb = Workbooks.Open(selectedFile)
        ' Ask the user to confirm if the correct file has been opened
        response = MsgBox("Have you opened the correct file?", vbExclamation + vbYesNo, "Confirm File")
    
        ' Check the user's response
        If response = vbNo Then
            Exit Sub
    
        End If
        
         ' ��ѡ�е��ļ�
        Set sourceWb = Workbooks.Open(selectedFile, ReadOnly:=True)
    
        ' ���ơ�Project Database�������������
        Call CopyProjectDatabase(sourceWb)
    
        ' ���ơ�BOM�������������
        Call CopyBOM(sourceWb)
    
        sourceWb.Close False ' �ر�Դ�ļ�

Exit Sub
ErrorHandler:
    ' Example of manually inserted line number and content
    HandleError "Error #" & Err.Number & ": " & Err.Description, procName
    'Resume Next ' Or other error recovery logic
End Sub
Sub CopyProjectDatabase(sourceWb As Workbook)
Const procName As String = "sourceWb"
On Error GoTo ErrorHandler

    Dim targetWs As Worksheet, sourceWs As Worksheet
    Dim targetTbl As ListObject, sourceTbl As ListObject
    Dim targetLastRow As Long, sourceLastRow As Long, i As Long
    Dim projectCol As Integer, updateTimeCol As Integer, salesCol As Integer
    Dim rowsCopied As Long
    rowsCopied = 0 ' ��ʼ��������

    Set targetWs = ThisWorkbook.Sheets("Project Database")
    Set sourceWs = sourceWb.Sheets("Project Database")
    
    Set targetTbl = targetWs.ListObjects(1)
    Set sourceTbl = sourceWs.ListObjects(1)

    ' ���������Ŀ���ơ�����ʱ������۸����˷ֱ����� A��B �� C
    projectCol = FindColumnInTable(sourceTbl, "��Ŀ����") ' ��Ŀ������
    If projectCol = 0 Then
        MsgBox "��Ŀ������δ�ҵ���"
        Exit Sub
    End If
        
    updateTimeCol = FindColumnInTable(sourceTbl, "����ʱ��") ' ����ʱ����
    If updateTimeCol = 0 Then
        MsgBox "��Ŀ������δ�ҵ���"
        Exit Sub
    End If
    
    salesCol = FindColumnInTable(sourceTbl, "���۸�����") ' ���۸�������
    If salesCol = 0 Then
        MsgBox "��Ŀ������δ�ҵ���"
        Exit Sub
    End If
    
    ' ��ȡĿ���Դ�������һ������
    targetLastRow = targetTbl.ListColumns(projectCol).DataBodyRange.Rows.count
    sourceLastRow = sourceTbl.ListColumns(projectCol).DataBodyRange.Rows.count
    
    ' �ԱȲ���������
    Dim sourceRow As Range
    For i = 1 To sourceTbl.ListRows.count
        Set sourceRow = sourceTbl.ListRows(i).Range
        If Not IsInTarget(targetTbl, sourceRow.Cells(1, projectCol).Value, _
                          sourceRow.Cells(1, updateTimeCol).Value, _
                          sourceRow.Cells(1, salesCol).Value, projectCol, updateTimeCol, salesCol) Then
              
                ' ��Ŀ��ListObject��ĩβ�������
                Dim newRow As listRow
                Set newRow = targetTbl.ListRows.Add
    
                ' �����ݴ�Դ�и��Ƶ�Ŀ����
                sourceRow.Copy
                newRow.Range.PasteSpecial xlPasteValues ' ճ��ֵ������
                Application.CutCopyMode = False ' ȡ��������ѡ��
                
                 rowsCopied = rowsCopied + 1 ' ���Ӽ�����
        End If
    Next i
    
    ' ��ʾ�������������
    MsgBox " �ɹ�������Ŀ��������" & rowsCopied, vbInformation, "�������"

Exit Sub
ErrorHandler:
    ' Example of manually inserted line number and content
    HandleError "Error #" & Err.Number & ": " & Err.Description, procName
    'Resume Next ' Or other error recovery logic
End Sub
Sub CopyBOM(sourceWb As Workbook)
Const procName As String = "CopyBOM"
On Error GoTo ErrorHandler

    Dim targetWs As Worksheet, sourceWs As Worksheet
    Dim targetTbl As ListObject, sourceTbl As ListObject
    Dim targetLastRow As Long, sourceLastRow As Long, i As Long
    Dim projectCol As Integer, updateTimeCol As Integer, salesCol As Integer, MNCol As Integer
    
    Dim rowsCopied As Long
    rowsCopied = 0 ' ��ʼ��������

    Set targetWs = ThisWorkbook.Sheets("BOM")
    Set sourceWs = sourceWb.Sheets("BOM")
    
    Set targetTbl = targetWs.ListObjects(1)
    Set sourceTbl = sourceWs.ListObjects(1)

    ' ���������Ŀ���ơ�����ʱ������۸����˷ֱ����� A��B �� C
    projectCol = FindColumnInTable(sourceTbl, "Project Name") ' ��Ŀ������
    If projectCol = 0 Then
        MsgBox "��Ŀ������δ�ҵ���"
        Exit Sub
    End If
        
    updateTimeCol = FindColumnInTable(sourceTbl, "update time") ' ����ʱ����
    If updateTimeCol = 0 Then
        MsgBox "����ʱ����δ�ҵ���"
        Exit Sub
    End If
    
    MNCol = FindColumnInTable(sourceTbl, "MN") ' MN��
    If MNCol = 0 Then
        MsgBox "MN��δ�ҵ���"
        Exit Sub
    End If
    
    salesCol = FindColumnInTable(sourceTbl, "����") ' ���۸�������
    If salesCol = 0 Then
        MsgBox "���۸�������δ�ҵ���"
        Exit Sub
    End If
    
    ' ��ȡĿ���Դ�������һ������
    targetLastRow = targetTbl.ListColumns(projectCol).DataBodyRange.Rows.count
    sourceLastRow = sourceTbl.ListColumns(projectCol).DataBodyRange.Rows.count
    
    ' �ԱȲ���������
    Dim sourceRow As Range
    For i = 1 To sourceTbl.ListRows.count
        Set sourceRow = sourceTbl.ListRows(i).Range
        If Not IsInTargetBOM(targetTbl, sourceRow.Cells(1, projectCol).Value, _
                          sourceRow.Cells(1, updateTimeCol).Value, _
                          sourceRow.Cells(1, MNCol).Value, _
                          sourceRow.Cells(1, salesCol).Value, projectCol, updateTimeCol, MNCol, salesCol) Then
              
                ' ��Ŀ��ListObject��ĩβ�������
                Dim newRow As listRow
                Set newRow = targetTbl.ListRows.Add
    
                ' �����ݴ�Դ�и��Ƶ�Ŀ����
                sourceRow.Copy
                newRow.Range.PasteSpecial xlPasteValues ' ճ��ֵ������
                Application.CutCopyMode = False ' ȡ��������ѡ��
                
                 rowsCopied = rowsCopied + 1 ' ���Ӽ�����
        End If
    Next i
    
    ' ��ʾ�������������
    MsgBox " �ɹ�����BOM��������" & rowsCopied, vbInformation, "�������"

Exit Sub
ErrorHandler:
    ' Example of manually inserted line number and content
    HandleError "Error #" & Err.Number & ": " & Err.Description, procName
    'Resume Next ' Or other error recovery logic
End Sub

Function IsInTarget(targetTbl As ListObject, ProjectName As String, updateTime As String, salesPerson As String, projectCol As Integer, updateTimeCol As Integer, salesCol As Integer) As Boolean
Const procName As String = "IsInTarget"
On Error GoTo ErrorHandler
    
    
    Dim row As listRow
    

    For Each row In targetTbl.ListRows
 
        
        If row.Range.Cells(1, projectCol).Value = ProjectName And _
           row.Range.Cells(1, updateTimeCol).Value = updateTime And _
           row.Range.Cells(1, salesCol).Value = salesPerson Then
            IsInTarget = True
            Exit Function
        End If
    Next row
    IsInTarget = False
    
    
Exit Function
ErrorHandler:
    ' Example of manually inserted line number and content
    HandleError "Error #" & Err.Number & ": " & Err.Description, procName
    'Resume Next ' Or other error recovery logic
End Function
Function IsInTargetBOM(targetTbl As ListObject, ProjectName As String, updateTime As String, MN As String, salesPerson As String, projectCol As Integer, updateTimeCol As Integer, MNCol As Integer, salesCol As Integer) As Boolean
Const procName As String = "IsInTargetBOM"
On Error GoTo ErrorHandler
    
    
    Dim row As listRow
    

    For Each row In targetTbl.ListRows
 
        
        If row.Range.Cells(1, projectCol).Value = ProjectName And _
           row.Range.Cells(1, updateTimeCol).Value = updateTime And _
           row.Range.Cells(1, MNCol).Value = MN And _
           row.Range.Cells(1, salesCol).Value = salesPerson Then
            IsInTargetBOM = True
            Exit Function
        End If
    Next row
    IsInTargetBOM = False
    
    
Exit Function
ErrorHandler:
    ' Example of manually inserted line number and content
    HandleError "Error #" & Err.Number & ": " & Err.Description, procName
    'Resume Next ' Or other error recovery logic
End Function



Function FindColumnInTable(tbl As ListObject, columnName As String) As Integer
Const procName As String = "FindColumnInTable"
On Error GoTo ErrorHandler

    Dim i As Integer
    For i = 1 To tbl.ListColumns.count
        If tbl.ListColumns(i).name = columnName Then
            FindColumnInTable = i
            Exit Function
        End If
    Next i
    FindColumnInTable = 0 ' ���δ�ҵ��У�����0
    
    
    
Exit Function
ErrorHandler:
    ' Example of manually inserted line number and content
    HandleError "Error #" & Err.Number & ": " & Err.Description, procName
    'Resume Next ' Or other error recovery logic
End Function

Sub TransferDataWithProjectName()
Const procName As String = "TransferDataWithProjectName"
On Error GoTo ErrorHandler

    Dim sourceWs As Worksheet, targetWs As Worksheet
    Dim sourceTbl As ListObject, targetTbl As ListObject
    Dim followUpCol As Integer, projectNameCol  As Integer, ReportTimeCol  As Integer, ReportSourceCol  As Integer, ProjectCategoryCol As Integer
    Dim CurrentSalesCol  As Integer, BrandCol  As Integer, PhaseStatusCol  As Integer, AgentCol As Integer, DistributorCol As Integer, SysIntergrateCol As Integer, DesignCol As Integer
    Dim i As Integer, targetRow As Integer, ReportSource As String
    Dim cell As Range, ProjectName As String, contents() As String, ReportTime As String, ProjectCategory As String, Design As String
    Dim CurrentSales As String, Brand As String, PhaseStatus As String, Agent As String, Distributor As String, SysIntergrate As String
    Dim dateCol As Integer, importDate As Date, QuotationCol As Integer, Quotation As String
    Dim ID As Integer, BeforProjectName As String
    
    ID = 0
    
    ' ����Դ��Ŀ�깤����ͱ��
    Set sourceWs = ThisWorkbook.Sheets("Source Database")
    Set targetWs = ThisWorkbook.Sheets("Project Database")
    Set sourceTbl = sourceWs.ListObjects(1)
    Set targetTbl = targetWs.ListObjects(1)

    ' �ҵ�Դ����е����������
    followUpCol = FindColumnInTable(sourceTbl, "��Ŀ������¼")
    projectNameCol = FindColumnInTable(sourceTbl, "��Ŀ����")
    ReportTimeCol = FindColumnInTable(sourceTbl, "����ʱ��")
    ReportSourceCol = FindColumnInTable(sourceTbl, "������Դ")
    ProjectCategoryCol = FindColumnInTable(sourceTbl, "��Ŀ����")
    CurrentSalesCol = FindColumnInTable(sourceTbl, "��ǰ����")
    BrandCol = FindColumnInTable(sourceTbl, "Ʒ��")
    PhaseStatusCol = FindColumnInTable(sourceTbl, "�׶�״̬")
    AgentCol = FindColumnInTable(sourceTbl, "������")
    DistributorCol = FindColumnInTable(sourceTbl, "������")
    SysIntergrateCol = FindColumnInTable(sourceTbl, "������")
    DesignCol = FindColumnInTable(sourceTbl, "��Ƶ�λ")
    QuotationCol = FindColumnInTable(sourceTbl, "��ۺϼ�")
    

    ' ȷ���ҵ�����
    If followUpCol = 0 Or projectNameCol = 0 Or ReportSourceCol = 0 Or ReportTimeCol = 0 Or ProjectCategoryCol = 0 Or CurrentSalesCol = 0 Or BrandCol = 0 Or PhaseStatusCol = 0 Then
        MsgBox "δ��Դ�����ҵ���Ҫ���С�"
        Exit Sub
    End If

    ' ��ʼ��Ŀ���к�
    targetRow = targetTbl.ListRows.count + 1
    

    ' ����Դ���е�ÿ����Ԫ��
    For Each cell In sourceTbl.ListColumns(followUpCol).DataBodyRange
    
        importDate = Date - 30 ' �����ʼ�����ǽ���
        ProjectName = cell.Offset(0, projectNameCol - followUpCol).Value ' ��ȡͬ�е���Ŀ����
        ReportTime = cell.Offset(0, ReportTimeCol - followUpCol).Value
        ReportSource = cell.Offset(0, ReportSourceCol - followUpCol).Value
        ProjectCategory = cell.Offset(0, ProjectCategoryCol - followUpCol).Value
        CurrentSales = cell.Offset(0, CurrentSalesCol - followUpCol).Value
        Brand = cell.Offset(0, BrandCol - followUpCol).Value
        PhaseStatus = cell.Offset(0, PhaseStatusCol - followUpCol).Value
        Agent = cell.Offset(0, AgentCol - followUpCol).Value
        Distributor = cell.Offset(0, DistributorCol - followUpCol).Value
        SysIntergrate = cell.Offset(0, SysIntergrateCol - followUpCol).Value
        Design = cell.Offset(0, DesignCol - followUpCol).Value
        Quotation = cell.Offset(0, QuotationCol - followUpCol).Value
        
        If BeforProjectName <> ProjectName Then
            ID = ID + 1
        End If
        
        contents = Split(cell.Value, Chr(10)) ' ʹ�û��з��ָ�

        ' ���ָ����������и��Ƶ�Ŀ����
        For i = LBound(contents) To UBound(contents)
            ' ��Ŀ����ĩβ�������
            targetTbl.ListRows.Add
            ' ���ơ���Ŀ������¼���͡���Ŀ���ơ���Ŀ����
      
            targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "��Ŀ������¼")).Value = "��ʷ��¼��" & contents(i)
            targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "��Ŀ����")).Value = ProjectName
            targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "����ʱ��")).Value = ReportTime
            targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "������Դ")).Value = ReportSource
            targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "Ʒ��")).Value = Brand
            targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "�׶�״̬")).Value = PhaseStatus
            targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "������")).Value = Distributor
            targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "��Ŀ����")).Value = ProjectCategory
            targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "���۸�����")).Value = CurrentSales
            targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "������")).Value = Agent
            targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "��Ʒ�")).Value = Design
            targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "������")).Value = SysIntergrate
            targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "������¼ʱ��")).Value = importDate
            targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "��ۺϼ�")).Value = Quotation
            targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "����ʱ��")).Value = Now
            targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "��Ŀ���")).Value = ID
            
        
            targetRow = targetRow + 1
            importDate = importDate + 2 ' �������������һ��
            BeforProjectName = ProjectName
            
        Next i
    Next cell
    
    targetWs.Shapes("ProjectUpButton").Visible = msoFalse
    
Exit Sub
ErrorHandler:
    ' Example of manually inserted line number and content
    HandleError "Error #" & Err.Number & ": " & Err.Description, procName
    'Resume Next ' Or other error recovery logic
End Sub

Sub TransferDataWithBOM()
Const procName As String = "TransferDataWithBOM"
On Error GoTo ErrorHandler

    Dim sourceWs As Worksheet, targetWs As Worksheet
    Dim sourceTbl As ListObject, targetTbl As ListObject
    Dim projectNameCol   As Integer, ProductNumberCol  As Integer, QuantityCol  As Integer, ProductNameCol As Integer
    Dim CurrentSalesCol  As Integer, BrandCol  As Integer, ProductModelCol  As Integer, ListPriceCol As Integer, TotalListPricerCol As Integer
    Dim i As Integer, targetRow As Integer, ReportSource As String
    Dim cell As Range, ProjectName As String, ProductNumber As String, Quantity As String, ProductName As String
    Dim CurrentSales As String, Brand As String, ProductModel As String, Specification As String, ListPrice As String, TotalListPrice As String
    Dim SpecificationCol As Integer, ID As Integer, BeforProjectName As String
    
    
    ' ����Դ��Ŀ�깤����ͱ��
    Set sourceWs = ThisWorkbook.Sheets("Source Bom")
    Set targetWs = ThisWorkbook.Sheets("BOM")
    Set sourceTbl = sourceWs.ListObjects(1)
    Set targetTbl = targetWs.ListObjects(1)

    ' �ҵ�Դ����е����������
    projectNameCol = FindColumnInTable(sourceTbl, "���ۻ��ᱨ��")
    ProductNameCol = FindColumnInTable(sourceTbl, "��Ʒ����")
    QuantityCol = FindColumnInTable(sourceTbl, "����")
    ProductModelCol = FindColumnInTable(sourceTbl, "��Ʒ�ͺ�")
    SpecificationCol = FindColumnInTable(sourceTbl, "���")
    BrandCol = FindColumnInTable(sourceTbl, "Ʒ��")
    ListPriceCol = FindColumnInTable(sourceTbl, "���")
    TotalListPricerCol = FindColumnInTable(sourceTbl, "��ۺϼ�")
    ProductNumberCol = FindColumnInTable(sourceTbl, "��Ʒ���")
    
   

    ' ȷ���ҵ�����
    If projectNameCol = 0 Or ProductNumberCol = 0 Or QuantityCol = 0 Or ProductModelCol = 0 Or SpecificationCol = 0 Or BrandCol = 0 Or ProductModelCol = 0 Or ListPriceCol = 0 Then
        MsgBox "δ��Դ�����ҵ���Ҫ���С�"
        Exit Sub
    End If

    ' ��ʼ��Ŀ���к�
    targetRow = targetTbl.ListRows.count + 1
     ID = 0
   

    ' ����Դ���е�ÿ����Ԫ��
    For Each cell In sourceTbl.ListColumns(projectNameCol).DataBodyRange
        'importDate = Date - 30 ' �����ʼ�����ǽ���

        
        ProjectName = cell.Offset(0, projectNameCol - projectNameCol).Value
        Quantity = cell.Offset(0, QuantityCol - projectNameCol).Value
        ProductModel = cell.Offset(0, ProductModelCol - projectNameCol).Value
        Specification = cell.Offset(0, SpecificationCol - projectNameCol).Value
        Brand = cell.Offset(0, BrandCol - projectNameCol).Value
        ListPrice = cell.Offset(0, ListPriceCol - projectNameCol).Value
        TotalListPricer = cell.Offset(0, TotalListPricerCol - projectNameCol).Value
        ProductNumber = cell.Offset(0, ProductNumberCol - projectNameCol).Value
        ProductName = cell.Offset(0, ProductNameCol - projectNameCol).Value
        
        If BeforProjectName <> ProjectName Then
            ID = ID + 1
        End If
          ' ��Ŀ����ĩβ�������
          targetTbl.ListRows.Add
          ' ���ơ���Ŀ������¼���͡���Ŀ���ơ���Ŀ����
    
          'targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "��Ŀ������¼")).Value = "��ʷ��¼��" & contents(i)
          targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "Project Name")).Value = ProjectName
          targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "Product name")).Value = ProductName
          targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "Model")).Value = ProductModel
          targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "Specification")).Value = Specification
          targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "Brand")).Value = Brand
          targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "Unit")).Value = "each"
          targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "Quantity")).Value = Quantity
          targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "Unit. price")).Value = ListPrice
          targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "Ext. Price")).Value = TotalListPricer
          targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "MN")).Value = ProductNumber
          targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "update time")).Value = Now
          targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "Sale")).Value = ThisWorkbook.Sheets("Config").Range("A1")
          targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "Discount")).Value = "100%"
          targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "Discount Unit Price")).Value = ListPrice
          targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "Discount Ext. Price")).Value = TotalListPricer
          targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "Date")).Value = Date
          targetTbl.ListRows(targetRow).Range.Cells(1, FindColumnInTable(targetTbl, "OrderNumber")).Value = ID
          
  
         
            targetRow = targetRow + 1
            BeforProjectName = ProjectName
            'importDate = importDate + 2 ' �������������һ��
            
     
    Next cell
    targetWs.Shapes("BomUpButton").Visible = msoFalse
    
Exit Sub
ErrorHandler:
    ' Example of manually inserted line number and content
    HandleError "Error #" & Err.Number & ": " & Err.Description, procName
    'Resume Next ' Or other error recovery logic
End Sub

Function getpychar(char) As String
Const procName As String = "getpychar"
On Error GoTo ErrorHandler

    tmp = 65536 + Asc(char)
    
    If (tmp >= 45217 And tmp <= 45252) Then
    
    getpychar = "A"
    
    ElseIf (tmp >= 45253 And tmp <= 45760) Then
    
    getpychar = "B"
    
    ElseIf (tmp >= 45761 And tmp <= 46317) Then
    
    getpychar = "C"
    
    ElseIf (tmp >= 46318 And tmp <= 46825) Then
    
    getpychar = "D"
    
    ElseIf (tmp >= 46826 And tmp <= 47009) Then
    
    getpychar = "E"
    
    ElseIf (tmp >= 47010 And tmp <= 47296) Then
    
    getpychar = "F"
    
    ElseIf (tmp >= 47297 And tmp <= 47613) Then
    
    getpychar = "G"
    
    ElseIf (tmp >= 47614 And tmp <= 48118) Then
    
    getpychar = "H"
    
    ElseIf (tmp >= 48119 And tmp <= 49061) Then
    
    getpychar = "J"
    
    ElseIf (tmp >= 49062 And tmp <= 49323) Then
    
    getpychar = "K"
    
    ElseIf (tmp >= 49324 And tmp <= 49895) Then
    
    getpychar = "L"
    
    ElseIf (tmp >= 49896 And tmp <= 50370) Then
    
    getpychar = "M"
    
    ElseIf (tmp >= 50371 And tmp <= 50613) Then
    
    getpychar = "N"
    
    ElseIf (tmp >= 50614 And tmp <= 50621) Then
    
    getpychar = "O"
    
    ElseIf (tmp >= 50622 And tmp <= 50905) Then
    
    getpychar = "P"
    
    ElseIf (tmp >= 50906 And tmp <= 51386) Then
    
    getpychar = "Q"
    
    ElseIf (tmp >= 51387 And tmp <= 51445) Then
    
    getpychar = "R"
    
    ElseIf (tmp >= 51446 And tmp <= 52217) Then
    
    getpychar = "S"
    
    ElseIf (tmp >= 52218 And tmp <= 52697) Then
    
    getpychar = "T"
    
    ElseIf (tmp >= 52698 And tmp <= 52979) Then
    
    getpychar = "W"
    
    ElseIf (tmp >= 52980 And tmp <= 53640) Then
    
    getpychar = "X"
    
    ElseIf (tmp >= 53679 And tmp <= 54480) Then
    
    getpychar = "Y"
    
    ElseIf (tmp >= 54481 And tmp <= 62289) Then
    
    getpychar = "Z"
    
    Else '����������ģ��򲻴���
    
    getpychar = char
    
    End If


Exit Function
ErrorHandler:
    ' Example of manually inserted line number and content
    HandleError "Error #" & Err.Number & ": " & Err.Description, procName
    'Resume Next ' Or other error recovery logic
End Function

'���ȡASC��

Function getpy(str)

For i = 1 To Len(str)

getpy = getpy & getpychar(Mid(str, i, 1))

Next i


Exit Function
ErrorHandler:
    ' Example of manually inserted line number and content
    HandleError "Error #" & Err.Number & ": " & Err.Description, procName
    'Resume Next ' Or other error recovery logic
End Function
Function GetFirstFourPY(str As Variant)
Const procName As String = "GetFirstFourPY"
On Error GoTo ErrorHandler

    Dim result As String
    Dim i As Integer
    Dim ch As String
    Dim pyChar As String
    Dim count As Integer
    Dim ws As Worksheet
    
    
    count = 0
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        If IsChinese(ch) Then
            pyChar = getpychar(ch)
            result = result & pyChar
            count = count + 1
            If count = 4 Then Exit For
        Else
            Exit For
        End If
    Next i

    If result <> "" Then
        GetFirstFourPY = str & "[" & result & "]"
    Else
        GetFirstFourPY = str
    End If
    
    
    
Exit Function
ErrorHandler:
    ' Example of manually inserted line number and content
    HandleError "Error #" & Err.Number & ": " & Err.Description, procName
    'Resume Next ' Or other error recovery logic
End Function

Function IsChinese(ch As String) As Boolean
Const procName As String = "IsChinese"
On Error GoTo ErrorHandler

    IsChinese = AscW(ch) > 127 ' �򵥵ķ�ʽ�ж��Ƿ�Ϊ�����ַ�
    
    
Exit Function
ErrorHandler:
    ' Example of manually inserted line number and content
    HandleError "Error #" & Err.Number & ": " & Err.Description, procName
    'Resume Next ' Or other error recovery logic
End Function


Sub DownloadFileFromGitHub()
    Dim url As String
    Dim filePath As String
    Dim fileName As String
    Dim script As String

    ' Set the URL of the file to be downloaded
    url = "https://github.com/James99309/VBA--Update-1.0/raw/main/��Ʒ����飨E-ANTD%20HP���ⶨ���״���ߣ�.doc"

    ' Set the name of the file
    fileName = "��Ʒ�����.doc" ' Modify as needed

    ' Get the desktop path using the function defined above
    filePath = GetDesktopPath() & fileName

    ' Construct the AppleScript command
    script = "do shell script ""curl -L -o '" & filePath & "' '" & url & "'"""

    ' Run the AppleScript command from VBA
    MacScript (script)

    MsgBox "File downloaded successfully to Desktop!", vbInformation
End Sub


Function GetDesktopPath() As String
    ' AppleScript����ص�ǰ�û�������·��
    Dim script As String
    script = "return path to desktop folder as string"
    
    ' ִ��AppleScript�����ȡ���
    GetDesktopPath = MacScript(script)
End Function
