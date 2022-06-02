object DMIntegrador: TDMIntegrador
  OldCreateOrder = False
  Left = 406
  Top = 193
  Height = 325
  Width = 542
  object mAvisos: TClientDataSet
    Active = True
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'ID'
        DataType = ftInteger
      end
      item
        Name = 'Nome'
        DataType = ftString
        Size = 100
      end
      item
        Name = 'Tipo'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'Obs'
        DataType = ftString
        Size = 200
      end
      item
        Name = 'Tipo_Reg'
        DataType = ftString
        Size = 10
      end
      item
        Name = 'Referencia'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 74
    Top = 69
    Data = {
      AB0000009619E0BD010000001800000006000000000003000000AB0002494404
      00010000000000044E6F6D650100490000000100055749445448020002006400
      045469706F0100490000000100055749445448020002001400034F6273010049
      000000010005574944544802000200C800085469706F5F526567010049000000
      0100055749445448020002000A000A5265666572656E63696101004900000001
      000557494454480200020014000000}
    object mAvisosID: TIntegerField
      FieldName = 'ID'
    end
    object mAvisosNome: TStringField
      FieldName = 'Nome'
      Size = 100
    end
    object mAvisosTipo: TStringField
      FieldName = 'Tipo'
    end
    object mAvisosObs: TStringField
      FieldName = 'Obs'
      Size = 200
    end
    object mAvisosTipo_Reg: TStringField
      FieldName = 'Tipo_Reg'
      Size = 10
    end
    object mAvisosReferencia: TStringField
      FieldName = 'Referencia'
    end
  end
  object dsmAviso: TDataSource
    DataSet = mAvisos
    Left = 123
    Top = 67
  end
  object frxReport1: TfrxReport
    Version = '5.6.8'
    DotMatrixReport = False
    IniFile = '\Software\Fast Reports'
    PreviewOptions.Buttons = [pbPrint, pbLoad, pbSave, pbExport, pbZoom, pbFind, pbOutline, pbPageSetup, pbTools, pbEdit, pbNavigator, pbExportQuick]
    PreviewOptions.Zoom = 1.000000000000000000
    PrintOptions.Printer = 'Default'
    PrintOptions.PrintOnSheet = 0
    ReportOptions.CreateDate = 42450.581960740700000000
    ReportOptions.LastChange = 42450.622028217590000000
    ScriptLanguage = 'PascalScript'
    StoreInDFM = False
    Left = 222
    Top = 147
  end
  object frxPDFExport1: TfrxPDFExport
    UseFileCache = True
    ShowProgress = True
    OverwritePrompt = False
    DataOnly = False
    PrintOptimized = False
    Outline = False
    Background = False
    HTMLTags = True
    Quality = 95
    Transparency = False
    Author = 'FastReport'
    Subject = 'FastReport PDF export'
    ProtectionFlags = [ePrint, eModify, eCopy, eAnnot]
    HideToolbar = False
    HideMenubar = False
    HideWindowUI = False
    FitWindow = False
    CenterWindow = False
    PrintScaling = False
    PdfA = False
    Left = 262
    Top = 147
  end
  object frxMailExport1: TfrxMailExport
    UseFileCache = True
    ShowProgress = True
    OverwritePrompt = False
    DataOnly = False
    ShowExportDialog = True
    SmtpPort = 25
    UseIniFile = True
    TimeOut = 60
    ConfurmReading = False
    UseMAPI = SMTP
    MAPISendFlag = 0
    Left = 302
    Top = 147
  end
  object frxRichObject1: TfrxRichObject
    Left = 334
    Top = 147
  end
  object frxmAvisos: TfrxDBDataset
    UserName = 'frxmAvisos'
    CloseDataSource = False
    FieldAliases.Strings = (
      'ID=ID'
      'Nome=Nome'
      'Tipo=Tipo'
      'Obs=Obs'
      'Tipo_Reg=Tipo_Reg'
      'Referencia=Referencia')
    DataSource = dsmAviso
    BCDToCurrency = False
    Left = 222
    Top = 195
  end
end
