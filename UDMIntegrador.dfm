object DMIntegrador: TDMIntegrador
  OldCreateOrder = False
  Left = 406
  Top = 193
  Height = 325
  Width = 542
  object mAvisos: TClientDataSet
    Active = True
    Aggregates = <>
    Params = <>
    Left = 74
    Top = 69
    Data = {
      8C0000009619E0BD0100000018000000050000000000030000008C0002494404
      00010000000000044E6F6D650100490000000100055749445448020002006400
      045469706F0100490000000100055749445448020002001400034F6273010049
      000000010005574944544802000200C800085469706F5F526567010049000000
      0100055749445448020002000A000000}
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
  end
  object dsmAviso: TDataSource
    DataSet = mAvisos
    Left = 123
    Top = 67
  end
end
