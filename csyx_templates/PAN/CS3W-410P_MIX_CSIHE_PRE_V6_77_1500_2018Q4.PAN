PVObject_=pvModule
  Version=6.77
  Flags=$00500443

  PVObject_Commercial=pvCommercial
    Comment=www.canadiansolar.com
    Flags=$0041
    Manufacturer=Canadian Solar Inc.
    Model=CS3W-410P 1500V HE
    DataSource=Preliminary 2019 TUV-SUD data
    YearBeg=2019
    Width=1.048
    Height=2.108
    Depth=0.040
    Diameter=24.900
    Weight=24.900
    NPieces=0
    PriceDate=22/07/16 16:29
    Currency=EUR
    Remarks, Count=3
      Str_1=Frame: Anodized aluminium alloy
      Str_2=Structure: 3.2mm Glass / EVA / Backsheet
      Str_3=Connections: Cable, T4 series
    End of Remarks=Connections: Cable, T4 series
  End of PVObject pvCommercial

  Technol=mtSiPoly
  NCelS=72
  NCelP=2
  NDiode=3
  SubModuleLayout=slTwinHalfCells
  GRef=1000
  TRef=25.0
  PNom=410.0
  PNomTolLow=0.00
  PNomTolUp=1.20
  Isc=11.060
  Voc=47.60
  Imp=10.490
  Vmp=39.10
  muISC=5.49
  muVocSpec=-138.0
  muPmpReq=-0.370
  RShunt=550
  Rp_0=2000
  Rp_Exp=5.50
  RSerie=0.251
  Gamma=1.027
  muGamma=-0.0004
  VMaxIEC=1500
  VMaxUL=1500
  Absorb=0.90
  ARev=3.200
  BRev=11.040
  RDiode=0.010
  VRevDiode=-0.70
  AirMassRef=1.500
  CellArea=137.8
  SandiaAMCorr=50.000
  RelEffic800=0.43
  RelEffic400=-1.26
  RelEffic200=-3.40

  PVObject_IAM=pvIAM
    Flags=$00
    IAMMode=UserProfile
    IAMProfile=TCubicProfile
      NPtsMax=9
      NPtsEff=9
      LastCompile=$B18D
      Mode=3
      Point_1=10.0,1.00000
      Point_2=20.0,1.00000
      Point_3=30.0,1.00000
      Point_4=40.0,0.99000
      Point_5=50.0,0.99000
      Point_6=60.0,0.97000
      Point_7=70.0,0.92000
      Point_8=80.0,0.76000
      Point_9=90.0,0.00000
    End of TCubicProfile
  End of PVObject pvIAM

  OperPoints, list of=3 tOperPoint
    Point_1=False,800,25.0,0.43,0.00,0.000,0.000,0.00
    Point_2=False,400,25.0,-1.26,0.00,0.000,0.000,0.00
    Point_3=False,200,25.0,-3.40,0.00,0.000,0.000,0.00
  End of List OperPoints
End of PVObject pvModule
