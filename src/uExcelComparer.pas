unit uExcelComparer;

interface

uses
  System.Classes;

type

  RExcelDifference = record
    Sheet: integer;
    Row: Integer;
  private
    procedure Init;
  public
    function Equals: Boolean;
  end;

  TExcelComparerErr = procedure(Sender: TObject; AMsg: string) of object;

  TExcelComparer = class(TObject)
  private
    FLastError: string;
    FOnError: TExcelComparerErr;
    FTempFilesList: TStringList;
    function CheckExcelInstalled: Boolean;
    procedure RunExcel(DisableAlerts: Boolean = False; Visible: boolean = false);
    function StopExcel: boolean;
    function TempFile(AExtension: string): string;
    property TempFilesList: TStringList read FTempFilesList;
  protected
    procedure DoError(AMsg: string);
  public
    constructor Create;
    destructor Destroy; override;
    function Compare(AFile1, AFile2: string; out ADiff: RExcelDifference): Boolean;
    function CompareCSV(AFile1, AFile2: string; out ADiff: RExcelDifference): Boolean;
    property LastError: string read FLastError;
    property OnError: TExcelComparerErr read FOnError write FOnError;
  end;

implementation

uses
  System.SysUtils, Vcl.Dialogs, Winapi.ActiveX, System.Win.ComObj,
  System.Variants, System.IOUtils;

type
  EExcelComparerException = Exception;

const ExcelApp = 'Excel.Application';

var
  MyExcel: OleVariant;

procedure TestErr(AMustBeTrue: Boolean; AMsg: string);
begin
if not AMustBeTrue then
  raise EExcelComparerException.Create(AMsg);
end;

constructor TExcelComparer.Create;
begin
  inherited;
  MyExcel:= Unassigned;
  FTempFilesList := TStringList.Create();
end;

destructor TExcelComparer.Destroy;

  procedure CleanUpTemp();
  {Delete all files produced by TempFile()}
  begin
  while TempFilesList.Count>0 do
    try
      DeleteFile(TempFilesList[0]);
    finally
    TempFilesList.Delete(0);
    end;
  end;

begin
  CleanUpTemp();
  FreeAndNil(FTempFilesList);
  inherited Destroy;
end;

function TExcelComparer.CheckExcelInstalled: Boolean;
{Searching CLSID OLE Excel}
var
  ClassID: TCLSID;
begin
  Result:=CLSIDFromProgID(PWideChar(WideString(ExcelApp)),ClassID)=S_OK;
  TestErr(Result,'Found no installed MS Excel on this machine.');
end;

function TExcelComparer.Compare(AFile1, AFile2: string; out ADiff: RExcelDifference): Boolean;
{Returns 0 for equal otherwise first diff row}

  procedure CompareRanges(const ARange1, ARange2: Variant);
  {Cell to cell compare. Extremly slow.}
  var
    i,j: Integer;
    iCol,iRow: integer;
  begin
  iCol:= ARange1.Columns.Count;
  iRow:= ARange1.Rows.Count;
  for i:=1 to iRow do
    for j:=1 to iCol do
      if VarCompareValue(ARange1.Cells[i,j].Value,ARange2.Cells[i,j].Value)<>vrEqual then
        begin
        ADiff.Row:=i;
        Break;
        end;
  end;

var
  WB1,WB2: OleVariant; //WorkBooks
  iSht: Integer;  //Iterator per sheet
  Rng1,Rng2: OleVariant;//Used ranges on WorkSheets
begin
FLastError:='';
Result:=False;
ADiff.Init();
RunExcel();
  try
  WB1:=MyExcel.Workbooks.Open(AFile1);
  WB2:=MyExcel.Workbooks.Open(AFile2);
  TestErr(WB1.Sheets.Count=WB2.Sheets.Count,'Different count of WorkSheets in files');
  for iSht:=1 to WB1.Sheets.Count do
    begin
    Rng1:=WB1.Sheets[iSht].UsedRange;
    Rng2:=WB2.Sheets[iSht].UsedRange;
    TestErr(Rng1.Columns.Count=Rng2.Columns.Count,'Different count of columns on sheet '+iSht.ToString);
    TestErr(Rng1.Rows.Count=Rng2.Rows.Count,'Different count of rows on sheet '+iSht.ToString);
    CompareRanges(Rng1,Rng2);
    if ADiff.Row<>0 then
      begin
      ADiff.Sheet:=iSht;
      Break;
      end;
    end;
  Result:=ADiff.Row=0;
  except on e: exception do DoError('Comparing files. '+e.Message);
  end;
StopExcel();
end;

function TExcelComparer.CompareCSV(AFile1, AFile2: string; out ADiff: RExcelDifference): Boolean;
{Compares only data, first saving every worksheet as CSV. Much faster.}

  function CompareCsvFiles(AFile1, AFile2: string): Integer;
  {CSV files compare. Returns diff row.}
  var
    i: Integer;
    slFile1,slFile2: TStringList;
  begin
  Result:=0;
  slFile1:=TStringList.Create();
  slFile2:=TStringList.Create();
    try
    slFile1.LoadFromFile(AFile1);
    slFile2.LoadFromFile(AFile2);
    TestErr(slFile1.Count=slFile2.Count,'Different row count.');
    i:=0;
    while (i<slFile1.Count) and (Result=0) do
      begin
      if not SameText(slFile1[i],slFile2[i]) then
        Result:=i;
      Inc(i);
      end;
    finally
      FreeAndNil(slFile1);
      FreeAndNil(slFile2);
    end;
  end;

type
  RCsvExcelPair = record
    File1,File2: string;
  end;
  TCsvExcelPairArray = array of RCsvExcelPair;

const
  xlCSV=6;	//file format CSV	*.csv
var
  WB1,WB2: OleVariant; //WorkBooks
  iSht: Integer;  //Iterator per sheet
  sFileCsv1, sFileCsv2: string;
  arrCSV: TCsvExcelPairArray;
  rCSV: RCsvExcelPair;
begin
Result:=False;
FLastError:='';
ADiff.Init();
RunExcel();
  try
  WB1:=MyExcel.Workbooks.Open(AFile1);
  WB2:=MyExcel.Workbooks.Open(AFile2);
  TestErr(WB1.Sheets.Count=WB2.Sheets.Count,'Different count of WorkSheets in files.');
  SetLength(arrCSV,integer(WB1.Sheets.Count));
  for iSht:=1 to WB1.Sheets.Count do
    begin
    sFileCsv1:=TempFile('.csv');
    sFileCsv2:=TempFile('.csv');
    WB1.Sheets[iSht].SaveAs(sFileCsv1,xlCSV);
    WB2.Sheets[iSht].SaveAs(sFileCsv2,xlCSV);
    arrCSV[iSht-1].File1:=sFileCsv1;
    arrCSV[iSht-1].File2:=sFileCsv2;
    end;
  MyExcel.Workbooks.Close;
  for iSht:=Low(arrCSV) to High(arrCSV) do
    begin
    rCsv:=arrCSV[iSht];
    ADiff.Row:=CompareCsvFiles(rCSV.File1,rCSV.File2);
    if ADiff.Row<>0 then
      begin
      ADiff.Sheet:=iSht+1;
      Break;
      end;
    end;
  Result:=ADiff.Row=0;
  except on e: exception do DoError('Comparing files CSV mode. '+e.Message);
  end;
StopExcel();
end;

procedure TExcelComparer.DoError(AMsg: string);
begin
  FLastError:=AMsg;
  if Assigned(FOnError) then
    begin
    FOnError(Self, AMsg);
    end;
end;

procedure TExcelComparer.RunExcel(DisableAlerts: Boolean = False; Visible: boolean = false);
{Just run if installed}
begin
  try
  CheckExcelInstalled();
  MyExcel:=CreateOleObject(ExcelApp);
    MyExcel.Application.EnableEvents:=DisableAlerts;
    MyExcel.Application.DisplayAlerts:= DisableAlerts;
    MyExcel.Visible:=Visible;
  except on e: exception do DoError(format('Starting excel. %s.',[e.Message]));
  end;
end;

function TExcelComparer.StopExcel: boolean;
{Stop if there is assigned one}
begin
Result:=false;
if not VarIsEmpty(MyExcel) then
  try
  if MyExcel.Visible then
    MyExcel.Visible:=false;
  MyExcel.Quit;
  MyExcel:=Unassigned;
  Result:=True;
  except  on e: exception do DoError(format('Stopping excel. %s.',[e.Message]));
  end;
end;

function TExcelComparer.TempFile(AExtension: string): string;
begin
  try
  Result:=TPath.GetTempFileName;
  TempFilesList.Add(Result);
  except on E: Exception do DoError('Getting temp file. '+e.Message);
  end;
end;

function RExcelDifference.Equals: Boolean;
begin
  Result := (Row=0) and (Sheet=0);
end;

procedure RExcelDifference.Init;
begin
  Row:=0;
  Sheet:=0;
end;

end.
