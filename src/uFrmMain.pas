unit uFrmMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls;

type
  TfrmMain = class(TForm)
    edtFirstFile: TButtonedEdit;
    edtSecondFile: TButtonedEdit;
    btnCompare: TButton;
    mmResult: TMemo;
    procedure btnCompareClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses
  uExcelComparer;


{$R *.dfm}

procedure TfrmMain.btnCompareClick(Sender: TObject);
var
  c: TExcelComparer;
  r: RExcelDifference;
begin
  c := TExcelComparer.Create();
    try
    if c.CompareCSV(edtFirstFile.Text,edtSecondFile.Text,r) then
      mmResult.Lines.Add('Files are equal.')
    else
      if c.LastError.IsEmpty then
        mmResult.Lines.Add(Format('Found diff. Sheet %d. Row %d.', [r.Sheet,r.Row]))
      else
        mmResult.Lines.Add('Something went wrong. '+c.LastError);
    finally
      c.Free;
    end;
end;

end.
