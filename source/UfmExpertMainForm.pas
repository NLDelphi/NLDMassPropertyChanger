unit UfmExpertMainForm;

interface

uses
  SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, ExtCtrls, StdCtrls, ComCtrls,
  ActnList, ImgList, ToolWin, Types;

type
  TfmExpertMainForm = class(TForm)
    ToolBar: TToolBar;
    tbnSave: TToolButton;
    Actions: TActionList;
    actFileSave: TAction;
    ilActions: TImageList;
    StringGrid: TStringGrid;
    procedure StringGridSetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: string);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure StringGridResize(Sender: TObject);
    procedure actFileSaveExecute(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure StringGridDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure StringGridSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
  private
    FComponentList: TInterfaceList;
    procedure FillComponentList;
    procedure PopulateGrid;
  private
    FModified: Boolean;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;

implementation

{$R *.dfm}

uses
  Printers, Windows,
  ToolsAPI, UfcGExpertsFuncs;

const // Do not localize.
  SHintPropertyName = 'Hint';
  SNamePropertyName = 'Name';

resourcestring
  SNotAvailable = '';

procedure TfmExpertMainForm.StringGridSetEditText(Sender: TObject; ACol,
  ARow: Integer; const Value: string);
begin
  FModified := True;
end;

procedure TfmExpertMainForm.FormClose(Sender: TObject; var Action:
  TCloseAction);
resourcestring
  SSaveChanges =
    'You have unsaved changes; do you want to save these changes before closing?';
begin
  if FModified then
    if MessageDlg(SSaveChanges, mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      actFileSave.Execute;
  Action := caFree;
end;

procedure TfmExpertMainForm.FillComponentList;

  procedure FillComponentListFromComponent(const AComponent: IOTAComponent);
  var
    i: Integer;
    RetrievedComponent: IOTAComponent;
  begin
    Assert(Assigned(AComponent));

    FComponentList.Add(AComponent);

    for i := 0 to AComponent.GetComponentCount - 1 do
    begin
      RetrievedComponent := AComponent.GetComponent(i);
      FillComponentListFromComponent(RetrievedComponent);
    end;
  end;

var
  Module: IOTAModule;
  FormEditor: IOTAFormEditor;
  RootComponent: IOTAComponent;
begin
  if not IsForm(GxOtaGetCurrentSourceFile) then
  begin
    MessageDlg('This expert is for use in form files only.', mtError, [mbOK], 0);
    Exit;
  end;

  Module := GxOtaGetCurrentModule;
  Assert(Assigned(Module));

  FormEditor := GxOtaGetFormEditorFromModule(Module);
  Assert(Assigned(FormEditor));

  RootComponent := FormEditor.GetRootComponent;
  Assert(Assigned(RootComponent));

  FillComponentListFromComponent(RootComponent);
end;

procedure TfmExpertMainForm.PopulateGrid;

  function ComponentName(const AComponent: IOTAComponent): string;
  begin
    Result := GxOtaGetComponentPropertyAsString(AComponent, SNamePropertyName);
    if Result = '' then
      Result := SNotAvailable;
  end;

  function ComponentClass(const AComponent: IOTAComponent): string;
  begin
    Result := AComponent.GetComponentType;
    if Result = '' then
      Result := SNotAvailable;
  end;

  function ParentName(const AComponent: IOTAComponent): string;
  begin
    Result := GxOtaGetComponentParentName(AComponent);
    if Result = '' then
      Result := SNotAvailable;
  end;

  function ComponentHint(const AComponent: IOTAComponent): string;
  begin
    Result := GxOtaGetComponentPropertyAsString(AComponent, 'Hint');
    if Result = '' then
      Result := SNotAvailable;
  end;

resourcestring
  SCellComponent = 'Component';
var
  CurrentRow: Integer;
  AComponent: IOTAComponent;
  AComponentName: string;
  i: Integer;
begin
  Assert(Assigned(FComponentList));

  StringGrid.RowCount := 1;
  StringGrid.ColCount := 4;

  StringGrid.Cells[0, 0] := SCellComponent;
  StringGrid.Cells[1, 0] := 'Parent';
  StringGrid.Cells[2, 0] := 'Class';
  StringGrid.Cells[3, 0] := SHintPropertyName;

  for i := 0 to FComponentList.Count - 1 do
  begin
    AComponent := FComponentList.Items[i] as IOTAComponent;
    Assert(Assigned(AComponent));

    AComponentName := ComponentName(AComponent);
    if Length(AComponentName) = 0 then
      Continue;

    CurrentRow := StringGrid.RowCount;
    StringGrid.RowCount := StringGrid.RowCount + 1;
    StringGrid.Cells[0, CurrentRow] := AComponentName;
    StringGrid.Cells[1, CurrentRow] := ParentName(AComponent);
    StringGrid.Cells[2, CurrentRow] := ComponentClass(AComponent);
    StringGrid.Cells[3, CurrentRow] := ComponentHint(AComponent);
    StringGrid.Objects[0, CurrentRow] := Pointer(i);
  end;

  StringGrid.FixedCols := 0;
  StringGrid.FixedRows := 1;
end;

constructor TfmExpertMainForm.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  FComponentList := TInterfaceList.Create;
  FModified := False;

  FillComponentList;
  PopulateGrid;
end;

destructor TfmExpertMainForm.Destroy;
begin
  FreeAndNil(FComponentList);

  inherited Destroy;
end;

procedure TfmExpertMainForm.StringGridResize(Sender: TObject);
const
  MinimumLastColWidth = 100;
var
  SendingGrid: TStringGrid;
  i: Integer;
  WidthDelta: Integer;
  LastColWidth: Integer;
begin
  SendingGrid := StringGrid;
  Assert(Assigned(SendingGrid));

  WidthDelta := SendingGrid.ClientWidth;
  for i := 0 to SendingGrid.ColCount - 1 do
    Dec(WidthDelta, SendingGrid.ColWidths[i]);
  Dec(WidthDelta, SendingGrid.ColCount * SendingGrid.GridLineWidth);

  LastColWidth := SendingGrid.ColWidths[SendingGrid.ColCount - 1];

  Inc(LastColWidth, WidthDelta);
  if LastColWidth < MinimumLastColWidth then
    LastColWidth := MinimumLastColWidth;

  SendingGrid.ColWidths[SendingGrid.ColCount - 1] := LastColWidth;
end;

procedure TfmExpertMainForm.actFileSaveExecute(Sender: TObject);
var
  i: Integer;
  ComponentIndex: Integer;
  AComponent: IOTAComponent;
  Value: string;
begin
  Assert(Assigned(FComponentList));

  for i := 1 to StringGrid.RowCount - 1 do
  begin
    ComponentIndex := Integer(StringGrid.Objects[0, i]);

    AComponent := FComponentList.Items[ComponentIndex] as IOTAComponent;
    Assert(Assigned(AComponent));

    Value := StringGrid.Cells[0, i];
    GxOtaSetComponentPropertyAsString(AComponent, 'Name', Value);

    Value := StringGrid.Cells[3, i];
    GxOtaSetComponentPropertyAsString(AComponent, 'Hint', Value);

  end;
  FModified := False;
  ModalResult := mrOK;
  PopulateGrid;
end;

procedure TfmExpertMainForm.FormKeyDown(Sender: TObject; var Key: Word; Shift:
  TShiftState);
begin
  if Key = VK_ESCAPE then
  begin
    Key := 0;
    Close;
  end;
end;

procedure TfmExpertMainForm.StringGridDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
begin
  with (Sender as TStringGrid) do
    if (ACol in [1, 2]) and (ARow >= FixedRows) then
      with Canvas do
      begin
        Brush.Color := clMedGray;
        FillRect(Rect);
        TextOut(Rect.Left + 2, Rect.Top +2, Cells[ACol, ARow]);
      end;
end;

procedure TfmExpertMainForm.StringGridSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
begin
  CanSelect := not (ACol in [1, 2]);
  actFileSave.Execute;
end;

end.

