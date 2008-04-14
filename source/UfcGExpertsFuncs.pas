unit UfcGExpertsFuncs;

interface

uses ToolsAPI;

type
  TPropertyValBuffer = packed record
    case Integer of
      0: (RawBuffer: array[0..255] of Byte);
      1: (SString: ShortString);
      2: (Int: Integer);
      3: (LString: Pointer);
    end;

function IsForm(const FileName: string): Boolean;

// Returns a fully qualified name of the current file,
// which could either be a form or unit (.pas/.cpp/.dfm/.xfm etc.).
// Returns a blank string if no file is currently selected.
function GxOtaGetCurrentSourceFile: string;

// Returns the current module; may return nil of there is no
// current module.
function GxOtaGetCurrentModule: IOTAModule;

// Returns a form editor for Module if it exists; nil otherwise.
function GxOtaGetFormEditorFromModule(const Module: IOTAModule): IOTAFormEditor;

function GxOtaGetComponentPropertyAsString(const AComponent: IOTAComponent;
  const PropertyName: string): string; overload;
function GxOtaGetComponentPropertyAsString(const AComponent: IOTAComponent;
  const PropertyIndex: Integer): string; overload;

function GxOtaGetPropertyIndexByName(const AComponent: IOTAComponent;
  const PropertyName: string): Integer;

// Returns the name of the component's parent if the component has a parent;
// returns an empty string otherwise.
function GxOtaGetComponentParentName(const AComponent: IOTAComponent): string;

function GxOtaSetComponentPropertyAsString(const AComponent: IOTAComponent;
  const PropertyName: string; const Value: string): Boolean; overload;
function GxOtaSetComponentPropertyAsString(const AComponent: IOTAComponent;
  const PropertyIndex: Integer; const Value: string): Boolean; overload;

implementation

uses SysUtils, TypInfo, Classes;

function IsForm(const FileName: string): Boolean;
var
  FileExt: string;
begin
  FileExt := UpperCase(ExtractFileExt(FileName));
  Result := (FileExt = '.DFM') or (FileExt = '.XFM');
end;

function GxOtaGetCurrentSourceFile: string;
var
  Module: IOTAModule;
  Editor: IOTAEditor;
begin
  Result := '';
  Module := GxOtaGetCurrentModule;
  if Module <> nil then
  begin
    Editor := Module.GetCurrentEditor;
    if Editor <> nil then
      Result := Editor.FileName;
  end;
end;

function GxOtaGetCurrentModule: IOTAModule;
var
  ModuleServices: IOTAModuleServices;
begin
  ModuleServices := BorlandIDEServices as IOTAModuleServices;
  Assert(Assigned(ModuleServices));

  Result := ModuleServices.CurrentModule;
end;

function GxOtaGetFormEditorFromModule(const Module: IOTAModule): IOTAFormEditor;
var
  i: Integer;
  Editor: IOTAEditor;
  FormEditor: IOTAFormEditor;
begin
  Result := nil;
  if not Assigned(Module) then
    Exit;
  for i := 0 to Module.GetModuleFileCount-1 do
  begin
    Editor := Module.GetModuleFileEditor(i);
    if Supports(Editor, IOTAFormEditor, FormEditor) then
    begin
      Assert(not Assigned(Result));
      Result := FormEditor;
      // In order to assert our assumptions that only one form
      // is ever associated with a module, do not call Break; here.
    end;
  end;
end;

function GxOtaGetComponentPropertyAsString(const AComponent: IOTAComponent; const PropertyName: string): string;
var
  PropertyIndex: Integer;
begin
  Assert(Assigned(AComponent));

  PropertyIndex := GxOtaGetPropertyIndexByName(AComponent, PropertyName);
  Result := GxOtaGetComponentPropertyAsString(AComponent, PropertyIndex);
end;

function GxOtaGetComponentPropertyAsString(const AComponent: IOTAComponent; const PropertyIndex: Integer): string;
var
  PropertyType: TTypeKind;
  Buffer: TPropertyValBuffer;
begin
  Assert(Assigned(AComponent));

  if PropertyIndex < 0 then
  begin
    Result := '';
    Exit;
  end;

  PropertyType := AComponent.GetPropType(PropertyIndex);

  Buffer.LString := nil;
  AComponent.GetPropValue(PropertyIndex, Buffer);

  case PropertyType of

    tkLString:
      begin
        Result := PChar(Buffer.LString);
      end;

    tkString:
      Result := Buffer.SString;

    tkInteger:
      Result := IntToStr(Buffer.Int);

  else
    Assert(False, 'GxOtaGetComponentPropertyAsString: Unhandled ' +
                  GetEnumName(TypeInfo(TTypeKind), Ord(PropertyType)));
  end;
end;

function GxOtaGetPropertyIndexByName(const AComponent: IOTAComponent; const PropertyName: string): Integer;
begin
  Assert(Assigned(AComponent));

  Result := AComponent.GetPropCount-1;
  while Result >= 0 do
  begin
    if SameText(PropertyName, AComponent.GetPropName(Result)) then
      Break;

    Dec(Result);
  end;
end;

function GxOtaGetComponentParentName(const AComponent: IOTAComponent): string;
resourcestring
  SNoParent = 'No Parent';
var
  Parent: TComponent;
  ComponentHandle: Pointer;
  ComponentObject: TObject;
  NativeComponent: TComponent;
begin
  Assert(Assigned(AComponent));

  // IOTAComponent.GetParent is broken in Delphi 5/6
  Result := SNoParent;
  ComponentHandle := AComponent.GetComponentHandle;
  Assert(Assigned(ComponentHandle));
  ComponentObject := TObject(ComponentHandle);
  if ComponentObject is TComponent then
  begin
    NativeComponent :=  TComponent(ComponentObject);
    Parent := NativeComponent.GetParentComponent;
    if Parent <> nil then
      Result := Parent.Name;
  end;
end;

function GxOtaSetComponentPropertyAsString(const AComponent: IOTAComponent;
  const PropertyName: string; const Value: string): Boolean;
var
  PropertyIndex: Integer;
begin
  Assert(Assigned(AComponent));

  PropertyIndex := GxOtaGetPropertyIndexByName(AComponent, PropertyName);
  Result := GxOtaSetComponentPropertyAsString(AComponent, PropertyIndex, Value);
end;

function GxOtaSetComponentPropertyAsString(const AComponent: IOTAComponent;
  const PropertyIndex: Integer; const Value: string): Boolean;
var
  PropertyType: TTypeKind;
  Buffer: TPropertyValBuffer;
begin
  Assert(Assigned(AComponent));

  if PropertyIndex < 0 then
  begin
    Result := False;
    Exit;
  end;

  PropertyType := AComponent.GetPropType(PropertyIndex);
  case PropertyType of

    tkLString:
      Buffer.LString := PChar(Value);

  else
    Assert(False, 'GxOtaSetComponentPropertyAsString: Unhandled ' +
                  GetEnumName(TypeInfo(TTypeKind), Ord(PropertyType)));
  end;

  Result := AComponent.SetProp(PropertyIndex, Buffer);
end;


end.
