unit UfcExpertInterface;

interface

uses Forms, windows, ExptIntf, ToolIntf;

type
  TNLDPropertyExpert = class(TIExpert)
   public
     constructor Create; virtual;
     destructor Destroy; override;

     { Expert Style }
     function GetStyle: TExpertStyle; override;
     function GetIDString: string; override;
     function GetName: string; override;
     function GetAuthor: string; override;
     function GetMenuText: string; override;
     function GetState: TExpertState; override;
     function GetGlyph: HICON; override;
     function GetComment: string; override;
     function GetPage: string; override;

     { Expert Action }
     procedure Execute; override;
   protected
     procedure OnClick(Sender: TIMenuItemIntf); virtual;
   private
     MenuItem: TIMenuItemIntf;
   end;

   procedure Register;

implementation

uses Dialogs, Menus, Classes, UfmExpertMainForm;

procedure Register;
begin
  RegisterLibraryExpert(TNLDPropertyExpert.Create)
end;

{ TNLDPropertyExpert }

constructor TNLDPropertyExpert.Create;
 var
   Main: TIMainMenuIntf;
   ToolsTools: TIMenuItemIntf;
   Tools: TIMenuItemIntf;
 begin
   inherited Create;
   MenuItem := nil;
   if ToolServices <> nil then
   begin
     Main := ToolServices.GetMainMenu;
     if Main <> nil then { we've got the main menu }
     try
       ToolsTools := Main.FindMenuItem('ToolsToolsItem');
       if ToolsTools <> nil then { we've got the suh-menuitem }
       try
         Tools := ToolsTools.GetParent;
         if Tools <> nil then { we've got the Tools menu }
         try
           MenuItem := Tools.InsertItem(ToolsTools.GetIndex+1,
                                       '&Property Modifier',
                                        'NLDPropModifier','',
                                         ShortCut(Ord('P'),[ssCtrl]),0,0,
                                        [mfEnabled, mfVisible], OnClick)
         finally
           Tools.DestroyMenuItem
         end
       finally
         ToolsTools.DestroyMenuItem
       end
     finally
       Main.Free
     end
   end
end;

destructor TNLDPropertyExpert.Destroy;
begin
   if MenuItem <> nil then MenuItem.DestroyMenuItem;
   inherited Destroy
end;

procedure TNLDPropertyExpert.Execute;
begin
  inherited;

end;

function TNLDPropertyExpert.GetAuthor: string;
begin
  Result:= 'Walter Heck (on behalf of http://www.nldelphi.com)'
end;

function TNLDPropertyExpert.GetComment: string;
begin

end;

function TNLDPropertyExpert.GetGlyph: HICON;
begin
  Result:=0;
end;

function TNLDPropertyExpert.GetIDString: string;
begin
  Result:='NLDelphi.TNLDPropertyExpert'
end;

function TNLDPropertyExpert.GetMenuText: string;
begin
  Result:='NLDelphi Dingetje'
end;

function TNLDPropertyExpert.GetName: string;
begin
  Result:='Property Modifier Wizard'
end;

function TNLDPropertyExpert.GetPage: string;
begin

end;

function TNLDPropertyExpert.GetState: TExpertState;
begin

end;

function TNLDPropertyExpert.GetStyle: TExpertStyle;
begin
  Result:=esAddIn;
end;

procedure TNLDPropertyExpert.OnClick(Sender: TIMenuItemIntf);
var
  fmExpertMainForm: TfmExpertMainForm;
begin
     fmExpertMainForm:=TfmExpertMainForm.Create(Application);
     fmExpertMainForm.Show;
end;

end.
 