unit helpers;

interface

uses
  System.SysUtils,
  System.Classes,
  System.JSON,
  System.Net.HttpClient,
  System.Net.URLClient,
  System.StrUtils,
  System.NetEncoding,
  System.Generics.Collections,
  key_press_helper,
  MicrosoftPlanner,
  MicrosoftApiAuthenticator;

type
  THelpers = class(TMsAdapter)
  private
    Fauthenticator: TMsAuthenticator;
    FVerbose: boolean;

    FPlanner: TMsPlanner;
  protected
  public
    constructor Create(Authenticator: TMsAuthenticator); reintroduce;
    destructor Destroy; override;

    function getAllPlanners: TArray<TMsPlannerGroup>;

    property Planner: TMsPlanner read FPlanner;

    class function New(TENANT_ID:string; CLINET_ID: string; REDIRECT_URI: string; REDIRECT_PORT: integer; SCOPE: TArray<string>; Verbose: boolean): THelpers; static;
  end;

implementation

{ THelpers }

constructor THelpers.Create(Authenticator: TMsAuthenticator);
begin
  inherited Create(Authenticator);
  self.Fauthenticator := nil;
  self.FPlanner := TMsPlanner.Create(Authenticator);
end;

destructor THelpers.Destroy;
var
  AAuthenticator: TMsAuthenticator;
begin
  self.FPlanner.Free;
  AAuthenticator := self.Fauthenticator;
  inherited;
  if AAuthenticator <> nil then
    AAuthenticator.Free;
end;

function THelpers.getAllPlanners: TArray<TMsPlannerGroup>;
var
  AIGroup: Integer;
  AGroup: TMsPlannerGroup;
  AIPlanner: Integer;
  APlanner: TMsPlannerPlanner;
  AIBucket: Integer;
  ABucket: TMsPlannerBucket;
begin
  Result := self.FPlanner.GetGroups;
  for AIGroup := 0 to Length(Result) -1 do
  begin
    AGroup := Result[AIGroup];
    self.FPlanner.GetPlanners(AGroup);
    for AIPlanner := 0 to Length(AGroup.Planners) -1 do
    begin
      APlanner := AGroup.Planners[AIPlanner];
      self.FPlanner.GetBuckets(APlanner);
      for AIBucket := 0 to Length(APlanner.Buckets) -1 do
      begin
        ABucket := APlanner.Buckets[AIBucket];
        self.FPlanner.GetTasks(ABucket);
        APlanner.Buckets[AIBucket] := ABucket;
      end;
      AGroup.Planners[AIPlanner] := APlanner;
    end;
    Result[AIGroup] := AGroup;
  end;
end;


class function THelpers.New(TENANT_ID:string; CLINET_ID: string; REDIRECT_URI: string; REDIRECT_PORT: integer; SCOPE: TArray<string>; Verbose: boolean): THelpers;
var
  AAuthenticator: TMsAuthenticator;
begin
  AAuthenticator := TMsAuthenticator.Create(
    ATDelegated,
    TMsClientInfo.Create(
      TENANT_ID,
      CLINET_ID,
      SCOPE,
      TRedirectUri.Create(REDIRECT_PORT, REDIRECT_URI), // YOUR REDIRECT URI (it must be localhost though)
      TMsTokenStorege.Create('microsoft_planner_cli')
    ),
    TMsClientEvents.Create(
    procedure(ResponseInfo: THttpServerResponse)
    begin
      ResponseInfo.ContentStream := TStringStream.Create('<title>Login Succes</title>This tab can be closed now :)');  // YOUR SUCCESS PAGE, do whatever you want here
    end,
    procedure(Error: TMsError)
    begin
      Writeln(Format(  // A premade error message, do whatever you want here
        ''
        + '%sStatus: . . . . . %d : %s'
        + '%sErrorName:  . . . %s'
        + '%sErrorDescription: %s'
        + '%sUrl:  . . . . . . %s %s'
        + '%sData: . . . . . . %s',
        [
          sLineBreak, error.HTTPStatusCode, error.HTTPStatusText,
          sLineBreak, error.HTTPerror_name,
          sLineBreak, error.HTTPerror_description,
          sLineBreak, error.HTTPMethod, error.HTTPurl,
          sLineBreak, error.HTTPerror_data
        ]
      ));
    end,
    procedure(out Cancel: boolean)
    begin
      Cancel := KeyPressed(0);  // Cancel the authentication if a key is pressed
      sleep(0); // if you refresh app-messages here you dont need the sleep
      // Application.ProcessMessages;
    end
    )
  );
  Result := THelpers.Create(AAuthenticator);
  Result.Fauthenticator := AAuthenticator;
  Result.FVerbose := Verbose;
end;

end.
