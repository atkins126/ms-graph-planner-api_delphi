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
  listing,
  MicrosoftApiAuthenticator;

type
  THelpers = class(TMsAdapter)
  private
    Fauthenticator: TMsAuthenticator;
    FVerbose: boolean;
    FOptions: TDictionary<string, string>;
    FPlanner: TMsPlanner;

    function getFields: TDictionary<string, string>;
  protected
  public
    constructor Create(Authenticator: TMsAuthenticator); reintroduce;
    destructor Destroy; override;

    function getAllPlanners: TArray<TMsPlannerGroup>;

    procedure list();

    procedure createItem();
    procedure updateItem();
    procedure deleteItem();

    property Planner: TMsPlanner read FPlanner;

    class function New(TENANT_ID:string; CLINET_ID: string; REDIRECT_URI: string; REDIRECT_PORT: integer; SCOPE: TArray<string>; Options: TDictionary<string, string>): THelpers; static;
  end;

var
  REQUESTERROR: Boolean;

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

procedure THelpers.list;
var
  AListing: TListing;
begin
  Alisting := Tlisting.Create(self.FOptions, self.FPlanner);
  AListing.doListing();
  WriteLn(Alisting.Text);
  Alisting.Free;
end;

class function THelpers.New(TENANT_ID:string; CLINET_ID: string; REDIRECT_URI: string; REDIRECT_PORT: integer; SCOPE: TArray<string>; Options: TDictionary<string, string>): THelpers;
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
    var
      req_headers: string;
      res_headers: string;
      headers: string;
      header: TNetHeader;
    begin
      REQUESTERROR := True;

      for header in Error.HTTPreq_Header do req_headers := Format('%s"%s": "%s", ', [req_headers, header.Name, header.Value]);
      for header in Error.HTTPres_Header do res_headers := Format('%s"%s": "%s", ', [res_headers, header.Name, header.Value]);
      headers := Format('{"ReqHeaders": {%s}, "ResHeaders": {%s}}', [req_headers.TrimRight([',', ' ']), res_headers.TrimRight([',', ' '])]);

      Writeln(Format(  // A premade error message, do whatever you want here
        ''
        + '%sStatus: . . . . . %d : %s'
        + '%sErrorName:  . . . %s'
        + '%sErrorDescription: %s'
        + '%sHeader: . . . . . %s'
        + '%sUrl:  . . . . . . %s %s'
        + '%sData: . . . . . . %s',
        [
          sLineBreak, error.HTTPStatusCode, error.HTTPStatusText,
          sLineBreak, error.HTTPerror_name,
          sLineBreak, error.HTTPerror_description,
          sLineBreak, headers,
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
  Result.FOptions := Options;
  Result.FVerbose := Options.ContainsKey('Verbose');
end;

procedure THelpers.createItem;
var
  AFields: TDictionary<string, string>;
  ANewBucket: TMsPlannerBucket;
  ANewTask: TMsPlannerTask;

  AListing: TListing;
begin
  if not self.FOptions.ContainsKey('Fields') then
  begin
    if self.FOptions.ContainsKey('Bucket') then
    begin

      write('Name: ');
      ReadLn(ANewBucket.Name);
      if ANewBucket.Name = '' then begin WriteLn('Name field is required'); exit; end;
      
      write('OrderHint ['''']: ');
      ReadLn(ANewBucket.OrderHint);

      write('PlannerId: ');
      ReadLn(ANewBucket.PlanId);
      if ANewBucket.PlanId = '' then begin WriteLn('PlannerId field is required'); exit; end;

      self.FPlanner.CreateBucket(ANewBucket);

      if not REQUESTERROR then
      begin
        AListing := Tlisting.Create(self.FOptions, self.FPlanner);
        AListing.writeBucket(ANewBucket);
        WriteLn(Alisting.Text);
        AListing.Free;
      end;
    end
    else if self.FOptions.ContainsKey('Task') then
    begin
      Write('Title: ');
      ReadLn(ANewTask.Title);
      if ANewtask.Title = '' then begin WriteLn('Title is required'); exit; end;
      
      write('PercentComplete [''0'']: ');
      ReadLn(ANewTask.PercentComplete);
      if ANewTask.PercentComplete = '' then ANewTask.PercentComplete := '0';

      Write('DueDate ['''']: ');
      ReadLn(ANewTask.DueDateTime);

      Write('BucketId: ');
      ReadLn(ANewTask.BucketId);
      if ANewtask.BucketId = '' then begin WriteLn('BucketId is required'); exit; end;

      self.FPlanner.CreateTask(ANewTask);

      if not REQUESTERROR then
      begin
        AListing := Tlisting.Create(self.FOptions, self.FPlanner);
        AListing.writeTask(ANewTask);
        WriteLn(Alisting.Text);
        AListing.Free;
      end;
    end
    else
    begin
      WriteLn('You must specify a Bucket or Task to create');
    end;
  end
  else
  begin
    if self.FOptions.ContainsKey('Bucket') then
    begin
      AFields := self.getFields;
      
      if not AFields.TryGetValue('Name', ANewBucket.Name) then begin WriteLn('Name field is required'); exit; end
      else if not AFields.TryGetValue('OrderHint', ANewBucket.OrderHint) then begin end
      else if not AFields.TryGetValue('PlannerId', ANewBucket.PlanId) then begin WriteLn('PlannerId field is required'); exit; end;
      self.FPlanner.CreateBucket(ANewBucket);
      AFields.Free;

      if not REQUESTERROR then
      begin
        AListing := Tlisting.Create(self.FOptions, self.FPlanner);
        AListing.writeBucket(ANewBucket);
        WriteLn(Alisting.Text);
        AListing.Free;
      end;
    end
    else if self.FOptions.ContainsKey('Task') then
    begin
      AFields := self.getFields;
      if not AFields.TryGetValue('Name', ANewTask.Title) then begin WriteLn('Name field is required'); exit; end
      else if not AFields.TryGetValue('PercentComplete', ANewTask.PercentComplete) then begin ANewTask.PercentComplete := '0'; end
      else if not AFields.TryGetValue('DueDate', ANewTask.DueDateTime) then begin end
      else if not AFields.TryGetValue('BucketId', ANewTask.BucketId) then begin WriteLn('BucketId field is required'); exit; end;
      self.FPlanner.CreateTask(ANewTask);
      AFields.Free;

      if not REQUESTERROR then
      begin
        AListing := Tlisting.Create(self.FOptions, self.FPlanner);
        AListing.writeTask(ANewTask);
        WriteLn(Alisting.Text);
        AListing.Free;
      end;
    end
    else
    begin
      WriteLn('You must specify a Bucket or Task to create');
    end;
  end;
end;

procedure THelpers.updateItem;
begin

end;

procedure THelpers.deleteItem;
begin

end;

function THelpers.getFields: TDictionary<string, string>;
var
  s: string;
  arr: TArray<string>;
  arr2: TArray<string>;
begin
  Result := nil;
  if self.FOptions.TryGetValue('Fields', s) then
  begin
    arr := s.Split([',']);
    Result := TDictionary<string, string>.Create;
    for s in arr do
    begin
      arr2 := s.Split(['=']);
      if Length(arr2) = 2 then
        Result.AddOrSetValue(arr2[0], arr2[1])
      else
        Result.AddOrSetValue(arr2[0], '');
    end;
  end;
end;

end.
