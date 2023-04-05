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
  MicrosoftApiAuthenticator;

type
  THelperPlanner = record
    Id: string;
    Title: string;
    Owner: string;
    CreatedDateTime: string;
  end;

  THelperGroup = record
    Id: string;
    DisplayName: string;
    Description: string;
    CreatedDateTime: string;
    Planners: TArray<THelperPlanner>;
  end;

  THelpers = class(TMsAdapter)
  private
    Fauthenticator: TMsAuthenticator;
    FVerbose: boolean;
    function buildUrl(s: string): string;
    procedure HandleError(AReq: IHttpRequest; ARes: IHTTPResponse);
  protected
  public
    constructor Create(Authenticator: TMsAuthenticator); reintroduce;
    destructor Destroy; override;

    function getAllPlanners: TArray<THelperGroup>;

    class function New(TENANT_ID:string; CLINET_ID: string; REDIRECT_URI: string; REDIRECT_PORT: integer; SCOPE: TArray<string>; Verbose: boolean): THelpers; static;
  end;

implementation

{ THelpers }

constructor THelpers.Create(Authenticator: TMsAuthenticator);
begin
  inherited Create(Authenticator);
  self.Fauthenticator := nil;
end;

destructor THelpers.Destroy;
var
  AAuthenticator: TMsAuthenticator;
begin
  AAuthenticator := self.Fauthenticator;
  inherited;
  if AAuthenticator <> nil then
    AAuthenticator.Free;
end;

function THelpers.buildUrl(s: string): string;
begin
  Result := 'https://graph.microsoft.com/v1.0/' + s;
end;

procedure THelpers.HandleError(AReq: IHttpRequest; ARes: IHTTPResponse);
var
  AErr: TMsError;
  AJ: TJSONValue;
  AJErr: TJSONValue;
begin
  AErr.HTTPStatusCode := ARes.StatusCode;
  AErr.HTTPStatusText := ARes.StatusText;
  AErr.HTTPurl := AReq.URL.ToString;
  AErr.HTTPMethod := AReq.MethodString;
  AErr.HTTPreq_Header := AReq.Headers;
  AErr.HTTPres_header := ARes.Headers;
  AErr.HTTPerror_data := ARes.ContentAsString;

  Aj := TJSONValue.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
  if Aj <> nil then
  begin
    if Aj.TryGetValue<TJsonValue>('error', AJErr) then
    begin
      AJErr.TryGetValue<string>('code', AErr.HTTPerror_name);
      AJErr.TryGetValue<string>('message', AErr.HTTPerror_description);
    end;
    Aj.Free;
  end;
  self.OnRequestError(AErr);
end;

function THelpers.getAllPlanners: TArray<THelperGroup>;
var
  // requests
  AReq: IHttpRequest;
  ARes: IHTTPResponse;

  // group data
  AJson: TJSONValue;
  AJsonGroups: TJSONArray;
  AIGroup: integer;
  AGroupId: string;

  // planner data
  AJsonPlannerRes: TJSONValue;
  AJsonPlanners: TJSONArray;
  AJsonPlanner: TJsonValue;

  // result
  ANewPlanner: THelperPlanner;
  ANewGroup: THelperGroup;
begin
  AReq := self.Http.GetRequest(sHTTPMethodGet, self.buildUrl('groups'));
  AReq.AddHeader('Content-Type', 'application/json');
  AReq.AddHeader('Accept', 'application/json');
  AReq.AddHeader('Authorization', self.Token);
  ARes := self.Http.Execute(AReq);

  if ARes.StatusCode = 200 then
  begin
    AJson := TJSONObject.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
    if AJson <> nil then
    begin
      if AJson.TryGetValue<TJSONArray>('value', AJsonGroups) then
      begin
        // iterate through all groups
        for AIGroup := 0 to AJsonGroups.Count - 1 do
        begin
          // print status
          //if self.FVerbose then
          //  WriteLn('Fetch Group', AIGroup, ' of ', AJsonGroups.Count, '...', #13); // @todo this doesn't work. why?
          // fetch planners
          if AJsonGroups.Items[AIGroup].TryGetValue<string>('id', AGroupId) then
          begin
            // parse group
            ANewGroup.Id := AGroupId;
            AJsonGroups.Items[AIGroup].TryGetValue<string>('displayName', ANewGroup.DisplayName);
            AJsonGroups.Items[AIGroup].TryGetValue<string>('description', ANewGroup.Description);
            AJsonGroups.Items[AIGroup].TryGetValue<string>('createdDateTime', ANewGroup.CreatedDateTime);
            ANewGroup.Planners := [];
            // fetch planners
            AReq := self.Http.GetRequest(sHTTPMethodGet, self.buildUrl('groups/' + AGroupId + '/planner/plans'));
            AReq.AddHeader('Content-Type', 'application/json');
            AReq.AddHeader('Accept', 'application/json');
            AReq.AddHeader('Authorization', self.Token);
            ARes := self.Http.Execute(AReq);
            if ARes.StatusCode = 200 then
            begin
              // fetch tasks
              AJsonPlannerRes := TJSONObject.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
              if AJsonPlannerRes <> nil then
              begin
                if AJsonPlannerRes.TryGetValue<TJSONArray>('value', AJsonPlanners) then
                begin
                  // iterate through all planners
                  for AJsonPlanner in AJsonPlanners do
                  begin
                    // parse planner
                    ANewPlanner.Id := AJsonPlanner.GetValue<string>('id');
                    AJsonPlanner.TryGetValue<string>('title', ANewPlanner.Title);
                    AJsonPlanner.TryGetValue<string>('owner', ANewPlanner.Owner);
                    AJsonPlanner.TryGetValue<string>('createdDateTime', ANewPlanner.CreatedDateTime);
                    // add planner to group
                    ANewGroup.Planners := ANewGroup.Planners + [ANewPlanner];
                  end;
                end;
                AJsonPlannerRes.Free;
              end;
              // add group to result
              if Length(ANewGroup.Planners) > 0 then
                Result := Result + [ANewGroup];
            end
            else
            begin
              self.HandleError(AReq, ARes);
            end;
          end;
        end;
      end;
      AJson.Free;
    end;
  end
  else
  begin
    self.HandleError(AReq, ARes);
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
