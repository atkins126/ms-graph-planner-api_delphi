unit MicrosoftPlanner;

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
  MicrosoftApiAuthenticator;

type
  TMsPlannerTask = record
    Id: string;
    Title: string;
    OrderHint: string;
    PlanId: string;
    CreatedDateTime: string;
    CompletedDateTime: string;
    PercentComplete: string;
    DueDateTime: string;
    HasDescription: string;
    PreviewType: string;
  end;

  TMsPlannerBucket = record
    Id: string;
    Name: string;
    OrderHint: string;
    PlanId: string;
    Tasks: TArray<TMsPlannerTask>;
  end;

  TMsPlannerPlanner = record
    Id: string;
    Title: string;
    Owner: string;
    CreatedDateTime: string;
    Buckets: TArray<TMsPlannerBucket>;
  end;

  TMsPlannerGroup = record
    Id: string;
    DisplayName: string;
    Description: string;
    CreatedDateTime: string;
    Planners: TArray<TMsPlannerPlanner>;
  end;

  TMsPlanner = class(TMsAdapter)
  private
    procedure handleError(AReq: IHttpRequest; ARes: IHTTPResponse);
    function buildUrl(s: string): string;
  protected
  public
    function GetGroups: TArray<TMsPlannerGroup>;
    procedure GetGroup(var Group: TMsPlannerGroup);
    procedure GetPlanners(var Group: TMsPlannerGroup);
    procedure GetPlanner(var Planner: TMsPlannerPlanner);
    procedure GetBuckets(var Planner: TMsPlannerPlanner);
    procedure GetBucket(var Bucket: TMsPlannerBucket);
    procedure GetTasks(var Bucket: TMsPlannerBucket);
    procedure GetTask(var Task: TMsPlannerTask);
    
    constructor Create(Authenticator: TMsAuthenticator); reintroduce;
    destructor Destroy; override;
    property id: string read FId;

    
    property id: string read FId;

    
  end;

implementation

{ TMsPlanner }

constructor TMsPlanner.Create(Authenticator: TMsAuthenticator; id: string);
constructor TMsPlanner.Create(Authenticator: TMsAuthenticator; id: string);
begin
  inherited Create(Authenticator);
end;

destructor TMsPlanner.Destroy;
begin

  inherited Destroy;
end;

procedure TMsPlanner.handleError(AReq: IHttpRequest; ARes: IHTTPResponse);
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

function TMsPlanner.buildUrl(s: string): string;
begin
  Result := 'https://graph.microsoft.com/v1.0/' + s;
end;

function TMsPlanner.GetGroups: TArray<TMsPlannerGroup>;
var
  // requests
  AReq: IHttpRequest;
  ARes: IHTTPResponse;

  // group data
  AJ: TJSONValue;
  AJArr: TJSONArray;
  AJVal: TJSONValue;
  i: integer;
  Group: TMsPlannerGroup;
begin
  AReq := self.Http.GetRequest(sHTTPMethodGet, self.buildUrl('me/transitiveMemberOf/microsoft.graph.group?$count=true&$select=displayName,id,createdDateTime,description'));
  AReq.AddHeader('Content-Type', 'application/json');
  AReq.AddHeader('Accept', 'application/json');
  AReq.AddHeader('Authorization', self.Token);
  ARes := self.Http.Execute(AReq);

  if ARes.StatusCode = 200 then
  begin
    AJ := TJSONObject.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
    if AJ <> nil then
    begin
      if AJ.TryGetValue<TJSONArray>('value', AJArr) then
      begin
        SetLength(Result, AJArr.Count);
        for i := 0 to AJArr.Count - 1 do
        begin
          AJVal := AJArr.Items[i];
          AJVal.TryGetValue<string>('id', Group.Id);
          AJVal.TryGetValue<string>('displayName', Group.DisplayName);
          AJVal.TryGetValue<string>('description', Group.Description);
          AJVal.TryGetValue<string>('createdDateTime', Group.CreatedDateTime);
          Result[i] := Group;
        end;
      end;
    end;
  end
  else
  begin
    self.handleError(AReq, ARes);
  end;
end;

procedure TMsPlanner.GetGroup(var Group: TMsPlannerGroup);
var
  // requests
  AReq: IHttpRequest;
  ARes: IHTTPResponse;

  // group data
  AJ: TJSONValue;
begin
  AReq := self.Http.GetRequest(sHTTPMethodGet, self.buildUrl('groups/' + Group.Id + '?$count=true&$select=id,displayName,description,createdDateTime'));
  AReq.AddHeader('Content-Type', 'application/json');
  AReq.AddHeader('Accept', 'application/json');
  AReq.AddHeader('Authorization', self.Token);
  ARes := self.Http.Execute(AReq);

  if ARes.StatusCode = 200 then
  begin
    AJ := TJSONObject.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
    if AJ <> nil then
    begin
      AJ.TryGetValue<string>('id', Group.Id);
      AJ.TryGetValue<string>('displayName', Group.DisplayName);
      AJ.TryGetValue<string>('description', Group.Description);
      AJ.TryGetValue<string>('createdDateTime', Group.CreatedDateTime);
    end;
  end
  else
  begin
    self.handleError(AReq, ARes);
  end;
end;

procedure TMsPlanner.GetPlanners(var Group: TMsPlannerGroup);
var
  // requests
  AReq: IHttpRequest;
  ARes: IHTTPResponse;

  // planner data
  AJ: TJSONValue;
  AJArr: TJSONArray;
  AJVal: TJSONValue;
  i: integer;
  Planner: TMsPlannerPlanner;
begin
  AReq := self.Http.GetRequest(sHTTPMethodGet, self.buildUrl('groups/' + Group.Id + '/planner/plans?$count=true&$select=id,title,owner,createdDateTime'));
  AReq.AddHeader('Content-Type', 'application/json');
  AReq.AddHeader('Accept', 'application/json');
  AReq.AddHeader('Authorization', self.Token);
  ARes := self.Http.Execute(AReq);

  if ARes.StatusCode = 200 then
  begin
    AJ := TJSONObject.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
    if AJ <> nil then
    begin
      if AJ.TryGetValue<TJSONArray>('value', AJArr) then
      begin
        SetLength(Group.Planners, AJArr.Count);
        for i := 0 to AJArr.Count - 1 do
        begin
          AJVal := AJArr.Items[i];
          AJVal.TryGetValue<string>('id', Planner.Id);
          AJVal.TryGetValue<string>('title', Planner.Title);
          AJVal.TryGetValue<string>('owner', Planner.Owner);
          AJVal.TryGetValue<string>('createdDateTime', Planner.CreatedDateTime);
          Group.Planners[i] := Planner;
        end;
      end;
    end;
  end
  else
  begin
    self.handleError(AReq, ARes);
  end;
end;

procedure TMsPlanner.GetPlanner(var Planner: TMsPlannerPlanner);
var
  // requests
  AReq: IHttpRequest;
  ARes: IHTTPResponse;

  // planner data
  AJ: TJSONValue;
begin
  AReq := self.Http.GetRequest(sHTTPMethodGet, self.buildUrl('planner/plans/' + Planner.Id + '?$count=true&$select=id,title,owner,createdDateTime'));
  AReq.AddHeader('Content-Type', 'application/json');
  AReq.AddHeader('Accept', 'application/json');
  AReq.AddHeader('Authorization', self.Token);
  ARes := self.Http.Execute(AReq);

  if ARes.StatusCode = 200 then
  begin
    AJ := TJSONObject.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
    if AJ <> nil then
    begin
      AJ.TryGetValue<string>('id', Planner.Id);
      AJ.TryGetValue<string>('title', Planner.Title);
      AJ.TryGetValue<string>('owner', Planner.Owner);
      AJ.TryGetValue<string>('createdDateTime', Planner.CreatedDateTime);
    end;
  end
  else
  begin
    self.handleError(AReq, ARes);
  end;
end;

procedure TMsPlanner.GetBuckets(var Planner: TMsPlannerPlanner);
var
  // requests
  AReq: IHttpRequest;
  ARes: IHTTPResponse;

  // bucket data
  AJ: TJSONValue;
  AJArr: TJSONArray;
  AJVal: TJSONValue;
  i: integer;
  Bucket: TMsPlannerBucket;
begin
  AReq := self.Http.GetRequest(sHTTPMethodGet, self.buildUrl('planner/plans/' + Planner.Id + '/buckets?$count=true&$select=id,name,orderHint,planId'));
  AReq.AddHeader('Content-Type', 'application/json');
  AReq.AddHeader('Accept', 'application/json');
  AReq.AddHeader('Authorization', self.Token);
  ARes := self.Http.Execute(AReq);

  if ARes.StatusCode = 200 then
  begin
    AJ := TJSONObject.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
    if AJ <> nil then
    begin
      if AJ.TryGetValue<TJSONArray>('value', AJArr) then
      begin
        SetLength(Planner.Buckets, AJArr.Count);
        for i := 0 to AJArr.Count - 1 do
        begin
          AJVal := AJArr.Items[i];
          AJVal.TryGetValue<string>('id', Bucket.Id);
          AJVal.TryGetValue<string>('name', Bucket.Name);
          AJVal.TryGetValue<string>('orderHint', Bucket.OrderHint);
          AJVal.TryGetValue<string>('planId', Bucket.PlanId);
          Planner.Buckets[i] := Bucket;
        end;
      end;
    end;
  end
  else
  begin
    self.handleError(AReq, ARes);
  end;
end;

procedure TMsPlanner.GetBucket(var Bucket: TMsPlannerBucket);
var
  // requests
  AReq: IHttpRequest;
  ARes: IHTTPResponse;

  // bucket data
  AJ: TJSONValue;
begin
  AReq := self.Http.GetRequest(sHTTPMethodGet, self.buildUrl('planner/buckets/' + Bucket.Id + '?$count=true&$select=id,name,orderHint,planId'));
  AReq.AddHeader('Content-Type', 'application/json');
  AReq.AddHeader('Accept', 'application/json');
  AReq.AddHeader('Authorization', self.Token);
  ARes := self.Http.Execute(AReq);

  if ARes.StatusCode = 200 then
  begin
    AJ := TJSONObject.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
    if AJ <> nil then
    begin
      AJ.TryGetValue<string>('id', Bucket.Id);
      AJ.TryGetValue<string>('name', Bucket.Name);
      AJ.TryGetValue<string>('orderHint', Bucket.OrderHint);
      AJ.TryGetValue<string>('planId', Bucket.PlanId);
    end;
  end
  else
  begin
    self.handleError(AReq, ARes);
  end;
end;

procedure TMsPlanner.GetTasks(var Bucket: TMsPlannerBucket);
var
  // requests
  AReq: IHttpRequest;
  ARes: IHTTPResponse;

  // task data
  AJ: TJSONValue;
  AJArr: TJSONArray;
  AJVal: TJSONValue;
  i: integer;
  Task: TMsPlannerTask;
begin
  AReq := self.Http.GetRequest(sHTTPMethodGet, self.buildUrl('planner/buckets/' + Bucket.Id + '/tasks?$count=true&$select=id,title,orderHint,planId,createdDateTime,completedDateTime,percentComplete,dueDateTime,hasDescription,previewType'));
  AReq.AddHeader('Content-Type', 'application/json');
  AReq.AddHeader('Accept', 'application/json');
  AReq.AddHeader('Authorization', self.Token);
  ARes := self.Http.Execute(AReq);

  if ARes.StatusCode = 200 then
  begin
    AJ := TJSONObject.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
    if AJ <> nil then
    begin
      if AJ.TryGetValue<TJSONArray>('value', AJArr) then
      begin
        SetLength(Bucket.Tasks, AJArr.Count);
        for i := 0 to AJArr.Count - 1 do
        begin
          AJVal := AJArr.Items[i];
          AJVal.TryGetValue<string>('id', Task.Id);
          AJVal.TryGetValue<string>('title', Task.Title);
          AJVal.TryGetValue<string>('orderHint', Task.OrderHint);
          AJVal.TryGetValue<string>('planId', Task.PlanId);
          AJVal.TryGetValue<string>('createdDateTime', Task.CreatedDateTime);
          AJVal.TryGetValue<string>('completedDateTime', Task.CompletedDateTime);
          AJVal.TryGetValue<string>('percentComplete', Task.PercentComplete);
          AJVal.TryGetValue<string>('dueDateTime', Task.DueDateTime);
          AJVal.TryGetValue<string>('hasDescription', Task.HasDescription);
          AJVal.TryGetValue<string>('previewType', Task.PreviewType);
          Bucket.Tasks[i] := Task;
        end;
      end;
    end;
  end
  else
  begin
    self.handleError(AReq, ARes);
  end;
end;

procedure TMsPlanner.GetTask(var Task: TMsPlannerTask);
var
  // requests
  AReq: IHttpRequest;
  ARes: IHTTPResponse;

  // task data
  AJ: TJSONValue;
begin
  AReq := self.Http.GetRequest(sHTTPMethodGet, self.buildUrl('planner/tasks/' + Task.Id));
  AReq.AddHeader('Content-Type', 'application/json');
  AReq.AddHeader('Accept', 'application/json');
  AReq.AddHeader('Authorization', self.Token);
  ARes := self.Http.Execute(AReq);

  if ARes.StatusCode = 200 then
  begin
    AJ := TJSONObject.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
    if AJ <> nil then
    begin
      AJ.TryGetValue<string>('id', Task.Id);
      AJ.TryGetValue<string>('title', Task.Title);
      AJ.TryGetValue<string>('orderHint', Task.OrderHint);
      AJ.TryGetValue<string>('planId', Task.PlanId);
      AJ.TryGetValue<string>('createdDateTime', Task.CreatedDateTime);
      AJ.TryGetValue<string>('completedDateTime', Task.CompletedDateTime);
      AJ.TryGetValue<string>('percentComplete', Task.PercentComplete);
      AJ.TryGetValue<string>('dueDateTime', Task.DueDateTime);
      AJ.TryGetValue<string>('hasDescription', Task.HasDescription);
      AJ.TryGetValue<string>('previewType', Task.PreviewType);
    end;
  end
  else
  begin
    self.handleError(AReq, ARes);
  end;
end;


end.