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
    BucketId: string;
    PlanId: string;
    CreatedDateTime: string;
    CompletedDateTime: string;
    PercentComplete: string;
    DueDateTime: string;
    HasDescription: string;
    PreviewType: string;
    ETag: string;
  end;

  TMsPlannerBucket = record
    Id: string;
    Name: string;
    OrderHint: string;
    PlanId: string;
    Tasks: TArray<TMsPlannerTask>;
    ETag: string;
  end;

  TMsPlannerPlanner = record
    Id: string;
    Title: string;
    Owner: string;
    CreatedDateTime: string;
    Buckets: TArray<TMsPlannerBucket>;
    ETag: string;
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

    function FExecuteRequest(AReq: IHttpRequest): IHTTPResponse;

    function GetValue(AJ: TJsonValue; AKey: string; var AValue: string): boolean;
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

    procedure CreateBucket(var Bucket: TMsPlannerBucket);
    procedure CreateTask(var Task: TMsPlannerTask);

    procedure UpdateBucket(var Bucket: TMsPlannerBucket);
    procedure UpdateTask(var Task: TMsPlannerTask);

    procedure DeleteBucket(var Bucket: TMsPlannerBucket);
    procedure DeleteTask(var Task: TMsPlannerTask);
    
    constructor Create(Authenticator: TMsAuthenticator); reintroduce;
    destructor Destroy; override;
  end;

implementation

{ TMsPlanner }

constructor TMsPlanner.Create(Authenticator: TMsAuthenticator);
begin
  inherited Create(Authenticator);
end;

destructor TMsPlanner.Destroy;
begin

  inherited Destroy;
end;

function TMsPlanner.FExecuteRequest(AReq: IHttpRequest): IHTTPResponse;
var
  AStatusCode: integer;
  ARetryAfterStr: string;
  ARetryAfter: integer;
  ARetryCount: integer;
const
  DefaultRetryAfter = 1;
  Max_Retry = 10;

  function GetHeaderValue(AHeaderName: string; AHeaders: TNetHeaders; var AValue: string): boolean;
  var
    AHeader: TNetHeader;
  begin
    AValue := '';
    Result := False;
    for AHeader in AHeaders do
    begin
      if SameText(AHeader.Name, AHeaderName) then
      begin
        AValue := AHeader.Value;
        Result := True;
        Break;
      end;
    end;
  end;
begin
  AStatusCode := 429;
  ARetryAfter := 0;
  ARetryCount := 0;
  while (AStatusCode = 429) and (ARetryCount <= Max_Retry) do
  begin
    inc(ARetryCount);
    Result := self.Http.Execute(AReq);
    AStatusCode := Result.StatusCode;
    if AStatusCode = 429 then
    begin
      self.handleError(AReq, Result);
      if GetHeaderValue('Retry-After', Result.Headers, ARetryAfterStr) then
      if TryStrToInt(ARetryAfterStr, ARetryAfter) then
      begin
        Sleep(ARetryAfter * 1000);
      end
      else
      begin
        Sleep(1000 * DefaultRetryAfter);
      end
      else
      begin
        Sleep(1000 * DefaultRetryAfter);
      end;
    end;
  end;
end;

function TMsPlanner.GetValue(AJ: TJsonValue; AKey: string; var AValue: string): boolean;
var
  AI: Integer;
  AJsonObj: TJSONObject;
begin
  Result := False;
  AValue := '';
  if AKey <> '' then
  begin
    AJsonObj := AJ as TJSONObject;
    for AI := 0 to AJsonObj.Count - 1 do
    begin
      if AJsonObj.Pairs[AI].JsonString.Value = AKey then
      begin
        AValue :=AJsonObj.Pairs[AI].JsonValue.Value;
        Result := True;
        Break;
      end;
    end;
  end;
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

  ARes := self.FExecuteRequest(AReq);

  if ARes.StatusCode = 200 then
  begin
    AJ := TJSONValue.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
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
      AJ.Free;
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
  ARes := self.FExecuteRequest(AReq);

  if ARes.StatusCode = 200 then
  begin
    AJ := TJSONValue.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
    if AJ <> nil then
    begin
      AJ.TryGetValue<string>('id', Group.Id);
      AJ.TryGetValue<string>('displayName', Group.DisplayName);
      AJ.TryGetValue<string>('description', Group.Description);
      AJ.TryGetValue<string>('createdDateTime', Group.CreatedDateTime);
      AJ.Free;
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
  ARes := self.FExecuteRequest(AReq);

  if ARes.StatusCode = 200 then
  begin
    AJ := TJSONValue.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
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
          self.GetValue(AJVal, '@odata.etag', Planner.ETag);
          Group.Planners[i] := Planner;
        end;
      end;
      AJ.Free;
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
  ARes := self.FExecuteRequest(AReq);

  if ARes.StatusCode = 200 then
  begin
    AJ := TJSONValue.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
    if AJ <> nil then
    begin
      AJ.TryGetValue<string>('id', Planner.Id);
      AJ.TryGetValue<string>('title', Planner.Title);
      AJ.TryGetValue<string>('owner', Planner.Owner);
      AJ.TryGetValue<string>('createdDateTime', Planner.CreatedDateTime);
      self.GetValue(AJ, '@odata.etag', Planner.ETag);
      AJ.Free;
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
  ARes := self.FExecuteRequest(AReq);

  if ARes.StatusCode = 200 then
  begin
    AJ := TJSONValue.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
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
          self.GetValue(AJVal, '@odata.etag', Bucket.ETag);
          Planner.Buckets[i] := Bucket;
        end;
      end;
      AJ.Free;
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
  ARes := self.FExecuteRequest(AReq);

  if ARes.StatusCode = 200 then
  begin
    AJ := TJSONValue.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
    if AJ <> nil then
    begin
      AJ.TryGetValue<string>('id', Bucket.Id);
      AJ.TryGetValue<string>('name', Bucket.Name);
      AJ.TryGetValue<string>('orderHint', Bucket.OrderHint);
      AJ.TryGetValue<string>('planId', Bucket.PlanId);
      self.GetValue(AJ, '@odata.etag', Bucket.ETag);
      AJ.Free;
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
  ARes := self.FExecuteRequest(AReq);

  if ARes.StatusCode = 200 then
  begin
    AJ := TJSONValue.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
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
          self.GetValue(AJVal, '@odata.etag', Task.ETag);
          Bucket.Tasks[i] := Task;
        end;
      end;
      AJ.Free;
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
  ARes := self.FExecuteRequest(AReq);

  if ARes.StatusCode = 200 then
  begin
    AJ := TJSONValue.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
    if AJ <> nil then
    begin
      AJ.TryGetValue<string>('id', Task.Id);
      AJ.TryGetValue<string>('title', Task.Title);
      AJ.TryGetValue<string>('orderHint', Task.OrderHint);
      AJ.TryGetValue<string>('bucketId', Task.BucketId);
      AJ.TryGetValue<string>('planId', Task.PlanId);
      AJ.TryGetValue<string>('createdDateTime', Task.CreatedDateTime);
      AJ.TryGetValue<string>('completedDateTime', Task.CompletedDateTime);
      AJ.TryGetValue<string>('percentComplete', Task.PercentComplete);
      AJ.TryGetValue<string>('dueDateTime', Task.DueDateTime);
      AJ.TryGetValue<string>('hasDescription', Task.HasDescription);
      AJ.TryGetValue<string>('previewType', Task.PreviewType);
      self.GetValue(AJ, '@odata.etag', Task.ETag);
      AJ.Free;
    end;
  end
  else
  begin
    self.handleError(AReq, ARes);
  end;
end;

procedure TMsPlanner.CreateTask(var Task: TMsPlannerTask);
var
  bucket: TMsPlannerBucket;

  // requests
  AReq: IHttpRequest;
  ARes: IHTTPResponse;
  APayload: TStringStream;

  // task data
  AJ: TJSONValue;
  AJObj: TJSONObject;
begin
  AJObj := TJSONObject.Create;

  if task.PlanId = '' then
  begin
    bucket.Id := Task.BucketId;
    self.GetBucket(bucket);
    Task.PlanId := bucket.PlanId;
  end;

  AJObj.AddPair('title', Task.Title);
  if Task.OrderHint <> '' then
    AJObj.AddPair('orderHint', Task.OrderHint);
  AJObj.AddPair('bucketId', Task.BucketId);
  AJObj.AddPair('planId', Task.PlanId);
  if Task.PercentComplete <> '' then
    AJObj.AddPair('percentComplete', Task.PercentComplete);
  if Task.DueDateTime <> '' then
    AJObj.AddPair('dueDateTime', Task.DueDateTime);

  AReq := self.Http.GetRequest(sHTTPMethodPost, self.buildUrl('planner/tasks'));
  AReq.AddHeader('Content-Type', 'application/json');
  AReq.AddHeader('Accept', 'application/json');
  AReq.AddHeader('Authorization', self.Token);
  APayload := TStringStream.Create(AJObj.ToJSON);
  AJObj.Free;
  AReq.SourceStream := APayload;
  ARes := self.FExecuteRequest(AReq);
  APayload.Free;

  if ARes.StatusCode = 201 then
  begin
    AJ := TJSONValue.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
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
      self.GetValue(AJ, '@odata.etag', Task.ETag);
      AJ.Free;
    end;
  end
  else
  begin
    self.handleError(AReq, ARes);
  end;
end;

procedure TMsPlanner.UpdateTask(var Task: TMsPlannerTask);
var
  OldTask: TMsPlannerTask;
  // requests
  AReq: IHttpRequest;
  ARes: IHTTPResponse;
  APayload: TStringStream;

  // task data
  AJ: TJSONValue;
  AJObj: TJSONObject;
begin
  if Task.ETag = '' then
  begin
    OldTask.Id := Task.Id;
    self.GetTask(OldTask);
    Task.ETag := OldTask.ETag;
  end;

  AJObj := TJSONObject.Create;
  if Task.Title <> '' then
    AJObj.AddPair('title', Task.Title);
  if Task.OrderHint <> '' then
    AJObj.AddPair('orderHint', Task.OrderHint);
  if Task.BucketId <> '' then
    AJObj.AddPair('bucketId', Task.BucketId);
  if Task.PercentComplete <> '' then
    AJObj.AddPair('percentComplete', Task.PercentComplete);
  if Task.DueDateTime <> '' then
    AJObj.AddPair('dueDateTime', Task.DueDateTime);

  AReq := self.Http.GetRequest(sHTTPMethodPatch, self.buildUrl('planner/tasks/' + Task.Id));
  AReq.AddHeader('Content-Type', 'application/json');
  AReq.AddHeader('Accept', 'application/json');
  AReq.AddHeader('Authorization', self.Token);
  AReq.AddHeader('Prefer', 'return=representation');
  AReq.AddHeader('If-Match', Task.ETag);
  APayload := TStringStream.Create(AJObj.ToJSON);
  AJObj.Free;
  AReq.SourceStream := APayload;
  ARes := self.FExecuteRequest(AReq);
  APayload.Free;

  if ARes.StatusCode = 200 then
  begin
    AJ := TJSONValue.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
    if AJ <> nil then
    begin
      AJ.TryGetValue<string>('id', Task.Id);
      AJ.TryGetValue<string>('title', Task.Title);
      AJ.TryGetValue<string>('orderHint', Task.OrderHint);
      AJ.TryGetValue<string>('bucketId', Task.BucketId);
      AJ.TryGetValue<string>('planId', Task.PlanId);
      AJ.TryGetValue<string>('createdDateTime', Task.CreatedDateTime);
      AJ.TryGetValue<string>('completedDateTime', Task.CompletedDateTime);
      AJ.TryGetValue<string>('percentComplete', Task.PercentComplete);
      AJ.TryGetValue<string>('dueDateTime', Task.DueDateTime);
      AJ.TryGetValue<string>('hasDescription', Task.HasDescription);
      AJ.TryGetValue<string>('previewType', Task.PreviewType);
      self.GetValue(AJ, '@odata.etag', Task.ETag);
      AJ.Free;
    end;
  end
  else
  begin
    self.handleError(AReq, ARes);
  end;
end;

procedure TMsPlanner.DeleteTask(var Task: TMsPlannerTask);
var
  // requests
  AReq: IHttpRequest;
  ARes: IHTTPResponse;
begin
  if Task.ETag = '' then
  begin
    self.GetTask(Task);
  end;

  AReq := self.Http.GetRequest(sHTTPMethodDelete, self.buildUrl('planner/tasks/' + Task.Id));
  AReq.AddHeader('Content-Type', 'application/json');
  AReq.AddHeader('Accept', 'application/json');
  AReq.AddHeader('Authorization', self.Token);
  AReq.AddHeader('If-Match', Task.ETag);
  ARes := self.FExecuteRequest(AReq);

  if ARes.StatusCode = 204 then
  begin
    Task.Id := '';
  end
  else
  begin
    self.handleError(AReq, ARes);
  end;
end;

procedure TMsPlanner.CreateBucket(var Bucket: TMsPlannerBucket);
var
  // requests
  AReq: IHttpRequest;
  ARes: IHTTPResponse;
  APayload: TStringStream;

  // bucket data
  AJ: TJSONValue;
  AJObj: TJSONObject;
begin
  AJObj := TJSONObject.Create;
  AJObj.AddPair('name', Bucket.Name);
  AJObj.AddPair('planId', Bucket.PlanId);
  if Bucket.OrderHint <> '' then
    AJObj.AddPair('orderHint', Bucket.OrderHint);

  AReq := self.Http.GetRequest(sHTTPMethodPost, self.buildUrl('planner/buckets'));
  AReq.AddHeader('Content-Type', 'application/json');
  AReq.AddHeader('Accept', 'application/json');
  AReq.AddHeader('Authorization', self.Token);
  APayload := TStringStream.Create(AJObj.ToJSON);
  AJObj.Free;
  AReq.SourceStream := APayload;
  ARes := self.FExecuteRequest(AReq);
  APayload.Free;

  if ARes.StatusCode = 201 then
  begin
    AJ := TJSONValue.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
    if AJ <> nil then
    begin
      AJ.TryGetValue<string>('id', Bucket.Id);
      AJ.TryGetValue<string>('name', Bucket.Name);
      AJ.TryGetValue<string>('planId', Bucket.PlanId);
      AJ.TryGetValue<string>('orderHint', Bucket.OrderHint);
      self.GetValue(AJ, '@odata.etag', Bucket.ETag);
      AJ.Free;
    end;
  end
  else
  begin
    self.handleError(AReq, ARes);
  end;
end;

procedure TMsPlanner.UpdateBucket(var Bucket: TMsPlannerBucket);
var
  OldBucket: TMsPlannerBucket;
  // requests
  AReq: IHttpRequest;
  ARes: IHTTPResponse;
  APayload: TStringStream;

  // bucket data
  AJ: TJSONValue;
  AJObj: TJSONObject;
begin
  if Bucket.ETag = '' then
  begin
    OldBucket.id := Bucket.Id;
    self.GetBucket(OldBucket);
    Bucket.ETag := OldBucket.ETag;
  end;

  AJObj := TJSONObject.Create;
  if Bucket.Name <> '' then
    AJObj.AddPair('name', Bucket.Name);
  if Bucket.OrderHint <> '' then
    AJObj.AddPair('orderHint', Bucket.OrderHint);

  AReq := self.Http.GetRequest(sHTTPMethodPatch, self.buildUrl('planner/buckets/' + Bucket.Id));
  AReq.AddHeader('Content-Type', 'application/json');
  AReq.AddHeader('Accept', 'application/json');
  AReq.AddHeader('Authorization', self.Token);
  AReq.AddHeader('If-Match', Bucket.ETag);
  AReq.AddHeader('Prefer', 'return=representation');
  APayload := TStringStream.Create(AJObj.ToJSON);
  AJObj.Free;
  AReq.SourceStream := APayload;
  ARes := self.FExecuteRequest(AReq);
  APayload.Free;

  if ARes.StatusCode = 200 then
  begin
    AJ := TJSONValue.ParseJSONValue(ARes.ContentAsString(TEncoding.UTF8));
    if AJ <> nil then
    begin
      AJ.TryGetValue<string>('id', Bucket.Id);
      AJ.TryGetValue<string>('name', Bucket.Name);
      AJ.TryGetValue<string>('planId', Bucket.PlanId);
      AJ.TryGetValue<string>('orderHint', Bucket.OrderHint);
      self.GetValue(AJ, '@odata.etag', Bucket.ETag);
      AJ.Free;
    end;
  end
  else
  begin
    self.handleError(AReq, ARes);
  end;
end;

procedure TMsPlanner.DeleteBucket(var Bucket: TMsPlannerBucket);
var
  // requests
  AReq: IHttpRequest;
  ARes: IHTTPResponse;
begin
  if Bucket.ETag = '' then
  begin
    self.GetBucket(Bucket);
  end;

  AReq := self.Http.GetRequest(sHTTPMethodDelete, self.buildUrl('planner/buckets/' + Bucket.Id));
  AReq.AddHeader('Content-Type', 'application/json');
  AReq.AddHeader('Accept', 'application/json');
  AReq.AddHeader('Authorization', self.Token);
  AReq.AddHeader('If-Match', Bucket.ETag);
  ARes := self.FExecuteRequest(AReq);

  if ARes.StatusCode = 204 then
  begin
    Bucket.Id := '';
  end
  else
  begin
    self.handleError(AReq, ARes);
  end;
end;

end.