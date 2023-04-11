unit listing;

interface

uses
  System.Generics.Collections,
  MicrosoftPlanner;

type
  Tlisting = class
  private
    FOptions: TDictionary<string, string>;
    FPlannerLib: TMsPlanner;

    FIndentation: Integer;
    FMsg: TArray<string>;
    procedure msg(s: string);

  
    function doWithId: Boolean;
    

    function getText: string;
  public
    constructor Create(Options: TDictionary<string, string>; planner: TMsPlanner);
    destructor Destroy; override;

    procedure doListing();

    property Text: string read getText;

    procedure writePlanner(p: TMsPlannerPlanner);
    procedure writeBucket(b: TMsPlannerBucket);
    procedure writeTask(t: TMsPlannerTask);
    procedure writeGroup(g: TMsPlannerGroup);
    property indentation: Integer read FIndentation write FIndentation;

  end;

implementation

uses
  System.SysUtils;

{ Tlisting }

constructor Tlisting.Create(Options: TDictionary<string, string>; planner: TMsPlanner);
begin
  inherited Create;
  FOptions := Options;
  FPlannerLib := planner;
end;

function Tlisting.doWithId: Boolean;
var
  AGroup: TMsPlannerGroup;
  AIPlanner: Integer;
  APlanner: TMsPlannerPlanner;
  AIBucket: Integer;
  ABucket: TMsPlannerBucket;
  ATask: TMsPlannerTask;
begin
  result := false;

  if self.FOptions.TryGetValue('Task', ATask.id) then
  begin
    if ATask.Id <> '' then
    begin
      self.FPlannerLib.GetTask(ATask);
      self.writeTask(ATask);
      exit(true);
    end;
  end;

  if self.FOptions.TryGetValue('Bucket', ABucket.Id) then
  begin
    if ABucket.Id <> '' then
    begin
      self.FPlannerLib.GetBucket(ABucket);
      self.writeBucket(ABucket);

      if self.FOptions.ContainsKey('Task') then
      begin
        inc(self.FIndentation);
        self.FPlannerLib.GetTasks(ABucket);
        for ATask in ABucket.Tasks do self.writeTask(ATask);
        dec(self.FIndentation);
      end;
      exit(true);
    end;
  end;

  if self.FOptions.TryGetValue('Planner', APlanner.Id) then
  begin
    if APlanner.Id <> '' then
    begin
      Self.FPlannerLib.GetPlanner(APlanner);
      self.writePlanner(APlanner);

      if self.FOptions.ContainsKey('Bucket') or self.FOptions.ContainsKey('Task') then
      begin
        inc(self.FIndentation);
        self.FPlannerLib.GetBuckets(APlanner);
        for AIBucket := 0 to Length(APlanner.Buckets) -1 do
        begin
          ABucket := APlanner.Buckets[AIBucket];
          self.writeBucket(ABucket);
          if self.FOptions.ContainsKey('Task') then
          begin
            inc(self.FIndentation);
            self.FPlannerLib.GetTasks(ABucket);
            for ATask in ABucket.Tasks do self.writeTask(ATask);
            dec(self.FIndentation);
          end;
        end;
        dec(self.FIndentation);
      end;
      exit(true);
    end;
  end;

  if self.FOptions.TryGetValue('Group', AGroup.Id) then
  begin
    if AGroup.Id <> '' then
    begin
      self.FPlannerLib.GetGroup(AGroup);
      self.writeGroup(AGroup);

      if self.FOptions.ContainsKey('Planner') or self.FOptions.ContainsKey('Bucket') or self.FOptions.ContainsKey('Task') then
      begin
        inc(self.FIndentation);
        self.FPlannerLib.GetPlanners(AGroup);
        for AIPlanner := 0 to Length(AGroup.Planners) -1 do
        begin
          APlanner := AGroup.Planners[AIPlanner];
          self.writePlanner(APlanner);
          if self.FOptions.ContainsKey('Bucket') or self.FOptions.ContainsKey('Task') then
          begin
            inc(self.FIndentation);
            self.FPlannerLib.GetBuckets(APlanner);
            for AIBucket := 0 to Length(APlanner.Buckets) -1 do
            begin
              ABucket := APlanner.Buckets[AIBucket];
              self.writeBucket(ABucket);
              if self.FOptions.ContainsKey('Task') then
              begin
                inc(self.FIndentation);
                self.FPlannerLib.GetTasks(ABucket);
                for ATask in ABucket.Tasks do self.writeTask(ATask);
                dec(self.FIndentation);
              end;
            end;
            dec(self.FIndentation);
          end;
        end;
        dec(self.FIndentation);
      end;
      exit(true);
    end;
  end;
end;

destructor Tlisting.Destroy;
begin
  inherited;
end;

procedure Tlisting.doListing;
var
  AGroups: TArray<TMsPlannerGroup>;
  AGroup: TMsPlannerGroup;
  AIGroup: Integer;
  APlan: TMsPlannerPlanner;
  AIPlan: Integer;
  ABucket: TMsPlannerBucket;
  AIBucket: Integer;
  ATask: TMsPlannerTask;
 // AITask: Integer;
begin
  if self.doWithId then exit;

  AGroups := self.FPlannerLib.GetGroups;
  
  if self.FOptions.ContainsKey('Task') then
  begin
    for AIGroup := 0 to Length(AGroups) -1 do
    begin
      AGroup := AGroups[AIGroup];
      self.writeGroup(AGroup);
      inc(self.FIndentation);
      self.FPlannerLib.GetPlanners(AGroup);
      AGroups[AIGroup] := AGroup;
      for AIPlan := 0 to Length(AGroup.Planners) -1 do
      begin
        APlan := AGroup.Planners[AIPlan];
        self.writePlanner(APlan);
        inc(self.FIndentation);
        self.FPlannerLib.GetBuckets(APlan);
        AGroup.Planners[AIPlan] := APlan;
        for AIBucket := 0 to Length(APlan.Buckets) -1 do
        begin
          ABucket := APlan.Buckets[AIBucket];
          self.writeBucket(ABucket);
          inc(self.FIndentation);
          self.FPlannerLib.GetTasks(ABucket);
          APlan.Buckets[AIBucket] := ABucket;
          for ATask in ABucket.Tasks do self.writeTask(ATask);
          dec(self.FIndentation);
        end;
        dec(self.FIndentation);
      end;
      dec(self.FIndentation);
    end;
  end
  else if self.FOptions.ContainsKey('Bucket') then
  begin
    for AIGroup := 0 to Length(AGroups) -1 do
    begin
      AGroup := AGroups[AIGroup];
      self.writeGroup(AGroup);
      inc(self.FIndentation);
      self.FPlannerLib.GetPlanners(AGroup);
      AGroups[AIGroup] := AGroup;
      for AIPlan := 0 to Length(AGroup.Planners) -1 do
      begin
        APlan := AGroup.Planners[AIPlan];
        self.writePlanner(APlan);
        inc(self.FIndentation);
        self.FPlannerLib.GetBuckets(APlan);
        AGroup.Planners[AIPlan] := APlan;
        for ABucket in APlan.Buckets do
        begin
          self.writeBucket(ABucket);
        end;
        dec(self.FIndentation);
      end;
      dec(self.FIndentation);
    end;
  end
  else if self.FOptions.ContainsKey('Planner') then
  begin
    for AIGroup := 0 to Length(AGroups) -1 do
    begin
      AGroup := AGroups[AIGroup];
      self.writeGroup(AGroup);
      inc(self.FIndentation);
      self.FPlannerLib.GetPlanners(AGroup);
      AGroups[AIGroup] := AGroup;
      for APlan in AGroup.Planners do
      begin
        self.writePlanner(APlan);
      end;
      dec(self.FIndentation);
    end;
  end
  else if self.FOptions.ContainsKey('Group') then
  begin
    for AIGroup := 0 to Length(AGroups) -1 do
    begin
      AGroup := AGroups[AIGroup];
      self.writeGroup(AGroup);
    end;
  end
  else
  begin
    for AIGroup := 0 to Length(AGroups) -1 do
    begin
      AGroup := AGroups[AIGroup];
      self.writeGroup(AGroup);
      inc(self.FIndentation);
      self.FPlannerLib.GetPlanners(AGroup);
      AGroups[AIGroup] := AGroup;
      for APlan in AGroup.Planners do
      begin
        self.writePlanner(APlan);
      end;
      dec(self.FIndentation);
    end;
  end;
end;

procedure Tlisting.msg(s: string);
begin
  FMsg := FMsG + [string.Create(' ', self.FIndentation*2) + s];
end;

procedure Tlisting.writePlanner(p: TMsPlannerPlanner);
begin
  msg('- Planner: ' + p.Title);
  msg('  Id: ' + p.Id);
  msg('  Created: ' + p.CreatedDateTime);
  msg('  Owner: ' + p.Owner);
end;

procedure Tlisting.writeBucket(b: TMsPlannerBucket);
begin
  self.msg('- Bucket: ' + b.Name);
  self.msg('  Id: ' + b.Id);
  self.msg('  Order Hint: ' + b.OrderHint);
  self.msg('  Plan Id: ' + b.PlanId);
end;

procedure Tlisting.writeTask(t: TMsPlannerTask);
begin
  self.msg('- Task: ' + t.Title);
  self.msg('  Id: ' + t.Id);
  self.msg('  Created: ' + t.CreatedDateTime);
  self.msg('  Due: ' + t.DueDateTime);
  self.msg('  Completed: ' + t.CompletedDateTime);
  self.msg('  Percent Complete: ' + t.PercentComplete);
  self.msg('  Order Hint: ' + t.OrderHint);
  self.msg('  Plan Id: ' + t.PlanId);
  self.msg('  Has Description: ' + t.HasDescription);
  self.msg('  Preview Type: ' + t.PreviewType);
end;

procedure Tlisting.writeGroup(g: TMsPlannerGroup);
begin
  self.msg('- Group: ' + g.DisplayName);
  self.msg('  Id: ' + g.Id);
  self.msg('  Description: ' + g.Description);
  self.msg('  Created: ' + g.CreatedDateTime);
end;

function Tlisting.GetText: string;
begin
  result := string.Join(sLineBreak, self.FMsg);
end;

end.