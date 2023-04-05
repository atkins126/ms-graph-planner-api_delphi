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
  MicrosoftApiAuthenticator;

type
  TMsPlanner = class(TMsAdapter)
  private
  protected
  public
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

end.