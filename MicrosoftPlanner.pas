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
    FId: string;
  protected
  public
    constructor Create(Authenticator: TMsAuthenticator; id: string); reintroduce;
    destructor Destroy; override;
    property id: string read FId;

    
  end;

implementation

{ TMsPlanner }

constructor TMsPlanner.Create(Authenticator: TMsAuthenticator; id: string);
begin
  inherited Create(Authenticator);
end;

destructor TMsPlanner.Destroy;
begin

  inherited Destroy;
end;

end.