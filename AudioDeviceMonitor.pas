unit AudioDeviceMonitor;

interface

uses
  System.Classes, Winapi.Windows, Winapi.ActiveX, System.Win.ComObj, Winapi.MMSystem;

const
  CLSID_MMDeviceEnumerator: TGUID = '{BCDE0395-E52F-467C-8E3D-C4579291692E}';
  IID_IMMDeviceEnumerator: TGUID = '{A95664D2-9614-4F35-A746-DE8DB63617E6}';
  IID_IMMNotificationClient: TGUID = '{7991EEC9-7E89-4D85-8390-6C703CEC60C0}';

type
  EDataFlow = (
    eRender,
    eCapture,
    eAll,
    EDataFlow_enum_count
  );

  ERole = (
    eConsole,
    eMultimedia,
    eCommunications,
    ERole_enum_count
  );

  PROPERTYKEY = record
    fmtid: TGUID;
    pid: DWORD;
  end;

  PROPVARIANT = record
    vt: Word;
    wReserved1, wReserved2, wReserved3: Word;
    case Integer of
      0: (cVal: Char);
      1: (bVal: Byte);
      2: (iVal: Smallint);
      3: (uiVal: Word);
      4: (lVal: Longint);
      5: (ulVal: LongWord);
      6: (intVal: Integer);
      7: (uintVal: Cardinal);
      8: (hVal: Largeint);
      9: (uhVal: Int64);
  end;

  IPropertyStore = interface(IUnknown)
    ['{886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99}']
    function GetCount(out cProps: DWORD): HResult; stdcall;
    function GetAt(iProp: DWORD; out pkey: PROPERTYKEY): HResult; stdcall;
    function GetValue(const key: PROPERTYKEY; var pv: PROPVARIANT): HResult; stdcall;
    function SetValue(const key: PROPERTYKEY; const propvar: PROPVARIANT): HResult; stdcall;
    function Commit: HResult; stdcall;
  end;

  IMMNotificationClient = interface(IUnknown)
    ['{7991EEC9-7E89-4D85-8390-6C703CEC60C0}']
    function OnDeviceStateChanged(const pwstrDeviceId: PWideChar; dwNewState: DWORD): HResult; stdcall;
    function OnDeviceAdded(const pwstrDeviceId: PWideChar): HResult; stdcall;
    function OnDeviceRemoved(const pwstrDeviceId: PWideChar): HResult; stdcall;
    function OnDefaultDeviceChanged(flow: EDataFlow; role: ERole; const pwstrDeviceId: PWideChar): HResult; stdcall;
    function OnPropertyValueChanged(const pwstrDeviceId: PWideChar; const key: PROPERTYKEY): HResult; stdcall;
  end;

  IMMDevice = interface(IUnknown)
    ['{D666063F-1587-4E43-81F1-B948E807363F}']
    function Activate(const iid: TGUID; dwClsCtx: DWORD; pActivationParams: Pointer; out ppInterface: Pointer): HResult; stdcall;
    function OpenPropertyStore(stgmAccess: DWORD; out ppProperties: IPropertyStore): HResult; stdcall;
    function GetId(out ppstrId: PWideChar): HResult; stdcall;
    function GetState(out pdwState: DWORD): HResult; stdcall;
  end;

  IMMDeviceCollection = interface(IUnknown)
    ['{0BD7A1BE-7A1A-44DB-8397-C0A0B755BFA8}']
    function GetCount(out pcDevices: UINT): HResult; stdcall;
    function Item(nDevice: UINT; out ppDevice: IMMDevice): HResult; stdcall;
  end;

  IMMDeviceEnumerator = interface(IUnknown)
    ['{A95664D2-9614-4F35-A746-DE8DB63617E6}']
    function EnumAudioEndpoints(dataFlow: EDataFlow; dwStateMask: DWORD; out ppDevices: IMMDeviceCollection): HResult; stdcall;
    function GetDefaultAudioEndpoint(dataFlow: EDataFlow; role: ERole; out ppEndpoint: IMMDevice): HResult; stdcall;
    function GetDevice(pwstrId: PWideChar; out ppDevice: IMMDevice): HResult; stdcall;
    function RegisterEndpointNotificationCallback(pClient: IMMNotificationClient): HResult; stdcall;
    function UnregisterEndpointNotificationCallback(pClient: IMMNotificationClient): HResult; stdcall;
  end;

  TAudioDeviceChangedEvent = procedure(Sender: TObject) of object;

  TAudioDeviceMonitor = class(TComponent, IMMNotificationClient)
  private
    FDeviceEnumerator: IMMDeviceEnumerator;
    FOnDefaultDeviceChangedEvent: TAudioDeviceChangedEvent;
  protected
    function OnDefaultDeviceChanged(flow: EDataFlow; role: ERole; const pwstrDeviceId: PWideChar): HResult; stdcall;
    function OnDeviceAdded(const pwstrDeviceId: PWideChar): HResult; stdcall;
    function OnDeviceRemoved(const pwstrDeviceId: PWideChar): HResult; stdcall;
    function OnDeviceStateChanged(const pwstrDeviceId: PWideChar; dwNewState: DWORD): HResult; stdcall;
    function OnPropertyValueChanged(const pwstrDeviceId: PWideChar; const key: PROPERTYKEY): HResult; stdcall;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  published
    property OnDefaultDeviceChangedEvent: TAudioDeviceChangedEvent read FOnDefaultDeviceChangedEvent write FOnDefaultDeviceChangedEvent;
  end;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('Fablesalive', [TAudioDeviceMonitor]);
end;

{ TAudioDeviceMonitor }

constructor TAudioDeviceMonitor.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  CoInitialize(nil);
  CoCreateInstance(CLSID_MMDeviceEnumerator, nil, CLSCTX_INPROC_SERVER, IID_IMMDeviceEnumerator, FDeviceEnumerator);
  FDeviceEnumerator.RegisterEndpointNotificationCallback(Self as IMMNotificationClient);
end;

destructor TAudioDeviceMonitor.Destroy;
begin
  if Assigned(FDeviceEnumerator) then
    FDeviceEnumerator.UnregisterEndpointNotificationCallback(Self as IMMNotificationClient);
  CoUninitialize;
  inherited Destroy;
end;

function TAudioDeviceMonitor.OnDefaultDeviceChanged(flow: EDataFlow; role: ERole; const pwstrDeviceId: PWideChar): HResult;
begin
  if Assigned(FOnDefaultDeviceChangedEvent) then
    FOnDefaultDeviceChangedEvent(Self);
  Result := S_OK;
end;

function TAudioDeviceMonitor.OnDeviceAdded(const pwstrDeviceId: PWideChar): HResult;
begin
  Result := S_OK;
end;

function TAudioDeviceMonitor.OnDeviceRemoved(const pwstrDeviceId: PWideChar): HResult;
begin
  Result := S_OK;
end;

function TAudioDeviceMonitor.OnDeviceStateChanged(const pwstrDeviceId: PWideChar; dwNewState: DWORD): HResult;
begin
  Result := S_OK;
end;

function TAudioDeviceMonitor.OnPropertyValueChanged(const pwstrDeviceId: PWideChar; const key: PROPERTYKEY): HResult;
begin
  Result := S_OK;
end;

end.

