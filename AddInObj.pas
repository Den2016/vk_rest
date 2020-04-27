unit AddInObj;

interface

uses  { Какие библиотеки используем }
{$IFDEF DEBUGDC}
  dbugintf, Dialogs,
{$ENDIF}
  ComServ, ComObj, ActiveX, SysUtils, Windows, AddInLib, Classes,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdCustomTCPServer, IdTCPServer, IdUDPBase, IdUDPClient,
  IdContext, Winapi.Messages, System.Net.URLClient, System.Net.HttpClient, System.Net.HttpClientComponent,
  System.WideStrUtils, System.JSON;

     const c_AddinName = 'vk_rest'; //Имя внешней компоненты

     //Количество свойств
     const c_PropCount = 5;

     //Идентификаторы свойств
     type TProperties = (
       propURL,
       propFileName,
       propFileResponse,
       propStatusCode,
       propTimeout
     );

     //Имена свойств, видимые из 1С
     //Порядок соблюдайте точно такой же, что и в TProperties
    const c_PropNames: Array[0..c_PropCount-1, 0..1] of WideString =
    (
      ('URL','Ссылка'),
      ('FileName','Файл'),
      ('FileResponse','ФайлОтвета'),
      ('StatusCode','КодСтатуса'),
      ('TimeOut','Таймаут')
    );

    //Количество методов
     const c_MethCount = 15;
    //Идентификаторы методов.
    type TMethods = (
       methRequest,
       methJSONSuccess, // проверка на успешный разбор JSON
       methJSONGetValueType, // получить тип значения JSON
       methJSONGetValue, // получить значение JSON
       methJSONArrayCount,  // получить размер массива
       methJSONGetArrayValue, // получить элемент массива, параметры - имя, строка, имя
       methJSONGetArrayValueType, // получить тип элемент массива, параметры - имя, строка, имя
       methSetRoot,  // сохранить ссылку на узел под заданным именем
       methSetRootFromArray, // получить значение сохраненного элемента JSON
       methClearRoot,  // получить размер массива, сохраненного под именем ранее
       methSetParentRoot, // получить элемент массива, сохраненного под именем ранее, параметры - имя массива, строка, колонка
       methGetPairsCount,
       methGetPairName,
       methGetPairValue,
       methParseFile

       );

    //Имена методов, видимые из 1С
     //Порядок соблюдайте точно такой же, что и в TMethods
    const c_MethNames: Array[0..c_MethCount-1,0..3] of WideString =
    (
    ('Request','ВыполнитьЗапрос','Выполнить','0'), // третий параметр - количество параметров метода
    ('Success','JSONРазобран','Разобран','0'), // третий параметр - количество параметров метода
    ('GetValueType','JSONТипЗначения','ТипЗначения','1'), // третий параметр - количество параметров метода
    ('GetValue','JSONЗначение','Значение','1'), // третий параметр - количество параметров метода
    ('GetArrayCount','JSONРазмерМассива','РазмерМассива','1'), // третий параметр - количество параметров метода
    ('GetArrayValue','JSONЗначениеИзМассива','ЗначениеИзМассива','3'), // третий параметр - количество параметров метода
    ('GetArrayValueType','JSONТипЗначенияИзМассива','ТипЗначенияИзМассива','2'), // третий параметр - количество параметров метода
    ('SetRoot','JSONУстановитьКорень','УстановитьКорень','1'), // параметр - имя поля для задания текущего root-объекта
    ('SetRootFromArray','JSONУстановитьКореньЗначениемИзМассива','УстановитьКореньИзМассива','2'), // параметр - имя поля для задания текущего root-объекта
    ('ClearRoot','JSONСброситьКорень','СброситьКорень','0'),
    ('SetParentRoot','JSONУстановитьПредыдущийКорень','УстановитьПредыдущийКорень','0'),
    ('GetPairsCount','JSONКоличествоПар','КоличествоПар','0'),
    ('GetPairName','JSONИмяПары','ИмяПары','1'), // индекс пары
    ('GetPairValue','JSONЗначениеПары','ЗначениеПары','1'), //индекс пары
    ('ParseFile','РазобратьФайл','РазобратьФайл','1') //разобрать файл с json
    );

const
{Нажмите Ctrl-Shift-G чтобы сгенерировать новый уникальный идентификатор GUID}
     CLSID_AddInObject : TGUID = '{BF974081-BFB7-4C48-94D7-39A3C9DF4FB8}';

type

  AddInObject = class(TComObject, IDispatch, IInitDone, ILanguageExtender)
  private
    objURL:string;
    objFileName:string;
    objFileResponse:string;
    objTimeout:Integer;
    NetHTTPClient:TNetHTTPClient;
    NetHTTPRequest:TNetHTTPRequest;
    objResponse:IHTTPResponse;
    JSON, RJSON:TJSONValue;
    stack:TArray<TJSONValue>;
    function SendRESTRequest:TJSONObject;
    function ParseFile(Filename:String):byte;
  public
    i1cv7: IDispatch;
    iStatus: IStatusLine;
//    iExtWindows: IExtWndsSupport;
    iError: IErrorLog;
    iEvent : IAsyncEvent;
    _App: OleVariant;
  protected
    { These two methods is convenient way to access function
      parameters from SAFEARRAY vector of variants }
    function GetNParam(var pArray : PSafeArray; lIndex: Integer ): OleVariant;
    procedure PutNParam(var pArray: PSafeArray; lIndex: Integer; var varPut: OleVariant);


    { IInitDone implementation }
    function Init(pConnection: IDispatch): HResult; stdcall;
    function Done: HResult; stdcall;
    function GetInfo(var pInfo: PSafeArray): HResult; stdcall;

    { ILanguageExtender implementation }
    function RegisterExtensionAs(var bstrExtensionName: WideString): HResult; stdcall;
    function GetNProps(var plProps: Integer): HResult; stdcall;
    function FindProp(const bstrPropName: WideString; var plPropNum: Integer): HResult; stdcall;
    function GetPropName(lPropNum, lPropAlias: Integer; var pbstrPropName: WideString): HResult; stdcall;
    function GetPropVal(lPropNum: Integer; var pvarPropVal: OleVariant): HResult; stdcall;
    function SetPropVal(lPropNum: Integer; var varPropVal: OleVariant): HResult; stdcall;
    function IsPropReadable(lPropNum: Integer; var pboolPropRead: Integer): HResult; stdcall;
    function IsPropWritable(lPropNum: Integer; var pboolPropWrite: Integer): HResult; stdcall;
    function GetNMethods(var plMethods: Integer): HResult; stdcall;
    function FindMethod(const bstrMethodName: WideString; var plMethodNum: Integer): HResult; stdcall;
    function GetMethodName(lMethodNum, lMethodAlias: Integer; var pbstrMethodName: WideString): HResult; stdcall;
    function GetNParams(lMethodNum: Integer; var plParams: Integer): HResult; stdcall;
    function GetParamDefValue(lMethodNum, lParamNum: Integer; var pvarParamDefValue: OleVariant): HResult; stdcall;
    function HasRetVal(lMethodNum: Integer; var pboolRetValue: Integer): HResult; stdcall;
    function CallAsProc(lMethodNum: Integer; var paParams: PSafeArray): HResult; stdcall;
    function CallAsFunc(lMethodNum: Integer; var pvarRetValue: OleVariant; var paParams: PSafeArray): HResult; stdcall;

    { IDispatch }
    function GetIDsOfNames(const IID: TGUID; Names: Pointer;
      NameCount, LocaleID: Integer; DispIDs: Pointer): HResult; virtual; stdcall;
    function GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; virtual; stdcall;
    function GetTypeInfoCount(out Count: Integer): HResult; virtual; stdcall;
    function Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult; virtual; stdcall;

    { IStatusLine }
    function SetStatusLine(const bstrSource: WideString): HResult; safecall;
    function ResetStatusLine(): HResult; safecall;

    procedure ShowErrorLog(fMessage:WideString);
  end;



implementation

//=======================  General functions  ================================
///////////////////////////////////////////////////////////////////////
function AddInObject.GetNParam(var pArray : PSafeArray; lIndex: Integer ): OleVariant;
var
  varGet : OleVariant;
begin
  SafeArrayGetElement(pArray,lIndex,varGet);
  GetNParam := varGet;
end;

///////////////////////////////////////////////////////////////////////
function AddInObject.ParseFile(Filename: String): byte;
Var FileStream:TFileStream;
    MemStream:TMemoryStream;
    S:String;
begin
  Result:=0;
  FileStream:=TFileStream.Create(Filename,fmOpenRead);
  MemStream:=TMemoryStream.Create;
  try
    MemStream.CopyFrom(FileStream,FileStream.Size);
    //ConvertStreamFromAnsiToUTF8(FileStream,MemStream); // конвертируем из файла ansi в utf-8
    MemStream.Seek(0,soFromEnd);
    MemStream.Write(#0#0,2);
    MemStream.Seek(0,soFromBeginning);
    FreeAndNil(JSON);
    S:=PAnsiChar(MemStream.Memory);
    JSON:=TJSONObject.ParseJSONValue(S);
    RJSON:=JSON;
    SetLength(stack,1);
    stack[0]:=JSON;
    if(JSON<> nil) then Result:=1;
  finally
    FreeAndNil(MemStream);
    FreeAndNil(FileStream);
  end;
end;

procedure AddInObject.PutNParam(var pArray: PSafeArray; lIndex: Integer; var varPut: OleVariant);
begin
  SafeArrayPutElement(pArray,lIndex,varPut);
end;



//======================= IInitDone interface ================================
///////////////////////////////////////////////////////////////////////
function AddInObject.Init(pConnection: IDispatch): HResult; stdcall;
//var  wnd: HWND;
begin
  i1cv7:=pConnection;
  NetHTTPClient:=TNetHTTPClient.Create(nil);
  NetHTTPRequest:=TNetHTTPRequest.Create(nil);
  NetHTTPRequest.Client:=NetHTTPClient;
  objTimeout:=3000;

  iError:=nil;
  pConnection.QueryInterface(IID_IErrorLog,iError);

  iStatus:=nil;
  pConnection.QueryInterface(IID_IStatusLine,iStatus);

  iEvent := nil;
  pConnection.QueryInterface(IID_IAsyncEvent,iEvent);
  iEvent.SetEventBufferDepth(300); //глубина буфера событий

  //iExtWindows:=nil;
  //pConnection.QueryInterface(IID_IExtWndsSupport,iExtWindows);

  {
  iExtWindows.GetAppMainFrame(wnd);
  Application.Handle := wnd;
   }
  _App:=pConnection;

  Init := S_OK;
end;

///////////////////////////////////////////////////////////////////////
function AddInObject.Done: HResult; stdcall;
begin
  FreeAndNil(NetHTTPRequest);
  FreeAndNil(NetHTTPClient);
  If ( iStatus <> nil ) then
    iStatus._Release();

//  If ( iExtWindows <> nil ) then
//    iExtWindows._Release();

  If ( iError <> nil ) then
    iError._Release();

  if (iEvent <> nil) then
    iEvent._Release();

  Done := S_OK;
end;
///////////////////////////////////////////////////////////////////////
function AddInObject.GetInfo(var pInfo: PSafeArray{(OleVariant)}): HResult; stdcall;
var  varInfo : OleVariant;
begin
  varInfo := '2000';
  PutNParam(pInfo,0,varInfo);

  GetInfo := S_OK;
end;

//======================= IStatusLine Interface ==============================
///////////////////////////////////////////////////////////////////////
function AddInObject.SetStatusLine(const bstrSource: WideString): HResult; safecall;
begin
  SetStatusLine:=S_OK;
end;
///////////////////////////////////////////////////////////////////////
function AddInObject.ResetStatusLine(): HResult; safecall;
begin
  //ResetStatusLine: = S_OK;
end;

//======================= ILanguageExtender Interface ========================
///////////////////////////////////////////////////////////////////////
function AddInObject.RegisterExtensionAs(var bstrExtensionName: WideString): HResult; stdcall;
begin
  bstrExtensionName := c_AddinName;
  RegisterExtensionAs := S_OK;
end;
///////////////////////////////////////////////////////////////////////
function AddInObject.GetNProps(var plProps: Integer): HResult; stdcall;
begin
     plProps := Integer(c_PropCount);
     GetNProps := S_OK;
end;
///////////////////////////////////////////////////////////////////////
function AddInObject.FindProp(const bstrPropName: WideString; var plPropNum: Integer): HResult; stdcall;
var
  NewPropName: WideString;
  i: Integer;
begin
     plPropNum := -1;

     NewPropName:=bstrPropName;

     for i:=0 to c_PropCount-1 do begin
       if (NewPropName=c_PropNames[i,0]) or (NewPropName=c_PropNames[i,1]) then begin
         plPropNum:=i;
         break;
       end;
     end;

     if (plPropNum = -1) then
       begin
         FindProp := S_FALSE;
         Exit;
       end;

     FindProp := S_OK;
end;
///////////////////////////////////////////////////////////////////////
function AddInObject.GetPropName(lPropNum, lPropAlias: Integer; var pbstrPropName: WideString): HResult; stdcall;
begin
     pbstrPropName := '';
     if (lPropAlias<>0) and (lPropAlias<>1) then begin
            GetPropName := S_FALSE;
            Exit;
     end;
     if (lPropNum<0) or (lPropNum>=c_PropCount) then begin
            GetPropName := S_FALSE;
            Exit;
     end;

     pbstrPropName := c_PropNames[lPropNum, lPropAlias];

     GetPropName := S_OK;
end;
///////////////////////////////////////////////////////////////////////
function AddInObject.GetPropVal(lPropNum: Integer; var pvarPropVal: OleVariant): HResult; stdcall;
//Здесь 1С читает значения свойств
begin
    VarClear(pvarPropVal);
    try
        case TProperties(lPropNum) of
            propURL: pvarPropVal := objURL;
            propFileName: pvarPropVal := objFileName;
            propFileResponse: pvarPropVal := objFileResponse;
            propStatusCode: pvarPropVal := objResponse.StatusCode;
            propTimeout: pvarPropVal := objTimeout;
            else
              GetPropVal := S_FALSE;
              Exit;
        end;
    except
        on E:Exception do begin
            ShowErrorLog(E.Message);
        end;
    end; //try
    GetPropVal := S_OK;
end;
///////////////////////////////////////////////////////////////////////
function AddInObject.SendRESTRequest:TJSONObject;
var FileStream:TFileStream;
    MemStream:TMemoryStream;
    R:IHTTPResponse;
    ResponseFilename:String;
    headers:TNetHeaders;
begin
  Result:=nil;
  ResponseFilename:=objFileResponse;
  if ResponseFilename='' then ResponseFilename:=objFileName+'.response.json';
  NetHTTPClient.AcceptCharSet:='UTF-8, *;q=0.8';
  NetHTTPClient.AllowCookies:=true;
  NetHTTPClient.Asynchronous:=false;
  NetHTTPClient.ConnectionTimeout:=objTimeout;
  NetHTTPClient.HandleRedirects:=True;
  NetHTTPClient.UserAgent:='vk_rest component';
  NetHTTPRequest.Client:=NetHTTPClient;
  FreeAndNil(JSON); // обнуляем результат предыдущего запроса, чтобы не было утечек памяти при частом использовании
  try
    MemStream:=TMemoryStream.Create; // стрим в памяти
    FileStream:=TFileStream.Create(objFileName,fmOpenReadWrite); // открываем файл
    SetLength(headers, 1); // задаем заголовки
    headers[Length(headers)-1].Name:='Content-type';
    headers[Length(headers)-1].Value:='application/json; charset=utf-8';
    try
      ConvertStreamFromAnsiToUTF8(FileStream,MemStream); // конвертируем из файла ansi в utf-8
      R:=NetHTTPRequest.Post(objURL,MemStream,nil,headers); // делаем запрос
      MemStream.Clear; // чистим стрим в памяти
      ConvertStreamFromUTF8ToAnsi(R.ContentStream,MemStream); // перегоняем обратно из utf8 в ansi
      FreeAndNil(FileStream); // закрываем файл,
      try
        // parsing response
        JSON:=TJSONObject.ParseJSONValue(R.ContentAsString);
        RJSON:=JSON;                // copy to currentroot
        SetLength(stack,1);   // обнуляем стек и помещаем в вершину корневой JSON
        stack[0]:=JSON;
      finally
      // в любом случае записываем ответ в файл
        FileStream:=TFileStream.Create(ResponseFilename,fmCreate or fmOpenWrite or fmShareCompat);
        FileStream.CopyFrom(MemStream,MemStream.Size);
      end;
    except
      on E:Exception do begin
        ShowErrorLog(E.Message);
      end;
    end;
  finally
    FreeAndNil(FileStream);
    FreeAndNil(MemStream);
  end;
  //Result:=TJSONObject(JSON);

end;

function AddInObject.SetPropVal(lPropNum: Integer; var varPropVal: OleVariant): HResult; stdcall;
//Здесь 1С устанавливает значения свойств
//Var X:Integer;
begin
     try
          case TProperties(lPropNum) of
              propURL: objURL:=varPropVal;
              propFileName: begin
                              if objFileName=objFileResponse then objFileResponse:=varPropVal;
                              objFileName:=varPropVal;
                              if objFileResponse='' then objFileResponse:=varPropVal;
                            end;
              propFileResponse: objFileResponse:=varPropVal;
              propTimeout: objTimeout:=varPropVal;
              else
                SetPropVal := S_FALSE;
                Exit;
          end;
     except
          on E:Exception do begin
              ShowErrorLog(E.Message);
          end;
     end; //try
     SetPropVal := S_OK;
end;
///////////////////////////////////////////////////////////////////////
function AddInObject.IsPropReadable(lPropNum: Integer; var pboolPropRead: Integer): HResult; stdcall;
{Здесь 1С узнает, можно ли читать свойства}
begin
//здесь все свойства читабельны
  pboolPropRead := 1;

//     case TProperties(lPropNum) of
//          propErrorMsg: pboolPropRead := 1;{1=можно читать свойство, 0=нет}
//     else
//            IsPropReadable := S_FALSE;
//            Exit;
//     end;
  IsPropReadable := S_OK;

end;
///////////////////////////////////////////////////////////////////////
function AddInObject.IsPropWritable(lPropNum: Integer; var pboolPropWrite: Integer): HResult; stdcall;
//Здесь 1С узнает, можно ли изменять свойство
begin
     case TProperties(lPropNum) of
          propURL: pboolPropWrite := 1;{1=можно записывать свойство, 0=нет}
          propFileName:pboolPropWrite := 1;{1=можно записывать свойство, 0=нет}
          propFileResponse:pboolPropWrite := 1;{1=можно записывать свойство, 0=нет}
          propStatusCode:pboolPropWrite := 0;{1=можно записывать свойство, 0=нет}
          propTimeout:pboolPropWrite := 1;{1=можно записывать свойство, 0=нет}
          else
            IsPropWritable := S_FALSE;
            Exit;
     end;

     IsPropWritable := S_OK;
end;


///////////////////////////////////////////////////////////////////////
function AddInObject.GetNMethods(var plMethods: Integer): HResult; stdcall;
begin
     plMethods := c_MethCount;
     GetNMethods := S_OK;
end;
///////////////////////////////////////////////////////////////////////
function AddInObject.FindMethod(const bstrMethodName: WideString; var plMethodNum: Integer): HResult; stdcall;
var NewMethodName: WideString;
var i:Integer;
begin
  NewMethodName := bstrMethodName;

     plMethodNum := -1;

     for i:=0 to c_MethCount-1 do begin
       if (NewMethodName=c_MethNames[i,0]) or (NewMethodName=c_MethNames[i,1]) or (NewMethodName=c_MethNames[i,2]) then begin
         plMethodNum := i;
         break;
       end;
     end;

     if (plMethodNum = -1) then
       begin
         FindMethod := S_FALSE;
         Exit;
       end;

     FindMethod := S_OK;

end;
///////////////////////////////////////////////////////////////////////
function AddInObject.GetMethodName(lMethodNum, lMethodAlias: Integer; var pbstrMethodName: WideString): HResult; stdcall;
begin

     pbstrMethodName := '';
     if (lMethodAlias<>0) and (lMethodAlias<>1) then begin
            Result := S_FALSE;
            Exit;
     end;
     if (lMethodNum<0) or (lMethodNum>=c_MethCount) then begin
            Result := S_FALSE;
            Exit;
     end;

     pbstrMethodName := c_MethNames[lMethodNum, lMethodAlias];

     GetMethodName := S_OK;

end;

///////////////////////////////////////////////////////////////////////
function AddInObject.GetNParams(lMethodNum: Integer; var plParams: Integer): HResult; stdcall;
//Здесь 1С узнает количество параметров у методов
begin

     plParams := StrToInt(c_MethNames[lMethodNum, 3]);
(*     plParams := 0;

     case TMethods(lMethodNum) of

          methGetContactShow: plParams := 1;{1 параметр}
          methGetContactStatus: plParams := 1;{1 параметр}
          methSendMessage: plParams := 2;{кому и сообщение}
          methSubscribe: plParams := 1;{кому}
          methSubscribeOK: plParams := 1;{кому}
          methUnSubscribe: plParams := 1;{кому}
          methSocketSend: plParams := 1;
          else
            begin
               GetNParams := S_FALSE;
               Exit;
            end;
     end;
  *)
     GetNParams := S_OK;

end;
///////////////////////////////////////////////////////////////////////
function AddInObject.GetParamDefValue(lMethodNum, lParamNum: Integer; var pvarParamDefValue: OleVariant): HResult; stdcall;
begin
  { Ther is no default value for any parameter }
  VarClear(pvarParamDefValue);
  GetParamDefValue := S_OK;
end;
///////////////////////////////////////////////////////////////////////
function AddInObject.HasRetVal(lMethodNum: Integer; var pboolRetValue: Integer): HResult; stdcall;
//Здесь 1С узнает, какие методы работают как функции
begin
  pboolRetValue := 1; //Все методы возвращают значение
  HasRetVal := S_OK;
end;



///////////////////////////////////////////////////////////////////////
function AddInObject.CallAsProc(lMethodNum: Integer; var paParams: PSafeArray{(OleVariant)}): HResult; stdcall;
//Здесь 1С выполняет код процедур
begin
    CallAsProc := S_FALSE;
end;

///////////////////////////////////////////////////////////////////////
function AddInObject.CallAsFunc(lMethodNum: Integer; var pvarRetValue: OleVariant; var paParams: PSafeArray): HResult; stdcall;
{Здесь 1С выполняет код функций}
var
//_to,_msg: String;
//    ss:TStringStream;
//    TagName: ShortString;
    s,a: String;
    fname: String;
//    AttrName, AttrValue: String;
    x:Integer;
    J:TJSONValue;
begin
  pvarRetValue:=0;
  try
    case TMethods(lMethodNum) of
      methRequest: begin
        SendRESTRequest;
      end;
      methJSONSuccess: begin
        pvarRetValue:=0;
        if Assigned(JSON) then pvarRetValue:=1;
      end;
      methJSONGetValueType: begin
        pvarRetValue:='';
        a:=GetNParam(paParams,0);
        if(a='') then begin
          s:=TJSONObject(RJSON).ClassName;
          pvarRetValue:=s;
        end else begin
          if TJSONObject(RJSON).Values[a]<>nil then begin
              s:=TJSONObject(RJSON).GetValue(a).ClassName;
              //if s='TJSONArray' then pvarRetValue:='Массив';
              //if s='TJSONString' then pvarRetValue:='Строка';
              //if s='TJSONNumber' then pvarRetValue:='Число';
              //if s='TJSONTrue' then pvarRetValue:='Число';
              //if s='TJSONFalse' then pvarRetValue:='Число';
              pvarRetValue:=s;
          end;
        end;
      end;
      methJSONGetValue:
        begin
          a:=GetNParam(paParams,0);
          if TJSONObject(RJSON).Values[a]=nil then raise Exception.Create('Неверное имя '+a);
          s:=TJSONObject(RJSON).Values[a].ClassName;
          if s='TJSONArray' then raise Exception.Create('Узел '+a+' является массивом');
          pvarRetValue:=TJSONObject(RJSON).Values[a].Value;
          if s='TJSONTrue' then pvarRetValue:=1;
          if s='TJSONFalse' then pvarRetValue:=0;
        end;
      methJSONArrayCount:
        begin  // получить размер массива
          a:=GetNParam(paParams,0);
          if(a='') then begin
            s:=TJSONObject(RJSON).ClassName;
            if s='TJSONArray' then begin
              pvarRetValue:=TJSONArray(RJSON).Count;
            end else  raise Exception.Create('Текущий корневой узел '+a+' не является массивом');
          end else begin
            if TJSONObject(RJSON).Values[a]=nil then raise Exception.Create('Неверное имя '+a);
            s:=TJSONObject(RJSON).Values[a].ClassName;
            if s='TJSONArray' then begin
              pvarRetValue:=TJSONArray(TJSONObject(RJSON).Values[a]).Count;
            end else  raise Exception.Create('Узел '+a+' не является массивом');
          end;
        end;
      methJSONGetArrayValueType:
        begin
          a:=GetNParam(paParams,0);
          x:=GetNParam(paParams,1);
          if(a='') then begin
            s:=TJSONObject(RJSON).ClassName;
            if s='TJSONArray' then begin
              J:=TJSONArray(RJSON).Items[x];
              s:=TJSONObject(J).ClassName;
              pvarRetValue:=s;
            end else  raise Exception.Create('Текущий корневой узел '+a+' не является массивом');
          end else begin
            if TJSONObject(RJSON).Values[a]=nil then raise Exception.Create('Неверное имя '+a);
            s:=TJSONObject(RJSON).Values[a].ClassName;
            if s='TJSONArray' then begin
              J:=TJSONArray(TJSONObject(RJSON).Values[a]).Items[x];
              s:=TJSONObject(J).ClassName;
              pvarRetValue:=s;
            end else  raise Exception.Create('Узел '+a+' не является массивом');
          end;

        end;
      methJSONGetArrayValue:
        begin // получить элемент массива, параметры - имя, строка, имя
          a:=GetNParam(paParams,0);
          x:=GetNParam(paParams,1);
          fname:=GetNParam(paParams,2);
          if(a='') then begin
            s:=TJSONObject(RJSON).ClassName;
            if s='TJSONArray' then begin
              J:=TJSONArray(RJSON).Items[x];
            end else  raise Exception.Create('Текущий корневой узел '+a+' не является массивом');

          end else begin
            if TJSONObject(RJSON).Values[a]=nil then raise Exception.Create('Неверное имя '+a);
            s:=TJSONObject(RJSON).Values[a].ClassName;
            if s='TJSONArray' then begin
              J:=TJSONArray(TJSONObject(RJSON).Values[a]).Items[x];
            end else  raise Exception.Create('Узел '+a+' не является массивом');
          end;
          if(fname='') then s:=TJSONObject(J).ClassName
            else s:=TJSONObject(J).Values[fname].ClassName;
          if s='TJSONArray' then raise Exception.Create('Узел '+fname+' является массивом');
          if s='TJSONObject' then raise Exception.Create('Узел '+fname+' является объектом');
          if(fname='') then pvarRetValue := J.Value
            else pvarRetValue:=TJSONObject(J).Values[fname].Value;
          if s='TJSONTrue' then pvarRetValue:=1;
          if s='TJSONFalse' then pvarRetValue:=0;

        end;
      methSetRoot:  // установить новый корень
        begin
          a:=GetNParam(paParams,0);
          if TJSONObject(RJSON).Values[a]=nil then raise Exception.Create('Неверное имя '+a);
          SetLength(stack,Length(stack)+1);
          stack[Length(stack)-1]:=RJSON;
          RJSON:=TJSONObject(RJSON).Values[a];
        end;
      methSetRootFromArray: // установить корень из массива, параметры - имя массива, индекс
        begin
          a:=GetNParam(paParams,0);
          x:=GetNParam(paParams,1);
          if(a='') then begin
            s:=TJSONObject(RJSON).ClassName;
            if s='TJSONArray' then begin
              J:=TJSONArray(RJSON).Items[x];
              SetLength(stack,Length(stack)+1);
              stack[Length(stack)-1]:=RJSON;
              RJSON:=J;
            end else  raise Exception.Create('Текущий корневой узел '+a+' не является массивом');
          end else begin
            if TJSONObject(RJSON).Values[a]=nil then raise Exception.Create('Неверное имя '+a);
            s:=TJSONObject(RJSON).Values[a].ClassName;
            if s='TJSONArray' then begin
              J:=TJSONArray(TJSONObject(RJSON).Values[a]).Items[x];
              SetLength(stack,Length(stack)+1);
              stack[Length(stack)-1]:=RJSON;
              RJSON:=J;
            end else  raise Exception.Create('Узел '+a+' не является массивом');
          end;
        end;
      methClearRoot:  // сбросить корень
        begin
          RJSON:=JSON;
          SetLength(stack,1);
          stack[0]:=RJSON;
        end;
      methSetParentRoot: // установить предыдущий корень
        begin
          RJSON:=stack[Length(stack)-1];
          SetLength(stack,Length(stack)-1);
        end;
      methGetPairsCount: //количество пар
        pvarRetValue:=TJSONObject(RJSON).Count;
      methGetPairName:
        begin
          s:=TJSONObject(RJSON).Pairs[GetNParam(paParams,0)].JsonString.Value;
          pvarRetValue:=s;
        end;
      methGetPairValue:
        begin
          s:=TJSONObject(RJSON).Pairs[GetNParam(paParams,0)].JsonValue.Value;
          pvarRetValue:=s;
        end;
      methParseFile:
        begin
          a:=GetNParam(paParams,0);
          pvarRetValue:=ParseFile(a);
        end;

      else begin
               CallAsFunc := S_FALSE;
               Exit;
               end;
          end; //case

  except

           on E:Exception do begin
             ShowErrorLog(E.Message);
             CallAsFunc := S_FALSE;
             Exit;
           end;

  end; //try
  CallAsFunc := S_OK;
end;
///////////////////////////////////////////////////////////////////////
function AddInObject.GetIDsOfNames(const IID: TGUID; Names: Pointer;
  NameCount, LocaleID: Integer; DispIDs: Pointer): HResult;
begin
  Result := E_NOTIMPL;
end;
///////////////////////////////////////////////////////////////////////
function AddInObject.GetTypeInfo(Index, LocaleID: Integer;
  out TypeInfo): HResult;
begin
  Result := E_NOTIMPL;
end;

///////////////////////////////////////////////////////////////////////
function AddInObject.GetTypeInfoCount(out Count: Integer): HResult;
begin
  Result := E_NOTIMPL;
end;


///////////////////////////////////////////////////////////////////////
function AddInObject.Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
  Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult;
begin
  Result := E_NOTIMPL;
end;

///////////////////////////////////////////////////////////////////////
procedure AddInObject.ShowErrorLog(fMessage:WideString);
var
  ErrInfo: PExcepInfo;
begin
  If Trim(fMessage) = '' then Exit;
  New(ErrInfo);
  ErrInfo^.bstrSource := c_AddinName;
  ErrInfo^.bstrDescription := fMessage;
  ErrInfo^.wCode:=1006;
  ErrInfo^.sCode:=E_FAIL; //генерация исключения в 1С
  iError.AddError(nil, ErrInfo);
end;

{
///////////////////////////////////////////////////////////////////////
//Процедура потока
procedure TMyThread.Execute;
var str: String;
begin
  try
     repeat
       str:=MyObject.g_cp.ReadString;
       str:=trim(str);
       if str<>'' then begin
         //MessageBox(0, pchar('Считан код: '+str), '*debug',0);
         MyObject.iEvent.ExternalEvent(c_AddinName, 'BarCodeValue', str);
       end;

       EnterCriticalSection(g_kb_CriticalSection);

         try
           if g_sz_barcodes.Count>0 then begin
             MyObject.iEvent.ExternalEvent(c_AddinName, 'BarCodeValue', g_sz_barcodes.Strings[0]);
             g_sz_barcodes.Delete(0);
             g_kb_str:='';
           end;
         except
         end;

       LeaveCriticalSection(g_kb_CriticalSection);
       sleep(500);
     until terminated;
  except
     on E:Exception do begin
       MyObject.ShowErrorLog('Ошибка чтения из COM-порта: '+E.Message);
     end;
  end;

end;

///////////////////////////////////////////////////////////////////////
constructor TMyThread.Create(prm_Obj:AddInObject);
begin
    inherited Create(False);
    MyObject:=prm_Obj;
end;

}

///////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////
procedure Close1C();
//var cnt: Integer;
begin
{     GetIniFile();
     cnt:=g_Delay;
     if cnt=0 then exit;
     repeat
       if g_Message<>'' then begin
         ShowBalloon(g_Message, 'Осталось '+IntToStr(cnt)+' секунд',750);
       end else begin
         Sleep(750);
       end;
       Sleep(250);
     Dec(cnt);
     until cnt=0;
}
     Windows.TerminateProcess(Windows.GetCurrentProcess(),1);
end;

{ TCloseTimer }




initialization
  ComServer.SetServerName('AddIn');
  TComObjectFactory.Create(ComServer,AddInObject,CLSID_AddInObject,
    c_AddinName,'V7 AddIn 2.0',ciSingleInstance);

//  InitializeCriticalSection(g_kb_CriticalSection);

finalization

//  DeleteCriticalSection(g_kb_CriticalSection);


end.
