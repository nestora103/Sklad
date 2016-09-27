unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComObj, ActiveX, StdCtrls, Grids, OleServer, ExcelXP,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdLPR, Menus,
  OleCtnrs, WordXP,DateUtils,xpman;
const
//константа заполняемых строк 1 документа. Выведена опытным путем
NumBStr=11;

type
  TForm1 = class(TForm)
    IdLPR1: TIdLPR;
    WordDocument1: TWordDocument;
    WordApplication1: TWordApplication;
    WordParagraphFormat1: TWordParagraphFormat;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Button1: TButton;
    StringGrid1: TStringGrid;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    Button2: TButton;
    SaveDialog1: TSaveDialog;
    OpenDialog1: TOpenDialog;
    OpenDialog2: TOpenDialog;
    GroupBox2: TGroupBox;
    Label5: TLabel;
    Edit1: TEdit;
    Button4: TButton;
    Label6: TLabel;
    Label4: TLabel;
    CheckBox1: TCheckBox;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Edit1Change(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

  //тип элемента массива для вывода в шаблон ворд
  type strInfoElem=record
    //номер требования или накладной. У требования номера не будет
    docNumber:string;
    //имя товара
    objName:string;
    //номенклатурный номер товара
    nomNumber:string;
    //код единиц измерения константа
    kod:string;
    // ед измерения
    kodAd:string;
    //затребовано товара шт.
    objRequer:string;
    //отпущено товара шт.
    objRequerOut:string;
    //цена одной единицы товара
    oneObjCost:string;
    //сумма без НДС .
    //заполняется для требований =objRequerOut*oneObjCost
    sumWithOutNDS:string;
  end;
var
  Form1: TForm1;

  //===================================
  //работа с Word
  //переменная ссылка на объект Excel
  PExelObj:variant;
  //переменная ссылка на объект Word
  PWordObj:variant;
  //Переменная ссылка на книгу
  PExelBookCurrent:variant;
  //переменая ссылка на активную книгу
  PExelBookActive:variant;
  //Переменная ссылка на активный лист
  PExelSheetActive:variant;
  //===================================

  //колич. заполненных столбцов в Excel файле
  rows:integer;
  //колич. заполненных строк в Excel файле
  cols:integer;

  i,j:integer;

  //строка для сохранения адресов выделеных ячеек
  strstr:string;
  sss:string;
  flagFill:boolean;
  //динимический массив строк для вывода в шаблон ворд
  masStrInfoElem: array of strInfoElem;
  iMasStrInfoElem:integer;

  docCreate:string;
  iBlanc:integer;
  FullProgPath: PChar;
  firstRun:boolean;
  //назв. прошлого открытого файла
  strFileOld:string;
implementation

{$R *.dfm}

//проверка установлен ли excel
function CheckExcelInstall:boolean;
var
  ClassID: TCLSID;
  Rez : HRESULT;
begin
// Ищем CLSID OLE-объекта
  Rez := CLSIDFromProgID(PWideChar(WideString('Excel.Application')), ClassID);
  if Rez = S_OK then  // Объект найден
    Result := true
  else
    Result := false;
end;

//проверка запущен ли Excel
function CheckExcelRun: boolean;
begin
  try
    PExelObj:=GetActiveOleObject('Excel.Application');
    Result:=True;
  except
    Result:=false
  end;
end;

procedure GenMonthYear();
var
dt: integer;
dt2: integer;
begin
//месяц и год генерации отчета
dt:=MonthOf(Now); 
dt2:=YearOf(Now);
docCreate:=IntToStr(dt)+'.'+IntToStr(dt2);
end;

procedure ReplaceLit(subStr:string;str:string);
begin
//выделили весь документ
PWordObj.ActiveDocument.Select;
//осущ. поиск по всему документу
PWordObj.Selection.Find.Forward:=true;
//подстроку которую надо искать
PWordObj.Selection.Find.Text:=subStr;
//заменить первое найденное значение нужным
if PWordObj.Selection.Find.Execute then PWordObj.Selection.Text:=str;
end;

function StrCorection(strIn:string):string;
var
i:integer;
strOut:string;
strOutCor:string;
buf:string;
flagMore:boolean;
flagE:boolean;
begin
i:=1;

//одна ячейка была выделена или несколько
flagMore:=false;
while i<=length(strIn) do
  begin
    if strIn[i]=',' then
      begin
        flagMore:=true;
        break;
      end;
    inc(i);  
  end;



i:=1;
strOut:='';
while i<=length(strIn) do
  begin
    if strIn[i]<>'$' then
      begin
        strOut:=strOut+strIn[i];
      end;
    inc(i);
  end;
//Проверим накладная или требования надо вбить в шаблон
if form1.RadioButton1.Checked then
  begin
    //требования. только D-адреса
    i:=1;
    strOutCor:=strOut;
    strOut:='';
    buf:='';
    while i<=length(strOutCor) do
      begin
        if strOutCor[i]<>',' then
          begin
            buf:=buf+strOutCor[i];
          end
        else
          begin
            if buf[1]='D' then
              begin
                strOut:=strOut+buf+',';
              end;
            buf:='';
          end;
        inc(i);
      end;
    if buf[1]='D' then
      begin
        strOut:=strOut+buf+',';
      end;
    buf:='';
  end
else
  begin
    //накладная. только B-адреса
    i:=1;
    strOutCor:=strOut;
    strOut:='';
    buf:='';

    if flagMore then
      begin
        flagE:=false;
        //выделено несколько
        while i<=length(strOutCor) do
          begin
            if strOutCor[i]<>',' then
              begin
                buf:=buf+strOutCor[i];
              end
            else
              begin
                if buf[1]='B' then
                  begin
                    if PExelSheetActive.Range[buf].Text<>'' then
                      begin
                        strOut:=strOut+buf+',';
                        flagE:=true;
                        break;
                      end;
                  end;
                buf:='';
              end;
            inc(i);
          end;
        if ((buf[1]='B') and (not flagE)) then
          begin
            if PExelSheetActive.Range[buf].Text<>'' then
              begin
                strOut:=strOut+buf+',';
              end;
          end;
        buf:='';
      end
    else
      begin
        //выделена одна
        while i<=length(strOutCor) do
          begin
            if strOutCor[i]<>#0 then
              begin
                buf:=buf+strOutCor[i];
              end;
            inc(i);
          end;
        if buf[1]='B' then
          begin
            if PExelSheetActive.Range[buf].Text<>'' then
              begin
                strOut:=strOut+buf+',';
              end;
          end;
        buf:='';
      end;


    if strOut='' then
      begin
        showMessage('Ошибка работы. Выбраны неверные значения. Программа будет завершена');
        //закрытие активной книги
        PExelBookActive.Close;
        //закрытие приложения Excel
        PExelObj.Application.Quit;
        PExelObj:=Unassigned;
        halt;
      end;
      
    //для накладной выбираем номер первого выделения. 1 ячейка с номером накладной
    i:=1;
    buf:='';
    while strOut[i]<>',' do
      begin
        buf:=buf+strOut[i];
        inc(i);
      end;
    strOut:=buf+',';
  end;
//удаление последнего символа (,) от последнего до последнего
Delete(strOut,length(strOut),length(strOut));
Result:=strOut;
end;

//заполнение инф для требований
procedure FillT(allAdrStr:string);
var
str:string;
numRecord:integer;
iStr:integer;
buf:string;
buf2:string;
iBuf:integer;

j:integer;

//временная переменная отпущенное количество
objRequerOut:integer;
//временная переменная цена за 1 штуку товара
oneObjCost:real;
begin
numRecord:=0;
iStr:=1;

//посчитаем количество записей. Количество букв в адресе
{while iStr<=length(allAdrStr) do
  begin
    //попадание от A до Z
    if ((ord(allAdrStr[iStr])>=65) and (ord(allAdrStr[iStr])<=90)) then
      begin
        inc(numRecord);
      end;
    inc(iStr)
  end;}
buf:='';
while iStr<=length(allAdrStr)+1 do
  begin
    if ((allAdrStr[iStr]=',') or (allAdrStr[iStr]=#0)) then
      begin
        //вытащим номер строки из адреса ячейки
        iBuf:=2;
        buf2:='';
        while iBuf<=length(buf) do
          begin
            if ((StrToInt(buf[iBuf])>=0) and (StrToInt(buf[iBuf])<=9)) then
              begin
                //цифра
                buf2:=buf2+buf[iBuf];
              end;
            inc(iBuf);
          end;
        j:=StrToInt(buf2);

        //Выделяем память под насчитаное записисей в массиве
        setLength(masStrInfoElem,iMasStrInfoElem+1);
        //заполняем номер требования
        masStrInfoElem[iMasStrInfoElem].docNumber:=PExelSheetActive.Range['C'+intTostr(j)].Text;
        //заполняем имя товара
        masStrInfoElem[iMasStrInfoElem].objName:=PExelSheetActive.Range['D'+intTostr(j)].Text;
        //ном. номер
        masStrInfoElem[iMasStrInfoElem].nomNumber:=PExelSheetActive.Range['H'+intTostr(j)].Text;
        //код
        masStrInfoElem[iMasStrInfoElem].kod:='08';
        //доп.
        masStrInfoElem[iMasStrInfoElem].kodAd:='шт';
        //затребовано товара
        masStrInfoElem[iMasStrInfoElem].objRequer:=PExelSheetActive.Range['E'+intTostr(j)].Text;
        //затребовано отпущено товара
        masStrInfoElem[iMasStrInfoElem].objRequerOut:=PExelSheetActive.Range['E'+intTostr(j)].Text;
        //цена за товар
        masStrInfoElem[iMasStrInfoElem].oneObjCost:=PExelSheetActive.Range['F'+intTostr(j)].Text;

        //для вычислений
        objRequerOut:=StrToInt(masStrInfoElem[iMasStrInfoElem].objRequerOut);
        oneObjCost:=StrToFloat(masStrInfoElem[iMasStrInfoElem].oneObjCost);

        //посчитаем для текущей строки сумму без НДС
        masStrInfoElem[iMasStrInfoElem].sumWithOutNDS:=
          FloatToStr(oneObjCost*objRequerOut);

        inc(iMasStrInfoElem);
        buf:='';
      end
    else
      begin
        buf:=buf+allAdrStr[iStr];
      end;
    inc(iStr);
  end;

//showMessage(intToStr(numRecord));

end;


//заполнение информации для накладной
procedure FillN(allAdrStr:string);
var
numRecord:integer;
strNac:string;
j:integer;
buf:string;
ssss:string;
jj:integer;
begin
iMasStrInfoElem:=0;
//считаем количество записей в накладной. с 1 потому что считать будем со след. строки
numRecord:=1;
//номер строки у нашего адреса
j:=2;
buf:='';
while j<=length(allAdrStr) do
  begin
    if ((StrToInt(allAdrStr[j])>=0) and (StrToInt(allAdrStr[j])<=9)) then
      begin
        //цифра
        buf:=buf+allAdrStr[j];
      end;
    inc(j);
  end;

j:=StrToInt(buf);
//переход на след. ячейку вниз по столбцу
j:=j+1;
while (true) do
  begin
    if PExelSheetActive.Range[allAdrStr[1]+intTostr(j)].Text='' then
      begin
        inc(numRecord);
        //если количество строк больше оговоренного объема, то присваиваем максимальное количество строк и выходим
        if numRecord>NumBStr then
          begin
            numRecord:=NumBStr;
            break;
          end;
      end
    else
      begin
        //выходим посчитали
        break;
      end;
    inc(j);
  end;
//showMessage(IntTostr(numRecord)); //при прогоне по шагам кажется не работает(глюк)
//Выделяем память под насчитаное записисей в массиве
setLength(masStrInfoElem,iMasStrInfoElem+numRecord);

//заполняем элементы массива
j:=StrToInt(buf);
jj:=1;
while jj<=numRecord do
  begin
    //заполняем номер накладной
    masStrInfoElem[iMasStrInfoElem].docNumber:=PExelSheetActive.Range[allAdrStr[1]+intTostr(j)].Text;
    //заполняем имя товара
    masStrInfoElem[iMasStrInfoElem].objName:=PExelSheetActive.Range['D'+intTostr(j)].Text;
    //ном. номер
    masStrInfoElem[iMasStrInfoElem].nomNumber:=PExelSheetActive.Range['H'+intTostr(j)].Text;
    //код
    masStrInfoElem[iMasStrInfoElem].kod:='08';
    //доп.
    masStrInfoElem[iMasStrInfoElem].kodAd:='шт';
    //затребовано товара
    masStrInfoElem[iMasStrInfoElem].objRequer:=PExelSheetActive.Range['E'+intTostr(j)].Text;
    //затребовано отпущено товара

    //не заполняем поле отпущено для накладных
    //masStrInfoElem[iMasStrInfoElem].objRequerOut:=PExelSheetActive.Range['E'+intTostr(j)].Text;
    //цена за товар
    masStrInfoElem[iMasStrInfoElem].oneObjCost:=PExelSheetActive.Range['F'+intTostr(j)].Text;
    inc(iMasStrInfoElem);
    inc(j);
    inc(jj);
  end;
//заполнили
end;


procedure TForm1.Button1Click(Sender: TObject);

begin

//проверим установлен ли excel
if not CheckExcelInstall then
  begin
    ShowMessage('Установленное ПО MS Excel на ПК не обнаружено. Установите его и перезапустите ПО');
    halt;
  end;

//сгенерировать месяц и год отчета.
GenMonthYear;

//при первом запуске программы запускаем Exсel c выбранным файлом
if firstRun then
  begin
    //Создание оъекта Excel.Application. Объект Excel.
    PExelObj:=CreateOleObject('Excel.Application');
    //Откроем Excel файл который нужно дополнить.
    if form1.OpenDialog1.Execute then PExelObj.WorkBooks.Open(form1.OpenDialog1.FileName,ReadOnly:=True);
    //разрешить вписывать номер счета для подсчета
    form1.Edit1.Enabled:=true;
    firstRun:=false;
    form1.Button1.Caption:='Перейти к файлу склада';
    ShowMessage('ВНИМАНИЕ!!! Количество заполняемых строк не должно превышать 11');
  end;


//Проверяем запущен ли Excel. если нет запускаем
if not CheckExcelRun then
  begin
    PExelObj:=CreateOleObject('Excel.Application');
  end;

//сделаем окно приложения Excel видимым
PExelObj.Visible:=true;

if not CheckExcelRun then
  begin
    //Откроем Excel файл который нужно дополнить.
    if form1.OpenDialog1.Execute then
      begin
        PExelObj.WorkBooks.Open(form1.OpenDialog1.FileName,ReadOnly:=True);
        strFileOld:=form1.OpenDialog1.FileName;
      end;
  end;
//получим ссылку на объект книги документа
PExelBookCurrent:=PExelObj.WorkBooks;
//получаем доступ к последней открытой книге. Делаем её активной.
PExelBookCurrent.Item[PExelObj.WorkBooks.Count].Activate;
//получение ссылки на активную книгу.
PExelBookActive:=PExelBookCurrent.Item[PExelObj.WorkBooks.Count];
//Получение ссылки на первый лист активной книги
PExelSheetActive:=PExelBookActive.Sheets.Item[1];

//активируем последнюю ячейку на активном листе.
//PExelSheetActive.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;
//Объект Selection тоже самое что объект Range. Свойство Теxt вернет стринт а Value

//Убираем предидущее выделение если оно было, выделяя ячейку А1
PExelObj.Range['A1'].Select ;

while PExelObj.Selection.Text='' do application.ProcessMessages;
form1.Button2.Enabled:=true;
form1.RadioButton2.Enabled:=false;
form1.RadioButton1.Enabled:=false;
form1.Button1.Enabled:=false;

while not flagFill do application.ProcessMessages;
//выд. адреса
strstr:=PExelObj.Selection.Address;
//коррекция выд. адресов
sss:=StrCorection(strstr);

//Если мы нажимали не туда куда надо то ошибка
if sss='' then
  begin
    showMessage('Ошибка работы. Выбраны неверные значения. Программа будет завершена');
    //закрытие активной книги
    PExelBookActive.Close;
    //закрытие приложения Excel
    PExelObj.Application.Quit;
    PExelObj:=Unassigned;
    halt;
  end;


//Проверяем накладная выбрана или требование. Заполняем
if form1.RadioButton1.Checked then
  begin
    //требования
    FillT(sss);
  end
else
  begin
    //накладная
    FillN(sss);
  end;


PWordObj:=CreateOleObject('Word.Application');
//невидимое окно ворд
PWordObj.Visible:=false;
PWordObj.Documents.Open(ExtractFileDir(ParamStr(0))+'\template'+'.doc');

if form1.RadioButton1.Checked then
  begin
    //требования
    //запускаем процедуру замены литеры на необходимое значение
    ReplaceLit('AAA','ТРЕБОВАНИЕ №');
    ReplaceLit('AAA','ТРЕБОВАНИЕ №');

    //заполним отправителя как получателя для требований
    //1 бланк
    ReplaceLit('BBB','0019');
    ReplaceLit('BBB','307');
    //2 бланк
    ReplaceLit('BBB','0019');
    ReplaceLit('BBB','307');
  end
else
  begin
    //накладная
    ReplaceLit('AAA','НАКЛАДНАЯ №');
    ReplaceLit('AAA','НАКЛАДНАЯ №');

    //заполним отправителя как пустое поле
    //1 бланк
    ReplaceLit('BBB','');
    ReplaceLit('BBB','');
    //2 бланк
    ReplaceLit('BBB','');
    ReplaceLit('BBB','');

  end;

//вставка номера дока
ReplaceLit('&num',masStrInfoElem[0].docNumber);
ReplaceLit('&num',masStrInfoElem[0].docNumber);

//вставка даты дока
//в зависимости от того нужно ставить дату или нет
if (form1.CheckBox1.Checked) then
  begin
    ReplaceLit('&tim',docCreate);
    ReplaceLit('&tim',docCreate);
  end
else
  begin
    ReplaceLit('&tim','');
    ReplaceLit('&tim','');
  end;


//заполняем первый бланк
iBlanc:=1;
while  iBlanc<=NumBStr do
  begin
    if iBlanc<=length(masStrInfoElem) then
      begin
        //наименование
        ReplaceLit('&nameD1',masStrInfoElem[iBlanc-1].objName);
        //номенклатурный
        ReplaceLit('&nomD1',masStrInfoElem[iBlanc-1].nomNumber);
        //код
        ReplaceLit('&kodD1',masStrInfoElem[iBlanc-1].kod);
        //код доп
        ReplaceLit('&kAD1',masStrInfoElem[iBlanc-1].kodAd);
        //затребованное количество
        ReplaceLit('&recD1',masStrInfoElem[iBlanc-1].objRequer);
        //отпущеное количество
        ReplaceLit('&recOD1',masStrInfoElem[iBlanc-1].objRequerOut);
        //цена за штуку
        ReplaceLit('&costD1',masStrInfoElem[iBlanc-1].oneObjCost);
        //сумма без НДС
        ReplaceLit('&cswNDS1',masStrInfoElem[iBlanc-1].sumWithOutNDS);
      end
    else
      begin
        //наименование
        ReplaceLit('&nameD1','');
        //номенклатурный
        ReplaceLit('&nomD1','');
        //код
        ReplaceLit('&kodD1','');
        //код доп
        ReplaceLit('&kAD1','');
        //затребованное количество
        ReplaceLit('&recD1','');
        //отпущеное количество
        ReplaceLit('&recOD1','');
        //цена за штуку
        ReplaceLit('&costD1','');
        //сумма без НДС
        ReplaceLit('&cswNDS1','');
      end;
    inc(iBlanc);
  end;

//заполняем второй бланк
iBlanc:=1;
while  iBlanc<=NumBStr do
  begin
    if iBlanc<=length(masStrInfoElem) then
      begin
        //наименование
        ReplaceLit('&nameD2',masStrInfoElem[iBlanc-1].objName);
        //номенклатурный
        ReplaceLit('&nomD2',masStrInfoElem[iBlanc-1].nomNumber);
        //код
        ReplaceLit('&kodD2',masStrInfoElem[iBlanc-1].kod);
        //код доп
        ReplaceLit('&kAD2',masStrInfoElem[iBlanc-1].kodAd);
        //затребованное количество
        ReplaceLit('&recD2',masStrInfoElem[iBlanc-1].objRequer);
        //отпущеное количество
        ReplaceLit('&recOD2',masStrInfoElem[iBlanc-1].objRequerOut);
        //цена за штуку
        ReplaceLit('&costD2',masStrInfoElem[iBlanc-1].oneObjCost);
        //сумма без НДС
        ReplaceLit('&cswNDS2',masStrInfoElem[iBlanc-1].sumWithOutNDS);
      end
    else
      begin
        //наименование
        ReplaceLit('&nameD2','');
        //номенклатурный
        ReplaceLit('&nomD2','');
        //код
        ReplaceLit('&kodD2','');
        //код доп
        ReplaceLit('&kAD2','');
        //затребованное количество
        ReplaceLit('&recD2','');
        //отпущеное количество
        ReplaceLit('&recOD2','');
        //цена за штуку
        ReplaceLit('&costD2','');
        //сумма без НДС
        ReplaceLit('&cswNDS2','');
      end;
    inc(iBlanc);
  end;

//видимое окно ворд
PWordObj.Visible:=true;

//сохранение активной книги под новым именем
if form1.SaveDialog1.Execute then
  begin
    PWordObj.ActiveDocument.SaveAs(form1.SaveDialog1.FileName+'.doc');
    PWordObj.ActiveDocument.Close(True); // сохраняем и закрываем Word
    //окончательно закрываем ворд
    PWordObj.Quit;
    //освобождаем память выделенную для ворда
    PWordObj:=UnAssigned;
  end
else
  begin
    //отмена
    PWordObj.ActiveDocument.Close;
    PWordObj.Quit;
    PWordObj:=UnAssigned;
  end;


form1.RadioButton1.Checked:=true;
form1.RadioButton1.Enabled:=true;
form1.RadioButton2.Enabled:=true;
form1.Button2.Enabled:=false;
form1.Button1.Enabled:=true;
flagFill:=false;
//освобождаем массив строк
masStrInfoElem:=nil;
iMasStrInfoElem:=0;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
firstRun:=true;
flagFill:=false;
form1.Button4.Enabled:=false;
form1.RadioButton1.Checked:=true;
form1.RadioButton1.Enabled:=true;
form1.RadioButton2.Enabled:=true;
form1.Button2.Enabled:=false;
form1.Button4.Enabled:=false;
form1.Edit1.Enabled:=false;

//по умолчанию дату не ставим
form1.CheckBox1.Checked:=false;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
//флаг для отлавливания выделеных ячеек
flagFill:=true;
end;

procedure TForm1.Button3Click(Sender: TObject);
var
//счетчик для перебора стобцов в документе
iMaxRows:integer;
//счетчик для перебора строк в документе
iMaxCols:integer;
//вспомогательная строка
srsr:string;
srsr2:string;
//переменные для работы с объектом Excel
PExelObj2:variant;
PExelBookCurrent2:variant;
PExelBookActive2:variant;
PExelSheetActive2:variant;

//счетчик перебора активных строк
iFind:integer;
//аккумулятор общей суммы заказа
orderAcum:real;

flag:boolean;
begin
//перессылаемся на уже запущеный процесс Excel
PExelObj2:=PExelObj;
//получим ссылку на объект книги документа
PExelBookCurrent2:=PExelObj2.WorkBooks;
//получаем доступ к последней открытой книге. Делаем её активной.
PExelBookCurrent2.Item[PExelObj2.WorkBooks.Count].Activate;
//получение ссылки на активную книгу.
PExelBookActive2:=PExelBookCurrent2.Item[PExelObj2.WorkBooks.Count];
//Получение ссылки на первый лист активной книги
PExelSheetActive2:=PExelBookActive2.Sheets.Item[1];
//флаг для возможности правильного подсчета строк и столбов,
//с учетом что может быть пустая строка
flag:=false;

iMaxRows:=2;
rows:=0;
while (not flag) do
  begin
    //высчитываем количество заполненных столбцов. Ориентируемся по заголовку таблицы в строке 2
    while PExelSheetActive2.Cells[2,iMaxRows].Text<>'' do        // строка,ряд
      begin
        srsr:=PExelSheetActive2.Cells[2,iMaxRows].Text;
        inc(iMaxRows);
        //заполненое количество столбцов
        inc(rows);
        flag:=true;
      end;
    //проверка на случай пустой строки
    if ((PExelSheetActive2.Cells[2,iMaxRows+1].Text<>'')and(flag)) then
      begin
        inc(iMaxRows);
        inc(rows);
        srsr:=PExelSheetActive2.Cells[2,iMaxRows+1].Text;
        flag:=false;
      end
    else
      begin
        flag:=true;
      end;
  end;


flag:=false;
iMaxCols:=3;
cols:=0;

while (not flag) do
  begin
    //высчитываем количество заполненных строк. Ориентируемся по столбцу суммы
    while PExelSheetActive2.Cells[iMaxCols,7].Text<>'' do        // строка,ряд
      begin
        srsr:=PExelSheetActive2.Cells[iMaxCols,7].Text;
        inc(iMaxCols);
        //заполненое количество строк
        inc(cols);
        flag:=true;
      end;
    //проверка на случай пустой строки
    if ((PExelSheetActive2.Cells[iMaxCols+1,7].Text<>'')and(flag)) then
      begin
        srsr:=PExelSheetActive2.Cells[iMaxCols+1,7].Text;
        inc(iMaxCols);
        inc(cols);
        flag:=false;
      end
    else
      begin
        flag:=true;
      end;
  end;


//если пуcтой номер заказа то выводим сообщение
if form1.Edit1.Text='' then
  begin
    form1.Label4.Caption:='Неверный номер заказа';
  end
else
  begin
     iFind:=1;
     orderAcum:=0.0;
     // перебирием все заполненые строки и считаем сумму
     while iFind<=cols do
      begin
        srsr:=PExelSheetActive2.Range['M'+intTostr(iFind+2)].Text;
        if PExelSheetActive2.Range['M'+intTostr(iFind+2)].Text=form1.Edit1.Text then
          begin
            {if iFind=272 then
              begin
              srsr:=PExelSheetActive2.Range['F'+intTostr(iFind+2)].Text;
              srsr2:=PExelSheetActive2.Range['E'+intTostr(iFind+2)].Text;
              end;}
            //= +цена за 1 наименование*количество заказаных
            orderAcum:=orderAcum+(StrToFloat(PExelSheetActive2.Range['F'+intTostr(iFind+2)].Text)*
            StrToFloat(PExelSheetActive2.Range['E'+intTostr(iFind+2)].Text));
          end;
        inc(iFind);
      end;

     //вывод результата
     form1.Label4.Caption:=FloatToStr(orderAcum);
     orderAcum:=0;
  end;

//закрытие активной книги
//PExelBookActive2.Close;
//закрытие приложения Excel
//PExelObj2.Application.Quit;
//PExelObj2:=Unassigned;


end;

procedure TForm1.Edit1Change(Sender: TObject);
begin
form1.Button4.Enabled:=true;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin



//Проверим запущен ли Excel. Не закрыли ли его до этого
{if (CheckExcelRun) then
  begin
    //завершаем работу с Excel
    PExelBookActive.Close;
    PExelBookCurrent.close;
    //закрытие приложения Excel
    PExelObj.Quit;
    PExelObj:=Unassigned;
  end;  }

halt;
end;

end.
