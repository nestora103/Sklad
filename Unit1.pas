unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComObj, ActiveX, StdCtrls, Grids, OleServer, ExcelXP,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdLPR, Menus,
  OleCtnrs, WordXP,DateUtils,xpman;
const
//��������� ����������� ����� 1 ���������. �������� ������� �����
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

  //��� �������� ������� ��� ������ � ������ ����
  type strInfoElem=record
    //����� ���������� ��� ���������. � ���������� ������ �� �����
    docNumber:string;
    //��� ������
    objName:string;
    //�������������� ����� ������
    nomNumber:string;
    //��� ������ ��������� ���������
    kod:string;
    // �� ���������
    kodAd:string;
    //����������� ������ ��.
    objRequer:string;
    //�������� ������ ��.
    objRequerOut:string;
    //���� ����� ������� ������
    oneObjCost:string;
    //����� ��� ��� .
    //����������� ��� ���������� =objRequerOut*oneObjCost
    sumWithOutNDS:string;
  end;
var
  Form1: TForm1;

  //===================================
  //������ � Word
  //���������� ������ �� ������ Excel
  PExelObj:variant;
  //���������� ������ �� ������ Word
  PWordObj:variant;
  //���������� ������ �� �����
  PExelBookCurrent:variant;
  //��������� ������ �� �������� �����
  PExelBookActive:variant;
  //���������� ������ �� �������� ����
  PExelSheetActive:variant;
  //===================================

  //�����. ����������� �������� � Excel �����
  rows:integer;
  //�����. ����������� ����� � Excel �����
  cols:integer;

  i,j:integer;

  //������ ��� ���������� ������� ��������� �����
  strstr:string;
  sss:string;
  flagFill:boolean;
  //������������ ������ ����� ��� ������ � ������ ����
  masStrInfoElem: array of strInfoElem;
  iMasStrInfoElem:integer;

  docCreate:string;
  iBlanc:integer;
  FullProgPath: PChar;
  firstRun:boolean;
  //����. �������� ��������� �����
  strFileOld:string;
implementation

{$R *.dfm}

//�������� ���������� �� excel
function CheckExcelInstall:boolean;
var
  ClassID: TCLSID;
  Rez : HRESULT;
begin
// ���� CLSID OLE-�������
  Rez := CLSIDFromProgID(PWideChar(WideString('Excel.Application')), ClassID);
  if Rez = S_OK then  // ������ ������
    Result := true
  else
    Result := false;
end;

//�������� ������� �� Excel
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
//����� � ��� ��������� ������
dt:=MonthOf(Now); 
dt2:=YearOf(Now);
docCreate:=IntToStr(dt)+'.'+IntToStr(dt2);
end;

procedure ReplaceLit(subStr:string;str:string);
begin
//�������� ���� ��������
PWordObj.ActiveDocument.Select;
//����. ����� �� ����� ���������
PWordObj.Selection.Find.Forward:=true;
//��������� ������� ���� ������
PWordObj.Selection.Find.Text:=subStr;
//�������� ������ ��������� �������� ������
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

//���� ������ ���� �������� ��� ���������
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
//�������� ��������� ��� ���������� ���� ����� � ������
if form1.RadioButton1.Checked then
  begin
    //����������. ������ D-������
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
    //���������. ������ B-������
    i:=1;
    strOutCor:=strOut;
    strOut:='';
    buf:='';

    if flagMore then
      begin
        flagE:=false;
        //�������� ���������
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
        //�������� ����
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
        showMessage('������ ������. ������� �������� ��������. ��������� ����� ���������');
        //�������� �������� �����
        PExelBookActive.Close;
        //�������� ���������� Excel
        PExelObj.Application.Quit;
        PExelObj:=Unassigned;
        halt;
      end;
      
    //��� ��������� �������� ����� ������� ���������. 1 ������ � ������� ���������
    i:=1;
    buf:='';
    while strOut[i]<>',' do
      begin
        buf:=buf+strOut[i];
        inc(i);
      end;
    strOut:=buf+',';
  end;
//�������� ���������� ������� (,) �� ���������� �� ����������
Delete(strOut,length(strOut),length(strOut));
Result:=strOut;
end;

//���������� ��� ��� ����������
procedure FillT(allAdrStr:string);
var
str:string;
numRecord:integer;
iStr:integer;
buf:string;
buf2:string;
iBuf:integer;

j:integer;

//��������� ���������� ���������� ����������
objRequerOut:integer;
//��������� ���������� ���� �� 1 ����� ������
oneObjCost:real;
begin
numRecord:=0;
iStr:=1;

//��������� ���������� �������. ���������� ���� � ������
{while iStr<=length(allAdrStr) do
  begin
    //��������� �� A �� Z
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
        //������� ����� ������ �� ������ ������
        iBuf:=2;
        buf2:='';
        while iBuf<=length(buf) do
          begin
            if ((StrToInt(buf[iBuf])>=0) and (StrToInt(buf[iBuf])<=9)) then
              begin
                //�����
                buf2:=buf2+buf[iBuf];
              end;
            inc(iBuf);
          end;
        j:=StrToInt(buf2);

        //�������� ������ ��� ���������� ��������� � �������
        setLength(masStrInfoElem,iMasStrInfoElem+1);
        //��������� ����� ����������
        masStrInfoElem[iMasStrInfoElem].docNumber:=PExelSheetActive.Range['C'+intTostr(j)].Text;
        //��������� ��� ������
        masStrInfoElem[iMasStrInfoElem].objName:=PExelSheetActive.Range['D'+intTostr(j)].Text;
        //���. �����
        masStrInfoElem[iMasStrInfoElem].nomNumber:=PExelSheetActive.Range['H'+intTostr(j)].Text;
        //���
        masStrInfoElem[iMasStrInfoElem].kod:='08';
        //���.
        masStrInfoElem[iMasStrInfoElem].kodAd:='��';
        //����������� ������
        masStrInfoElem[iMasStrInfoElem].objRequer:=PExelSheetActive.Range['E'+intTostr(j)].Text;
        //����������� �������� ������
        masStrInfoElem[iMasStrInfoElem].objRequerOut:=PExelSheetActive.Range['E'+intTostr(j)].Text;
        //���� �� �����
        masStrInfoElem[iMasStrInfoElem].oneObjCost:=PExelSheetActive.Range['F'+intTostr(j)].Text;

        //��� ����������
        objRequerOut:=StrToInt(masStrInfoElem[iMasStrInfoElem].objRequerOut);
        oneObjCost:=StrToFloat(masStrInfoElem[iMasStrInfoElem].oneObjCost);

        //��������� ��� ������� ������ ����� ��� ���
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


//���������� ���������� ��� ���������
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
//������� ���������� ������� � ���������. � 1 ������ ��� ������� ����� �� ����. ������
numRecord:=1;
//����� ������ � ������ ������
j:=2;
buf:='';
while j<=length(allAdrStr) do
  begin
    if ((StrToInt(allAdrStr[j])>=0) and (StrToInt(allAdrStr[j])<=9)) then
      begin
        //�����
        buf:=buf+allAdrStr[j];
      end;
    inc(j);
  end;

j:=StrToInt(buf);
//������� �� ����. ������ ���� �� �������
j:=j+1;
while (true) do
  begin
    if PExelSheetActive.Range[allAdrStr[1]+intTostr(j)].Text='' then
      begin
        inc(numRecord);
        //���� ���������� ����� ������ ������������ ������, �� ����������� ������������ ���������� ����� � �������
        if numRecord>NumBStr then
          begin
            numRecord:=NumBStr;
            break;
          end;
      end
    else
      begin
        //������� ���������
        break;
      end;
    inc(j);
  end;
//showMessage(IntTostr(numRecord)); //��� ������� �� ����� ������� �� ��������(����)
//�������� ������ ��� ���������� ��������� � �������
setLength(masStrInfoElem,iMasStrInfoElem+numRecord);

//��������� �������� �������
j:=StrToInt(buf);
jj:=1;
while jj<=numRecord do
  begin
    //��������� ����� ���������
    masStrInfoElem[iMasStrInfoElem].docNumber:=PExelSheetActive.Range[allAdrStr[1]+intTostr(j)].Text;
    //��������� ��� ������
    masStrInfoElem[iMasStrInfoElem].objName:=PExelSheetActive.Range['D'+intTostr(j)].Text;
    //���. �����
    masStrInfoElem[iMasStrInfoElem].nomNumber:=PExelSheetActive.Range['H'+intTostr(j)].Text;
    //���
    masStrInfoElem[iMasStrInfoElem].kod:='08';
    //���.
    masStrInfoElem[iMasStrInfoElem].kodAd:='��';
    //����������� ������
    masStrInfoElem[iMasStrInfoElem].objRequer:=PExelSheetActive.Range['E'+intTostr(j)].Text;
    //����������� �������� ������

    //�� ��������� ���� �������� ��� ���������
    //masStrInfoElem[iMasStrInfoElem].objRequerOut:=PExelSheetActive.Range['E'+intTostr(j)].Text;
    //���� �� �����
    masStrInfoElem[iMasStrInfoElem].oneObjCost:=PExelSheetActive.Range['F'+intTostr(j)].Text;
    inc(iMasStrInfoElem);
    inc(j);
    inc(jj);
  end;
//���������
end;


procedure TForm1.Button1Click(Sender: TObject);

begin

//�������� ���������� �� excel
if not CheckExcelInstall then
  begin
    ShowMessage('������������� �� MS Excel �� �� �� ����������. ���������� ��� � ������������� ��');
    halt;
  end;

//������������� ����� � ��� ������.
GenMonthYear;

//��� ������ ������� ��������� ��������� Ex�el c ��������� ������
if firstRun then
  begin
    //�������� ������ Excel.Application. ������ Excel.
    PExelObj:=CreateOleObject('Excel.Application');
    //������� Excel ���� ������� ����� ���������.
    if form1.OpenDialog1.Execute then PExelObj.WorkBooks.Open(form1.OpenDialog1.FileName,ReadOnly:=True);
    //��������� ��������� ����� ����� ��� ��������
    form1.Edit1.Enabled:=true;
    firstRun:=false;
    form1.Button1.Caption:='������� � ����� ������';
    ShowMessage('��������!!! ���������� ����������� ����� �� ������ ��������� 11');
  end;


//��������� ������� �� Excel. ���� ��� ���������
if not CheckExcelRun then
  begin
    PExelObj:=CreateOleObject('Excel.Application');
  end;

//������� ���� ���������� Excel �������
PExelObj.Visible:=true;

if not CheckExcelRun then
  begin
    //������� Excel ���� ������� ����� ���������.
    if form1.OpenDialog1.Execute then
      begin
        PExelObj.WorkBooks.Open(form1.OpenDialog1.FileName,ReadOnly:=True);
        strFileOld:=form1.OpenDialog1.FileName;
      end;
  end;
//������� ������ �� ������ ����� ���������
PExelBookCurrent:=PExelObj.WorkBooks;
//�������� ������ � ��������� �������� �����. ������ � ��������.
PExelBookCurrent.Item[PExelObj.WorkBooks.Count].Activate;
//��������� ������ �� �������� �����.
PExelBookActive:=PExelBookCurrent.Item[PExelObj.WorkBooks.Count];
//��������� ������ �� ������ ���� �������� �����
PExelSheetActive:=PExelBookActive.Sheets.Item[1];

//���������� ��������� ������ �� �������� �����.
//PExelSheetActive.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;
//������ Selection ���� ����� ��� ������ Range. �������� ��xt ������ ������ � Value

//������� ���������� ��������� ���� ��� ����, ������� ������ �1
PExelObj.Range['A1'].Select ;

while PExelObj.Selection.Text='' do application.ProcessMessages;
form1.Button2.Enabled:=true;
form1.RadioButton2.Enabled:=false;
form1.RadioButton1.Enabled:=false;
form1.Button1.Enabled:=false;

while not flagFill do application.ProcessMessages;
//���. ������
strstr:=PExelObj.Selection.Address;
//��������� ���. �������
sss:=StrCorection(strstr);

//���� �� �������� �� ���� ���� ���� �� ������
if sss='' then
  begin
    showMessage('������ ������. ������� �������� ��������. ��������� ����� ���������');
    //�������� �������� �����
    PExelBookActive.Close;
    //�������� ���������� Excel
    PExelObj.Application.Quit;
    PExelObj:=Unassigned;
    halt;
  end;


//��������� ��������� ������� ��� ����������. ���������
if form1.RadioButton1.Checked then
  begin
    //����������
    FillT(sss);
  end
else
  begin
    //���������
    FillN(sss);
  end;


PWordObj:=CreateOleObject('Word.Application');
//��������� ���� ����
PWordObj.Visible:=false;
PWordObj.Documents.Open(ExtractFileDir(ParamStr(0))+'\template'+'.doc');

if form1.RadioButton1.Checked then
  begin
    //����������
    //��������� ��������� ������ ������ �� ����������� ��������
    ReplaceLit('AAA','���������� �');
    ReplaceLit('AAA','���������� �');

    //�������� ����������� ��� ���������� ��� ����������
    //1 �����
    ReplaceLit('BBB','0019');
    ReplaceLit('BBB','307');
    //2 �����
    ReplaceLit('BBB','0019');
    ReplaceLit('BBB','307');
  end
else
  begin
    //���������
    ReplaceLit('AAA','��������� �');
    ReplaceLit('AAA','��������� �');

    //�������� ����������� ��� ������ ����
    //1 �����
    ReplaceLit('BBB','');
    ReplaceLit('BBB','');
    //2 �����
    ReplaceLit('BBB','');
    ReplaceLit('BBB','');

  end;

//������� ������ ����
ReplaceLit('&num',masStrInfoElem[0].docNumber);
ReplaceLit('&num',masStrInfoElem[0].docNumber);

//������� ���� ����
//� ����������� �� ���� ����� ������� ���� ��� ���
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


//��������� ������ �����
iBlanc:=1;
while  iBlanc<=NumBStr do
  begin
    if iBlanc<=length(masStrInfoElem) then
      begin
        //������������
        ReplaceLit('&nameD1',masStrInfoElem[iBlanc-1].objName);
        //��������������
        ReplaceLit('&nomD1',masStrInfoElem[iBlanc-1].nomNumber);
        //���
        ReplaceLit('&kodD1',masStrInfoElem[iBlanc-1].kod);
        //��� ���
        ReplaceLit('&kAD1',masStrInfoElem[iBlanc-1].kodAd);
        //������������� ����������
        ReplaceLit('&recD1',masStrInfoElem[iBlanc-1].objRequer);
        //��������� ����������
        ReplaceLit('&recOD1',masStrInfoElem[iBlanc-1].objRequerOut);
        //���� �� �����
        ReplaceLit('&costD1',masStrInfoElem[iBlanc-1].oneObjCost);
        //����� ��� ���
        ReplaceLit('&cswNDS1',masStrInfoElem[iBlanc-1].sumWithOutNDS);
      end
    else
      begin
        //������������
        ReplaceLit('&nameD1','');
        //��������������
        ReplaceLit('&nomD1','');
        //���
        ReplaceLit('&kodD1','');
        //��� ���
        ReplaceLit('&kAD1','');
        //������������� ����������
        ReplaceLit('&recD1','');
        //��������� ����������
        ReplaceLit('&recOD1','');
        //���� �� �����
        ReplaceLit('&costD1','');
        //����� ��� ���
        ReplaceLit('&cswNDS1','');
      end;
    inc(iBlanc);
  end;

//��������� ������ �����
iBlanc:=1;
while  iBlanc<=NumBStr do
  begin
    if iBlanc<=length(masStrInfoElem) then
      begin
        //������������
        ReplaceLit('&nameD2',masStrInfoElem[iBlanc-1].objName);
        //��������������
        ReplaceLit('&nomD2',masStrInfoElem[iBlanc-1].nomNumber);
        //���
        ReplaceLit('&kodD2',masStrInfoElem[iBlanc-1].kod);
        //��� ���
        ReplaceLit('&kAD2',masStrInfoElem[iBlanc-1].kodAd);
        //������������� ����������
        ReplaceLit('&recD2',masStrInfoElem[iBlanc-1].objRequer);
        //��������� ����������
        ReplaceLit('&recOD2',masStrInfoElem[iBlanc-1].objRequerOut);
        //���� �� �����
        ReplaceLit('&costD2',masStrInfoElem[iBlanc-1].oneObjCost);
        //����� ��� ���
        ReplaceLit('&cswNDS2',masStrInfoElem[iBlanc-1].sumWithOutNDS);
      end
    else
      begin
        //������������
        ReplaceLit('&nameD2','');
        //��������������
        ReplaceLit('&nomD2','');
        //���
        ReplaceLit('&kodD2','');
        //��� ���
        ReplaceLit('&kAD2','');
        //������������� ����������
        ReplaceLit('&recD2','');
        //��������� ����������
        ReplaceLit('&recOD2','');
        //���� �� �����
        ReplaceLit('&costD2','');
        //����� ��� ���
        ReplaceLit('&cswNDS2','');
      end;
    inc(iBlanc);
  end;

//������� ���� ����
PWordObj.Visible:=true;

//���������� �������� ����� ��� ����� ������
if form1.SaveDialog1.Execute then
  begin
    PWordObj.ActiveDocument.SaveAs(form1.SaveDialog1.FileName+'.doc');
    PWordObj.ActiveDocument.Close(True); // ��������� � ��������� Word
    //������������ ��������� ����
    PWordObj.Quit;
    //����������� ������ ���������� ��� �����
    PWordObj:=UnAssigned;
  end
else
  begin
    //������
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
//����������� ������ �����
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

//�� ��������� ���� �� ������
form1.CheckBox1.Checked:=false;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
//���� ��� ������������ ��������� �����
flagFill:=true;
end;

procedure TForm1.Button3Click(Sender: TObject);
var
//������� ��� �������� ������� � ���������
iMaxRows:integer;
//������� ��� �������� ����� � ���������
iMaxCols:integer;
//��������������� ������
srsr:string;
srsr2:string;
//���������� ��� ������ � �������� Excel
PExelObj2:variant;
PExelBookCurrent2:variant;
PExelBookActive2:variant;
PExelSheetActive2:variant;

//������� �������� �������� �����
iFind:integer;
//����������� ����� ����� ������
orderAcum:real;

flag:boolean;
begin
//������������� �� ��� ��������� ������� Excel
PExelObj2:=PExelObj;
//������� ������ �� ������ ����� ���������
PExelBookCurrent2:=PExelObj2.WorkBooks;
//�������� ������ � ��������� �������� �����. ������ � ��������.
PExelBookCurrent2.Item[PExelObj2.WorkBooks.Count].Activate;
//��������� ������ �� �������� �����.
PExelBookActive2:=PExelBookCurrent2.Item[PExelObj2.WorkBooks.Count];
//��������� ������ �� ������ ���� �������� �����
PExelSheetActive2:=PExelBookActive2.Sheets.Item[1];
//���� ��� ����������� ����������� �������� ����� � �������,
//� ������ ��� ����� ���� ������ ������
flag:=false;

iMaxRows:=2;
rows:=0;
while (not flag) do
  begin
    //����������� ���������� ����������� ��������. ������������� �� ��������� ������� � ������ 2
    while PExelSheetActive2.Cells[2,iMaxRows].Text<>'' do        // ������,���
      begin
        srsr:=PExelSheetActive2.Cells[2,iMaxRows].Text;
        inc(iMaxRows);
        //���������� ���������� ��������
        inc(rows);
        flag:=true;
      end;
    //�������� �� ������ ������ ������
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
    //����������� ���������� ����������� �����. ������������� �� ������� �����
    while PExelSheetActive2.Cells[iMaxCols,7].Text<>'' do        // ������,���
      begin
        srsr:=PExelSheetActive2.Cells[iMaxCols,7].Text;
        inc(iMaxCols);
        //���������� ���������� �����
        inc(cols);
        flag:=true;
      end;
    //�������� �� ������ ������ ������
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


//���� ��c��� ����� ������ �� ������� ���������
if form1.Edit1.Text='' then
  begin
    form1.Label4.Caption:='�������� ����� ������';
  end
else
  begin
     iFind:=1;
     orderAcum:=0.0;
     // ���������� ��� ���������� ������ � ������� �����
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
            //= +���� �� 1 ������������*���������� ���������
            orderAcum:=orderAcum+(StrToFloat(PExelSheetActive2.Range['F'+intTostr(iFind+2)].Text)*
            StrToFloat(PExelSheetActive2.Range['E'+intTostr(iFind+2)].Text));
          end;
        inc(iFind);
      end;

     //����� ����������
     form1.Label4.Caption:=FloatToStr(orderAcum);
     orderAcum:=0;
  end;

//�������� �������� �����
//PExelBookActive2.Close;
//�������� ���������� Excel
//PExelObj2.Application.Quit;
//PExelObj2:=Unassigned;


end;

procedure TForm1.Edit1Change(Sender: TObject);
begin
form1.Button4.Enabled:=true;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin



//�������� ������� �� Excel. �� ������� �� ��� �� �����
{if (CheckExcelRun) then
  begin
    //��������� ������ � Excel
    PExelBookActive.Close;
    PExelBookCurrent.close;
    //�������� ���������� Excel
    PExelObj.Quit;
    PExelObj:=Unassigned;
  end;  }

halt;
end;

end.
