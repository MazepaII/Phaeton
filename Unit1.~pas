//КООРДИНАТЫ СОЗДАТЕЛЯ ПРОГРАММЫ
//Электронная почта: Ltlvfpfqbpfzws@yandex.ru
unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, math, XPMaN, ExtCtrls, jpeg, scale, Menus, Grids,
  ActiveX, Printers, UTCLStringGrid, ComCtrls, Buttons,ComObj;
    const
/// Обводка ящеек в Экселе
  xlDiagonalDown = 5;
  xlDiagonalUp = 6;
  xlEdgeBottom = 9;
  xlEdgeLeft = 7;
  xlEdgeRight = 10;
  xlEdgeTop = 8;
  xlInsideHorizontal = 12;
  xlInsideVertical = 11;
//Состояние окна Excel:
  xlMaximized = $FFFFEFD7 ;//(или -4137)  развернуть Excel на весь экран.
  xlNormal = $FFFFEFD1    ;//(или -4143) - восстановить Excel.
  xlMinimized = $FFFFEFD4 ;//(или -4140) - свернуть Excel на панель задач.

type
  TForm1 = class(TForm)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    Label1: TLabel;
    Shape_IS_OBV: TShape;
    CLStringGrid1: TCLStringGrid;
    CLStringGrid2: TCLStringGrid;
    Panel7: TPanel;
    Panel8: TPanel;
    Panel9: TPanel;
    Panel10: TPanel;
    Label42: TLabel;
    Panel11: TPanel;
    Label43: TLabel;
    Label44: TLabel;
    Panel12: TPanel;
    Label45: TLabel;
    Panel_IS_H: TPanel;
    Shape1: TShape;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N5: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    TabSheet2: TTabSheet;
    Label57: TLabel;
    Shape_P_OBV: TShape;
    CLStringGrid3: TCLStringGrid;
    Panel13: TPanel;
    Label47: TLabel;
    Label48: TLabel;
    CLStringGrid4: TCLStringGrid;
    Panel_P_H: TPanel;
    Shape2: TShape;
    Label28: TLabel;
    Label25: TLabel;
    Button1: TButton;
    Panel_Button_2: TPanel;
    Button2: TButton;
    Panel14: TPanel;
    Label2: TLabel;
    Panel15: TPanel;
    Label4: TLabel;
    Panel18: TPanel;
    Label6: TLabel;
    Label7: TLabel;
    Panel17: TPanel;
    Label8: TLabel;
    Label9: TLabel;
    Panel16: TPanel;
    Label10: TLabel;
    Label16: TLabel;
    Label3: TLabel;
    Panel19: TPanel;
    Label5: TLabel;
    Panel20: TPanel;
    Label33: TLabel;
    Label34: TLabel;
    Label41: TLabel;
    Label11: TLabel;
    Panel21: TPanel;
    Label12: TLabel;
    Panel22: TPanel;
    Label13: TLabel;
    Image_save_1: TImage;
    PrintDialog1: TPrintDialog;
    SaveDialog2: TSaveDialog;
    SaveDialog1: TSaveDialog;
    Image_save_2: TImage;
    N6: TMenuItem;
    N7: TMenuItem;
    M_Excel: TMenuItem;
    M_Bloknot: TMenuItem;
    Shape_KPACH_2: TShape;
    Image_printer_1: TImage;
    Image_printer_2: TImage;
    Shape_KPACH_1: TShape;
    printer_1: TImage;
    Save_1: TImage;
    Save_2: TImage;
    Printer_2: TImage;
    Memo1: TMemo;
    Label14: TLabel;
    Label15: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Button1_Cxem: TButton;
    Button2_Cxem: TButton;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure CLStringGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure CLStringGrid3DrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure CLStringGrid2DrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure CLStringGrid4DrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure CLStringGrid2KeyPress(Sender: TObject; var Key: Char);
    procedure SetMultiLineButton(AParent: TWinControl) ;
    procedure Button2Click(Sender: TObject);
    procedure Button2MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure CLStringGrid2KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure CLStringGrid4KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Save_1Click(Sender: TObject);
    procedure M_ExcelClick(Sender: TObject);
    procedure M_BloknotClick(Sender: TObject);
    procedure printer_1Click(Sender: TObject);
    procedure Button1_CxemClick(Sender: TObject); // Делаем многострочный Caption у Button

  private
function UpDownSymbol(RangeH: String;StartUp, LengthUp, StartDown,LengthDown: integer ):boolean;
function ExcelPamka(Pamka: String;UP,Down,Left,Right: integer): integer ;  //Рамка Excel
function RoundSignificant(num: Extended; col: integer): Extended;  //Округлить
function IsOLEObjectInstalled(Name: String): boolean; //Установлен ли Excel
function Excel_Raschet: boolean; //Excel РАСЧЕТ         ([([([(((ПРОЦИДУРА)]))])])
function Excel_Exit: boolean; //ЗАКРЫТЬ Excel           ([([([(((ПРОЦИДУРА)]))])])
function Memo1_Raschet: boolean; // Блокнот Ввод ([([([(((ПРОЦИДУРА)]))])])
    { Private declarations }
Procedure CMDialogKey(Var Msg: TWMKey); message CM_DIALOGKEY;  //Запрет клавишы [Tab и Enter]
Procedure WMDisplayChange(var message:TMessage); message WM_DISPLAYCHANGE;  //Обнаружение изменений видеорежима

  public
    { Public declarations }
  end;

var
  Form1: TForm1;
//Дано
n_kac, ZY, D_1, Delta, S_TP, Delta_kac: Extended ;
//Расчет
n_TP_1, n_TP_2, Z, N, S_kac, a, S_TP_Y, b, b_k_a, b_k_a_1, b_k_a_2, ZY_b_k_a_1, ZY_b_k_a_2: Extended ;
//Счетчик
Chetcik1_1, Chetcik0_95 : Extended ;
  //изменение цвета текста
  indeks,indeks_RAS,Stroka,Stolbech: integer;
  // Excel
    ExcelApp, Workbook, Range, Cell1, Cell2, ArrayData  : Variant;
    TemplateFile : String;
    BeginCol, BeginRow, i, j : integer;
    RowCount, ColCount : integer;
implementation

uses Unit2, Unit3;

{$R *.dfm}

//Запрет клавишы [Tab и Enter]    ((ПРОЦЕДУРА))
procedure TForm1.CMDialogKey(var Msg: TWMKey);
begin
 Msg.Result := 0
end;

//Обнаружение изменений видеорежима    ((ПРОЦЕДУРА))
procedure TForm1.WMDisplayChange(var message: TMessage);
var   FullProgPath: PChar;
begin
inherited;
Application.Messagebox('Обнаружено изменение видеорежима. Программа будет перезапущена','Видеорежим', MB_ICONASTERISK or mb_ok) ;
FullProgPath := PChar(Application.ExeName);
WinExec(FullProgPath, SW_SHOW); // Or better use the CreateProcess function
Application.Terminate; // or: Close;
end;


//ОКРКГЛЕНИЕ чисел       ([([([(((ПРОЦИДУРА)]))])])
function TForm1.RoundSignificant(num: Extended; col: integer): Extended;
var
  counter, MaxValue, MinValue, PreSign: integer;
  operand: Extended;
begin
  if (col <= 0) or (num = 0)
    then
      begin
        result := 0;
        Exit;
      end;
  try
    MaxValue := Trunc(IntPower(10, col));
  except
    result := num;
    Exit;
  end;
  MinValue := MaxValue div 10;
  counter := 0;
  PreSign := Sign(num);
  operand := Abs(num);
  while operand <= MinValue do
    begin
      operand := operand * 10;
      counter := counter + 1;
    end;
  while operand > MaxValue do
    begin
      operand := operand / 10;
      counter := counter - 1;
    end;
  result := Round(operand) / IntPower(10, counter) * PreSign;
end;

//АВТО ШИРИНА СТОЛБЦОВ       ((ПРОЦЕДУРА))
type
TGridHack = class(TCustomGrid);
procedure ResizeStringGrid(_Grid: TCustomGrid);
var
Col, Row: integer;
Grid: TGridHack;
MaxWidth: integer;
ColWidth: integer;
ColText: string;
MaxRow: integer;
Pribavka: integer;
ColWidths: array of integer;
begin
Grid := TGridHack(_Grid);
SetLength(ColWidths, Grid.ColCount);
Pribavka:=50;                // 12 шрифт (Times New Roman)
MaxRow := Grid.RowCount+1;   // Максимальное число сталбцов
if MaxRow > Grid.RowCount then
MaxRow := Grid.RowCount;
for Col := 0 to Grid.ColCount - 1 do
begin
MaxWidth := 0;
for Row := 0 to MaxRow - 1 do
begin
ColText  := Grid.GetEditText(Col, Row);
ColWidth := Grid.Canvas.TextWidth(ColText);
if ColWidth > MaxWidth then
MaxWidth := ColWidth;
end;
if goVertLine in Grid.Options then
Inc(MaxWidth, Grid.GridLineWidth);
ColWidths[Col]      := MaxWidth + Pribavka;
Grid.ColWidths[Col] := ColWidths[Col];
end;
end;

// выделить столбец
procedure SelectCol( StringGrid: TStringGrid; ColNumber: integer );
var
  NewSel: TGridRect;
begin
   with StringGrid do
   begin
      if ( ColNumber > FixedCols-1 ) and ( ColNumber < ColCount ) then
      begin
         NewSel.Left := ColNumber;
         NewSel.Top := FixedRows;
         NewSel.Right := ColNumber;
         NewSel.Bottom := RowCount - 1;
         Selection := NewSel;
      end;
   end;
end;

// Делаем многострочный Caption у Button
procedure TForm1.SetMultiLineButton(AParent: TWinControl);
var j : integer;
 ah : THandle;
begin
for j := 0 to AParent.ControlCount - 1 do
if (AParent.Controls[j] is TButton) then
begin
 ah := (AParent.Controls[j] as TButton).Handle;
 SetWindowLong(ah, GWL_STYLE,
 GetWindowLong(ah, GWL_STYLE) OR
 BS_MULTILINE) ;
end;
 end;

//ПОДСТРОЧНЫЕ  НАДСТРОЧНЫЕ СИМВОЛЫ             ([([([(((ПРОЦИДУРА)]))])])
function TForm1.UpDownSymbol(RangeH: String; StartUp, LengthUp, StartDown,LengthDown: integer): boolean;
begin
// Подстрочные символы
if (StartDown<>777) and (LengthDown<>777) then
ExcelApp.Range[RangeH].Characters[StartDown, LengthDown].Font.Subscript:= True  ;
// Надстрочные символы
if (StartUp<>777) and (LengthUp<>777) then
ExcelApp.Range[RangeH].Characters[StartUp, LengthUp].Font.Superscript:= True  ;
end;

//РАМКА Excel   ([([([(((ПРОЦИДУРА)]))])])     ([([([(((ПРОЦИДУРА)]))])])
function TForm1.ExcelPamka(Pamka: String;UP,Down,Left,Right: integer): integer ;
begin
ExcelApp.Range[Pamka].Select;
ExcelApp.Selection.Borders[xlEdgeBottom].Weight := UP;
ExcelApp.Selection.Borders[xlEdgeTop].Weight := Down;
ExcelApp.Selection.Borders[xlEdgeLeft].Weight := Left;
ExcelApp.Selection.Borders[xlEdgeRight].Weight := Right;
end;

//Установлен ли Excel
function TForm1.IsOLEObjectInstalled(Name: String): boolean;
var
  ClassID: TCLSID;
  Rez : HRESULT;
begin
  // Ищем CLSID OLE-объекта
  Rez := CLSIDFromProgID(PWideChar(WideString(Name)), ClassID);

  if Rez = S_OK then  // Объект найден
    Result := true
  else
    Result := false;
end;



//Excel Ввод    ([([([(((ПРОЦИДУРА)]))])])     ([([([(((ПРОЦИДУРА)]))])])
function TForm1.Excel_Raschet: boolean;
var
   //Сохранить Ексель
    DirPath, FullFileName : String;
   //Номера строк
   XS_1,XS_2: String  ;
   //Номера строк
   X_1,X_2: integer  ;
   //Расширение файла при сохранениеи
   Racshir: String;
begin
// Координаты левого верхнего угла области, в которую будем выводить данные
  BeginCol := 1;
  BeginRow := 1;
// Размеры выводимого массива данных
  // Столбцы
  RowCount := 300;
  // Строчки
  ColCount := 5;

// Создание Excel
  ExcelApp := CreateOleObject('Excel.Application');
// Отключаем реакцию Excel на события, чтобы ускорить вывод информации
  ExcelApp.Application.EnableEvents := false;
 //  Создаем Книгу (Workbook)
   Workbook := ExcelApp.WorkBooks.Add();
// Создаем Вариантный Массив, который заполним выходными данными
ArrayData := VarArrayCreate([1, RowCount, 1, ColCount], varVariant);

//  (((Исходные данные)))
     i:=2 ; //Строка
     j:=1 ; //Столбец
     ArrayData[i,j] := '№';   i:=i+1 ;
     for i:=i to 8 do
       ArrayData[i,j] := i-2;
     i:=2 ; //Строка
     j:=2 ; //Столбец
     ArrayData[i,j] := 'Название';                   i:=i+1 ;
     ArrayData[i,j] := 'Диаметр несущей обечайки';   i:=i+1 ;
     ArrayData[i,j] := 'Число кассет ';              i:=i+1 ;
     ArrayData[i,j] := 'Зазор между кассетами';      i:=i+1 ;
     ArrayData[i,j] := 'Толщина кожуха кассеты';     i:=i+1 ;
     ArrayData[i,j] := 'Шаг трубок';                 i:=i+1 ;
     ArrayData[i,j] := 'Число нечетных рядов трубок (уточняем в процессе расчета)';   i:=i+1 ;
     i:=2 ; //Строка
     j:=3 ; //Столбец
     ArrayData[i,j] := 'Обозначения';   i:=i+1 ;
     ArrayData[i,j] := 'D1';            i:=i+1 ;
     ArrayData[i,j] := 'n';             i:=i+1 ;
     ArrayData[i,j] := 'Dкас';          i:=i+1 ;
     ArrayData[i,j] := 'D';             i:=i+1 ;
     ArrayData[i,j] := 'Sтр';           i:=i+1 ;
     ArrayData[i,j] := 'Zнечет';        i:=i+1 ;
     i:=2 ; //Строка
     j:=4 ; //Столбец
     ArrayData[i,j] := 'Единицы измерения';   i:=i+1 ;
     ArrayData[i,j] := 'мм';   i:=i+1 ;
     ArrayData[i,j] := '';     i:=i+1 ;
     ArrayData[i,j] := 'мм';   i:=i+1 ;
     ArrayData[i,j] := 'мм';   i:=i+1 ;
     ArrayData[i,j] := 'мм';   i:=i+1 ;
     ArrayData[i,j] := '';     i:=i+1 ;
     i:=2 ; //Строка
     j:=5 ; //Столбец
     ArrayData[i,j] := 'Значения';      i:=i+1 ;
     ArrayData[i,j] := D_1;             i:=i+1 ;
     ArrayData[i,j] := n_kac;           i:=i+1 ;
     ArrayData[i,j] := Delta_kac;       i:=i+1 ;
     ArrayData[i,j] := Delta;           i:=i+1 ;
     ArrayData[i,j] := S_TP;            i:=i+1 ;
     ArrayData[i,j] := ZY;              i:=i+1 ;

// (((Результаты расчета)))
     i:=10 ; //Строка
     j:=1 ; //Столбец
     for i:=i to 19 do
       ArrayData[i,j] := i-9;
     i:=10 ; //Строка
     j:=2 ; //Столбец
     ArrayData[i,j] := 'Шаг кассеты';                          i:=i+1 ;
     ArrayData[i,j] := 'Размер в свету';                       i:=i+1 ;
     ArrayData[i,j] := 'Размер в свету по нормали ';           i:=i+1 ;
     ArrayData[i,j] := 'Вертикальный шаг трубок (по треугольной решетки)';   i:=i+1 ;
     ArrayData[i,j] := 'Число трубок в нечетных рядах';        i:=i+1 ;
     ArrayData[i,j] := 'Число трубок в четных рядах';          i:=i+1 ;
     ArrayData[i,j] := 'Общее число трубок';                   i:=i+1 ;
     ArrayData[i,j] := 'Новое число нечетных рядов трубок ';   i:=i+1 ;
     ArrayData[i,j] := 'Общее число рядов трубок';             i:=i+1 ;
     i:=10 ; //Строка
     j:=3 ; //Столбец
     ArrayData[i,j] := 'Sкас';        i:=i+1 ;
     ArrayData[i,j] := 'a';           i:=i+1 ;
     ArrayData[i,j] := 'b';           i:=i+1 ;
     ArrayData[i,j] := 'S''тр';       i:=i+1 ;
     ArrayData[i,j] := 'nнечет';      i:=i+1 ;
     ArrayData[i,j] := 'nчет';        i:=i+1 ;
     ArrayData[i,j] := 'N';           i:=i+1 ;
     ArrayData[i,j] := 'Z''нечет';    i:=i+1 ;
     ArrayData[i,j] := 'Z';           i:=i+1 ;
     ArrayData[i,j] := 'b/a';         i:=i+1 ;
     i:=10 ; //Строка
     j:=4 ; //Столбец
     ArrayData[i,j] := 'мм';        i:=i+1 ;
     ArrayData[i,j] := 'мм';        i:=i+1 ;
     ArrayData[i,j] := 'мм';        i:=i+1 ;
     ArrayData[i,j] := 'мм';        i:=i+1 ;
     i:=10 ; //Строка
     j:=5 ; //Столбец
     ArrayData[i,j] := S_kac;   i:=i+1 ;
     ArrayData[i,j] := a;       i:=i+1 ;
     ArrayData[i,j] := b;       i:=i+1 ;
     ArrayData[i,j] := S_TP_Y;  i:=i+1 ;
     ArrayData[i,j] := n_TP_1;  i:=i+1 ;
     ArrayData[i,j] := n_TP_2;  i:=i+1 ;
     ArrayData[i,j] := N;       i:=i+1 ;
     ArrayData[i,j] := ZY;      i:=i+1 ;
     ArrayData[i,j] := Z;       i:=i+1 ;
     ArrayData[i,j] := b_k_a;   i:=i+1 ;


// Левая верхняя ячейка области, в которую будем выводить данные
  Cell1 := WorkBook.WorkSheets[1].Cells[BeginRow, BeginCol];

// Правая нижняя ячейка области, в которую будем выводить данные
  Cell2 := WorkBook.WorkSheets[1].Cells[BeginRow  + RowCount - 1, BeginCol +
ColCount - 1];

// Область, в которую будем выводить данные
  Range := WorkBook.WorkSheets[1].Range[Cell1, Cell2];

// А вот и сам вывод данных
  // Намного быстрее поячеечного присвоения
  Range.Value := ArrayData;

XS_1:='3' ;
XS_2:='19' ;
// Выравнивание и Шрафт
ExcelApp.Range['A'+XS_1+':A'+XS_2].Select;
ExcelApp.Selection.HorizontalAlignment:=3;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 11;
ExcelApp.Selection.Font.Name := 'Calibri';
ExcelApp.Range['B'+XS_1+':B'+XS_2].Select;
ExcelApp.Selection.HorizontalAlignment:=2;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 11;
ExcelApp.Selection.Font.Name := 'Cambria';
ExcelApp.Range['C'+XS_1+':C'+XS_2].Select;
ExcelApp.Selection.HorizontalAlignment:=3;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 13;
ExcelApp.Selection.Font.Name := 'Calibri';
ExcelApp.Range['D'+XS_1+':D'+XS_2].Select;
ExcelApp.Selection.HorizontalAlignment:=3;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 11;
ExcelApp.Selection.Font.Name := 'Cambria';
ExcelApp.Range['E'+XS_1+':E'+XS_2].Select;
ExcelApp.Selection.HorizontalAlignment:=2;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 11;
ExcelApp.Selection.Font.Name := 'Times New Roman';
  // Заголовок
  ExcelApp.Range['A2:E2'].Select;
  ExcelApp.Selection.WrapText:=True;
  ExcelApp.Selection.HorizontalAlignment:=3;
  ExcelApp.Selection.VerticalAlignment:=2 ;
  ExcelApp.Selection.Font.Size := 12;
  ExcelApp.Selection.Font.Name := 'Times New Roman';
  ExcelApp.Range['D2:D2'].Font.Size := 11.5;

  // Перенос по слову
ExcelApp.Range['B3:B19'].Select;
ExcelApp.Selection.WrapText:=True;

// Ширена и Высота столбцов и строк
    //ширена
  ExcelApp.Range['A:A', EmptyParam].EntireColumn.ColumnWidth :=4.71 ;
  ExcelApp.Range['B:B', EmptyParam].EntireColumn.ColumnWidth :=37.57 ;
  ExcelApp.Range['C:C', EmptyParam].EntireColumn.ColumnWidth :=13.14 ;
  ExcelApp.Range['D:D', EmptyParam].EntireColumn.ColumnWidth :=12.43 ;
  ExcelApp.Range['E:E', EmptyParam].EntireColumn.ColumnWidth :=15.29 ;
  //Высота   (только ДАНО!!!!)
   ExcelApp.Range['A1:A19', EmptyParam].EntireRow.autofit ;

// Обьядинение ячеек и их Шрифт и текст
ExcelApp.Range['A1:E1'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Selection.HorizontalAlignment:=3;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 14;
ExcelApp.Selection.Font.Name := 'Times New Roman';
ExcelApp.Selection:='Исходные данные'  ;
ExcelApp.Range['A9:E9'].Select;
ExcelApp.Selection.MergeCells:=True ;
ExcelApp.Selection.HorizontalAlignment:=3;
ExcelApp.Selection.VerticalAlignment:=2 ;
ExcelApp.Selection.Font.Size := 14;
ExcelApp.Selection.Font.Name := 'Times New Roman';
ExcelApp.Selection:='Результаты расчета';

//Рисуем ячейки
  //Дано
X_1:=2 ;
X_2:=8 ;
     for i:=X_1 to X_2 do
       begin
         XS_1:=floattostr(i) ;
         ExcelPamka('A'+XS_1+':E'+XS_1,2,2,-4138,-4138) ;
       end;
  //Результаты расчета
X_1:=10 ;
X_2:=19 ;
     for i:=X_1 to X_2 do
       begin
         XS_1:=floattostr(i) ;
         ExcelPamka('A'+XS_1+':E'+XS_1,2,2,-4138,-4138) ;
       end;
//обводка целеком
XS_1:='2' ;
XS_2:='8' ;
ExcelPamka('A'+XS_1+':A'+XS_2,-4138,-4138,-4138,-4138) ;
ExcelPamka('B'+XS_1+':B'+XS_2,-4138,-4138,-4138,-4138) ;
ExcelPamka('C'+XS_1+':C'+XS_2,-4138,-4138,-4138,-4138) ;
ExcelPamka('D'+XS_1+':D'+XS_2,-4138,-4138,-4138,-4138) ;
ExcelPamka('E'+XS_1+':E'+XS_2,-4138,-4138,-4138,-4138) ;
XS_1:='10' ;
XS_2:='19' ;
ExcelPamka('A'+XS_1+':A'+XS_2,-4138,-4138,-4138,-4138) ;
ExcelPamka('B'+XS_1+':B'+XS_2,-4138,-4138,-4138,-4138) ;
ExcelPamka('C'+XS_1+':C'+XS_2,-4138,-4138,-4138,-4138) ;
ExcelPamka('D'+XS_1+':D'+XS_2,-4138,-4138,-4138,-4138) ;
ExcelPamka('E'+XS_1+':E'+XS_2,-4138,-4138,-4138,-4138) ;

// Подстрочные символы
UpDownSymbol('C3:C8',777,777,2,8) ;
UpDownSymbol('C10:C10',777,777,2,8) ;
UpDownSymbol('C13:C13',777,777,3,8) ;
UpDownSymbol('C14:C15',777,777,2,8) ;
UpDownSymbol('C17:C17',777,777,3,8) ;

// Шрифт Symbol
ExcelApp.Range['C5'].Characters[0,1].Font.Name := 'Symbol';
ExcelApp.Range['C6'].Characters[0,1].Font.Name := 'Symbol';

// ВЫДЕЛЕНИЕ КОНЕЧНОЙ ЯЧЕЙКИ
ExcelApp.Range['A2'].Select;  //(Видно только результаты расчета)
END;

//ЗАКРЫТЬ Excel   ([([([(((ПРОЦИДУРА)]))])])        ([([([(((ПРОЦИДУРА)]))])])
function TForm1.Excel_Exit: boolean;
begin
  // Закрыть Excel
  ExcelApp.DisplayAlerts := False;
  ExcelApp.Quit; // закрыть Excel
  ExcelApp := Unassigned;
end;

////Окно печати  Excel   ([([([(((ПРОЦИДУРА)]))])]) ([([([(((ПРОЦИДУРА)]))])])
Function ShowPrintDialog:boolean;
begin
ShowPrintDialog:=true;
Form1.WindowState := wsMinimized;
try
  ShowPrintDialog:=ExcelApp.Dialogs.Item[8].Show;
  Form1.WindowState := wsNormal;
except
  ShowPrintDialog:=false;
  Form1.WindowState := wsNormal;
end;
End;

// Блокнот Ввод        ([([([(((ПРОЦИДУРА)]))])])  ([([([(((ПРОЦИДУРА)]))])])
function TForm1.Memo1_Raschet: boolean;
begin
     Memo1.Clear ;
Memo1.Lines.Add('                           Исходные данные');
Memo1.Lines.Add('');
Memo1.Lines.Add('Название                        Обозначения    Единицы     Значения');
Memo1.Lines.Add('                                              измерения');
Memo1.Lines.Add('Диаметр несущей обечайки           D1            мм           '+FloatToStr(D_1));;
Memo1.Lines.Add('Число кассет                       n                          '+FloatToStr(n_kac));
Memo1.Lines.Add('Зазор между кассетами              Дельта_кас    мм           '+FloatToStr(Delta_kac));
Memo1.Lines.Add('Толщина кожуха кассеты             Дельта        мм           '+FloatToStr(Delta));
Memo1.Lines.Add('Шаг трубок                         Sтр           мм           '+FloatToStr(S_TP));
Memo1.Lines.Add('Число нечетных рядов трубок        Zнечет                     '+FloatToStr(ZY));
Memo1.Lines.Add('');
Memo1.Lines.Add('                         Результаты расчета');
Memo1.Lines.Add('');
Memo1.Lines.Add('Шаг кассеты                        Sкас          мм           '+FloatToStr(S_kac));
Memo1.Lines.Add('Размер в свету                     a             мм           '+FloatToStr(a));
Memo1.Lines.Add('Расстояние по нормали '  );
Memo1.Lines.Add('диаметру обечайки между ');
Memo1.Lines.Add('стенками кассеты                   b             мм           '+FloatToStr(b));
Memo1.Lines.Add('Вертикальный шаг трубок '  );
Memo1.Lines.Add('(по треугольной решетки)           S''тр          мм           '+FloatToStr(S_TP_Y));
Memo1.Lines.Add('Число трубок в нечетных рядах      nнечет                     '+FloatToStr(n_TP_1));
Memo1.Lines.Add('Число трубок в четных рядах        nчет                       '+FloatToStr(n_TP_2));
Memo1.Lines.Add('Общее число трубок                 N                          '+FloatToStr(N));
Memo1.Lines.Add('Новое число нечетных рядов трубок  Z''нечет                    '+FloatToStr(ZY));
Memo1.Lines.Add('Общее число рядов трубок           Z                          '+FloatToStr(Z));
Memo1.Lines.Add('                                   b/a                        '+FloatToStr(b_k_a));
end;

// КНОПКА 1
procedure TForm1.Button1Click(Sender: TObject);
var
//Проверка заполнения полей
i_i: integer ;

begin
//Проверка заполнения полей
//пустота
with CLStringGrid2 do
  begin
    for i_i:= 1 to RowCount-FixedRows do
     begin
      if Cells[0,i_i]='' then
        begin
          Application.Messagebox('Поля то заполнять надо!','Ахтунг !!!', mb_iconerror or mb_ok);
          Exit; /// (Поля то заполнять надо)
        end;
      end;
    //нуми
    for i_i:= 1 to RowCount-FixedRows do
     begin
      if strtofloat(Cells[0,i_i])=0 then
        begin
          Application.Messagebox('Ноль, он и в Африке ноль','Ахтунг !!!', mb_iconerror or mb_ok);
          Exit; /// (Ноль, он и в Африке ноль)
        end;
      end;
  end;

//ДАНО
with CLStringGrid2 do
 begin
   //[1] Число касет
   D_1:=strtofloat(Cells[0,1]);
   //[2] Число рядов трубок
   n_kac:=strtofloat(Cells[0,2]);
   //[3] Внешний диаметр [мм]
   Delta_kac:=strtofloat(Cells[0,3]);
   //[4] Толщина кожуха касеты  [мм]
   Delta:=strtofloat(Cells[0,4]);
   //[5] Шаг трубок
   S_TP:=strtofloat(Cells[0,5]);
   //[6] Зазор между касетами
   ZY:=strtofloat(Cells[0,6]);
 end;


//0) Перевод [мм] в [м]
// Дано
//[9] D1       [Внешний диаметр]
//D_1:=D_1*power(10,-3);
//[4] dтр      [Диаметр трубки]
//d_TP:=d_TP*power(10,-3);

//РАСЧЕТ
//1)Шаг кассеты
try  S_kac:=2*PI*D_1/(n_kac-2*PI) ;    except   if Application.Messagebox('Ошибка (1)   [Шак касеты].','Ахтунг !!!', mb_iconerror or mb_ok) = mrOk    then   Exit; end;
                                       if   S_kac<=0   then    begin Application.Messagebox('Ошибка (1)   [Шак касеты - отрецательное значение].','Ахтунг !!!', mb_iconerror or mb_ok) ; Exit; end;
//2)
try  a:=S_kac-Delta_kac-2*Delta;                  except   if Application.Messagebox('Ошибка (2)   [а].','Ахтунг !!!', mb_iconerror or mb_ok) = mrOk    then    Exit;  end;
                                       if   a<=0       then  begin Application.Messagebox('Ошибка (2)   [а - отрецательное значение].','Ахтунг !!!', mb_iconerror or mb_ok); Exit; end;
//3)Число трубок во втором ряду
try  n_TP_2:= Int(a/S_TP) ;            except   if Application.Messagebox('Ошибка (3)   [Число трубок в первом ряду].','Ахтунг !!!', mb_iconerror or mb_ok) = mrOk    then    Exit;  end;
                                       if   n_TP_2<=0  then  begin Application.Messagebox('Ошибка (3)   [Число трубок в первом ряду - отрецательное значение].','Ахтунг !!!', mb_iconerror or mb_ok); Exit; end;
//4)Число трубок в первом ряду
try  n_TP_1:=n_TP_2-1 ;                except   if Application.Messagebox('Ошибка (4)   [Число трубок во втором ряду].','Ахтунг !!!', mb_iconerror or mb_ok) = mrOk    then   Exit;  end;
                                       if   n_TP_1<=0  then  begin Application.Messagebox('Ошибка (4)   [Число трубок во втором ряду - отрецательное значение].','Ахтунг !!!', mb_iconerror or mb_ok); Exit; end;
//5) Вертикальный шаг трубок
try  S_TP_Y:=SIN(PI/3)*S_TP ;          except   if Application.Messagebox('Ошибка (5)   [Вертикальный шак трубок].','Ахтунг !!!', mb_iconerror or mb_ok) = mrOk    then   Exit;  end;
                                       if   S_TP_Y<=0  then  begin Application.Messagebox('Ошибка (5)   [Вертикальный шак трубок - отрецательное значение].','Ахтунг !!!', mb_iconerror or mb_ok); Exit; end;
//6) Общее число рядов
try  Z:=2*ZY-1 ;                       except   if Application.Messagebox('Ошибка (6)   [Общее число рядов].','Ахтунг !!!', mb_iconerror or mb_ok) = mrOk    then    Exit;  end;
                                       if   Z<=0       then  begin Application.Messagebox('Ошибка (6)   [Общее число рядов - отрецательное значение].','Ахтунг !!!', mb_iconerror or mb_ok); Exit; end;
//7)
try  b:=(Z-1)*S_TP*SIN(PI/3)+S_TP ;    except   if Application.Messagebox('Ошибка (7)   [b].','Ахтунг !!!', mb_iconerror or mb_ok) = mrOk    then   Exit;  end;
                                       if   b<=0       then  begin Application.Messagebox('Ошибка (7)   [b - отрецательное значение].','Ахтунг !!!', mb_iconerror or mb_ok); Exit; end;
//8)a/b
try  b_k_a:=b/a ;                      except   if Application.Messagebox('Ошибка (8)   [a/b].','Ахтунг !!!', mb_iconerror or mb_ok) = mrOk    then   Exit;  end;

While (b_k_a<=0.98) or (b_k_a>=1.1)do  //  Если не попали в интервал 0,95<=b/a<=1,1 то поехали
  begin
  Chetcik1_1:=0 ;
  Chetcik0_95:=0 ;
   if (b/a>= 1.1) then
       begin
         ZY:=ZY-1  ;
         Z:=2*ZY-1 ;
         b:=(Z-1)*S_TP*SIN(PI/3)+S_TP ;
         b_k_a:=b/a ;
         Chetcik1_1:=1 ;
         b_k_a_1:=b/a ;
         ZY_b_k_a_1:=ZY ;
       end;
   if (b/a<= 0.98) then
       begin
         ZY:=ZY+1  ;
         Z:=2*ZY-1 ;
         b:=(Z-1)*S_TP*SIN(PI/3)+S_TP ;
         b_k_a:=b/a ;
         Chetcik0_95:=1 ;
         b_k_a_2:=b/a ;
         ZY_b_k_a_2:=ZY ;
       end;
  if (b/a<= 1.1)    and  (b/a>= 0.98)   then b_k_a:=1 ; //Если начинаем скакать от менишего b/a к большему то выходим из  [[While]]
  if (Chetcik1_1=1) and (Chetcik0_95=1) then            //Если начинаем скакать от менишего a/b к большему то выходим из  [[While]] при этом берется ZY при котором b/a имет наименьшее значение
// //    begin
// //      if (b_k_a_1)< 1 then
// //      if (1-b_k_a_1)<(b_k_a_2-1) then ZY:=ZY_b_k_a_1
// //                                 else ZY:=ZY_b_k_a_2
// //                      else
// //                                 if (b_k_a_1-1)<(1-b_k_a_2) then ZY:=ZY_b_k_a_1
// //                                 else ZY:=ZY_b_k_a_2 ;
b_k_a:=1 ;
// //   end;
  end;
/////9)Число трубок (Чет; нечет)
/////try  if  Odd(Trunc(ZY)) then  N:=(ZY-1)*n_TP_2+ZY*n_TP_1-7    // нечет
/////                        else  N:=(ZY-1)*n_TP_2+ZY*n_TP_1-4 ;  // чет
/////{ОШИБКА} except   if Application.Messagebox('Ошибка (9)   Чет; нечет.','Ахтунг !!!', mb_iconerror or mb_ok) = mrOk    then   Exit;  end;
/////{ОШИБКА}          if   N<=0       then  begin Application.Messagebox('Ошибка (9)   [Чет; нечет - отрецательное значение].','Ахтунг !!!', mb_iconerror or mb_ok); Exit; end;

//9)Полное число трубок
try N:=(ZY-1)*n_TP_2+ZY*n_TP_1 ;       except   if Application.Messagebox('Ошибка (9)   Чет; нечет.','Ахтунг !!!', mb_iconerror or mb_ok) = mrOk    then   Exit;  end;
                                       if   N<=0       then  begin Application.Messagebox('Ошибка (9)   [Чет; нечет - отрецательное значение].','Ахтунг !!!', mb_iconerror or mb_ok); Exit; end;


//Расчет закончен
indeks_RAS:=1 ;
//Округление результатов
S_kac:=Round(S_kac) ;
a:=Round(a) ;
b:=Round(b) ;
S_TP_Y:=Round(S_TP_Y*10)/10 ;
n_TP_1:=Round(n_TP_1) ;
n_TP_2:=Round(n_TP_2) ;
N:=Round(N) ;
ZY:=Round(ZY) ;
Z:=Round(Z) ;
b_k_a:=RoundSignificant(b/a,3) ;
//Вывод данных
with CLStringGrid4 do
  begin
   Cells[0,1]:=floattostr(S_kac);
   Cells[0,2]:=floattostr(a);
   Cells[0,3]:=floattostr(b);
   Cells[0,4]:=floattostr(S_TP_Y);
   Cells[0,5]:=floattostr(n_TP_1);
   Cells[0,6]:=floattostr(n_TP_2);
   Cells[0,7]:=floattostr(N);
   Cells[0,8]:=floattostr(ZY);
   Cells[0,9]:=floattostr(Z);
   Cells[0,10]:=floattostr(b_k_a);
  end;
//Почернение
Shape_KPACH_1.Brush.Color:=clBlack;
Shape_KPACH_2.Brush.Color:=clBlack;
//Выделение Таблицы
CLStringGrid2.SetFocus;
//Переход на стр 2
PageControl1.ActivePage:=TabSheet2 ;
//Размер формы
PageControl1.Height:=Round(PageControl1.Height*1.5) ;
Form1.Height:=PageControl1.Height+CLStringGrid1.RowHeights[1];
end;

  //КНОПКА 2
//нажатие
procedure TForm1.Button2Click(Sender: TObject);
begin
//Октивна первая страница
PageControl1.ActivePage:=TabSheet1 ;
//Размер формы
PageControl1.Height:=Round(PageControl1.Height/1.5) ;
Form1.Height:=PageControl1.Height+CLStringGrid1.RowHeights[1];
end;

  //КНОПКА 1-Схема
//нажатие
procedure TForm1.Button1_CxemClick(Sender: TObject);
begin
Form3.Show ;
end;

//Движение мыши
procedure TForm1.Button2MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
var
  //снять выделение в StringGrid
  hGridRect: TGridRect;
begin
//снять выделение в StringGrid
hGridRect.Top := -1;
hGridRect.Left := -1;
hGridRect.Right := -1;
hGridRect.Bottom := -1;
CLStringGrid4.Selection := hGridRect;
end;

//СОЗДАНИЕ ФОРМЫ
procedure TForm1.FormCreate(Sender: TObject);
begin
//МАШТАБИРОВАНИЕ ФОРМЫ                  (!!!!!!)                (!!!!!!)
//ScaleForm(form1) ;
//Октивна первая страница
PageControl1.ActivePage:=TabSheet1 ;
end;

//ПОЯВЛЕНИЕ ФОРМЫ
procedure TForm1.FormShow(Sender: TObject);
var
  x, y, w: integer;
  s: string;
//  MaxWidth: integer;
  //Авто ширена столбца
  i,j,maxWidth: integer ;
  //снять выделение в StringGrid
  hGridRect: TGridRect;
  //Намер ряда (Высота)
  Row_Heights: integer ;
begin
// Установлен ли Excel  [[[[[Меню ((Печать))]]]]]
if not IsOLEObjectInstalled('Excel.Application') then
     begin
        M_Bloknot.Checked:=true   ; //(Октивен чекет Блокнот) ('Класс не зарегистрирован')
        M_Excel.Enabled:=False;
     end
  else
        M_Excel.Checked:=true   ; //(Октивен чекет Excel) ('Класс найден');

//снять выделение в StringGrid
//hGridRect.Top := -1;
//hGridRect.Left := -1;
//hGridRect.Right := -1;
//hGridRect.Bottom := -1;
//CLStringGrid4.Selection := hGridRect;
//МАШТАБИРОВАНИЕ ФОРМЫ           (!!!!!!)           (!!!!!!)
//   if screen.width=1280 then            //[1280]
//     begin
//       //Прокрутка
//       if screen.Height<1024 then
//          begin
//           VertScrollBar.Range:=Panel_Vce.Height-6; } {Устанавливаем диапазон вертикальной прокрутки.}
//           VertScrollBar.Visible:=True;                 {Показываем вертикальную полосу прокрутки.}
//          end;
//       //Сдвиг элементов
//       form1.Height:=318 ;
//       form1.Width:= 498 ;
//     end;
//   if screen.width=1152 then            //[1152]
//     begin
//     //Прокрутка
//       if screen.Height<864 then
//          begin
//          VertScrollBar.Range:=Panel_Vce.Height-5; } {Устанавливаем диапазон вертикальной прокрутки.}
//           VertScrollBar.Visible:=True;                 {Показываем вертикальную полосу прокрутки.}
//          end;
//       //Сдвиг элементов
//       form1.Height:=form1.Height+24 ;
//     end;
//   if screen.width=1024 then            //[1024]
//     begin
//     //Прокрутка
//       if screen.Height<768 then
//          begin
//          VertScrollBar.Range:=Panel_Vce.Height-6; } {Устанавливаем диапазон вертикальной прокрутки.}
//           VertScrollBar.Visible:=True;                 {Показываем вертикальную полосу прокрутки.}
//          end;
//       //Сдвиг элементов
//       form1.Height:=form1.Height+25 ;
//
//     end;
//   if screen.width=800 then            //[800]
//     begin
//     //Прокрутка
//       if screen.Height<600 then
//          begin
//  {         VertScrollBar.Range:=Panel_Vce.Height-3; } {Устанавливаем диапазон вертикальной прокрутки.}
//           VertScrollBar.Visible:=True;                 {Показываем вертикальную полосу прокрутки.}
//          end;
//       //Сдвиг элементов
//       form1.Height:=form1.Height+28 ;
//     end;

CLStringGrid1.Cells[0,0]:='№' ;
CLStringGrid1.Cells[0,1]:='1' ;
CLStringGrid1.Cells[0,2]:='2' ;
CLStringGrid1.Cells[0,3]:='3' ;
CLStringGrid1.Cells[0,4]:='4' ;
CLStringGrid1.Cells[0,5]:='5' ;
CLStringGrid1.Cells[0,6]:='6' ;
CLStringGrid1.Cells[1,0]:='Обозначения' ;
CLStringGrid1.Cells[2,0]:='измерения' ;      ////////!!!
CLStringGrid1.Cells[2,1]:='[мм]' ;
CLStringGrid1.Cells[2,2]:='' ;
CLStringGrid1.Cells[2,3]:='[мм]' ;
CLStringGrid1.Cells[2,4]:='[мм]' ;
CLStringGrid1.Cells[2,5]:='[мм]' ;
CLStringGrid1.Cells[2,6]:='' ;
CLStringGrid1.Cells[3,0]:='Название' ;
CLStringGrid1.Cells[3,1]:=' Диаметр несущей обечайки' ;
CLStringGrid1.Cells[3,2]:=' Число кассет' ;
CLStringGrid1.Cells[3,3]:=' Зазор между кассетами' ;
CLStringGrid1.Cells[3,4]:=' Толщина кожуха кассеты' ;
CLStringGrid1.Cells[3,5]:=' Шаг трубок' ;
CLStringGrid1.Cells[3,6]:=' Число нечетных рядов' ;

with CLStringGrid2 do
  begin
CLStringGrid2.Cells[0,0]:='Значения' ;
CLStringGrid2.Cells[0,1]:='1500' ;
CLStringGrid2.Cells[0,2]:='52' ;
CLStringGrid2.Cells[0,3]:='10' ;
CLStringGrid2.Cells[0,4]:='2' ;
CLStringGrid2.Cells[0,5]:='17' ;
CLStringGrid2.Cells[0,6]:='8' ;
  end;

with CLStringGrid3 do
  begin
Cells[0,0]:='№' ;
  for i:= 1 to Rowcount-1 do  Cells[0,i]:=floattostr(i) ;
Cells[1,0]:='Обозначения' ;
Cells[2,0]:='измерения' ;      ////////!!!
  for i:= 1 to 4 do      Cells[2,i]:='[мм]' ;
Cells[3,0]:='Название' ;
Cells[3,1]:=' Шаг касеты' ;
Cells[3,2]:=' Размер в свету' ;
Cells[3,3]:=' Размер в свету по нормали' ;
Cells[3,7]:=' Общее число трубок' ;          

  end;

with CLStringGrid4 do
  begin
Cells[0,0]:='Значения' ;
  end;


//Выделение Таблицы
CLStringGrid2.SetFocus;
CLStringGrid2.Row:=1;    //Строка
CLStringGrid2.Col:=0;   //Столбец



//Авто ширена столбцов
ResizeStringGrid(CLStringGrid1) ;
ResizeStringGrid(CLStringGrid2) ;
ResizeStringGrid(CLStringGrid3) ;
//Авто ширена и высота столбцов
with CLStringGrid1 do
  begin
    //Высота столбца[0]
    RowHeights[0] := Round(RowHeights[0] * 1.8);
    RowHeights[6] := Round(RowHeights[6] * 2.4);
    //Ширену таблицы подганяет
    Width:=ColWidths[0];
      for i:= 1 to colcount-1 do
        Width:=Width+ColWidths[i]+GridLineWidth*2;
      Height:=Height ;

    //Высоте таблицы подганяет
    Height:=RowHeights[0];
      for i:= 1 to RowCount-1 do
        Height:=Height+RowHeights[i]+GridLineWidth*2;
      Height:=Height-2 ;

    Panel7.Left:=Left+ColWidths[0]+Round(ColWidths[1]/2.5);
    Panel8.Left:=Left+ColWidths[0]+Round(ColWidths[1]/2.5);
    Panel9.Left:=Left+ColWidths[0]+Round(ColWidths[1]/2.5);
    Panel10.Left:=Left+ColWidths[0]+Round(ColWidths[1]/2.5);
    Panel11.Left:=Left+ColWidths[0]+Round(ColWidths[1]/2.5);
    Panel12.Left:=Left+ColWidths[0]+Round(ColWidths[1]/2.5);

    Panel7.top:=0 ;
    Panel8.top:=0 ;
    Panel9.top:=0 ;
    Panel10.top:=0 ;
    Panel11.top:=0 ;
    Panel12.top:=0 ;
    Panel7.top:=top+RowHeights[0]+Round(RowHeights[1]*0.25);
    Panel8.top:=top+RowHeights[0]+RowHeights[1]+Round(RowHeights[2]*0.3);
    Panel9.top:=top+RowHeights[0]+RowHeights[1]+RowHeights[2]+Round(RowHeights[3]*0.3);
    Panel10.top:=top+RowHeights[0]+RowHeights[1]+RowHeights[2]+RowHeights[3]+Round(RowHeights[4]*0.3);
    Panel11.top:=top+RowHeights[0]+RowHeights[1]+RowHeights[2]+RowHeights[3]+RowHeights[4]+Round(RowHeights[5]*0.4);
    Panel12.top:=top+RowHeights[0]+RowHeights[1]+RowHeights[2]+RowHeights[3]+RowHeights[4]+RowHeights[5]+Round(RowHeights[6]*0.45);
  end;

with CLStringGrid2 do
  begin
    //Ширену таблицы подганяет
    Width:=ColWidths[0];
      for i:= 1 to colcount-1 do
      Width:=Width+ColWidths[i]+GridLineWidth*2;
   //Координаты
    Left:=CLStringGrid1.Left+CLStringGrid1.Width ;
    Left:=Left+1 ;
    Top:=CLStringGrid1.Top ;
    Height:=CLStringGrid1.Height ;
   //Высота ячейкек
   for i:= 0 to RowCount-1 do
    RowHeights[i]:=CLStringGrid1.RowHeights[i] ;
  end;

with CLStringGrid3 do
  begin
    //Высота столбца[0]
    RowHeights[0] := Round(RowHeights[0] * 1.8);
   //Координаты
    Left:=CLStringGrid1.Left ;
    Top:=CLStringGrid1.Top ;
    //Высота столбца[1],[2]
    RowHeights[4] := Round(RowHeights[4] * 1.8);
    RowHeights[5] := Round(RowHeights[5] * 1.8);
    RowHeights[6] := Round(RowHeights[6] * 1.8);
    RowHeights[8] := Round(RowHeights[8] * 1.8);
    RowHeights[9] := Round(RowHeights[9] * 1.8);
    //Ширену таблицы подганяет
      for i:= 0 to colcount-1 do
        ColWidths[i]:=CLStringGrid1.ColWidths[i];
      Width:=CLStringGrid1.Width;
    //Высоте таблицы подганяет
    Height:=RowHeights[0];
      for i:= 1 to RowCount-1 do
        Height:=Height+RowHeights[i]+GridLineWidth*2;
      Height:=Height-6 ;

      Panel13.Left:=Left+ColWidths[0]+Round(ColWidths[1]/2.5);
      Panel14.Left:=Left+ColWidths[0]+Round(ColWidths[1]/2.5);
      Panel15.Left:=Left+ColWidths[0]+Round(ColWidths[1]/2.5);

      Row_Heights:=0 ;
      for i:= 0 to 0 do  Row_Heights:=Row_Heights+RowHeights[i] ;
          Panel13.top:=top+Row_Heights+Round(RowHeights[0]*0.07);
      Row_Heights:=0  ;
      for i:= 0 to 1 do  Row_Heights:=Row_Heights+RowHeights[i] ;
          Panel14.top:=top+Row_Heights+Round(RowHeights[1]*0.25);
      Row_Heights:=0 ;
      for i:= 0 to 2 do  Row_Heights:=Row_Heights+RowHeights[i] ;
          Panel15.top:=top+Row_Heights+Round(RowHeights[2]*0.25);
      Row_Heights:=0 ;
      for i:= 0 to 3 do  Row_Heights:=Row_Heights+RowHeights[i] ;
          Panel16.top:=top+Row_Heights+Round(RowHeights[3]*0.65);
      Row_Heights:=0 ;
      for i:= 0 to 4 do  Row_Heights:=Row_Heights+RowHeights[i] ;
          Panel17.top:=top+Row_Heights+Round(RowHeights[4]*0.3);
      Row_Heights:=0 ;
      for i:= 0 to 5 do  Row_Heights:=Row_Heights+RowHeights[i] ;
          Panel18.top:=top+Row_Heights+Round(RowHeights[5]*0.35);
      Row_Heights:=0 ;
      for i:= 0 to 6 do  Row_Heights:=Row_Heights+RowHeights[i] ;
          Panel19.top:=top+Row_Heights+Round(RowHeights[6]*0.2);
      Row_Heights:=0 ;
      for i:= 0 to 7 do  Row_Heights:=Row_Heights+RowHeights[i] ;
          Panel20.top:=top+Row_Heights+Round(RowHeights[7]*0.8);
      Row_Heights:=0 ;
      for i:= 0 to 8 do  Row_Heights:=Row_Heights+RowHeights[i] ;
          Panel21.top:=top+Row_Heights+Round(RowHeights[8]*0.5);
      Row_Heights:=0 ;
      for i:= 0 to 9 do  Row_Heights:=Row_Heights+RowHeights[i] ;
          Panel22.top:=top+Row_Heights+Round(RowHeights[9]*0.3);

    Panel13.Left:=Left+ColWidths[0]+Round(ColWidths[1]/2.5);
    Panel14.Left:=Left+ColWidths[0]+Round(ColWidths[1]/2.5);
    Panel15.Left:=Left+ColWidths[0]+Round(ColWidths[1]/2.5);
    Panel16.Left:=Left+ColWidths[0]+Round(ColWidths[1]/2.5);
    Panel17.Left:=Left+ColWidths[0]+Round(ColWidths[1]/2.5);
    Panel18.Left:=Left+ColWidths[0]+Round(ColWidths[1]/2.5);
    Panel19.Left:=Left+ColWidths[0]+Round(ColWidths[1]/2.5);
    Panel20.Left:=Left+ColWidths[0]+Round(ColWidths[1]/2.5);
    Panel21.Left:=Left+ColWidths[0]+Round(ColWidths[1]/2.5);
    Panel22.Left:=Left+ColWidths[0]+Round(ColWidths[1]/2.5);
  end;

with CLStringGrid4 do
  begin
    //Ширена табл
      Width:=CLStringGrid2.Width ;
   //Высота табл
      Height:=CLStringGrid3.Height ;
    //Ширена ячейкек
    for i:= 0 to colcount-1 do
      ColWidths[i]:=CLStringGrid2.ColWidths[i];
   //Координаты
    Left:=CLStringGrid2.Left ;
    Top:=CLStringGrid3.Top ;
    //Высота табл
    Height:=CLStringGrid3.Height ;
   //Высота ячейкек
   for i:= 0 to RowCount-1 do
    RowHeights[i]:=CLStringGrid3.RowHeights[i] ;
  end;

with Label1 do    //Исходные данные
  begin
    //Координаты
    Left:=CLStringGrid1.Left+Round((CLStringGrid1.Width+CLStringGrid2.Width)/2)-Round(Width/2) ;
    Top:=Button1.Top+Round(Button1.Height/4) ;//Round((CLStringGrid1.Top+CLStringGrid1.Height-CLStringGrid3.Top-CLStringGrid3.Height)/10.1) ;
  end;

with Label57 do    //Результаты расчета
  begin
    //Координаты
    Left:=CLStringGrid1.Left+Round((CLStringGrid1.Width+CLStringGrid2.Width)/2)-Round(Width/2) ;
    Top:=Label1.Top ;
  end;

with Button1 do    //Кнопка
 begin
   //Координаты
   Left:=CLStringGrid1.Left+CLStringGrid1.Width+CLStringGrid2.Width-Width+2 ;
   Top:=Round((CLStringGrid1.Top-Height)/2)-2 ;
 end;

with Panel_IS_H do    //Линия Ввертик (Исходные данные)
  begin
   //Координаты  (Panel)
   Left:=CLStringGrid1.Left+CLStringGrid1.Width ;
   Top:=CLStringGrid1.Top ;
   //Высота  (Panel)
   Height:=CLStringGrid1.Height ;
   //Высота  (Линии)
   Shape1.Height:=CLStringGrid1.Height;
  end;
with Panel_P_H do    //Линия Ввертик (Результаты расчета)
  begin
   //Координаты  (Panel)
   Left:=CLStringGrid3.Left+CLStringGrid3.Width ;
   Top:=CLStringGrid3.Top ;
   //Высота  (Panel)
   Height:=CLStringGrid3.Height ;
   //Высота  (Линии)
   Shape2.Height:=CLStringGrid3.Height;
  end;
with Shape_IS_OBV do    //Обводка (Исходные данные)
  begin
   //Координаты
   Left:=CLStringGrid1.Left-1 ;
   Top:=CLStringGrid1.Top-1 ;
   //Высота
   Height:=CLStringGrid1.Height+2 ;
   //Ширена
   Width:=CLStringGrid1.Width+CLStringGrid2.Width+3 ;
  end;
with Shape_P_OBV do    //Обводка (Результаты расчета)
  begin
   //Координаты
   Left:=CLStringGrid3.Left-1 ;
   Top:=CLStringGrid3.Top-1 ;
   //Высота
   Height:=CLStringGrid3.Height+2 ;
   //Ширена
   Width:=CLStringGrid3.Width+CLStringGrid4.Width+3 ;
  end;

with Form1 do   //Размер формы
  begin
    Width:=CLStringGrid1.Left*2+CLStringGrid1.Width+CLStringGrid2.Width+4 ;
    Height:=PageControl1.Height+CLStringGrid1.RowHeights[1];
  end;

with Panel_Button_2 do  //Делаем кнопку 2
  begin
  //Делаем многострочный Caption у Button 2
   SetMultiLineButton (Panel_Button_2);
   //Высота и  Ширена
   Button2.Width:=Button1.Width  ;
   Button2.Height:=Button1.Height  ;
   Height:=Button2.Height ;
   Width:=Button2.Width ;
   //Координаты
   Left:=CLStringGrid3.Left+CLStringGrid3.Width+CLStringGrid4.Width-Width+2 ;
   Top:=Round((CLStringGrid3.Top-Height)/2)-2 ;
  end;
with Button1_Cxem do  //Кнопка 1-Cхема
  begin
   //Высота
   Height:=Button1.Height  ;
   //Координаты
   Left:=Button1.Left-Width-2 ;
   Top:=Button1.Top ;
  end;
with Button2_Cxem do  //Кнопка 2-Cхема
  begin
   //Высота и  Ширена
   Height:=Button1_Cxem.Height ;
   Width:=Button1_Cxem.Width ;
   //Координаты
   Left:=Button1_Cxem.Left ;
   Top:=Button1_Cxem.Top ;
  end;
with Image_save_1 do  //(КАРТИНКА) Сохранить
  begin
   //Координаты
   Left:=CLStringGrid1.Left-1;
   Top:=Round((CLStringGrid1.Top-Height)/2)-2 ;
  end;
with Image_save_2 do  //(КАРТИНКА) Сохранить
  begin
   //Координаты
   Left:=Image_save_1.Left ;
   Top:=Image_save_1.Top ;
  end;
with Image_printer_1 do  //(КАРТИНКА) Принтер
  begin
   //Координаты
   Left:=Image_save_1.Left+Image_save_1.Width+8;
   Top:=Image_save_1.Top ;
  end;
with Image_printer_2 do  //(КАРТИНКА) Принтер
  begin
   //Координаты
   Left:=Image_printer_1.Left ;
   Top:=Image_printer_1.Top ;
  end;
with Printer_1 do  //Принтер
  begin
   //Координаты
   Left:=Image_printer_1.Left ;
   Top:=Image_printer_1.Top ;
  end;
with Printer_2 do  //Принтер
  begin
   //Координаты
   Left:=Image_printer_1.Left ;
   Top:=Image_printer_1.Top ;
  end;
with Save_1 do  //Сохранить
  begin
   //Координаты
   Left:=Image_save_1.Left ;
   Top:=Image_save_1.Top ;
  end;
with Save_2 do  //Сохранить
  begin
   //Координаты
   Left:=Image_save_1.Left ;
   Top:=Image_save_1.Top ;
  end;
with Shape_KPACH_1 do  //(КАРТИНКА) Цветовичек
  begin
   //Координаты
   Left:=Image_save_1.Left-1;
   Top:=Image_save_1.Top-1 ;
   //Высота и  Ширена
   Height:=Image_save_1.Height+3 ;
   Width:=Image_save_1.Width+Image_printer_1.Width+10 ;
   //Цвет
   Brush.Color:=clRed;
  end;
with Shape_KPACH_2 do  //Цветовичек
  begin
   //Координаты
   Left:=Shape_KPACH_1.Left ;
   Top:=Shape_KPACH_1.Top ;
   //Высота и  Ширена
   Height:=Shape_KPACH_1.Height ;
   Width:=Shape_KPACH_1.Width ;
   //Цвет
   Brush.Color:=Shape_KPACH_1.Brush.Color;
   pen.Mode:=Shape_KPACH_1.pen.Mode ;
  end;
with Memo1 do  //Цветовичек
  begin
   //Координаты
   Left:=-2*Width ;
   Top:=0 ;
  end;
end;
///    [[[МЕНЮ]]]

//           ((((ФАЙЛ))))
// Выход               (F12)
procedure TForm1.N2Click(Sender: TObject);
begin
Close;
end;
// Расчитать           (ПРОБЕЛ)
procedure TForm1.N5Click(Sender: TObject);
begin
if (Form3.Active) and (PageControl1.ActivePage=TabSheet1) then begin  PageControl1.SetFocus; CLStringGrid2.SetFocus; Exit;  end;
if (Form3.Active) and (PageControl1.ActivePage=TabSheet2) then begin  PageControl1.SetFocus; CLStringGrid4.SetFocus; Exit;  end;
if PageControl1.ActivePage=TabSheet1 then begin Button1.Click; Exit;  end;
if PageControl1.ActivePage=TabSheet2 then begin Button2.Click; Exit;  end;

end;
//     [[[ПАРАМЕТРЫ]]]
//               (((ПЕЧАТЬ)))
// Excel
procedure TForm1.M_ExcelClick(Sender: TObject);
begin
M_Excel.Checked:=true   ;

end;
// Блокнот
procedure TForm1.M_BloknotClick(Sender: TObject);
begin
M_Bloknot.Checked:=true   ;

end;
//           ((((Справка))))
// О программе (Перейти к Form2)     (F1)
procedure TForm1.N4Click(Sender: TObject);
begin
Form2.ShowModal;
end;




procedure TForm1.CLStringGrid1DrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
const
  N = 7;
var
  st: string;
  kol_vo_strok, dl, i: integer;
  // Убераем выделение ячейки при старте
   StringGrid: TStringGrid;
   Can: TCanvas;
   // Центрирование
   cr:TRect;
Format: Word;
C: array[0..255] of Char;
    s:string;
    Flag: Cardinal;
   // Центрирование по вертикали
    text: string;
begin
//Как сменить цвет выделения в TStringGrid?
   if ( gdSelected in State ) then
 begin
   StringGrid := Sender as TStringGrid;
   Can := StringGrid.Canvas;
   Can.Font := StringGrid.Font;
   if (ARow >= StringGrid.FixedRows) and (ACol >= StringGrid.FixedCols)
    then Can.Brush.Color := StringGrid.Color
    else Can.Brush.Color := StringGrid.FixedColor;
   If ActiveControl = Sender then // Эта строка "выключает" выделение
   if (gdSelected in State) then
     begin
       Can.Font.Color := clHighlightText;
       Can.Brush.Color := clHighlight;
     end;
   Can.FillRect(Rect);
   Can.TextOut(Rect.Left+2,Rect.Top+2, StringGrid.Cells[ACol, ARow]);
 end;

// Центрирование по Центру
if ((ACol=0) and (ARow>=0))
or ((ACol=1) and (ARow>=0))
or ((ACol=2) and (ARow>=0))
or ((ACol=3) and  (ARow=0))
or ((ACol=4) and  (ARow=0)) or ((ACol=5) and  (ARow=0))
 then
   begin
     with CLStringGrid1 do
       begin
        text:=Cells[ACol,ARow];
        Canvas.Brush.Color:=clWindow;
        Canvas.FillRect(Rect);
        // Центрирование по вертикали
        DrawText(Canvas.Handle, PChar(text), Length(text), Rect, DT_CENTER or DT_VCENTER or DT_SINGLELINE);
      end;
  end;
// Центрирование по леву
if ((ACol=3) and (ARow>=1))
 then
   begin
     with CLStringGrid1 do
       begin
        text:=Cells[ACol,ARow];
        Canvas.Brush.Color:=clWindow;
        Canvas.FillRect(Rect);
        // Центрирование по вертикали
        DrawText(Canvas.Handle, PChar(text), Length(text), Rect,DT_VCENTER or DT_SINGLELINE);
      end;
  end;
//Многострочный TStringGrid
if (ACol=2) and (ARow=0) then
begin
  with CLStringGrid1 do
    begin
      Canvas.FillRect(Rect);
      s := 'Единицы измерения';
      Flag := DT_VCENTER or DT_CENTER or DT_WORDBREAK  ;
      DrawText(Canvas.Handle,PChar(s), length(s),Rect,Flag);
    end;
end;
if (ACol=3) and (ARow=6)then
begin
  with CLStringGrid1 do
    begin
      Canvas.FillRect(Rect);
      s := ' Число нечетных рядов           трубок (уточняем в                  процессе расчета)';
      Flag := DT_VCENTER or DT_VCENTER or DT_WORDBREAK ;
      DrawText(Canvas.Handle,PChar(s), length(s),Rect,Flag);
    end;
end;

end;

procedure TForm1.CLStringGrid3DrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
Var
  // Убераем выделение ячейки при старте
   StringGrid: TStringGrid;
   Can: TCanvas;
   // Центрирование
   cr:TRect;
   text: String;
 //Многострочный TStringGrid
   s:string;
   Flag: Cardinal;
begin
//Как сменить цвет выделения в TStringGrid?
   if ( gdSelected in State ) then
 begin
   StringGrid := Sender as TStringGrid;
   Can := StringGrid.Canvas;
   Can.Font := StringGrid.Font;
   if (ARow >= StringGrid.FixedRows) and (ACol >= StringGrid.FixedCols)
    then Can.Brush.Color := StringGrid.Color
    else Can.Brush.Color := StringGrid.FixedColor;
   If ActiveControl = Sender then // Эта строка "выключает" выделение
   if (gdSelected in State) then
   begin
     Can.Font.Color := clHighlightText;
     Can.Brush.Color := clHighlight;
   end;
   Can.FillRect(Rect);
   Can.TextOut(Rect.Left+2,Rect.Top+2, StringGrid.Cells[ACol, ARow]);
 end;
// Центрирование по вертикали
if ((ACol=0) and (ARow>=0))
or ((ACol=1) and (ARow=0))
or ((ACol=2) and (ARow>=0))
or ((ACol=3) and (ARow=0))
 then
   begin
     with CLStringGrid3 do
       begin
        text:=Cells[ACol,ARow];
        Canvas.Brush.Color:=clWindow;
        Canvas.FillRect(Rect);
        // Центрирование по вертикали
        DrawText(Canvas.Handle, PChar(text), Length(text), Rect, DT_CENTER or DT_VCENTER or DT_SINGLELINE);
      end;
  end;
 //Многострочный TStringGrid
if (ACol=2) and (ARow=0) then   //Единицы измерения
begin
  with CLStringGrid3 do
    begin
      Canvas.FillRect(Rect);
      s := 'Единицы измерения';
      Flag :=  DT_CENTER or DT_VCENTER or DT_WORDBREAK ;
      DrawText(Canvas.Handle,PChar(s), length(s),Rect,Flag);
    end;
end;
if (ACol=3) and (ARow=4) then    //  Вертикальный шаг трубок
begin
  with CLStringGrid3 do
    begin
      Canvas.FillRect(Rect);
      s := ' Вертикальный шаг                  трубок ';
      Flag :=  DT_VCENTER or DT_WORDBREAK ;
      DrawText(Canvas.Handle,PChar(s), length(s),Rect,Flag);
    end;
end;
if (ACol=3) and (ARow=5) then    //  Число трубок в нечетных рядах
begin
  with CLStringGrid3 do
    begin
      Canvas.FillRect(Rect);
      s := ' Число трубок в                        нечетных рядах';
      Flag :=  DT_VCENTER or DT_WORDBREAK ;
      DrawText(Canvas.Handle,PChar(s), length(s),Rect,Flag);
    end;
end;
if (ACol=3) and (ARow=6) then    // Число трубок в четных рядах
begin
  with CLStringGrid3 do
    begin
      Canvas.FillRect(Rect);
      s := ' Число трубок в                        четных рядах';
      Flag :=  DT_VCENTER or DT_WORDBREAK ;
      DrawText(Canvas.Handle,PChar(s), length(s),Rect,Flag);
    end;
end;if (ACol=3) and (ARow=8) then    //  Новое число нечетных рядов трубок
begin
  with CLStringGrid3 do
    begin
      Canvas.FillRect(Rect);
      s := ' Новое число нечетных           рядов трубок';
      Flag :=  DT_VCENTER or DT_WORDBREAK ;
      DrawText(Canvas.Handle,PChar(s), length(s),Rect,Flag);
    end;
end;
if (ACol=3) and (ARow=9) then    //
begin
  with CLStringGrid3 do
    begin
      Canvas.FillRect(Rect);
      s := ' Общее число рядов                 трубок';
      Flag :=  DT_VCENTER or DT_WORDBREAK ;
      DrawText(Canvas.Handle,PChar(s), length(s),Rect,Flag);
    end;
end;

end;

procedure TForm1.CLStringGrid2DrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  // Убераем выделение ячейки при старте
   StringGrid: TStringGrid;
   Can: TCanvas;
   // Центрирование
   cr:TRect;
   text: String;
 //Многострочный TStringGrid
   s:string;
   Flag: Cardinal;
begin
// Центрирование по леву
if ((ACol=0) and (ARow>=1))
 then
   begin
     with CLStringGrid2 do
       begin
        text:=Cells[ACol,ARow];
        Canvas.Brush.Color:=clWindow;
        Canvas.FillRect(Rect);
        // Центрирование по вертикали
        DrawText(Canvas.Handle, PChar(text), Length(text), Rect,DT_VCENTER or DT_SINGLELINE);
      end;
  end;
// Центрирование по вертикали
if ((ACol=0) and (ARow=0))
 then
   begin
     with CLStringGrid2 do
       begin
        text:=Cells[ACol,ARow];
        Canvas.Brush.Color:=clWindow;
        Canvas.FillRect(Rect);
        // Центрирование по вертикали
        DrawText(Canvas.Handle, PChar(text), Length(text), Rect, DT_CENTER or DT_VCENTER or DT_SINGLELINE);
      end;
  end;


end;
procedure TForm1.CLStringGrid4DrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  // Убераем выделение ячейки при старте
   StringGrid: TStringGrid;
   Can: TCanvas;
   // Центрирование
   cr:TRect;
   text: String;
 //Многострочный TStringGrid
   s:string;
   Flag: Cardinal;
begin
//Как сменить цвет выделения в TStringGrid?
   if ( gdSelected in State ) and (indeks=0) then
 begin
   StringGrid := Sender as TStringGrid;
   Can := StringGrid.Canvas;
   Can.Font := StringGrid.Font;
   if (ARow >= StringGrid.FixedRows) and (ACol >= StringGrid.FixedCols)
    then Can.Brush.Color := StringGrid.Color
    else Can.Brush.Color := StringGrid.FixedColor;
   If ActiveControl = Sender then // Эта строка "выключает" выделение
   if (gdSelected in State) then
   begin
     Can.Font.Color := clHighlightText;
     Can.Brush.Color := clHighlight;
   end;
   Can.FillRect(Rect);
   DrawText(Canvas.Handle, PChar(text), Length(text), Rect, DT_VCENTER or DT_SINGLELINE);
   indeks:=1 ;
 end;

// Центрирование по леву
if ((ACol=0) and (ARow>=1))
 then
   begin
     with CLStringGrid4 do
       begin
        text:=Cells[ACol,ARow];
        Canvas.Brush.Color:=clWindow;
        Canvas.FillRect(Rect);
        // Центрирование по вертикали
        DrawText(Canvas.Handle, PChar(text), Length(text), Rect,DT_VCENTER or DT_SINGLELINE);
      end;
  end;
// Центрирование по вертикали
if ((ACol=0) and (ARow=0))
 then
   begin
     with CLStringGrid4 do
       begin
        text:=Cells[ACol,ARow];
        Canvas.Brush.Color:=clWindow;
        Canvas.FillRect(Rect);
        // Центрирование по вертикали
        DrawText(Canvas.Handle, PChar(text), Length(text), Rect, DT_CENTER or DT_VCENTER or DT_SINGLELINE);
      end;
  end;
end;



procedure TForm1.CLStringGrid2KeyPress(Sender: TObject; var Key: Char);
var
  k:integer;
//снять выделение в StringGrid
  hGridRect: TGridRect;
//Координаты выделение
  X,Y:integer;
begin
with CLStringGrid2 do
  begin
    X:=Row ;
    Y:=Col ;
    if (X=1) or (X=2) then //Ввод только целых чисел (строки 1 и 2)
        begin
          if  key in['.', ','] then
              key:=#0;
        end
     else      //замена точки на запятую (остальные строчки)
        begin
         if  Key='.' then
             Key:=',';
        end;
  end;
//стирать кнопкой [Esc]
if  Key=Chr(VK_ESCAPE) then
Key:=Chr( VK_BACK );
//чудо ввод
with CLStringGrid2 do
   begin
   if Cells[Y,X]='0' then          //Если ввелиноль, то после него
    if not(key in [',','.',#8]) then   //должна стоять запятая
          begin                        //либо только удалить его
           key:=#0;
           beep ;
          end;
   if key in['0'..'9',',','.',#8] then //разрешаем вводить только числа
     begin
     if (key=',') or (key='.') then //проверка для только одной запятой
       begin
       if Cells[Y,X]='' then key:=#0;
         For k:=1 to Length(Cells[Y,X]) do
         begin
          if (Cells[Y,X][k]=',') or (Cells[Y,X][k]=',')then
                   begin
                    key:=#0;
                    beep ;
                end;
         end;
       end;
     end else key:=#0;
end;
//Краснение
if key in['0'..'9',#8] then
   begin
     Shape_KPACH_1.Brush.Color:=clRed;
     Shape_KPACH_2.Brush.Color:=clRed;
   end;
End;


procedure TForm1.CLStringGrid2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
//Координаты выделение
  X,Y:integer;
begin
with CLStringGrid2 do
  begin
    X:=Row ;
    Y:=Col ;
    //Кнопка Вверх и Вниз (Переход от крайней нижней к крайней верхней ячейке)
    if  (Key=Word(VK_DOWN)) and (X=RowCount-FixedRows) then  begin Key:=Word(VK_PRIOR);  Exit; end;
    if  (Key=Word(VK_UP)) and (X=FixedRows) then  begin Key:=Word(VK_NEXT);  Exit; end;
  end;
end;

procedure TForm1.CLStringGrid4KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
//Координаты выделение
  X,Y:integer;
begin
with CLStringGrid4 do
  begin
    X:=Row ;
    Y:=Col ;
    //Кнопка Вверх и Вниз (Переход от крайней нижней к крайней верхней ячейке)
    if  (Key=Word(VK_DOWN)) and (X=RowCount-FixedRows) then  begin Key:=Word(VK_PRIOR);  Exit; end;
    if  (Key=Word(VK_UP)) and (X=FixedRows) then  begin Key:=Word(VK_NEXT);  Exit; end;
  end;
end;

//______________[[[[[[[[[СОХРАНИТЬ]]]]]]]]]]____________________________________
procedure TForm1.Save_1Click(Sender: TObject);
type
 XlFileFormat = TOleEnum;
const
 xlAddIn = $00000012;
 xlCSV = $00000006;
 xlCSVMac = $00000016;
 xlCSVMSDOS = $00000018;
 xlCSVWindows = $00000017;
 xlDBF2 = $00000007;
 xlDBF3 = $00000008;
 xlDBF4 = $0000000B;
 xlDIF = $00000009;
 xlExcel2 = $00000010;
 xlExcel2FarEast = $0000001B;
 xlExcel3 = $0000001D;
 xlExcel4 = $00000021;
 xlExcel5 = $00000027;
 xlExcel7 = $00000027;
 xlExcel9795 = $0000002B;
 xlExcel4Workbook = $00000023;
 xlIntlAddIn = $0000001A;
 xlIntlMacro = $00000019;
 xlWorkbookNormal = $FFFFEFD1;
 xlSYLK = $00000002;
 xlTemplate = $00000011;
 xlCurrentPlatformText = $FFFFEFC2;
 xlTextMac = $00000013;
 xlTextMSDOS = $00000015;
 xlTextPrinter = $00000024;
 xlTextWindows = $00000014;
 xlWJ2WD1 = $0000000E;
 xlWK1 = $00000005;
 xlWK1ALL = $0000001F;
 xlWK1FMT = $0000001E;
 xlWK3 = $0000000F;
 xlWK4 = $00000026;
 xlWK3FM3 = $00000020;
 xlWKS = $00000004;
 xlWorks2FarEast = $0000001C;
 xlWQ1 = $00000022;
 xlWJ3 = $00000028;
 xlWJ3FJ3 = $00000029;
 xlUnicodeText = $0000002A;
 xlHtml = $0000002C;
var
   //Расширение файла при сохранениеи
   Racshir: String;
   // Кнопка в диологовом окошке
   temp:Word;
begin
//############Окна собщений############
//Проверка нажали ли на кнопку расчитать
if indeks_RAS<>1 then
  begin
    if Application.Messagebox('Какой смысл в пустом отчете?','Ахтунг !!!', mb_iconerror or mb_ok) = mrOk    then   Exit;
  end;
//Нужен ли вам старый отчет
if Shape_KPACH_2.Brush.Color=clRed then
  begin
      temp:=Application.Messagebox('Исходные данные были изменены. Перерасчет не был произведен. Продолжить выполнение операции?','Сохранение', MB_ICONQUESTION+MB_YESNO+MB_DEFBUTTON1);
      case temp of
        idYES:  ;
        idNO: Exit ;
      end;
  end;
//############СОХРАНИТЬ############
SaveDialog1.Title:='Сохранение результатов' ;
SaveDialog1.FilterIndex:=0 ;
SaveDialog1.filename := ExtractFilePath( Application.ExeName )+'Отчет';
  if SaveDialog1.Execute then
     begin
       Case SaveDialog1.FilterIndex of
               1: Racshir:='.xlsx' ;
               2: Racshir:='.xls'  ;
               3: Racshir:='.txt'      ;
           end;
    //Excel (2007) *.xlsx
      if SaveDialog1.FilterIndex=1 then
         begin
           if FileExists(SaveDialog1.FileName) then   // если такой файл уже есть то перезаписываем
               begin
                   Excel_Raschet ; //(ВЫВОД В Excel)
                   ExcelApp.ActiveSheet.SaveAs(SaveDialog1.filename);
                   Excel_Exit ; //(Закрыть Excel)
                   Exit ;
               end
           else   //если его нет то прибовляем к названию "Racshir"
               begin
                  Excel_Raschet ; //(ВЫВОД В Excel)
                  ExcelApp.ActiveSheet.SaveAs(SaveDialog1.filename+Racshir);
                  Excel_Exit ; //(Закрыть Excel)
                  Exit ;
               end;
        end ;
    //Excel (1995-1997) *.xls
      if SaveDialog1.FilterIndex=2 then
         begin
           if FileExists(SaveDialog1.FileName) then   // если такой файл уже есть то перезаписываем
               begin
                   Excel_Raschet ; //(ВЫВОД В Excel)
                   ExcelApp.ActiveSheet.SaveAs(SaveDialog1.filename, xlExcel5);
                   Excel_Exit ; //(Закрыть Excel)
                   Exit ;
               end
           else   //если его нет то прибовляем к названию "Racshir"
               begin
                  Excel_Raschet ; //(ВЫВОД В Excel)
                  ExcelApp.ActiveSheet.SaveAs(SaveDialog1.filename+Racshir, xlExcel5);
                  Excel_Exit ; //(Закрыть Excel)
                  Exit ;
               end;
        end ;
    //Блокнот *.txt
      if SaveDialog1.FilterIndex=3 then
         begin
           if FileExists(SaveDialog1.FileName) then   // если такой файл уже есть то перезаписываем
               begin
                   Memo1_Raschet ;
                   Memo1.Lines.SaveToFile(SaveDialog1.filename);
                   Exit ;
               end
           else   //если его нет то прибовляем к названию "Racshir"
               begin
                   Memo1_Raschet ;
                   Memo1.Lines.SaveToFile(SaveDialog1.filename+Racshir);
                   Exit ;
               end;
         end ;
     end
  else  Exit;

end;

//______________[[[[[[[[[ПЕЧАТЬ]]]]]]]]]]_______________________________________
procedure TForm1.printer_1Click(Sender: TObject);
var
   // Кнопка в диологовом окошке
   temp:Word;
   //Печать блокнот
   Line: TextFile;
   P: integer;
begin
//############Окна собщений############
//Проверка нажали ли на кнопку расчитать
if indeks_RAS<>1 then
  begin
    if Application.Messagebox('Какой смысл в пустом отчете?','Ахтунг !!!', mb_iconerror or mb_ok) = mrOk    then   Exit;
  end;
//Нужен ли вам старый отчет
if Shape_KPACH_2.Brush.Color=clRed then
  begin
      temp:=Application.Messagebox('Исходные данные были изменены. Перерасчет не был произведен. Продолжить выполнение операции?','Сохранение', MB_ICONQUESTION+MB_YESNO+MB_DEFBUTTON1);
      case temp of
        idYES:  ;
        idNO: Exit ;
      end;
  end;
//############ПЕЧАТЬ############
 if M_Excel.Checked=true then   //(Октивен чекет Excel)
    begin
       Excel_Raschet ; //(ВЫВОД В Excel)
       ExcelApp.WindowState := xlMinimized;
       ExcelApp.Visible := True ; //(Делаем Excel видимым)
       ShowPrintDialog;  //(Окно печати  Excel )
//PrintOut([From], [To], [Copies], [Preview], [ActivePrinter], [PrintToFile], [Collate], [PrToFileName])
      if not VarIsEmpty(ExcelApp) then
         begin
           Excel_Exit ; //(Закрыть Excel)
           Exit;
         end;
    end
 else
    begin
       If PrintDialog1.Execute then
           begin
             AssignPrn(Line);
             ReWrite(Line);
             Printer.Canvas.Font := Memo1.Font;
             for P := 0 to Memo1.Lines.Count -1 do Writeln (Line, Memo1.Lines[P]);
             System.CloseFile(Line);
           end;
    end;
end;




END.
