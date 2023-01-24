//КООРДИНАТЫ СОЗДАТЕЛЯ ПРОГРАММЫ
//Электронная почта: Ltlvfpfqbpfzws@yandex.ru
unit Unit2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, XPMaN, StdCtrls, jpeg, ExtCtrls;

type
  TForm2 = class(TForm)
    Memo1: TMemo;
    Button1: TButton;
    Image1: TImage;
    TimerKoordinatbI: TTimer;
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure TimerKoordinatbITimer(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
  //Процидуры
function NachaloTekct: boolean;  //Начальный текст     ((ПРОЦЕДУРА))
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

{$R *.dfm}

//______________[[[[[[[[[Начальный текст]]]]]]]]]_____(ПРОЦЕДУРА))_________
function TForm2.NachaloTekct: boolean;
begin
Memo1.Lines.Clear;
Memo1.Lines.Add('Программа для расчета');
Memo1.Lines.Add('числа параллельных трубок в ПГ');
Memo1.Lines.Add('');
Memo1.Lines.Add('Version 1.4  build 7  (2011)');
Memo1.Lines.Add('');
Memo1.Lines.Add('Санкт-Петербургский государственый');
Memo1.Lines.Add('морской технический университет');
Memo1.Lines.Add('');
Memo1.Lines.Add('Автор: Мазилевский И.И.');
Memo1.Lines.Add('            Ревков М.В.');
end;

//______________[[[[[[[[[Кнопка "ОК"]]]]]]]]]]_____________________________
procedure TForm2.Button1Click(Sender: TObject);
begin
TimerKoordinatbI.Enabled:=False ;  //ТАЙМЕР  Координаты создателя программы
Close;
NachaloTekct ; //Начальный текст
end;

//____________[[[[[[[[[ОТКРЫТИЕ  О программе]]]]]]]]]]_____________________
procedure TForm2.FormShow(Sender: TObject);
begin
TimerKoordinatbI.Enabled:=true ;  //ТАЙМЕР  Координаты создателя программы
Button1.SetFocus ;
end;

//____________[[[[[[[[[Закрытие ОКНА]]]]]]]]]_______________________________
procedure TForm2.FormClose(Sender: TObject; var Action: TCloseAction);
begin
TimerKoordinatbI.Enabled:=False ;  //ТАЙМЕР  Координаты создателя программы
Close;
NachaloTekct ; //Начальный текст
end;

//____________[[[[[[[[[ТАЙМЕР  ((Координаты создателя программы))]]]]]]]]]___
procedure TForm2.TimerKoordinatbITimer(Sender: TObject);
begin
  Memo1.Lines.Clear;
  Memo1.Lines.Add('КООРДИНАТЫ СОЗДАТЕЛЯ ПРОГРАММЫ') ;
  Memo1.Lines.Add('') ;
  Memo1.Lines.Add('Ф.И.О.:') ;
  Memo1.Lines.Add('Мазилевский Илья Игоревич') ;
  Memo1.Lines.Add('') ;
  Memo1.Lines.Add('Мобильный телефон номер:') ;
  Memo1.Lines.Add('+7(921)434-75-16') ;
  Memo1.Lines.Add('') ;
  Memo1.Lines.Add('Электронная почта:') ;
  Memo1.Lines.Add('Ltlvfpfqbpfzws@yandex.ru') ;
  TimerKoordinatbI.Enabled:=False ;
end;


end.
