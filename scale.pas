unit scale;

interface

uses
  Forms, WinTypes, WinProcs, SysUtils;

procedure ScaleForm(Sender: TObject);

implementation

procedure ScaleForm(Sender: TObject);

const

  {измените это так, чтобы это соответствовало
  режиму разрешения во время разработки}
  DesignScrY: LongInt = 1024;
  DesignScrX: LongInt = 1280;
  DesignBorder: LongInt = 8; {значение в Панели Управления + 1}

var

  SystemScrY: LongInt;
  SystemScrX: LongInt;
  SystemBorder: LongInt;
  OldHeight: LongInt;
  OldWidth: LongInt;

begin

  SystemScrY := GetSystemMetrics(SM_CYSCREEN);
  SystemScrX := GetSystemMetrics(SM_CXSCREEN);
  SystemBorder := GetSystemMetrics(SM_CYFRAME);
  with Sender as TForm do
  begin
    Scaled := True;
    AutoScroll := False;
    Top := Top * SystemScrX div DesignScrX;
    Left := Left * SystemScrX div DesignScrX;
    OldHeight := Height + (DesignBorder - SystemBorder) * 2;
    OldWidth := Width + (DesignBorder - SystemBorder) * 2;
    ScaleBy((OldWidth * SystemScrX div DesignScrX - SystemBorder * 2),
      (OldWidth - DesignBorder * 2));
    {
    Для форм не имеющих границ измените предшествующие
    три строки следующим способом:

    OldHeight := Height;
    OldWidth  := Width;
    ScaleBy(SystemScrX, DesignScrX);
    }

    Height := OldHeight * SystemScrY div DesignScrY;
    Width := OldWidth * SystemScrX div DesignScrX;
  end;
end;

begin
end.
