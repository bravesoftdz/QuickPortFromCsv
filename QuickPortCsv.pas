{
Quick Port From Csv
This Delphi Script for quick EDA working for Altium
Should be follow ADOH-DelphiScriptReference from Altium Wiki
# Class D
}
const
    Idx_Position = 0;
    Idx_Group = 1;
    Idx_NetLabel = 2;
    Idx_PortLabel = 3;
    Idx_SideLabel = 4;
    Idx_VisibleOpt = 5;
    Min_CoLNums = 5;
    Idx_RootInfo = 0;
    Idx_StartX = 5;
    Idx_StartY = 6;
{
      Col Index version 1
}
    __iter = 0;
    __posX = 1;
    __posY = 2;
    __left = 0;
    __bottom = 1;
    __right = 2;
    __top = 3;

Procedure MessageWirteLn(cls, txt, src, doc, cbprs, cbparm : String);
var
    WSM         : IWorkSpace;
    MM          : IMessagesManager;
    ImageIndex  : Integer;
    F           : Boolean;
Begin
    // Obtain the Workspace Manager interface
    WSM := GetWorkSpace;
    If WSM = Nil Then Exit;

    // Obtain the Messages Panel interface
    MM := WSM.DM_MessagesManager;
    If MM = Nil Then Exit;

    // Tick Icon for the lines in the Message panel

    ImageIndex := 3;

    // Clear out messages from the Message panel...
    //MM.ClearMessages;
    MM.BeginUpdate;

    F := False;
    MM.AddMessage(cls,txt,src,doc,cbprs,cbparm, ImageIndex,F );
    MM.EndUpdate;

    // Display the Messages panel in DXP.
    WSM.DM_ShowMessageView;
End;

Procedure MessageClean();
var
    WSM         : IWorkSpace;
    MM          : IMessagesManager;
Begin
    // Obtain the Workspace Manager interface
    WSM := GetWorkSpace;
    If WSM = Nil Then Exit;

    // Obtain the Messages Panel interface
    MM := WSM.DM_MessagesManager;
    If MM = Nil Then Exit;

    // Clear out messages from the Message panel...
    MM.ClearMessages;
    MM.BeginUpdate;
    MM.EndUpdate;
    WSM.DM_ShowMessageView;
End;

Procedure printd(cls_str, msg_str : String);
Var
   dummy : integer;
Begin
      MessageWirteLn(cls_str, msg_str, 'debug', 'debug', 'debug', 'debug');
End;

Function MeasureTextLen(PortName : String): Integer;
Var
    Len         : Integer;
    I           : Integer;
    C           : Ansichar;
    Sum         : Integer;
    StrSize     : Integer;
Begin
    Sum := 0;
    StrSize := length(PortName);
    // I don't know pascal's For-loop Middle with 0 or 1
    For I := 1 to StrSize Do
    Begin
        // only works with New Roman 10px , Standard is Mils
        C := PortName[I];
        Case C of
            'A' : Sum := Sum + 69 ;
            'B' : Sum := Sum + 61 ;
            'C' : Sum := Sum + 66 ;
            'D' : Sum := Sum + 53 ;
            'E' : Sum := Sum + 52 ;
            'F' : Sum := Sum + 52 ;
            'G' : Sum := Sum + 73 ;
            'H' : Sum := Sum + 76 ;
            'I' : Sum := Sum + 29 ;
            'J' : Sum := Sum + 38 ;
            'K' : Sum := Sum + 62 ;
            'L' : Sum := Sum + 50 ;
            'M' : Sum := Sum + 96 ;
            'N' : Sum := Sum + 80 ;
            'O' : Sum := Sum + 81 ;
            'P' : Sum := Sum + 60 ;
            'Q' : Sum := Sum + 81 ;
            'R' : Sum := Sum + 64 ;
            'S' : Sum := Sum + 57 ;
            'T' : Sum := Sum + 56 ;
            'U' : Sum := Sum + 73 ;
            'V' : Sum := Sum + 66 ;
            'W' : Sum := Sum + 100 ;
            'X' : Sum := Sum + 63 ;
            'Y' : Sum := Sum + 59 ;
            'Z' : Sum := Sum + 61 ;
            'a' : Sum := Sum + 54 ;
            'b' : Sum := Sum + 63 ;
            'c' : Sum := Sum + 50 ;
            'd' : Sum := Sum + 63 ;
            'e' : Sum := Sum + 56 ;
            'f' : Sum := Sum + 33 ;
            'g' : Sum := Sum + 63 ;
            'h' : Sum := Sum + 61 ;
            'i' : Sum := Sum + 26 ;
            'j' : Sum := Sum + 26 ;
            'k' : Sum := Sum + 52 ;
            'l' : Sum := Sum + 26 ;
            'm' : Sum := Sum + 92 ;
            'n' : Sum := Sum + 61 ;
            'o' : Sum := Sum + 63 ;
            'p' : Sum := Sum + 63 ;
            'q' : Sum := Sum + 63 ;
            'r' : Sum := Sum + 36 ;
            's' : Sum := Sum + 46 ;
            't' : Sum := Sum + 37 ;
            'u' : Sum := Sum + 61 ;
            'v' : Sum := Sum + 51 ;
            'w' : Sum := Sum + 78 ;
            'x' : Sum := Sum + 49 ;
            'y' : Sum := Sum + 52 ;
            'z' : Sum := Sum + 49 ;
            '_' : Sum := Sum + 45 ;
            '-' : Sum := Sum + 43 ;
            ' ' : Sum := Sum + 37 ;
            '.' : Sum := Sum + 23 ;
            '/' : Sum := Sum + 42 ;
            '$' : Sum := Sum + 58 ;
            '^' : Sum := Sum + 73 ;
            '#' : Sum := Sum + 64 ;
            '<' : Sum := Sum + 73 ;
            '>' : Sum := Sum + 73 ;
            '1' : Sum := Sum + 47 ;
            '2' : Sum := Sum + 47 ;
            '3' : Sum := Sum + 47 ;
            '4' : Sum := Sum + 47 ;
            '5' : Sum := Sum + 47 ;
            '6' : Sum := Sum + 47 ;
            '7' : Sum := Sum + 47 ;
            '8' : Sum := Sum + 47 ;
            '9' : Sum := Sum + 47 ;
            '0' : Sum := Sum + 47 ;
            Else Sum := Sum + 100;
        End;
    End;
    If (Sum mod 50) = 0 Then
       Result := Sum
    else
       Result := ((Sum div 50) + 1) * 50;
       Result := Sum;

End;

Procedure NewPort(MiddleX, MiddleY, PortLen : Integer, PortName : String, Side : Integer);
Var
    SchDoc     : ISch_Document;
    AName       : TDynamicString;
    Orientation : TRotationBy90;
    AElectrical : TPinElectrical;
    SchPort     : ISch_Port;
    Loc         : TLocation;
    CurView     : IServerDocumentView;
    StartX, StartY : Integer;
Begin
    If SchServer = Nil Then Exit;
    SchDoc := SchServer.GetCurrentSchDocument;
    If SchDoc = Nil Then Exit;

    SchPort := SchServer.SchObjectFactory(ePort,eCreate_GlobalCopy);
    If SchPort = Nil Then Exit;

    StartX := MiddleX;
    StartY := MiddleY;

    // Port's Location must Middle with Left or bottom Anchor
    Case Side of
        1 : StartX := MiddleX - PortLen;
        2 : StartY := MiddleY - PortLen;
    End;

    // Set Direction.
    Case Side of
        1 : SchPort.Style := ePortRight;
        2 : SchPort.Style := ePortTop;
        3 : SchPort.Style := ePortLeft;
        4 : SchPort.Style := ePortBottom;
    End;

    {.......................................................................
        Case Style of
        This If future using for Power of directical option.
     .......................................................................}
    //SchPort.Alignment := eHorizontalLeftAlign;

    SchPort.Location  := Point(MilsToCoord(StartX),MilsToCoord(StartY));
    // SchPort.Style     := ePortRight;
    SchPort.IOType    := ePortUnspecIfied;
    //SchPort.Alignment := eHorizontalCentreAlign;
    SchPort.Width     := MilsToCoord(PortLen);
    SchPort.Height    := MilsToCoord(100);
    //SchPort.AreaColor := $FFFF;
    SchPort.TextColor := $000080;
    SchPort.Name      := PortName;
    SchDoc.RegisterSchObjectInContainer(SchPort);
End;

Procedure NewWire(MiddleX, MiddleY, EndX, EndY, Side : Integer);
Var
    SchDoc          : ISch_Document;
    SchWire         : ISch_Wire;
Begin
    If SchServer = Nil Then Exit;
    SchDoc := SchServer.GetCurrentSchDocument;
    If SchDoc = Nil Then Exit;

    SchWire := SchServer.SchObjectFactory(eWire,eCreate_GlobalCopy);
    If SchWire = Nil Then Exit;
    //LineWidth := eSmall;
    // but this should be value , small
    SchWire.SetState_LineWidth := eSmall;
    Schwire.VerticesCount := 0;
    Schwire.Location := Point(MilsToCoord(MiddleX), MilsToCoord(MiddleY));
    Schwire.InsertVertex := 1;
    SchWire.SetState_Vertex(1, Point(MilsToCoord(EndX), MilsToCoord(EndY)));
    Schwire.InsertVertex := 2;
    SchWire.SetState_Vertex(2, Point(MilsToCoord(MiddleX), MilsToCoord(MiddleY)));
    SchDoc.RegisterSchObjectInContainer(SchWire);
End;

Procedure NewNet(MiddleX, MiddleY, EndX, EndY, LabelGap : Integer, NetName : String, Side : Integer ,IgnAng : Integer);
Var
    SchNetlabel     : ISch_Netlabel;
    SchPort         : ISch_Port;
    SchDoc          : ISch_Document;
    CurView         : IServerDocumentView;
    LabelX, LabelY  : Integer;
Begin
    If SchServer = Nil Then Exit;
    SchDoc := SchServer.GetCurrentSchDocument;
    If SchDoc = Nil Then Exit;
    SchNetlabel := SchServer.SchObjectFactory(eNetlabel,eCreate_GlobalCopy);
    If SchNetlabel = Nil Then Exit;
    LabelY := MiddleY;
    LabelX := MiddleX;
    // justIfication and orientation and start point.
    Case Side of
        1: LabelX := MiddleX + LabelGap;
        2: LabelY := EndY - LabelGap;
        3: LabelX := EndX + LabelGap;
        4: LabelY := MiddleY - LabelGap;
        End;
    //LabelX := MiddleX + EndGap;
    //LabelY := MiddleY + EndGap;
    SchNetlabel.Location    := Point(MilsToCoord(LabelX), MilsToCoord(LabelY));
    SchNetlabel.Text        := NetName;
    SchNetlabel.Orientation := eRotate0;
    Case Side of
        1: SchNetlabel.Orientation := eRotate0;
        2: SchNetlabel.Orientation := eRotate270;
        3: SchNetlabel.Orientation := eRotate0;
        {3: SchNetlabel.JustIfication := eJustIfy_BottomLeft;}
        4: SchNetlabel.Orientation := eRotate270;
        else SchNetlabel.Orientation := eRotate0;
        end;
    SchDoc.RegisterSchObjectInContainer(SchNetlabel);
End;

Procedure CreatePin(MiddleX, MiddleY, Side, WireLen : Integer, PortName, NetName : String, IgnBit : Integer);
Var
    EndX, EndY, StartX, StartY, PortLen : Integer;
    Len : Integer;
Begin

    EndX := MiddleX;
    EndY := MiddleY;
    
    Case Side of
        1 : EndX := (MiddleX+WireLen);
        2 : EndY := (MiddleY+WireLen);
        3 : EndX := (MiddleX-WireLen);
        4 : EndY := (MiddleY-WireLen);
        Else Exit;
    End;
    If (IgnBit and 1) = 0 Then
    Begin
        NewNet(MiddleX, MiddleY, EndX, EndY, 100, NetName, Side, 1);
    End;
    If (IgnBit and 2) = 0 Then
    Begin
        NewWire(MiddleX, MiddleY, EndX, EndY, Side);
    End;
    If (IgnBit and 4) = 0 Then
    Begin
        PortLen := MeasureTextLen(PortName)+200;
        NewPort(MiddleX, MiddleY, PortLen, PortName, Side);
    End;
End;


function _splitstring(const S, Delimiters: char): TStringList;
var
    newarr : TStringList;
    newstr : String;
    i : integer;
    j : integer;
Begin
     j := 0;
     newstr := '';
     newarr := TStringList.Create;

       For i := 1 to length(S) Do
       Begin
            If S[i] = Delimiters Then
            Begin
               newarr.add(newstr);
               newstr := '';
               j := j + 1;
            End
            Else
            Begin
                 newstr := newstr + S[i];
            End
       End;
       If i > 0 Then
       Begin
            newarr.add(newstr);
       end;
       Result := newarr;
End;

function GetSide(const S): Integer;
Var
    C : AnsiChar;
Begin
    //printd('asdf', S);
    If Length(S) = 0 Then
        Result := 1
    Else
    Begin
    C := S[1];
        Case C of
             'L' : Result := 0;
             'l' : Result := 0;
             '1' : Result := 0;
             'B' : Result := 1;
             'b' : Result := 1;
             '2' : Result := 1;
             'R' : Result := 2;
             'r' : Result := 2;
             '2' : Result := 2;
             'T' : Result := 3;
             't' : Result := 3;
             '3' : Result := 3;
             Else Result := 4;
        End;
    End;
End;

function GetVisibleState(const S): Integer;
Var
    C : AnsiChar;
Begin
    If Length(S) = 0 Then
        Result := 1
    Else
    Begin
        C := S[1];
        Case C of
            'N' : Result := 0;
            'X' : Result := 0;
            'n' : Result := 0;
            'x' : Result := 0;
            '-' : Result := 0;
            Else Result := 1;
        End;
    End;
End;

function PlaceViaCsv(path : string);
Var
    tfIn            : TextFile;
    rowstore        : TList;
    colstore        : TStringList;
    s               : string;
    i, j            : integer;
    currentNetSize  : Integer;
    NetMaxSize      : Integer;
    txtrow          : integer;
    currentSide     : String;
    C               : AnsiChar;
    pinSide         : Integer;
    hor, ver        : Integer;
    InitX, InitY    : Integer;
    pinX, pinY      : Integer;
    WireLen         : Integer;
    SideCount       : Array[0..4] of Integer;
    EachSideIter    : Array[0..4, 0..3] of Integer;
    IgnOpt          : Integer;
    plabel        : String;
    nlabel        : String;
Begin
    // Messages for notIfy Reading this
    printd('ReadText',path);
    // Set the name of the file that will be read
    AssignFile(tfIn, path);
    rowstore := TList.Create;
    txtrow := 0;
    NetMaxSize := 0;
    for i:=0 to 4 Do
    Begin
        SideCount[i] := 0;
    End;
    reset(tfIn);

    // Keep reading lines until the end of the file is reached
    while not eof(tfIn) do
    Begin
        readln(tfIn, s);
        //printd('ReadText', s);
        colstore := _splitstring(s, ',');
        If colstore.GetCount = 4 Then
        Begin
            colstore.add('');
        End;
        rowstore.add(colstore);
        {
        For i := 0 to colstore.GetCount - 1 Do
        Begin
            printd(IntToStr(i), colstore[i]);
        End;
        }
        If(txtrow > 0) Then
        Begin
            currentNetSize := MeasureTextLen(colstore[Idx_NetLabel]);
            If currentNetSize > NetMaxSize Then
            Begin
                NetMaxSize := currentNetSize;
            end;

            pinSide := GetSide(colstore[Idx_SideLabel]);
            SideCount[pinSide] := SideCount[pinSide] + 1;
        End;
        txtrow := txtrow + 1;
    End;
    //  For debug
    {
    for i := 0 to rowstore.Count -1 Do
    Begin
         for j := 0 to rowstore[i].GetCount -1 Do
         Begin
              printd(IntToStr(i)+ ' -> '+ IntToStr(j), rowstore[i][j]);
         End;
    End;
    }
    // Done so close the file
    CloseFile(tfIn);
  // Add SideCounts and NetMaxSize and in the Tlist
    printd('Bottom Count', IntToStr(SideCount[0]));
    printd('Left Count', IntToStr(SideCount[1]));
    printd('Right Count', IntToStr(SideCount[2]));
    printd('Top Count', IntToStr(SideCount[3]));
    printd('Unknown Count', IntToStr(SideCount[4]));

    InitX := StrToInt(rowstore[0][Idx_StartX]);
    InitY := StrToInt(rowstore[0][Idx_StartY]);
    WireLen := NetMaxSize + 200;
    //InitX := 5000;
    //InitY := 1000;

    If SideCount[0] >= SideCount[2] Then
        ver := SideCount[0]
    Else
        ver := SideCount[2];

    If SideCount[1] >= SideCount[3] Then
        hor := SideCount[1]
    Else
        hor := SideCount[3];
    
    If ver < 1 Then
    Begin
        ver := 1;
    End;
    If hor < 1 Then
    Begin 
        hor := 1;
    End;

    EachSideIter[__left,__iter] := 1;
    EachSideIter[__left,__posX] := InitX-WireLen;
    EachSideIter[__left,__posY] := InitY-200;

    EachSideIter[__bottom,__iter] := 1;
    EachSideIter[__bottom,__posX] := InitX+200;
    EachSideIter[__bottom,__posY] := InitY-400-((ver-1)*100)-WireLen;

    EachSideIter[__right,__iter] := 1;
    EachSideIter[__right,__posX] := InitX+400+((hor-1)*100)+WireLen;
    EachSideIter[__right,__posY] := InitY-200-((ver-1)*100);

    EachSideIter[__top,__iter] := 1;
    EachSideIter[__top,__posX] := InitX+200+((hor-1)*100);
    EachSideIter[__top,__posY] := InitY+WireLen;

    printd('Total Row', IntToStr(rowstore.Count -1));
    for i :=1 To rowstore.Count -1 Do
    Begin
        pinSide := GetSide(rowstore[i][Idx_SideLabel]);
        pinX := EachSideIter[pinSide,__posX];
        pinY := EachSideIter[pinSide,__posY];
        plabel := rowstore[i][Idx_PortLabel];
        nlabel := rowstore[i][Idx_NetLabel];
        IgnOpt := 0;
        If pinSide >= 4 Then
            printd('PlaceFromCSV', 'Unknown Pin Info Table')
        Else
        Begin
            Case pinSide of
                0 : EachSideIter[__left,__posY] := pinY - 100;
                1 : EachSideIter[__bottom,__posX] := pinX + 100;
                2 : EachSideIter[__right,__posY] := pinY + 100;
                3 : EachSideIter[__top,__posX] := pinX - 100;
            End;
            //CreatePin(pinX,pinY,pinSide+1,WireLen,rowstore[i][Idx_PortLabel],rowstore[i][Idx_NetLabel]);
            If GetVisibleState(rowstore[i][Idx_VisibleOpt]) = 1 Then
            Begin

                If (Length(plabel) = 0) Then
                Begin
                    If (Length(nlabel) >= 1) Then
                    Begin
                        IgnOpt := IgnOpt Or 4;
                    End
                    Else
                    Begin
                        IgnOpt := IgnOpt Or 7;
                    End;
                End
                Else
                Begin
                    If (Length(nlabel) = 0) Then
                    Begin
                        nlabel := plabel;
                    End;
                End;
                // printd('finalMk', nLabel + ' - ' + IntToStr(pinSide));
                CreatePin(pinX,pinY,pinSide+1,WireLen,plabel,nlabel, IgnOpt);
            End;
        End;
    End;

end;

Procedure main;
Var
    x,y : Integer;
    MiddleX, MiddleY, EndX, EndY, Side, iter, wLen : Integer;
    Len : Integer;
    dummyP, dummyN : String;
Begin
{
}
    {.... First Init .... }
    MessageClean();
    // demo code                        \
    PlaceViaCsv('C:\demo.csv');
    //CreatePin(100,-100,4,300,'dummy1','dm1', 0);
    //CreatePin(500,-600,2,300,'dummy2','dm2', 0);
End;