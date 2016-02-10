namespace SSH.Console

(*
    NOTE: We used a fork of Renci to fix a bug. I expect the bug has been fixed in Renci by now. If not let me know and I will make my Renci fork public. 
          The bug either causes the connection not to work or the console to hang. You'll know if it's there.
*)

open System
open System.Globalization
open System.Windows
open System.Windows.Controls
open System.Windows.Data
open System.Windows.Input
open System.Windows.Media
open System.Windows.Media.Animation
open System.Windows.Shapes
open Renci.SshNet
open Renci.SshNet.Common
open System.ComponentModel
open System.Diagnostics
open System.Reactive.Linq
open System.Text
open System.Threading.Tasks
open System.Windows.Data
open System.Windows.Input
open System.Runtime.InteropServices
open System.Windows.Input
open System.Text
open System.Collections.Generic
open System.ComponentModel
open System.Collections.Generic
open System.Text


[<AutoOpen>]
module Utils =

    open System
    open System.Threading.Tasks

    let inline defaultTo def x = match x with | null -> def | _ -> x

    let awaitTask (t : Task) : Async<unit> =
        let getException (t : Task) : Exception =
            if t.IsFaulted then
                let ex =
                    if t.Exception.InnerExceptions.Count = 1 then
                        t.Exception.InnerException
                    else
                        t.Exception :> Exception
                ex
            else
                null
        async {
            let! ex = t.ContinueWith getException |> Async.AwaitTask
            if ex <> null then raise ex
        }

    let ignoreErrorsInFn f x =
        try
            f x
        with
        | ex -> ()

    module Option =
        let defaultTo def x = match x with | None -> def | Some value -> value

    module List =
        let choosei f = List.mapi f >> List.choose id

        let getPartitionLocations selector =
            let partitionAnItem s t =
                let value = selector t
                match s with
                | [] -> [ value, 0, 1 ]
                | (value0, start, len) :: xs ->
                    if value = value0 then (value0, start, len + 1) :: xs
                    else (value, start + len, 1) :: (value0, start, len) :: xs
            Seq.fold partitionAnItem []

    type System.Windows.Threading.DispatcherObject with
        member this.InvokeOnDispatcher f =
            this.Dispatcher.BeginInvoke(Action f) |> ignore

type Coord = {
        X : int
        Y : int
    } with
    static member Origin = { Coord.X = 0; Y = 0 }

type Selection = {
        Start : Coord
        End : Coord
    } with
    static member Empty(coord) = { Selection.Start = coord; End = coord }
    static member Null = Selection.Empty(Coord.Origin)

    member this.IsEmpty = this.Start = this.End
    member this.First = min this.Start this.End
    member this.Last = max this.Start this.End
    member this.ExtendTo position =
        if position = this.End then this
        else { Selection.Start = this.Start; End = position }

module Colours =
    let ColourMap =
        dict [
            // Primary 3-bit (8 colors)
            (00, 0x000000); (01, 0x800000); (02, 0x008000); (03, 0x808000); (04, 0x000080); (05, 0x800080); (06, 0x008080); (07, 0xc0c0c0);
            // Equivalent "bright" versions of original 8 colors
            (08, 0x808080); (09, 0xff0000); (10, 0x00ff00); (11, 0xffff00); (12, 0x0000ff); (13, 0xff00ff); (14, 0x00ffff); (15, 0xffffff);
            // Strictly ascending
            (16, 0x000000); (17, 0x00005f); (18, 0x000087); (19, 0x0000af); (20, 0x0000d7); (21, 0x0000ff); (22, 0x005f00); (23, 0x005f5f);
            (24, 0x005f87); (25, 0x005faf); (26, 0x005fd7); (27, 0x005fff); (28, 0x008700); (29, 0x00875f); (30, 0x008787); (31, 0x0087af);
            (32, 0x0087d7); (33, 0x0087ff); (34, 0x00af00); (35, 0x00af5f); (36, 0x00af87); (37, 0x00afaf); (38, 0x00afd7); (39, 0x00afff);
            (40, 0x00d700); (41, 0x00d75f); (42, 0x00d787); (43, 0x00d7af); (44, 0x00d7d7); (45, 0x00d7ff); (46, 0x00ff00); (47, 0x00ff5f);
            (48, 0x00ff87); (49, 0x00ffaf); (50, 0x00ffd7); (51, 0x00ffff); (52, 0x5f0000); (53, 0x5f005f); (54, 0x5f0087); (55, 0x5f00af);
            (56, 0x5f00d7); (57, 0x5f00ff); (58, 0x5f5f00); (59, 0x5f5f5f); (60, 0x5f5f87); (61, 0x5f5faf); (62, 0x5f5fd7); (63, 0x5f5fff);
            (64, 0x5f8700); (65, 0x5f875f); (66, 0x5f8787); (67, 0x5f87af); (68, 0x5f87d7); (69, 0x5f87ff); (70, 0x5faf00); (71, 0x5faf5f);
            (72, 0x5faf87); (73, 0x5fafaf); (74, 0x5fafd7); (75, 0x5fafff); (76, 0x5fd700); (77, 0x5fd75f); (78, 0x5fd787); (79, 0x5fd7af);
            (80, 0x5fd7d7); (81, 0x5fd7ff); (82, 0x5fff00); (83, 0x5fff5f); (84, 0x5fff87); (85, 0x5fffaf); (86, 0x5fffd7); (87, 0x5fffff);
            (88, 0x870000); (89, 0x87005f); (90, 0x870087); (91, 0x8700af); (92, 0x8700d7); (93, 0x8700ff); (94, 0x875f00); (95, 0x875f5f);
            (96, 0x875f87); (97, 0x875faf); (98, 0x875fd7); (99, 0x875fff); (100, 0x878700); (101, 0x87875f); (102, 0x878787); (103, 0x8787af);
            (104, 0x8787d7); (105, 0x8787ff); (106, 0x87af00); (107, 0x87af5f); (108, 0x87af87); (109, 0x87afaf); (110, 0x87afd7); (111, 0x87afff);
            (112, 0x87d700); (113, 0x87d75f); (114, 0x87d787); (115, 0x87d7af); (116, 0x87d7d7); (117, 0x87d7ff); (118, 0x87ff00); (119, 0x87ff5f);
            (120, 0x87ff87); (121, 0x87ffaf); (122, 0x87ffd7); (123, 0x87ffff); (124, 0xaf0000); (125, 0xaf005f); (126, 0xaf0087); (127, 0xaf00af);
            (128, 0xaf00d7); (129, 0xaf00ff); (130, 0xaf5f00); (131, 0xaf5f5f); (132, 0xaf5f87); (133, 0xaf5faf); (134, 0xaf5fd7); (135, 0xaf5fff);
            (136, 0xaf8700); (137, 0xaf875f); (138, 0xaf8787); (139, 0xaf87af); (140, 0xaf87d7); (141, 0xaf87ff); (142, 0xafaf00); (143, 0xafaf5f);
            (144, 0xafaf87); (145, 0xafafaf); (146, 0xafafd7); (147, 0xafafff); (148, 0xafd700); (149, 0xafd75f); (150, 0xafd787); (151, 0xafd7af);
            (152, 0xafd7d7); (153, 0xafd7ff); (154, 0xafff00); (155, 0xafff5f); (156, 0xafff87); (157, 0xafffaf); (158, 0xafffd7); (159, 0xafffff);
            (160, 0xd70000); (161, 0xd7005f); (162, 0xd70087); (163, 0xd700af); (164, 0xd700d7); (165, 0xd700ff); (166, 0xd75f00); (167, 0xd75f5f);
            (168, 0xd75f87); (169, 0xd75faf); (170, 0xd75fd7); (171, 0xd75fff); (172, 0xd78700); (173, 0xd7875f); (174, 0xd78787); (175, 0xd787af);
            (176, 0xd787d7); (177, 0xd787ff); (178, 0xd7af00); (179, 0xd7af5f); (180, 0xd7af87); (181, 0xd7afaf); (182, 0xd7afd7); (183, 0xd7afff);
            (184, 0xd7d700); (185, 0xd7d75f); (186, 0xd7d787); (187, 0xd7d7af); (188, 0xd7d7d7); (189, 0xd7d7ff); (190, 0xd7ff00); (191, 0xd7ff5f);
            (192, 0xd7ff87); (193, 0xd7ffaf); (194, 0xd7ffd7); (195, 0xd7ffff); (196, 0xff0000); (197, 0xff005f); (198, 0xff0087); (199, 0xff00af);
            (200, 0xff00d7); (201, 0xff00ff); (202, 0xff5f00); (203, 0xff5f5f); (204, 0xff5f87); (205, 0xff5faf); (206, 0xff5fd7); (207, 0xff5fff);
            (208, 0xff8700); (209, 0xff875f); (210, 0xff8787); (211, 0xff87af); (212, 0xff87d7); (213, 0xff87ff); (214, 0xffaf00); (215, 0xffaf5f);
            (216, 0xffaf87); (217, 0xffafaf); (218, 0xffafd7); (219, 0xffafff); (220, 0xffd700); (221, 0xffd75f); (222, 0xffd787); (223, 0xffd7af);
            (224, 0xffd7d7); (225, 0xffd7ff); (226, 0xffff00); (227, 0xffff5f); (228, 0xffff87); (229, 0xffffaf); (230, 0xffffd7); (231, 0xffffff);
            // Gray-scale range
            (232, 0x080808); (233, 0x121212); (234, 0x1c1c1c); (235, 0x262626); (236, 0x303030); (237, 0x3a3a3a); (238, 0x444444); (239, 0x4e4e4e);
            (240, 0x585858); (241, 0x626262); (242, 0x6c6c6c); (243, 0x767676); (244, 0x808080); (245, 0x8a8a8a); (246, 0x949494); (247, 0x9e9e9e);
            (248, 0xa8a8a8); (249, 0xb2b2b2); (250, 0xbcbcbc); (251, 0xc6c6c6); (252, 0xd0d0d0); (253, 0xdadada); (254, 0xe4e4e4); (255, 0xeeeeee) ]

type Colour =
    | Black
    | Red
    | Green
    | Yellow
    | Blue
    | Magenta
    | Cyan
    | White
    | BrightBlack
    | BrightRed
    | BrightGreen
    | BrightYellow
    | BrightBlue
    | BrightMagenta
    | BrightCyan
    | BrightWhite
    | RGB of (byte * byte * byte)
    | Palette of byte with
    static member RgbFromInt (n : int) = (byte (n >>> 16) &&& 0xFFuy, byte (n >>> 8) &&& 0xFFuy, byte n &&& 0xFFuy)
    member this.AsRgb =
        match this with
        | Black -> Colour.RgbFromInt(Colours.ColourMap.[0])
        | Red -> Colour.RgbFromInt(Colours.ColourMap.[1])
        | Green -> Colour.RgbFromInt(Colours.ColourMap.[2])
        | Yellow -> Colour.RgbFromInt(Colours.ColourMap.[3])
        | Blue -> Colour.RgbFromInt(Colours.ColourMap.[4])
        | Magenta -> Colour.RgbFromInt(Colours.ColourMap.[5])
        | Cyan -> Colour.RgbFromInt(Colours.ColourMap.[6])
        | White -> Colour.RgbFromInt(Colours.ColourMap.[7])
        | BrightBlack -> Colour.RgbFromInt(Colours.ColourMap.[8])
        | BrightRed -> Colour.RgbFromInt(Colours.ColourMap.[9])
        | BrightGreen -> Colour.RgbFromInt(Colours.ColourMap.[10])
        | BrightYellow -> Colour.RgbFromInt(Colours.ColourMap.[11])
        | BrightBlue -> Colour.RgbFromInt(Colours.ColourMap.[12])
        | BrightMagenta -> Colour.RgbFromInt(Colours.ColourMap.[13])
        | BrightCyan -> Colour.RgbFromInt(Colours.ColourMap.[14])
        | BrightWhite -> Colour.RgbFromInt(Colours.ColourMap.[15])
        | RGB (r, g, b) -> (r, g, b)
        | Palette n -> Colour.RgbFromInt(Colours.ColourMap.[int n])


[<Flags>]
type InputModifiers =
    | None = 0
    | Shift = 1
    | Control = 2
    | Alt = 4

type InputCommand =
    | Character of char
    | Return
    | Tab
    | Backspace
    | Insert
    | Delete
    | Left
    | Right
    | Up
    | Down
    | PageUp
    | PageDown
    | Home
    | End
    | Escape
    | Function of (byte * InputModifiers)
    | PastedText of string

type FontWeight =
    | Normal
    | Bold
    | Light

[<Flags>]
type FontStyle =
    | Normal = 0x0
    | Oblique = 0x1
    | Italic = 0x2
    | Underline = 0x4
    | Blink = 0x8
    | Inverse = 0x10
    | Hidden = 0x20

[<Flags>]
type CursorStyle =
    | Blinking = 0
    | Steady = 1
    | Block = 0
    | Underline = 2
    | Bar = 4

type ScrollRegion = (int * int) option

type OutputCommand =
    | Character of char * ScrollRegion option
    | Tab
    | SetTabStop
    | ClearTabStop
    | ClearAllTabStops
    | MoveTo of Coord
    | MoveToColumn of int
    | MoveToRow of int
    | MoveLeft of int * ScrollRegion option
    | MoveRight of int * ScrollRegion option
    | MoveUp of int * ScrollRegion option
    | MoveDown of int * ScrollRegion option
    | MoveToLineStart
    | MoveToLineEnd
    | ScrollUp of int * ScrollRegion
    | ScrollDown of int * ScrollRegion
    | ClearScreen
    | ClearToScreenStart
    | ClearToScreenEnd
    | ClearLine
    | ClearToLineStart
    | ClearToLineEnd
    | InsertLines of int
    | DeleteLines of int
    | SetCurrentAttributesDefault
    | SetCurrentForegroundColour of Colour
    | ClearCurrentForegroundColour
    | SetCurrentBackgroundColour of Colour
    | ClearCurrentBackgroundColour
    | SetCurrentFontWeight of FontWeight
    | SetCurrentFontStyle of FontStyle
    | ClearCurrentFontStyle of FontStyle
    | SaveCursor of int
    | RestoreCursor of int
    | SetCursorStyle of CursorStyle
    | UseScreenBuffer of int
    | ScrollUpThroughBuffer of int
    | Bell
    | SetAudioVolume of float
    | SetWindowTitle of string
    | SetIconName of string


module XTerm =

    module ControlChars =
        /// <summary>Ignored on input (not stored in input buffer; see full duplex protocol).</summary>
        [<Literal>]
        let NUL = 0x00uy
        /// <summary>Transmit answerback message.</summary>
        [<Literal>]
        let ENQ = 0x05uy
        /// <summary>Sound bell tone from keyboard.</summary>
        [<Literal>]
        let BEL = 0x07uy
        /// <summary>Move the cursor to the left one character position, unless it is at the left margin, in which case no action occurs.</summary>
        [<Literal>]
        let BS = 0x08uy
        /// <summary>Move the cursor to the next tab stop, or to the right margin if no further tab stops are present on the line.</summary>
        [<Literal>]
        let HT = 0x09uy
        /// <summary>This code causes a line feed or a new line operation (see new line mode).</summary>
        [<Literal>]
        let LF = 0x0Auy
        /// <summary>Interpreted as LF.</summary>
        [<Literal>]
        let VT = 0x0Buy
        /// <summary>Interpreted as LF.</summary>
        [<Literal>]
        let FF = 0x0Cuy
        /// <summary>Move cursor to the left margin on the current line.</summary>
        [<Literal>]
        let CR = 0x0Duy
        /// <summary>Invoke G1 character set, as designated by SCS control sequence.</summary>
        [<Literal>]
        let SO = 0x0Euy
        /// <summary>Select G0 character set, as selected by ESC ( sequence.</summary>
        [<Literal>]
        let SI = 0x0Fuy
        /// <summary>Causes terminal to resume transmission.</summary>
        [<Literal>]
        let XON = 0x11uy
        /// <summary>Causes terminal to stop transmitted all codes except XOFF and XON.</summary>
        [<Literal>]
        let XOFF = 0x13uy
        /// <summary>If sent during a control sequence, the sequence is immediately terminated and not executed and the error character is displayed.</summary>
        [<Literal>]
        let CAN = 0x18uy
        /// <summary>Interpreted as CAN.</summary>
        [<Literal>]
        let SUB = 0x1Auy
        /// <summary>Invokes a control sequence.</summary>
        [<Literal>]
        let ESC = 0x1Buy
        [<Literal>]
        let FS = 0x1Cuy
        [<Literal>]
        let GS = 0x1Duy
        [<Literal>]
        let RS = 0x1Euy
        [<Literal>]
        let US = 0x1Fuy

        /// <summary>Ignored on input (not stored in input buffer).</summary>
        [<Literal>]
        let DEL = 0x7Fuy

        [<Literal>]
        let IND = 0x84uy
        [<Literal>]
        let NEL = 0x85uy
        [<Literal>]
        let HTS = 0x88uy
        [<Literal>]
        let RI = 0x8Duy
        [<Literal>]
        let SS2 = 0x8Euy
        [<Literal>]
        let SS3 = 0x8Fuy
        [<Literal>]
        let DCS = 0x90uy
        [<Literal>]
        let SPA = 0x96uy
        [<Literal>]
        let EPA = 0x97uy
        [<Literal>]
        let SOS = 0x98uy
        [<Literal>]
        let DECID = 0x9Auy
        [<Literal>]
        let CSI = 0x9Buy
        [<Literal>]
        let ST = 0x9Cuy
        [<Literal>]
        let OSC = 0x9Duy
        [<Literal>]
        let PM = 0x9Euy
        [<Literal>]
        let APC = 0x9Fuy


    let private getAsciiChar b = (Encoding.ASCII.GetChars [| b |]).[0]
    let private getCharAsciiCodes (c : char) = Encoding.ASCII.GetBytes [| c |] |> List.ofArray
    let private getStrAsciiCodes (s : string) = Encoding.ASCII.GetBytes s |> List.ofArray


    type CharacterSets =
        | DecSpecialCharacterLineDrawing
        | UnitedKingdom
        | UnitedStates
        | Dutch
        | Finnish
        | French
        | FrenchCanadian
        | German
        | Italian
        | NorwegianDanish
        | Spanish
        | Swedish
        | Swiss
    with
        static member FromCode =
            function
            | '0' -> DecSpecialCharacterLineDrawing
            | 'A' -> UnitedKingdom
            | 'B' -> UnitedStates
            | '4' -> Dutch
            | 'C'
            | '5' -> Finnish
            | 'R' -> French
            | 'Q' -> FrenchCanadian
            | 'K' -> German
            | 'Y' -> Italian
            | 'E'
            | '6' -> NorwegianDanish
            | 'Z' -> Spanish
            | 'H'
            | '7' -> Swedish
            | '=' -> Swiss
            | _ -> invalidArg "c" "Character did not represent a valid character set"

    type CustomModes =
        | KeypadAsCodes
        | Use8BitControls

    type State(scrollRegion : ScrollRegion,
                modes : Set<int>, customModes : Set<CustomModes>,
                characterSets : Map<int, CharacterSets>,
                incompleteCtrlSeq : CtrlSeq option, incompleteOSCtrlSeq : OSCtrlSeq option) =
        new() = State(None, Set.empty, Set.empty, Map.empty, None, None)

        member this.ScrollRegion = scrollRegion
        member this.ClearScrollRegion = State(None, modes, customModes, characterSets, None, None)
        member this.UpdateScrollRegion (top, afterBottom) =
            let top = top |> max 0
            let afterBottom = afterBottom |> max (top + 1)
            State(Some (top, afterBottom), modes, customModes, characterSets, None, None)
        member this.GetDeviceIndependantRow y =
            let topLine =
                if modes.Contains 6 && scrollRegion.IsSome then
                    fst scrollRegion.Value + 1
                else
                    1
            y - topLine |> max 0
        member this.GetMoveToCommand x y =
            [ { Coord.X = x - 1 |> max 0; Y = this.GetDeviceIndependantRow y } |> OutputCommand.MoveTo ]


        member this.IsInCtrlSeq = incompleteCtrlSeq.IsSome || incompleteOSCtrlSeq.IsSome
        member this.ClearCtrlSeq = State(scrollRegion, modes, customModes, characterSets, None, None)
        member this.UpdateCtrlSeq ctrlSeq = State(scrollRegion, modes, customModes, characterSets, Some ctrlSeq, incompleteOSCtrlSeq)
        member this.ProcessNextCtrlSeqChar =
            match incompleteCtrlSeq with
            | None ->
                match incompleteOSCtrlSeq with
                | None -> invalidOp "Can't ProcessNextCtrlSeqChar when incompleteCtrlSeq and incompleteOSCtrlSeq are None"
                | Some osCtrlSeq -> osCtrlSeq.ProcessNextChar this
            | Some ctrlSeq -> ctrlSeq.ProcessNextChar this

        member this.IsInOSCtrlSeq = incompleteOSCtrlSeq.IsSome
        member this.EndOSCtrlSeq =
            match incompleteOSCtrlSeq with
            | None -> invalidOp "Can't EndOSCtrlSeq when incompleteOSCtrlSeq is None"
            | Some osCtrlSeq ->
                State(scrollRegion, modes, customModes, characterSets, None, None),
                    osCtrlSeq.Execute
        member this.UpdateOSCtrlSeq osCtrlSeq = State(scrollRegion, modes, customModes, characterSets, None, Some osCtrlSeq)

        member this.SetMode nonAnsi mode =
            State(scrollRegion, modes.Add mode, customModes, characterSets, None, None),
                if not nonAnsi then
                    match mode with
                    | 2 -> []                       // Keyboard Action Mode (AM)
                    | 4 -> []                       // Insert Mode (IRM)
                    | 12 -> []                      // Send/receive (SRM)
                    | 20 -> []                      // Automatic Newline (LNM)
                    | _ -> []
                else
                    match mode with
                    | 1 -> []                       // Application Cursor Keys (DECCKM). 
                    | 2 -> []                       // Designate USASCII for character sets G0-G3 (DECANM), and set VT100 mode. 
                    | 3 -> []                       // 132 Column Mode (DECCOLM). 
                    | 4 -> []                       // Smooth (Slow) Scroll (DECSCLM). 
                    | 5 -> []                       // Reverse Video (DECSCNM). 
                    | 6 ->                          // Origin Mode (DECOM). 
                        this.GetMoveToCommand 1 1
                    | 7 -> []                       // Wraparound Mode (DECAWM). 
                    | 8 -> []                       // Auto-repeat Keys (DECARM). 
                    | 9 -> []                       // Send Mouse X & Y on button press. See the section Mouse Tracking. 
                    | 10 -> []                      // Show toolbar (rxvt). 
                    | 12 -> []                      // Start Blinking Cursor (att610). 
                    | 18 -> []                      // Print form feed (DECPFF). 
                    | 19 -> []                      // Set print extent to full screen (DECPEX). 
                    | 25 -> []                      // Show Cursor (DECTCEM). 
                    | 30 -> []                      // Show scrollbar (rxvt). 
                    | 35 -> []                      // Enable font-shifting functions (rxvt). 
                    | 38 -> []                      // Enter Tektronix Mode (DECTEK). 
                    | 40 -> []                      // Allow 80 → 132 Mode. 
                    | 41 -> []                      // more(1) fix (see curses resource). 
                    | 42 -> []                      // Enable Nation Replacement Character sets (DECNRCM). 
                    | 44 -> []                      // Turn On Margin Bell. 
                    | 45 -> []                      // Reverse-wraparound Mode. 
                    | 46 -> []                      // Start Logging. This is normally disabled by a compile-time option. 
                    | 47 ->                         // Use Alternate Screen Buffer. (This may be disabled by the titeInhibit resource).
                        [ OutputCommand.UseScreenBuffer 1 ]
                    | 66 -> []                      // Application keypad (DECNKM). 
                    | 67 -> []                      // Backarrow key sends backspace (DECBKM). 
                    | 69 -> []                      // Enable left and right margin mode (DECLRMM), VT420 and up. 
                    | 95 -> []                      // Do not clear screen when DECCOLM is set/reset (DECNCSM), VT510 and up. 
                    | 1000 -> []                    // Send Mouse X & Y on button press and release. See the section Mouse Tracking. 
                    | 1001 -> []                    // Use Hilite Mouse Tracking. 
                    | 1002 -> []                    // Use Cell Motion Mouse Tracking. 
                    | 1003 -> []                    // Use All Motion Mouse Tracking. 
                    | 1004 -> []                    // Send FocusIn/FocusOut events. 
                    | 1005 -> []                    // Enable UTF-8 Mouse Mode. 
                    | 1006 -> []                    // Enable SGR Mouse Mode. 
                    | 1007 -> []                    // Enable Alternate Scroll Mode. 
                    | 1010 -> []                    // Scroll to bottom on tty output (rxvt). 
                    | 1015 -> []                    // Enable urxvt Mouse Mode. 
                    | 1011 -> []                    // Scroll to bottom on key press (rxvt). 
                    | 1034 -> []                    // Interpret "meta" key, sets eighth bit. (enables the eightBitInput resource). 
                    | 1035 -> []                    // Enable special modifiers for Alt and NumLock keys. (This enables the numLock resource). 
                    | 1036 -> []                    // Send ESC when Meta modifies a key. (This enables the metaSendsEscape resource). 
                    | 1037 -> []                    // Send DEL from the editing-keypad Delete key. 
                    | 1039 -> []                    // Send ESC when Alt modifies a key. (This enables the altSendsEscape resource). 
                    | 1040 -> []                    // Keep selection even if not highlighted. (This enables the keepSelection resource). 
                    | 1041 -> []                    // Use the CLIPBOARD selection. (This enables the selectToClipboard resource). 
                    | 1042 -> []                    // Enable Urgency window manager hint when Control-G is received. (This enables the bellIsUrgent resource). 
                    | 1043 -> []                    // Enable raising of the window when Control-G is received. (enables the popOnBell resource). 
                    | 1047 ->                       // Use Alternate Screen Buffer. (This may be disabled by the titeInhibit resource).
                        [ OutputCommand.UseScreenBuffer 1; OutputCommand.ClearScreen ]
                    | 1048 ->                       // Save cursor as in DECSC. (This may be disabled by the titeInhibit resource).
                        [ OutputCommand.SaveCursor 0 ]
                    | 1049 ->                       // Save cursor as in DECSC and use Alternate Screen Buffer, clearing it first. (This may be disabled by the titeInhibit resource). This combines the effects of the 1047 and 1048 modes. Use this with terminfo-based applications rather than the 47 mode.
                        [ OutputCommand.SaveCursor 0; OutputCommand.UseScreenBuffer 1; OutputCommand.ClearScreen ]
                    | 1050 -> []                    // Set terminfo/termcap function-key mode. 
                    | 1051 -> []                    // Set Sun function-key mode. 
                    | 1052 -> []                    // Set HP function-key mode. 
                    | 1053 -> []                    // Set SCO function-key mode. 
                    | 1060 -> []                    // Set legacy keyboard emulation (X11R6). 
                    | 1061 -> []                    // Set VT220 keyboard emulation. 
                    | 2004 -> []                    // Set bracketed paste mode.
                    | _ -> []
        member this.ResetMode nonAnsi mode =
            State(scrollRegion, modes.Remove mode, customModes, characterSets, None, None),
                if not nonAnsi then
                    match mode with
                    | 2 -> []                       // Keyboard Action Mode (AM). 
                    | 4 -> []                       // Replace Mode (IRM). 
                    | 12 -> []                      // Send/receive (SRM). 
                    | 20 -> []                      // Normal Linefeed (LNM).
                    | _ -> []
                else
                    match mode with
                    | 1 -> []                       // Normal Cursor Keys (DECCKM). 
                    | 2 -> []                       // Designate VT52 mode (DECANM). 
                    | 3 -> []                       // 80 Column Mode (DECCOLM). 
                    | 4 -> []                       // Jump (Fast) Scroll (DECSCLM). 
                    | 5 -> []                       // Normal Video (DECSCNM). 
                    | 6 ->                          // Normal Cursor Mode (DECOM). 
                        this.GetMoveToCommand 1 1
                    | 7 -> []                       // No Wraparound Mode (DECAWM). 
                    | 8 -> []                       // No Auto-repeat Keys (DECARM). 
                    | 9 -> []                       // Don’t send Mouse X & Y on button press. 
                    | 10 -> []                      // Hide toolbar (rxvt). 
                    | 12 -> []                      // Stop Blinking Cursor (att610). 
                    | 18 -> []                      // Don’t print form feed (DECPFF). 
                    | 19 -> []                      // Limit print to scrolling region (DECPEX). 
                    | 25 -> []                      // Hide Cursor (DECTCEM). 
                    | 30 -> []                      // Don’t show scrollbar (rxvt). 
                    | 35 -> []                      // Disable font-shifting functions (rxvt). 
                    | 40 -> []                      // Disallow 80 → 132 Mode. 
                    | 41 -> []                      // No more(1) fix (see curses resource). 
                    | 42 -> []                      // Disable Nation Replacement Character sets (DECNRCM). 
                    | 44 -> []                      // Turn Off Margin Bell. 
                    | 45 -> []                      // No Reverse-wraparound Mode. 
                    | 46 -> []                      // Stop Logging. (This is normally disabled by a compile-time option). 
                    | 47 ->                         // Use Normal Screen Buffer. 
                        [ OutputCommand.UseScreenBuffer 0 ]
                    | 66 -> []                      // Numeric keypad (DECNKM). 
                    | 67 -> []                      // Backarrow key sends delete (DECBKM). 
                    | 69 -> []                      // Disable left and right margin mode (DECLRMM), VT420 and up. 
                    | 95 -> []                      // Clear screen when DECCOLM is set/reset (DECNCSM), VT510 and up. 
                    | 1000 -> []                    // Don’t send Mouse X & Y on button press and release. See the section Mouse Tracking. 
                    | 1001 -> []                    // Don’t use Hilite Mouse Tracking. 
                    | 1002 -> []                    // Don’t use Cell Motion Mouse Tracking. 
                    | 1003 -> []                    // Don’t use All Motion Mouse Tracking. 
                    | 1004 -> []                    // Don’t send FocusIn/FocusOut events. 
                    | 1005 -> []                    // Disable UTF-8 Mouse Mode. 
                    | 1006 -> []                    // Disable SGR Mouse Mode. 
                    | 1007 -> []                    // Disable Alternate Scroll Mode. 
                    | 1010 -> []                    // Don’t scroll to bottom on tty output (rxvt). 
                    | 1015 -> []                    // Disable urxvt Mouse Mode. 
                    | 1011 -> []                    // Don’t scroll to bottom on key press (rxvt). 
                    | 1034 -> []                    // Don’t interpret "meta" key. (This disables the eightBitInput resource). 
                    | 1035 -> []                    // Disable special modifiers for Alt and NumLock keys. (This disables the numLock resource). 
                    | 1036 -> []                    // Don’t send ESC when Meta modifies a key. (This disables the metaSendsEscape resource). 
                    | 1037 -> []                    // Send VT220 Remove from the editing-keypad Delete key. 
                    | 1039 -> []                    // Don’t send ESC when Alt modifies a key. (This disables the altSendsEscape resource). 
                    | 1040 -> []                    // Do not keep selection when not highlighted. (This disables the keepSelection resource). 
                    | 1041 -> []                    // Use the PRIMARY selection. (This disables the selectToClipboard resource). 
                    | 1042 -> []                    // Disable Urgency window manager hint when Control-G is received. (This disables the bellIsUrgent resource). 
                    | 1043 -> []                    // Disable raising of the window when Control-G is received. (This disables the popOnBell resource). 
                    | 1047 ->                       // Use Normal Screen Buffer, clearing screen first if in the Alternate Screen. (This may be disabled by the titeInhibit resource). 
                        [ OutputCommand.ClearScreen; OutputCommand.UseScreenBuffer 0 ]
                    | 1048 ->                       // Restore cursor as in DECRC. (This may be disabled by the titeInhibit resource). 
                        [ OutputCommand.RestoreCursor 0 ]
                    | 1049 ->                       // Use Normal Screen Buffer and restore cursor as in DECRC. (This may be disabled by the titeInhibit resource). This combines the effects of the 1047 and 1048 modes. Use this with terminfo-based applications rather than the 4 7 mode. 
                        [ OutputCommand.UseScreenBuffer 0; OutputCommand.RestoreCursor 0 ]
                    | 1050 -> []                    // Reset terminfo/termcap function-key mode. 
                    | 1051 -> []                    // Reset Sun function-key mode. 
                    | 1052 -> []                    // Reset HP function-key mode. 
                    | 1053 -> []                    // Reset SCO function-key mode. 
                    | 1060 -> []                    // Reset legacy keyboard emulation (X11R6). 
                    | 1061 -> []                    // Reset keyboard emulation to Sun/PC style. 
                    | 2004 -> []                    // Reset bracketed paste mode.
                    | _ -> []

        member this.SetCustomMode mode = State(scrollRegion, modes, customModes.Add mode, characterSets, None, None)
        member this.ResetCustomMode mode = State(scrollRegion, modes, customModes.Remove mode, characterSets, None, None)

        member this.ScrollRegionIfWrapping = if modes.Contains 7 then Some scrollRegion else None

        member this.SetCharSet setNo cs = State(scrollRegion, modes, customModes, characterSets.Add(setNo, cs), None, None)

    and CtrlSeq(prefix : string, reversedParameters : int list, nonAnsi : bool) =
        new(prefix : string) = CtrlSeq(prefix, [], false)
        new() = CtrlSeq("")
        member this.ProcessNextChar (state : State) (c : char) : (State * OutputCommand list) =
            let moveTo x y = state.GetMoveToCommand x y
            if prefix = "" then
                match c with
                | 'M' -> state.ClearCtrlSeq, [ OutputCommand.MoveUp (1, Some state.ScrollRegion) ]      // Reverse Index (RI)
                | 'D' -> state.ClearCtrlSeq, [ OutputCommand.MoveDown (1, Some state.ScrollRegion) ]    // Index (IND)
                | 'E' -> state.ClearCtrlSeq, [ OutputCommand.MoveDown (1, Some state.ScrollRegion); OutputCommand.MoveToLineStart ]    // Next Line (NEL)
                | 'H' -> state.ClearCtrlSeq, [ OutputCommand.SetTabStop ]                               // Tab Set (HTS)
                | '[' -> CtrlSeq(c.ToString(), [ 0 ], nonAnsi) |> state.UpdateCtrlSeq, []                     // Control Sequence Introducer (CSI)
                | 'N' -> state.ClearCtrlSeq, []     // Single Shift Select of G2 Character Set (affects next character only) (SS2)
                | 'O' -> state.ClearCtrlSeq, []     // Single Shift Select of G3 Character Set (affects next character only) (SS3)
                | 'P' -> state.ClearCtrlSeq, []     // Device Control String (DCS)
                | 'V' -> state.ClearCtrlSeq, []     // Start of Guarded Area (SPA)
                | 'W' -> state.ClearCtrlSeq, []     // End of Guarded Area (EPA)
                | 'X' -> state.ClearCtrlSeq, []     // Start of String (SOS)
                | '|' -> state.ClearCtrlSeq, []     // Return Terminal ID (DECID)
                | ']' -> OSCtrlSeq() |> state.UpdateOSCtrlSeq, []     // Operating System Command (OSC)
                | '\\' ->                           // String Terminator (ST)
                    if state.IsInOSCtrlSeq then
                        state.EndOSCtrlSeq
                    else
                        state.ClearCtrlSeq, []
                | '^' -> state.ClearCtrlSeq, []     // Privacy Message (PM)
                | '_' -> state.ClearCtrlSeq, []     // Application Program Command (APC)
                | '6' -> state.ClearCtrlSeq, []     // Back Index (DECBI)
                | '7' -> state.ClearCtrlSeq, [ OutputCommand.SaveCursor 0 ]     // Save Cursor (DECSC)
                | '8' -> state.ClearCtrlSeq, [ OutputCommand.RestoreCursor 0 ]     // Restore Cursor (DECRC)
                | '9' -> state.ClearCtrlSeq, []     // Forward Index (DECFI)
                | '=' -> state.SetCustomMode KeypadAsCodes, []      // Keypad Application Mode (DECKPAM)
                | '>' -> state.ResetCustomMode KeypadAsCodes, []    // Keypad Numeric Mode (DECKPNM)
                | 'C' -> state.ClearCtrlSeq, []     // Full Reset (RIS)
                | 'n' -> state.ClearCtrlSeq, []     // Invoke the G2 Character Set as GL (LS2).
                | 'o' -> state.ClearCtrlSeq, []     // Invoke the G3 Character Set as GL (LS3).
                // | '|' -> state.ClearCtrlSeq, []     // Invoke the G3 Character Set as GR (LS3R).
                | '}' -> state.ClearCtrlSeq, []     // Invoke the G2 Character Set as GR (LS2R).
                | '~' -> state.ClearCtrlSeq, []     // Invoke the G1 Character Set as GR (LS1R).
                | _ -> CtrlSeq(c.ToString(), [], nonAnsi) |> state.UpdateCtrlSeq, []
            elif not reversedParameters.IsEmpty && c >= '0' && c <= '9' then
                CtrlSeq(prefix, reversedParameters.Head * 10 + (int c - int '0') :: reversedParameters.Tail, nonAnsi) |> state.UpdateCtrlSeq, []
            elif not reversedParameters.IsEmpty && c = ';' then
                CtrlSeq(prefix, 0 :: reversedParameters, nonAnsi) |> state.UpdateCtrlSeq, []
            elif reversedParameters = [ 0 ] && c = '?' then
                CtrlSeq(prefix, reversedParameters, true) |> state.UpdateCtrlSeq, []
            else
            let parameters : int list = List.rev reversedParameters
            let chooseParams (f : int -> OutputCommand option) = state.ClearCtrlSeq, List.choose f parameters
            let collectParams (f : int -> OutputCommand list) = state.ClearCtrlSeq, List.collect f parameters
            let foldParams (f : State -> int -> State * OutputCommand list) =
                let applyFAndAppendCmds (state, cmds) p =
                    let (newState, newCmds) = f state p
                    newState, cmds @ newCmds
                List.fold applyFAndAppendCmds (state, []) parameters
            let sumMinOneParams() = List.sumBy (max 1) parameters
            match prefix with
            | "[" ->
                match c with
                | 'b'                   // Repeat the preceding graphic character (REP)
                | 'X'                   // Erase Characters (ECH)
                | 'P'                   // Delete Characters (DCH)
                | '@' -> collectParams (fun p -> OutputCommand.Character (' ', state.ScrollRegionIfWrapping) |> List.replicate p)   // Insert Blank Character(s) (ICH)
                | 'D' -> state.ClearCtrlSeq, [ OutputCommand.MoveLeft (sumMinOneParams(), None) ]       // Cursor Backward (CUB)
                | 'a'                                                                                   // Character Position Relative (HPR)
                | 'C' -> state.ClearCtrlSeq, [ OutputCommand.MoveRight (sumMinOneParams(), None) ]      // Cursor Forward (CUF)
                | 'A' -> state.ClearCtrlSeq, [ OutputCommand.MoveUp (sumMinOneParams(), None) ]         // Cursor Up (CUU)
                | 'e'                                                                                   // Line Position Relative (VPR)
                | 'B' -> state.ClearCtrlSeq, [ OutputCommand.MoveDown (sumMinOneParams(), None) ]       // Cursor Down (CUD)
                | 'E' -> state.ClearCtrlSeq, [ OutputCommand.MoveDown (sumMinOneParams(), Some state.ScrollRegion) ]    // Cursor Next Line (CNL)
                | 'F' -> state.ClearCtrlSeq, [ OutputCommand.MoveUp (sumMinOneParams(), Some state.ScrollRegion) ]    // Cursor Preceding Line (CPL)
                | 'I' -> state.ClearCtrlSeq, List.replicate (sumMinOneParams()) OutputCommand.Tab       // Cursor Forward Tabulation (CHT)
                | 'Z' -> state.ClearCtrlSeq, List.replicate (-sumMinOneParams()) OutputCommand.Tab       // Cursor Forward Tabulation (CHT)
                | 'J' ->
                    if nonAnsi then         // Selective Erase in Display (DECSED)
                        collectParams (function
                            | 0 -> [ OutputCommand.ClearToScreenEnd ]
                            | 1 -> [ OutputCommand.ClearToScreenStart ]
                            | 2 -> [ OutputCommand.ClearScreen ]
                            | _ -> [])
                    else                    // Erase in Display (ED)
                        collectParams (function
                            | 0 -> [ OutputCommand.ClearToScreenEnd ]
                            | 1 -> [ OutputCommand.ClearToScreenStart ]
                            | 2 -> [ OutputCommand.ClearScreen ]
                            | _ -> [])
                | 'K' ->
                    if nonAnsi then         // Selective Erase in Line (DECSEL)
                        collectParams (function
                            | 0 -> [ OutputCommand.ClearToLineEnd ]
                            | 1 -> [ OutputCommand.ClearToLineStart ]
                            | 2 -> [ OutputCommand.ClearLine ]
                            | _ -> [])
                    else                    // Erase in Line (EL)
                        collectParams (function
                            | 0 -> [ OutputCommand.ClearToLineEnd ]
                            | 1 -> [ OutputCommand.ClearToLineStart ]
                            | 2 -> [ OutputCommand.ClearLine ]
                            | _ -> [])
                | 'G'                       // Cursor Character Absolute (CHA)
                | '`' ->                    // Character Position Absolute (HPA)
                    state.ClearCtrlSeq,
                        match parameters with
                        | [ x ] -> [ x - 1 |> max 0 |> OutputCommand.MoveToColumn ]
                        | _ -> []   // Too many parameters
                | 'd' ->                    // Character Position Absolute (HPA)
                    state.ClearCtrlSeq,
                        match parameters with
                        | [ y ] -> [ state.GetDeviceIndependantRow y |> OutputCommand.MoveToRow ]
                        | _ -> []   // Too many parameters
                | 'H'                       // Cursor Position (CUP)
                | 'f' ->                    // Horizontal and Vertical Position (HVP)
                    state.ClearCtrlSeq,
                        match parameters with
                        | [ y; x ] -> moveTo x y
                        | [ y ] -> moveTo 1 y
                        | _ -> []   // Too many parameters
                | 'L' -> state.ClearCtrlSeq, [ sumMinOneParams() |> OutputCommand.InsertLines ]         // Insert Lines (IL)
                | 'M' -> state.ClearCtrlSeq, [ sumMinOneParams() |> OutputCommand.DeleteLines ]         // Delete Lines (DL)
                | 'S' -> state.ClearCtrlSeq, [ OutputCommand.ScrollUp (sumMinOneParams(), state.ScrollRegion) ]     // Scroll Up (SU)
                | 'T' -> state.ClearCtrlSeq, [ OutputCommand.ScrollDown (sumMinOneParams(), state.ScrollRegion) ]   // Scroll Down (SD)
                | 'm' ->                    // Character Attributes (SGR)
                    match parameters with
                    | [ 38; 2; r; g; b ] ->
                        state.ClearCtrlSeq,
                            [ RGB (byte (min 255 r), byte (min 255 g), byte (min 255 b))
                                |> OutputCommand.SetCurrentForegroundColour ]
                    | [ 38; 5; idx ] ->
                        state.ClearCtrlSeq,
                            [ Palette (byte (min 255 idx)) |> OutputCommand.SetCurrentForegroundColour ]
                    | [ 48; 2; r; g; b ] ->
                        state.ClearCtrlSeq,
                            [ RGB (byte (min 255 r), byte (min 255 g), byte (min 255 b))
                                |> OutputCommand.SetCurrentBackgroundColour ]
                    | [ 48; 5; idx ] ->
                        state.ClearCtrlSeq,
                            [ Palette (byte (min 255 idx)) |> OutputCommand.SetCurrentBackgroundColour ]
                    | _ -> 
                    chooseParams (function
                        | 0 -> OutputCommand.SetCurrentAttributesDefault |> Some
                        | 1 -> OutputCommand.SetCurrentFontWeight Bold |> Some
                        | 2 -> OutputCommand.SetCurrentFontWeight Light |> Some
                        | 4 -> OutputCommand.SetCurrentFontStyle FontStyle.Underline |> Some
                        | 5 -> OutputCommand.SetCurrentFontStyle FontStyle.Blink |> Some
                        | 7 -> OutputCommand.SetCurrentFontStyle FontStyle.Inverse |> Some
                        | 8 -> OutputCommand.SetCurrentFontStyle FontStyle.Hidden |> Some
                        | 22 -> OutputCommand.SetCurrentFontWeight Normal |> Some
                        | 24 -> OutputCommand.ClearCurrentFontStyle FontStyle.Underline |> Some
                        | 25 -> OutputCommand.ClearCurrentFontStyle FontStyle.Blink |> Some
                        | 27 -> OutputCommand.ClearCurrentFontStyle FontStyle.Inverse |> Some
                        | 28 -> OutputCommand.ClearCurrentFontStyle FontStyle.Hidden |> Some
                        | 30 -> OutputCommand.SetCurrentForegroundColour Black |> Some
                        | 31 -> OutputCommand.SetCurrentForegroundColour Red |> Some
                        | 32 -> OutputCommand.SetCurrentForegroundColour Green |> Some
                        | 33 -> OutputCommand.SetCurrentForegroundColour Yellow |> Some
                        | 34 -> OutputCommand.SetCurrentForegroundColour Blue |> Some
                        | 35 -> OutputCommand.SetCurrentForegroundColour Magenta |> Some
                        | 36 -> OutputCommand.SetCurrentForegroundColour Cyan |> Some
                        | 37 -> OutputCommand.SetCurrentForegroundColour White |> Some
                        | 39 -> OutputCommand.ClearCurrentForegroundColour |> Some
                        | 40 -> OutputCommand.SetCurrentBackgroundColour Black |> Some
                        | 41 -> OutputCommand.SetCurrentBackgroundColour Red |> Some
                        | 42 -> OutputCommand.SetCurrentBackgroundColour Green |> Some
                        | 43 -> OutputCommand.SetCurrentBackgroundColour Yellow |> Some
                        | 44 -> OutputCommand.SetCurrentBackgroundColour Blue |> Some
                        | 45 -> OutputCommand.SetCurrentBackgroundColour Magenta |> Some
                        | 46 -> OutputCommand.SetCurrentBackgroundColour Cyan |> Some
                        | 47 -> OutputCommand.SetCurrentBackgroundColour White |> Some
                        | 49 -> OutputCommand.ClearCurrentBackgroundColour |> Some
                        | 90 -> OutputCommand.SetCurrentForegroundColour BrightBlack |> Some
                        | 91 -> OutputCommand.SetCurrentForegroundColour BrightRed |> Some
                        | 92 -> OutputCommand.SetCurrentForegroundColour BrightGreen |> Some
                        | 93 -> OutputCommand.SetCurrentForegroundColour BrightYellow |> Some
                        | 94 -> OutputCommand.SetCurrentForegroundColour BrightBlue |> Some
                        | 95 -> OutputCommand.SetCurrentForegroundColour BrightMagenta |> Some
                        | 96 -> OutputCommand.SetCurrentForegroundColour BrightCyan |> Some
                        | 97 -> OutputCommand.SetCurrentForegroundColour BrightWhite |> Some
                        | 100 -> OutputCommand.SetCurrentBackgroundColour BrightBlack |> Some
                        | 101 -> OutputCommand.SetCurrentBackgroundColour BrightRed |> Some
                        | 102 -> OutputCommand.SetCurrentBackgroundColour BrightGreen |> Some
                        | 103 -> OutputCommand.SetCurrentBackgroundColour BrightYellow |> Some
                        | 104 -> OutputCommand.SetCurrentBackgroundColour BrightBlue |> Some
                        | 105 -> OutputCommand.SetCurrentBackgroundColour BrightMagenta |> Some
                        | 106 -> OutputCommand.SetCurrentBackgroundColour BrightCyan |> Some
                        | 107 -> OutputCommand.SetCurrentBackgroundColour BrightWhite |> Some
                        | _ -> None)
                | 'g' ->    // TBC - Tabulation Clear
                    chooseParams (function
                        | 0 -> Some OutputCommand.ClearTabStop
                        | 3 -> Some OutputCommand.ClearAllTabStops
                        | _ -> None)
                | 'l' ->    // Reset Mode (RM)
                    foldParams (fun state mode -> state.ResetMode nonAnsi mode)
                | 'h' ->    // Set Mode (SM)
                    foldParams (fun state mode -> state.SetMode nonAnsi mode)
                | 'r' ->
                    match parameters with
                    | [ top; bottom ] -> state.UpdateScrollRegion (top - 1, bottom), moveTo 1 1
                    | _ -> state.ClearScrollRegion, moveTo 1 1
                | ' '
                | '!'
                | '>' -> CtrlSeq(prefix + c.ToString(), reversedParameters, nonAnsi) |> state.UpdateCtrlSeq, []
                | _ -> state.ClearCtrlSeq, []
            | "[>" ->
                match c with
                | 'm' -> state.ClearCtrlSeq, []                         // Set or reset resource-values used by xterm to decide whether to construct escape sequences holding information about the modifiers pressed with a given key
                | 'p' -> state.ClearCtrlSeq, []                         // Set resource value pointerMode. This is used by xterm to decide whether to hide the pointer cursor as the user types.
                | _ -> state.ClearCtrlSeq, []
            | "[!" ->
                match c with
                | 'p' -> state.ClearCtrlSeq, []                         // Soft terminal reset (DECSTR)
                | _ -> state.ClearCtrlSeq, []
            | "[ " ->
                match c with
                | 'q' ->                                                // Set cursor style (DECSCUSR, VT520)
                    chooseParams(function
                    | 0 -> CursorStyle.Blinking ||| CursorStyle.Block |> OutputCommand.SetCursorStyle |> Some
                    | 1 -> CursorStyle.Blinking ||| CursorStyle.Block |> OutputCommand.SetCursorStyle |> Some
                    | 2 -> CursorStyle.Steady ||| CursorStyle.Block |> OutputCommand.SetCursorStyle |> Some
                    | 3 -> CursorStyle.Blinking ||| CursorStyle.Underline |> OutputCommand.SetCursorStyle |> Some
                    | 4 -> CursorStyle.Steady ||| CursorStyle.Underline |> OutputCommand.SetCursorStyle |> Some
                    | 5 -> CursorStyle.Blinking ||| CursorStyle.Bar |> OutputCommand.SetCursorStyle |> Some
                    | 6 -> CursorStyle.Steady ||| CursorStyle.Bar |> OutputCommand.SetCursorStyle |> Some
                    | _ -> None)
                | _ -> state.ClearCtrlSeq, []
            | " " ->
                match c with
                | 'F' -> state.ResetCustomMode Use8BitControls, []      // 7-bit controls (S7C1T)
                | 'G' -> state.SetCustomMode Use8BitControls, []        // 8-bit controls (S8C1T)
                | 'L' -> state.ClearCtrlSeq, []                         // Set ANSI conformance level 1 (dpANS X3.134.1)
                | 'M' -> state.ClearCtrlSeq, []                         // Set ANSI conformance level 2 (dpANS X3.134.1)
                | 'N' -> state.ClearCtrlSeq, []                         // Set ANSI conformance level 3 (dpANS X3.134.1)
                | _ -> state.ClearCtrlSeq, []
            | "#" ->
                match c with
                | '3' -> state.ClearCtrlSeq, []                         // DEC double-height line, top half (DECDHL)
                | '4' -> state.ClearCtrlSeq, []                         // DEC double-height line, bottom half (DECDHL)
                | '5' -> state.ClearCtrlSeq, []                         // DEC single-width line (DECSWL)
                | '6' -> state.ClearCtrlSeq, []                         // DEC double-width line (DECDWL)
                | '8' -> state.ClearCtrlSeq, []                         // DEC Screen Alignment Test (DECALN)
                | _ -> state.ClearCtrlSeq, []
            | "%" ->
                match c with
                | '@' -> state.ClearCtrlSeq, []                         // Select default ISO 8859-1 character set (ISO 2022)
                | 'G' -> state.ClearCtrlSeq, []                         // Select UTF-8 character set (ISO 2022)
                | _ -> state.ClearCtrlSeq, []
            | "(" ->
                state.SetCharSet 0 (CharacterSets.FromCode c), []       // Designate G0 Character Set (ISO 2022, VT100)
            | ")"
            | "-" ->
                state.SetCharSet 1 (CharacterSets.FromCode c), []       // Designate G0 Character Set (ISO 2022, VT100)
            | "*"
            | "." ->
                state.SetCharSet 2 (CharacterSets.FromCode c), []       // Designate G0 Character Set (ISO 2022, VT100)
            | "+"
            | "/" ->
                state.SetCharSet 3 (CharacterSets.FromCode c), []       // Designate G0 Character Set (ISO 2022, VT100)
            | _ ->
                state.ClearCtrlSeq, []

    and OSCtrlSeq(parameter : int, text : string) =
        new() = OSCtrlSeq(0, null)
        member this.ProcessNextChar (state : State) (c : char) : (State * OutputCommand list) =
            if text = null && c >= '0' && c <= '9' then
                OSCtrlSeq(parameter * 10 + (int c - int '0'), null) |> state.UpdateOSCtrlSeq, []
            elif text = null && c = ';' then
                OSCtrlSeq(parameter, String.Empty) |> state.UpdateOSCtrlSeq, []
            else
                OSCtrlSeq(parameter, text + c.ToString()) |> state.UpdateOSCtrlSeq, []
        member this.Execute =
            match parameter with
            | 0 -> [ OutputCommand.SetIconName text; OutputCommand.SetWindowTitle text ]
            | 1 -> [ OutputCommand.SetIconName text ]
            | 2 -> [ OutputCommand.SetWindowTitle text ]
            | _ -> []


    let Process (outputFromServer : IObservable<byte[]>, inputFromConsole : IObservable<InputCommand>) : (IObservable<byte[]> * IObservable<OutputCommand seq>) =

        let processInput : InputCommand -> byte list =
            function
            | InputCommand.Character char -> getCharAsciiCodes char
            | InputCommand.Return -> [ ControlChars.CR ]
            | InputCommand.Tab -> [ ControlChars.HT ]
            | InputCommand.Backspace -> [ ControlChars.BS ]
            | InputCommand.Insert -> ControlChars.ESC :: getStrAsciiCodes "[2~"
            | InputCommand.Delete -> ControlChars.ESC :: getStrAsciiCodes "[3~"
            | InputCommand.Left -> ControlChars.ESC :: getStrAsciiCodes "[D"
            | InputCommand.Right -> ControlChars.ESC :: getStrAsciiCodes "[C"
            | InputCommand.Up -> ControlChars.ESC :: getStrAsciiCodes "[A"
            | InputCommand.Down -> ControlChars.ESC :: getStrAsciiCodes "[B"
            | InputCommand.PageUp -> ControlChars.ESC :: getStrAsciiCodes "[5~"
            | InputCommand.PageDown -> ControlChars.ESC :: getStrAsciiCodes "[6~"
            | InputCommand.Home -> ControlChars.ESC :: getStrAsciiCodes "[1~"
            | InputCommand.End -> ControlChars.ESC :: getStrAsciiCodes "[4~"
            | InputCommand.Escape -> [ ControlChars.ESC ]
            | InputCommand.Function (value, mods) ->
                let getModedString (n : int) =
                    if mods = InputModifiers.None then
                        ControlChars.ESC :: getStrAsciiCodes (String.Format("[{0}~", n))
                    else
                        let modCode =
                            (if mods.HasFlag InputModifiers.Shift then 1 else 0)
                                + (if mods.HasFlag InputModifiers.Alt then 2 else 0)
                                + (if mods.HasFlag InputModifiers.Control then 4 else 0)
                                + 1
                        ControlChars.ESC :: getStrAsciiCodes (String.Format("[{0};{1}~", n, modCode))
                match value with
                | 1uy -> ControlChars.SS3 :: getStrAsciiCodes "P"
                | 2uy -> ControlChars.SS3 :: getStrAsciiCodes "Q"
                | 3uy -> ControlChars.SS3 :: getStrAsciiCodes "R"
                | 4uy -> ControlChars.SS3 :: getStrAsciiCodes "S"
                | 5uy -> getModedString 15
                | 6uy -> getModedString 17
                | 7uy -> getModedString 18
                | 8uy -> getModedString 19
                | 9uy -> getModedString 20
                | 10uy -> getModedString 21
                | 11uy -> getModedString 23
                | 12uy -> getModedString 24
                | _ -> []
            | InputCommand.PastedText text -> text.Replace("\r\n", "\r").Replace("\n", "\r") |> getStrAsciiCodes

        let processOutputByte (state : State, _) (c : byte) : (State * OutputCommand list) =
            match c with
            | ControlChars.DEL -> state, []
            | ControlChars.BS -> state, [ OutputCommand.MoveLeft (1, state.ScrollRegionIfWrapping) ]  // Backspace (BS)
            | ControlChars.HT -> state, [ OutputCommand.Tab ]                                       // Horizontal Tab (HT)
            | ControlChars.VT                                                                       // Vertical Tab (VT)
            | ControlChars.FF                                                                       // Form Feed or New Page (NP)
            | ControlChars.LF -> state, [ OutputCommand.MoveDown (1, Some state.ScrollRegion) ]     // Line Feed or New Line (NL)
            | ControlChars.CR -> state, [ OutputCommand.MoveToLineStart ]                           // Carriage Return (CR)
            | ControlChars.ESC -> CtrlSeq() |> state.UpdateCtrlSeq, []
            | ControlChars.CAN
            | ControlChars.SUB -> state.ClearCtrlSeq, []
            | ControlChars.ENQ -> state, []                                                         // Return Terminal Status (ENQ)
            | ControlChars.SI -> state, []              // Shift In (SI) Switch to Standard Character Set (invoke the G0 character set) (the default).
            | ControlChars.SO -> state, []              // Shift Out (SO) Switch to Alternate Character Set (invoke the G1 character set)
            | ControlChars.IND -> state, [ OutputCommand.MoveDown (1, Some state.ScrollRegion) ]    // Index (IND)
            | ControlChars.NEL -> state, [ OutputCommand.MoveDown (1, Some state.ScrollRegion); OutputCommand.MoveToLineStart ]    // Next Line (NEL)
            | ControlChars.HTS -> state, [ OutputCommand.SetTabStop ]                               // Tab Set (HTS)
            | ControlChars.RI -> state, [ OutputCommand.MoveUp (1, Some state.ScrollRegion) ]       // Reverse Index (RI)
            | ControlChars.CSI -> CtrlSeq("[") |> state.UpdateCtrlSeq, []
            | ControlChars.OSC ->                       // Operating System Command (OSC)
                OSCtrlSeq() |> state.UpdateOSCtrlSeq, []
            | ControlChars.ST ->
                if state.IsInOSCtrlSeq then
                    state.EndOSCtrlSeq
                else
                    state, []
            | ControlChars.BEL ->                                                                   // Bell (BEL)
                if state.IsInOSCtrlSeq then
                    state.EndOSCtrlSeq
                else
                    state, [ OutputCommand.Bell ]
            | _ when c >= 0x20uy && state.IsInCtrlSeq ->
                state.ProcessNextCtrlSeqChar (getAsciiChar c)
            | _ when c >= 0x20uy && c < 0x7Fuy ->
                state, [ OutputCommand.Character (getAsciiChar c, state.ScrollRegionIfWrapping) ]
            | _ -> state, []

        let processOutputData (state : State, _) (data : byte[]) : (State * OutputCommand seq) =
            let processed = Array.scan processOutputByte (state, []) data
            fst processed.[processed.Length - 1],
                processed |> Seq.map snd |> Seq.concat

        let preamble =
            Observable.ToObservable [
                    //[ TelnetCommands.IAC; TelnetCommands.WILL; TelnetOptions.WindowSize ]  // NAWS
                ]
        Observable.map processInput inputFromConsole |> Observable.merge preamble
                |> Observable.filter (List.isEmpty >> not) |> Observable.map Array.ofList,
            outputFromServer |> Observable.scan processOutputData (State(), Seq.empty) |> Observable.map snd

type LineBuffer<'Char>(originalWidth : int, maxLength : int, blankChar : 'Char) =
    let mutable width = originalWidth
    let mutable lines : 'Char[] List = List<'Char[]>()

    let newLine() = Array.create width blankChar

    member this.Width
        with get() = width
        and set value = width <- value

    member this.Count = lines.Count

    member this.Item
        with get(x : int, y : int) =
            if y < lines.Count && x < lines.[y].Length then
                lines.[y].[x]
            else blankChar
        and set (x : int, y : int) (c : 'Char) =
            if y >= lines.Count then
                // Add as many extra lines as y requires (these will be removed when next scroll if y > maxLength)
                lines.AddRange(Seq.init (y + 1 - lines.Count) (fun _ -> newLine()))
            if x >= lines.[y].Length then
                if x >= width then invalidArg "x" "x must be less than the current width"
                // Extend width of line to match new console size
                let newLine = newLine()
                lines.[y].CopyTo(newLine, 0)
                lines.[y] <- newLine
            lines.[y].[x] <- c

    member this.Line(y : int) : 'Char[] =
        if y < lines.Count then lines.[y]   // Return whole array even though it may not match width
        else Array.empty

    member this.InsertLines (count : int) (insertPoint : int) =
        lines.InsertRange(insertPoint, Seq.init count (fun _ -> newLine()))
    member this.DeleteLines (count : int) (removePoint : int) =
        lines.RemoveRange(removePoint, count)
    member this.TrimToMaxLength() =
        // Trim buffer to maxLength lines
        let extraLines = lines.Count - maxLength
        if extraLines > 0 then
            lines.RemoveRange(maxLength, extraLines)

    member this.ClearLine (y : int) =
        if y < lines.Count then
            lines.[y] <- newLine()
    member this.ClearLineToStart (inclusiveX : int, y : int) =
        if y < lines.Count then
            let line = lines.[y]
            for x in 0 .. inclusiveX do line.[x] <- blankChar
    member this.ClearLineToEnd (inclusiveX : int, y : int) =
        if y < lines.Count then
            let line = lines.[y]
            for x in inclusiveX .. line.Length - 1 do line.[x] <- blankChar

type Attributes = {
    Background : Colour option
    Foreground : Colour option
    FontStyle : FontStyle
    FontWeight : FontWeight
} with
    static member Default = { Background = None; Foreground = None; FontStyle = FontStyle.Normal; FontWeight = FontWeight.Normal }

type Display(startupSize : Coord, maxBufferLines : int) as this =
    let mutable size = startupSize
    let mutable currentBufferIdx = 0
    let newBuffer() = LineBuffer<char * Attributes>(size.X, maxBufferLines, (' ', Attributes.Default))
    let mutable buffer = newBuffer()
    let buffers = new Dictionary<int, LineBuffer<char * Attributes>>()
    let mutable cursorPosition = Coord.Origin
    let savedCursorPositions = new Dictionary<int, Coord>()
    let mutable selection = Selection.Null
    let mutable bufferScrollPosition = 0
    let mutable currentAttributes = Attributes.Default
    let mutable windowTitle : string = null
    let mutable iconName : string = null
    // Screen buffer has line 0 at the bottom
    let reverseY (y : int) = (size.Y - 1) - y
    let isCoordOutsideSize (coord : Coord) = coord.X > size.X || coord.Y >= size.Y
    let moveCoordInsideSize (coord : Coord) = { Coord.X = min coord.X size.X; Y = min coord.Y (size.Y - 1) }

    let defaultTabStop = 8
    let mutable manualTabStops : int list = []

    let propertyChanged = Event<_, _>()
    let onPropertyChanged propertyName =
        propertyChanged.Trigger(this, PropertyChangedEventArgs(propertyName))
    let onLinesChanged() = onPropertyChanged "Lines"


    member this.CursorPosition
        with get() = cursorPosition
        and private set value =
            if this.CursorPosition <> value then
                if isCoordOutsideSize value then invalidArg "value" "CursorPosition must be within size"
                cursorPosition <- value
                onPropertyChanged "CursorPosition"
                onPropertyChanged "DisplayRelativeCursorPosition"

    member this.Selection
        with get() = selection
        and private set (value : Selection) =
            if this.Selection <> value then
                if isCoordOutsideSize value.Start || isCoordOutsideSize value.End then invalidArg "value" "Selection must be within size"
                selection <- value
                onPropertyChanged "Selection"
                onPropertyChanged "DisplayRelativeSelection"

    member this.Size
        with get() = size
        and set value =
            size <- value
            // Move cursor and selection inside new size
            this.CursorPosition <- moveCoordInsideSize this.CursorPosition
            this.Selection <-
                { Selection.Start = moveCoordInsideSize this.Selection.Start;
                    End = moveCoordInsideSize this.Selection.End }
            // Change width of all buffers
            buffer.Width <- value.X
            buffers.Values |> Seq.iter (fun buf -> buf.Width <- value.X)
            // Fire notifications
            onLinesChanged()
            onPropertyChanged "Size"

    member this.Lines : (char * Attributes)[] seq =
        // Screen buffer has line 0 at the bottom
        seq { for line in 0 .. size.Y - 1 -> buffer.Line (reverseY line + bufferScrollPosition) }

    member this.BufferScrollPosition
        with get() = bufferScrollPosition
        and set value =
            let value = value |> min (buffer.Count - size.Y) |> max 0
            if bufferScrollPosition <> value then
                bufferScrollPosition <- value
                onLinesChanged()
                onPropertyChanged "DisplayRelativeCursorPosition"
                onPropertyChanged "DisplayRelativeSelection"

    member this.DisplayRelativeCursorPosition
        with get() = { cursorPosition with Y = cursorPosition.Y + this.BufferScrollPosition }

    member this.DisplayRelativeSelection
        with get() =
            { Selection.Start =  { this.Selection.Start with Y = this.Selection.Start.Y + this.BufferScrollPosition };
                End = { this.Selection.End with Y = this.Selection.End.Y + this.BufferScrollPosition } }
        and set (value : Selection) =
            this.Selection <-
                { Selection.Start = { value.Start with Y = value.Start.Y - this.BufferScrollPosition };
                    End = { value.End with Y = value.End.Y - this.BufferScrollPosition } }

    member this.WindowTitle
        with get() = windowTitle
        and set value =
            windowTitle <- value
            onPropertyChanged "WindowTitle"
    member this.IconName
        with get() = iconName
        and set value =
            iconName <- value
            onPropertyChanged "IconName"


    member this.TypeChar (c : char) (scrollRegion : ScrollRegion option) =
        if this.CursorPosition.X = size.X then
            // At right margin
            if scrollRegion.IsSome then
                this.MoveRight 1 scrollRegion
            buffer.[this.CursorPosition.X - 1, reverseY this.CursorPosition.Y] <- (c, currentAttributes)
        else
            buffer.[this.CursorPosition.X, reverseY this.CursorPosition.Y] <- (c, currentAttributes)
            this.MoveRight 1 scrollRegion
        onLinesChanged()

    member this.Tab() =
        let succeedingManualTabStops = List.filter ((<) this.CursorPosition.X) manualTabStops
        let nextTabStop =
            if succeedingManualTabStops.IsEmpty then
                this.CursorPosition.X + defaultTabStop - (this.CursorPosition.X % defaultTabStop)
            else
                List.min succeedingManualTabStops
        if nextTabStop < size.X then
            this.CursorPosition <- { this.CursorPosition with X = nextTabStop }
        else
            this.MoveToLineEnd()
    member this.SetTabStop() =
        if List.exists ((=) this.CursorPosition.X) manualTabStops |> not then
            manualTabStops <- this.CursorPosition.X :: manualTabStops
    member this.ClearTabStop() =
        manualTabStops <- List.filter ((<>) this.CursorPosition.X) manualTabStops
    member this.ClearAllTabStops() =
        manualTabStops <- []

    member this.MoveTo coord =
        this.CursorPosition <- coord
    member this.MoveLeft (n : int) (scrollRegion : ScrollRegion option) =
        if this.CursorPosition.X - n < 0 then
            if scrollRegion.IsNone then
                this.CursorPosition <- { this.CursorPosition with X = 0 }
            else
                let scrollage = ((this.CursorPosition.X - n) / size.X) - 1
                this.MoveUp -scrollage scrollRegion
                this.CursorPosition <- { this.CursorPosition with X = (this.CursorPosition.X - n) % size.X + size.X }
        else
            this.CursorPosition <- { this.CursorPosition with X = this.CursorPosition.X - n }
    member this.MoveRight n (scrollRegion : ScrollRegion option) =
        if this.CursorPosition.X + n > size.X then
            if scrollRegion.IsNone then
                this.CursorPosition <- { this.CursorPosition with X = size.X - 1 }
            else
                let scrollage = (this.CursorPosition.X + n) / size.X
                this.MoveDown scrollage scrollRegion
                this.CursorPosition <- { this.CursorPosition with X = (this.CursorPosition.X + n) % size.X }
        else
            this.CursorPosition <- { this.CursorPosition with X = this.CursorPosition.X + n }
    member this.MoveUp n (scrollRegion : ScrollRegion option) =
        let top = scrollRegion |> Option.map (Option.map fst >> Option.defaultTo 0) |> Option.defaultTo 0
        if this.CursorPosition.Y >= top then
            this.CursorPosition <- { this.CursorPosition with Y = max (this.CursorPosition.Y - n) top }
            match scrollRegion with
            | Some scrollRegion ->
                this.ScrollDown (top - (this.CursorPosition.Y - n)) scrollRegion
            | None -> ()
        else
            this.CursorPosition <- { this.CursorPosition with Y = max (this.CursorPosition.Y - n) 0 }
    member this.MoveDown n (scrollRegion : ScrollRegion option) =
        let afterBottom = scrollRegion |> Option.map (Option.map snd >> Option.defaultTo size.Y) |> Option.defaultTo size.Y
        if this.CursorPosition.Y < afterBottom then
            this.CursorPosition <- { this.CursorPosition with Y = min (this.CursorPosition.Y + n) (afterBottom - 1) }
            match scrollRegion with
            | Some scrollRegion ->
                this.ScrollUp ((this.CursorPosition.Y + n) - (afterBottom - 1)) scrollRegion
            | None -> ()
        else
            this.CursorPosition <- { this.CursorPosition with Y = min (this.CursorPosition.Y + n) (size.Y - 1) }
    member this.MoveToLineStart() =
        this.CursorPosition <- { this.CursorPosition with X = 0 }
    member this.MoveToLineEnd() =
        this.CursorPosition <- { this.CursorPosition with X = size.X - 1 }

    member this.ScrollUp n scrollRegion =
        if n > 0 then
            let (top, afterBottom) = scrollRegion |> Option.defaultTo (0, size.Y)
            buffer.InsertLines n (reverseY (afterBottom - 1))
            if top = 0 then
                buffer.TrimToMaxLength()
                if this.BufferScrollPosition <> 0 then
                    // Have scrolled up and numer of lines below current top line has increased
                    this.BufferScrollPosition <- this.BufferScrollPosition + n
            else
                buffer.DeleteLines n (reverseY (top - 1))
            onLinesChanged()
            let newSelY y =
                if y >= afterBottom then y
                elif top = 0 then y - n
                elif y >= top then max (y - n) (top - 1)
                else y
            this.Selection <-
                { Selection.Start = { this.Selection.Start with Y = newSelY this.Selection.Start.Y };
                    End = { this.Selection.End with Y = newSelY this.Selection.End.Y } }
    member this.ScrollDown n scrollRegion =
        if n > 0 then
            let (top, afterBottom) = scrollRegion |> Option.defaultTo (0, size.Y)
            buffer.InsertLines n (reverseY (top - 1))
            buffer.DeleteLines n (reverseY (afterBottom - 1))
            onLinesChanged()
            let newSelY y =
                if y >= afterBottom || y < top then y
                else min (y + n) (afterBottom - 1)
            this.Selection <-
                { Selection.Start = { this.Selection.Start with Y = newSelY this.Selection.Start.Y };
                    End = { this.Selection.End with Y = newSelY this.Selection.End.Y } }

    member this.ClearScreen() =
        for line in 0 .. size.Y - 1 do buffer.ClearLine (reverseY line)
        onLinesChanged()
    member this.ClearToScreenStart() =
        for y in 0 .. this.CursorPosition.Y - 1 do buffer.ClearLine (reverseY y)
        this.ClearToLineStart()
    member this.ClearToScreenEnd() =
        for y in this.CursorPosition.Y + 1 .. size.Y - 1 do buffer.ClearLine (reverseY y)
        // Do this second as it will fire property changed
        this.ClearToLineEnd()
    member this.ClearLine() =
        buffer.ClearLine (reverseY this.CursorPosition.Y)
        onLinesChanged()
    member this.ClearToLineStart() =
        buffer.ClearLineToStart (this.CursorPosition.X, reverseY this.CursorPosition.Y)
        onLinesChanged()
    member this.ClearToLineEnd() =
        buffer.ClearLineToEnd (this.CursorPosition.X, reverseY this.CursorPosition.Y)
        onLinesChanged()

    member this.InsertLines n =
        buffer.InsertLines n (reverseY this.CursorPosition.Y)
        onLinesChanged()
    member this.DeleteLines n =
        buffer.DeleteLines n (reverseY this.CursorPosition.Y)
        onLinesChanged()

    member this.SetCurrentAttributesDefault() =
        currentAttributes <- Attributes.Default
    member this.SetCurrentForegroundColour colour =
        currentAttributes <- { currentAttributes with Foreground = colour }
    member this.SetCurrentBackgroundColour colour =
        currentAttributes <- { currentAttributes with Background = colour }
    member this.SetCurrentFontWeight fontWeight =
        currentAttributes <- { currentAttributes with FontWeight = fontWeight }
    member this.SetCurrentFontStyle fontStyle =
        currentAttributes <- { currentAttributes with FontStyle = currentAttributes.FontStyle ||| fontStyle }
    member this.ClearCurrentFontStyle fontStyle =
        currentAttributes <- { currentAttributes with FontStyle = currentAttributes.FontStyle &&& ~~~fontStyle }

    member this.SaveCursor (n : int) =
        savedCursorPositions.[n] <- this.CursorPosition
    member this.RestoreCursor (n : int) =
        this.CursorPosition <-
            if savedCursorPositions.ContainsKey n then savedCursorPositions.[n] else Coord.Origin
    member this.UseScreenBuffer (n : int) =
        buffers.[currentBufferIdx] <- buffer
        this.BufferScrollPosition <- 0
        currentBufferIdx <- n
        buffer <-
            if buffers.ContainsKey n then
                let buffer = buffers.[n]
                buffers.Remove n |> ignore
                buffer
            else
                newBuffer()
        onLinesChanged()

    interface INotifyPropertyChanged with
        [<CLIEvent>]
        member this.PropertyChanged = propertyChanged.Publish


type Audio =
    static member Bell() = Console.Beep()
    static member Volume
        with get() = 1.
        and set (v : float) = ()


module OutputCommandHandler =
    let ProcessOutput (display : Display) : OutputCommand -> unit =
        function
        | OutputCommand.Character (c, scrollRegion) -> display.TypeChar c scrollRegion
        | OutputCommand.Tab -> display.Tab()
        | SetTabStop -> display.SetTabStop()
        | ClearTabStop -> display.ClearTabStop()
        | ClearAllTabStops -> display.ClearAllTabStops()
        | OutputCommand.MoveTo coord -> display.MoveTo coord
        | OutputCommand.MoveToColumn x -> display.MoveTo { display.CursorPosition with X = x }
        | OutputCommand.MoveToRow y -> display.MoveTo { display.CursorPosition with Y = y }
        | OutputCommand.MoveLeft (n, scrollRegion) -> display.MoveLeft n scrollRegion
        | OutputCommand.MoveRight (n, scrollRegion) -> display.MoveRight n scrollRegion
        | OutputCommand.MoveUp (n, scrollRegion) -> display.MoveUp n scrollRegion
        | OutputCommand.MoveDown (n, scrollRegion) -> display.MoveDown n scrollRegion
        | OutputCommand.MoveToLineStart -> display.MoveToLineStart()
        | OutputCommand.MoveToLineEnd -> display.MoveToLineEnd()
        | OutputCommand.ScrollUp (n, scrollRegion) -> display.ScrollUp n scrollRegion
        | OutputCommand.ScrollDown (n, scrollRegion) -> display.ScrollDown n scrollRegion
        | OutputCommand.ClearScreen -> display.ClearScreen()
        | OutputCommand.ClearToScreenStart -> display.ClearToScreenStart()
        | OutputCommand.ClearToScreenEnd -> display.ClearToScreenEnd()
        | OutputCommand.ClearLine -> display.ClearLine()
        | OutputCommand.ClearToLineStart -> display.ClearToLineStart()
        | OutputCommand.ClearToLineEnd -> display.ClearToLineEnd()
        | InsertLines n -> display.InsertLines n
        | DeleteLines n -> display.DeleteLines n
        | OutputCommand.SetCurrentAttributesDefault -> display.SetCurrentAttributesDefault()
        | OutputCommand.SetCurrentForegroundColour colour -> Some colour |> display.SetCurrentForegroundColour
        | OutputCommand.ClearCurrentForegroundColour -> display.SetCurrentForegroundColour None
        | OutputCommand.SetCurrentBackgroundColour colour -> Some colour |> display.SetCurrentBackgroundColour
        | OutputCommand.ClearCurrentBackgroundColour -> display.SetCurrentBackgroundColour None
        | OutputCommand.SetCurrentFontWeight fontWeight -> display.SetCurrentFontWeight fontWeight
        | OutputCommand.SetCurrentFontStyle fontStyle -> display.SetCurrentFontStyle fontStyle
        | OutputCommand.ClearCurrentFontStyle fontStyle -> display.ClearCurrentFontStyle fontStyle
        | SaveCursor n -> display.SaveCursor n
        | RestoreCursor n -> display.RestoreCursor n
        | UseScreenBuffer n -> display.UseScreenBuffer n
        | ScrollUpThroughBuffer n -> display.BufferScrollPosition <- display.BufferScrollPosition + n
        | OutputCommand.Bell -> Audio.Bell()
        | OutputCommand.SetAudioVolume v -> Audio.Volume <- v
        | SetCursorStyle cursorStyle -> ()  // console.CursorStyle <- cursorStyle
        | SetWindowTitle text -> display.WindowTitle <- text
        | SetIconName text -> display.IconName <- text

module private KeyConverter =
    type MapType =
        | MAPVK_VK_TO_VSC = 0x0u
        | MAPVK_VSC_TO_VK = 0x1u
        | MAPVK_VK_TO_CHAR = 0x2u
        | MAPVK_VSC_TO_VK_EX = 0x3u

    [< DllImport("user32.dll") >]
    extern bool GetKeyboardState(byte[] lpKeyState)
    [< DllImport("user32.dll") >]
    extern uint32 MapVirtualKey(uint32 uCode, MapType uMapType)
    [< DllImport("user32.dll") >]
    extern int ToUnicode(uint32 wVirtKey, uint32 wScanCode, byte[] lpKeyState, [< Out >] [< MarshalAs(UnmanagedType.LPWStr, SizeParamIndex = 4s) >] StringBuilder pwszBuff, int cchBuff, uint32 wFlags)

    let GetCharFromKey (key : Key) : char option =
        let virtualKey : uint32 = KeyInterop.VirtualKeyFromKey key |> uint32
        let keyboardState : byte[] = Array.zeroCreate 256
        GetKeyboardState keyboardState |> ignore

        let scanCode : uint32 = MapVirtualKey(virtualKey, MapType.MAPVK_VK_TO_VSC)
        let stringBuilder : StringBuilder = StringBuilder(2)

        let result : int = ToUnicode(virtualKey, scanCode, keyboardState, stringBuilder, stringBuilder.Capacity, 0u)
        match result with
            | -1 -> None
            | 0 -> None
            | _ -> Some stringBuilder.[0]

module KeyInputHandler =

    let KeyboardModifiers() : InputModifiers =
        let mods = Keyboard.Modifiers
        InputModifiers.None
            ||| if mods.HasFlag ModifierKeys.Shift then InputModifiers.Shift else InputModifiers.None
            ||| if mods.HasFlag ModifierKeys.Control then InputModifiers.Control else InputModifiers.None
            ||| if mods.HasFlag ModifierKeys.Alt then InputModifiers.Alt else InputModifiers.None

    let ProcessInput (args : KeyEventArgs) : InputCommand option =
        args.Handled <- true
        match args.Key with
        | Key.Return -> Some InputCommand.Return
        | Key.Tab -> Some InputCommand.Tab
        | Key.Back -> Some InputCommand.Backspace
        | Key.Insert -> Some InputCommand.Insert
        | Key.Delete -> Some InputCommand.Delete
        | Key.Left -> Some InputCommand.Left
        | Key.Right -> Some InputCommand.Right
        | Key.Up -> Some InputCommand.Up
        | Key.Down -> Some InputCommand.Down
        | Key.PageUp -> Some InputCommand.PageUp
        | Key.PageDown -> Some InputCommand.PageDown
        | Key.Home -> Some InputCommand.Home
        | Key.End -> Some InputCommand.End
        | Key.Escape -> Some InputCommand.Escape
        | key when key >= Key.F1 && key <= Key.F24 ->
            InputCommand.Function (byte (key - Key.F1) + 1uy, KeyboardModifiers()) |> Some
        //| Key.Pause
        | _ ->
            match KeyConverter.GetCharFromKey args.Key with
            | Some c -> InputCommand.Character c |> Some
            | None ->
                args.Handled <- false
                None


type TextEventArgs(text : string) =
    inherit EventArgs()
    member this.Text = text

type ConsoleSession(ssh : SshClient, display : Display,
                    keyDown : IObservable<KeyEventArgs>,
                    mouseWheel : IObservable<MouseWheelEventArgs>,
                    textPasted : IObservable<TextEventArgs>) =
    static let DebugPrint (host : string) (out : bool) (data : byte[]) : unit =
        let header = String.Format((if out then "client -> {0}:" else "{0} -> client:"), host)
        let formatChar (c : byte) =
            if c >= 0x20uy && c < 0x7Fuy then
                String.Format("   {0} ('{1}')", c, (char c))
            else
                String.Format("   {0}", c)
        String.Join(Environment.NewLine, header :: (Seq.map formatChar data |> List.ofSeq)) |> Debug.WriteLine

    let keysFromConsole : IObservable<InputCommand> = Observable.choose KeyInputHandler.ProcessInput keyDown
    let pastedIntoConsole : IObservable<InputCommand> =
        Observable.map (fun (args : TextEventArgs) -> InputCommand.PastedText args.Text) textPasted
    let inputFromConsole : IObservable<InputCommand> = Observable.merge keysFromConsole pastedIntoConsole

    let displayResizes : IObservable<unit> =
        (display :> INotifyPropertyChanged).PropertyChanged
            |> Observable.filter (fun args -> args.PropertyName = "Size")
            |> Observable.map ignore

    let scrollCommands : IObservable<OutputCommand seq> =
        let scrollDisplay notches = OutputCommand.ScrollUpThroughBuffer (notches * display.Size.Y / 4) |> Seq.singleton
        mouseWheel |> Observable.map ((fun args -> args.Delta / 120) >> scrollDisplay)

    let mutable stream : ShellStream = null

    member this.Connect() =
            ssh.Connect()
            stream <- ssh.CreateShellStream("xterm", uint32 display.Size.X, uint32 display.Size.Y, 0u, 0u, 0)
            let outputFromServer : IObservable<byte[]> =
                Observable.map (fun (args : ShellDataEventArgs) -> args.Data) stream.DataReceived
            let (inputToServer : IObservable<byte[]>, outputToConsole : IObservable<OutputCommand seq>) =
                XTerm.Process (outputFromServer, inputFromConsole)
            let outputToConsole = Observable.merge outputToConsole scrollCommands
            outputToConsole.Add (Seq.iter (OutputCommandHandler.ProcessOutput display |> ignoreErrorsInFn))
            displayResizes.Add (fun () ->
                stream.SendWindowChangeRequest(uint32 display.Size.X, uint32 display.Size.Y, 0u, 0u))
            inputToServer.Add (fun (data : byte[]) -> stream.Write(data, 0, data.Length); stream.Flush())

    interface IDisposable with
        member this.Dispose() =
            if stream <> null then stream.Dispose()

type ConsoleControl() as this =
    inherit Control()
    let border = Border()
    
    let caret = Rectangle()
    let caretAnimation = DoubleAnimation(1., 0., Duration(TimeSpan.FromSeconds(1.))) :> AnimationTimeline
    [<Literal>]
    static let caretLineWidth : float = 2.

    let getFormattedText (text : string) : FormattedText =
        FormattedText(text,
            CultureInfo.CurrentUICulture, this.FlowDirection,
            Typeface(this.FontFamily, this.FontStyle, this.FontWeight, this.FontStretch),
            this.FontSize, this.Foreground)
    let getCharacterSize() : Size =
        let sampleChar = getFormattedText "A"
        Size(sampleChar.Width, sampleChar.Height)
    let getCharacterOffset (coord : Coord) : Vector =
        let charSize = getCharacterSize()
        Vector(charSize.Width * (float coord.X), charSize.Height * (float coord.Y))
    let getCharacterAt (pt : Point) =
        let charSize = getCharacterSize()
        let pt : Vector = pt - this.Origin
        let y = int (pt.Y / charSize.Height)
        if y < 0 then
            Coord.Origin
        elif y >= this.SizeInChars.Y then
            { Coord.X = this.SizeInChars.X - 1; Y = this.SizeInChars.Y - 1 }
        else
            { Coord.X = int (pt.X / charSize.Width + 0.5) |> min this.SizeInChars.X |> max 0; Y = y }

    let textPasted = Event<Handler<TextEventArgs>, TextEventArgs>()


    static let OnBorderBrushChanged (d : DependencyObject) (e : DependencyPropertyChangedEventArgs) =
        let ctrl = d :?> ConsoleControl
        ctrl.Border.BorderBrush <- e.NewValue :?> Brush
    static let OnBorderThicknessChanged (d : DependencyObject) (e : DependencyPropertyChangedEventArgs) =
        let ctrl = d :?> ConsoleControl
        ctrl.Border.BorderThickness <- e.NewValue :?> Thickness
    static let OnUseBlockCaretChanged (d : DependencyObject) (e : DependencyPropertyChangedEventArgs) =
        let ctrl = d :?> ConsoleControl
        ctrl.UpdateCaretAnimation()
    static let OnSelectionChanged (d : DependencyObject) (e : DependencyPropertyChangedEventArgs) =
        let ctrl = d :?> ConsoleControl
        // Refresh state of Copy command
        CommandManager.InvalidateRequerySuggested()
    static let OnCanCopy (sender : Object) (e : CanExecuteRoutedEventArgs) =
        if not e.Handled then
            let ctrl = sender :?> ConsoleControl
            e.CanExecute <- String.IsNullOrEmpty ctrl.SelectedText |> not
            e.Handled <- true
    static let OnCopy (sender : Object) (e : ExecutedRoutedEventArgs) =
        if not e.Handled then
            let ctrl = sender :?> ConsoleControl
            Clipboard.SetText(ctrl.SelectedText, TextDataFormat.Text)
            ctrl.Selection <- Selection.Null
            e.Handled <- true
    static let OnCanPaste (sender : Object) (e : CanExecuteRoutedEventArgs) =
        if not e.Handled then
            let text : string = Clipboard.GetText TextDataFormat.Text
            e.CanExecute <- String.IsNullOrEmpty text |> not
            e.Handled <- true
    static let OnPaste (sender : Object) (e : ExecutedRoutedEventArgs) =
        if not e.Handled then
            let text : string = Clipboard.GetText TextDataFormat.Text
            if String.IsNullOrEmpty text |> not then
                let ctrl = sender :?> ConsoleControl
                ctrl.OnTextPasted text
            e.Handled <- true
    static do
        ConsoleControl.PaddingProperty.OverrideMetadata(typeof<ConsoleControl>,
            FrameworkPropertyMetadata(Thickness(caretLineWidth),
                FrameworkPropertyMetadataOptions.AffectsMeasure ||| FrameworkPropertyMetadataOptions.AffectsArrange
                    ||| FrameworkPropertyMetadataOptions.AffectsRender))
        ConsoleControl.BackgroundProperty.OverrideMetadata(typeof<ConsoleControl>,
            FrameworkPropertyMetadata(Brushes.Black, FrameworkPropertyMetadataOptions.AffectsRender))
        ConsoleControl.ForegroundProperty.OverrideMetadata(typeof<ConsoleControl>, FrameworkPropertyMetadata(Brushes.White))
        ConsoleControl.BorderBrushProperty.OverrideMetadata(typeof<ConsoleControl>, FrameworkPropertyMetadata(OnBorderBrushChanged))
        ConsoleControl.BorderThicknessProperty.OverrideMetadata(typeof<ConsoleControl>,
            FrameworkPropertyMetadata(Thickness(),
                FrameworkPropertyMetadataOptions.AffectsMeasure ||| FrameworkPropertyMetadataOptions.AffectsArrange
                    ||| FrameworkPropertyMetadataOptions.AffectsRender,
                OnBorderThicknessChanged))
        ConsoleControl.FontFamilyProperty.OverrideMetadata(typeof<ConsoleControl>, FrameworkPropertyMetadata(FontFamily("Courier New")))
        ConsoleControl.CursorProperty.OverrideMetadata(typeof<ConsoleControl>, FrameworkPropertyMetadata(Cursors.IBeam))
        
        CommandManager.RegisterClassCommandBinding(typeof<ConsoleControl>, CommandBinding(ApplicationCommands.Copy, ExecutedRoutedEventHandler(OnCopy), CanExecuteRoutedEventHandler(OnCanCopy)))
        CommandManager.RegisterClassCommandBinding(typeof<ConsoleControl>, CommandBinding(ApplicationCommands.Paste, ExecutedRoutedEventHandler(OnPaste)))

    do
        this.Foreground <- Brushes.White    // There seems to be no non-hack way to avoid inheriting from the TextElement.Foreground property
        caretAnimation.AutoReverse <- true
        caretAnimation.RepeatBehavior <- RepeatBehavior.Forever
        caret.SetBinding(Rectangle.FillProperty, Binding("Foreground", Source = this)) |> ignore
        let menu = ContextMenu()
        MenuItem(Command = ApplicationCommands.Copy, CommandTarget = this) |> menu.Items.Add |> ignore
        MenuItem(Command = ApplicationCommands.Paste, CommandTarget = this) |> menu.Items.Add |> ignore
        this.ContextMenu <- menu

    static let cursorPositionProperty : DependencyProperty =
        DependencyProperty.Register("CursorPosition", typeof<Coord>, typeof<ConsoleControl>,
            FrameworkPropertyMetadata(Coord.Origin, FrameworkPropertyMetadataOptions.AffectsArrange))
    static let sizeInCharsProperty : DependencyProperty =
        DependencyProperty.Register("SizeInChars", typeof<Coord>, typeof<ConsoleControl>,
            FrameworkPropertyMetadata({ Coord.X = 80; Y = 24 },
                FrameworkPropertyMetadataOptions.BindsTwoWayByDefault
                    ||| FrameworkPropertyMetadataOptions.AffectsMeasure ||| FrameworkPropertyMetadataOptions.AffectsArrange ||| FrameworkPropertyMetadataOptions.AffectsRender))
    static let linesProperty : DependencyProperty =
        DependencyProperty.Register("Lines", typeof<(char * Attributes)[] seq>, typeof<ConsoleControl>,
            FrameworkPropertyMetadata([] :> (char * Attributes)[] seq, FrameworkPropertyMetadataOptions.AffectsRender))
    static let selectionProperty : DependencyProperty =
        DependencyProperty.Register("Selection", typeof<Selection>, typeof<ConsoleControl>,
            FrameworkPropertyMetadata(Selection.Null,
                FrameworkPropertyMetadataOptions.BindsTwoWayByDefault ||| FrameworkPropertyMetadataOptions.AffectsRender,
                OnSelectionChanged))
    static let selectionForegroundProperty : DependencyProperty =
        DependencyProperty.Register("SelectionForeground", typeof<Brush>, typeof<ConsoleControl>, UIPropertyMetadata(Brushes.White))
    static let selectionBackgroundProperty : DependencyProperty =
        DependencyProperty.Register("SelectionBackground", typeof<Brush>, typeof<ConsoleControl>, UIPropertyMetadata(Brushes.Navy))
    static let useBlockCaretProperty : DependencyProperty =
        DependencyProperty.Register("UseBlockCaret", typeof<bool>, typeof<ConsoleControl>, UIPropertyMetadata(false))


    member private this.Border : Border = border
    member private this.Origin = Point(this.Padding.Left + this.BorderThickness.Left, this.Padding.Top + this.BorderThickness.Top)


    member this.CursorPosition
        with get() = this.GetValue(cursorPositionProperty) :?> Coord
        and set (value : Coord) = this.SetValue(cursorPositionProperty, value)
    static member CursorPositionProperty : DependencyProperty = cursorPositionProperty

    member this.SizeInChars
        with get() = this.GetValue(sizeInCharsProperty) :?> Coord
        and set (value : Coord) = this.SetValue(sizeInCharsProperty, value)
    static member SizeInCharsProperty : DependencyProperty = sizeInCharsProperty

    member this.Lines
        with get() = this.GetValue(linesProperty) :?> (char * Attributes)[] seq
        and set (value : (char * Attributes)[] seq) = this.SetValue(linesProperty, value)
    static member LinesProperty : DependencyProperty = linesProperty

    member this.Selection
        with get() = this.GetValue(selectionProperty) :?> Selection
        and set (value : Selection) = this.SetValue(selectionProperty, value)
    static member SelectionProperty : DependencyProperty = selectionProperty

    member this.SelectionForeground
        with get() = this.GetValue(selectionForegroundProperty) :?> Brush
        and set (value : Brush) = this.SetValue(selectionForegroundProperty, value)
    static member SelectionForegroundProperty : DependencyProperty = selectionForegroundProperty

    member this.SelectionBackground
        with get() = this.GetValue(selectionBackgroundProperty) :?> Brush
        and set (value : Brush) = this.SetValue(selectionBackgroundProperty, value)
    static member SelectionBackgroundProperty : DependencyProperty = selectionBackgroundProperty

    member this.UseBlockCaret
        with get() = this.GetValue(useBlockCaretProperty) :?> bool
        and set (value : bool) = this.SetValue(useBlockCaretProperty, value)

    static member UseBlockCaretProperty : DependencyProperty = useBlockCaretProperty


    member this.SelectedText : string =
        let fy = this.Selection.First.Y |> max 0
        let ly = this.Selection.Last.Y |> min (this.SizeInChars.Y - 1)
        if ly < fy then
            null
        else
            let lines = this.Lines |> Seq.skip fy |> Seq.take (ly + 1 - fy) |> List.ofSeq
            let rec getStrings (firstPos : int) (lastPos : int) : (char * Attributes)[] list -> string list =
                let stringFromLine = Seq.map fst >> (fun cs -> String(Array.ofSeq cs).TrimEnd())
                function
                | [ l ] ->
                    if firstPos >= l.Length then
                        [ String.Empty ]
                    else
                        [ l |> Seq.skip firstPos
                            |> if lastPos < Array.length l then Seq.take (lastPos - firstPos) else id
                            |> stringFromLine ]
                | l :: ls -> (l |> Seq.skip firstPos |> stringFromLine) :: getStrings 0 lastPos ls
                | [] -> []  // Selection covered an area where lines haven't been created
            let textLines : string list = getStrings this.Selection.First.X this.Selection.Last.X lines
            String.Join(Environment.NewLine, textLines)


    override this.OnLostKeyboardFocus (args : KeyboardFocusChangedEventArgs) =
        this.UpdateCaretAnimation()
        this.RemoveVisualChild(caret)
        base.OnLostKeyboardFocus args
    override this.OnGotKeyboardFocus (args : KeyboardFocusChangedEventArgs) =
        this.AddVisualChild(caret)
        this.UpdateCaretAnimation()
        base.OnGotKeyboardFocus args
    member private this.UpdateCaretAnimation() =
        if this.HasEffectiveKeyboardFocus && not this.UseBlockCaret then
            caret.BeginAnimation(Line.OpacityProperty, caretAnimation)
        else
            caret.BeginAnimation(Line.OpacityProperty, null)


    override this.MeasureOverride (constr : Size) : Size =
        let charSize = getCharacterSize()
        if this.CursorPosition.Y >= this.SizeInChars.Y then
            // Scrolled off bottom of screen
            caret.Measure (Size(0., 0.))
        elif this.UseBlockCaret then
            caret.Measure charSize
        else
            caret.Measure (Size(caretLineWidth, charSize.Height))
        let overheadWidth = this.Padding.Left + this.BorderThickness.Left + this.Padding.Right + this.BorderThickness.Right
        let overheadHeight = this.Padding.Top + this.BorderThickness.Top + this.Padding.Bottom + this.BorderThickness.Bottom
        let sizeInChars = this.SizeInChars
        if constr.Width <> infinity && constr.Height <> infinity then
            this.SizeInChars <- {
                Coord.X = (constr.Width - overheadWidth) / charSize.Width |> int;
                Y = (constr.Height - overheadHeight) / charSize.Height |> int
            }
        elif constr.Width <> infinity then
            this.SizeInChars <- {
                this.SizeInChars with X = (constr.Width - overheadWidth) / charSize.Width |> int }
        elif constr.Height <> infinity then
            this.SizeInChars <- {
                this.SizeInChars with Y = (constr.Height - overheadHeight) / charSize.Height |> int }
        let textAreaSize : Vector = getCharacterOffset this.SizeInChars
        Size(textAreaSize.X + overheadWidth, textAreaSize.Y + overheadHeight)

    override this.ArrangeOverride (arrangeBounds : Size) : Size =
        let charPos : Point = this.Origin + getCharacterOffset(this.CursorPosition)
        if this.CursorPosition.Y >= this.SizeInChars.Y then
            // Scrolled off bottom of screen
            caret.Arrange (Rect(0., 0., 0., 0.))
        elif this.UseBlockCaret then
            caret.Arrange(Rect(charPos, getCharacterSize()))
        else
            caret.Arrange(Rect(charPos.X - caretLineWidth / 2., charPos.Y, caretLineWidth, getCharacterSize().Height))
        base.ArrangeOverride(arrangeBounds)

    override this.OnRender (drawingContext : DrawingContext) =
        // Draw background
        base.OnRender(drawingContext)
        drawingContext.DrawRectangle(this.Background |> defaultTo (Brushes.Transparent :> Brush), null, Rect(0., 0., this.ActualWidth, this.ActualHeight))
        let origin = this.Origin
        let charSize = getCharacterSize()
        this.Lines |> Seq.iteri (fun lineNo line ->
            let pos = Point(origin.X, origin.Y + float lineNo * charSize.Height)
            // Get FormattedText
            let line =
                Array.map (fun (c, a) ->
                        if a.FontStyle.HasFlag FontStyle.Hidden then (' ', a) else (c, a))
                    line
            let ft = getFormattedText (String(Array.map fst line))
            // Set font weight
            let getFontWeight (_, a : Attributes) =
                if a.FontStyle.HasFlag FontStyle.Blink then
                    FontWeights.Black
                else
                    match a.FontWeight with
                    | FontWeight.Normal -> FontWeights.Normal
                    | FontWeight.Bold -> FontWeights.Bold
                    | FontWeight.Light -> FontWeights.Light
            line |> List.getPartitionLocations getFontWeight
                |> List.filter (fun (fontWeight, start, len) -> fontWeight <> FontWeights.Normal)
                |> List.iter ft.SetFontWeight
            // Set font style
            line |> List.getPartitionLocations (fun (_, a) -> a.FontStyle.HasFlag FontStyle.Underline)
                |> List.filter (fun (isSet, start, len) -> isSet)
                |> List.iter (fun (isSet, start, len) -> ft.SetTextDecorations(TextDecorations.Underline, start, len))
            line |> List.getPartitionLocations (fun (_, a) -> a.FontStyle.HasFlag FontStyle.Italic || a.FontStyle.HasFlag FontStyle.Oblique)
                |> List.filter (fun (isSet, start, len) -> isSet)
                |> List.iter (fun (isSet, start, len) -> ft.SetFontStyle(FontStyles.Italic, start, len))
            // Set colours
            let getColourBrush (c : Colour) : Brush option =
                let (r, g, b) = c.AsRgb
                SolidColorBrush(Color.FromRgb(r, g, b)) :> Brush |> Some
            let extractOption =
                function
                | (None, _, _) -> None
                | (Some brush, start, len) -> Some (brush, start, len)
            let getForegroundBrush (_, a : Attributes) =
                if a.FontStyle.HasFlag FontStyle.Inverse then
                    match a.Background with
                    | None -> Some this.Background
                    | Some colour -> getColourBrush colour
                else
                    match a.Foreground with
                    | None -> None
                    | Some colour -> getColourBrush colour
            line |> List.getPartitionLocations getForegroundBrush
                |> List.choose extractOption
                |> List.iter ft.SetForegroundBrush
            let getBackgroundBrush (_, a : Attributes) =
                if a.FontStyle.HasFlag FontStyle.Inverse then
                    match a.Foreground with
                    | None -> Some this.Foreground
                    | Some colour -> getColourBrush colour
                else
                    match a.Background with
                    | None -> None
                    | Some colour -> getColourBrush colour
            line |> List.getPartitionLocations getBackgroundBrush
                |> List.choose extractOption
                |> List.iter (fun (brush, start, len) ->
                    drawingContext.DrawGeometry(brush, null,
                        ft.BuildHighlightGeometry(pos, start, len)))
            // Draw selection
            if not this.Selection.IsEmpty && this.Selection.First.Y <= lineNo && this.Selection.Last.Y >= lineNo then
                let from = if this.Selection.First.Y = lineNo then this.Selection.First.X else 0
                let afterLast = (if this.Selection.Last.Y = lineNo then this.Selection.Last.X else this.SizeInChars.X)
                let textLen = (min afterLast line.Length) - from
                // Draw selection background
                if this.SelectionBackground <> null then
                    drawingContext.DrawRectangle(this.SelectionBackground, Pen(this.SelectionBackground, 1.),
                        Rect(origin + getCharacterOffset { Coord.X = from; Y = lineNo },
                            origin + getCharacterOffset { Coord.X = afterLast; Y = lineNo + 1 }))
                // Change relevant text to selection Foreground brush
                if textLen > 0 && this.SelectionForeground <> null then
                    ft.SetForegroundBrush(this.SelectionForeground, from, textLen)
            // Draw text
            drawingContext.DrawText(ft, pos)
        )


    override this.VisualChildrenCount : int = if this.HasEffectiveKeyboardFocus then 2 else 1

    override this.GetVisualChild (index : int) : Visual =
        match index with
        | 0 -> border :> Visual
        | 1 when this.HasEffectiveKeyboardFocus -> caret :> Visual
        | _ -> invalidArg "index" "There are fewer controls than the index requested"


    member private this.OnTextPasted text =
        textPasted.Trigger(this, TextEventArgs text)

    [<CLIEvent>]
    member this.TextPasted = textPasted.Publish

    override this.OnKeyDown (args : KeyEventArgs) =
        let mods = Keyboard.Modifiers
        if args.Key = Key.Tab && mods.HasFlag ModifierKeys.Control then
            this.MoveFocus(
                new TraversalRequest(
                    if mods.HasFlag ModifierKeys.Shift then FocusNavigationDirection.Previous
                    else FocusNavigationDirection.Next)
                ) |> ignore
            args.Handled <- true
        base.OnKeyDown args

    override this.OnMouseDown (args : MouseButtonEventArgs) =
        if args.ChangedButton = MouseButton.Left then
            this.CaptureMouse() |> ignore
            let charClicked = args.GetPosition this |> getCharacterAt
            this.Selection <-
                if Keyboard.Modifiers.HasFlag ModifierKeys.Shift then
                    this.Selection.ExtendTo charClicked
                else
                    Selection.Empty charClicked
            // Temporary way to set focus until work out why it's not already happening
            this.Focus() |> ignore
        base.OnMouseDown args
    override this.OnMouseUp (args : MouseButtonEventArgs) =
        if args.ChangedButton = MouseButton.Left then
            this.ReleaseMouseCapture()
        base.OnMouseUp args

    override this.OnMouseMove (args : MouseEventArgs) =
        if args.LeftButton = MouseButtonState.Pressed && this.IsMouseCaptureWithin then
            this.Selection <- args.GetPosition this |> getCharacterAt |> this.Selection.ExtendTo
        base.OnMouseMove args
