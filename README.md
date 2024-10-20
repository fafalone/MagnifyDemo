# MagnifyDemo
Screen Magnification API Demo

![image](https://github.com/user-attachments/assets/39b941dd-c311-4fee-95ef-b66e02820c9a)

This is a quick port of the Windows SDK example for the [Magnification API](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/magapi/entry-magapi-sdk).

**UPDATE:** Version 1.1 adds undocumented smoothing functionality.

>[!WARNING]
>The `useSmoothing` setting does not work in the VB6 IDE. It will crash. However it works in compiled VB6 exes and twinBASIC IDE.

There's two versions:

1) The original twinBASIC version, made using WinDevLib, so all APIs/UDTs/etc were already defined.

2) A backported VB6 version, which uses local copies of the defs. It's still 64bit compatible and can readily be imported to tB and compiled for either 32 or 64 bit.

At the top of the code you'll find options for the zoom factor (default 2.0) and whether to invert colors (default false).

>[!NOTE]
>MSDN documentation says the Magnification API is not supported under WOW64; but at least for the basic functionality in this project, it works without issue in both VB6 and tB 32bit on my 64bit Win10.
>It's recommended you use the twinBASIC version in 64bit mode for 64bit Windows just in case this isn't true in all versions or you want to expand functionality.
