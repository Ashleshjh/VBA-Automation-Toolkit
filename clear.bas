Option Explicit

Sub clear()
    ' Purpose: Demonstrates different methods to clear cell data and properties.
    ' Author: Ashlesh JH

    ' 1. Removes only the styling (Bold, Color, Fonts), keeping the text/values
    Range("a1:a10").ClearFormats

    ' 2. Removes only the comments/notes attached to the cells
    Range("a1:a10").ClearComments

    ' 3. Removes clickable hyperlinks (keeps the display text)
    Range("a1:a10").ClearHyperlinks

    ' 4. Clear EVERYTHING (Values, Formats, Comments, and Hyperlinks)
    ' This resets the cells to a completely blank state.
    Range("a1:a10").clear
End Sub
