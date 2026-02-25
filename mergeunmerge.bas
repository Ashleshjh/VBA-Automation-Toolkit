Sub mergeunmerge()

    ' Attempt to merge cells A1 through B5
    Range("a1:b5") = Merge
    
    ' Attempt to unmerge cells A1 through B5
    Range("a1:b5") = UnMerge

End Sub