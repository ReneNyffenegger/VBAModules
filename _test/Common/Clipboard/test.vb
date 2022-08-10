option explicit

sub main() ' {
    dim txt as string

    txt = ClipboardToText
    msgBox "Content of clipboard is" & vbNewline & vbNewline & txt

    txt = "line one" & vbNewline & "line two" & vbNewline & "line three" & vbNewline

    textToClipboard txt
    msgBox "Clipboard should contain" & vbNewline & vbNewline & txt, vbOkOnly, "Verify clipboard"

end sub ' }
