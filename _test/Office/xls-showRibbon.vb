option explicit

sub main() ' {

    createButton range(cells(2,2), cells(3,4)), "Show Ribbon", "show_ribbon"
    createButton range(cells(5,2), cells(6,4)), "Hide Ribbon", "hide_ribbon"

end sub ' }

sub show_ribbon() ' {
    showRibbon true
end sub ' }

sub hide_ribbon() ' {
    showRibbon false
end sub ' }
