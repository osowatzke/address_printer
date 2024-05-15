function setMargins(selection, margins)
    selection.PageSetup.LeftMargin = margins(1);
    selection.PageSetup.RightMargin = margins(2);
    selection.PageSetup.TopMargin = margins(3);
    selection.PageSetup.BottomMargin = margins(4);
end