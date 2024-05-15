function printAddress(selection, addr)
    for i = 1:length(addr.name)
        if i == length(addr.name)
            selection.Paragraphs.LineSpacing = 18;
        else
            selection.Paragraphs.LineSpacing = 12;
        end
        selection.Paragraphs.SpaceBefore = 0;
        selection.Paragraphs.SpaceAfter = 0;
        selection.TypeText(upper(addr.name{i}));
        if i ~= length(addr.name) || ~isempty(addr.addr)
            selection.TypeParagraph;
        end
    end
    selection.Font.Name='GeosansLight';
    selection.Font.Size=12;
    for i = 1:length(addr.addr)
        if i == length(addr.addr)
            selection.Paragraphs.LineSpacing = 18;
        else
            selection.Paragraphs.LineSpacing = 12;
        end
        selection.Paragraphs.SpaceBefore = 0;
        selection.Paragraphs.SpaceAfter = 0;
        selection.TypeText(upper(addr.addr{i}));
        if i ~= length(addr.addr)
            selection.TypeParagraph;
        end
    end
end