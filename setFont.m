function setFont(selection, addr)
    fontSel = cellfun(@(x) any(regexp(x,'[0-9]')), addr.name);
    fontSel = any(fontSel);
    if fontSel
        selection.Font.Name='Aparajita';
        selection.Font.Size=20;
    else
        selection.Font.Name='Cormorant Garamond';
        selection.Font.Size=16;
    end
end