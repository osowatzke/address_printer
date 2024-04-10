word = actxserver('Word.Application');
word.Visible = 0;
document=word.Documents.Add;
selection=word.Selection;

ppi = 72;

selection.PageSetup.VerticalAlignment = 'wdAlignVerticalCenter';
selection.PageSetup.PageHeight = 5.25*ppi;
selection.PageSetup.PageWidth = 7.25*ppi;

selection.Paragraphs.Alignment = 'wdAlignParagraphCenter';

addr = load_addresses;
for i = 1:length(addr)
    printAddress(selection,addr{i});
    if i ~= length(addr)
        pageBreak(selection);
    end
end

currFilePath = fileparts(mfilename('fullpath'));
addrFileName = fullfile(currFilePath,'addresses.docx');
document.SaveAs2(addrFileName);
word.Quit;

function pageBreak(selection)
    selection.InsertNewPage;
    selection.MoveUp(4,1);
    selection.MoveUp(5,1);
    selection.MoveUntil(char(13));
    selection.Delete;
    selection.MoveDown(4,1);
end

function printAddress(selection, addr)
    fontSel = cellfun(@(x) any(regexp(x,'[0-9]')), addr.name);
    fontSel = any(fontSel);
    if fontSel
        selection.Font.Name='Aparajita';
        selection.Font.Size=20;
    else
        selection.Font.Name='Cormorant Garamond';
        selection.Font.Size=16;
    end
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