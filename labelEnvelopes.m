word = actxserver('Word.Application');
word.Visible = 1;
document = word.Documents.Add;
page = document.ActiveWindow.ActivePane.Pages.Item(1);
width = page.Width;
height = page.Height;
selection = word.Selection;
margins = 0.25*ones(1,4)*ppi;
setMargins(selection,margins);

ppi = 72;

selection.PageSetup.PageHeight = 5.25*ppi;
selection.PageSetup.PageWidth = 7.25*ppi;

addr = loadAddresses();

for i = 1:length(addr)
    if i ~= 1
        selection.InsertNewPage();
    end
    printReturnAddress(selection);
    textbox = createTextbox(document, ppi);
    textbox.Select();
    printAddress(selection, addr{i});
    word.Selection.GoTo(3,-1);
    selection.MoveUntil(char(13));
end

currFilePath = fileparts(mfilename('fullpath'));
addrDocxFileName = fullfile(currFilePath,'Addresses.docx');
addrPdfFileName = fullfile(currFilePath,'Addresses.pdf');
document.SaveAs2(addrDocxFileName);
document.SaveAs2(addrPdfFileName,17);
word.Quit();

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