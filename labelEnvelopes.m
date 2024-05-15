word = actxserver('Word.Application');
word.Visible = 0;
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
    setFont(selection, addr{i});
    printReturnAddress(selection);
    textbox = createTextbox(document, ppi);
    textbox.Select();
    setFont(selection, addr{i});
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