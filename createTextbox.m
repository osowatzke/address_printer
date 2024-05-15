function textbox = createTextbox(document, ppi)
    textbox = document.Shapes.AddTextbox(1, 0, 0, 7.25*ppi, 5.25*ppi);
    textbox.TextFrame.TextRange.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';
    textbox.TextFrame.VerticalAnchor = 'msoAnchorMiddle';
end