function U = AddTextToSlide(Slide,Info)

%     U = Slide.Shapes.AddTextbox(1,Info.Left,Info.Top,Info.Width,Info.Height);
    U = Slide.Shapes.AddTextbox(1,Info.Left,Info.Top,Info.Width,Info.Height);
    if(isfield(Info, 'Text'));
        U.TextFrame.TextRange.Text = Info.Text;
    end
    
    if(isfield(Info, 'Name')); 
        U.TextFrame.TextRange.Font.Name = Info.Name;
    else
        U.TextFrame.TextRange.Font.Name = 'Times New Roman';
    end
    
    if(isfield(Info, 'Size'));
        U.TextFrame.TextRange.Font.Size = Info.Size;
    else
        U.TextFrame.TextRange.Font.Size = 24;
    end
    
    if(isfield(Info, 'Color'));
        U.TextFrame.TextRange.Font.Color.RGB = RGB_int(Info.Color);
    end

    if(isfield(Info, 'Spacing'));
        U.TextFrame2.TextRange.Font.Spacing=Info.Spacing;
%     else
%         U.TextFrame2.TextRange.Font.Spacing = 0.8;
    end
    U.TextFrame.TextRange.ParagraphFormat.Alignment = 'ppAlignCenter';
    Slide.invoke;
    
end