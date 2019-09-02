function U = AddLineToSlide(Slide,Info)

%     U = Slide.Shapes.AddTextbox(1,Info.Left,Info.Top,Info.Width,Info.Height);
    U = Slide.Shapes.AddLine(Info.P1(1),Info.P1(2),Info.P2(1),Info.P2(2));
    if(isfield(Info, 'Weight'));
       U.Line.Weight = Info.Weight;
    else
        U.Line.Weight = 3;
    end
    if(isfield(Info, 'Color'));
      U.Line.ForeColor.RGB = RGB_int(Color);
    else
        U.Line.ForeColor.RGB = RGB_int([0,0,0]);
    end
%     U.Line.ForeColor.RGB = RGB_int([255,0,0])
%     if(isfield(Info, 'Text'));
%         U.TextFrame.TextRange.Text = Info.Text;
%     end
%     
%     if(isfield(Info, 'Name')); 
%         U.TextFrame.TextRange.Font.Name = Info.Name;
%     else
%         U.TextFrame.TextRange.Font.Name = 'Times New Roman';
%     end
%     
%     if(isfield(Info, 'Size'));
%         U.TextFrame.TextRange.Font.Size = Info.Size;
%     else
%         U.TextFrame.TextRange.Font.Size = 24;
%     end
%     
%     if(isfield(Info, 'Color'));
%         U.TextFrame.TextRange.Font.Color.RGB = RGB_int(Info.Color);
%     end
% 
%     if(isfield(Info, 'Spacing'));
%         U.TextFrame2.TextRange.Font.Spacing=Info.Spacing;
% %     else
% %         U.TextFrame2.TextRange.Font.Spacing = 0.8;
%     end
%     U.TextFrame.TextRange.ParagraphFormat.Alignment = 'ppAlignCenter';
    Slide.invoke;
    
end