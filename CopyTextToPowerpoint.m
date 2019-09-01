function U = CopyTextToPowerpoint(Matlab_Obj,Ax,Powerpoint_Obj)
    
    [X_Location,Y_Location] = Matlab2Powerpoint_Position(Matlab_Obj,Ax,Powerpoint_Obj);

    U = AddTextToSlide(Powerpoint_Obj.Parent,struct(...
        'Text',Matlab_Obj.String,...
        'Top',Y_Location,'Left',X_Location,...
        'Height',100,'Width',200,...
        'Size',Ax.XAxis.FontSize*1.5,'Name',Ax.XAxis.FontName));
    %msoTextOrientationHorizontal
    % U.Width = U.TextFrame.TextRange.BoundWidth+10;
    U.Left = U.Left - U.Width/2;
    U.Top = U.Top - U.Height/2;
%     U.Top = U.Top - U.TextFrame.TextRange.BoundWidth/2;
     U.Rotation = - Matlab_Obj.Rotation;
    
    
end