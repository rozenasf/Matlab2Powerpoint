function [X_Location,Y_Location] = Matlab2Powerpoint_Position(Matlab_Obj,Ax,Powerpoint_Obj)
    Pos = Matlab_Obj.Position;
    Axes_Position = Ax.Position;

    MatlabSize = Matlab_Obj.Extent;
    X_extent = MatlabSize(1) + MatlabSize(3)/2;
    Pos(1) = X_extent;
    
    Xlim = Ax.XLim;
    X_Axes_fraction = (Pos(1) - Xlim(1)) ./ (Xlim(2) - Xlim(1));
    X_Figure_fraction = X_Axes_fraction * Axes_Position(3) + Axes_Position(1);
    X_Location = Powerpoint_Obj.Left + (Powerpoint_Obj.Width) * X_Figure_fraction;

    % Pos(2) = -1;
    

    Ylim = Ax.YLim;
    Y_extent = MatlabSize(2) + MatlabSize(4)/2;
    Pos(2) = Y_extent;
    Y_Axes_fraction = (Pos(2) - Ylim(1)) ./ (Ylim(2) - Ylim(1));
    Y_Figure_fraction = 1 - (Y_Axes_fraction * Axes_Position(4) + Axes_Position(2));
    Y_Location = Powerpoint_Obj.Top + (Powerpoint_Obj.Height) * Y_Figure_fraction;
    
    
end