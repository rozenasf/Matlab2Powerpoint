function U = AddBoxToPowerpoint(Ax,Powerpoint_Obj)
U={};
    Xlim=xlim(Ax);Ylim = ylim(Ax);
    Xlist = Xlim([1,2,2,1,1]);
    Ylist = Ylim([1,1,2,2,1]);
    X_Location=[];Y_Location=[];
    for i=1:numel(Xlist)
        [X_Location(i),Y_Location(i)] = Matlab2Powerpoint_Position(struct('Extent',[Xlist(i),Ylist(i),0,0]),Ax,Powerpoint_Obj);
    end
    for i=1:numel(Xlist)-1
        U{i} = AddLineToSlide(Powerpoint_Obj.Parent,struct('P1',[X_Location(i),Y_Location(i)],'P2',[X_Location(i+1),Y_Location(i+1)]));
    end
end