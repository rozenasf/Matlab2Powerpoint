function ReplaceImage(Obj,Image)

Properties_To_Keep = {'name','left','top','width','height'};
    OLD = struct();
    for name = Properties_To_Keep
        OLD.(name{1}) = Obj.(name{1});
    end
    OLD.ZOrderPosition = Obj.ZOrderPosition;
%     MatlabPPT{ind,3}.Shapes
    Image =  Obj.Parent.Shapes.AddPicture(Image, 'msoFalse', 'msoTrue', Obj.left, Obj.top, Obj.width*4, Obj.height*4);
%     Image =  Obj.Parent.Shapes.AddPicture(Image, 'msoFalse', 'msoTrue', Obj.left, Obj.top, Obj.width, Obj.height);
Obj.Delete;
    for name = Properties_To_Keep
        Image.(name{1}) = OLD.(name{1});
    end
    invoke(Image,'ZOrder',1);
    for i=1:(OLD.ZOrderPosition-1)
        invoke(Image,'ZOrder',2);
    end
end