function Obj = Obj_From_Placeholder(PlaceHolder,MatlabPPT)
    ind = find(strcmp(MatlabPPT(:,1),PlaceHolder));

    if(numel(ind)==0);Obj=[];return;end
    Obj = MatlabPPT{ind,2};
end
