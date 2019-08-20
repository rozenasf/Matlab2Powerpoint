function Update_All_Powerpoint_Figure()
    MatlabPPT = RefreshPPT();
    objects_names=MatlabPPT(:,1);
    for i=1:numel(objects_names)
        if(findstr('figure',objects_names{i}))
            Update_Single_Powerpoint_Figure(str2num(objects_names{2}(7:end)));
        end
    end
end