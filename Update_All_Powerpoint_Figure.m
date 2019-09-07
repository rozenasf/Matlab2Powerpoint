function Update_All_Powerpoint_Figure()
figHandles = get(0, 'Children');
for i=1:numel(figHandles)
    Update_Powerpoint_Figure(figHandles(i),struct('Key','p','Modifier','alt'))
    drawnow
end
%     MatlabPPT = RefreshPPT();
%     objects_names=MatlabPPT(:,1);
%     for i=1:numel(objects_names)
%         if(findstr('figure',objects_names{i}))
%             Update_Single_Powerpoint_Figure(str2num(objects_names{i}(7:end)));
%         end
%     end
end