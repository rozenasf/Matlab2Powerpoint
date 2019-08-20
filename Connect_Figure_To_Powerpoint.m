function Connect_Figure_To_Powerpoint()
set(groot,'defaultfigureWindowKeyPressFcn',@(a,b)Update_Powerpoint_Figure(a.Number,b))

% handles=findall(0,'type','figure');
% for i=1:numel(handles)
%     set(handles(i),'WindowKeyPressFcn',@(a,b)Update_Powerpoint_Figure(a.Number,b))
% end
end