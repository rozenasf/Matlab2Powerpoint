function Connect_Figure_To_Powerpoint()
handles=findall(0,'type','figure');
for i=1:numel(handles)
    set(handles(i),'WindowKeyPressFcn',@(a,b)Update_Powerpoint_Figure(handles(i).Number,b))
end
end