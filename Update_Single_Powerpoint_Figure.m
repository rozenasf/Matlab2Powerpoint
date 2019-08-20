function Update_Single_Powerpoint_Figure(fig_number)

FigureNumber=fig_number;
%         resolution='300';
        mresolution='2';
        Path = [pwd,'\'];
        
        handles=findall(0,'type','figure');
        ind = find([handles.Number] == FigureNumber);
        if(isempty(ind)); return ;end
        figure_handle = handles(ind);
        
        % figure_handle.Position
        
        [MatlabPPT] = RefreshPPT();
        Obj = Obj_From_Placeholder(['figure',num2str(FigureNumber)],MatlabPPT);
        if(isempty(Obj));return;end
        new_height = Obj.Height./Obj.Width * figure_handle.Position(3);
        figure_handle.Position(2) = figure_handle.Position(2)  - (new_height - figure_handle.Position(4));
        figure_handle.Position(4) = new_height;
        Ax = figure_handle.CurrentAxes;
        Backup_color = get(Ax,'color');
        set(Ax,'color','none');
% export_fig(figure_handle,[Path,'temp.png'],'-dpng',['-r',resolution],'-transparent');
export_fig(figure_handle,[Path,'temp.png'],'-dpng',['-m',mresolution],'-transparent');
set(Ax,'color',Backup_color);
%         print(figure_handle,[Path,'temp.jpg'],'-dpng',['-r',resolution]);
        ReplaceImage(Obj,[Path,'temp.png']);
end