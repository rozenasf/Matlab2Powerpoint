function Update_Powerpoint_Figure(fig_obj,b)
if (strcmp(b.Key,'p') && ~isempty(b.Modifier) && sum(strcmp(b.Modifier,'alt')))
    system_dependent('COM_SafeArraySingleDim', 1)
    if(sum(strcmp(b.Modifier,'shift')))
        %         Update_All_Powerpoint_Figure();
        %         Fig = Powerpoint_Tunnel(fig_obj.Children(end));
        %         if(isempty(Fig.FrameAux))
        %             Fig.Resize;
        %             %             Backup_List.Ax = {'color','Position','YGrid','XGrid'};
        % %             Backup_List.Fx = {};
        %             Fig.BackupFigure({'color','Position'},{},1);
        %             Fig.BackupFigure({'YGrid','XGrid'},{},2);
        %             try Fig.DrawLabes;end
        %             try Fig.DrawTitle;end
        %             Fig.DrawBox;
        %             Fig.DrawTickValues;
        %             Fig.DrawAllCurves;
        %             Fig.RestoreFigure(1);
        %             Fig.GeneratePNG;
        %             Fig.RestoreFigure(2);
        %             Fig.UnHide;
        %             Fig.InjectPNG;
        %         end
        %         Fig.GroupAll(['@',Fig.name]);
  
    else
%         Update_Single_Powerpoint_Figure(fig_obj.Number);
                Options = struct(...
                'Line',struct('Color',[0,0,0],'Weight',1) ...
                ,'Box',struct('Color',[0,0,0],'Weight',3) ...
                ,'Text',struct('Color',[0,0,0],'FontSize',24,'FontName','Times New Roman') ...
                ,'Ticks',struct('Color',[0,0,0],'Weight',2.25,'Length',6) ...
                ,'Grid',struct('Color',[0.5,0.5,0.5],'Weight',0.5)...;
                );
        Fig = Powerpoint_Tunnel(fig_obj.Children(end),Options);
        Fig.Resize;
        Fig.GeneratePNG;
        Fig.InjectPNG;

    end
end
end