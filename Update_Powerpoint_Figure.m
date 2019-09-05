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
        %         Update_Single_Powerpoint_Figure(fig_number);
        %         Fig = Powerpoint_Tunnel(fig_obj.Children(end));
        %         Fig.Resize;
        %         Fig.GeneratePNG;
        %         Fig.InjectPNG;
        Fx = fig_obj;
        Ax = fig_obj.Children(end);
        Fig = Powerpoint_Tunnel(Ax);
        if(strcmp(Fig.GetSuffix(Fig.Frame.name),'Ax'))
            BaseName = Fig.RemoveStage(Fig.Frame.name);
            Fx_Frame = Fig.Find_pp_Frame([BaseName,'_Fx']);
            if(isempty(Fx_Frame))
                Fx_temp = Fig.Slide.Shapes.AddShape(1,0,0,100,100) ;
                Fx_temp.name = [BaseName,'_Fx'];
            end
        end
        Fig.GroupAll('figure')
        Fig = Powerpoint_Tunnel(fig_obj.Children(end));
        
        % if(strcmp(Fig.GetSuffix(Fig.Frame.name),'Ax'))
        if(Fig.IsGroup(Fig.Frame))
            BaseName = Fig.Frame.name;
            [Objects, Names] = Fig.GetAllShapes(Fig.Frame.GroupItems);
            Fig.Frame.Ungroup;
            Fig.Refresh;
            
            
            Ax_Frame = Fig.Find_pp_Frame([BaseName,'_Ax']);
            Fx_Frame = Fig.Find_pp_Frame([BaseName,'_Fx']);
            %     if(isempty(Fx_Frame))
            %         Fx_Frame = Fig.Slide.Shapes.AddShape(1,0,0,100,100);
            %         Fx_Frame.name = [BaseName,'_Fx'];
            %     end
%             Ax=gca;
%             Fx=gcf;
            for i=1:2
                Ax.Position(3) = Ax_Frame.Width / Fx.Position(3);
                Ax.Position(4) = Ax_Frame.Height / Fx.Position(4);
                
                Ax.OuterPosition(2) = 0.05;
                Ax.OuterPosition(1) = 0.05;
                
                Scale_X = Ax.OuterPosition(3);
                Fx.Position(3) = Fx.Position(3) * Scale_X * 1.1;
                Scale_Y = Ax.OuterPosition(4);
                Fx.Position(4) = Fx.Position(4) * Scale_Y * 1.1;
            end
            
            Fx_Frame.Width = Ax_Frame.Width / Ax.Position(3);
            Fx_Frame.Height = Ax_Frame.Height / Ax.Position(4);
            
            Fx_Frame.Left = Ax_Frame.Left - Fx_Frame.Width * Ax.Position(1);
            Fx_Frame.Top = Ax_Frame.Top - Fx_Frame.Height * (1-(Ax.Position(2)+Ax.Position(4)));
            
            Fig.GeneratePNG;
            Fig.Frame = Fx_Frame;
            Fig.InjectPNG;
            
            
            Fig.Frame = Fig.GroupByName(Fig.Slide, Names,[]);
        else
            Fig.Resize;
            Fig.GeneratePNG;
            Fig.InjectPNG;
        end
        
    end
end
end