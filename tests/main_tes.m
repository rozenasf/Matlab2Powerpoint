figure(1);clf
Fx=gcf;Fx.UserData.Scale=1;
X=linspace(2,16,100);
plot(X,sin(X),'linewidth',3);
hold on
plot(X,cos(X),'linewidth',3);
grid on
TXYL('Hi','X [um]','Y [um]','OK','No!')
%%
system_dependent('COM_SafeArraySingleDim', 1)
Fig = Powerpoint_Tunnel(gca);

            Fig.Resize;
%             Fig.BackupFigure;
%             Fig.DrawLabes;
%             Fig.DrawTitle;
%             Fig.DrawBox;
%             Fig.DrawTickValues;
%             Fig.DrawAllCurves;
%             Fig.RestoreFigure;
             Fig.GeneratePNG;
%             Fig.UnHide;
            Fig.InjectPNG;
  
%         Fig.GroupAll;
%%
Fig = Powerpoint_Tunnel(gca);
Fig.Frame
Fig.GroupAll;
%%
system_dependent('COM_SafeArraySingleDim', 1)
figure1 = Powerpoint_Tunnel(gca);
figure1.Resize;
figure1.BackupFigure;
try figure1.DrawLabes;end
try figure1.DrawTitle;end
figure1.DrawBox;
figure1.DrawTickValues;
figure1.DrawAllCurves;
figure1.RestoreFigure;
figure1.GeneratePNG;
figure1.UnHide;
figure1.InjectPNG;
figure1.GroupAll;
% figure1.Slide.Shapes.Range({'Rectangle 87';'Rectangle 85';'Rectangle 84'}).Group
%%
Fig = Powerpoint_Tunnel(fig_obj.Children(end));

            Fig.Resize;
            Fig.BackupFigure;
            Fig.DrawLabes;
            Fig.DrawTitle;
            Fig.DrawBox;
            Fig.DrawTickValues;
            Fig.DrawAllCurves;
            Fig.RestoreFigure;
            Fig.GeneratePNG;
            Fig.UnHide;
            Fig.InjectPNG;
  
        Fig.GroupAll;
%%
%%

Fig = Powerpoint_Tunnel(gca);
Fig.GroupAll('figure')
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
    Ax=gca;
    Fx=gcf;
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

%%

Ax.Children(1)
%%
        Fig = Powerpoint_Tunnel(gca);
        if(Fig.IsGroup(Fig.Frame))
           [Objects, Names] = Fig.GetAllShapes(Fig.Frame.GroupItems);
           Fig.Frame.Ungroup;
           Fig.Refresh;
           Fig.Frame = Fig.Find_pp_Frame('figure1_Ax');
           Fig.Resize;
%            Fig.GeneratePNG;
%            Fig.InjectPNG;
Frame_Fx = Fig.Find_pp_Frame('figure1_Fx');
        Frame_Ax = Fig.Find_pp_Frame('figure1_Ax');
        
        Frame_Ax.Width
        
        Ax_Width = Fig.Ax.Position(3)*Fig.Fx.Position(3)
        
           Fig.Frame = Fig.GroupByName(Fig.Slide, Names,[]);
        else
            disp('is not group');
        end
        %%
        
        %%
        Fig.Resize;
        Fig.GeneratePNG;
        Fig.InjectPNG;