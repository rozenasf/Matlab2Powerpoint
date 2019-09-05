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
Ax.Children(1)
%%
        Fig = Powerpoint_Tunnel(gca);
        %%
        Fig.Resize;
        Fig.GeneratePNG;
        Fig.InjectPNG;