for any questions, please contact me.

Add to startup, or run once:
addpath('######\matlab2powerpoint');
Connect_Figure_To_Powerpoint();

How to use:
1. In powerpoint, draw a rectangle
2. In "selection pane" change the name of the rectangle to "figure%d". for example, "figure5"
3. plot something in figure(5) in matlab.
4. while the fiugre window is active, and no tools are being used, press ALT+p keyboard combination.
this will scale figure(5) to the size of the rectange, and replace it with the figure. 
you can now move the image around, and scale it. when pressing again ALT+p, the figure will update, and all the text will look normal again.

Advance:
1. from code, Fig2Powerpoint(5) will do the same as ALT+p
2. draw figure from powerpoint elements, such as text, lines and curves. an example:
%%
% Update_All_Powerpoint_Figure();
Options = struct(...
    'Line',struct('Color',[0,0,0],'Weight',1) ...
    ,'Box',struct('Color',[0,0,0],'Weight',2) ...
    ,'Text',struct('Color',[0,0,0],'FontSize',12,'FontName','Times New Roman') ...
    ,'Ticks',struct('Color',[0,0,0],'Weight',1.5,'Length',6) ...
    ,'Grid',struct('Color',[0.5,0.5,0.5],'Weight',0.5)...;
    );

fig_obj=gcf;
Ax=gca;
Ax.XTick=Ax.XTick(((Ax.XTick>=min(xlim())) .* (Ax.XTick<=max(xlim())))==1);
Ax.YTick=Ax.YTick(((Ax.YTick>=min(ylim())) .* (Ax.YTick<=max(ylim())))==1);
Fig = Powerpoint_Tunnel(fig_obj.Children(end),Options);

Fig.Resize;
Backup_List.Ax = {'color','Position','YGrid','XGrid'};
Backup_List.Fx = {};
Fig.BackupFigure({'color','Position'},{},1);
Fig.BackupFigure({'YGrid','XGrid'},{},2);
%      try Fig.DrawLabes;end
%      try Fig.DrawTitle;end
Fig.DrawBox;
Fig.DrawTickValues;
Fig.DrawAllCurves;
Fig.DrawBoundingLines;
Fig.RestoreFigure(1);
Fig.GeneratePNG;
Fig.RestoreFigure(2);
Fig.UnHide;
%     Fig.InjectPNG;
% end


Fig.GroupAll(['@',Fig.name]);
%% regroup
Fig.GroupAll(['@',Fig.name]);


%% Load A figure. dont use figure 1
figureNumber=13;
Fx=openfig(['Q:\ilani group\Documents\Our Papers\MA graphene\figures\Fig_Files\fig',num2str(figureNumber),'.fig']);
 get(0,'Children');
 figure(figureNumber)
 copyobj(allchild(1),figureNumber);
 delete( Fx)
