
%%


%  U.Top = U.Top - U.TextFrame.TextRange.BoundHeight;

%%
return
%%
figure(1);clf
Fx=gcf;Fx.UserData.Scale=1;
X=linspace(2,16,100);
plot(X,sin(X),'linewidth',3);
hold on
plot(X,cos(X),'linewidth',3);
TXYL('Hi','X [um]','Y [um]','OK','No!')
%%
Ax=gca;
Pos1 = Ax.Position;
Ax.YAxis.Label.Visible = 'off';
Ax.XAxis.Label.Visible = 'off';
Ax.XAxis.Visible = 'off';
Ax.YAxis.Visible = 'off';
Ax.Title.Visible = 'off';
Ax.Box = 'off';
XTicksLabels = Ax.XAxis.TickLabels;
Ax.XAxis.TickLabels = [];
YTicksLabels = Ax.YAxis.TickLabels;
Ax.YAxis.TickLabels = [];
Ax.Position = Pos1;
Update_Single_Powerpoint_Figure(1);
Ax.YAxis.Label.Visible = 'on';
Ax.XAxis.Label.Visible = 'on';
Ax.Title.Visible = 'on';
Ax.Box = 'on';
Ax.XAxis.Visible = 'on';
Ax.YAxis.Visible = 'on';
Ax.XAxis.TickLabels = XTicksLabels;
Ax.YAxis.TickLabels = YTicksLabels;

%

[MatlabPPT] = RefreshPPT();
Powerpoint_Obj = MatlabPPT{1,2};
Ticks=GetTickLabelsObjects(gca);

Ax=gca;
CopyTextToPowerpoint(Ax.XLabel,Ax,Powerpoint_Obj);
CopyTextToPowerpoint(Ax.YLabel,Ax,Powerpoint_Obj);
CopyTextToPowerpoint(Ax.Title,Ax,Powerpoint_Obj);

for i=1:numel(Ticks)
   CopyTextToPowerpoint(Ticks{i},Ax,Powerpoint_Obj); 
end


AddBoxToPowerpoint(Ax,Powerpoint_Obj);

for i=1:numel(Ticks)
    delete(Ticks{i});
end
%%
 