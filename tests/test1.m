figure(1);clf
Fx=gcf;Fx.UserData.Scale=1;
X=linspace(2,16,100);
plot(X,sin(X),'linewidth',3);
hold on
plot(X,cos(X),'linewidth',3);
TXYL('Hi','X [um]','Y [um]','OK','No!')
%%

try U1.Delete; clear U1;end; 
try U2.Delete; clear U2;end; 
try U3.Delete; clear U3;end; 
Ax=gca;
U1 = CopyTextToPowerpoint(Ax.XLabel,Ax,Powerpoint_Obj);
U2 = CopyTextToPowerpoint(Ax.YLabel,Ax,Powerpoint_Obj);
U3 = CopyTextToPowerpoint(Ax.Title,Ax,Powerpoint_Obj);
%  U.Top = U.Top - U.TextFrame.TextRange.BoundHeight;
