fullpath = mfilename('fullpath');
[pathstr1,~,~] = fileparts(fullpath);
addpath(genpath(pathstr1));

% Drive into toPPT folder.
cd(pathstr1);
clear pathstr1 fullpath;
%%
toPPT('setPageFormat','16:9');
%%
% figure(1);pcolor(rand(10,10));
% toPPT(figure1,'pos%',[posPercentageX,posPercentageY],'Height%',30,'posAnker','NW');
toPPT(gcf,'pos%',[25,25],'Height%',30,'posAnker','NW');
%%
saveFilename = 'test';

toPPT('saveFilename',saveFilename);
%%
ppt         = actxserver('PowerPoint.Application');
ppt.Visible = 1;
% pres        = ppt.Presentations.Add;
pres = get(ppt,'ActivePresentation');
slide3=pres.Slides.Item(3);

myFIg=figure(); %create the figure to test
  plot([0 1],[1 1],'m');
  print('-dmeta',myFIg) % copy to clipboard
  
pic = slide3.Shapes.PasteSpecial().Item(1);
  

%% Analyze Powerpoint
tic
ppt         = actxserver('PowerPoint.Application');
ppt.Visible = 1;
pres = get(ppt,'ActivePresentation');
Number_Of_Slides = pres.Slides.Count;
Slide = {};
Properties = {'name','top','left','width','height'};
for i=1:Number_Of_Slides
    Slide{i}=[];
    Slide{i}.obj=pres.Slides.Item(i);
    for j=1:Slide{i}.obj.Shapes.Count
        Slide{i}.Item{j}=[];
        Slide{i}.Item{j}.obj = Slide{i}.obj.Shapes.Item(j);
    end
end
toc
%%
title_slide = invoke( pres.Slides,'Add',1,1   ); % Add Title Slide
title       = get( title_slide.Shapes.Item(1) ); % Get Handle to title shape
subtitle    = get( title_slide.Shapes.Item(2) ); % Get Handle to subtitle shape
set(title.TextFrame.TextRange   ,'Text','Title'   );  % Set Title Text to 'Title'
set(subtitle.TextFrame.TextRange,'Text','Subtitle');  % Set Subtitle Text to 'Subitle'
%%


%%
ToReplace = 'HIHI';



myFIg=figure(1); %create the figure to test
  plot(rand(10,2));
  print('-dmeta',myFIg) % copy to clipboard



%%

tic
FigureNumber=34;
resolution='300';
Path = 'C:\Users\asafr\Desktop\jrichter24-toPPT-ce09cf0\jrichter24-toPPT-ce09cf0\';

handles=findall(0,'type','figure');
ind = find([handles.Number] == FigureNumber);
if(isempty(ind)); return ;end
figure_handle = handles(ind);

% figure_handle.Position

[MatlabPPT, Change] = RefreshPPT();
Obj = Obj_From_Placeholder(['figure',num2str(FigureNumber)],MatlabPPT);

figure_handle.Position(4) = Obj.Height./Obj.Width * figure_handle.Position(3);
print(figure_handle,[Path,'temp.jpg'],'-dpng',['-r',resolution]);
ReplaceImage(Obj,[Path,'temp.jpg']);
toc
%%
Connect_Figure_To_Powerpoint()
