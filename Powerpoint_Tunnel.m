classdef Powerpoint_Tunnel < handle
    properties
        MatlabPPT
        Slide
        Frame
        FrameAux
        IsFrameGroup
        Fx
        Ax
        Frame_Size %[Left,Top,Width,Height]
        SizeFactor
        mresolution
        Path
        Options
        Hiden_Objects
        name
    end
    
    
    methods
        
        
        function obj = Powerpoint_Tunnel(Ax,Options)
            if(~exist('Options'))
                obj.Options = struct(...
                'Line',struct('Color',[0,0,0],'Weight',1) ...
                ,'Box',struct('Color',[0,0,0],'Weight',3) ...
                ,'Text',struct('Color',[0,0,0],'FontSize',24,'FontName','Times New Roman') ...
                ,'Ticks',struct('Color',[0,0,0],'Weight',2.25,'Length',6) ...
                ,'Grid',struct('Color',[0.5,0.5,0.5],'Weight',0.5)...;
                );
            else
                obj.Options = Options;
            end
            obj.Path = pwd;
            obj.Fx = Ax.Parent;
            obj.Ax = Ax;
            obj.Refresh;
            obj.name = ['figure',num2str(obj.Fx.Number)];
            obj.Frame = obj.Find_pp_Frame(obj.name);
            if(isempty(obj.Frame))
                obj.Frame = obj.Find_pp_Frame([obj.name ,'_Ax']);
            end
            obj.Slide = obj.Frame.Parent;
            obj.IsFrameGroup = obj.IsGroup(obj.Frame);
            [~,Names] = obj.GetAllShapes(obj.Slide.Shapes);
            obj.FrameAux = obj.FindPrefix(Names,['@',obj.name]);
            obj.SizeFactor = 2;
            obj.mresolution = 2;
            
            obj.Hiden_Objects={};
        end
        
        
        function Refresh(obj)
            obj.MatlabPPT = RefreshPPT();
        end
        
        
        function InjectFigure(obj)
            obj.Resize;
            obj.GeneratePNG;
            obj.InjectPNG;
        end
        
        
        function [Frame,Slide] = Find_pp_Frame(obj ,name)
            Frame_cell = Obj_From_Placeholder(name,obj.MatlabPPT);
            if(numel(Frame_cell)>0)
                Frame = Frame_cell{1};
                Slide = Frame.Parent;
            else
                Frame=[];Slide=[];
            end
        end
        
        
        function BackupFigure(obj,Backup_List_Ax,Backup_List_Fx,Index)
            %             Backup_List.Ax = {'color','Position','YGrid','XGrid'};
            %             Backup_List.Fx = {};
            Backup_List.Ax = Backup_List_Ax;
            Backup_List.Fx = Backup_List_Fx;
            names = {'Ax','Fx'};
            for j=1:numel(names)
                obj.(names{j}).UserData.Backup{Index} = struct();
                for i = 1:numel(Backup_List.(names{j}))
                    obj.(names{j}).UserData.Backup{Index}.(Backup_List.(names{j}){i}) = ...
                        get(obj.(names{j}),(Backup_List.(names{j}){i}));
                end
            end
        end
        
        
        function RestoreFigure(obj,Index)
            Types = {'Ax','Fx'};
            for j=1:numel(Types)
                Type = Types{j};
                Names = fieldnames(obj.(Type).UserData.Backup{Index});
                for i = 1:numel(Names)
                    set(obj.(Type),Names{i},...
                        obj.(Type).UserData.Backup{Index}.(Names{i}));
                end
            end
        end
        
        
        function RestorePosition(obj)
            set(obj.Ax,'Position',...
                obj.Ax.UserData.Backup.Position);
            
        end
        
        
        function Hide(obj,Object)
            Position = obj.Ax.Position;
            Object.Visible = 'off';
            obj.Ax.Position = Position;
            obj.Hiden_Objects {end+1} = Object;
        end
        
        
        function UnHide(obj)
            for i=1:numel(obj.Hiden_Objects)
                obj.Hiden_Objects{i}.Visible = 'on';
            end
            obj.Hiden_Objects = {};
        end
        
        
        function Resize(obj)
            Figure = obj.Fx;
            
            new_height = obj.Frame.Height*obj.SizeFactor;
            new_width = obj.Frame.Width.*obj.SizeFactor;
            
            if(~isempty(obj.GetSuffix(obj.Frame.name)) &&  strcmp(obj.GetSuffix(obj.Frame.name),'Ax') )
                new_width = new_width ./ obj.Ax.Position(3);
                new_height = new_height ./ obj.Ax.Position(4);
            end
            
            obj.Fx.Position(2) = Figure.Position(2)  - (new_height - Figure.Position(4));
            obj.Fx.Position(1) = Figure.Position(1)  - (new_width - Figure.Position(3));
            obj.Fx.Position(4) = new_height;
            obj.Fx.Position(3) = new_width;
            
            
            
            drawnow;
            
            % obj.Ax.OuterPosition=[0.05,0.05,0.9,0.9];
            drawnow
            if(abs(obj.Fx.Position(3)-new_width)>1e-3 || abs(obj.Fx.Position(4)-new_height)>1e-3)
                
                warning('The figure scaled is too small\large. change scaling by: Fx=gcf;Fx.UserData.Scale=#');
            end
        end
        
        
        
        
        function GeneratePNG(obj)
            Color = obj.Ax.Color;
            set(obj.Ax,'color','none');
            export_fig(obj.Fx,[obj.Path,'temp.png'],'-nocrop','-dpng',['-m',num2str(obj.mresolution)],'-transparent');
            %             obj.RestoreFigure;
            set(obj.Ax,'color',Color);
        end
        
        
        function InjectPNG(obj)
            
            ImagePath = [obj.Path,'temp.png'];
            
            Properties_To_Keep = {'name','left','top','width','height'};
            OLD = struct();
            for name = Properties_To_Keep
                OLD.(name{1}) = obj.Frame.(name{1});
            end
            OLD.ZOrderPosition = obj.Frame.ZOrderPosition;
            Image =  ...
                obj.Slide.Shapes.AddPicture(ImagePath, 'msoFalse', 'msoTrue', ...
                obj.Frame.left, obj.Frame.top,...
                obj.Frame.width*4, obj.Frame.height*4);
            
            obj.Frame.Delete;
            for name = Properties_To_Keep
                Image.(name{1}) = OLD.(name{1});
            end
            invoke(Image,'ZOrder',1);
            for i=1:(OLD.ZOrderPosition-1)
                invoke(Image,'ZOrder',2);
            end
            obj.Frame=Image;
        end
        
        function Rectangle = DrawRectangle(obj,x1,y1,x2,y2,Options)
%             X2_Moved = obj.Ax2PP_x(x2);Y2_Moved = obj.Ax2PP_y(y2);
            Rectangle = obj.Slide.Shapes.AddShape('msoshaperectangle',obj.Ax2PP_x(x1),obj.Ax2PP_y(y1),obj.Ax2PP_x(x2)-obj.Ax2PP_x(x1),obj.Ax2PP_y(y2)-obj.Ax2PP_y(y1));
            try
                Rectangle.Line.Weight=Options.Weight;
                Rectangle.Line.ForeColor.RGB = RGB_int(Options.Color);
            end
            Rectangle.Fill.ForeColor.SchemeColor = 'ppBackground';
        end
        
        
        function Line = DrawLine(obj,x1,y1,x2,y2,Options,Offset)
            X2_Moved = obj.Ax2PP_x(x2);Y2_Moved = obj.Ax2PP_y(y2);
            if(exist('Offset','var'));X2_Moved=X2_Moved+Offset(1);Y2_Moved=Y2_Moved+Offset(2);end
            Line = obj.Slide.Shapes.AddLine(...
                obj.Ax2PP_x(x1),obj.Ax2PP_y(y1),...
                X2_Moved,Y2_Moved);
            if(~exist('Options','var'));Options = obj.Options.Line;end
            
            Line.Line.Weight = Options.Weight;
            Line.Line.ForeColor.RGB = RGB_int(Options.Color);
            
            obj.Slide.invoke;
            
        end
        
        
        function Curves = DrawAllCurves(obj)
            Curves = {};
            for i=1:numel(obj.Ax.Children)
                Object = obj.Ax.Children(numel(obj.Ax.Children)-i+1);
                if(isa(Object,'matlab.graphics.chart.primitive.Line'))
                    Curves{i} = DrawCurve(obj, Object,xlim(obj.Ax),ylim(obj.Ax));
                    Curves{i}.Line.ForeColor.RGB = RGB_int(Object.Color);
                    Curves{i}.Line.Weight = Object.LineWidth;%*1.5;
                    Curves{i}.name = ['@', obj.name, '_Data_', num2str(i)];
                    obj.Hide(Object);
                    
                end
            end
        end
        
        
        function OutputCurve = DrawCurve(obj, LineObj, Xlim, Ylim)
            if(~exist('Xlim'));Xlim=[-inf,inf];end
            if(~exist('Ylim'));Ylim=[-inf,inf];end
            IND=1;
            while((LineObj.XData(IND)<Xlim(1)) || (LineObj.YData(IND)>Ylim(2)) || (LineObj.YData(IND)<Ylim(1)))
                IND=IND+1;
            end
            if(IND==1)
                Curve=obj.Slide.Shapes.BuildFreeform('msoEditingCorner',obj.Ax2PP_x(LineObj.XData(IND)),obj.Ax2PP_y(LineObj.YData(IND)) );
            else
                i=IND;
                FFF=polyfit(LineObj.XData((i-1):i),LineObj.YData((i-1):i),1);
                YCollOnX = polyval(FFF,Xlim(1));
                
                if(YCollOnX>Ylim(2))
                    FFF=polyfit(LineObj.YData((i-1):i),LineObj.XData((i-1):i),1);
                    XCollOnY1 = (polyval(FFF,Ylim(2)));
                    Curve=obj.Slide.Shapes.BuildFreeform('msoEditingCorner',obj.Ax2PP_x(XCollOnY1),obj.Ax2PP_y(Ylim(2)));
                elseif(YCollOnX<Ylim(1))
                    FFF=polyfit(LineObj.YData((i-1):i),LineObj.XData((i-1):i),1);
                    XCollOnY2 = (polyval(FFF,Ylim(1)));
                    Curve=obj.Slide.Shapes.BuildFreeform('msoEditingCorner',obj.Ax2PP_x(XCollOnY2),obj.Ax2PP_y(Ylim(1)));
                else
                    Curve=obj.Slide.Shapes.BuildFreeform('msoEditingCorner',obj.Ax2PP_x(Xlim(1)),obj.Ax2PP_y(YCollOnX));
                end
                
            end
            
            for i=(IND+1):numel(LineObj.XData)
                X=LineObj.XData(i);Y=LineObj.YData(i);
                if( X>Xlim(2) )
                    FFF=polyfit(LineObj.XData((i-1):i),LineObj.YData((i-1):i),1);
                    YCollOnX = polyval(FFF,Xlim(2));
                    Curve.AddNodes('msoSegmentCurve','msoEditingCorner',obj.Ax2PP_x(Xlim(2)),obj.Ax2PP_y(YCollOnX));
                    break;
                end
                if( Y>Ylim(2) )
                    if((i>1 && LineObj.YData(i-1)<Ylim(2)))
                        FFF=polyfit(LineObj.YData((i-1):i),LineObj.XData((i-1):i),1);
                        X = polyval(FFF,Ylim(2));
                    elseif((i<numel(LineObj.YData) && LineObj.YData(i+1)<Ylim(2)))
                        FFF=polyfit(LineObj.YData(i:(i+1)),LineObj.XData(i:(i+1)),1);
                        X = polyval(FFF,Ylim(2));
                    end
                    Y=Ylim(2);
                end
                if( Y<Ylim(1) )
                    if(i>1 && LineObj.YData(i-1)>Ylim(1))
                        FFF=polyfit(LineObj.YData((i-1):i),LineObj.XData((i-1):i),1);
                        X = polyval(FFF,Ylim(1));
                    elseif((i<numel(LineObj.YData) && LineObj.YData(i+1)>Ylim(1)))
                        FFF=polyfit(LineObj.YData(i:(i+1)),LineObj.XData(i:(i+1)),1);
                        X = polyval(FFF,Ylim(1));
                    end
                    Y=Ylim(1);
                end

                    Curve.AddNodes('msoSegmentCurve','msoEditingCorner',obj.Ax2PP_x(X),obj.Ax2PP_y(Y));
                
            end
            Curve.ConvertToShape;
            obj.Slide.invoke;
            OutputCurve = obj.Slide.Shapes.Item(obj.Slide.Shapes.Count);
            %             figure1.MatlabPPT{21,2}.Line.ForeColor.RGB=1
        end
        
        
        function TextBox = DrawText(obj, Text, x, y, Options)
            if(~exist('Options','var'));Options = obj.Options.Text;end
            TextBox = obj.Slide.Shapes.AddTextbox(1, 0, 0, 100, 100);
            
            TextBox.TextFrame.TextRange.Text = Text;
            TextBox.TextFrame.TextRange.Font.Name = Options.FontName;
            TextBox.TextFrame.TextRange.Font.Size = Options.FontSize;
            TextBox.Rotation = Options.Rotation;
            TextBox.TextFrame.TextRange.Font.Color.RGB = RGB_int(Options.Color );
            if(isfield(Options,'Alignment'))
                TextBox.TextFrame.TextRange.ParagraphFormat.Alignment = Options.Alignment;
            else
                TextBox.TextFrame.TextRange.ParagraphFormat.Alignment = 'ppAlignCenter';
            end
            TextBox.Left = obj.Ax2PP_x(x) - TextBox.Width/2;
            TextBox.Top = obj.Ax2PP_y(y) - TextBox.Height/2;
            obj.Slide.invoke;
        end
        
        
        function TitleBox = DrawTitle(obj)
            
            Object = obj.Ax.Title;
            Pos = [Object.Extent(1) + Object.Extent(3)/2,Object.Extent(2) + Object.Extent(4)/2];
            TitleBox = obj.DrawText(Object.String, ...
                Pos(1),...
                Pos(2),...
                struct('FontSize',Object.FontSize*1.5,'Color',Object.Color,'FontName',Object.FontName,'Rotation',0) );
            TitleBox.name = ['@', obj.name,'_Labels_Title'];
            obj.Hide(Object);
            
        end
        
        
        function LabesBoxs = DrawLabes(obj)
            Names = {'XAxis','YAxis'};
            for i=1:numel(Names)
                Object = obj.Ax.(Names{i}).Label;
                Pos = [Object.Extent(1) + Object.Extent(3)/2,Object.Extent(2) + Object.Extent(4)/2];
                LabesBoxs{i} = obj.DrawText(Object.String, ...
                    Pos(1),...
                    Pos(2),...
                    struct('FontSize',Object.FontSize*1.5,'Color',Object.Color,'FontName',Object.FontName,'Rotation',-Object.Rotation) );
                LabesBoxs{i}.name = ['@', obj.name,'_Labels_',Names{i}];
                obj.Hide(Object);
            end
            
        end
        
        
        function TickValues = DrawTickValues(obj)
            
            Ylim=ylim(obj.Ax);
            Xlim=xlim(obj.Ax);
            TickValues={};
            
            Object = obj.Ax.XAxis;
            for i=1:numel(Object.TickValues)
                Pos = [Object.TickValues(i),Ylim(1)];Text = [' ',Object.TickLabels{i},' '];
                
%                 TickValues{end+1} = obj.DrawText(Text, ...
%                     Pos(1),...
%                     Pos(2),...
%                     struct('FontSize',Object.FontSize*1.5,'Color',Object.Color,'FontName',Object.FontName,'Rotation',0) );
                    TickValues{end+1} = obj.DrawText(Text, ...
                    Pos(1),...
                    Pos(2),...
                    struct('FontSize',obj.Options.Text.FontSize,'Color',obj.Options.Text.Color,'FontName',obj.Options.Text.FontName,'Rotation',0) );
                
                TickValues{end}.name = ['@', obj.name,'_Axes_XTickValues_',num2str(i)];
                TickValues{end}.Top = TickValues{end}.Top + TickValues{end}.TextFrame.TextRange.Font.Size/1.5;
            end
            obj.Hide(Object);
            
            Object = obj.Ax.YAxis;
            for i=1:numel(Object.TickValues)
                Pos = [Xlim(1),Object.TickValues(i)];Text = [' ',Object.TickLabels{i},' '];
%                 TickValues{end+1} = obj.DrawText(Text, ...
%                     Pos(1),...
%                     Pos(2),...
%                     struct('FontSize',Object.FontSize*1.5,'Color',Object.Color,'FontName',Object.FontName,'Rotation',0, ...
%                     'Alignment','ppAlignRight') );
TickValues{end+1} = obj.DrawText(Text, ...
                    Pos(1),...
                    Pos(2),...
                    struct('FontSize',obj.Options.Text.FontSize,'Color',obj.Options.Text.Color,'FontName',obj.Options.Text.FontName,'Rotation',0, ...
                    'Alignment','ppAlignRight') );
                TickValues{end}.name = ['@', obj.name,'_Axes_YTickValues_',num2str(i)];
                TickValues{end}.Left = TickValues{end}.Left - TickValues{end}.Width/2;
            end
            obj.Hide(Object);
        end
        
        function DrawBoundingLines(obj)
            Xlim = xlim(obj.Ax);Ylim = ylim(obj.Ax);
            Xlist = Xlim ([1,2,2,1,1]);
            Ylist = Ylim ([1,1,2,2,1]);
            Box={};
            for i=1:4
                Box{i} = obj.DrawLine(Xlist(i),Ylist(i),Xlist(i+1),Ylist(i+1),obj.Options.Box);
                Box{i}.name = ['@', obj.name, '_Shadow_', num2str(i)];
            end
 
        end
        function [Box,Ticks,Grid] = DrawBox(obj)
            Xlim = xlim(obj.Ax);Ylim = ylim(obj.Ax);
            Xlist = Xlim ([1,2,2,1,1]);
            Ylist = Ylim ([1,1,2,2,1]);
            Box = {};Ticks={};Grid={};
            Box{5} = DrawRectangle(obj,min(Xlist),max(Ylist),max(Xlist),min(Ylist),obj.Options.Box);
            Box{5}.name = ['@', obj.name, '_Axes_Box_Rect'];
            for x = obj.Ax.XAxis.TickValues
                if(strcmp(obj.Ax.XGrid,'on'))
                    if(~(x==Xlim(1)) && ~(x==Xlim(2)))
                    Grid{end+1} = obj.DrawLine((x),(Ylim(1)),(x),(Ylim(2)),obj.Options.Grid);
                    end
                end
                Ticks{end+1} = obj.DrawLine((x),(Ylim(1)),(x),(Ylim(1)),obj.Options.Ticks,[0,-obj.Options.Ticks.Length]);
                Ticks{end+1} = obj.DrawLine((x),(Ylim(2)),(x),(Ylim(2)),obj.Options.Ticks,[0,obj.Options.Ticks.Length]);
                
            end
            for y = obj.Ax.YAxis.TickValues
                if(strcmp(obj.Ax.YGrid,'on'))
                    if(~(y==Ylim(1)) && ~(y==Ylim(2)))
                    Grid{end+1} = obj.DrawLine((Xlim(1)),(y),(Xlim(2)),(y),obj.Options.Grid);
                    end
                end
                Ticks{end+1} = obj.DrawLine((Xlim(1)),(y),(Xlim(1)),(y),obj.Options.Ticks,[obj.Options.Ticks.Length,0]);
                Ticks{end+1} = obj.DrawLine((Xlim(2)),(y),(Xlim(2)),(y),obj.Options.Ticks,[-obj.Options.Ticks.Length,0]);
                
            end
%             Xlim = xlim(obj.Ax);Ylim = ylim(obj.Ax);
%             Xlist = Xlim ([1,2,2,1,1]);
%             Ylist = Ylim ([1,1,2,2,1]);
%             for i=1:4
%                 Box{i} = obj.DrawLine(Xlist(i),Ylist(i),Xlist(i+1),Ylist(i+1),obj.Options.Box);
%                 Box{i}.name = ['@', obj.name, '_Axes_Box_Line_', num2str(i)];
%             end
            
            
            for i=1:numel(Ticks)
                Ticks{i}.name = ['@', obj.name, '_Axes_Box_Tick_', num2str(i)];
            end
            
            
            for i=1:numel(Grid)
                Grid{i}.name = ['@', obj.name, '_Axes_Box_Grid_', num2str(i)];
            end
            
            
            if(strcmp(obj.Ax.XGrid,'on'));obj.Ax.XGrid='off';end
            if(strcmp(obj.Ax.YGrid,'on'));obj.Ax.YGrid='off';end
            
        end
        
        
        function GroupAll(obj,Prefix)
            Depth = @(Names)cellfun(@(X)sum(X == '_'),Names);
            
            %Reindex Slide:
            [Objects,Names] = obj.GetAllShapes(obj.Slide.Shapes);
            Names = obj.FindPrefix(Names,Prefix);%['@',obj.name]
            
            while(1)
                [MaxDepth,ind] = max(Depth(Names));
                if(MaxDepth == 0);return;end
                NewGroup = obj.RemoveStage(Names{ind});
                NewGroupNames = obj.FindPrefix(Names, NewGroup);
                
                obj.GroupByName(obj.Slide, NewGroupNames, NewGroup)
                Names = setdiff(Names, NewGroupNames);
                Names{end+1} = NewGroup;
            end
            
            
            
            
        end
        function Cut_Name = RemoveStage(obj, Name)
            All_Split = strfind(Name,'_');
            Cut_Name = Name(1:(All_Split(end)-1));
        end
        %Tools
        
        function Suffix = GetSuffix(obj, Name)
            All_Split = strfind(Name,'_');
            if(isempty(All_Split));Suffix = '';return;end
            Suffix = Name((All_Split(end)+1):end);
        end
        
        function out = IsGroup(obj,Object)
            try
                Object.GroupItems;
                out = 1;
            catch
                out = 0;
            end
        end
        
        function OutNames = FindPrefix(~, Names, prefix)
            OutNames={};
            for j = 1:numel(Names)
                Name = Names{j};
                Find = strfind(Name,prefix);
                if ((numel(Find)==1) && Find ==1)
                    OutNames{end+1} = Name;
                end
            end
        end
        function [Objects, Names] = GetAllShapes(~, Shapes)
            Objects = {};Names={};
            for i = 1:Shapes.Count
                Objects{i} = Shapes.Item(i);
                Names{i} = Objects{i}.name;
            end
        end
        
        function NewGroup = GroupByName(obj, Parent ,Names, GroupName)
            if(isempty(GroupName));GroupName = obj.RemoveStage(Names{1});end
            if( numel(Names) > 1 )
                NewGroup = Parent.Shapes.Range(Names(:)).Group;
                NewGroup.name = GroupName;
            else
                Parent.Shapes.Range(Names(:)).name = GroupName;
            end
        end
        %
        %         function Ungroup(~, Parent, GroupName)
        %             Parent.Shapes.Range(GroupName(:)).UnGroup
        %         end
        %
        function NewGroup = GroupObjects(obj, Objects, GroupName)
            NewGroup = Objects{1}.Parent.Shapes.Range(obj.GetObjectsNames(Objects)).Group;
            NewGroup.name = GroupName;
        end
        
        
        function Names = GetObjectsNames(obj,Objects)
            Names = cell(numel(Objects),1);
            for i=1:numel(Objects)
                Names{i,1} = Objects{i}.name;
            end
        end
        
        
        
        function Delete(obj, Objects)
            if(iscell(Objects))
                for i = 1:numel(Objects)
                    obj.Delete(Objects{i});
                end
            else
                Objects.delete;
            end
        end
        
        
        function left = Ax2PP_x(obj,x)
            Xlim = xlim(obj.Ax);
            X_Axes_fraction = (x - Xlim(1)) ./ (Xlim(2) - Xlim(1));
            X_Figure_fraction = X_Axes_fraction * obj.Ax.Position(3) + obj.Ax.Position(1);
            left = obj.Frame.Left + obj.Frame.Width * X_Figure_fraction;
        end
        function top = Ax2PP_y(obj,y)
            Ylim = ylim(obj.Ax);
            Y_Axes_fraction = (y - Ylim(1)) ./ (Ylim(2) - Ylim(1));
            Y_Figure_fraction = 1 - (Y_Axes_fraction * obj.Ax.Position(4) + obj.Ax.Position(2));
            top = obj.Frame.Top + obj.Frame.Height * Y_Figure_fraction;
        end
        
    end
end