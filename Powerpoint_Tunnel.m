classdef Powerpoint_Tunnel < handle
    properties
        MatlabPPT
        Slide
        Frame
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
        
        
        function obj = Powerpoint_Tunnel(Ax)
            obj.Path = pwd;
            obj.Fx = Ax.Parent;
            obj.Ax = Ax;
            obj.Refresh;
            obj.name = ['figure',num2str(obj.Fx.Number)];
            obj.Find_pp_Frame(obj.name);
            obj.SizeFactor = 1;
            obj.mresolution = 2;
            obj.Options = struct(...
                'Line',struct('Color',[0,0,0],'Weight',1) ...
                ,'Box',struct('Color',[0,0,0],'Weight',3) ...
                ,'Text',struct('Color',[0,0,0],'FontSize',24,'FontName','Times New Roman') ...
                ,'Ticks',struct('Color',[0,0,1],'Weight',1) ...
                );
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
        
        function Find_pp_Frame(obj,name)
            Frame_cell = Obj_From_Placeholder(name,obj.MatlabPPT);
            obj.Frame = Frame_cell{1};
            obj.Slide = obj.Frame.Parent;
        end
        
        
        function BackupFigure(obj)
            Backup_List.Ax = {'color','Position'};
            Backup_List.Fx = {};
            names = {'Ax','Fx'};
            for j=1:numel(names)
                obj.(names{j}).UserData.Backup = struct();
                for i = 1:numel(Backup_List.(names{j}))
                    obj.(names{j}).UserData.Backup.(Backup_List.(names{j}){i}) = ...
                        get(obj.(names{j}),(Backup_List.(names{j}){i}));
                end
            end
        end
        
        
        function RestoreFigure(obj)
            Types = {'Ax','Fx'};
            for j=1:numel(Types)
                Type = Types{j};
                Names = fieldnames(obj.(Type).UserData.Backup);
                for i = 1:numel(Names)
                    set(obj.(Type),Names{i},...
                        obj.(Type).UserData.Backup.(Names{i}));
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
            new_height = obj.Frame.Height*obj.SizeFactor;
            new_width = obj.Frame.Width.*obj.SizeFactor;
            obj.Fx.Position(2) = obj.Fx.Position(2)  - (new_height - obj.Fx.Position(4));
            obj.Fx.Position(1) = obj.Fx.Position(1)  - (new_width - obj.Fx.Position(3));
            obj.Fx.Position(4) = new_height;
            obj.Fx.Position(3) = new_width;
            drawnow;
            if(abs(obj.Fx.Position(3)-new_width)>1e-3 || abs(obj.Fx.Position(4)-new_height)>1e-3)
                warning('The figure scaled is too small\large. change scaling by: Fx=gcf;Fx.UserData.Scale=#');
            end
        end
        
        
        function GeneratePNG(obj)
            obj.BackupFigure;
            set(obj.Ax,'color','none');
            export_fig(obj.Fx,[obj.Path,'temp.png'],'-nocrop','-dpng',['-m',num2str(obj.mresolution)],'-transparent');
            obj.RestoreFigure;
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
        
        
        function Line = DrawLine(obj,x1,y1,x2,y2,Options)
           Line = obj.Slide.Shapes.AddLine(...
               obj.Ax2PP_x(x1),obj.Ax2PP_y(y1),...
               obj.Ax2PP_x(x2),obj.Ax2PP_y(y2));
           if(~exist('Options','var'));Options = obj.Options.Line;end
           
               Line.Line.Weight = Options.Weight;
               Line.Line.ForeColor.RGB = RGB_int(Options.Color);

           obj.Slide.invoke;
           
        end
        
        
        function Curves = DrawAllCurves(obj)
            Curves = {};
            for i=1:numel(obj.Ax.Children)
                Object = obj.Ax.Children(i);
                if(isa(Object,'matlab.graphics.chart.primitive.Line'))
                    Curves{i} = DrawCurve(obj, Object);
                    Curves{i}.Line.ForeColor.RGB = RGB_int(Object.Color);
                    Curves{i}.Line.Weight = Object.LineWidth*1.5;
                    Curves{i}.name = [obj.name, '_Curve_', num2str(i)];
                    obj.Hide(Object);
                    
                end
            end
        end
        
        function OutputCurve = DrawCurve(obj, LineObj)
            Curve=obj.Slide.Shapes.BuildFreeform('msoEditingCorner',obj.Ax2PP_x(LineObj.XData(1)),obj.Ax2PP_y(LineObj.YData(1)) );
            for i=2:numel(LineObj.XData)
                Curve.AddNodes('msoSegmentCurve','msoEditingCorner',obj.Ax2PP_x(LineObj.XData(i)),obj.Ax2PP_y(LineObj.YData(i)));
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
            TextBox.TextFrame.TextRange.Font.Color.RGB = RGB_int(Options.Color * 0); %EVERYTHING BLACK!
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
            TitleBox.name = [obj.name,'_Title'];
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
                LabesBoxs{i}.name = [obj.name,'_',Names{i}];
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
                TickValues{end+1} = obj.DrawText(Text, ...
                    Pos(1),...
                    Pos(2),...
                    struct('FontSize',Object.FontSize*1.5,'Color',Object.Color,'FontName',Object.FontName,'Rotation',0) );
                TickValues{end}.name = [obj.name,'_XTickValue_',num2str(i)];
                TickValues{end}.Top = TickValues{end}.Top + TickValues{end}.TextFrame.TextRange.Font.Size/1.5;
            end
            obj.Hide(Object);
            
            Object = obj.Ax.YAxis;
            for i=1:numel(Object.TickValues)
                Pos = [Xlim(1),Object.TickValues(i)];Text = [' ',Object.TickLabels{i},' '];
                TickValues{end+1} = obj.DrawText(Text, ...
                    Pos(1),...
                    Pos(2),...
                    struct('FontSize',Object.FontSize*1.5,'Color',Object.Color,'FontName',Object.FontName,'Rotation',0, ...
                'Alignment','ppAlignRight') );
                TickValues{end}.name = [obj.name,'_YTickValue_',num2str(i)];
                TickValues{end}.Left = TickValues{end}.Left - TickValues{end}.Width/2;
            end
            obj.Hide(Object);
        end
        
        
        function Box = DrawBox(obj)
           Xlim = xlim(obj.Ax);Ylim = ylim(obj.Ax);
           Xlist = Xlim ([1,2,2,1,1]);
           Ylist = Ylim ([1,1,2,2,1]);
           Box = {};
           for i=1:4
              Box{i} = obj.DrawLine(Xlist(i),Ylist(i),Xlist(i+1),Ylist(i+1),obj.Options.Box);
              Box{i}.name = [obj.name, '_BoxLine_', num2str(i)];
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