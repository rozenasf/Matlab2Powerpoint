classdef Fit_Data < handle
    properties
        line
        Markers
        Type
        Text
        TextObj
        fit
        Func
        Range
        X
        Y
        YFit
        X_0
        Y_0
        dfit
        dX_0
        dY_0
        ddfit
        ddX_0
        ddY_0
        Ax
        Color
        a
        b
        c
        d
        e
        f
        g
        Xlist
        Ylist
        DontCalculateExtremum
        Std
        AVR
    end
    methods
        function obj=Fit_Data(Type,RelevantLine,InitialGuess,Ax)
            obj.Ax=Ax;
            obj.TextObj={};
            obj.Type=Type;
            IDX=find((RelevantLine.XData>=obj.Ax.XLim(1)) .* (RelevantLine.XData<=obj.Ax.XLim(2)) .* ...
                (RelevantLine.YData>=obj.Ax.YLim(1)) .* (RelevantLine.YData<=obj.Ax.YLim(2)));
            obj.X=RelevantLine.XData(IDX);
            obj.Y=RelevantLine.YData(IDX);
            obj.AVR=mean(obj.Y);
            BlackList=find(isnan(obj.X) + isnan(obj.Y));
            obj.X(BlackList)=[];
            obj.Y(BlackList)=[];
            obj.DontCalculateExtremum=0;
            if(numel(obj.X)==0);return;end
            switch(class(obj.Type))
                case 'double'
                    obj=Polynom_Fit(obj,obj.Type);
                case 'function_handle'
                    switch nargin(obj.Type)
                        case 1
                            obj=BasicFunction_Fit(obj,obj.Type);
                        case 2
                            obj=GeneralFunction_Fit(obj,obj.Type,InitialGuess);
                            obj.DontCalculateExtremum=1;
                        otherwise
                            disp('for general fit, use F(x,a) where a is vector');
                    end
                    
            end
            
            XFit=[min(obj.X),max(obj.X)];
            obj.YFit=obj.Func(obj.X);
            obj.line=line(obj.Ax,1,1,'color','m','UserData','Fit');
            obj.Markers=line(obj.Ax,1,1,...
                'Color','red','LineStyle','none','Marker','.','UserData','Fit');
            set(obj.line,'ButtonDownFcn',@(a,b)obj.OnPress(1,a,b));
            set(obj.Markers,'ButtonDownFcn',@(a,b)obj.OnPress(2,a,b));
            
            obj.Range=XFit;
            
            obj.Text=[obj.Text,'a=[',sprintf('%.2e,' , obj.fit)];obj.Text=obj.Text(1:end-1);obj.Text=[obj.Text,']'];
            obj.Text=[obj.Text,' \n'];
            
            %obj.Show;
            %fprintf(obj.Text);
            obj.Range=obj.Ax.XLim;
            obj.Std=std(obj.YFit-obj.Y);
            
        end
        %function for users
        function out=Eval(obj,Values)
            out=obj.Func(Values);
            obj.Extend(Values);
        end
        function Extend(obj,Values)
           for i=1:numel(Values)
              if(max(obj.Range)<Values(i))
                  obj.Range(2)=Values(i);
              elseif(min(obj.Range)>Values(i))
                  obj.Range(1)=Values(i);
              end
           end
        end
        function out=FindAllZeros(obj,F,Lim)
            try
            out=[];
            Split=100;
            Ranges=linspace(Lim(1),Lim(2),Split);
            for i=2:(Split-2)
                if( (sign(F(Ranges(i)))>0) ~= (sign(F(Ranges(i+1)))>0) )
                    out=[out;fzero(F,Ranges(i:i+1))];
                end
            end
            end
        end
        function obj=ZeroAndExtremum(obj)
            syms x;
            
            dF=matlabFunction(diff(obj.Func,x));
            ddF=matlabFunction(diff(dF,x));
            
            obj.X_0=obj.FindAllZeros(obj.Func,obj.Ax.XLim);
            obj.dX_0=obj.FindAllZeros(dF,obj.Ax.XLim);
            obj.ddX_0=obj.FindAllZeros(ddF,obj.Ax.XLim);
            
            obj.Y_0=obj.Func(0);
            obj.dY_0=obj.Func(obj.dX_0);
            obj.ddY_0=obj.Func(obj.ddX_0);
        end
        
        function obj=Polynom_Fit(obj,Type)
            obj.fit=polyfit(obj.X,obj.Y,Type);
            obj.fit=obj.fit';
            
            obj.X_0=obj.KeepReal(roots(obj.fit));
            
            %obj.X_0=obj.X_0(find((obj.X_0<obj.Ax.XLim(2)) .* (obj.X_0>obj.Ax.XLim(1))));
            obj.Y_0=obj.fit(end);
            
            obj.Func=@(x)polyval(obj.fit,x);
            if(Type>1)
                obj.dfit=polyder(obj.fit);
                obj.dX_0=obj.KeepReal(roots(obj.dfit));
                %                 obj.dX_0(isreal(obj.dX_0));
                %obj.dX_0=obj.dX_0(find((obj.dX_0<obj.Ax.XLim(2)) .* (obj.dX_0>obj.Ax.XLim(1))));
                obj.dY_0=obj.Func(obj.dX_0);
            end
            if(Type>2)
                obj.ddfit=polyder(obj.dfit);
                obj.ddX_0=obj.KeepReal(roots(obj.ddfit));
                obj.ddX_0(isreal(obj.ddX_0));
                %obj.ddX_0=obj.ddX_0(find((obj.ddX_0<obj.Ax.XLim(2)) .* (obj.ddX_0>obj.Ax.XLim(1))));
                obj.ddY_0=obj.Func(obj.ddX_0);
            end
            
            obj.Text=[];
        end
        function out=KeepReal(obj,in)
            out=in(imag(in)==0);
        end
        function obj=GeneralFunction_Fit(obj,Type,InitialGuess)
            
            ToSearch=@(a)sum((Type(obj.X,a)-obj.Y).^2);
            options = optimset('TolFun',1e-6,'TolX',1e-6);
            out=fminsearch(ToSearch,InitialGuess,options);
            obj.fit=out';
            obj.Func=@(x)Type(x,obj.fit);
            obj.Text=func2str(obj.Type);
            
        end
        function obj=BasicFunction_Fit(obj,Type)
            
            %XN=(X-a)/b;
            %Y=a(3).*F(X.*a(1)+a(2))+a(4);
            
            [XN,Xa,Xb]=obj.NormelizeAxis(obj.X);
            [YN,Ya,Yb]=obj.NormelizeAxis(obj.Y);
            ToSearch=@(a)sum((obj.FunScale(Type,XN,a)-YN).^2);
            options = optimset('TolFun',1e-6,'TolX',1e-6);
            out=fminsearch(ToSearch,[1,0,1,0],options);
            
            out(3)=out(3)*Yb;
            out(4)=out(4)*Yb+Ya;
            
            out(1)=out(1)/Xb;
            out(2)=out(2)-Xa*out(1);
            
            obj.fit=out';
            obj.Func=@(x)obj.FunScale(Type,x,obj.fit);
            
            obj.Text=func2str(obj.Type);
            obj.Text=[obj.Text,'\n a(3).*F(X.*a(1)+a(2))+a(4) \n'];
        end
        function OnPress(obj,Which,a,b)
            fprintf(obj.Text);
            assignin('base','ans',obj);
            switch (b.Button)
                case 1
                    ah = get( obj.line, 'Parent' );
                    p = get( ah, 'CurrentPoint' );
                    
                    Xlist=obj.Xlist;
                    Ylist=obj.Ylist;
                    switch Which
                        case 1
                            x=p(1,1);y=obj.Func(p(1,1));
                            obj.TextObj{end+1}.text=text(p(1,1),obj.Func(p(1,1)),[sprintf(obj.Text),...
                                sprintf('(%0.2e,%0.2e)',x,y)]);
                            k=numel(obj.TextObj);
                            set(obj.TextObj{k}.text,'ButtonDownFcn',@(~,~)obj.DeleteAndClear(k));
                            
                            obj.TextObj{k}.Point=line(obj.Ax,x,y,'Color','r','Marker','o');
                        case 2
                            [~,ii]=min(abs(Xlist-p(1)));
                            x=Xlist(ii);y=Ylist(ii);
                            obj.TextObj{end+1}.text=text(Xlist(ii),Ylist(ii),...
                                sprintf('(%0.2e,%0.2e)',x,y));
                            k=numel(obj.TextObj);
                            set(obj.TextObj{end}.text,'ButtonDownFcn',@(~,~)obj.DeleteAndClear(k));
                            
                            
                    end
                    assignin('base','x',x);
                    assignin('base','y',y);
                    fprintf('(x,y)=(%e,%e)\n',x,y);
                case 3
                    obj.Range=obj.Ax.XLim;
            end
        end
        
        function DeleteAndClear(obj,Index)
            delete(obj.TextObj{Index}.text);
            if(isfield(obj.TextObj{Index},'Point'));
                delete(obj.TextObj{Index}.Point);
            end
            obj.TextObj{Index}=[];
        end
        function ExtendToX_0(obj)
            Total=[obj.Range,obj.X_0'];
            obj.Range=[min(Total),max(Total)];
        end
        function set.Range(obj,Value)
            obj.Range=Value;
            switch(class(obj.Type))
                case 'function_handle'
                    if(~obj.DontCalculateExtremum)
                        obj=ZeroAndExtremum(obj);
                    end
            end
            N=abs(round(numel(obj.X)*diff(obj.Range)/(max(obj.X)-min(obj.X))));
            if(N>1e5);error('cant extend! too many points!');end
            XR=linspace(obj.Range(1),obj.Range(2),N);
            obj.line.XData=XR;
            obj.line.YData=real(obj.Func(XR));

            obj.Xlist=[0,obj.X_0',obj.dX_0',obj.ddX_0'];
            obj.Ylist=[obj.Y_0,obj.Func(obj.X_0'),obj.dY_0',obj.ddY_0'];
            IDX=find((obj.Xlist<obj.Range(2)) .* (obj.Xlist>obj.Range(1)));
            obj.Markers.XData=obj.Xlist(IDX);
            obj.Markers.YData=obj.Ylist(IDX);
            obj=obj.AddToABC();
        end
        function [XN,a,b]=NormelizeAxis(obj,X)
            a=mean(X);
            b=std(X);
            XN=(X-a)/b;
        end
        function Y=FunScale(obj,F,X,a)
            Y=a(3).*F(X.*a(1)+a(2))+a(4);
        end
        function obj=AddToABC(obj)
            ABC='abcdef';
            for i=1:min([numel(obj.fit),6])
                obj.(ABC(i))=obj.fit(i);
            end
        end
        function delete(obj)
            try delete(obj.line); end
            try delete(obj.Markers); end
            for i=1:numel(obj.TextObj)
                try delete(obj.TextObj{i});end
            end
            
        end
    end
end