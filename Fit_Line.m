function varargout=Fit_Line(varargin)
            if(nargout==0)
                try (evalin('base','delete(Fit)'));end
            end
            obj.Ax=gca;
            Lines=obj.Ax.Children;
            %             RelevantLines=Lines(end);
            RelevantLines=Lines;
            for j=numel(RelevantLines):(-1):1
                if(strcmp(RelevantLines(j).UserData,'Fit'))
                    RelevantLines(j)=[];
                end
            end
            Type=1;
            InitialGuess=[];
            Xlim=xlim(obj.Ax);
            Ylim=ylim(obj.Ax);
            InputType='Regular';
            RequestOutput=[];
            OverideName=[];
            obj=Fit_Data.empty;
            RequestOutputValues={};
            RequestOutput={};
            for i=1:numel(varargin)
                input=varargin{i};
                switch class(input)
                    case 'double'
                        switch InputType
                            case 'Regular';
                                if(i==1)
                                    Type=input;
                                else
                                    InitialGuess=input;
                                end
                            case 'Xlim'
                                xlim(input)
                            case 'Ylim'
                                ylim(input)
                            case 'RequestOutput'
                                RequestOutputValues{numel(RequestOutput)}=input;
                            otherwise

                        end
                        InputType='Regular';
                    case 'char'
                        switch input
                            case 'all'
                                disp('FitData: It is alweys all, no need for it')
                            case 'X'
                                InputType='Xlim';
                            case 'Y'
                                InputType='Ylim';
                            otherwise
                                %if(isprop(obj,input))
                                    RequestOutput{end+1}=input;
                                    InputType='RequestOutput';
                                %end
                        end
                    case 'function_handle'
                        Type=input;
                    case 'matlab.graphics.chart.primitive.Line'
                        RelevantLines=input;
                end
            end
                obj=Fit_Data.empty;
                for i=1:numel(RelevantLines)
                    if(~isempty(RelevantLines(1).XData))
                        obj(i)=Fit_Data(Type,RelevantLines(numel(RelevantLines)-i+1),InitialGuess,gca);
                    end
                end
            xlim(Xlim);
            ylim(Ylim);
            varargout={};
            switch nargout
                case 0
                    assignin('base','Fit',obj);
                otherwise
                    if(isempty(RequestOutput))
                        varargout={obj};
                    else
                        for j=1:numel(RequestOutput)
                            if(isprop(obj(1),RequestOutput{j}))
                                varargout{j}=[obj.(RequestOutput{j})];
                            elseif(ismethod(obj(1),RequestOutput{j}))
                                varargout{j}=nan(numel(RequestOutputValues{j}),numel(obj));
                                for i=1:numel(obj)
                                    varargout{j}(:,i)=obj(i).(RequestOutput{j})(RequestOutputValues{j})';
                                end
                            else
                                varargout={obj};
                            end
                        end
                    end
            end
end