function TXYL(Title,Xlabel,Ylabel,varargin)
if(exist('Title') && ~isempty(Title))
    if((~iscell(Title) && Title(1)=='$') || (iscell(Title) && Title{1}(1)=='$'))
        title(Title,'Interpreter','latex');
    else
        title(Title);
    end
end
if(exist('Xlabel') && ~isempty(Xlabel))
    if(Xlabel(1)=='$')
        xlabel(Xlabel,'Interpreter','latex');
    else
        xlabel(Xlabel);
    end
end
if(exist('Ylabel') && ~isempty(Ylabel))
    if(Ylabel(1)=='$')
        ylabel(Ylabel,'Interpreter','latex');
    else
        ylabel(Ylabel);
    end
end
if(numel(varargin)>0)
    if(iscell(varargin{2}))
        Legend={};
        for i=1:numel(varargin{2}{1})
            Legend{i}=sprintf(varargin{1},varargin{2}{1}(i));
        end
        Location='northeast';
        for k=3:numel(varargin)
           if(strcmp( varargin{k},'location'));
               Location=varargin{k+1};
               k=k-1;
               break;
           end
        end
        legend(Legend,'Interpreter','latex','location',Location);
    else
        Location='northeast';
        for k=1:numel(varargin)
           if(strcmp( varargin{k},'location'));
               Location=varargin{k+1};
               k=k-1;
               break;
           end
        end
        legend(varargin(1:k),'location',Location);%'Interpreter','latex',
    end
end
Ax=gca;Ax.FontName='Times New Roman';Ax.FontSize=16;
Fx=gcf;Fx.Color='w';
Ax=gca;Ax.LineWidth=2;
Ax.YColor=[0,0,0];
Ax.XColor=[0,0,0];
end
