function ChangePlotColor(varargin)
Colormap=@jet;
ColorbarEnable=0;
Linewidth=[];
i=0;Ax=[];
while i<nargin
    i=i+1;
    switch(class(varargin{i}))
        case  'function_handle'
            Colormap=varargin{i};
        case 'char'
            switch varargin{i}
                case 'colorbar'
                    ColorbarEnable=1;ColorbarValues=varargin{i+1};i=i+1;
                case 'linewidth'
                    Linewidth=varargin{i+1};i=i+1;
            end
        case 'matlab.graphics.axis.Axes'
            Ax=varargin{i};
    end
end
if(isempty(Ax))
Ax=gca;
end
    Lines=Ax.Children;
    
    FullColormap=Colormap(numel(Lines));
    for j=1:numel(Lines)
        i=numel(Lines)-j+1;
        if(~isempty(Linewidth));Lines(j).LineWidth=Linewidth;end
        Lines(j).Color=FullColormap(1+mod(i-1,size(FullColormap,1)),:);
    end
%     
    if(ColorbarEnable)
        colormap(FullColormap);
        Colorbar=colorbar();
        Lin=linspace(0,1,numel(ColorbarValues)+1);Lin=(Lin(1:end-1)+Lin(2:end))/2;
        Colorbar.Ticks=Lin;
        Colorbar.TickLabels=cellfun(@(x){num2str(x)},num2cell(ColorbarValues));
    end
end