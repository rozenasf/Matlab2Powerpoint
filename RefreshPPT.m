function MatlabPPT = RefreshPPT(varargin)

    ppt         = actxserver('PowerPoint.Application');
    ppt.Visible = 1;
    pres = get(ppt,'ActivePresentation');
    Number_Of_Slides = pres.Slides.Count;
    Slide = {};
    Properties = {'name','top','left','width','height'};
    Items = {};
    k=1;
    for i=1:Number_Of_Slides
        Slide{i}=[];
        Slide{i}.obj=pres.Slides.Item(i);
        for j=1:Slide{i}.obj.Shapes.Count
            Slide{i}.Item{j}=[];
            Slide{i}.Item{j}.obj = Slide{i}.obj.Shapes.Item(j);
            Items{k,1} = Slide{i}.obj.Shapes.Item(j).name;
            Items{k,2} = Slide{i}.obj.Shapes.Item(j);
            Items{k,3} = Slide{i}.obj;
            k = k+1;
        end
    end
    MatlabPPT = Items;
end