function Colorbar_gen = ExtractColormap()
    image = imclipboard('paste');
    if(size(image,2) > size(image,1))
       image = permute(image,[2,1,3]);
    end
    Raw_Colorbar = squeeze(median(image,2));
    Colorbar_gen = @(n)interp1(linspace(0,1,size(Raw_Colorbar,1)),double(Raw_Colorbar)/255,linspace(0,1,n));
end