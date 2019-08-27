Ax=gca;
Polygon=line(1,1,'color','m');
PolygonCoordinates=[];
Done=0;
while(~Done)
    [x,y]=ginput(1);
if(isempty(PolygonCoordinates) || ~isequal([x,y],PolygonCoordinates(end,:)))
    PolygonCoordinates=[PolygonCoordinates;[x,y]];

else
    PolygonCoordinates=[PolygonCoordinates;PolygonCoordinates(1,:)];
   Done=1; 
end
        Polygon.XData=PolygonCoordinates(:,1);
    Polygon.YData=PolygonCoordinates(:,2);
end
PolygonCoordinates