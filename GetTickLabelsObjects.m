function Ticks=GetTickLabelsObjects(Ax)

Ylim=ylim(Ax);
Xlim=xlim(Ax);

Ticks={};
for i=1:numel(Ax.XAxis.TickValues)
   Ticks{end+1} = text(Ax.XAxis.TickValues(i),Ylim(1),sprintf([Ax.XAxis.TickLabels{i}]));
   Ticks{end}.FontSize = Ax.XAxis.FontSize;
   Ticks{end}.FontName = Ax.XAxis.FontName;
   Ticks{end}.HorizontalAlignment='center';
   Ticks{end}.Position(2) = Ticks{end}.Position(2) - Ticks{i}.Extent(4)/(4/3);
end
for i=1:numel(Ax.YAxis.TickValues)
   Ticks{end+1} = text(Xlim(1),Ax.YAxis.TickValues(i),[' ',Ax.YAxis.TickLabels{i},' ']);
   Ticks{end}.FontSize = Ax.YAxis.FontSize;
   Ticks{end}.FontName = Ax.YAxis.FontName;
   Ticks{end}.HorizontalAlignment='right';
   Ticks{end}.Position(1) = Ticks{end}.Position(1) - Ticks{i}.Extent(3)/1*0;
   
end

end