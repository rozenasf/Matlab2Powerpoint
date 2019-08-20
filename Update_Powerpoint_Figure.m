function Update_Powerpoint_Figure(fig_number,b)
if (strcmp(b.Key,'p') && ~isempty(b.Modifier) && sum(strcmp(b.Modifier,'alt')))
    if(sum(strcmp(b.Modifier,'shift')))
        Update_All_Powerpoint_Figure();
    else
        Update_Single_Powerpoint_Figure(fig_number);
    end
end
end