function out = RGB_int (RGB)
RGB=round((RGB*1.0)*255);
    red = RGB(1); green = RGB(2); blue = RGB(3);
    out = red  + (green * 256) + (blue * 65536);
end