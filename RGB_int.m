function out = RGB_int (RGB)
    red = RGB(1); green = RGB(2); blue = RGB(3);
    out = red  + (green * 256) + (blue * 65536);
end