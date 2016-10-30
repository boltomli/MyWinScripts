$fonts = (New-Object -ComObject Shell.Application).Namespace(0x14)
dir -s *.ttf | %{ $fonts.CopyHere($_.FullName) }
dir -s *.ttc | %{ $fonts.CopyHere($_.FullName) }
dir -s *.otf | %{ $fonts.CopyHere($_.FullName) }
