module PPTX

using XMLDict
using EzXML
using DataStructures
using ZipArchives:
    ZipBufferReader, ZipWriter, zip_commitfile, zip_newfile, zip_nentries,
    zip_name, zip_names, zip_name_collision, zip_isdir, zip_readentry, zip_iscompressed

import Tables
import Tables: columns, columnnames, rows

import Colors: Colorant, hex, @colorant_str

export Presentation, Slide, TextBox, TextStyle, Picture, Table, TableCell, list_layoutnames

include("AbstractShape.jl")
include("constants.jl")
include("TextBox.jl")
include("Picture.jl")
include("Tables.jl")
include("Slide.jl")
include("Presentation.jl")
include("xml_utils.jl")
include("xml_ppt_utils.jl")
include("write.jl")

end
