module PPTX

using DataStructures
using EzXML
using VideoIO: openvideo
using XMLDict
using ZipArchives:
    ZipBufferReader, ZipWriter, zip_commitfile, zip_newfile, zip_nentries,
    zip_name, zip_names, zip_name_collision, zip_isdir, zip_readentry, zip_iscompressed

import Tables
import Tables: columns, columnnames, rows

import Colors: Colorant, hex, @colorant_str

export Presentation, Slide, TextBox, TextStyle, Picture, Table, TableCell, list_layoutnames, Video

include("AbstractShape.jl")
include("constants.jl")
include("TextBox.jl")
include("Picture.jl")
include("Tables.jl")
include("Slide.jl")
include("Presentation.jl")
include("Video.jl")
include("xml_utils.jl")
include("xml_ppt_utils.jl")
include("write.jl")

end
