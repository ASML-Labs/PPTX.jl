module PPTX

using XMLDict
using EzXML
using DataStructures
using ZipArchives:
    ZipBufferReader, ZipWriter, zip_commitfile, zip_newfile, zip_nentries,
    zip_name, zip_name_collision, zip_isdir, zip_readentry, zip_iscompressed

import Tables
import Tables: columns, columnnames, rows

import GridLayoutBase
import GridLayoutBase: GridLayout, LayoutObservables, BBox

import Observables
import Observables: Observable

export Presentation, Slide, TextBox, Picture, Table

export ShapeLayout

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
include("layout.jl")

end
