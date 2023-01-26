
"""
Interface functions:
* `Presentation`: Presentation constructor, used for making a presentation.
* `Slide`: Slide constructor, used for making a slide.
* `TextBox`: TextBox constructor, used for adding text to slides.
* `Picture`: Picture constructor, used for adding pictures to slides.
* `write`: write a presentation to a file.
"""
module PPTX

using XMLDict
using EzXML
using DataStructures

import Tables
import Tables: columns, columnnames, rows

import Pkg.PlatformEngines: exe7z

export Presentation, Slide, TextBox, Picture, Table

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
