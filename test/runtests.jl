using PPTX
using Test
import PPTX: slides, shapes, rid

include("testConstructors.jl")
include("testTables.jl")
include("testHyperlinks.jl")
include("testSlideXML.jl")
include("testWriting.jl")
include("testLayout.jl")