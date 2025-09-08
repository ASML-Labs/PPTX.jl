using Colors
using PPTX
using Test

import PPTX: slides, shapes, rid

include("testAbstractShape.jl")
include("testHyperlinks.jl")
include("testLayout.jl")
include("testPicture.jl")
include("testPresentation.jl")
# include("testPresentationState.jl") # This testset seems outdated and or obsolete
include("testSlide.jl")
include("testSlideXML.jl")
include("testTables.jl")
include("testTextBox.jl")
include("testVideo.jl")
include("testWriting.jl")