using PPTX
using Test
import PPTX: slides, shapes, rid
import p7zip_jll
import Documenter

if !p7zip_jll.is_available()
    # see source code of Pkg.PlatformEngines.exe7z()
    # we encountered some unzipping issues on ubuntu runners in github
    @warn "p7zip_jll binary is not available"
end

include("testConstructors.jl")
include("testTables.jl")
include("testHyperlinks.jl")
include("testSlideXML.jl")
include("testWriting.jl")

@testset "Doctests" begin
    Documenter.doctest(PPTX)
end