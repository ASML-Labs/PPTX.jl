const _EMUS_PER_INCH = 914400
const _EMUS_PER_CENTIPOINT = 127
const _EMUS_PER_CM = 360000
const _EMUS_PER_MM = 36000
const _EMUS_PER_PT = 12700

mm_to_emu(::Nothing) = nothing
mm_to_emu(x) = Int(round(x * _EMUS_PER_MM))
mm_to_emu(x::AbstractArray{<:Real}) = mm_to_emu.(x)

inch_to_emu(::Nothing) = nothing
inch_to_emu(x::Real) = Int(round(x * _EMUS_PER_INCH))

points_to_emu(::Nothing) = nothing
points_to_emu(x::Real) = Int(round(x * _EMUS_PER_PT))

const TEMPLATE_DIR = abspath(joinpath(@__DIR__, "..", "templates"))
const ASSETS_DIR = abspath(joinpath(@__DIR__, "..", "assets"))
const TESTDATA_DIR = abspath(joinpath(@__DIR__, "..", "test/testdata"))

# we use layoutSlide1 for the first title slide, and layoutSlide2 for all other slides
const TITLE_SLIDE_LAYOUT = 1
const DEFAULT_SLIDE_LAYOUT = 2

include_dependency(joinpath(TEMPLATE_DIR, "tableStyles.xml"))
const DEFAULT_TABLE_STYLE_DATA = read(joinpath(TEMPLATE_DIR, "tableStyles.xml"))

include_dependency(joinpath(TEMPLATE_DIR, "no-slides.pptx"))
const DEFAULT_TEMPLATE_DATA = read(joinpath(TEMPLATE_DIR, "no-slides.pptx"))