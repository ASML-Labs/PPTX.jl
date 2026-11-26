# presentation properties that are not for the user
# these may be gathered from the .pptx template upon writing
Base.@kwdef mutable struct PresentationState
    size::Union{Nothing, SlideSize} = nothing
end

"""
```julia
Presentation(
    slides::Vector{Slide}=Slide[];
    title::String="My Presentation",
    author::String="PPTX.jl",
)
```

Type to contain the final presentation you want to write to .pptx.

If `isempty(slides)` then we add a first slide with the Title slide layout.

# Examples
```jldoctest
julia> using PPTX

julia> pres = Presentation(; title = "My Presentation", author = "PPTX.jl")
Presentation with 1 slide
 title is "My Presentation"
 author is "PPTX.jl"

```
"""
struct Presentation
    title::String
    author::String
    slides::Vector{Slide}
    _state::PresentationState
    function Presentation(slides::Vector{Slide}, author::String, title::String)
        pres = new(title, author, Slide[], PresentationState())
        if isempty(slides)
            slides = [Slide(; title=title, layout=TITLE_SLIDE_LAYOUT, slide_nr = 1)]
        end
        for slide in slides
            push!(pres, slide)
        end
        return pres
    end
end
slides(p::Presentation) = p.slides

# keyword argument constructor
function Presentation(
    slides::Vector{Slide}=Slide[]; title::String="My Presentation", author::String="PPTX.jl"
)
    return Presentation(slides, author, title)
end

function new_rid(pres::Presentation)
    if isempty(slides(pres))
        return 6
    else
        return maximum(rid.(slides(pres))) + 1
    end
end

function Base.push!(pres::Presentation, slide::Slide)
    slide.rid = new_rid(pres)
    slide.slide_nr = length(slides(pres)) + 1
    return push!(slides(pres), slide)
end

function slide_size(p::Presentation)
    p_sz = p._state.size
    sz = isnothing(p_sz) ? SlideSize() : p_sz
    return sz
end

# default show used by Array show
function Base.show(io::IO, p::Presentation)
    compact = get(io, :compact, true)
    return print(io, _show_string(p, compact))
end

# default show used by display() on the REPL
function Base.show(io::IO, mime::MIME"text/plain", p::Presentation)
    compact = get(io, :compact, false)
    return print(io, _show_string(p, compact))
end

function _show_string(p::Presentation, compact::Bool)
    show_string = ""
    nslides = length(p.slides)
    if nslides == 1
        slide_string = "slide"
    else
        slide_string = "slides"
    end
    show_string *= "Presentation with $(nslides) $slide_string"
    if !compact
        show_string *= "\n title is \"$(p.title)\""
        show_string *= "\n author is \"$(p.author)\""
    end
    return show_string
end

function append_slide_relationships!(doc::EzXML.Document, p::Presentation)
    rels = root(doc)
    for node in collect(elements(rels))
        haskey(node, "Type") && endswith(node["Type"], "/slide") && unlink!(node)
    end
    existing = Set(node["Id"] for node in elements(rels))
    next_id = 1
    for slide in slides(p)
        while "rId$next_id" in existing
            next_id += 1
        end
        slide.rid = next_id
        rid = "rId$next_id"
        push!(existing, rid)
        rel = ElementNode("Relationship")
        link_node!(rel, "Id" => rid)
        link_node!(rel, "Type" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide")
        link_node!(rel, "Target" => "slides/$(slide_fname(slide))")
        link!(rels, rel)
        next_id += 1
    end
    doc
end

function append_slides_to_presentation!(doc::EzXML.Document, p::Presentation)
    pres = root(doc)
    slide_list = findfirst("//*[local-name()='sldIdLst']", pres)
    if isnothing(slide_list)
        slide_size = findfirst("//*[local-name()='sldSz']", pres)
        isnothing(slide_size) && error("template presentation is missing p:sldSz")
        slide_list = ElementNode("p:sldIdLst")
        linkprev!(slide_size, slide_list)
    else
        for node in collect(elements(slide_list))
            unlink!(node)
        end
    end
    next_id = 256
    for slide in slides(p)
        slide_id = ElementNode("p:sldId")
        link_node!(slide_id, "id" => string(next_id))
        link_node!(slide_id, "r:id" => "rId$(slide.rid)")
        link!(slide_list, slide_id)
        next_id += 1
    end
    doc
end

function update_presentation_state!(p::Presentation, template::ZipBufferReader)
    doc = EzXML.parsexml(zip_readentry(template, "ppt/presentation.xml"))
    r = root(doc)
    n = findfirst("//p:sldSz", r)
    cx, cy = n["cx"], n["cy"]
    sz = SlideSize(parse(Int, cx), parse(Int, cy))
    p._state.size = sz
    return nothing
end

# the slide filename (e.g. slide2.xml) can be used in the slide relationships to enable hyperlinks for example
# We need to make sure that slide N is also really written to slideN.xml therefore slide numbers are updated just before writing 
function update_slide_nrs!(p::Presentation)
    for (index, slide) in enumerate(slides(p))
        slide.slide_nr = index
    end
end
