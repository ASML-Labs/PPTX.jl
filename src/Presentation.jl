struct PresentationSize
    x::Int # EMUs
    y::Int # EMUs
end

# presentation properties that are not for the user
# these may be gathered from the .pptx template upon writing
Base.@kwdef mutable struct PresentationState
    size::Union{Nothing, PresentationSize} = nothing
end

"""
```julia
Presentation(
    slides::Vector{Slide}=Slide[];
    title::String="unknown",
    author::String="unknown",
)
```

Type to contain the final presentation you want to write to .pptx.

If `isempty(slides)` then we add a first slide with the Title slide layout.

# Examples
```jldoctest
julia> using PPTX

julia> pres = Presentation(; title = "My Presentation")
Presentation with 1 slide
 title is "My Presentation"
 author is "unknown"

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
            slides = [Slide(; title=title, layout=TITLE_SLIDE_LAYOUT)]
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
    slides::Vector{Slide}=Slide[]; title::String="unknown", author::String="unknown"
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
    return push!(slides(pres), slide)
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

function make_relationships(p::Presentation)::AbstractDict
    ids = ["rId1", "rId2", "rId3", "rId4", "rId5"]
    relationship_tag_begin = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    types = ["slideMaster", "theme", "presProps", "viewProps", "tableStyles"]
    types = relationship_tag_begin .* types
    targets = [
        "slideMasters/slideMaster1.xml",
        "theme/theme1.xml",
        "presProps.xml",
        "viewProps.xml",
        "tableStyles.xml",
    ]
    relationships = OrderedDict(
        "Relationships" => Dict[OrderedDict(
            "xmlns" => "http://schemas.openxmlformats.org/package/2006/relationships"
        )],
    )
    for (id, type, target) in Base.zip(ids, types, targets)
        push!(
            relationships["Relationships"],
            OrderedDict(
                "Relationship" =>
                    OrderedDict("Id" => id, "Type" => type, "Target" => target),
            ),
        )
    end

    for (slide_idx, slide) in enumerate(slides(p))
        slide_rel = OrderedDict(
            "Relationship" => OrderedDict(
                "Id" => "rId$(slide.rid)",
                "Type" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                "Target" => "slides/slide$slide_idx.xml",
            ),
        )
        push!(relationships["Relationships"], slide_rel)
    end
    return relationships
end

function make_presentation(p::Presentation)
    xml_pres = OrderedDict("p:presentation" => main_attributes())
    push!(xml_pres["p:presentation"], OrderedDict("saveSubsetFonts" => "1"))

    push!(
        xml_pres["p:presentation"],
        OrderedDict(
            "p:sldMasterIdLst" => OrderedDict(
                "p:sldMasterId" => OrderedDict("id" => "2147483648", "r:id" => "rId1")
            ),
        ),
    )

    slide_id_list = Dict[]
    for (idx, slide) in enumerate(slides(p))
        push!(
            slide_id_list,
            OrderedDict(
                "p:sldId" => OrderedDict("id" => "$(idx+255)", "r:id" => "rId$(slide.rid)")
            ),
        )
    end

    push!(xml_pres["p:presentation"], OrderedDict("p:sldIdLst" => slide_id_list))

    sldSz = p._state.size
    push!(
        xml_pres["p:presentation"],
        OrderedDict("p:sldSz" => OrderedDict("cx" => "$(sldSz.x)", "cy" => "$(sldSz.y)")),
    )
    push!(
        xml_pres["p:presentation"],
        OrderedDict("p:notesSz" => OrderedDict("cx" => "6858000", "cy" => "9144000")),
    )
    return xml_pres
end

function update_presentation_state!(p::Presentation, ppt_dir=".")
    doc = readxml(joinpath(ppt_dir, "presentation.xml"))
    r = root(doc)
    n = findfirst("//p:sldSz", r)
    cx, cy = n["cx"], n["cy"]
    sz = PresentationSize(parse(Int, cx), parse(Int, cy))
    p._state.size = sz
    return nothing
end