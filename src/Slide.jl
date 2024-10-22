"""
```julia
Slide(
    shapes::Vector{AbstractShape}=AbstractShape[];
    title::String="",
    layout::Int=1,
)
```

* `shapes::Vector{AbstractShape}` shapes to add to the PowerPoint, can also be pushed afterwards
* `title::String` title text placed inside the title textbox found in the slide layout
* `layout::Int` which slide layout to use. Typically 1 is the title slide and 2 is the text slide.

Make a Slide for a powerpoint Presentation.

You can `push!` any `AbstractShape` types into this slide, such as a `TextBox` or `Picture`.

# Examples
```julia
julia> using PPTX

julia> slide = Slide(; title="Hello Title", layout=2)
Slide("Hello Title", PPTX.AbstractShape[], 0, 2)

julia> text = TextBox("Hello world!")
TextBox
 content is "Hello world!"
 offset_x is 1800000 EMUs
 offset_y is 1800000 EMUs
 size_x is 1440000 EMUs
 size_y is 1080000 EMUs

julia> push!(slide, text);

julia> slide
Slide("Hello Title", PPTX.AbstractShape[TextBox], 0, 2)

```
"""
mutable struct Slide
    title::String
    shapes::Vector{AbstractShape}
    rid::Int
    layout::Int
    slide_nr::Int
    function Slide(
        title::String,
        shapes::Vector{<:AbstractShape},
        rid::Int=0,
        layout::Int=DEFAULT_SLIDE_LAYOUT,
        slide_nr::Int=0
    )
        slide = new(title, AbstractShape[], rid, layout, slide_nr)
        for shape in shapes
            push!(slide, shape)
        end
        return slide
    end
end
shapes(s::Slide) = s.shapes
rid(s::Slide) = s.rid
slide_nr(s::Slide) = s.slide_nr
Slide(shapes::Vector{<:AbstractShape}; kwargs...) = Slide(; shapes=shapes, kwargs...)

function Slide(;
    title::String="",
    shapes::Vector{<:AbstractShape}=AbstractShape[],
    rid::Int=0,
    layout::Int=DEFAULT_SLIDE_LAYOUT,
    slide_nr::Int=0
)
    return Slide(title, shapes, rid, layout, slide_nr)
end

function new_rid(slide::Slide)
    # When there is no rid shapes the new rid should be 2
    if isempty(slide.shapes)
        return 2
    else # default 2, or it should be the at least 1 bigger than the shape with the higest rid
        return max(2, maximum(rid.(shapes(slide))) + 1)
    end
end

slide_fname(s::Slide) = "slide$(s.slide_nr).xml"

function Base.push!(slide::Slide, shape::AbstractShape)
    if has_rid(shape)
        push!(shapes(slide), set_rid(shape, new_rid(slide)))
    else
        push!(shapes(slide), shape)
    end
end

function make_slide(s::Slide, relationship_map::Dict = slide_relationship_map(s))::AbstractDict
    xml_slide = OrderedDict("p:sld" => main_attributes())

    spTree = init_sptree()
    initial_max_id = 1
    for (index, shape) in enumerate(shapes(s))
        id = index + initial_max_id
        push!(spTree["p:spTree"], make_xml(shape, id, relationship_map))
    end

    push!(xml_slide["p:sld"], OrderedDict("p:cSld" => [spTree]))
    return xml_slide
end

# spTree means 'shape tree' I guess
function init_sptree()::AbstractDict
    return Dict(
        "p:spTree" => Any[
            Dict(
                "p:nvGrpSpPr" => [
                    OrderedDict(
                        "p:cNvPr" =>
                            [OrderedDict("id" => "1"), OrderedDict("name" => "")],
                    ),
                    OrderedDict("p:cNvGrpSpPr" => missing),
                    OrderedDict("p:nvPr" => missing),
                ],
            ),
            Dict(
                "p:grpSpPr" => OrderedDict(
                    "a:xfrm" => [
                        OrderedDict(
                            "a:off" => [OrderedDict("x" => "0"), OrderedDict("y" => "0")],
                        ),
                        OrderedDict(
                            "a:ext" => [OrderedDict("cx" => "0"), OrderedDict("cy" => "0")],
                        ),
                        OrderedDict(
                            "a:chOff" => [OrderedDict("x" => "0"), OrderedDict("y" => "0")],
                        ),
                        OrderedDict(
                            "a:chExt" => [
                                OrderedDict("cx" => "0"),
                                OrderedDict("cy" => "0"),
                            ],
                        ),
                    ],
                ),
            ),
        ],
    )
end

function type_schema(s::Slide)
    return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
end

function relationship_xml(s::Slide, r_id::Integer)
    return Dict(
        "Relationship" => [
            Dict("Id" => "rId$r_id"),
            Dict("Type" => type_schema(s)),
            Dict("Target" => slide_fname(s)),
        ],
    )
end

function type_schema(url::String)
    return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
end

# <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://github.com/ASML-Labs/PPTX.jl" TargetMode="External"/>
function relationship_xml(url::AbstractString, r_id::Integer)
    return Dict(
        "Relationship" => [
            Dict("Id" => "rId$r_id"),
            Dict("Type" => type_schema(url)),
            Dict("Target" => string(url)),
            Dict("TargetMode" => "External")
        ],
    )
end

function slide_relationship_map(s::Slide)
    d = Dict{Union{Slide, AbstractShape, AbstractString}, Int}()
    r_id = 1 # first rid is reserved by slideLayout
    for shape in shapes(s)
        if has_rid(shape) && !haskey(d, shape)
            r_id += 1
            d[shape] = r_id
        elseif has_hyperlink(shape) && !haskey(d, shape.hlink)
            r_id += 1
            d[shape.hlink] = r_id
        end
    end
    return d
end

function make_slide_relationships(s::Slide, relationship_map::Dict = slide_relationship_map(s))::AbstractDict
    xml_slide_rels = OrderedDict("Relationships" => Dict[])
    push!(
        xml_slide_rels["Relationships"],
        OrderedDict(
            "xmlns" => "http://schemas.openxmlformats.org/package/2006/relationships"
        ),
    )
    push!(
        xml_slide_rels["Relationships"],
        OrderedDict(
            "Relationship" => OrderedDict(
                "Id" => "rId1",
                "Type" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
                "Target" => "../slideLayouts/slideLayout$(s.layout).xml",
            ),
        ),
    )
    used_r_ids = [1]
    for shape in shapes(s)
        r_id = 1
        if has_rid(shape)
            r_id = relationship_map[shape]
            r_shape = shape
        end
        if has_hyperlink(shape)
            r_id = relationship_map[shape.hlink]
            r_shape = shape.hlink
        end
        if r_id ∉ used_r_ids
            push!(xml_slide_rels["Relationships"], relationship_xml(r_shape, r_id))
            push!(used_r_ids, r_id)
        end
    end
    return xml_slide_rels
end