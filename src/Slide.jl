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

julia> text = TextBox("Hello world!")

julia> push!(slide, text)

julia> pres = Presentation()

julia> push!(pres, slide)

julia> write("hello_world.pptx", pres)
```
"""
mutable struct Slide
    title::String
    shapes::Vector{AbstractShape}
    rid::Int
    layout::Int
    function Slide(
        title::String,
        shapes::Vector{<:AbstractShape},
        rid::Int=0,
        layout::Int=DEFAULT_SLIDE_LAYOUT,
    )
        slide = new(title, AbstractShape[], rid, layout)
        for shape in shapes
            push!(slide, shape)
        end
        return slide
    end
end
shapes(s::Slide) = s.shapes
rid(s::Slide) = s.rid

Slide(shapes::Vector{<:AbstractShape}; kwargs...) = Slide(; shapes=shapes, kwargs...)

function Slide(;
    title::String="",
    shapes::Vector{<:AbstractShape}=AbstractShape[],
    rid::Int=0,
    layout::Int=DEFAULT_SLIDE_LAYOUT,
)
    return Slide(title, shapes, rid, layout)
end

function new_rid(slide::Slide)
    if isempty(shapes(slide))
        return 2
    else
        return maximum(rid.(shapes(slide))) + 1
    end
end

function Base.push!(slide::Slide, shape::AbstractShape)
    if has_rid(shape)
        push!(shapes(slide), set_rid(shape, new_rid(slide)))
    else
        push!(shapes(slide), shape)
    end
end

function make_slide(s::Slide)::AbstractDict
    xml_slide = OrderedDict("p:sld" => main_attributes())

    spTree = init_sptree()
    initial_max_id = 1
    for (index, shape) in enumerate(shapes(s))
        id = index + initial_max_id
        push!(spTree["p:spTree"], make_xml(shape, id))
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

function make_slide_relationships(s::Slide)::AbstractDict
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
    for shape in shapes(s)
        if has_rid(shape)
            push!(xml_slide_rels["Relationships"], relationship_xml(shape))
        end
    end
    return xml_slide_rels
end