using FileIO, ImageIO

"""
```julia
Picture(source::String; top::Int=0, left::Int=0, size::Int = 40)
```

* `source::String` path of image file
* `top::Int` mm from the top
* `left::Int` mm from the left

Internally the sizes are converted EMUs.

# Examples
```julia
julia> using PPTX

julia> img = Picture(joinpath(PPTX.ASSETS_DIR, "cauliflower.jpg"))
Picture
 source is "./cauliflower.jpg"
 offset_x is 0 EMUs
 offset_y is 0 EMUs
 size_x is 1440000 EMUs
 size_y is 1475072 EMUs

```
"""
struct Picture <: AbstractShape
    source::String
    offset_x::Int
    offset_y::Int
    size_x::Int
    size_y::Int
    rid::Int
end

function Picture(source::String; top::Real=0, left::Real=0, offset_x::Real=left, offset_y::Real=top, size::Real=40, rid::Int=0)
    ratio = image_aspect_ratio(source)
    size_x = Int(round(size * _EMUS_PER_MM))
    size_y = Int(round(size_x / ratio))
    return Picture(
        source,
        Int(round(offset_x * _EMUS_PER_MM)),
        Int(round(offset_y * _EMUS_PER_MM)),
        size_x,
        size_y,
        rid,
    )
end

set_rid(s::Picture, i::Int) = Picture(s.source, s.offset_x, s.offset_y, s.size_x, s.size_y, i)
rid(s::Picture) = s.rid
has_rid(s::Picture) = true

function _show_string(p::Picture, compact::Bool)
    show_string = "Picture"
    if !compact
        show_string *= "\n source is \"$(p.source)\""
        show_string *= "\n offset_x is $(p.offset_x) EMUs"
        show_string *= "\n offset_y is $(p.offset_y) EMUs"
        show_string *= "\n size_x is $(p.size_x) EMUs"
        show_string *= "\n size_y is $(p.size_y) EMUs"
    end
    return show_string
end

function make_xml(shape::Picture, id::Int)
    cNvPr = Dict("p:cNvPr" => [Dict("id" => "$id"), Dict("name" => "Picture")])
    cNvPicPr = Dict("p:cNvPicPr" => Dict("a:picLocks" => Dict("noChangeAspect" => "1")))
    nvPr = Dict("p:nvPr" => missing)
    nvPicPr = Dict("p:nvPicPr" => [cNvPr, cNvPicPr, nvPr])

    blip = Dict("a:blip" => Dict("r:embed" => "rId$(rid(shape))"))
    stretch = Dict("a:stretch" => Dict("a:fillRect" => missing))
    blipFill = Dict("p:blipFill" => [blip, stretch])

    xfrm = Dict(
        "a:xfrm" => [
            Dict(
                "a:off" => [
                    Dict("x" => "$(shape.offset_x)"),
                    Dict("y" => "$(shape.offset_y)"),
                ],
            ),
            Dict(
                "a:ext" =>
                    [Dict("cx" => "$(shape.size_x)"), Dict("cy" => "$(shape.size_y)")],
            ),
        ],
    )
    prstgeom = Dict("a:prstGeom" => [Dict("prst" => "rect"), Dict("a:avLst" => missing)])
    spPr = Dict("p:spPr" => [xfrm, prstgeom])

    return Dict("p:pic" => [nvPicPr, blipFill, spPr])
end

function type_schema(p::Picture)
    return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
end
filename(p::Picture) = splitpath(p.source)[end]

function relationship_xml(p::Picture)
    return Dict(
        "Relationship" => [
            Dict("Id" => "rId$(rid(p))"),
            Dict("Type" => type_schema(p)),
            Dict("Target" => "../media/$(filename(p))"),
        ],
    )
end

function copy_picture(w::ZipWriter, p::Picture)
    dest_path = "ppt/media/$(filename(p))"
    # save any file being written so zip_name_collision is correct.
    zip_commitfile(w) 
    if !zip_name_collision(w, dest_path)
        open(p.source) do io
            zip_newfile(w, dest_path)
            write(w, io)
            zip_commitfile(w)
        end
    end
end

function image_aspect_ratio(path::String)
    img = load(path)
    height, width = size(img)
    return width / height
end

