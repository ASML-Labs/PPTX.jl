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

julia> img = Picture(joinpath(artifact"pptx_data", "assets", "cauliflower.jpg"))
Picture
 source is "./cauliflower.jpg"
 offset_x is 0 EMUs
 offset_y is 0 EMUs
 size_x is 1440000 EMUs
 size_y is 1475072 EMUs

```

Optionally, you can set the `size_x` and `size_y` manually for filetypes not supported by FileIO, such as SVG.
```julia
julia> using PPTX

julia> img = Picture(joinpath(artifact"pptx_data", "assets", "julia_logo.svg"); size_x=40, size_y=30)
Picture
 source is "./julia_logo.svg"
 offset_x is 0 EMUs
 offset_y is 0 EMUs
 size_x is 1440000 EMUs
 size_y is 1080000 EMUs

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

function Picture(
    source::String;
    top::Real=0,
    left::Real=0,
    offset_x::Real=left,
    offset_y::Real=top,
    size::Real=40,
    size_x::Real=size,
    size_y::Union{Nothing, Real}=nothing,
    rid::Int=0,
)
    scaled_size_x = Int(round(size_x * _EMUS_PER_MM))
    if isnothing(size_y)
        ratio = image_aspect_ratio(source)
        scaled_size_y = Int(round(scaled_size_x / ratio))
    else
        scaled_size_y = Int(round(size_y * _EMUS_PER_MM))
    end
    return Picture(
        source,
        Int(round(offset_x * _EMUS_PER_MM)),
        Int(round(offset_y * _EMUS_PER_MM)),
        scaled_size_x,
        scaled_size_y,
        rid,
    )
end

function set_rid(s::Picture, i::Int)
    return Picture(s.source, s.offset_x, s.offset_y, s.size_x, s.size_y, i)
end
rid(s::Picture) = s.rid
has_rid(s::Picture) = true

function _show_string(p::Picture, compact::Bool)
    show_string = "Picture"
    if !compact
        show_string *= "\n source is $(repr(p.source))"
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
    if endswith(lowercase(path), ".svg")
        doc = readxml(path)
        r = root(doc)
        m = match(r"(?<height>\d*)", r["height"])
        height = isnothing(m) ? 1 : parse(Float64, m[:height])
        m = match(r"(?<width>\d*)", r["width"])
        width = isnothing(m) ? 1 : parse(Float64, m[:width])
        return width / height
    end
    if endswith(lowercase(path), ".wmf")
        # https://winprotocoldoc.blob.core.windows.net/productionwindowsarchives/MS-WMF/%5bMS-WMF%5d.pdf
        f = read(path)
        # 2.3.2.3 META_PLACEABLE Record 
        key = reinterpret(UInt32, f[1:4])[1]
        iswmf = key == 0x9ac6cdd7
        # 2.2.2.18 Rect Object
        left, top, right, bottom = reinterpret(Int16, f[7:14])
        width = right - left
        height = bottom - top
        iswmf && return width / height
    end
    if endswith(lowercase(path), ".emf")
        # https://winprotocoldoc.blob.core.windows.net/productionwindowsarchives/MS-EMF/%5bMS-EMF%5d.pdf
        f = read(path)
        # 2.3.4.2 EMR_HEADER Record Types
        type = reinterpret(UInt32, f[1:4])[1]
        isemf = type == 0x00000001
        # 2.2.9 Header Object
        left, top, right, bottom = reinterpret(Int32, f[9:24]) # MS-WMF 2.2.2.19
        width = right - left
        height = bottom - top
        isemf && return width / height
    end
    
    local img
    try
        img = load(path)
    catch e
        if e isa ErrorException && contains(e.msg, "No applicable_loaders found")
            error("Cannot load image to determine aspect ratio, consider setting `size_x` and `size_y` manually.")
        else
            rethrow(e)
        end
    end
    height, width = size(img)
    return width / height
end
