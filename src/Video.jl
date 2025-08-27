using UUIDs

struct Video <: AbstractShape
    source::String
    offset_x::Int
    offset_y::Int
    size_x::Int
    size_y::Int
    rid::Int
    _uuid::String
    function Video(source::String, offset_x::Int, offset_y::Int, size_x::Int, size_y::Int, rid::Int)
        new(source, offset_x, offset_y, size_x, size_y, rid, string(UUIDs.uuid4()))
    end
end

function Video(
    source::String;
    top::Real=0,
    left::Real=0,
    offset=(left, top),
    offset_x::Real=offset[1],
    offset_y::Real=offset[2],
    size::Real=40,
    size_x::Real=size,
    size_y::Union{Nothing, Real}=nothing,
    rid::Int=0,
)
    scaled_size_x = Int(round(size_x * _EMUS_PER_MM))
    scaled_size_y = isnothing(size_y) ? scaled_size_x : Int(round(size_y * _EMUS_PER_MM))
    return Video(
        source,
        Int(round(offset_x * _EMUS_PER_MM)),
        Int(round(offset_y * _EMUS_PER_MM)),
        scaled_size_x,
        scaled_size_y,
        rid,
    )
end

function set_rid(v::Video, i::Int)
    return Video(v.source, v.offset_x, v.offset_y, v.size_x, v.size_y, i)
end
rid(v::Video) = v.rid
has_rid(v::Video) = true

function _show_string(v::Video, compact::Bool)
    show_string = "Video"
    if !compact
        show_string *= "\n source is $(repr(v.source))"
        show_string *= "\n offset_x is $(v.offset_x) EMUs"
        show_string *= "\n offset_y is $(v.offset_y) EMUs"
        show_string *= "\n size_x is $(v.size_x) EMUs"
        show_string *= "\n size_y is $(v.size_y) EMUs"
    end
    return show_string
end

function make_xml(shape::Video, id::Int, relationship_map::Dict)
    rel_id = relationship_map[shape]
    video_rid = "rId$rel_id"

    cNvPr = Dict("p:cNvPr" => [
        Dict("id" => "$id"),
        Dict("name" => "Video"),
        Dict("a:hlinkClick" => [
            Dict("r:id" => ""),
            Dict("action" => "ppaction://media")
        ]),
        Dict("a:extLst" => [
            Dict("a:ext" => [
                Dict("uri" => "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"),
                Dict("a16:creationId" => [
                    Dict("xmlns:a16" => "http://schemas.microsoft.com/office/drawing/2014/main"),
                    Dict("id" => "{$(UUIDs.uuid4())}")
                ])
            ])
        ])
    ])

    cNvPicPr = Dict("p:cNvPicPr" => [
        Dict("a:picLocks" => Dict("noChangeAspect" => "1"))
    ])

    nvPr = Dict("p:nvPr" => [
        Dict("a:videoFile" => [Dict("r:link" => "rId$(rel_id+1)")]),
        Dict("p:extLst" => [
            Dict("p:ext" => [
                Dict("uri" => "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"),
                Dict("p14:media" => [
                    Dict("xmlns:p14" => "http://schemas.microsoft.com/office/powerpoint/2010/main"),
                    Dict("r:embed" => video_rid)
                ])
            ])
        ])
    ])

    nvPicPr = Dict("p:nvPicPr" => [cNvPr, cNvPicPr, nvPr])

    blip = Dict("a:blip" => [Dict("r:embed" => "rId$(rel_id+2)")])
    stretch = Dict("a:stretch" => [Dict("a:fillRect" => missing)])
    blipFill = Dict("p:blipFill" => [blip, stretch])

    xfrm = Dict("a:xfrm" => [
        Dict("a:off" => [Dict("x" => "$(shape.offset_x)"), Dict("y" => "$(shape.offset_y)")]),
        Dict("a:ext" => [Dict("cx" => "$(shape.size_x)"), Dict("cy" => "$(shape.size_y)")])
    ])
    prstgeom = Dict("a:prstGeom" => [
        Dict("prst" => "rect"),
        Dict("a:avLst" => missing)
    ])
    spPr = Dict("p:spPr" => [xfrm, prstgeom])

    return Dict("p:pic" => [nvPicPr, blipFill, spPr])
end

function type_schema(v::Video; it = 0)
    if it == 0
        return "http://schemas.microsoft.com/office/2007/relationships/media"
    else
        return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
    end
end

function filename(v::Video)
    source_filename, extension = splitext(splitpath(v.source)[end])
    return "$(source_filename)$(v._uuid)$extension"
end

function relationship_xml(v::Video, r_id::Integer; it::Integer=0)
    return Dict(
            "Relationship" => [
                Dict("Id" => "rId$r_id"),
                Dict("Type" => type_schema(v; it = it)),
                Dict("Target" => "../media/$(filename(v))"),
            ],
        )
end

function copy_shape(w::ZipWriter, v::Video)
    dest_path_vid = "ppt/media/$(filename(v))"
    zip_commitfile(w)
    if !zip_name_collision(w, dest_path_vid)
        open(v.source) do io
            zip_newfile(w, dest_path_vid)
            write(w, io)
            zip_commitfile(w)
        end
    end

    dest_path_thumbnail = "ppt/media/$(thumbnail_name(v))"
    thumbnail = create_thumbnail(v)
    # Save to a temporary PNG file
    tmp_path = tempname() * ".png"
    save(tmp_path, thumbnail)

    # Write the PNG file into the ZIP
    zip_commitfile(w)
    if !zip_name_collision(w, dest_path_thumbnail)
        zip_newfile(w, dest_path_thumbnail)
        open(tmp_path, "r") do io
            write(w, io)
        end
        zip_commitfile(w)
    end
end

function thumbnail_name(v::Video)
    name_without_extension = split(filename(v), ".")[begin]
    return name_without_extension*"_thumbnail.png"
end

function create_thumbnail(v::Video)
    video = openvideo(v.source)
    # Read the first frame
    frame = read(video)
    return frame
end
