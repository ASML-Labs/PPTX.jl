abstract type AbstractShape end

# the 'relative identifier' is used to link shapes in the PowerPoint XML
set_rid!(s::AbstractShape, i::Int) = nothing
has_rid(s::AbstractShape) = false

## If AbstractShape does not have rId return 0
rid(s::AbstractShape) = 0

# struct to hold common shape geometry properties
# (could be re-used in the future inside the shape)
struct Geometry
    offset_x::Int
    offset_y::Int
    size_x::Int
    size_y::Int
end

# default do not change geometries (this is used for GridLayout, where the geometry is determined by the cell span, not the shape)
set_geometry(s::AbstractShape, geom::Geometry) = s
function get_geometry(s::AbstractShape)
    return Geometry(s.offset_x, s.offset_y, s.size_x, s.size_y)
end

# default no geometry rescaling (e.g. for TextBox, where the text should fill the cell span without changing the aspect ratio)
geometry_in_span(s::AbstractShape, geom::Geometry, keepratio::Bool) = geom

# geometry rescaling re-used by Picture and Video, keeps the ratio constant
function geometry_in_span(shape_geom::Geometry, span::Geometry, keepratio::Bool)
    if !keepratio || shape_geom.size_x <= 0 || shape_geom.size_y <= 0 || span.size_x <= 0 || span.size_y <= 0
        return span
    end

    # look at the scaling factor for each dimension and use the smaller one to ensure the shape fits within the span
    scale = min(span.size_x / shape_geom.size_x, span.size_y / shape_geom.size_y)
    # rescale the shape dimensions by the calculated factor
    scaled_x = Int(round(shape_geom.size_x * scale))
    scaled_y = Int(round(shape_geom.size_y * scale))
    centered_x = span.offset_x + Int(round((span.size_x - scaled_x) / 2))
    centered_y = span.offset_y + Int(round((span.size_y - scaled_y) / 2))
    return Geometry(centered_x, centered_y, scaled_x, scaled_y)
end

# default show used by Array show
function Base.show(io::IO, shape::AbstractShape)
    compact = get(io, :compact, true)
    return print(io, _show_string(shape, compact))
end

# default show used by display() on the REPL
function Base.show(io::IO, mime::MIME"text/plain", shape::AbstractShape)
    compact = get(io, :compact, false)
    return print(io, _show_string(shape, compact))
end

function _show_string(shape::AbstractShape, compact::Bool)
    return "$(typeof(shape))"
end

function hlink_xml(hlink, relationship_map::Dict)
    rel_id = relationship_map[hlink]
    Dict("a:hlinkClick" => Dict("r:id" => "rId$rel_id", "action" => "ppaction://hlinksldjump"))
end
# for urls
function hlink_xml(hlink::String, relationship_map::Dict)
    rel_id = relationship_map[hlink]
    Dict("a:hlinkClick" => Dict("r:id" => "rId$rel_id"))
end
has_hyperlink(s::AbstractShape) = hasfield(typeof(s), :hlink) && !isnothing(s.hlink)