abstract type AbstractShape end

# the 'relative identifier' is used to link shapes in the PowerPoint XML
set_rid!(s::AbstractShape, i::Int) = nothing
has_rid(s::AbstractShape) = false

## If AbstractShape does not have rId return 0
rid(s::AbstractShape) = 0

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

hlink_xml(hlink) = Dict("a:hlinkClick" => Dict("rId" => "rId$(rid(hlink))", "action" => "ppaction://hlinksldjump"))
has_hyperlink(s::AbstractShape) = hasfield(typeof(s), :hlink) && !isnothing(s.hlink)