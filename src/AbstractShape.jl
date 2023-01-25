abstract type AbstractShape end

# the 'relative identifier' is used to link shapes in the PowerPoint XML
set_rid!(s::AbstractShape, i::Int) = nothing
has_rid(s::AbstractShape) = false

## If AbstractShape does not have rId return 0
rid(s::AbstractShape) = 0