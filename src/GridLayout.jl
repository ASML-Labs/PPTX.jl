"""
    GridLayout(
        nrows::Int, 
        ncols::Int; 
        padding::Real=5.0, 
        margins=(left = padding, right = padding, top = padding, bottom = padding)
    )

A declarative grid layout that distributes shapes evenly across a slide.

Shapes are placed in cells via `setindex!` using integer indices, ranges or colons:

```julia
layout = GridLayout(2, 2)
layout[1, 1] = TextBox("Title")
layout[2, :] = Picture("plot.png")  # spans both columns in row 2
push!(slide, layout)
```

Cell positions and sizes are resolved at write time from the actual slide dimensions.
`padding` (in mm) is applied between cells and at the slide edges.
Shapes are stretched to fill their assigned cell area (including any column/row span).

Rules:
- Empty cells are silently skipped.
- Assigning to overlapping cell regions raises an `ArgumentError`.
"""
mutable struct GridLayout <: AbstractShape
    nrows::Int
    ncols::Int
    padding::Float64 # mm to EMUs
    margins::Margins # mm to EMUs
    _entries::Vector{Tuple{AbstractShape,UnitRange{Int},UnitRange{Int}}}
end

function GridLayout(
    nrows::Int, 
    ncols::Int; 
    padding::Real=5.0, 
    margins=(left = padding, right = padding, top = padding, bottom = padding),
)
    nrows > 0 || throw(ArgumentError("nrows must be positive"))
    ncols > 0 || throw(ArgumentError("ncols must be positive"))
    padding >= 0 || throw(ArgumentError("padding must be non-negative"))
    return GridLayout(
        nrows,
        ncols,
        mm_to_emu(padding),
        Margins(margins),
        Tuple{AbstractShape,UnitRange{Int},UnitRange{Int}}[],
    )
end

# ----- index normalization -----

function _to_range(idx, n::Int, dim::String)::UnitRange{Int}
    if idx isa Colon
        return 1:n
    elseif idx isa Integer
        i = Int(idx)
        1 <= i <= n || throw(BoundsError("$dim index $i is out of range 1:$n"))
        return i:i
    elseif idx isa UnitRange
        r = UnitRange{Int}(idx)
        (1 <= first(r) && last(r) <= n) ||
            throw(BoundsError("$dim range $r is out of bounds 1:$n"))
        return r
    else
        throw(ArgumentError("unsupported index type $(typeof(idx)) for $dim"))
    end
end

# ----- setindex! -----

function Base.setindex!(layout::GridLayout, shape::AbstractShape, row_idx, col_idx)
    row_range = _to_range(row_idx, layout.nrows, "row")
    col_range = _to_range(col_idx, layout.ncols, "col")

    # conflict detection: two rectangular spans conflict when they overlap in both dimensions
    for (_, er, ec) in layout._entries
        if !isempty(intersect(row_range, er)) && !isempty(intersect(col_range, ec))
            throw(ArgumentError(
                "Cell span ($row_range, $col_range) conflicts with existing entry at ($er, $ec)"
            ))
        end
    end

    push!(layout._entries, (shape, row_range, col_range))
    return shape
end

# ----- show -----

function _show_string(layout::GridLayout, compact::Bool)
    s = "GridLayout($(layout.nrows)×$(layout.ncols))"
    if !compact
        n = length(layout._entries)
        s *= " with $n assigned cell$(n == 1 ? "" : "s")"
    end
    return s
end
