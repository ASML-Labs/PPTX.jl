struct GridMargins # in EMUs
    left::Int
    right::Int
    top::Int
    bottom::Int
end

function Base.:(==)(m1::GridMargins, m2::GridMargins)
    return m1.left == m2.left && m1.right == m2.right && m1.top == m2.top && m1.bottom == m2.bottom
end

_grid_margin_mm(x::Real) = mm_to_emu(x)
_grid_margin_mm(::Nothing) = 0

function GridMargins(; left=0, right=0, top=0, bottom=0)
    return GridMargins(
        _grid_margin_mm(left),
        _grid_margin_mm(right),
        _grid_margin_mm(top),
        _grid_margin_mm(bottom),
    )
end

GridMargins(x::Real) = GridMargins(; left=x, right=x, top=x, bottom=x)
GridMargins(nt::NamedTuple) = GridMargins(; nt...)

"""
    GridLayout(
        nrows::Int, 
        ncols::Int; 
        padding::Real=5.0,
        margins=(left = padding, right = padding, top = padding, bottom = padding),
        rescale::Bool=true,
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
When `rescale=true` (default), pictures/videos preserve their aspect ratio and are
centered within the assigned cell span. With `rescale=false`, they are stretched
to fill the assigned span.

Rules:
- Empty cells are silently skipped.
- Assigning to overlapping cell regions raises an `ArgumentError`.
"""
mutable struct GridLayout <: AbstractShape
    nrows::Int
    ncols::Int
    padding::Float64
    margins::GridMargins
    rescale::Bool
    _entries::Vector{Tuple{AbstractShape,UnitRange{Int},UnitRange{Int}}}
end

function shapes(layout::GridLayout)
    return getfield.(layout._entries, 1)
end

function rid(layout::GridLayout)
    return maximum(rid.(shapes(layout)))
end

function GridLayout(
    nrows::Int, 
    ncols::Int; 
    padding::Real=5.0,
    margins=(left = padding, right = padding, top = padding, bottom = padding),
    rescale::Bool=true,
)
    nrows > 0 || throw(ArgumentError("nrows must be positive"))
    ncols > 0 || throw(ArgumentError("ncols must be positive"))
    padding >= 0 || throw(ArgumentError("padding must be non-negative"))
    grid_margins = GridMargins(margins)
    min_margin = min(grid_margins.left, grid_margins.right, grid_margins.top, grid_margins.bottom)
    min_margin >= 0 || throw(ArgumentError("margins must be non-negative"))
    return GridLayout(
        nrows,
        ncols,
        mm_to_emu(padding),
        grid_margins,
        rescale,
        Tuple{AbstractShape,UnitRange{Int},UnitRange{Int}}[],
    )
end

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

function layout_bounds(
    layout::GridLayout,
    row_range::UnitRange{Int},
    col_range::UnitRange{Int},
    slide_size_x::Int,
    slide_size_y::Int,
)
    left_margin = layout.margins.left
    right_margin = layout.margins.right
    top_margin = layout.margins.top
    bottom_margin = layout.margins.bottom
    gap = Int(round(layout.padding))

    usable_width = slide_size_x - left_margin - right_margin - (layout.ncols - 1) * gap
    usable_height = slide_size_y - top_margin - bottom_margin - (layout.nrows - 1) * gap
    usable_width > 0 || throw(ArgumentError("GridLayout has non-positive usable width"))
    usable_height > 0 || throw(ArgumentError("GridLayout has non-positive usable height"))

    cell_width = usable_width / layout.ncols
    cell_height = usable_height / layout.nrows

    first_col = first(col_range)
    last_col = last(col_range)
    first_row = first(row_range)
    last_row = last(row_range)

    offset_x = Int(round(left_margin + (first_col - 1) * (cell_width + gap)))
    offset_y = Int(round(top_margin + (first_row - 1) * (cell_height + gap)))
    size_x = Int(round((last_col - first_col + 1) * cell_width + (length(col_range) - 1) * gap))
    size_y = Int(round((last_row - first_row + 1) * cell_height + (length(row_range) - 1) * gap))

    return Geometry(offset_x, offset_y, size_x, size_y)
end

function _show_string(layout::GridLayout, compact::Bool)
    s = "GridLayout($(layout.nrows)×$(layout.ncols))"
    if !compact
        n = length(layout._entries)
        s *= " with $n assigned cell$(n == 1 ? "" : "s")"
    end
    return s
end

function copy_shape(w::ZipWriter, p::GridLayout)
    for (nested_shape, _, _) in p._entries
        copy_shape(w, nested_shape)
    end
end
