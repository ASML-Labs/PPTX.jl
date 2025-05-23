"""
```julia
Table(;
    content,
    offset_x::Real = 50,
    offset_y::Real = 50,
    size_x::Real = 150,
    size_y::Real = 100,
    column_widths::Vector{<:Real}, # set size per column
    row_heights::Vector{<:Real}, # set size per row
    header::Bool = true, # whether to automatically write the columnnames as headers
    bandrow::Bool = true, # whether to use alternating coloring per row
)
```

A Table to be used on a Slide.

The content can be anything that adheres to a `Tables.jl` interface.

Offsets and sizes are in millimeters, but will be converted to EMU.

To style each cell individually see `TableCell`.

# Examples
```jldoctest
julia> using PPTX, DataFrames

julia> df = DataFrame(a = [1,2], b = [3,4], c = [5,6])
2×3 DataFrame
 Row │ a      b      c     
     │ Int64  Int64  Int64 
─────┼─────────────────────
   1 │     1      3      5
   2 │     2      4      6

julia> t = Table(content=df, size_x=30)
Table
 content isa DataFrames.DataFrame
 offset_x is 1800000 EMUs
 offset_y is 1800000 EMUs
 size_x is 1080000 EMUs
 size_y is 3600000 EMUs

```
"""
struct Table <: AbstractShape
    content # anything that adheres to Tables.jl interfaces
    offset_x::Int # EMUs
    offset_y::Int # EMUs
    size_x::Int # EMUs
    size_y::Int # EMUs
    column_widths::Union{Nothing, Vector{Int}}
    row_heights::Union{Nothing, Vector{Int}}
    header::Bool
    bandrow::Bool
    style_id::String
    function Table(
        content,
        offset_x::Real, # millimeters
        offset_y::Real, # millimeters
        size_x::Real, # millimeters
        size_y::Real, # millimeters
        column_widths::Union{Nothing, Vector{Int}} = nothing,
        row_heights::Union{Nothing, Vector{Int}} = nothing,
        header::Bool = true,
        bandrow::Bool = true,
        style_id::String="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",
    )
        # input is in mm
        return new(
            content,
            mm_to_emu(offset_x),
            mm_to_emu(offset_y),
            mm_to_emu(size_x),
            mm_to_emu(size_y),
            mm_to_emu(column_widths),
            mm_to_emu(row_heights),
            header,
            bandrow,
            style_id,
        )
    end
end

# keyword argument constructor
function Table(;
    content,
    offset=(50,50),
    offset_x::Real=offset[1], # millimeters
    offset_y::Real=offset[2], # millimeters
    size=(150, 100),
    size_x::Real=size[1], # millimeters
    size_y::Real=size[2], # millimeters
    column_widths=nothing,
    row_heights=nothing,
    header::Bool=header_default(content),
    bandrow::Bool=true,
)
    t = Table(content, offset_x, offset_y, size_x, size_y, column_widths, row_heights, header, bandrow)
    check_size(t)
    return t
end

Table(content; kwargs...) = Table(; content=content, kwargs...)

header_default(x) = true
header_default(x::AbstractArray) = false

Tables.columnnames(t::Table) = Tables.columnnames(t.content)
Tables.columns(t::Table) = get_columns(t.content)
Tables.rows(t::Table) = get_rows(t.content)
ncols(t::Table) = length(get_columns(t))
nrows(t::Table) = length(get_rows(t))

get_columns(t::Table) = get_columns(t.content)
get_columns(t) = Tables.columns(t)
get_columns(m::AbstractMatrix) = eachcol(m)
get_rows(t::Table) = get_rows(t.content)
get_rows(t) = Tables.rows(t)
get_rows(m::AbstractMatrix) = eachrow(m)

function check_size(t::Table)
    if !isnothing(t.column_widths)
        @assert ncols(t) == length(t.column_widths) "column_widths does not match number of columns"
    end
    if !isnothing(t.row_heights)
        @assert nrows(t) == length(t.row_heights) "row_heights does not match number of rows"
    end
end

function _show_string(t::Table, compact::Bool)
    show_string = "Table"
    if !compact
        show_string *= "\n content isa $(typeof(t.content))"
        show_string *= "\n offset_x is $(t.offset_x) EMUs"
        show_string *= "\n offset_y is $(t.offset_y) EMUs"
        show_string *= "\n size_x is $(t.size_x) EMUs"
        show_string *= "\n size_y is $(t.size_y) EMUs"
    end
    return show_string
end

function make_xml(t::Table, id::Integer, relationship_map::Dict = Dict())
    nvGraphicFramePr = make_nvGraphicFramePr(t, id)
    xfrm = make_xfrm(t)
    tbl = make_xml_table(t)
    # a:graphic => a:graphicData => a:tbl
    uri = Dict("uri"=>"http://schemas.openxmlformats.org/drawingml/2006/table")
    graphic = Dict("a:graphic" => Dict("a:graphicData" => [uri, tbl]))
    return Dict("p:graphicFrame" => [nvGraphicFramePr, xfrm, graphic])
end

struct Line
    width::Int
    color::String
    dash::String # solid, dot, dash, dashDot, lgDash, lgDashDot, sysDash, sysDashDotDot
    function Line(
        width::Real,
        color,
        dash::AbstractString = "solid",
    )
        return new(points_to_emu(width), hex_color(color), dash_string(dash))
    end
end

function dash_string(dash)
    dash = string(dash)
    values = ["solid", "dot", "dash", "dashDot", "lgDash", "lgDashDot", "sysDash", "sysDashDotDot"]
    @assert dash in values "unsupported dash value \"$dash\" must be one of $values"
    return dash
end

function Line(;
    width = 1,
    color = colorant"black",
    dash = "solid"
    )
    return Line(width, color, dash_string(dash))
end

Base.convert(::Type{Line}, x::NamedTuple) = Line(;x...)

Base.@kwdef struct TableLines
    left::Union{Nothing, Line} = nothing
    right::Union{Nothing, Line} = nothing
    top::Union{Nothing, Line} = nothing
    bottom::Union{Nothing, Line} = nothing
end

TableLines(t::TableLines) = t
TableLines(nt::NamedTuple) = TableLines(;nt...)

function has_lines(lines::TableLines)
    return !isnothing(lines.left) || !isnothing(lines.right) || !isnothing(lines.top) || !isnothing(lines.bottom)
end

#= example
<a:lnR w="38100" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
    <a:srgbClr val="FF0101"/>
</a:solidFill>
<a:prstDash val="solid"/>
<a:round/>
<a:headEnd type="none" w="med" len="med"/>
<a:tailEnd type="none" w="med" len="med"/>
</a:lnR>
=#
function make_xml(line::Line, type::String = "R")
    return Dict("a:ln$type" => [
        Dict("w" => string(line.width)),
        Dict("cap" => "flat"),
        Dict("cmpd" => "sng"),
        Dict("algn" => "ctr"),
        solid_fill_color(line.color),
        Dict("a:prstDash" => Dict("val" => line.dash)),
        Dict("a:round" => missing),
        Dict("a:headEnd" => [Dict("type" => "none"), Dict("w" => "med"), Dict("len" => "med")]),
        Dict("a:tailEnd" => [Dict("type" => "none"), Dict("w" => "med"), Dict("len" => "med")]),
        ]
    )
end

# margins set to 0.1 mm PPTX.points_to_emu(x)*10 ?
# <a:tcPr marL="36000" marR="36000" marT="36000" marB="36000">
function make_table_xml(m::Margins)
    attr = []
    if !isnothing(m.left)
        push!(attr, Dict("marL" => string(m.left)))
    end
    if !isnothing(m.right)
        push!(attr, Dict("marR" => string(m.right)))
    end
    if !isnothing(m.top)
        push!(attr, Dict("marT" => string(m.top)))
    end
    if !isnothing(m.bottom)
        push!(attr, Dict("marB" => string(m.bottom)))
    end
    return attr
end

"""
```julia
TableCell(
    content; # text
    textstyle = TextStyle(),
    color = nothing, # background color of the table element
    anchor = nothing, # anchoring of text in the cell, can be "top", "bottom" or "center"
    lines,
    margins,
)
```

Create a styled TableCell for use inside a table/dataframe.

# Example

```julia
julia> t = TableCell(4; color = :green, textstyle=(color=:blue,))
TableCell
 text is 4
 textstyle has
  color is 0000FF
 background color is 008000

```
"""
struct TableCell
    textbody::TextBody
    color::Union{Nothing, Missing, String} # hex color
    lines::TableLines
    anchor::Union{Nothing, String} # "top", "bottom", "center"
    direction::Union{Nothing, String} # "vert" or "vert270", nothing gives horizontal
    margins::Margins
end

function TableCell(content; kwargs...)
    return TableCell(;content, kwargs...)
end

function TableCell(;
    content,
    text_style = TextStyle(),
    textstyle = text_style,
    style = textstyle,
    color = nothing,
    lines = TableLines(),
    anchor = nothing,
    direction = nothing,
    margins = Margins(),
)
    textbody = TextBody(; text=string(content), style=TextStyle(style), body_properties=nothing)
    return TableCell(
        textbody,
        hex_color(color),
        TableLines(lines),
        anchor_string(anchor),
        text_direction(direction),
        Margins(margins)
    )
end

function has_tc_properties(c::TableCell)
    return !isnothing(c.color) || has_lines(c.lines)
end

has_margins(c::TableCell) = has_margins(c.margins)

anchor_string(::Nothing) = nothing
function anchor_string(x)
    s = string(x)
    @assert s in ("top", "bottom", "center") "unknown table cell anchor $s, must be top, bottom or center"
    return s
end

text_direction(::Nothing) = nothing
function text_direction(x)
    s = string(x)
    @assert s in ("vert", "vert270")
    return s
end

function Base.show(io::IO, t::TableCell)
    print(io, "TableCell($(t.textbody.text))")
end

function Base.show(io::IO, ::MIME"text/plain", t::TableCell)
    whitespace = 1
    print(io, summary(t))
    text = t.textbody.text
    print(io, "\n" * " "^whitespace * "text is $text")
    if has_non_defaults(t.textbody.style)
        print(io, "\n" * " "^whitespace * "textstyle has" )
        print_style_properties(io, t.textbody.style; whitespace=whitespace+1, only_non_default=true)
    end
    if !isnothing(t.color)
        print(io, "\n" * " "^whitespace * "background color is $(t.color)")
    end
    if !isnothing(t.anchor)
        print(io, "\n" * " "^whitespace * "anchor is $(t.anchor)")
    end
end

#= Example
<p:nvGraphicFramePr>
    <p:cNvPr id="4" name="Table 4">
        <a:extLst>
            <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
                <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{FBDC5980-B3C0-4562-ADB2-48A8BBF2EFB1}"/>
            </a:ext>
        </a:extLst>
    </p:cNvPr>
    <p:cNvGraphicFramePr>
        <a:graphicFrameLocks noGrp="1"/>
    </p:cNvGraphicFramePr>
    <p:nvPr>
        <p:ph idx="1"/>
        <p:extLst>
            <p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}">
                <p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="166699280"/>
            </p:ext>
        </p:extLst>
    </p:nvPr>
</p:nvGraphicFramePr>
=#
function make_nvGraphicFramePr(t::Table, id::Integer)
    cNvPrExtLst = Dict(
        "a:extLst" => [
            Dict(
                "a:ext" => [
                    Dict("uri" => "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"),
                    Dict(
                        "a16:creationId" => [
                            "xmlns:a16" => "http://schemas.microsoft.com/office/drawing/2014/main",
                            "id" => "{FBDC5980-B3C0-4562-ADB2-48A8BBF2EFB1}",
                        ],
                    ),
                ],
            ),
        ],
    )

    cNvPr = Dict(
        "p:cNvPr" => [Dict("id" => "$id"), Dict("name" => "Table $id"), cNvPrExtLst]
    )

    cNvGraphicFramePr = Dict(
        "p:cNvGraphicFramePr" => Dict("a:graphicFrameLocks" => Dict("noGrp" => "1"))
    )

    nvPrextLst = Dict(
        "p:extLst" => [
            Dict(
                "p:ext" => [
                    Dict("uri" => "{D42A27DB-BD31-4B8C-83A1-F6EECF244321}"),
                    Dict(
                        "p14:modId" => [
                            "xmlns:p14" => "http://schemas.microsoft.com/office/powerpoint/2010/main",
                            "val" => "166699280", # does this need to be unique? hardcoding for now
                        ],
                    ),
                ],
            ),
        ],
    )

    nvPr = Dict("p:nvPr" => [Dict("p:ph" => Dict("idx" => "1")), nvPrextLst])

    nvGraphicFramePr = Dict("p:nvGraphicFramePr" => [cNvPr, cNvGraphicFramePr, nvPr])

    return nvGraphicFramePr
end

#= Example
<p:xfrm>
    <a:off x="838200" y="2635522"/>
    <a:ext cx="10515597" cy="2225040"/>
</p:xfrm>
=#
function make_xfrm(t::Table)
    offset = Dict("a:off" => [Dict("x" => "$(t.offset_x)"), Dict("y" => "$(t.offset_y)")])
    extend = Dict("a:ext" => [Dict("cx" => "$(t.size_x)"), Dict("cy" => "$(t.size_y)")])
    return Dict("p:xfrm" => [offset, extend])
end

#= Example 'skeleton' of the XML
<a:tbl>
    <a:tblPr>...</a:tblPr>
    <a:tblGrid>...</a:tblGrid>
    <a:tr>...</a:tr>
    ...
    <a:tr>...</a:tr>
</a:tbl>
=#
function make_xml_table(t::Table)
    tblPr = make_tblPr(t)
    tblGrid = make_tblGrid(t)
    return Dict("a:tbl" => [tblPr, tblGrid, make_rows(t)...])
end

#= Example
<a:tblPr firstRow="1" bandRow="1">
    <a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>
</a:tblPr>
=#
function make_tblPr(t::Table)
    tblPr = Dict(
        "a:tblPr" => [
            Dict("firstRow" => ppt_bool(t.header)),
            Dict("bandRow" => ppt_bool(t.bandrow)),
            Dict("a:tableStyleId" => PlainTextBody(t.style_id)),
        ],
    )
    return tblPr
end

#= tblGrid has a gridCol per column
<a:tblGrid>
    <a:gridCol w="3505199">
        <a:extLst>...</a:extLst>
    </a:gridCol>
    ...
</a:tblGrid>
=#
function make_tblGrid(t::Table)
    if isnothing(t.column_widths)
        nr_of_columns = ncols(t)
        column_width = t.size_x ÷ nr_of_columns
        grid = [make_gridCol(column_width) for _ in 1:nr_of_columns]
    else
        check_size(t)
        grid = [make_gridCol(w) for w in t.column_widths]
    end
    return Dict("a:tblGrid" => grid)
end

#= Example
<a:gridCol w="3505199">
    <a:extLst>
        <a:ext uri="{9D8B030D-6E8A-4147-A177-3AD203B41FA5}">
            <a16:colId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" val="4181689228"/>
        </a:ext>
    </a:extLst>
</a:gridCol>
=#
function make_gridCol(width::Integer)
    extLst = make_single_val_extLst("{9D8B030D-6E8A-4147-A177-3AD203B41FA5}", "colId")
    return Dict(
        "a:gridCol" => [
            Dict("w" => "$width"),
            extLst
        ]
    )
end

#=
<a:tr h="370840"><a:tc><a:txBody>...</a:txBody></a:tc><a:tr>
...
<a:tr h="370840"><a:tc><a:txBody>...</a:txBody></a:tc><a:tr>
=#
function make_rows(t::Table)::Vector
    tr_list = []

    if isnothing(t.row_heights)
        if t.header
            nr_of_rows = nrows(t) + 1
        else
            nr_of_rows = nrows(t)
        end
        height = t.size_y ÷ nr_of_rows
        heights = fill(height, nrows(t))
        header_height = height
    else
        if t.header
            if t.size_y > sum(t.row_heights)
                header_height = t.size_y - sum(t.row_heights)
            else
                header_height = t.row_heights[1]
            end
        else
            heights = t.row_heights
            header_height = 0
        end
    end

    # we also push the column names as a row
    if t.header
        push!(tr_list, make_xml_row(Tables.columnnames(t), header_height))
    end

    for (index, row) in enumerate(Tables.rows(t))
        push!(tr_list, make_xml_row(row, heights[index]))
    end
    return tr_list
end

#=
<a:tr h="370840">
    <a:tc><a:txBody>...</a:txBody><a:tcPr/></a:tc>
    ...
    <a:tc><a:txBody>...</a:txBody><a:tcPr/></a:tc>
    <a:extLst>
        <a:ext uri="{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}">
            <a16:rowId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" val="102079860"/>
        </a:ext>
    </a:extLst>
</a:tr>
=#
function make_xml_row(row, height::Integer)
    tc_list = []
    for element in row
        tc = make_table_cell(element)
        push!(tc_list, tc)
    end
    extLst = make_single_val_extLst("{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}", "rowId")
    tr = Dict("a:tr" => [Dict("h" => "$height"), tc_list..., extLst])
    return tr
end

function make_table_cell(element)
    text = PlainTextBody(string(element))
    return make_table_cell(text)
end

function make_table_cell(text::TextBody)
    tc_properties = Dict("a:tcPr" => missing)
    tc = Dict("a:tc" => [make_textbody_xml(text, "a"), tc_properties])
    return tc
end

function make_table_cell(element::TableCell)
    if has_tc_properties(element)
        tcPr = []

        if !isnothing(element.anchor)
            push!(tcPr, make_anchor(element))
        end

        if has_margins(element)
            append!(tcPr, make_table_xml(element.margins))
        end

        # <a:tcPr vert="vert"> or <a:tcPr vert="vert270">
        if !isnothing(element.direction)
            push!(tcPr, Dict("vert" => element.direction))
        end

        lines = element.lines
        if !isnothing(lines.left)
            push!(tcPr, make_xml(lines.left, "L"))
        end
        if !isnothing(lines.right)
            push!(tcPr, make_xml(lines.right, "R"))
        end
        if !isnothing(lines.top)
            push!(tcPr, make_xml(lines.top, "T"))
        end
        if !isnothing(lines.bottom)
            push!(tcPr, make_xml(lines.bottom, "B"))
        end
        if !isnothing(element.color)
            push!(tcPr, solid_fill_color(element.color))
        end
    else
        tcPr = missing
    end
    tc_properties = Dict("a:tcPr" => tcPr)
    tc = Dict("a:tc" => [make_textbody_xml(element.textbody, "a"), tc_properties])
    return tc
end

function solid_fill_color(color::Missing)
    #<a:solidFill>
    #    <a:sysClr val="windowText" lastClr="000000"/>
    #</a:solidFill>
    return Dict("a:noFill" => missing)
end

function make_anchor(t::TableCell)
    if isnothing(t.anchor)
        return nothing
    elseif t.anchor == "center"
        anchor = "ctr"
    elseif t.anchor == "top"
        anchor = "t"
    elseif t.anchor == "bottom"
        anchor = "b"
    else
        error("unknown table cell anchor \"$(t.anchor)\"")
    end
    return Dict("anchor" => anchor)
end

function make_single_val_extLst(uri::String, type::String, val::Integer = rand(UInt32))
    return Dict(
        "a:extLst" => [
            Dict(
                "a:ext" => [
                    Dict("uri" => uri),
                    Dict(
                        "a16:$type" => [
                            "xmlns:a16" => "http://schemas.microsoft.com/office/drawing/2014/main",
                            "val" => "$val",
                        ],
                    ),
                ],
            ),
        ],
    )
end