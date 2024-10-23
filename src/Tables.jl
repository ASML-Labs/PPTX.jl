"""
```julia
Table(;
    content,
    offset_x::Real = 50,
    offset_y::Real = 50,
    size_x::Real = 150,
    size_y::Real = 100,
)
```

A Table to be used on a Slide.

The content can be anything that adheres to a `Tables.jl` interface.

Offsets and sizes are in millimeters, but will be converted to EMU.

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
    style_id::String
    function Table(
        content,
        offset_x::Real, # millimeters
        offset_y::Real, # millimeters
        size_x::Real, # millimeters
        size_y::Real, # millimeters
        style_id::String="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",
    )
        # input is in mm
        return new(
            content,
            Int(round(offset_x * _EMUS_PER_MM)),
            Int(round(offset_y * _EMUS_PER_MM)),
            Int(round(size_x * _EMUS_PER_MM)),
            Int(round(size_y * _EMUS_PER_MM)),
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
)
    return Table(content, offset_x, offset_y, size_x, size_y)
end

Table(content; kwargs...) = Table(; content=content, kwargs...)

Tables.columnnames(t::Table) = Tables.columnnames(t.content)
Tables.columns(t::Table) = Tables.columns(t.content)
Tables.rows(t::Table) = Tables.rows(t.content)
ncols(t::Table) = length(Tables.columns(t))
nrows(t::Table) = length(Tables.rows(t))

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
    dash::String
    function Line(
        width::Real,
        color::Union{AbstractString, Colorant},
        dash::AbstractString = "solid",
    )
        return new(points_to_emu(width), hex_color(color), string(dash))
    end
end

Base.@kwdef struct TableLines
    left::Union{Nothing, Line} = nothing
    right::Union{Nothing, Line} = nothing
    top::Union{Nothing, Line} = nothing
    bottom::Union{Nothing, Line} = nothing
end

function has_lines(lines::TableLines)
    return !isnothing(lines.left) && !isnothing(lines.right) && !isnothing(lines.top) && !isnothing(lines.bottom)
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
        Dict("w" => line.width),
        Dict("cap" => "flat"),
        Dict("cmpd" => "sng"),
        Dict("algn" => "ctr"),
        solid_fill_color(line.color),
        Dict("a:prstDash" => Dict("val" => line.dash)),
        Dict("round" => missing),
        Dict("a:headEnd" => [Dict("type" => "none"), Dict("w" => "med"), Dict("len" => "med")]),
        Dict("a:tailEnd" => [Dict("type" => "none"), Dict("w" => "med"), Dict("len" => "med")]),
        ]
    )
end

"""
```julia
TableElement(
    content; # text
    textstyle = TextStyle(),
    color = nothing, # background color of the table element
)

Create a styled TableElement for use inside a table/dataframe.

# Example

```julia
julia> t = TableElement(4; color = colorant"green", textstyle=(color=colorant"blue",))
TableElement
 text is 4
 textstyle has
  color is 0000FF
 background color is 008000

```
"""
struct TableElement
    textbody::TextBody
    color::Union{Nothing, String} # hex color
    lines::TableLines
end

function TableElement(content; kwargs...)
    return TableElement(;content, kwargs...)
end

function TableElement(;
    content,
    text_style = TextStyle(),
    textstyle = text_style,
    style = textstyle,
    color::Union{Nothing, String, Colorant} = nothing,
    lines::TableLines = TableLines(),
)
    textbody = TextBody(string(content), style)
    return TableElement(textbody, hex_color(color), lines)
end

function has_tc_properties(element::TableElement)
    return !isnothing(element.color) || has_lines(element.lines)
end

function Base.show(io::IO, t::TableElement)
    print(io, "TableElement($(t.textbody.text))")
end

function Base.show(io::IO, ::MIME"text/plain", t::TableElement)
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
            Dict("firstRow" => "1"),
            Dict("bandRow" => "1"),
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
    nr_of_columns = ncols(t)
    column_width = t.size_x ÷ nr_of_columns
    return Dict("a:tblGrid" => [make_gridCol(column_width) for _ in 1:nr_of_columns])
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

    nr_of_rows = nrows(t) + 1
    height = t.size_y ÷ nr_of_rows

    # we also push the column names as a row
    push!(tr_list, make_xml_row(Tables.columnnames(t), height))

    for row in Tables.rows(t)
        push!(tr_list, make_xml_row(row, height))
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
        tc = make_table_element(element)
        push!(tc_list, tc)
    end
    extLst = make_single_val_extLst("{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}", "rowId")
    tr = Dict("a:tr" => [Dict("h" => "$height"), tc_list..., extLst])
    return tr
end

function make_table_element(element)
    text = PlainTextBody(string(element))
    return make_table_element(text)
end

function make_table_element(text::TextBody)
    tc_properties = Dict("a:tcPr" => missing)
    tc = Dict("a:tc" => [make_textbody_xml(text, "a"), tc_properties])
    return tc
end

function make_table_element(element::TableElement)
    if has_tc_properties(element)
        if !isnothing(element.color)
            tcPr = [solid_fill_color(element.color)]
        else
            tcPr = missing
        end
    else
        tcPr = missing
    end
    
    tc_properties = Dict("a:tcPr" => tcPr)
    tc = Dict("a:tc" => [make_textbody_xml(element.textbody, "a"), tc_properties])
    return tc
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