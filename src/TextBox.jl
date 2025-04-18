"""
```julia
TextStyle(
    bold = false,
    italic = false,
    underscore = false,
    strike = false,
    fontsize = nothing,
    typeface = nothing, # or a string, like \"Courier New\"
    color = nothing, # anything compliant with Colors.jl
    align = nothing, # "left", "right" or "center"
)
```

Style of the text inside a `TextBox`.
You can use Colors.jl colorants for the text color, or directly provide a HEX string, or a symbol like :white.

```jldoctest
julia> using PPTX

julia> style = TextStyle(bold=true, color=:red)
TextStyle
 bold is true
 italic is false
 underscore is false
 strike is false
 fontsize is nothing
 typeface is nothing
 color is FF0000
 align is nothing

julia> text = TextBox(content = "hello"; style)
TextBox
 content is "hello"
 content.style has
  bold is true
  color is FF0000
 offset_x is 1800000 EMUs
 offset_y is 1800000 EMUs
 size_x is 1440000 EMUs
 size_y is 1080000 EMUs

```
"""
struct TextStyle
    bold::Bool
    italic::Bool
    underscore::Bool
    strike::Bool
    fontsize::Union{Nothing, Float64}
    typeface::Union{Nothing, String}
    color::Union{Nothing, String}
    align::Union{Nothing, String}
end

function TextStyle(;
    bold::Bool = false,
    italic::Bool = false,
    underscore::Bool = false,
    strike::Bool = false,
    fontsize = nothing, # nothing will use default font
    typeface = nothing, # exact string of the type face
    color = nothing, # hex string
    align = nothing, # "left", "center", "right"
    )
    return TextStyle(
        bold,
        italic,
        underscore,
        strike,
        parse_fontsize(fontsize),
        typeface,
        hex_color(color),
        align_string(align),
    )
end

parse_fontsize(::Nothing) = nothing
parse_fontsize(x::Real) = Float64(x)

hex_color(t::TextStyle) = hex_color(t.color)
hex_color(c::String) = c
hex_color(c::Colorant) = hex(c)
hex_color(::Nothing) = nothing
hex_color(::Missing) = missing
hex_color(c::Symbol) = hex(parse(Colorant, c))

align_string(::Nothing) = nothing
function align_string(x)
    s = string(x)
    @assert s in ("left", "center", "right") "unknown text align $s, must be left, center or right"
    return s
end

TextStyle(style::TextStyle) = style

function TextStyle(style::AbstractDict{String})
    kw_pairs = [Symbol(lowercase(k)) => v for (k,v) in style]
    return TextStyle(; kw_pairs...)
end

function TextStyle(nt::NamedTuple)
    return TextStyle(; nt...)
end

function Base.show(io::IO, ::MIME"text/plain", t::TextStyle)
    print(io, summary(t))
    print_style_properties(io, t)
end

function has_non_defaults(t::TextStyle)
    for p in propertynames(t)
        prop = getproperty(t, p)
        if !(isnothing(prop) || prop == false)
            return true
        end
    end
    return false
end

function print_style_properties(io::IO, t::TextStyle; whitespace::Int=1, only_non_default=false)
    for p in propertynames(t)
        prop = getproperty(t, p)
        if only_non_default
            if isnothing(prop) || prop == false
                continue
            end
        end
        if p == :color
            print(io, "\n" * " "^whitespace * "$p is $(hex_color(prop))")
        else
            print(io, "\n" * " "^whitespace * "$p is $prop")
        end
    end
end

function style_properties_string(t::TextStyle, whitespace::Int=1)
    io = IOBuffer()
    print_style_properties(io, t; whitespace, only_non_default=true)
    return String(take!(io))
end


struct Margins
    left::Union{Nothing, Int}
    right::Union{Nothing, Int}
    top::Union{Nothing, Int}
    bottom::Union{Nothing, Int}
end

function Margins(;
    left = nothing,
    right = nothing,
    top = nothing,
    bottom = nothing,
    )
    return Margins(margin(left), margin(right), margin(top), margin(bottom))
end

Margins(t::Margins) = t
Margins(nt::NamedTuple) = Margins(;nt...)

margin(::Nothing) = nothing
margin(x) = mm_to_emu(x*10)

function has_margins(m::Margins)
    return !isnothing(m.left) || !isnothing(m.right) || !isnothing(m.top) || !isnothing(m.bottom) 
end

Base.@kwdef struct TextBody
    text::String
    style::TextStyle = TextStyle()
    margins::Margins = Margins()
    wrap::Bool = false
    body_properties::Union{Nothing, AbstractVector} = default_body_properties()
end

has_margins(t::TextBody) = has_margins(t.margins)

# <a:bodyPr wrap="none" lIns="108000" tIns="180000" rIns="108000" bIns="108000" rtlCol="0">
# <a:spAutoFit/>
# </a:bodyPr>
function make_body_properties_xml(m::Margins)
    attr = []
    if !isnothing(m.left)
        push!(attr, Dict("lIns" => string(m.left)))
    end
    if !isnothing(m.top)
        push!(attr, Dict("tIns" => string(m.top)))
    end
    if !isnothing(m.right)
        push!(attr, Dict("rIns" => string(m.right)))
    end
    if !isnothing(m.bottom)
        push!(attr, Dict("bIns" => string(m.bottom)))
    end
    return attr
end

function TextBody(text::AbstractString; kwargs...)
    return TextBody(;text=convert(String, text), kwargs...)
end

function TextBody(text::AbstractString, style::AbstractDict)
    return TextBody(;text=convert(String, text), style=TextStyle(style))
end

function TextBody(text::AbstractString, style::NamedTuple)
    return TextBody(;text=convert(String, text), style=TextStyle(;style...))
end

function TextBody(text::AbstractString, style::TextStyle)
    return TextBody(;text=convert(String, text), style=style)
end

function PlainTextBody(text::AbstractString)
    return TextBody(;text=convert(String, text), body_properties=nothing)
end

Base.convert(::Type{TextBody}, s::AbstractString) = TextBody(String(s))
Base.String(s::TextBody) = s.text

function default_body_properties()
    return [
        Dict("rtlCol" => "0"),
        Dict("a:spAutoFit" => missing),
    ]
end

"""
```julia
function TextBox(;
    content::String = "",
    offset = (50,50), # millimeters
    offset_x = offset[1],
    offset_y = offset[2],
    size = (40,30), # millimeters
    size_x = size[1],
    size_y = size[2],
    hlink = nothing, # hyperlink
    color = nothing, # use hex string, or Colorant
    linecolor = nothing, # use hex string, or Colorant
    linewidth = nothing, # use value in points, e.g. 3
    rotation = nothing, # use a value in degrees, e.g. 90
    textstyle = (italic = false, bold = false, fontsize = nothing),
    margins = nothing, # e.g. (left=0.1, right=0.1, bottom=0.1, top=0.1) in millimeters
    wrap = false, # wrap text in shape or not
)
```

A TextBox to be used on a Slide.
Offsets and sizes are in millimeters, but will be converted to EMU.

See [`TextStyle`](@ref TextStyle) for more text style options.

# Examples
```jldoctest
using PPTX

text = TextBox(
    content="Hello world!",
    offset=(100, 50),
    size=(30,50),
    textstyle=(color=:white, bold=true),
    color=:blue,
    linecolor=:black,
    linewidth=3
)

# output

TextBox
 content is "Hello world!"
 content.style has
  bold is true
  color is FFFFFF
 offset_x is 3600000 EMUs
 offset_y is 1800000 EMUs
 size_x is 1080000 EMUs
 size_y is 1800000 EMUs
 color is 0000FF
 linecolor is 000000
 linewidth is 38100 EMUs

```
"""
struct TextBox<: AbstractShape
    content::TextBody
    offset_x::Int # EMUs
    offset_y::Int # EMUs
    size_x::Int # EMUs
    size_y::Int # EMUs
    hlink::Union{Nothing, Any}
    color::Union{Nothing, String}
    linecolor::Union{Nothing, String}
    linewidth::Union{Nothing, Int}
    rotation::Union{Nothing, Float64}
    function TextBox(
        content::AbstractString,
        offset_x::Real, # millimeters
        offset_y::Real, # millimeters
        size_x::Real, # millimeters
        size_y::Real, # millimeters
        style = TextStyle(),
        hlink::Union{Nothing, Any} = nothing,
        color = nothing,
        linecolor = nothing,
        linewidth::Union{Nothing, Real} = 1,
        rotation::Union{Nothing, Real} = nothing,
        margins = Margins(),
        wrap = false,
    )
        # input is in mm
        return new(
            TextBody(
                string(content),
                TextStyle(style),
                Margins(margins),
                wrap,
                default_body_properties()
            ),
            mm_to_emu(offset_x),
            mm_to_emu(offset_y),
            mm_to_emu(size_x),
            mm_to_emu(size_y),
            hlink,
            hex_color(color),
            hex_color(linecolor),
            points_to_emu(linewidth),
            rotation_value(rotation),
        )
    end
end

mm_to_emu(::Nothing) = nothing
mm_to_emu(x) = Int(round(x * _EMUS_PER_MM))
mm_to_emu(x::AbstractArray{<:Real}) = mm_to_emu.(x)

points_to_emu(x::Nothing) = nothing
points_to_emu(x::Real) = Int(round(x * 12700))

rotation_value(::Nothing) = nothing
function rotation_value(x::Real)
    return mod(Float64(x), 360.0)
end

# keyword argument constructor
function TextBox(;
    content::AbstractString="",
    offset=(50,50),
    offset_x::Real=offset[1], # millimeters
    offset_y::Real=offset[2], # millimeters
    size=(40,30),
    size_x::Real=size[1], # millimeters
    size_y::Real=size[2], # millimeters
    text_style=TextStyle(),
    textstyle=text_style,
    style=textstyle,
    hlink::Union{Nothing, Any}=nothing,
    color=nothing,
    linecolor=nothing,
    linewidth::Union{Nothing, Int}=nothing,
    rotation::Union{Nothing, Real}=nothing,
    margins=Margins(),
    wrap=false,
)
    return TextBox(
        content,
        offset_x,
        offset_y,
        size_x,
        size_y,
        style,
        hlink,
        hex_color(color),
        hex_color(linecolor),
        linewidth,
        rotation,
        margins,
        wrap,
    )
end

TextBox(content::String; kwargs...) = TextBox(;content=content, kwargs...)

function _show_string(p::TextBox, compact::Bool)
    show_string = "TextBox"
    if !compact
        show_string *= "\n content is \"$(String(p.content))\""
        if has_non_defaults(p.content.style)
            show_string *= "\n content.style has"
            show_string *= style_properties_string(p.content.style, 2)
        end
        show_string *= "\n offset_x is $(p.offset_x) EMUs"
        show_string *= "\n offset_y is $(p.offset_y) EMUs"
        show_string *= "\n size_x is $(p.size_x) EMUs"
        show_string *= "\n size_y is $(p.size_y) EMUs"
        if !isnothing(p.color)
            show_string *= "\n color is $(p.color)"
        end
        if !isnothing(p.linecolor)
            show_string *= "\n linecolor is $(p.linecolor)"
        end
        if !isnothing(p.linewidth)
            show_string *= "\n linewidth is $(p.linewidth) EMUs"
        end
        if !isnothing(p.rotation)
            show_string *= "\n rotation is $(p.rotation) degrees"
        end
    end
    return show_string
end

function text_style_xml(t::TextBody)
    return text_style_xml(t.style)
end

function text_style_xml(t::TextStyle)
    style = Vector{Dict{String}}()
    push!(style, Dict("lang" => "en-US"))
    if t.bold
        push!(style, Dict("b" => "1"))
    end

    if t.italic
        push!(style, Dict("i" => "1"))
    end

    if t.underscore
        push!(style, Dict("u" => "sng"))
    end

    if t.strike
        push!(style, Dict("strike" => "sngStrike"))
    end

    if !isnothing(t.fontsize)
        sz = string(Int(round(t.fontsize*100)))
        push!(style, Dict("sz" => sz))
    end

    # no idea what this does
    push!(style, Dict("dirty" => "0"))
    # I think this sets the wiggly error notification
    push!(style, Dict("err" => "0"))

    if !isnothing(t.color)
        push!(style, solid_fill_color(hex_color(t.color)))
    end

    # there's a lot going on here, but seems only providing the latin typeface is sufficient
    #<a:latin typeface="Courier New" panose="02070309020205020404" pitchFamily="49" charset="0"/>
    #<a:cs typeface="Courier New" panose="02070309020205020404" pitchFamily="49" charset="0"/>
    if !isnothing(t.typeface)
        push!(style, Dict("a:latin" => Dict("typeface" => string(t.typeface))))
    end

    return style
end

function solid_fill_color(color::String)
    return Dict("a:solidFill" => Dict("a:srgbClr" => Dict("val" => color)))
end

function make_xml(t::TextBox, id::Integer, relationship_map::Dict)
    cNvPr = Dict("p:cNvPr" => Dict[Dict("id" => "$id"), Dict("name" => "TextBox")])
    if has_hyperlink(t)
        push!(cNvPr["p:cNvPr"], hlink_xml(t.hlink, relationship_map))
    end

    cNvSpPr = Dict("p:cNvSpPr" => Dict("txBox" => "1"))
    nvPr = Dict("p:nvPr" => missing)

    nvSpPr = Dict("p:nvSpPr" => [cNvPr, cNvSpPr, nvPr])

    xfrm = []
    if !isnothing(t.rotation)
        # degrees * 60000 apparently
        push!(xfrm, Dict("rot" => string(Int(t.rotation*60_000))))
    end
    offset = Dict("a:off" => [Dict("x" => "$(t.offset_x)"), Dict("y" => "$(t.offset_y)")])
    push!(xfrm, offset)
    extend = Dict("a:ext" => [Dict("cx" => "$(t.size_x)"), Dict("cy" => "$(t.size_y)")])
    push!(xfrm, extend)

    spPr_content = [
        Dict("a:xfrm" => xfrm),
        Dict("a:prstGeom" => [Dict("prst" => "rect"), Dict("a:avLst" => missing)]),
    ]
    if !isnothing(t.color)
        push!(spPr_content, solid_fill_color(t.color))
    end
    if !isnothing(t.linecolor)
        if isnothing(t.linewidth)
            w = points_to_emu(1)
        else
            w = t.linewidth
        end
        push!(spPr_content, Dict("a:ln" => [
                    Dict("w" => string(w)),
                    solid_fill_color(t.linecolor)
                ]
            )
        )
    end

    spPr = Dict("p:spPr" => spPr_content)

    txBody = make_textbody_xml(t)

    return Dict("p:sp" => [nvSpPr, spPr, txBody])
end

function make_textbody_xml(t::TextBody, txBodyNameSpace="p")
    ap = []

    algn = make_textalign(t.style)
    if !isnothing(algn)
        push!(ap, Dict("a:pPr" => algn))
    end

    ar = Dict(
        "a:r" => [Dict("a:rPr" => text_style_xml(t)), Dict("a:t" => t)],
    )
    push!(ap, ar)

    if t.wrap
        wrap = "square"
    else
        wrap = "none"
    end
    bodyPr = Dict{String}[Dict{String,String}("wrap" => wrap)]

    if !isnothing(t.body_properties)
        append!(bodyPr, t.body_properties)
    end
    if has_margins(t)
        append!(bodyPr, make_body_properties_xml(t.margins))
    end

    txBody = Dict(
        "$txBodyNameSpace:txBody" => [
            Dict(
                "a:bodyPr" => bodyPr,
            ),
            Dict("a:lstStyle" => missing),
            Dict(
                "a:p" => ap,
            ),
        ],
    )
    return txBody
end

make_textbody_xml(t::TextBox) = make_textbody_xml(t.content)

function make_textalign(t::TextStyle)
    if isnothing(t.align)
        return nothing
    elseif t.align == "center"
        align = "ctr"
    elseif t.align == "right"
        align = "r"
    elseif t.align == "left"
        align = "l"
    else
        error("unknown text align \"$(t.align)\"")
    end
    return Dict("algn" => align)
end