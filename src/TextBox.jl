"""
```julia
TextStyle(
    bold = false,
    italic = false,
    underscore = false,
    strike = false,
    fontsize = nothing,
    color = nothing,
)
```

Style of the text inside a `TextBox`.
You can use Colors.jl colorants for the text color, or directly provide a HEX string.

```jldoctest
julia> using PPTX, Colors

julia> style = TextStyle(bold=true, color=colorant"red")
TextStyle
 bold is true
 italic is false
 underscore is false
 strike is false
 fontsize is nothing
 color is FF0000

julia> text = TextBox(content = "hello"; style)
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
Base.@kwdef struct TextStyle
    bold::Bool = false
    italic::Bool = false
    underscore::Bool = false
    strike::Bool = false
    fontsize::Union{Nothing, Float64} = nothing # nothing will use default font
    color::Union{Nothing, String, Colorant} = nothing # or hex string
end

hex_color(t::TextStyle) = hex_color(t.color)
hex_color(c::String) = c
hex_color(c::Colorant) = hex(c)

function TextStyle(style::AbstractDict{String})
    kw_pairs = [Symbol(lowercase(k)) => v for (k,v) in style]
    return TextStyle(; kw_pairs...)
end

function Base.show(io::IO, ::MIME"text/plain", t::TextStyle)
    print(io, summary(t))
    print_style_properties(io, t)
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

Base.@kwdef struct TextBody
    text::String
    style::TextStyle = TextStyle()
    body_properties::Union{Nothing, AbstractVector} = default_body_properties()
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
        Dict("wrap" => "square"),
        Dict("rtlCol" => "0"),
        Dict("a:spAutoFit" => missing),
    ]
end

"""
```
function TextBox(;
    content::String = "",
    offset = (50,50),
    offset_x::Real = offset[1], # millimeters
    offset_y::Real = offset[2], # millimeters
    size = (40,30),
    size_x::Real = size[1], # millimeters
    size_y::Real = size[2], # millimeters
    style = (italic = false, bold = false, fontsize = nothing),
)
```

A TextBox to be used on a Slide.
Offsets and sizes are in millimeters, but will be converted to EMU.

See `TextStyle` for more text style options.

# Examples
```jldoctest
julia> using PPTX

julia> text = TextBox(content="Hello world!", offset=(100, 50), size=(30,50), style = (italic = true, fontsize = 24))
TextBox
 content is "Hello world!"
 content.style has
  italic is true
  fontsize is 24.0
 offset_x is 3600000 EMUs
 offset_y is 1800000 EMUs
 size_x is 1080000 EMUs
 size_y is 1800000 EMUs

```
"""
struct TextBox<: AbstractShape
    content::TextBody
    offset_x::Int # EMUs
    offset_y::Int # EMUs
    size_x::Int # EMUs
    size_y::Int # EMUs
    hlink::Union{Nothing, Any}
    function TextBox(
        content::AbstractString,
        offset_x::Real, # millimeters
        offset_y::Real, # millimeters
        size_x::Real, # millimeters
        size_y::Real, # millimeters
        style = TextStyle(),
        hlink::Union{Nothing, Any} = nothing
    )
        # input is in mm
        return new(
            TextBody(content, style),
            Int(round(offset_x * _EMUS_PER_MM)),
            Int(round(offset_y * _EMUS_PER_MM)),
            Int(round(size_x * _EMUS_PER_MM)),
            Int(round(size_y * _EMUS_PER_MM)),
            hlink
        )
    end
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
    style = TextStyle(),
    hlink::Union{Nothing, Any}=nothing
)
    return TextBox(
        content,
        offset_x,
        offset_y,
        size_x,
        size_y,
        style,
        hlink
    )
end

TextBox(content::String; kwargs...) = TextBox(;content=content, kwargs...)

function _show_string(p::TextBox, compact::Bool)
    show_string = "TextBox"
    if !compact
        show_string *= "\n content is \"$(String(p.content))\""
        show_string *= "\n content.style has"
        show_string *= style_properties_string(p.content.style, 2)
        show_string *= "\n offset_x is $(p.offset_x) EMUs"
        show_string *= "\n offset_y is $(p.offset_y) EMUs"
        show_string *= "\n size_x is $(p.size_x) EMUs"
        show_string *= "\n size_y is $(p.size_y) EMUs"
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

    push!(style, Dict("dirty" => "0"))

    if !isnothing(t.color)
        clr = Dict("a:srgbClr" => Dict("val" => t.color))
        push!(style, Dict("a:solidFill" => clr))
    end
    return style
end

function make_xml(t::TextBox, id::Integer, relationship_map::Dict)
    cNvPr = Dict("p:cNvPr" => Dict[Dict("id" => "$id"), Dict("name" => "TextBox")])
    if has_hyperlink(t)
        push!(cNvPr["p:cNvPr"], hlink_xml(t.hlink, relationship_map))
    end

    cNvSpPr = Dict("p:cNvSpPr" => Dict("txBox" => "1"))
    nvPr = Dict("p:nvPr" => missing)

    nvSpPr = Dict("p:nvSpPr" => [cNvPr, cNvSpPr, nvPr])

    offset = Dict("a:off" => [Dict("x" => "$(t.offset_x)"), Dict("y" => "$(t.offset_y)")])
    extend = Dict("a:ext" => [Dict("cx" => "$(t.size_x)"), Dict("cy" => "$(t.size_y)")])

    spPr = Dict(
        "p:spPr" => [
            Dict("a:xfrm" => [offset, extend]),
            Dict("a:prstGeomt" => [Dict("prst" => "rect"), Dict("a:avLst" => missing)]),
            Dict("a:noFill" => missing),
        ],
    )

    txBody = make_textbody_xml(t)

    return Dict("p:sp" => [nvSpPr, spPr, txBody])
end

function make_textbody_xml(t::TextBody, txBodyNameSpace="p")
    txBody = Dict(
        "$txBodyNameSpace:txBody" => [
            Dict(
                "a:bodyPr" => t.body_properties,
            ),
            Dict("a:lstStyle" => missing),
            Dict(
                "a:p" => Dict(
                    "a:r" => [Dict("a:rPr" => text_style_xml(t)), Dict("a:t" => t)],
                ),
            ),
        ],
    )
    return txBody
end

make_textbody_xml(t::TextBox) = make_textbody_xml(t.content)