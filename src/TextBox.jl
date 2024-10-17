Base.@kwdef struct TextStyle
    bold::Bool = false
    italic::Bool = false
    fontsize::Union{Nothing, Float64} = nothing # nothing will use default font
end

function TextStyle(style::AbstractDict{String})
    return TextStyle(;
        bold = get(style, "bold", false),
        italic = get(style, "italic", false),
        fontsize = get(style, "fontsize", nothing),
    )
end

function Base.show(io::IO, ::MIME"text/plain", t::TextStyle)
    print(io, summary(t))
    print_style_properties(io, t)
end

function print_style_properties(io::IO, t::TextStyle, whitespace::Int=1)
    for p in propertynames(t)
        print(io, "\n" * " "^whitespace * "$p is $(getproperty(t, p))")
    end
end

function style_properties_string(t::TextStyle, whitespace::Int=1)
    io = IOBuffer()
    print_style_properties(io, t, whitespace)
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

# Examples
```jldoctest
julia> using PPTX

julia> text = TextBox(content="Hello world!", offset=(100, 50), size=(30,50), style = (italic = true, fontsize = 24))
TextBox
 content is "Hello world!"
 content.style has
  bold is false
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
    style = [Dict("lang" => "en-US")]
    if t.bold
        push!(style, Dict("b" => "1"))
    end

    if t.italic
        push!(style, Dict("i" => "1"))
    end

    if !isnothing(t.fontsize)
        sz = string(Int(round(t.fontsize*100)))
        push!(style, Dict("sz" => sz))
    end

    push!(style, Dict("dirty" => "0"))
    return style
end

function make_xml(t::TextBox, id::Int=1)
    cNvPr = Dict("p:cNvPr" => Dict[Dict("id" => "$id"), Dict("name" => "TextBox")])
    if has_hyperlink(t)
        push!(cNvPr["p:cNvPr"], hlink_xml(t.hlink))
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